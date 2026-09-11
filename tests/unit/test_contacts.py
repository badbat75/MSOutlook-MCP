"""Unit tests for the contact tools and the address book they read.

The request behind these tools: the Birthdays calendar showed three entries on
7 June for one child, and that calendar gives only the day and the month. The
contacts behind them had to be found, compared and merged, which Graph does not
make easy: /me/contacts is the default folder only, a name cannot be searched
for with $filter, an update that leaves out displayName may regenerate it, and a
delete has been seen answering 404 while carrying itself out. These tests pin
what the tools do about each, against a fake Graph shaped like the real mailbox.
"""

from datetime import date

import anyio
import httpx
import pytest
from pydantic import ValidationError

from outlook_mcp import contacts as contacts_module
from outlook_mcp.app import mcp
from outlook_mcp.contacts import (
    BIRTHDAY_TIME,
    birthday_date,
    birthday_value,
    matches,
    month_day,
    sort_key,
)
from outlook_mcp.models import (
    DeleteContactInput,
    GetContactInput,
    ListContactsInput,
    UpdateContactInput,
)
from outlook_mcp.tools import contacts as tools

NEXT_LINK = "https://graph.microsoft.com/v1.0/me/contactFolders/F-phone/contacts?$skip=1"


def contact(cid, name, **fields):
    return {"id": cid, "displayName": name, **fields}


# The two duplicates from the report, as the real mailbox holds them: one filed
# with a surname, one synced from a phone with the endearment in the suffix.
KEEPER = contact(
    "C-keeper", "Gabriele De Simoni", givenName="Gabriele", surname="De Simoni",
    fileAs="De Simoni, Gabriele", birthday="2016-06-07T11:59:00Z", emailAddresses=[],
    createdDateTime="2020-05-31T10:30:03Z",
)
DUPLICATE = contact(
    "C-dup", "Gabriele, Amore di Papà", givenName="Gabriele", generation="Amore di Papà",
    birthday="2016-06-07T11:59:00Z",
    emailAddresses=[{"name": "Gabriele", "address": "gabriele@example.com"}],
    createdDateTime="2026-06-06T07:27:20Z",
)
NAMESAKE = contact(
    "C-namesake", "Gabriele Giambartolomei", birthday="1974-05-27T11:59:00Z",
    emailAddresses=[{"name": "G", "address": "gg@example.com"}], mobilePhone="+39 338 000",
)
PHONE_ONLY = contact("C-nameless", None, mobilePhone="379 135 2578")
ON_PHONE = contact("C-phone1", "Papà", homePhones=["+39 06 000"])
ON_PHONE_PAGE_2 = contact("C-phone2", "Zoe", emailAddresses=[{"address": "zoe@example.com"}])
NESTED = contact("C-nested", "Nicolò", birthday="2012-06-07T11:59:00Z")


def not_found(method, url):
    request = httpx.Request(method, f"https://graph.microsoft.com/v1.0{url}")
    response = httpx.Response(
        404, request=request,
        json={"error": {"code": "ErrorItemNotFound", "message": "The specified object was not found in the store."}},
    )
    return httpx.HTTPStatusError("404", request=request, response=response)


class FakeGraph:
    """A mailbox with a default folder, a phone's folder in it, and one nested deeper."""

    def __init__(self):
        self.folders = {
            "F-default": ("Contacts", [KEEPER, DUPLICATE, NAMESAKE, PHONE_ONLY]),
            "F-phone": ("HUAWEI P40 Pro", [ON_PHONE, ON_PHONE_PAGE_2]),
            "F-nested": ("Family", [NESTED]),
        }
        self.children = {"F-default": ["F-phone"], "F-phone": ["F-nested"], "F-nested": []}
        self.deleted_items = []
        self.delete_answers_404 = False
        self.calls = []

    def _contact(self, cid):
        for _, items in self.folders.values():
            for item in items:
                if item["id"] == cid:
                    return item
        return None

    async def get(self, endpoint, params=None, headers=None):
        self.calls.append(("GET", endpoint, params))
        if endpoint == "/me/contactFolders/contacts":
            return {"id": "F-default", "displayName": "Contacts"}
        if endpoint == "/me/contactFolders":
            return {"value": [{"id": f, "displayName": self.folders[f][0]} for f in self.children["F-default"]]}
        if endpoint.endswith("/childFolders"):
            parent = endpoint.split("/")[3]
            return {"value": [{"id": f, "displayName": self.folders[f][0]} for f in self.children[parent]]}
        if endpoint == "/me/contactFolders/deleteditems/contacts":
            return {"value": list(self.deleted_items)}
        if endpoint == "/me/contactFolders/F-phone/contacts":
            # Two pages, the second behind an absolute @odata.nextLink.
            return {"value": [ON_PHONE], "@odata.nextLink": NEXT_LINK}
        if endpoint == NEXT_LINK:
            return {"value": [ON_PHONE_PAGE_2]}
        if endpoint.startswith("/me/contactFolders/") and endpoint.endswith("/contacts"):
            return {"value": list(self.folders[endpoint.split("/")[3]][1])}
        if endpoint.startswith("/me/contactFolders/"):
            folder = self.folders.get(endpoint.split("/")[3])
            if folder is None:
                raise not_found("GET", endpoint)
            return {"displayName": folder[0]}
        if endpoint.startswith("/me/contacts/"):
            found = self._contact(endpoint.split("/")[3])
            if found is None:
                raise not_found("GET", endpoint)
            return dict(found)
        raise AssertionError(f"unexpected GET {endpoint}")

    async def patch(self, endpoint, json_data=None):
        self.calls.append(("PATCH", endpoint, json_data))
        return {**self._contact(endpoint.split("/")[3]), **json_data}

    async def delete(self, endpoint):
        self.calls.append(("DELETE", endpoint, None))
        cid = endpoint.split("/")[3]
        for name, items in self.folders.values():
            for item in list(items):
                if item["id"] == cid:
                    items.remove(item)
                    self.deleted_items.append({**item, "id": "moved-" + cid})
        if self.delete_answers_404:
            raise not_found("DELETE", endpoint)
        return {"status": "success"}

    async def post(self, endpoint, json_data=None):
        self.calls.append(("POST", endpoint, json_data))
        return {"status": "success"}

    def sent(self, method):
        bodies = [call[2] for call in self.calls if call[0] == method]
        assert len(bodies) == 1, f"expected one {method}, got {self.calls}"
        return bodies[0]


@pytest.fixture
def graph(monkeypatch):
    fake = FakeGraph()
    monkeypatch.setattr(tools, "get_graph", lambda ctx: fake)
    monkeypatch.setattr(contacts_module, "RECHECK_SECONDS", 0)
    return fake


def list_contacts(**fields):
    return anyio.run(tools.outlook_list_contacts, ListContactsInput(**fields), None)


def get_contact(cid):
    return anyio.run(tools.outlook_get_contact, GetContactInput(contact_id=cid), None)


def update(cid="C-keeper", **fields):
    return anyio.run(tools.outlook_update_contact, UpdateContactInput(contact_id=cid, **fields), None)


def delete(cid):
    return anyio.run(tools.outlook_delete_contact, DeleteContactInput(contact_id=cid), None)


class TestMatching:
    def test_accents_and_case_do_not_matter(self):
        assert matches(DUPLICATE, "papa")
        assert matches(DUPLICATE, "GABRIELE")
        assert matches(contact("x", "Nicolò"), "nicolo")

    def test_every_name_and_address_is_looked_at(self):
        assert matches(KEEPER, "de simoni, gab")      # fileAs
        assert matches(DUPLICATE, "@example.com")      # an address
        assert matches(contact("x", None, nickName="Gabri"), "gabri")

    def test_a_missing_field_is_not_an_error(self):
        assert not matches(PHONE_ONLY, "gabriele")

    def test_no_match(self):
        assert not matches(NAMESAKE, "de simoni")


class TestDaysOfTheYear:
    @pytest.mark.parametrize("given", ["2026-06-07", "2016-06-07", "--06-07", " 2026-06-07 "])
    def test_a_full_date_or_a_month_day(self, given):
        assert month_day(given) == (6, 7)

    def test_29_february_is_a_day_someone_is_born_on(self):
        assert month_day("--02-29") == (2, 29)

    @pytest.mark.parametrize("given", ["07-06", "07/06", "07/06/2026", "June 7", "2026-02-30", "--13-01"])
    def test_anything_ambiguous_or_impossible_is_refused(self, given):
        # "07-06" is 7 June to the person who reported the bug and 6 July to
        # an ISO reader: refusing it is the only answer that is never wrong.
        with pytest.raises(ValueError, match="ISO date"):
            month_day(given)

    def test_the_model_refuses_it_before_graph_is_asked(self):
        with pytest.raises(ValidationError):
            ListContactsInput(birthday="07-06")


class TestBirthdays:
    def test_the_stored_instant_is_read_as_its_utc_date(self):
        # 11:59 UTC, what Outlook writes and every birthday in the real mailbox
        # carries, is the date the Birthdays calendar shows.
        assert birthday_date("2016-06-07T11:59:00Z") == date(2016, 6, 7)

    @pytest.mark.parametrize("value", [None, "", "garbage"])
    def test_no_birthday(self, value):
        assert birthday_date(value) is None

    def test_a_date_is_written_the_way_outlook_writes_it(self):
        assert BIRTHDAY_TIME == "T11:59:00Z"
        assert birthday_value("2016-06-07") == "2016-06-07T11:59:00Z"

    def test_an_empty_string_removes_it(self):
        assert birthday_value("") is None

    @pytest.mark.parametrize("value", ["07/06/2016", "2016-6-7", "2016-02-30"])
    def test_a_malformed_date_is_refused(self, value):
        with pytest.raises(ValueError):
            birthday_value(value)


class TestSorting:
    def test_by_name_ignoring_accents_with_the_nameless_last(self):
        people = [PHONE_ONLY, contact("z", "Zoe"), contact("e", "Émile"), contact("a", "anna")]
        assert [c["id"] for c in sorted(people, key=sort_key)] == ["a", "e", "z", "C-nameless"]


class TestListing:
    def test_the_birthday_from_the_report_finds_both_duplicates(self, graph):
        result = list_contacts(birthday="2026-06-07")

        assert "Gabriele De Simoni" in result
        assert "Gabriele, Amore di Papà" in result
        assert "Giambartolomei" not in result
        # The nested folder is walked too: another child born on 7 June.
        assert "Nicolò" in result
        assert "3 in all folders" in result

    def test_every_folder_is_read_nested_and_paged(self, graph):
        result = list_contacts(top=200)

        assert "7 in all folders" in result
        assert "Folder: HUAWEI P40 Pro" in result
        assert "Folder: Family" in result
        assert "Zoe" in result, "the second page, behind @odata.nextLink"
        assert ("GET", NEXT_LINK, None) in graph.calls

    def test_the_default_folder_is_read_by_its_well_known_name(self, graph):
        list_contacts()
        assert graph.calls[0][1] == "/me/contactFolders/contacts"

    def test_a_search_spans_the_folders(self, graph):
        result = list_contacts(search="papa")
        assert "Gabriele, Amore di Papà" in result
        assert "Folder: HUAWEI P40 Pro" in result

    def test_what_a_summary_shows(self, graph):
        result = list_contacts(search="giambartolomei")
        assert "Email: gg@example.com" in result
        assert "Phone: +39 338 000 (mobile)" in result
        assert "Birthday: 1974-05-27" in result
        assert "ID: `C-namesake`" in result

    def test_pages_say_where_the_next_one_starts(self, graph):
        result = list_contacts(top=2)
        assert "showing 1-2" in result
        assert "skip=2" in result

        last = list_contacts(top=2, skip=6)
        assert "showing 7-7" in last
        assert "skip=" not in last.split("\n\n")[-1]

    def test_nothing_found_is_said(self, graph):
        assert list_contacts(search="nobody") == "No contacts matching 'nobody'."


class TestReading:
    def test_everything_a_duplicate_holds_is_shown(self, graph):
        result = get_contact("C-dup")
        assert "# Gabriele, Amore di Papà" in result
        assert "**Suffix:** Amore di Papà" in result
        assert "Gabriele <gabriele@example.com>" in result
        assert "**Birthday:** 2016-06-07" in result
        assert "**Created:** 2026-06-06T07:27:20Z" in result

    def test_the_folder_is_named(self, graph):
        graph.folders["F-default"][1][1] = {**DUPLICATE, "parentFolderId": "F-default"}
        assert "**Folder:** Contacts" in get_contact("C-dup")

    def test_an_unreadable_folder_is_left_out_not_fatal(self, graph):
        graph.folders["F-default"][1][0] = {**KEEPER, "parentFolderId": "F-gone"}
        result = get_contact("C-keeper")
        assert "Gabriele De Simoni" in result
        assert "Folder" not in result

    def test_an_unknown_id_is_a_404_message(self, graph):
        assert get_contact("C-nope").startswith("Error 404")


class TestUpdating:
    def test_the_display_name_goes_along_when_a_name_part_changes(self, graph):
        # Measured: a PATCH of givenName alone turned "Zz Test MCP (temporaneo)"
        # into "Qq Test MCP". The display name is what the contact, and its
        # birthday, are known by.
        update(given_name="Gabri")
        assert graph.sent("PATCH") == {"givenName": "Gabri", "displayName": "Gabriele De Simoni"}

    def test_a_given_display_name_wins(self, graph):
        update(display_name="Gabriele")
        assert graph.sent("PATCH") == {"displayName": "Gabriele"}

    def test_merging_the_duplicate_s_address_into_the_keeper(self, graph):
        result = update(email_addresses=["gabriele@example.com"], nickname="Amore di Papà")

        body = graph.sent("PATCH")
        assert body["emailAddresses"] == [{"address": "gabriele@example.com", "name": "gabriele@example.com"}]
        assert body["nickName"] == "Amore di Papà"
        assert body["displayName"] == "Gabriele De Simoni"
        assert "Email: gabriele@example.com" in result

    def test_an_address_the_contact_had_keeps_its_name(self, graph):
        update("C-dup", email_addresses=["GABRIELE@example.com", "new@example.com"])
        assert graph.sent("PATCH")["emailAddresses"] == [
            {"address": "GABRIELE@example.com", "name": "Gabriele"},
            {"address": "new@example.com", "name": "new@example.com"},
        ]

    def test_an_empty_list_removes_every_address(self, graph):
        update("C-dup", email_addresses=[])
        assert graph.sent("PATCH")["emailAddresses"] == []

    def test_a_birthday_is_sent_at_outlook_s_time(self, graph):
        update(birthday="2016-06-07")
        assert graph.sent("PATCH")["birthday"] == "2016-06-07T11:59:00Z"

    def test_an_empty_birthday_removes_it(self, graph):
        update(birthday="")
        assert graph.sent("PATCH")["birthday"] is None

    def test_notes_and_phones(self, graph):
        update(personal_notes="merged from the phone", mobile_phone="", home_phones=["+39 06 1"])
        body = graph.sent("PATCH")
        assert body["personalNotes"] == "merged from the phone"
        assert body["mobilePhone"] == ""
        assert body["homePhones"] == ["+39 06 1"]

    def test_nothing_to_do_is_said_and_nothing_is_asked(self, graph):
        assert update() == "No updates specified."
        assert graph.calls == []

    def test_outlook_s_limits_are_refused_before_graph_sees_them(self):
        with pytest.raises(ValidationError):
            UpdateContactInput(contact_id="x", email_addresses=["a@x", "b@x", "c@x", "d@x"])
        with pytest.raises(ValidationError):
            UpdateContactInput(contact_id="x", home_phones=["1", "2", "3"])
        with pytest.raises(ValidationError):
            UpdateContactInput(contact_id="x", birthday="07/06/2016")


class TestDeleting:
    def test_it_goes_to_deleted_items_and_is_named(self, graph):
        result = delete("C-dup")

        assert ("DELETE", "/me/contacts/C-dup", None) in graph.calls
        assert not [c for c in graph.calls if c[0] == "POST"], "never permanentDelete"
        assert "**Gabriele, Amore di Papà** moved to Deleted Items" in result

    def test_there_is_no_permanent_option(self):
        # Graph's permanentDelete left contacts with a birthday in Deleted Items
        # instead of purging them: an option that says "not recoverable" and
        # does not mean it is worse than none.
        with pytest.raises(ValidationError):
            DeleteContactInput(contact_id="x", permanent=True)

    def test_a_404_for_a_delete_that_happened_is_not_reported_as_a_failure(self, graph):
        graph.delete_answers_404 = True
        result = delete("C-dup")
        assert "moved to Deleted Items" in result

    def test_a_404_for_a_delete_that_did_not_happen_is_reported(self, graph):
        async def delete_nothing(endpoint):
            graph.calls.append(("DELETE", endpoint, None))
            raise not_found("DELETE", endpoint)

        graph.delete = delete_nothing
        assert delete("C-dup").startswith("Error 404")

    def test_an_unknown_id_is_never_deleted(self, graph):
        assert delete("C-nope").startswith("Error 404")
        assert not [c for c in graph.calls if c[0] == "DELETE"]


class TestWhatTheClientSees:
    @staticmethod
    def schema_of(tool_name):
        tools_by_name = {tool.name: tool for tool in anyio.run(mcp.list_tools)}
        schema = tools_by_name[tool_name].input_schema
        params = schema["properties"]["params"]
        if "$ref" in params:
            params = schema["$defs"][params["$ref"].rsplit("/", 1)[-1]]
        return params["properties"]

    def test_the_four_tools_are_registered(self):
        names = {tool.name for tool in anyio.run(mcp.list_tools)}
        assert {
            "outlook_list_contacts", "outlook_get_contact",
            "outlook_update_contact", "outlook_delete_contact",
        } <= names

    def test_update_takes_the_fields_a_merge_needs(self):
        properties = self.schema_of("outlook_update_contact")
        for field in ("display_name", "given_name", "surname", "email_addresses", "birthday", "personal_notes"):
            assert field in properties

    def test_delete_takes_an_id_and_nothing_else(self):
        assert set(self.schema_of("outlook_delete_contact")) == {"contact_id"}
