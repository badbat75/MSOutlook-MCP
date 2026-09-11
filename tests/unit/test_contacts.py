"""Unit tests for the contact tools and the address book they read.

The request behind these tools: the Birthdays calendar showed three entries on
7 June for one child, and that calendar gives only the day and the month. The
contacts behind them had to be found, compared and merged, which Graph does not
make easy: /me/contacts is the default folder only, a name cannot be searched
for with $filter, an update that leaves out displayName may regenerate it, and a
delete has been seen answering 404 while carrying itself out. These tests pin
what the tools do about each, against a fake Graph shaped like the real mailbox.

The folder tools came with the next request: fold the folder a phone synced
into back into the default one. Graph cannot move a contact, only copy it and
delete the original, and a copy carrying a photo was seen to vanish once its
original was deleted; the move tests pin the copy, the checks around it and
the refusals.
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
    ContactFolder,
    birthday_date,
    birthday_value,
    copy_of,
    find_folder,
    lost_fields,
    matches,
    month_day,
    sort_key,
)
from outlook_mcp.models import (
    CreateContactInput,
    DeleteContactFolderInput,
    DeleteContactInput,
    GetContactInput,
    ListContactsInput,
    MoveContactsInput,
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


def server_error(method, url):
    request = httpx.Request(method, f"https://graph.microsoft.com/v1.0{url}")
    response = httpx.Response(500, request=request, json={"error": {"code": "boom", "message": "boom"}})
    return httpx.HTTPStatusError("500", request=request, response=response)


class FakeGraph:
    """A mailbox with a default folder, a phone's folder in it, and one nested deeper."""

    def __init__(self):
        self.folders = {
            "F-default": ("Contacts", [KEEPER, DUPLICATE, NAMESAKE, PHONE_ONLY]),
            "F-phone": ("HUAWEI P40 Pro", [ON_PHONE, ON_PHONE_PAGE_2]),
            "F-nested": ("Family", [NESTED]),
        }
        self.children = {"F-default": ["F-phone"], "F-phone": ["F-nested"], "F-nested": []}
        # Items Graph counts in a folder ($count) without listing them as contacts.
        self.hidden = {}
        self.photos = {}
        # Properties a create leaves out of the contact it makes.
        self.drop_on_create = set()
        self.deleted_items = []
        self.deleted_folders = []
        self.delete_answers_404 = False
        self.created = 0
        self.calls = []

    def _locate(self, cid):
        for folder, (_, items) in self.folders.items():
            for item in items:
                if item["id"] == cid:
                    return folder, item
        return None, None

    def _contact(self, cid):
        return self._locate(cid)[1]

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
        if endpoint == NEXT_LINK:
            return {"value": list(self.folders["F-phone"][1][1:])}
        if endpoint.startswith("/me/contactFolders/") and endpoint.endswith("/contacts"):
            folder = endpoint.split("/")[3]
            items = list(self.folders[folder][1])
            if (params or {}).get("$count") == "true":
                return {"value": items[:1], "@odata.count": len(items) + self.hidden.get(folder, 0)}
            if folder == "F-phone" and len(items) > 1:
                # Two pages, the second behind an absolute @odata.nextLink.
                return {"value": items[:1], "@odata.nextLink": NEXT_LINK}
            return {"value": items}
        if endpoint.startswith("/me/contactFolders/"):
            folder = self.folders.get(endpoint.split("/")[3])
            if folder is None:
                raise not_found("GET", endpoint)
            return {"displayName": folder[0]}
        if endpoint.startswith("/me/contacts/"):
            folder, found = self._locate(endpoint.split("/")[3])
            if found is None:
                raise not_found("GET", endpoint)
            return {"parentFolderId": folder, **found}
        raise AssertionError(f"unexpected GET {endpoint}")

    async def get_bytes(self, endpoint):
        self.calls.append(("GET BYTES", endpoint, None))
        return self.photos.get(endpoint.split("/")[3])

    async def patch(self, endpoint, json_data=None):
        self.calls.append(("PATCH", endpoint, json_data))
        return {**self._contact(endpoint.split("/")[3]), **json_data}

    async def delete(self, endpoint):
        self.calls.append(("DELETE", endpoint, None))
        if endpoint.startswith("/me/contactFolders/"):
            folder = endpoint.split("/")[3]
            self.deleted_folders.append(folder)
            for children in self.children.values():
                if folder in children:
                    children.remove(folder)
            return {"status": "success"}
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
        if endpoint.startswith("/me/contactFolders/") and endpoint.endswith("/contacts"):
            folder = endpoint.split("/")[3]
            self.created += 1
            kept = {k: v for k, v in json_data.items() if k not in self.drop_on_create}
            item = {
                **kept, "id": f"C-new{self.created}", "parentFolderId": folder,
                "createdDateTime": "2026-09-11T18:00:00Z",
            }
            self.folders[folder][1].append(item)
            return dict(item)
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


def create(**fields):
    return anyio.run(tools.outlook_create_contact, CreateContactInput(**fields), None)


def list_folders():
    return anyio.run(tools.outlook_list_contact_folders, None)


def delete_folder(folder):
    return anyio.run(tools.outlook_delete_contact_folder, DeleteContactFolderInput(folder=folder), None)


def move(ids, destination):
    return anyio.run(
        tools.outlook_move_contacts,
        MoveContactsInput(contact_ids=ids, destination_folder=destination),
        None,
    )


def posted(graph):
    return [(call[1], call[2]) for call in graph.calls if call[0] == "POST"]


def deleted(graph):
    return [call[1] for call in graph.calls if call[0] == "DELETE"]


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

    def test_the_company(self, graph):
        update(company_name="ACME")
        assert graph.sent("PATCH") == {"companyName": "ACME", "displayName": "Gabriele De Simoni"}

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


PHONE_FOLDER = "HUAWEI P40 Pro (contacts synced by Link to Windows)"
FOLDERS = [
    ContactFolder("F-default", "Contacts"),
    ContactFolder("F-phone", PHONE_FOLDER, 1, "F-default"),
    ContactFolder("F-nested", "Famiglia Città", 2, "F-phone"),
]


class TestFindingAFolder:
    def test_the_default_folder_by_its_well_known_name(self):
        assert find_folder(FOLDERS, "contacts") == FOLDERS[0]
        assert find_folder(FOLDERS, " CONTACTS ") == FOLDERS[0]

    def test_a_name_ignoring_case_and_accents(self):
        assert find_folder(FOLDERS, "famiglia citta").id == "F-nested"
        assert find_folder(FOLDERS, PHONE_FOLDER.lower()).id == "F-phone"

    def test_an_id(self):
        assert find_folder(FOLDERS, "F-phone").name == PHONE_FOLDER

    def test_a_fragment_is_not_a_name_and_the_folders_are_named(self):
        # Contacts are moved into a folder found this way, and folders deleted.
        with pytest.raises(ValueError, match="The contact folders are: 'Contacts', 'HUAWEI"):
            find_folder(FOLDERS, "HUAWEI")

    def test_two_folders_of_one_name_need_an_id(self):
        twins = FOLDERS + [ContactFolder("F-twin", "Famiglia Città", 1, "F-default")]
        with pytest.raises(ValueError, match="by its ID"):
            find_folder(twins, "Famiglia Città")


class TestListingFolders:
    def test_nested_with_counts_and_ids(self, graph):
        result = list_folders()

        assert "3, 7 contacts in all" in result
        assert "- **Contacts** (default): 4 contacts | ID: `F-default`" in result
        # The phone's folder counted across both of its pages.
        assert "\n  - **HUAWEI P40 Pro**: 2 contacts | ID: `F-phone`" in result
        assert "\n    - **Family**: 1 contact | ID: `F-nested`" in result

    def test_items_graph_counts_but_does_not_list_are_shown(self, graph):
        # Measured: the real default folder counts 164 items and lists 162.
        graph.hidden["F-default"] = 2
        assert "4 contacts and 2 other items not listed as contacts" in list_folders()


class TestListingOneFolder:
    def test_only_that_folder_not_the_ones_inside_it(self, graph):
        result = list_contacts(folder="huawei p40 pro", top=200)

        assert "2 in HUAWEI P40 Pro" in result
        assert "Papà" in result and "Zoe" in result
        assert "Nicolò" not in result, "Family is inside it, and not asked for"
        assert "Gabriele" not in result

    def test_the_default_folder_by_its_well_known_name(self, graph):
        result = list_contacts(folder="contacts", top=200)
        assert "4 in Contacts" in result
        assert "C-phone1" not in result and "C-nested" not in result

    def test_nothing_found_names_the_folder(self, graph):
        assert list_contacts(folder="Family", search="zoe") == "No contacts matching 'zoe' in Family."

    def test_an_unknown_folder_says_which_there_are(self, graph):
        result = list_contacts(folder="Nope")
        assert result.startswith("Error")
        assert "'Contacts', 'HUAWEI P40 Pro', 'Family'" in result


class TestCreating:
    def test_into_the_default_folder(self, graph):
        result = create(
            given_name="Zoe", surname="Rossi", company_name="ACME", mobile_phone="+39 1",
            email_addresses=["zoe@example.com"], birthday="2016-06-07",
        )

        assert posted(graph) == [("/me/contactFolders/F-default/contacts", {
            "givenName": "Zoe", "surname": "Rossi", "companyName": "ACME", "mobilePhone": "+39 1",
            "birthday": "2016-06-07T11:59:00Z",
            "emailAddresses": [{"address": "zoe@example.com", "name": "zoe@example.com"}],
        })]
        assert "Contact created in **Contacts**" in result
        assert "Folder: Contacts" in result
        assert "ID: `C-new1`" in result

    def test_into_a_folder_named_by_its_name(self, graph):
        result = create(folder="family", display_name="Nonna")
        assert posted(graph) == [("/me/contactFolders/F-nested/contacts", {"displayName": "Nonna"})]
        assert "Contact created in **Family**" in result

    def test_the_display_name_is_left_to_outlook_unless_given(self, graph):
        create(given_name="Zoe", surname="Rossi")
        assert "displayName" not in posted(graph)[0][1]

    def test_an_unknown_folder_creates_nothing(self, graph):
        assert create(folder="Nope", given_name="Zoe").startswith("Error")
        assert posted(graph) == []

    def test_a_contact_needs_something_to_know_it_by(self):
        with pytest.raises(ValidationError, match="at least a name"):
            CreateContactInput()
        with pytest.raises(ValidationError, match="at least a name"):
            CreateContactInput(birthday="2016-06-07", personal_notes="who?")
        assert CreateContactInput(mobile_phone="+39 1").mobile_phone == "+39 1"

    def test_outlook_s_limits_are_refused_before_graph_sees_them(self):
        with pytest.raises(ValidationError):
            CreateContactInput(given_name="x", email_addresses=["a@x", "b@x", "c@x", "d@x"])
        with pytest.raises(ValidationError):
            CreateContactInput(given_name="x", business_phones=["1", "2", "3"])
        with pytest.raises(ValidationError):
            CreateContactInput(given_name="x", birthday="07/06/2016")

    def test_an_empty_birthday_is_no_birthday(self, graph):
        create(given_name="Zoe", birthday="")
        assert "birthday" not in posted(graph)[0][1]


class TestDeletingFolders:
    def test_an_empty_folder_goes_to_deleted_items(self, graph):
        graph.folders["F-nested"] = ("Family", [])
        result = delete_folder("family")

        assert ("DELETE", "/me/contactFolders/F-nested", None) in graph.calls
        assert "Contact folder **Family** moved to Deleted Items" in result

    def test_a_folder_with_contacts_is_left_alone(self, graph):
        result = delete_folder("Family")
        assert result.startswith("Not deleted") and "1 contact" in result
        assert graph.deleted_folders == []

    def test_items_that_are_not_contacts_count_too(self, graph):
        graph.folders["F-nested"] = ("Family", [])
        graph.hidden["F-nested"] = 1
        result = delete_folder("Family")
        assert result.startswith("Not deleted") and "1 other item" in result
        assert graph.deleted_folders == []

    def test_a_folder_with_subfolders_is_left_alone(self, graph):
        graph.folders["F-phone"] = ("HUAWEI P40 Pro", [])
        result = delete_folder("HUAWEI P40 Pro")
        assert result.startswith("Not deleted") and "has subfolders (Family)" in result
        assert graph.deleted_folders == []

    def test_the_default_folder_is_never_deleted(self, graph):
        assert "default contact folder" in delete_folder("contacts")
        assert graph.deleted_folders == []

    def test_an_unknown_folder_is_an_error(self, graph):
        assert delete_folder("Nope").startswith("Error")
        assert graph.deleted_folders == []


# A contact the way the phone's folder holds it on the real mailbox, with what
# Graph adds to it on reading: the sync's category, an address of empty
# strings, the personal-account echo of the first email address, and the
# properties only Graph sets.
SYNCED = contact(
    "C-phone1", "Papà", givenName="Papà", mobilePhone="+39 333 000", homePhones=["+39 06 000"],
    categories=[PHONE_FOLDER], birthday="1950-01-02T11:59:00Z",
    emailAddresses=[{"name": "Papà", "address": "papa@example.com"}],
    primaryEmailAddress={"name": "Papà", "address": "papa@example.com"},
    homeAddress={"street": "", "city": "Roma", "state": "", "countryOrRegion": "", "postalCode": ""},
    changeKey="EQAAAB", createdDateTime="2024-03-01T10:00:00Z",
    lastModifiedDateTime="2024-03-02T10:00:00Z", **{"@odata.etag": 'W/"EQAAAB"'},
)


class TestCopying:
    def test_what_graph_sets_itself_stays_behind(self):
        assert copy_of({**SYNCED, "parentFolderId": "F-phone"}) == {
            "displayName": "Papà", "givenName": "Papà", "mobilePhone": "+39 333 000",
            "homePhones": ["+39 06 000"], "categories": [PHONE_FOLDER],
            "birthday": "1950-01-02T11:59:00Z",
            "emailAddresses": [{"name": "Papà", "address": "papa@example.com"}],
            "homeAddress": {"street": "", "city": "Roma", "state": "", "countryOrRegion": "", "postalCode": ""},
        }

    def test_a_faithful_copy_lost_nothing(self):
        # A new item has its own ID, change key and dates: none of that is loss.
        assert lost_fields(SYNCED, {**copy_of(SYNCED), "id": "C-new", "changeKey": "x"}) == []

    def test_a_value_missing_or_changed_is_lost(self):
        copy = {**copy_of(SYNCED), "mobilePhone": "+39 000"}
        del copy["categories"]
        assert lost_fields(SYNCED, copy) == ["categories", "mobilePhone"]

    def test_an_empty_property_has_nothing_to_lose(self):
        empty = {"homeAddress": {"street": "", "city": ""}, "nickName": "", "children": [], "birthday": None}
        assert lost_fields(empty, {}) == []


class TestMoving:
    @pytest.fixture
    def phone(self, graph):
        graph.folders["F-phone"][1][0] = dict(SYNCED)
        return graph

    def test_out_of_the_phone_folder_into_the_default_one(self, phone):
        result = move(["C-phone1"], "contacts")

        assert posted(phone) == [("/me/contactFolders/F-default/contacts", copy_of(SYNCED))]
        assert deleted(phone) == ["/me/contacts/C-phone1"]
        assert "Moved 1 of 1** to **Contacts**" in result
        assert "✅ **Papà**: new ID `C-new1`" in result
        assert "originals are in Deleted Items" in result
        assert [c["id"] for c in phone.folders["F-phone"][1]] == ["C-phone2"]
        assert phone.folders["F-default"][1][-1]["categories"] == [PHONE_FOLDER]

    def test_the_copy_is_made_before_the_original_goes(self, phone):
        move(["C-phone1"], "contacts")
        writes = [call[0] for call in phone.calls if call[0] in ("POST", "DELETE")]
        assert writes == ["POST", "DELETE"]

    def test_several_in_one_call(self, phone):
        result = move(["C-phone1", "C-phone2"], "Contacts")
        assert "Moved 2 of 2" in result
        assert phone.folders["F-phone"][1] == []

    def test_one_already_there_is_left_alone(self, graph):
        result = move(["C-keeper"], "contacts")
        assert "⏭️ **Gabriele De Simoni**: already in Contacts" in result
        assert "Moved 0 of 1" in result
        assert posted(graph) == [] and deleted(graph) == []

    def test_a_contact_with_a_photo_is_left_where_it_is(self, phone):
        # Measured: a copy carrying the same photo vanished once the original
        # was deleted, from every folder and from Deleted Items.
        phone.photos["C-phone1"] = (b"\xff\xd8", "image/jpeg")
        result = move(["C-phone1"], "contacts")
        assert "not moved, it has a photo" in result
        assert posted(phone) == [] and deleted(phone) == []

    def test_an_unknown_id_is_reported_and_the_rest_still_move(self, phone):
        result = move(["C-nope", "C-phone2"], "contacts")
        assert "❌ `C-nope`: no such contact" in result
        assert "Moved 1 of 2" in result

    def test_a_copy_missing_something_is_undone_and_the_original_kept(self, phone):
        phone.drop_on_create = {"homePhones"}
        result = move(["C-phone1"], "contacts")

        assert "the copy came back without homePhones" in result
        assert deleted(phone) == ["/me/contacts/C-new1"]
        assert phone._contact("C-phone1") is not None
        assert "Moved 0 of 1" in result

    def test_a_404_for_a_delete_that_happened_still_moves(self, phone):
        phone.delete_answers_404 = True
        assert "Moved 1 of 1" in move(["C-phone1"], "contacts")

    def test_an_original_that_cannot_be_deleted_stops_the_call(self, phone):
        async def refuse(endpoint):
            phone.calls.append(("DELETE", endpoint, None))
            raise server_error("DELETE", endpoint)

        phone.delete = refuse
        result = move(["C-phone1", "C-phone2"], "contacts")

        assert "copied into Contacts (new ID `C-new1`)" in result
        assert "delete the original, `C-phone1`" in result
        assert "Not attempted: `C-phone2`" in result
        assert "Moved 0 of 2" in result

    def test_an_error_before_the_copy_stops_the_call_with_nothing_changed(self, phone):
        async def fail(endpoint, json_data=None):
            raise server_error("POST", endpoint)

        phone.post = fail
        result = move(["C-phone1", "C-phone2"], "contacts")

        assert "❌ `C-phone1`: Error 500" in result
        assert "Not attempted: `C-phone2`" in result
        assert deleted(phone) == []

    def test_an_unknown_destination_moves_nothing(self, phone):
        assert move(["C-phone1"], "Nope").startswith("Error")
        assert posted(phone) == []

    def test_a_call_takes_one_to_twenty_five(self):
        with pytest.raises(ValidationError):
            MoveContactsInput(contact_ids=[], destination_folder="contacts")
        with pytest.raises(ValidationError):
            MoveContactsInput(contact_ids=[f"C{i}" for i in range(26)], destination_folder="contacts")
        with pytest.raises(ValidationError):
            MoveContactsInput(contact_ids=["C1", ""], destination_folder="contacts")
        assert len(MoveContactsInput(contact_ids=[f"C{i}" for i in range(25)], destination_folder="x").contact_ids) == 25


class TestWhatTheClientSees:
    @staticmethod
    def tool(tool_name):
        return {tool.name: tool for tool in anyio.run(mcp.list_tools)}[tool_name]

    @classmethod
    def schema_of(cls, tool_name):
        schema = cls.tool(tool_name).input_schema
        params = schema["properties"]["params"]
        if "$ref" in params:
            params = schema["$defs"][params["$ref"].rsplit("/", 1)[-1]]
        return params["properties"]

    def test_the_contact_tools_are_registered(self):
        names = {tool.name for tool in anyio.run(mcp.list_tools)}
        assert {
            "outlook_list_contacts", "outlook_get_contact", "outlook_create_contact",
            "outlook_update_contact", "outlook_delete_contact",
            "outlook_list_contact_folders", "outlook_move_contacts",
            "outlook_delete_contact_folder",
        } <= names

    def test_a_move_takes_the_contacts_and_where_to(self):
        assert set(self.schema_of("outlook_move_contacts")) == {"contact_ids", "destination_folder"}

    def test_listing_folders_takes_nothing(self):
        assert not self.tool("outlook_list_contact_folders").input_schema.get("properties")

    def test_create_takes_a_folder_and_the_fields_of_an_update(self):
        created = set(self.schema_of("outlook_create_contact"))
        updated = set(self.schema_of("outlook_update_contact"))
        assert created - updated == {"folder"}
        assert updated - created == {"contact_id"}

    def test_deleting_a_folder_takes_the_folder_and_nothing_else(self):
        assert set(self.schema_of("outlook_delete_contact_folder")) == {"folder"}

    def test_update_takes_the_fields_a_merge_needs(self):
        properties = self.schema_of("outlook_update_contact")
        for field in ("display_name", "given_name", "surname", "email_addresses", "birthday", "personal_notes"):
            assert field in properties

    def test_delete_takes_an_id_and_nothing_else(self):
        assert set(self.schema_of("outlook_delete_contact")) == {"contact_id"}
