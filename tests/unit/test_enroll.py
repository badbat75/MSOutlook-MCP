"""Unit tests for the enrollment flow table in outlook_mcp.enroll.

The routes themselves need a browser and Entra, but the bookkeeping between the
two of them is pure logic, and it is where the mix-ups would be: a sign-in that
can be replayed, or one user finishing another's and having their tokens filed
under the wrong name. The pages they render are checked here too, because they
echo values that come from outside the code.
"""

import anyio
import pytest

from outlook_mcp import enroll


@pytest.fixture(autouse=True)
def empty_table():
    enroll._pending.clear()
    yield
    enroll._pending.clear()


@pytest.fixture
def clock(monkeypatch):
    """A monotonic clock the test moves by hand."""
    now = {"t": 1000.0}
    monkeypatch.setattr(enroll.time, "monotonic", lambda: now["t"])
    return now


def start(user, state="state-1"):
    flow = {"state": state, "auth_uri": "https://login.example/authorize"}
    enroll._remember(flow, user)
    return flow


class TestTakingAFlow:
    def test_the_user_who_started_it_gets_it_back(self):
        flow = start("ada@example.com")
        assert enroll._take("state-1", "ada@example.com") is flow

    def test_it_can_only_be_taken_once(self):
        # The callback consumes the flow, so a replayed callback URL cannot
        # redeem the same authorization code a second time.
        start("ada@example.com")
        enroll._take("state-1", "ada@example.com")
        assert enroll._take("state-1", "ada@example.com") is None

    def test_another_user_cannot_finish_it(self):
        # Otherwise bob could complete ada's sign-in and have ada's tokens
        # written into bob's cache, which is the one mix-up this mode exists
        # to prevent.
        start("ada@example.com")
        assert enroll._take("state-1", "bob@example.com") is None

    def test_a_failed_take_still_consumes_the_flow(self):
        start("ada@example.com")
        enroll._take("state-1", "bob@example.com")
        assert enroll._take("state-1", "ada@example.com") is None

    @pytest.mark.parametrize("variant", ["ADA@example.com", " ada@example.com "])
    def test_casing_and_padding_from_the_proxy_are_tolerated(self, variant):
        flow = start("ada@example.com")
        assert enroll._take("state-1", variant) is flow

    def test_an_unknown_state_is_not_found(self):
        start("ada@example.com")
        assert enroll._take("state-other", "ada@example.com") is None

    def test_an_expired_flow_is_not_found(self, clock):
        start("ada@example.com")
        clock["t"] += enroll.FLOW_TTL_SECONDS + 1
        assert enroll._take("state-1", "ada@example.com") is None

    def test_a_flow_inside_the_window_is_still_good(self, clock):
        flow = start("ada@example.com")
        clock["t"] += enroll.FLOW_TTL_SECONDS - 1
        assert enroll._take("state-1", "ada@example.com") is flow


class TestTheTableIsBounded:
    def test_expired_entries_are_pruned_on_the_next_start(self, clock):
        start("ada@example.com", "old")
        clock["t"] += enroll.FLOW_TTL_SECONDS + 1
        start("bob@example.com", "new")
        assert list(enroll._pending) == ["new"]

    def test_it_never_grows_past_the_limit(self):
        # Anyone the proxy lets through can start a sign-in, so an abandoned
        # tab must not be able to fill memory.
        for i in range(enroll.MAX_PENDING_FLOWS + 10):
            start("ada@example.com", f"state-{i}")
        assert len(enroll._pending) <= enroll.MAX_PENDING_FLOWS

    def test_the_oldest_is_dropped_first(self, clock):
        for i in range(enroll.MAX_PENDING_FLOWS):
            start("ada@example.com", f"state-{i}")
            clock["t"] += 1
        start("ada@example.com", "newest")
        assert "state-0" not in enroll._pending
        assert "newest" in enroll._pending


class _Request:
    """The two things the routes read off a Starlette request."""

    def __init__(self, headers, query_params):
        self.headers = headers
        self.query_params = query_params


class _Config:
    enrollment_enabled = True
    is_http = True
    has_credentials = True
    user_header = "X-Auth-Email"
    client_id = "client-id"
    client_secret = "secret"
    tenant_id = "common"
    cache_directory = None  # replaced per test, so nothing is written under $HOME


class _MsalApp:
    def __init__(self, result):
        self._result = result

    def acquire_token_by_auth_code_flow(self, flow, params):
        return self._result


@pytest.fixture
def enrolling(monkeypatch, tmp_path):
    """Everything around the callback stubbed, so only its rendering is tested."""

    def configure(result):
        config = _Config()
        config.cache_directory = tmp_path
        monkeypatch.setattr(enroll, "get_config", lambda: config)
        monkeypatch.setattr(
            enroll, "_msal_app", lambda creds, cache=None: _MsalApp(result)
        )
        monkeypatch.setattr(enroll, "save_token_cache", lambda cache, path: None)

        def finish(user, state="state-1"):
            request = _Request({"x-auth-email": user}, {"state": state})
            return anyio.run(enroll.enroll_callback, request)

        return finish

    return configure


class TestPagesEscapeUntrustedValues:
    """These pages echo three values that do not come from the code.

    The address the proxy asserted is the one that matters: a proxy location
    forwarding without replacing the identity header hands it to the client
    whole, which is the misconfiguration SETUP.md warns about. Reflecting it
    unescaped would turn that mistake into script execution on the origin that
    holds the proxy's own session cookie.
    """

    def test_text_neutralises_markup(self):
        assert enroll._text("<script>alert(1)</script>") == (
            "&lt;script&gt;alert(1)&lt;/script&gt;"
        )

    def test_the_title_is_escaped(self):
        body = enroll._page("<b>hello</b>", "<p>ok</p>").body.decode()
        assert "<b>hello</b>" not in body
        assert "&lt;b&gt;hello&lt;/b&gt;" in body

    def test_the_body_stays_the_callers_markup(self):
        # _page cannot escape the body, the callers build it. What it must not
        # do is mangle it, or every page would render its own tags as text.
        assert "<p>ok</p>" in enroll._page("Title", "<p>ok</p>").body.decode()

    def test_the_authorized_page_escapes_the_asserted_address(self, enrolling):
        finish = enrolling({"access_token": "a-token"})
        evil = 'ada@example.com<img src=x onerror="alert(1)">'
        start(evil)

        response = finish(evil)

        body = response.body.decode()
        assert response.status_code == 200
        assert "<img" not in body
        assert "&lt;img" in body

    def test_a_failed_authorization_escapes_entras_message(self, enrolling):
        finish = enrolling(
            {"error": "invalid_grant", "error_description": "<script>x</script>"}
        )
        start("ada@example.com")

        response = finish("ada@example.com")

        body = response.body.decode()
        assert response.status_code == 400
        assert "<script>x</script>" not in body
        assert "&lt;script&gt;" in body
