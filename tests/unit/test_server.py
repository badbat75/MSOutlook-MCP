"""Unit tests for the entry point's transport security settings.

The rest of server.py is argument parsing and a call to mcp.run(). What is worth
pinning is _transport_security(), because getting it wrong is invisible until a
request arrives through the proxy and comes back 421 with no explanation.
"""

from outlook_mcp.config import ServerConfig
from outlook_mcp.server import _transport_security

PROXIED = ServerConfig(
    transport="http", bind_host="127.0.0.1", bind_port=3015,
    allowed_hosts=("ai.example.com",),
)


class TestNoAllowedHosts:
    def test_nothing_configured_means_no_override(self):
        # None is not "no protection": it is "say nothing", so the SDK applies
        # its own rule. A loopback bind then accepts localhost only, which is
        # right for a server nobody proxies.
        assert _transport_security(ServerConfig()) is None
        assert _transport_security(ServerConfig(transport="http")) is None


class TestProxiedHosts:
    def test_the_protection_stays_on(self):
        # The fix widens the allowed list; it must never disable the check,
        # which would leave a loopback server answering any Host at all.
        assert _transport_security(PROXIED).enable_dns_rebinding_protection is True

    def test_the_proxy_hostname_is_accepted(self):
        assert "ai.example.com" in _transport_security(PROXIED).allowed_hosts

    def test_the_hostname_is_accepted_with_a_port_too(self):
        # nginx sends the bare name with `proxy_set_header Host $host` and the
        # name plus port with $http_host. Both are the same deployment, so a
        # change of one line in the proxy must not take the server down.
        assert "ai.example.com:*" in _transport_security(PROXIED).allowed_hosts

    def test_loopback_still_works(self):
        # Health checks and anything else on the host itself talk to the port
        # directly, and they must keep working after the proxy is configured.
        hosts = _transport_security(PROXIED).allowed_hosts
        assert {"127.0.0.1:*", "localhost:*", "[::1]:*"} <= set(hosts)

    def test_an_unlisted_host_is_not_in_the_list(self):
        assert "evil.example.com" not in _transport_security(PROXIED).allowed_hosts

    def test_origins_cover_the_proxy_over_tls_and_loopback(self):
        origins = _transport_security(PROXIED).allowed_origins
        assert "https://ai.example.com" in origins
        assert "http://127.0.0.1:*" in origins

    def test_every_configured_host_is_covered(self):
        config = ServerConfig(
            transport="http", allowed_hosts=("a.example.com", "b.example.com")
        )
        hosts = _transport_security(config).allowed_hosts
        for name in ("a.example.com", "b.example.com"):
            assert name in hosts
            assert f"{name}:*" in hosts
