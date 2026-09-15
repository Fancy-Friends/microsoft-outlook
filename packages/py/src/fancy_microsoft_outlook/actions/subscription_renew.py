# GENERATED FILE — do not edit.
#
# Emitted from provider/actions/subscription-renew.json by weaver's generator.
# A hand-edit here is destroyed by the next protocol sync, which is worse than
# being rejected, because it works until it silently does not. Fix
# provider/actions/subscription-renew.json (or weaver's template/) and
# regenerate:
#
# npm run provider -- microsoft_outlook

"""Extend a change-notification subscription before it expires. The host's
subscription machinery calls this when the lease is due.

PATCH /v1.0/subscriptions/{id} —
https://learn.microsoft.com/en-us/graph/api/subscription-update?view=graph-rest-1.0

This describes the request. `call` resolves the connection, picks the
estate, and either calls Microsoft Outlook or calls the faker.
"""

from __future__ import annotations

from typing import Any
from urllib.parse import quote

from .._runtime import CallResult, ConnectorConfigError, Mode, call
from ..service import descriptor

OPERATION = "subscription_renew"
METHOD = "PATCH"
PATH = "/v1.0/subscriptions/{id}"
SIDE_EFFECTS = "idempotent"


def body(config: dict[str, Any]) -> dict[str, Any]:
    """Build the JSON body for one call, failing loudly and specifically."""
    if config.get("id") is None or config.get("id") == "":
        raise ConnectorConfigError(
            "subscription_renew: \"id\" is required (Subscription ID)."
        )

    if config.get("expirationDateTime") is None or config.get("expirationDateTime") == "":
        raise ConnectorConfigError(
            "subscription_renew: \"expirationDateTime\" is required (Expires at)."
        )

    out: dict[str, Any] = {}
    _value = config.get("expirationDateTime")
    if _value is None or _value == "":
        raise ConnectorConfigError("subscription_renew: \"expirationDateTime\" is required.")

    out["expirationDateTime"] = str(_value)

    return out



def path(config: dict[str, Any]) -> str:
    """The request path, with each config value URL-ENCODED into it.

    `PATH` above is the TEMPLATE, which is what the descriptor advertises;
    this is what a caller sends. A value interpolated raw changes WHICH URL is
    called — a range like `Sheet1!A:B`, or a sheet named `Q1/Q2` — and the
    provider answers 404 about the document rather than about the encoding.
    """
    return (
        "/v1.0/subscriptions/"
        + quote(str(config.get("id") or ""), safe="")
    )

def subscription_renew(
    config: dict[str, Any],
    *,
    credentials: dict[str, str | None] | None = None,
    mode: Mode = "auto",
    connection_id: str | None = None,
    attempts: int = 3,
) -> CallResult:
    """Extend a change-notification subscription before it expires. The host's subscription
    machinery calls this when the lease is due.
    """
    return call(
        descriptor(),
        operation=OPERATION,
        method=METHOD,
        path=PATH,
        json_body=body(config),
        config=config,
        credentials=credentials,
        mode=mode,
        connection_id=connection_id,
        attempts=attempts,
    )
