# GENERATED FILE — do not edit.
#
# Emitted from provider/actions/subscription-create.json by weaver's
# generator.
# A hand-edit here is destroyed by the next protocol sync, which is worse than
# being rejected, because it works until it silently does not. Fix
# provider/actions/subscription-create.json (or weaver's template/) and
# regenerate:
#
# npm run provider -- microsoft_outlook

"""Subscribe to change notifications on the connected account's calendar
events. The host's subscription machinery calls this; it is not a node most
workflows need.

POST /v1.0/subscriptions —
https://learn.microsoft.com/en-us/graph/api/subscription-post-subscriptions?view=graph-rest-1.0

This describes the request. `call` resolves the connection, picks the
estate, and either calls Microsoft Outlook or calls the faker.
"""

from __future__ import annotations

from typing import Any

from .._runtime import CallResult, ConnectorConfigError, Mode, call
from ..service import descriptor

OPERATION = "subscription_create"
METHOD = "POST"
PATH = "/v1.0/subscriptions"
SIDE_EFFECTS = "unsafe-to-replay"


def body(config: dict[str, Any]) -> dict[str, Any]:
    """Build the JSON body for one call, failing loudly and specifically."""
    if config.get("resource") is None or config.get("resource") == "":
        raise ConnectorConfigError(
            "subscription_create: \"resource\" is required (Resource)."
        )

    if config.get("notificationUrl") is None or config.get("notificationUrl") == "":
        raise ConnectorConfigError(
            "subscription_create: \"notificationUrl\" is required (Notification URL)."
        )

    if config.get("expirationDateTime") is None or config.get("expirationDateTime") == "":
        raise ConnectorConfigError(
            "subscription_create: \"expirationDateTime\" is required (Expires at)."
        )

    if config.get("clientState") is None or config.get("clientState") == "":
        raise ConnectorConfigError(
            "subscription_create: \"clientState\" is required (Client state)."
        )

    out: dict[str, Any] = {}
    _value = config.get("changeType")
    if _value is not None and _value != "":
        out["changeType"] = str(_value)
    _value = config.get("resource")
    if _value is None or _value == "":
        raise ConnectorConfigError("subscription_create: \"resource\" is required.")

    out["resource"] = str(_value)
    _value = config.get("notificationUrl")
    if _value is None or _value == "":
        raise ConnectorConfigError("subscription_create: \"notificationUrl\" is required.")

    out["notificationUrl"] = str(_value)
    _value = config.get("lifecycleNotificationUrl")
    if _value is not None and _value != "":
        out["lifecycleNotificationUrl"] = str(_value)
    _value = config.get("expirationDateTime")
    if _value is None or _value == "":
        raise ConnectorConfigError("subscription_create: \"expirationDateTime\" is required.")

    out["expirationDateTime"] = str(_value)
    _value = config.get("clientState")
    if _value is None or _value == "":
        raise ConnectorConfigError("subscription_create: \"clientState\" is required.")

    out["clientState"] = str(_value)

    return out


def subscription_create(
    config: dict[str, Any],
    *,
    credentials: dict[str, str | None] | None = None,
    mode: Mode = "auto",
    connection_id: str | None = None,
    attempts: int = 3,
) -> CallResult:
    """Subscribe to change notifications on the connected account's calendar events. The host's
    subscription machinery calls this; it is not a node most workflows need.
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
