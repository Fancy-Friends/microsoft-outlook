# GENERATED FILE — do not edit.
#
# Emitted from provider/actions/event-list.json by weaver's generator.
# A hand-edit here is destroyed by the next protocol sync, which is worse than
# being rejected, because it works until it silently does not. Fix
# provider/actions/event-list.json (or weaver's template/) and regenerate:
#
# npm run provider -- microsoft_outlook

"""List events from the connected account's Outlook calendar.

GET /v1.0/me/events —
https://learn.microsoft.com/en-us/graph/api/user-list-events?view=graph-rest-1.0

This describes the request. `call` resolves the connection, picks the
estate, and either calls Microsoft Outlook or calls the faker.
"""

from __future__ import annotations

from typing import Any

from .._runtime import CallResult, ConnectorConfigError, Mode, call
from ..service import descriptor

OPERATION = "event_list"
METHOD = "GET"
PATH = "/v1.0/me/events"
SIDE_EFFECTS = "none"


def body(config: dict[str, Any]) -> dict[str, Any]:
    """Build the form body for one call, failing loudly and specifically."""
    top = config.get("top")
    if top is not None and top != "":
        try:
            _n = float(top)
        except (TypeError, ValueError):
            _n = None
        if _n is None or _n != int(_n) or _n < 1 or _n > 999:
            raise ConnectorConfigError(
                "event_list: \"top\" must be a integer, got "
                f"{top!r}."
            )

    out: dict[str, Any] = {}
    _value = config.get("top")
    out["$top"] = int(float(_value)) if _value is not None and _value != "" else 50
    _value = config.get("filter")
    if _value is not None and _value != "":
        out["$filter"] = str(_value)
    _value = config.get("orderby")
    if _value is not None and _value != "":
        out["$orderby"] = str(_value)
    _value = config.get("select")
    if _value is not None and _value != "":
        out["$select"] = str(_value)

    return out


def event_list(
    config: dict[str, Any],
    *,
    credentials: dict[str, str | None] | None = None,
    mode: Mode = "auto",
    connection_id: str | None = None,
    attempts: int = 3,
) -> CallResult:
    """List events from the connected account's Outlook calendar."""
    return call(
        descriptor(),
        operation=OPERATION,
        method=METHOD,
        path=PATH,
        form=body(config),
        config=config,
        credentials=credentials,
        mode=mode,
        connection_id=connection_id,
        attempts=attempts,
    )
