# GENERATED FILE — do not edit.
#
# Emitted from provider/triggers/calendar-changed.json by weaver's generator.
# A hand-edit here is destroyed by the next protocol sync, which is worse than
# being rejected, because it works until it silently does not. Fix
# provider/triggers/calendar-changed.json (or weaver's template/) and
# regenerate:
#
# npm run provider -- microsoft_outlook

"""Microsoft Outlook's subscription trigger — the delivery contract.

Kept beside the service descriptor rather than inside a node, because the
way a delivery is verified is a fact about MICROSOFT OUTLOOK.

A SUBSCRIPTION: Microsoft Outlook stops delivering unless somebody renews
it, forever, and if nobody does the workflow stops firing with no error
anywhere. LEASE says where the expiry is read from and how early to renew;
the host runs ONE renewal scheduler for every expiring trigger.
"""

from __future__ import annotations

from typing import Any

from .._runtime import Verification, handshake_response, verify_shared_token
from ..faker import respond
from ..service import SERVICE

OPERATION = "calendar_changed"
DELIVERY = "subscription"
SETUP = (
    "The host calls subscription_create with its own URL for this trigger, an expiry of now "
    "plus this lifetime, and the connection's clientState, and must answer Graph's "
    "validationToken challenge on that URL as plain text within 10 seconds or the "
    "subscription is refused. It stores the returned id and expirationDateTime. 3600 seconds "
    "before expiry it calls subscription_renew; a 404 there means the subscription is gone "
    "and it re-lists (event_list) and creates again, because notifications during the gap are "
    "gone. Every notification is a batch: the host compares each item's clientState with the "
    "connection's clientState, answers 202, and then processes the items. A `missed` "
    "lifecycle event means re-list; `reauthorizationRequired` means renew; "
    "`subscriptionRemoved` means create again and re-list."
)
SUBSCRIPTION_TTL = 604800

# The lease this subscription carries: where the provider's expiry sits in the
# subscription_create response, how it is spelled, how early the host renews, and what it
# calls when the lease is due.
LEASE = {
    "expiresAtFrom": "expirationDateTime",
    "expiresAtUnit": "rfc3339",
    "renewBeforeSeconds": 3600,
    "renewOperation": "subscription_renew",
}

# Which of this package's actions create, renew and stop the subscription.
# `renew` is None where the provider cannot renew and the host calls create again.
SUBSCRIPTION = {
    "create": "subscription_create",
    "renew": "subscription_renew",
    "stop": "subscription_delete",
}

# Where Microsoft Outlook echoes the token it was given when the subscription was created:
# a dotted path into the JSON body; `[]` means every element of the batch, all of which must match.
# Neither is a secret: S105 reads any *TOKEN* name as a hardcoded password, and
# these say WHERE the token arrives, not what it is.
TOKEN_IN = "body"  # noqa: S105
TOKEN_NAME = "value[].clientState"  # noqa: S105

# WHICH credential holds the token — a field name, not a secret (S105).
SECRET_CREDENTIAL = "clientState"  # noqa: S105


def verify_delivery(
    raw: str,
    headers: dict[str, str],
    clientstate: str | None,
    now: int | None = None,
) -> Verification:
    """Verify one inbound Microsoft Outlook delivery.
    
    The host calls this BEFORE starting a run, with the body exactly as
    received. The token is the connection's `clientState`; `now` is accepted for
    symmetry with signed schemes and unused, because an echoed token carries no
    timestamp.
    """
    del now

    return verify_shared_token(
        raw=raw,
        headers=headers,
        secret=clientstate,
        placement=TOKEN_IN,
        name=TOKEN_NAME,
    )


# The query parameter Microsoft Outlook sends as its challenge before it delivers anything.
HANDSHAKE_PARAM = "validationToken"


def answer_handshake(query: dict[str, str | list[str]]) -> dict[str, Any] | None:
    """Answer Microsoft Outlook's challenge — what to send back (200, text/plain,
    the decoded `validationToken`), or None when the request is not a challenge
    at all. The host's routing layer asks this before mounting the route.
    """
    return handshake_response(HANDSHAKE_PARAM, query)


def sample_event(config: dict[str, Any] | None = None) -> Any:
    """A faked sample event, so the trigger is runnable before any of the setup
    above.
    
    An author can see the real field names and wire the downstream nodes against
    them before the provider has ever been contacted.
    """
    from .._fake import FakeValues, seed_for_call

    resolved = config or {}
    fake = FakeValues(seed_for_call(SERVICE, OPERATION, resolved))

    return respond(OPERATION, {"config": resolved, "fake": fake})
