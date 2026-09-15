# GENERATED FILE — do not edit.
#
# Emitted from provider/fixtures/ by weaver's generator.
# A hand-edit here is destroyed by the next protocol sync, which is worse than
# being rejected, because it works until it silently does not. Fix
# provider/fixtures/ (or weaver's template/) and regenerate:
#
# npm run provider -- microsoft_outlook

"""The Microsoft Outlook faker.

Bit-for-bit identical to the TypeScript and PHP fakers: the same FNV-1a seed
and the same xorshift32 sequence, so a golden fixture asserts the exact
faked payload and ALL THREE runtimes have to produce it. That turns the
faker into a parity test rather than a convenience — which matters, because
cross-runtime drift does not fail loudly. It completes, down one path, with
no error.
"""

from __future__ import annotations

from typing import Any

from ._fake import FakeValues


def _event_get(config: dict[str, Any], fake: FakeValues) -> Any:
    bound_id = (
        str(_v)
        if (_v := config.get("id")) is not None and _v != ""
        else fake.id("AAMkAGI2")
    )

    return {
        "@odata.context": (
                              "https://graph.microsoft.com/v1.0/$metadata#users('ada%40examp"
                              "le.test')/events/$entity"
                          ),
        "@odata.etag": "W/\"ZlnW4RIAV06KYYwlrfNZvQAALfZeRQ==\"",
        "id": bound_id,
        "createdDateTime": "2026-09-01T09:00:00.0000000Z",
        "lastModifiedDateTime": "2026-09-10T15:30:00.0000000Z",
        "subject": "Design review",
        "bodyPreview": "Quarterly design review with the platform team.",
        "isCancelled": False,
        "isAllDay": False,
        "showAs": "busy",
        "webLink": (
                       "https://outlook.office365.com/owa/?itemid=AAMkAGI2&exvsurl=1&path=/c"
                       "alendar/item"
                   ),
        "start": {
            "dateTime": "2026-09-22T14:00:00.0000000",
            "timeZone": "UTC",
        },
        "end": {
            "dateTime": "2026-09-22T15:00:00.0000000",
            "timeZone": "UTC",
        },
        "location": {
            "displayName": "Room 4",
        },
        "organizer": {
            "emailAddress": {
                "name": "Ada Example",
                "address": "ada@example.test",
            },
        },
        "attendees": [
            {
                "type": "required",
                "status": {
                    "response": "accepted",
                    "time": "2026-09-02T10:00:00.0000000Z",
                },
                "emailAddress": {
                    "name": "Grace Example",
                    "address": "grace@example.test",
                },
            },
        ],
    }


def _event_list(config: dict[str, Any], fake: FakeValues) -> Any:
    return {
        "@odata.context": (
                              "https://graph.microsoft.com/v1.0/$metadata#users('ada%40examp"
                              "le.test')/events"
                          ),
        "value": [
            {
                "@odata.etag": "W/\"ZlnW4RIAV06KYYwlrfNZvQAALfZeRQ==\"",
                "id": fake.id("AAMkAGI2"),
                "createdDateTime": "2026-09-01T09:00:00.0000000Z",
                "lastModifiedDateTime": "2026-09-10T15:30:00.0000000Z",
                "subject": "Design review",
                "bodyPreview": "Quarterly design review with the platform team.",
                "isCancelled": False,
                "isAllDay": False,
                "showAs": "busy",
                "webLink": (
                               "https://outlook.office365.com/owa/?itemid=AAMkAGI2&exvsurl=1"
                               "&path=/calendar/item"
                           ),
                "start": {
                    "dateTime": "2026-09-22T14:00:00.0000000",
                    "timeZone": "UTC",
                },
                "end": {
                    "dateTime": "2026-09-22T15:00:00.0000000",
                    "timeZone": "UTC",
                },
                "organizer": {
                    "emailAddress": {
                        "name": "Ada Example",
                        "address": "ada@example.test",
                    },
                },
            },
        ],
    }


def _subscription_create(config: dict[str, Any], fake: FakeValues) -> Any:
    bound_resource = (
        str(_v)
        if (_v := config.get("resource")) is not None and _v != ""
        else "me/events"
    )
    bound_changetype = (
        str(_v)
        if (_v := config.get("changeType")) is not None and _v != ""
        else "created,updated,deleted"
    )
    bound_notificationurl = (
        str(_v)
        if (_v := config.get("notificationUrl")) is not None and _v != ""
        else "https://host.example.test/hooks/microsoft-outlook"
    )
    bound_clientstate = (
        str(_v)
        if (_v := config.get("clientState")) is not None and _v != ""
        else fake.id("cs")
    )

    return {
        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#subscriptions/$entity",
        "id": "-".join([fake.hex(8), fake.hex(4), fake.hex(4), fake.hex(4), fake.hex(12)]),
        "resource": bound_resource,
        "applicationId": "-".join(
            [fake.hex(8), fake.hex(4), fake.hex(4), fake.hex(4), fake.hex(12)]
        ),
        "changeType": bound_changetype,
        "clientState": bound_clientstate,
        "notificationUrl": bound_notificationurl,
        "expirationDateTime": "2026-09-22T18:23:45.9356913Z",
        "creatorId": "-".join(
            [fake.hex(8), fake.hex(4), fake.hex(4), fake.hex(4), fake.hex(12)]
        ),
        "latestSupportedTlsVersion": "v1_2",
        "notificationContentType": "application/json",
    }


def _subscription_delete(config: dict[str, Any], fake: FakeValues) -> Any:
    return {}


def _subscription_renew(config: dict[str, Any], fake: FakeValues) -> Any:
    bound_id = (
        str(_v)
        if (_v := config.get("id")) is not None and _v != ""
        else "-".join([fake.hex(8), fake.hex(4), fake.hex(4), fake.hex(4), fake.hex(12)])
    )
    bound_expirationdatetime = (
        str(_v)
        if (_v := config.get("expirationDateTime")) is not None and _v != ""
        else "2026-09-29T18:23:45.9356913Z"
    )

    return {
        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#subscriptions/$entity",
        "id": bound_id,
        "resource": "me/events",
        "applicationId": "-".join(
            [fake.hex(8), fake.hex(4), fake.hex(4), fake.hex(4), fake.hex(12)]
        ),
        "changeType": "created,updated,deleted",
        "notificationUrl": "https://host.example.test/hooks/microsoft-outlook",
        "expirationDateTime": bound_expirationdatetime,
        "latestSupportedTlsVersion": "v1_2",
        "notificationContentType": "application/json",
    }


def _calendar_changed(config: dict[str, Any], fake: FakeValues) -> Any:
    bound_subscriptionid = "-".join(
        [fake.hex(8), fake.hex(4), fake.hex(4), fake.hex(4), fake.hex(12)]
    )
    bound_eventid = fake.id("AAMkAGI2")

    return {
        "value": [
            {
                "subscriptionId": bound_subscriptionid,
                "subscriptionExpirationDateTime": "2026-09-22T18:23:45.9356913Z",
                "changeType": (
                    str(_v)
                    if (_v := config.get("sample")) is not None and _v != ""
                    else "updated"
                ),
                "resource": "Users/ada%40example.test/Events/AAMkAGI2",
                "tenantId": "-".join(
                    [fake.hex(8), fake.hex(4), fake.hex(4), fake.hex(4), fake.hex(12)]
                ),
                "clientState": "verified-before-injection",
                "resourceData": {
                    "@odata.type": "#Microsoft.Graph.Event",
                    "@odata.id": "Users/ada%40example.test/Events/AAMkAGI2",
                    "@odata.etag": "W/\"ZlnW4RIAV06KYYwlrfNZvQAALfZeRQ==\"",
                    "id": bound_eventid,
                },
            },
        ],
    }


def respond(operation: str, request: dict[str, Any]) -> Any:
    """Dispatch to the fixture for one operation."""
    config: dict[str, Any] = request.get("config") or {}
    fake: FakeValues = request["fake"]

    if operation == "event_get":
        return _event_get(config, fake)

    if operation == "event_list":
        return _event_list(config, fake)

    if operation == "subscription_create":
        return _subscription_create(config, fake)

    if operation == "subscription_delete":
        return _subscription_delete(config, fake)

    if operation == "subscription_renew":
        return _subscription_renew(config, fake)

    if operation == "calendar_changed":
        return _calendar_changed(config, fake)

    # A faker asked for an operation it has no shape for must SAY so. Making
    # something up would produce a green run whose output silently has none of
    # the fields the author is about to reference.
    raise ValueError(
        f'microsoft_outlook: no fake response is defined for "{operation}". '
        "Add a fixture under provider/fixtures/ and regenerate — a connector without a faker "
        "cannot be developed against, tested, or demonstrated."
    )
