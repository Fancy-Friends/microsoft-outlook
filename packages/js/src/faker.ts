/**
 * GENERATED FILE — do not edit.
 *
 * Emitted from provider/fixtures/ by weaver's generator.
 * A hand-edit here is destroyed by the next protocol sync, which is worse than
 * being rejected, because it works until it silently does not. Fix
 * provider/fixtures/ (or weaver's template/) and regenerate:
 *
 *     npm run provider -- microsoft_outlook
 */

/**
 * The Microsoft Outlook faker.
 *
 * Shapes, not behaviour: the goal is that a downstream node sees the field
 * NAMES Microsoft Outlook actually publishes, so an author can wire {{
 * $json.data.id }} against a fake and have it keep working against the real
 * thing.
 *
 * Deterministic — same inputs, same output. A faker returning a fresh uuid
 * every call cannot be asserted on, so its fixtures degrade to "it did not
 * throw", which is the assertion that catches nothing.
 */

import type { ConnectorFaker, FakeRequest } from "@particle-academy/fancy-connector-core";

function fakeEventGet({ config, fake }: FakeRequest): unknown {
  const boundId = (config.id !== undefined && config.id !== null && config.id !== "" ? String(config.id) : fake.id("AAMkAGI2"));

  return {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('ada%40example.test')/events/$entity",
    "@odata.etag": "W/\"ZlnW4RIAV06KYYwlrfNZvQAALfZeRQ==\"",
    "id": boundId,
    "createdDateTime": "2026-09-01T09:00:00.0000000Z",
    "lastModifiedDateTime": "2026-09-10T15:30:00.0000000Z",
    "subject": "Design review",
    "bodyPreview": "Quarterly design review with the platform team.",
    "isCancelled": false,
    "isAllDay": false,
    "showAs": "busy",
    "webLink": "https://outlook.office365.com/owa/?itemid=AAMkAGI2&exvsurl=1&path=/calendar/item",
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
  };
}

function fakeEventList({ config, fake }: FakeRequest): unknown {
  return {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('ada%40example.test')/events",
    "value": [
      {
        "@odata.etag": "W/\"ZlnW4RIAV06KYYwlrfNZvQAALfZeRQ==\"",
        "id": fake.id("AAMkAGI2"),
        "createdDateTime": "2026-09-01T09:00:00.0000000Z",
        "lastModifiedDateTime": "2026-09-10T15:30:00.0000000Z",
        "subject": "Design review",
        "bodyPreview": "Quarterly design review with the platform team.",
        "isCancelled": false,
        "isAllDay": false,
        "showAs": "busy",
        "webLink": "https://outlook.office365.com/owa/?itemid=AAMkAGI2&exvsurl=1&path=/calendar/item",
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
  };
}

function fakeSubscriptionCreate({ config, fake }: FakeRequest): unknown {
  const boundResource = (config.resource !== undefined && config.resource !== null && config.resource !== "" ? String(config.resource) : "me/events");
  const boundChangetype = (config.changeType !== undefined && config.changeType !== null && config.changeType !== "" ? String(config.changeType) : "created,updated,deleted");
  const boundNotificationurl = (config.notificationUrl !== undefined && config.notificationUrl !== null && config.notificationUrl !== "" ? String(config.notificationUrl) : "https://host.example.test/hooks/microsoft-outlook");
  const boundClientstate = (config.clientState !== undefined && config.clientState !== null && config.clientState !== "" ? String(config.clientState) : fake.id("cs"));

  return {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#subscriptions/$entity",
    "id": `${fake.hex(8)}-${fake.hex(4)}-${fake.hex(4)}-${fake.hex(4)}-${fake.hex(12)}`,
    "resource": boundResource,
    "applicationId": `${fake.hex(8)}-${fake.hex(4)}-${fake.hex(4)}-${fake.hex(4)}-${fake.hex(12)}`,
    "changeType": boundChangetype,
    "clientState": boundClientstate,
    "notificationUrl": boundNotificationurl,
    "expirationDateTime": "2026-09-22T18:23:45.9356913Z",
    "creatorId": `${fake.hex(8)}-${fake.hex(4)}-${fake.hex(4)}-${fake.hex(4)}-${fake.hex(12)}`,
    "latestSupportedTlsVersion": "v1_2",
    "notificationContentType": "application/json",
  };
}

function fakeSubscriptionDelete({ config, fake }: FakeRequest): unknown {
  return {};
}

function fakeSubscriptionRenew({ config, fake }: FakeRequest): unknown {
  const boundId = (config.id !== undefined && config.id !== null && config.id !== "" ? String(config.id) : `${fake.hex(8)}-${fake.hex(4)}-${fake.hex(4)}-${fake.hex(4)}-${fake.hex(12)}`);
  const boundExpirationdatetime = (config.expirationDateTime !== undefined && config.expirationDateTime !== null && config.expirationDateTime !== "" ? String(config.expirationDateTime) : "2026-09-29T18:23:45.9356913Z");

  return {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#subscriptions/$entity",
    "id": boundId,
    "resource": "me/events",
    "applicationId": `${fake.hex(8)}-${fake.hex(4)}-${fake.hex(4)}-${fake.hex(4)}-${fake.hex(12)}`,
    "changeType": "created,updated,deleted",
    "notificationUrl": "https://host.example.test/hooks/microsoft-outlook",
    "expirationDateTime": boundExpirationdatetime,
    "latestSupportedTlsVersion": "v1_2",
    "notificationContentType": "application/json",
  };
}

function fakeCalendarChanged({ config, fake }: FakeRequest): unknown {
  const boundSubscriptionid = `${fake.hex(8)}-${fake.hex(4)}-${fake.hex(4)}-${fake.hex(4)}-${fake.hex(12)}`;
  const boundEventid = fake.id("AAMkAGI2");

  return {
    "value": [
      {
        "subscriptionId": boundSubscriptionid,
        "subscriptionExpirationDateTime": "2026-09-22T18:23:45.9356913Z",
        "changeType": (config.sample !== undefined && config.sample !== null && config.sample !== "" ? String(config.sample) : "updated"),
        "resource": "Users/ada%40example.test/Events/AAMkAGI2",
        "tenantId": `${fake.hex(8)}-${fake.hex(4)}-${fake.hex(4)}-${fake.hex(4)}-${fake.hex(12)}`,
        "clientState": "verified-before-injection",
        "resourceData": {
          "@odata.type": "#Microsoft.Graph.Event",
          "@odata.id": "Users/ada%40example.test/Events/AAMkAGI2",
          "@odata.etag": "W/\"ZlnW4RIAV06KYYwlrfNZvQAALfZeRQ==\"",
          "id": boundEventid,
        },
      },
    ],
  };
}

export const microsoftOutlookFaker: ConnectorFaker = (operation, request) => {
  switch (operation) {
    case "event_get":
      return fakeEventGet(request);

    case "event_list":
      return fakeEventList(request);

    case "subscription_create":
      return fakeSubscriptionCreate(request);

    case "subscription_delete":
      return fakeSubscriptionDelete(request);

    case "subscription_renew":
      return fakeSubscriptionRenew(request);

    case "calendar_changed":
      return fakeCalendarChanged(request);

    default:
      // A faker asked for an operation it has no shape for must SAY so. Making
      // something up would produce a green run whose output silently has none
      // of the fields the author is about to reference.
      throw new Error(
        `microsoft_outlook: no fake response is defined for "${operation}". ` +
          "Add a fixture under provider/fixtures/ and regenerate — a connector without a faker " +
          "cannot be developed against, tested, or demonstrated.",
      );
  }
};
