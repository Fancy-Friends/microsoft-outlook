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
 * The golden fixtures.
 *
 * Deterministic on purpose: the same seed produces the same bytes in
 * TypeScript, PHP and Python, so this file and its twins in the other packages
 * assert the SAME values. That turns the faker into a parity test rather than
 * a convenience — which matters, because cross-runtime drift does not fail
 * loudly. It completes, down one path, with no error.
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import { fakeRequest } from "@particle-academy/fancy-connector-core";

import { microsoftOutlookFaker } from "../src/faker.js";

test("event_get fakes the shape Microsoft Outlook publishes", () => {
  const config = {};

  const faked = microsoftOutlookFaker("event_get", fakeRequest("microsoft_outlook", "event_get", config));

  assert.deepEqual(faked, {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('ada%40example.test')/events/$entity",
    "@odata.etag": "W/\"ZlnW4RIAV06KYYwlrfNZvQAALfZeRQ==\"",
    "id": "AAMkAGI2_fake_4eccb4b0d301",
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
      "timeZone": "UTC"
    },
    "end": {
      "dateTime": "2026-09-22T15:00:00.0000000",
      "timeZone": "UTC"
    },
    "location": {
      "displayName": "Room 4"
    },
    "organizer": {
      "emailAddress": {
        "name": "Ada Example",
        "address": "ada@example.test"
      }
    },
    "attendees": [
      {
        "type": "required",
        "status": {
          "response": "accepted",
          "time": "2026-09-02T10:00:00.0000000Z"
        },
        "emailAddress": {
          "name": "Grace Example",
          "address": "grace@example.test"
        }
      }
    ]
  });
});

test("event_list fakes the shape Microsoft Outlook publishes", () => {
  const config = {};

  const faked = microsoftOutlookFaker("event_list", fakeRequest("microsoft_outlook", "event_list", config));

  assert.deepEqual(faked, {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('ada%40example.test')/events",
    "value": [
      {
        "@odata.etag": "W/\"ZlnW4RIAV06KYYwlrfNZvQAALfZeRQ==\"",
        "id": "AAMkAGI2_fake_9a231b22d1a3",
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
          "timeZone": "UTC"
        },
        "end": {
          "dateTime": "2026-09-22T15:00:00.0000000",
          "timeZone": "UTC"
        },
        "organizer": {
          "emailAddress": {
            "name": "Ada Example",
            "address": "ada@example.test"
          }
        }
      }
    ]
  });
});

test("subscription_create fakes the shape Microsoft Outlook publishes", () => {
  const config = {};

  const faked = microsoftOutlookFaker("subscription_create", fakeRequest("microsoft_outlook", "subscription_create", config));

  assert.deepEqual(faked, {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#subscriptions/$entity",
    "id": "4743be5e-2184-10d8-1aa8-c17076db3ac5",
    "resource": "me/events",
    "applicationId": "8b47cca0-28ad-a735-7c8e-eb27866b66b7",
    "changeType": "created,updated,deleted",
    "clientState": "cs_fake_6621468790c1",
    "notificationUrl": "https://host.example.test/hooks/microsoft-outlook",
    "expirationDateTime": "2026-09-22T18:23:45.9356913Z",
    "creatorId": "223e4280-a7a8-4fcd-b2a5-bead3a3a24a4",
    "latestSupportedTlsVersion": "v1_2",
    "notificationContentType": "application/json"
  });
});

test("subscription_delete fakes the shape Microsoft Outlook publishes", () => {
  const config = {};

  const faked = microsoftOutlookFaker("subscription_delete", fakeRequest("microsoft_outlook", "subscription_delete", config));

  assert.deepEqual(faked, {});
});

test("subscription_renew fakes the shape Microsoft Outlook publishes", () => {
  const config = {};

  const faked = microsoftOutlookFaker("subscription_renew", fakeRequest("microsoft_outlook", "subscription_renew", config));

  assert.deepEqual(faked, {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#subscriptions/$entity",
    "id": "5ceaa0a0-172a-4973-7579-dcfdbb6519b4",
    "resource": "me/events",
    "applicationId": "5c761487-4231-2970-f5fa-2023840644ce",
    "changeType": "created,updated,deleted",
    "notificationUrl": "https://host.example.test/hooks/microsoft-outlook",
    "expirationDateTime": "2026-09-29T18:23:45.9356913Z",
    "latestSupportedTlsVersion": "v1_2",
    "notificationContentType": "application/json"
  });
});

test("calendar_changed fakes the shape Microsoft Outlook publishes", () => {
  const config = {
    "sample": "updated"
  };

  const faked = microsoftOutlookFaker("calendar_changed", fakeRequest("microsoft_outlook", "calendar_changed", config));

  assert.deepEqual(faked, {
    "value": [
      {
        "subscriptionId": "a15234ea-1e2f-bdbc-a4dd-0d03f54606ca",
        "subscriptionExpirationDateTime": "2026-09-22T18:23:45.9356913Z",
        "changeType": "updated",
        "resource": "Users/ada%40example.test/Events/AAMkAGI2",
        "tenantId": "905d610b-b3ff-d562-2f44-4e30083ca9c2",
        "clientState": "verified-before-injection",
        "resourceData": {
          "@odata.type": "#Microsoft.Graph.Event",
          "@odata.id": "Users/ada%40example.test/Events/AAMkAGI2",
          "@odata.etag": "W/\"ZlnW4RIAV06KYYwlrfNZvQAALfZeRQ==\"",
          "id": "AAMkAGI2_fake_f78e5b59a049"
        }
      }
    ]
  });
});

test("an operation with no fixture throws rather than inventing a shape", () => {
  assert.throws(() => microsoftOutlookFaker("no_such_operation", fakeRequest("microsoft_outlook", "no_such_operation", {})), /no fake response/);
});
