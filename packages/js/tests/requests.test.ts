/**
 * GENERATED FILE — do not edit.
 *
 * Emitted from provider/actions/ by weaver's generator.
 * A hand-edit here is destroyed by the next protocol sync, which is worse than
 * being rejected, because it works until it silently does not. Fix
 * provider/actions/ (or weaver's template/) and regenerate:
 *
 *     npm run provider -- microsoft_outlook
 */

/**
 * What Microsoft Outlook actually receives.
 *
 * Every assertion below is about the request rather than the response, and
 * none of it touches the network: the transport is a stub that records what it
 * was handed.
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import type { PreparedRequest } from "@particle-academy/fancy-connector-core";

import { microsoftOutlookEventGet } from "../src/actions/event-get.js";
import { microsoftOutlookEventList } from "../src/actions/event-list.js";
import { microsoftOutlookSubscriptionCreate } from "../src/actions/subscription-create.js";
import { microsoftOutlookSubscriptionDelete } from "../src/actions/subscription-delete.js";
import { microsoftOutlookSubscriptionRenew } from "../src/actions/subscription-renew.js";

/** Capture the prepared request instead of sending it. */
function capture() {
  const seen: PreparedRequest[] = [];

  return {
    seen,
    transport: async (request: PreparedRequest) => {
      seen.push(request);

      return { status: 200, body: JSON.stringify({ id: "captured" }), headers: {} };
    },
  };
}

const CREDENTIALS = {
  "clientId": "test_clientId",
  "clientSecret": "test_clientSecret",
  "accessToken": "test_accessToken",
  "refreshToken": "test_refreshToken",
  "clientState": "test_clientState"
};

test("event_get sends GET /v1.0/me/events/{id}", async () => {
  const { seen, transport } = capture();

  await microsoftOutlookEventGet({
    config: {
      "id": "example-id"
    },
    credentials: CREDENTIALS,
    mode: "live",
    transport,
  });

  assert.equal(seen.length, 1);
  assert.equal(seen[0]!.method, "GET");
  assert.ok(new URL(seen[0]!.url).pathname.endsWith("/v1.0/me/events/example-id"), seen[0]!.url);

  assert.deepEqual(
    Object.fromEntries(new URL(seen[0]!.url).searchParams),
    {},
  );
});

test("event_list sends GET /v1.0/me/events", async () => {
  const { seen, transport } = capture();

  await microsoftOutlookEventList({
    config: {
      "top": 999,
      "filter": "example-filter",
      "orderby": "example-orderby",
      "select": "example-select"
    },
    credentials: CREDENTIALS,
    mode: "live",
    transport,
  });

  assert.equal(seen.length, 1);
  assert.equal(seen[0]!.method, "GET");
  assert.ok(new URL(seen[0]!.url).pathname.endsWith("/v1.0/me/events"), seen[0]!.url);

  assert.deepEqual(
    Object.fromEntries(new URL(seen[0]!.url).searchParams),
    {
      "$top": "999",
      "$filter": "example-filter",
      "$orderby": "example-orderby",
      "$select": "example-select"
    },
  );
});

test("subscription_create sends POST /v1.0/subscriptions", async () => {
  const { seen, transport } = capture();

  await microsoftOutlookSubscriptionCreate({
    config: {
      "changeType": "created,updated,deleted",
      "resource": "example-resource",
      "notificationUrl": "example-notificationUrl",
      "lifecycleNotificationUrl": "example-lifecycleNotificationUrl",
      "expirationDateTime": "example-expirationDateTime",
      "clientState": "example-clientState"
    },
    credentials: CREDENTIALS,
    mode: "live",
    transport,
  });

  assert.equal(seen.length, 1);
  assert.equal(seen[0]!.method, "POST");
  assert.ok(new URL(seen[0]!.url).pathname.endsWith("/v1.0/subscriptions"), seen[0]!.url);

  assert.deepEqual(JSON.parse(String(seen[0]!.body ?? "{}")), {
    "changeType": "created,updated,deleted",
    "resource": "example-resource",
    "notificationUrl": "example-notificationUrl",
    "lifecycleNotificationUrl": "example-lifecycleNotificationUrl",
    "expirationDateTime": "example-expirationDateTime",
    "clientState": "example-clientState"
  });
});

test("subscription_delete sends DELETE /v1.0/subscriptions/{id}", async () => {
  const { seen, transport } = capture();

  await microsoftOutlookSubscriptionDelete({
    config: {
      "id": "example-id"
    },
    credentials: CREDENTIALS,
    mode: "live",
    transport,
  });

  assert.equal(seen.length, 1);
  assert.equal(seen[0]!.method, "DELETE");
  assert.ok(new URL(seen[0]!.url).pathname.endsWith("/v1.0/subscriptions/example-id"), seen[0]!.url);

  assert.deepEqual(
    Object.fromEntries(new URL(seen[0]!.url).searchParams),
    {},
  );
});

test("subscription_renew sends PATCH /v1.0/subscriptions/{id}", async () => {
  const { seen, transport } = capture();

  await microsoftOutlookSubscriptionRenew({
    config: {
      "id": "example-id",
      "expirationDateTime": "example-expirationDateTime"
    },
    credentials: CREDENTIALS,
    mode: "live",
    transport,
  });

  assert.equal(seen.length, 1);
  assert.equal(seen[0]!.method, "PATCH");
  assert.ok(new URL(seen[0]!.url).pathname.endsWith("/v1.0/subscriptions/example-id"), seen[0]!.url);

  assert.deepEqual(JSON.parse(String(seen[0]!.body ?? "{}")), {
    "expirationDateTime": "example-expirationDateTime"
  });
});

test("the credential is placed the way the provider wants it", async () => {
  const { seen, transport } = capture();

  await microsoftOutlookEventGet({
    config: {
      "id": "example-id"
    },
    credentials: CREDENTIALS,
    mode: "live",
    transport,
  });

  assert.equal(seen[0]!.headers.Authorization, "Bearer test_accessToken");
});

test("a missing required field is refused BEFORE anything is sent", async () => {
  // Nothing was attempted, so there is nothing to classify — and the message names
  // the field, rather than letting the provider answer three frames later with
  // "invalid request".
  const { seen, transport } = capture();

  await assert.rejects(
    microsoftOutlookEventGet({
      config: {},
      credentials: CREDENTIALS,
      mode: "live",
      transport,
    }),
    new RegExp("id"),
  );

  assert.equal(seen.length, 0, "the request must not have been sent");
});
