/**
 * GENERATED FILE — do not edit.
 *
 * Emitted from provider/triggers/calendar-changed.json by weaver's generator.
 * A hand-edit here is destroyed by the next protocol sync, which is worse than
 * being rejected, because it works until it silently does not. Fix
 * provider/triggers/calendar-changed.json (or weaver's template/) and regenerate:
 *
 *     npm run provider -- microsoft_outlook
 */

/**
 * Microsoft Outlook's subscription trigger — the delivery contract.
 *
 * Kept beside the service descriptor rather than inside a node, because the
 * way a delivery is verified is a fact about MICROSOFT OUTLOOK. Two Microsoft
 * Outlook triggers must not be able to disagree about how a delivery is
 * verified.
 *
 * A SUBSCRIPTION: Microsoft Outlook stops delivering unless somebody renews
 * it, forever, and if nobody does the workflow stops firing with no error
 * anywhere. The lease below says where the expiry is read from and how early
 * to renew; the host runs ONE renewal scheduler for every expiring trigger.
 */

import { handshakeResponse, verifyDelivery, type ChallengeHandshake, type InboundDelivery, type LeaseDeclaration, type SharedTokenScheme, type TriggerDescriptor, type WebhookVerification } from "@particle-academy/fancy-connector-core";
import { microsoftOutlookFaker } from "../faker.js";

/** Where Microsoft Outlook echoes the token it was given when the subscription was created. */
export const MICROSOFT_OUTLOOK_CALENDAR_CHANGED_TOKEN_SCHEME: SharedTokenScheme = { kind: "shared-token", in: "body", path: "value[].clientState" };

/**
 * Microsoft Outlook's challenge before it delivers anything: a request
 * carrying `?validationToken=…` must be answered with that value as plain
 * text, or the subscription is never created.
 * `answerMicrosoftOutlookHandshake` is the pure half; the host's routing layer
 * asks it before mounting the route.
 */
export const MICROSOFT_OUTLOOK_CALENDAR_CHANGED_HANDSHAKE: ChallengeHandshake = { kind: "echo-query", param: "validationToken" };

/**
 * The lease this subscription carries: where the provider's expiry sits in the
 * `subscription_create` response (`expirationDateTime`, rfc3339), how early
 * the host renews, and what it calls when the lease is due. The host builds
 * the value with the core's `leaseFromResponse` and runs one scheduler for
 * every expiring trigger.
 */
export const MICROSOFT_OUTLOOK_CALENDAR_CHANGED_LEASE: LeaseDeclaration = {
  expiresAtFrom: "expirationDateTime",
  expiresAtUnit: "rfc3339",
  renewBeforeSeconds: 3600,
  renewOperation: "subscription_renew",
};

/**
 * Which of this package's actions create, renew and stop the subscription.
 * `renew` is null where the provider cannot renew and the host calls create
 * again — after stopping the old one, and after re-listing, because
 * notifications during the gap are gone.
 */
export const MICROSOFT_OUTLOOK_CALENDAR_CHANGED_SUBSCRIPTION = { create: "subscription_create", renew: "subscription_renew", stop: "subscription_delete" } as const;

export const MICROSOFT_OUTLOOK_CALENDAR_CHANGED: TriggerDescriptor = {
  service: "microsoft_outlook",
  operation: "calendar_changed",
  delivery: "subscription",
  setup:
    "The host calls subscription_create with its own URL for this trigger, an expiry of now plus this lifetime, and the connection's clientState, and must answer Graph's validationToken challenge on that URL as plain text within 10 seconds or the subscription is refused. It stores the returned id and expirationDateTime. 3600 seconds before expiry it calls subscription_renew; a 404 there means the subscription is gone and it re-lists (event_list) and creates again, because notifications during the gap are gone. Every notification is a batch: the host compares each item's clientState with the connection's clientState, answers 202, and then processes the items. A `missed` lifecycle event means re-list; `reauthorizationRequired` means renew; `subscriptionRemoved` means create again and re-list.",
  subscriptionTtl: 604800,
  lease: MICROSOFT_OUTLOOK_CALENDAR_CHANGED_LEASE,
  verification: {
    scheme: MICROSOFT_OUTLOOK_CALENDAR_CHANGED_TOKEN_SCHEME,
    handshake: MICROSOFT_OUTLOOK_CALENDAR_CHANGED_HANDSHAKE,
  },
  faker: microsoftOutlookFaker,
};

/**
 * Verify one inbound Microsoft Outlook delivery.
 *
 * The host calls this BEFORE starting a run, with the body exactly as
 * received. Re-serialised JSON changes key order and whitespace, and produces
 * a mismatch that looks precisely like a wrong secret — hours of debugging the
 * wrong thing.
 *
 * The secret is the connection's `clientState`.
 */
export function verifyMicrosoftOutlookDelivery(
  delivery: InboundDelivery,
  clientState: string | undefined,
  now?: number,
): Promise<WebhookVerification> {
  return verifyDelivery(MICROSOFT_OUTLOOK_CALENDAR_CHANGED, delivery, clientState, now);
}

/**
 * Answer Microsoft Outlook's challenge — what to send back (200, text/plain,
 * the decoded `validationToken`), or undefined when the request is not a
 * challenge at all.
 */
export function answerMicrosoftOutlookHandshake(
  query: Record<string, string | string[] | undefined>,
): ReturnType<typeof handshakeResponse> {
  return handshakeResponse(MICROSOFT_OUTLOOK_CALENDAR_CHANGED_HANDSHAKE, query);
}
