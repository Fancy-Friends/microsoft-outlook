/**
 * GENERATED FILE — do not edit.
 *
 * Emitted from provider/manifest.json by weaver's generator.
 * A hand-edit here is destroyed by the next protocol sync, which is worse than
 * being rejected, because it works until it silently does not. Fix
 * provider/manifest.json (or weaver's template/) and regenerate:
 *
 *     npm run provider -- microsoft_outlook
 */

/**
 * Microsoft Outlook, as one service descriptor shared by every Microsoft
 * Outlook operation.
 *
 * @particle-academy/fancy-connector-core carries what is true of ALL
 * connectors. This carries what is true of Microsoft Outlook: its base URL,
 * its auth scheme, its idempotency header, and its faker.
 *
 * ## The sandbox trap, written down where it is used
 *
 * Microsoft's test estate is a Microsoft 365 Developer Program sandbox: a
 * separate E5 tenant with its own users, its own calendars and its own OAuth
 * consent, on the SAME host. Eligibility is limited -- Visual Studio
 * Professional or Enterprise subscribers, ISV Success / Microsoft AI Cloud
 * Partner Program members, or Premier/Unified Support customers -- and the
 * tenant expires after 90 days of no development activity. Without one, every
 * call reaches a real mailbox.
 */

import type { ConnectorMode, PreparedRequest, ServiceDescriptor } from "@particle-academy/fancy-connector-core";

import { microsoftOutlookFaker } from "./faker.js";

/**
 * The connector API version this package was GENERATED against.
 *
 * A literal, never imported. An imported constant lets an upgrade rewrite the
 * very claim it exists to detect, after which the copy agrees with itself
 * forever.
 */
export const CONNECTOR_API_VERSION = 1;

export const MICROSOFT_OUTLOOK_BASE_URLS = {
  "live": "https://graph.microsoft.com",
  "sandbox": "https://graph.microsoft.com"
} as const;

/** Credential keys a remote call cannot proceed without. */
export const MICROSOFT_OUTLOOK_REQUIRES = [
  "accessToken",
  "refreshToken",
  "clientId",
  "clientSecret"
] as const;

/**
 * Apply Microsoft Outlook's auth scheme to an outgoing request.
 *
 *
 *
 * The mode is passed in because for some providers auth and estate are the
 * same decision expressed in the URL; here it is unused, and saying so is
 * cheaper than wondering later whether it was forgotten.
 */
export function microsoftOutlookAuthorize(
  credentials: Record<string, string | undefined>,
  request: PreparedRequest,
  _mode: ConnectorMode,
): void {
  request.headers.Authorization = `Bearer ${credentials.accessToken ?? ""}`;
}

/** The Microsoft Outlook service, for the TypeScript runtime. */
export const MICROSOFT_OUTLOOK: ServiceDescriptor = {
  service: "microsoft_outlook",
  title: "Microsoft Outlook",
  sandbox: "separate-account",
  baseUrls: { ...MICROSOFT_OUTLOOK_BASE_URLS },
  requires: [...MICROSOFT_OUTLOOK_REQUIRES],
  authorize: microsoftOutlookAuthorize,
  faker: microsoftOutlookFaker,
};
