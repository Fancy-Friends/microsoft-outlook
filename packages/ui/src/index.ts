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
 * Microsoft Outlook's node kinds for fancy-flow.
 *
 * Install this on every host. The TypeScript executors live in the js
 * package's `./flow` subpath; PHP and Python hosts run their own and need only
 * this.
 */

export * from "./service.js";
export * from "./kinds/event-get.js";
export * from "./kinds/event-list.js";
export * from "./kinds/subscription-create.js";
export * from "./kinds/subscription-delete.js";
export * from "./kinds/subscription-renew.js";
export * from "./kinds/calendar-changed.js";

import type { NodeKindDefinition } from "@particle-academy/fancy-flow/engine";
import { microsoftOutlookEventGetKind } from "./kinds/event-get.js";
import { microsoftOutlookEventListKind } from "./kinds/event-list.js";
import { microsoftOutlookSubscriptionCreateKind } from "./kinds/subscription-create.js";
import { microsoftOutlookSubscriptionDeleteKind } from "./kinds/subscription-delete.js";
import { microsoftOutlookSubscriptionRenewKind } from "./kinds/subscription-renew.js";
import { microsoftOutlookCalendarChangedTriggerKind } from "./kinds/calendar-changed.js";

/** Every Microsoft Outlook kind, for a host that registers the lot. */
export const MICROSOFT_OUTLOOK_KINDS: NodeKindDefinition[] = [
  microsoftOutlookEventGetKind,
  microsoftOutlookEventListKind,
  microsoftOutlookSubscriptionCreateKind,
  microsoftOutlookSubscriptionDeleteKind,
  microsoftOutlookSubscriptionRenewKind,
  microsoftOutlookCalendarChangedTriggerKind,
];
