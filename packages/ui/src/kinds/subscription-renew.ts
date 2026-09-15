/**
 * GENERATED FILE — do not edit.
 *
 * Emitted from provider/actions/subscription-renew.json by weaver's generator.
 * A hand-edit here is destroyed by the next protocol sync, which is worse than
 * being rejected, because it works until it silently does not. Fix
 * provider/actions/subscription-renew.json (or weaver's template/) and regenerate:
 *
 *     npm run provider -- microsoft_outlook
 */

/**
 * Outlook subscription renewal — Extend a change-notification subscription
 * before it expires. The host's subscription machinery calls this when the
 * lease is due.
 *
 * https://learn.microsoft.com/en-us/graph/api/subscription-update?view=graph-rest-1.0
 */

import type { NodeKindDefinition } from "@particle-academy/fancy-flow/engine";
import { defineConnectorKind, summarize, type OutputField } from "@particle-academy/fancy-flow/connectors";
import { microsoftOutlookMeta } from "../service.js";

export const MICROSOFT_OUTLOOK_SUBSCRIPTION_RENEW_KIND = "@particle-academy/microsoft_outlook_subscription_renew";
export const MICROSOFT_OUTLOOK_SUBSCRIPTION_RENEW_OPERATION = "subscription_renew";

export const MICROSOFT_OUTLOOK_SUBSCRIPTION_RENEW_META = microsoftOutlookMeta("action", "renew a subscription", "https://learn.microsoft.com/en-us/graph/api/subscription-update?view=graph-rest-1.0");

/**
 * What this node emits — the "ingredients" a downstream node can reference.
 *
 * fancy-flow reads `outputShape` off the kind and offers it in the variable
 * picker, so declaring it is the whole of the work: an author configuring the
 * next node picks `{{ $json.data.id }}` off a list instead of typing a path
 * and hoping.
 */
export const MICROSOFT_OUTLOOK_SUBSCRIPTION_RENEW_OUTPUT: OutputField[] = [
  {
    "path": "data.id",
    "type": "string",
    "description": "The subscription id."
  },
  {
    "path": "data.expirationDateTime",
    "type": "string",
    "description": "The new expiry, ISO 8601 UTC. The trigger's lease is rebuilt from this."
  },
  {
    "path": "mode",
    "type": "string",
    "description": "Which estate this ran against: fake, sandbox or live."
  }
];

export const microsoftOutlookSubscriptionRenewKind: NodeKindDefinition = defineConnectorKind(MICROSOFT_OUTLOOK_SUBSCRIPTION_RENEW_META, {
  name: MICROSOFT_OUTLOOK_SUBSCRIPTION_RENEW_KIND,
  aliases: ["microsoft_outlook_subscription_renew"],
  label: "Outlook subscription renewal",
  description: "Extend a change-notification subscription before it expires. The host's subscription machinery calls this when the lease is due.",
  inputs: [{ id: "in" }],
  outputs: [{ id: "out" }],
  sideEffects: "idempotent",
  outputShape: MICROSOFT_OUTLOOK_SUBSCRIPTION_RENEW_OUTPUT,
  configSchema: [
    {
      "type": "text",
      "key": "id",
      "label": "Subscription ID",
      "required": true,
      "description": "The id subscription_create answered with."
    },
    {
      "type": "text",
      "key": "expirationDateTime",
      "label": "Expires at",
      "required": true,
      "description": "The new expiry, as an ISO 8601 instant in UTC. The host computes now + the trigger's lifetime, capped at 10080 minutes for events."
    }
  ],
  defaultConfig: {
    "mode": "auto"
  },
  renderBody: ({ config }) =>
    summarize(MICROSOFT_OUTLOOK_SUBSCRIPTION_RENEW_META, config as Record<string, unknown>, "renew a subscription"),
});
