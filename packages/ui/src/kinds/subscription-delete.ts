/**
 * GENERATED FILE — do not edit.
 *
 * Emitted from provider/actions/subscription-delete.json by weaver's generator.
 * A hand-edit here is destroyed by the next protocol sync, which is worse than
 * being rejected, because it works until it silently does not. Fix
 * provider/actions/subscription-delete.json (or weaver's template/) and regenerate:
 *
 *     npm run provider -- microsoft_outlook
 */

/**
 * Outlook subscription removal — Delete a change-notification subscription.
 * The host's subscription machinery calls this when a trigger is removed.
 *
 * https://learn.microsoft.com/en-us/graph/api/subscription-delete?view=graph-rest-1.0
 */

import type { NodeKindDefinition } from "@particle-academy/fancy-flow/engine";
import { defineConnectorKind, summarize, type OutputField } from "@particle-academy/fancy-flow/connectors";
import { microsoftOutlookMeta } from "../service.js";

export const MICROSOFT_OUTLOOK_SUBSCRIPTION_DELETE_KIND = "@particle-academy/microsoft_outlook_subscription_delete";
export const MICROSOFT_OUTLOOK_SUBSCRIPTION_DELETE_OPERATION = "subscription_delete";

export const MICROSOFT_OUTLOOK_SUBSCRIPTION_DELETE_META = microsoftOutlookMeta("action", "delete a subscription", "https://learn.microsoft.com/en-us/graph/api/subscription-delete?view=graph-rest-1.0");

/**
 * What this node emits — the "ingredients" a downstream node can reference.
 *
 * fancy-flow reads `outputShape` off the kind and offers it in the variable
 * picker, so declaring it is the whole of the work: an author configuring the
 * next node picks `{{ $json.data.id }}` off a list instead of typing a path
 * and hoping.
 */
export const MICROSOFT_OUTLOOK_SUBSCRIPTION_DELETE_OUTPUT: OutputField[] = [
  {
    "path": "mode",
    "type": "string",
    "description": "Which estate this ran against. Graph answers 204 with no body, so this is the only thing published."
  }
];

export const microsoftOutlookSubscriptionDeleteKind: NodeKindDefinition = defineConnectorKind(MICROSOFT_OUTLOOK_SUBSCRIPTION_DELETE_META, {
  name: MICROSOFT_OUTLOOK_SUBSCRIPTION_DELETE_KIND,
  aliases: ["microsoft_outlook_subscription_delete"],
  label: "Outlook subscription removal",
  description: "Delete a change-notification subscription. The host's subscription machinery calls this when a trigger is removed.",
  inputs: [{ id: "in" }],
  outputs: [{ id: "out" }],
  sideEffects: "idempotent",
  outputShape: MICROSOFT_OUTLOOK_SUBSCRIPTION_DELETE_OUTPUT,
  configSchema: [
    {
      "type": "text",
      "key": "id",
      "label": "Subscription ID",
      "required": true,
      "description": "The id subscription_create answered with."
    }
  ],
  defaultConfig: {
    "mode": "auto"
  },
  renderBody: ({ config }) =>
    summarize(MICROSOFT_OUTLOOK_SUBSCRIPTION_DELETE_META, config as Record<string, unknown>, "delete a subscription"),
});
