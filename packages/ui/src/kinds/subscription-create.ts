/**
 * GENERATED FILE — do not edit.
 *
 * Emitted from provider/actions/subscription-create.json by weaver's generator.
 * A hand-edit here is destroyed by the next protocol sync, which is worse than
 * being rejected, because it works until it silently does not. Fix
 * provider/actions/subscription-create.json (or weaver's template/) and regenerate:
 *
 *     npm run provider -- microsoft_outlook
 */

/**
 * Outlook subscription — Subscribe to change notifications on the connected
 * account's calendar events. The host's subscription machinery calls this; it
 * is not a node most workflows need.
 *
 * https://learn.microsoft.com/en-us/graph/api/subscription-post-subscriptions?view=graph-rest-1.0
 *
 * `unsafe-to-replay`.
 */

import type { NodeKindDefinition } from "@particle-academy/fancy-flow/engine";
import { defineConnectorKind, summarize, type OutputField } from "@particle-academy/fancy-flow/connectors";
import { microsoftOutlookMeta } from "../service.js";

export const MICROSOFT_OUTLOOK_SUBSCRIPTION_CREATE_KIND = "@particle-academy/microsoft_outlook_subscription_create";
export const MICROSOFT_OUTLOOK_SUBSCRIPTION_CREATE_OPERATION = "subscription_create";

export const MICROSOFT_OUTLOOK_SUBSCRIPTION_CREATE_META = microsoftOutlookMeta("action", "subscribe to changes", "https://learn.microsoft.com/en-us/graph/api/subscription-post-subscriptions?view=graph-rest-1.0");

/**
 * What this node emits — the "ingredients" a downstream node can reference.
 *
 * fancy-flow reads `outputShape` off the kind and offers it in the variable
 * picker, so declaring it is the whole of the work: an author configuring the
 * next node picks `{{ $json.data.id }}` off a list instead of typing a path
 * and hoping.
 */
export const MICROSOFT_OUTLOOK_SUBSCRIPTION_CREATE_OUTPUT: OutputField[] = [
  {
    "path": "data.id",
    "type": "string",
    "description": "The subscription id. subscription_renew and subscription_delete need it."
  },
  {
    "path": "data.resource",
    "type": "string",
    "description": "What is watched."
  },
  {
    "path": "data.changeType",
    "type": "string",
    "description": "Which changes raise a notification."
  },
  {
    "path": "data.notificationUrl",
    "type": "string",
    "description": "Where notifications go."
  },
  {
    "path": "data.expirationDateTime",
    "type": "string",
    "description": "When it expires, ISO 8601 UTC with a seven-digit fraction. The subscription trigger's lease is built from this."
  },
  {
    "path": "mode",
    "type": "string",
    "description": "Which estate this ran against: fake, sandbox or live."
  }
];

export const microsoftOutlookSubscriptionCreateKind: NodeKindDefinition = defineConnectorKind(MICROSOFT_OUTLOOK_SUBSCRIPTION_CREATE_META, {
  name: MICROSOFT_OUTLOOK_SUBSCRIPTION_CREATE_KIND,
  aliases: ["microsoft_outlook_subscription_create"],
  label: "Outlook subscription",
  description: "Subscribe to change notifications on the connected account's calendar events. The host's subscription machinery calls this; it is not a node most workflows need.",
  inputs: [{ id: "in" }],
  outputs: [{ id: "out" }],
  sideEffects: "unsafe-to-replay",
  outputShape: MICROSOFT_OUTLOOK_SUBSCRIPTION_CREATE_OUTPUT,
  configSchema: [
    {
      "type": "select",
      "key": "changeType",
      "label": "Changes",
      "default": "created,updated,deleted",
      "description": "Which changes raise a notification. Graph accepts a comma-separated combination.",
      "options": [
        {
          "value": "created,updated,deleted",
          "label": "created, updated and deleted"
        },
        {
          "value": "created",
          "label": "created only"
        },
        {
          "value": "updated",
          "label": "updated only"
        },
        {
          "value": "deleted",
          "label": "deleted only"
        }
      ]
    },
    {
      "type": "text",
      "key": "resource",
      "label": "Resource",
      "required": true,
      "default": "me/events",
      "description": "The Graph resource to watch, without the base URL. `me/events` is the connected account's calendar events."
    },
    {
      "type": "text",
      "key": "notificationUrl",
      "label": "Notification URL",
      "required": true,
      "description": "The HTTPS URL the host mounts for this trigger. Graph validates it with a validationToken challenge before creating the subscription, then POSTs change notifications to it."
    },
    {
      "type": "text",
      "key": "lifecycleNotificationUrl",
      "label": "Lifecycle notification URL",
      "description": "Where Graph sends `reauthorizationRequired`, `subscriptionRemoved` and `missed` events. May be the notification URL itself. Optional, but without it a subscription Graph removed is one nobody hears about."
    },
    {
      "type": "text",
      "key": "expirationDateTime",
      "label": "Expires at",
      "required": true,
      "description": "When the subscription expires, as an ISO 8601 instant in UTC. The host computes now + the trigger's lifetime; Graph caps an event subscription at 10080 minutes (under seven days) and raises anything under 45 minutes to 45."
    },
    {
      "type": "text",
      "key": "clientState",
      "label": "Client state",
      "required": true,
      "description": "The connection's clientState. The host fills this from the connection -- never type a value here. Graph echoes it inside every notification, and that echo is how a delivery is verified. At most 128 characters."
    }
  ],
  defaultConfig: {
    "mode": "auto"
  },
  renderBody: ({ config }) =>
    summarize(MICROSOFT_OUTLOOK_SUBSCRIPTION_CREATE_META, config as Record<string, unknown>, "subscribe to changes"),
});
