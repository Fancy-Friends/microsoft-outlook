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
 * Outlook calendar change — Start a run when an event is created, updated or
 * deleted on a subscribed Outlook calendar.
 *
 * https://learn.microsoft.com/en-us/graph/change-notifications-delivery-webhooks?view=graph-rest-1.0
 *
 * Delivery: subscription. The host calls subscription_create with its own URL
 * for this trigger, an expiry of now plus this lifetime, and the connection's
 * clientState, and must answer Graph's validationToken challenge on that URL
 * as plain text within 10 seconds or the subscription is refused. It stores
 * the returned id and expirationDateTime. 3600 seconds before expiry it calls
 * subscription_renew; a 404 there means the subscription is gone and it
 * re-lists (event_list) and creates again, because notifications during the
 * gap are gone. Every notification is a batch: the host compares each item's
 * clientState with the connection's clientState, answers 202, and then
 * processes the items. A `missed` lifecycle event means re-list;
 * `reauthorizationRequired` means renew; `subscriptionRemoved` means create
 * again and re-list.
 */

import type { NodeKindDefinition } from "@particle-academy/fancy-flow/engine";
import { defineConnectorKind, summarize, type OutputField } from "@particle-academy/fancy-flow/connectors";
import { microsoftOutlookMeta } from "../service.js";

export const MICROSOFT_OUTLOOK_CALENDAR_CHANGED_TRIGGER_KIND = "@particle-academy/microsoft_outlook_calendar_changed_trigger";
export const MICROSOFT_OUTLOOK_CALENDAR_CHANGED_TRIGGER_OPERATION = "calendar_changed";

export const MICROSOFT_OUTLOOK_CALENDAR_CHANGED_TRIGGER_META = microsoftOutlookMeta("trigger", "a calendar change", "https://learn.microsoft.com/en-us/graph/change-notifications-delivery-webhooks?view=graph-rest-1.0");

/**
 * What this node emits — the "ingredients" a downstream node can reference.
 *
 * fancy-flow reads `outputShape` off the kind and offers it in the variable
 * picker, so declaring it is the whole of the work: an author configuring the
 * next node picks `{{ $json.data.id }}` off a list instead of typing a path
 * and hoping.
 */
export const MICROSOFT_OUTLOOK_CALENDAR_CHANGED_TRIGGER_OUTPUT: OutputField[] = [
  {
    "path": "value",
    "type": "array",
    "description": "The notifications in this delivery -- a batch, possibly spanning several subscriptions. Each item is described below."
  },
  {
    "path": "value[].subscriptionId",
    "type": "string",
    "description": "Which subscription this item belongs to."
  },
  {
    "path": "value[].subscriptionExpirationDateTime",
    "type": "string",
    "description": "When that subscription expires, ISO 8601 UTC -- Graph's hint for when to renew."
  },
  {
    "path": "value[].changeType",
    "type": "string",
    "description": "created, updated or deleted."
  },
  {
    "path": "value[].resource",
    "type": "string",
    "description": "The changed resource's path, e.g. Users/{id}/Events/{id}."
  },
  {
    "path": "value[].resourceData.id",
    "type": "string",
    "description": "The changed event's id. event_get reads the rest."
  },
  {
    "path": "value[].resourceData.@odata.type",
    "type": "string",
    "description": "#Microsoft.Graph.Event."
  },
  {
    "path": "value[].tenantId",
    "type": "string",
    "description": "The tenant the change happened in."
  },
  {
    "path": "value[].lifecycleEvent",
    "type": "string",
    "description": "Only on a LIFECYCLE notification: subscriptionRemoved, missed or reauthorizationRequired. Absent on a change notification."
  }
];

export const microsoftOutlookCalendarChangedTriggerKind: NodeKindDefinition = defineConnectorKind(MICROSOFT_OUTLOOK_CALENDAR_CHANGED_TRIGGER_META, {
  name: MICROSOFT_OUTLOOK_CALENDAR_CHANGED_TRIGGER_KIND,
  aliases: ["microsoft_outlook_calendar_changed_trigger"],
  label: "Outlook calendar change",
  description: "Start a run when an event is created, updated or deleted on a subscribed Outlook calendar.",
  icon: "📅",
  inputs: [],
  outputs: [{ id: "out" }],
  sideEffects: "none",
  outputShape: MICROSOFT_OUTLOOK_CALENDAR_CHANGED_TRIGGER_OUTPUT,
  configSchema: [
    {
      "type": "select",
      "key": "sample",
      "label": "Sample change (fake mode)",
      "default": "updated",
      "description": "Which changeType the faked notification carries.",
      "options": [
        {
          "value": "created",
          "label": "created"
        },
        {
          "value": "updated",
          "label": "updated"
        },
        {
          "value": "deleted",
          "label": "deleted"
        }
      ]
    }
  ],
  defaultConfig: {
    "mode": "auto",
    "sample": "updated"
  },
  renderBody: ({ config }) =>
    summarize(MICROSOFT_OUTLOOK_CALENDAR_CHANGED_TRIGGER_META, config as Record<string, unknown>, "a calendar change"),
});
