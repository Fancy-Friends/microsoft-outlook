/**
 * GENERATED FILE — do not edit.
 *
 * Emitted from provider/actions/event-list.json by weaver's generator.
 * A hand-edit here is destroyed by the next protocol sync, which is worse than
 * being rejected, because it works until it silently does not. Fix
 * provider/actions/event-list.json (or weaver's template/) and regenerate:
 *
 *     npm run provider -- microsoft_outlook
 */

/**
 * Outlook calendar events — List events from the connected account's Outlook
 * calendar.
 *
 * https://learn.microsoft.com/en-us/graph/api/user-list-events?view=graph-rest-1.0
 */

import type { NodeKindDefinition } from "@particle-academy/fancy-flow/engine";
import { defineConnectorKind, summarize, type OutputField } from "@particle-academy/fancy-flow/connectors";
import { microsoftOutlookMeta } from "../service.js";

export const MICROSOFT_OUTLOOK_EVENT_LIST_KIND = "@particle-academy/microsoft_outlook_event_list";
export const MICROSOFT_OUTLOOK_EVENT_LIST_OPERATION = "event_list";

export const MICROSOFT_OUTLOOK_EVENT_LIST_META = microsoftOutlookMeta("action", "list events", "https://learn.microsoft.com/en-us/graph/api/user-list-events?view=graph-rest-1.0");

/**
 * What this node emits — the "ingredients" a downstream node can reference.
 *
 * fancy-flow reads `outputShape` off the kind and offers it in the variable
 * picker, so declaring it is the whole of the work: an author configuring the
 * next node picks `{{ $json.data.id }}` off a list instead of typing a path
 * and hoping.
 */
export const MICROSOFT_OUTLOOK_EVENT_LIST_OUTPUT: OutputField[] = [
  {
    "path": "data.value",
    "type": "array",
    "description": "The events on this page, each shaped like event_get's output."
  },
  {
    "path": "data.@odata.nextLink",
    "type": "string",
    "description": "Present when there is another page: the full URL of the next one."
  },
  {
    "path": "mode",
    "type": "string",
    "description": "Which estate this ran against: fake, sandbox or live."
  }
];

export const microsoftOutlookEventListKind: NodeKindDefinition = defineConnectorKind(MICROSOFT_OUTLOOK_EVENT_LIST_META, {
  name: MICROSOFT_OUTLOOK_EVENT_LIST_KIND,
  aliases: ["microsoft_outlook_event_list"],
  label: "Outlook calendar events",
  description: "List events from the connected account's Outlook calendar.",
  inputs: [{ id: "in" }],
  outputs: [{ id: "out" }],
  sideEffects: "none",
  outputShape: MICROSOFT_OUTLOOK_EVENT_LIST_OUTPUT,
  configSchema: [
    {
      "type": "number",
      "key": "top",
      "label": "Page size",
      "min": 1,
      "max": 999,
      "default": 50,
      "description": "Events per page. An @odata.nextLink in the response means there are more."
    },
    {
      "type": "text",
      "key": "filter",
      "label": "Filter",
      "placeholder": "lastModifiedDateTime ge 2026-09-01T00:00:00Z",
      "description": "An OData $filter. Leave blank for everything."
    },
    {
      "type": "text",
      "key": "orderby",
      "label": "Order",
      "placeholder": "start/dateTime desc",
      "description": "An OData $orderby. Leave blank for Graph's default order."
    },
    {
      "type": "text",
      "key": "select",
      "label": "Fields",
      "placeholder": "id,subject,start,end",
      "description": "An OData $select, comma separated. Leave blank for Graph's default set."
    }
  ],
  defaultConfig: {
    "mode": "auto"
  },
  renderBody: ({ config }) =>
    summarize(MICROSOFT_OUTLOOK_EVENT_LIST_META, config as Record<string, unknown>, "list events"),
});
