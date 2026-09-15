/**
 * GENERATED FILE — do not edit.
 *
 * Emitted from provider/actions/event-get.json by weaver's generator.
 * A hand-edit here is destroyed by the next protocol sync, which is worse than
 * being rejected, because it works until it silently does not. Fix
 * provider/actions/event-get.json (or weaver's template/) and regenerate:
 *
 *     npm run provider -- microsoft_outlook
 */

/**
 * Outlook calendar event — Read one event from the connected account's Outlook
 * calendar.
 *
 * https://learn.microsoft.com/en-us/graph/api/event-get?view=graph-rest-1.0
 */

import type { NodeKindDefinition } from "@particle-academy/fancy-flow/engine";
import { defineConnectorKind, summarize, type OutputField } from "@particle-academy/fancy-flow/connectors";
import { microsoftOutlookMeta } from "../service.js";

export const MICROSOFT_OUTLOOK_EVENT_GET_KIND = "@particle-academy/microsoft_outlook_event_get";
export const MICROSOFT_OUTLOOK_EVENT_GET_OPERATION = "event_get";

export const MICROSOFT_OUTLOOK_EVENT_GET_META = microsoftOutlookMeta("action", "read an event", "https://learn.microsoft.com/en-us/graph/api/event-get?view=graph-rest-1.0");

/**
 * What this node emits — the "ingredients" a downstream node can reference.
 *
 * fancy-flow reads `outputShape` off the kind and offers it in the variable
 * picker, so declaring it is the whole of the work: an author configuring the
 * next node picks `{{ $json.data.id }}` off a list instead of typing a path
 * and hoping.
 */
export const MICROSOFT_OUTLOOK_EVENT_GET_OUTPUT: OutputField[] = [
  {
    "path": "data.id",
    "type": "string",
    "description": "The event id."
  },
  {
    "path": "data.subject",
    "type": "string",
    "description": "The title."
  },
  {
    "path": "data.bodyPreview",
    "type": "string",
    "description": "The first lines of the body, as text."
  },
  {
    "path": "data.start.dateTime",
    "type": "string",
    "description": "Start, in start.timeZone. Graph writes seven fractional digits."
  },
  {
    "path": "data.end.dateTime",
    "type": "string",
    "description": "End, in end.timeZone."
  },
  {
    "path": "data.organizer.emailAddress.address",
    "type": "string",
    "description": "Who organises it."
  },
  {
    "path": "data.isCancelled",
    "type": "boolean",
    "description": "Whether the event was cancelled."
  },
  {
    "path": "data.webLink",
    "type": "string",
    "description": "The event in Outlook on the web."
  },
  {
    "path": "data.lastModifiedDateTime",
    "type": "string",
    "description": "ISO 8601 last modification, UTC."
  },
  {
    "path": "mode",
    "type": "string",
    "description": "Which estate this ran against: fake, sandbox or live."
  }
];

export const microsoftOutlookEventGetKind: NodeKindDefinition = defineConnectorKind(MICROSOFT_OUTLOOK_EVENT_GET_META, {
  name: MICROSOFT_OUTLOOK_EVENT_GET_KIND,
  aliases: ["microsoft_outlook_event_get"],
  label: "Outlook calendar event",
  description: "Read one event from the connected account's Outlook calendar.",
  inputs: [{ id: "in" }],
  outputs: [{ id: "out" }],
  sideEffects: "none",
  outputShape: MICROSOFT_OUTLOOK_EVENT_GET_OUTPUT,
  configSchema: [
    {
      "type": "text",
      "key": "id",
      "label": "Event ID",
      "required": true,
      "description": "The event's id, as published by a notification's resourceData.id or by event_list."
    }
  ],
  defaultConfig: {
    "mode": "auto"
  },
  renderBody: ({ config }) =>
    summarize(MICROSOFT_OUTLOOK_EVENT_GET_META, config as Record<string, unknown>, "read an event"),
});
