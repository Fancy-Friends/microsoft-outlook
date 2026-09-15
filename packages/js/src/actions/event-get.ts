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
 * Read one event from the connected account's Outlook calendar.
 *
 * GET /v1.0/me/events/{id} —
 * https://learn.microsoft.com/en-us/graph/api/event-get?view=graph-rest-1.0
 *
 * Notice what is NOT here: no key, no base URL, no mode check, no retry loop,
 * no fake/real branch. This describes the request; callConnector resolves the
 * connection, picks the estate, and either calls Microsoft Outlook or calls
 * the faker.
 */

import {
  callConnector,
  type ConnectorResult,
  type RequestedMode,
  type Transport,
} from "@particle-academy/fancy-connector-core";
import { MICROSOFT_OUTLOOK } from "../service.js";

export const EVENT_GET_OPERATION = "event_get";

export type EventGetOptions = {
  /** The node's resolved config. Keys: id. */
  config: Record<string, unknown>;
  credentials?: Record<string, string | undefined>;
  mode?: RequestedMode;
  connectionId?: string | null;
  input?: unknown;
  attempts?: number;
  /** Override the transport. The only way to exercise this without a network. */
  transport?: Transport;
};

export async function microsoftOutlookEventGet(options: EventGetOptions): Promise<ConnectorResult> {
  const config = options.config ?? {};

  if (config.id === undefined || config.id === null || config.id === "") {
    throw new Error(`event_get: "id" is required (Event ID).`);
  }

  return callConnector(MICROSOFT_OUTLOOK, {
    operation: EVENT_GET_OPERATION,
    config,
    input: options.input,
    ...(options.credentials === undefined ? {} : { credentials: options.credentials }),
    ...(options.mode === undefined ? {} : { mode: options.mode }),
    ...(options.connectionId === undefined ? {} : { connectionId: options.connectionId }),
    ...(options.attempts === undefined ? {} : { attempts: options.attempts }),
    ...(options.transport === undefined ? {} : { transport: options.transport }),
    request: {
      method: "GET",
      path: `/v1.0/me/events/${encodeURIComponent(String(config.id))}`,
      query: {},
    },
  });
}
