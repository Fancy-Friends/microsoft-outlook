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
 * List events from the connected account's Outlook calendar.
 *
 * GET /v1.0/me/events —
 * https://learn.microsoft.com/en-us/graph/api/user-list-events?view=graph-rest-1.0
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

export const EVENT_LIST_OPERATION = "event_list";

export type EventListOptions = {
  /** The node's resolved config. Keys: top, filter, orderby, select. */
  config: Record<string, unknown>;
  credentials?: Record<string, string | undefined>;
  mode?: RequestedMode;
  connectionId?: string | null;
  input?: unknown;
  attempts?: number;
  /** Override the transport. The only way to exercise this without a network. */
  transport?: Transport;
};

export async function microsoftOutlookEventList(options: EventListOptions): Promise<ConnectorResult> {
  const config = options.config ?? {};

  {
    const n = Number(config.top);
    const given = config.top !== undefined && config.top !== null && config.top !== "";
    if (given && !(Number.isInteger(n) && n >= 1 && n <= 999)) {
      throw new Error(
        `event_list: "top" must be a integer, got ${JSON.stringify(config.top)}.`,
      );
    }
  }

  return callConnector(MICROSOFT_OUTLOOK, {
    operation: EVENT_LIST_OPERATION,
    config,
    input: options.input,
    ...(options.credentials === undefined ? {} : { credentials: options.credentials }),
    ...(options.mode === undefined ? {} : { mode: options.mode }),
    ...(options.connectionId === undefined ? {} : { connectionId: options.connectionId }),
    ...(options.attempts === undefined ? {} : { attempts: options.attempts }),
    ...(options.transport === undefined ? {} : { transport: options.transport }),
    request: {
      method: "GET",
      path: "/v1.0/me/events",
      query: {
        "$top": config.top !== undefined && config.top !== null && config.top !== "" ? Math.trunc(Number(config.top)) : 50,
        ...(config.filter !== undefined && config.filter !== null && config.filter !== "" ? { "$filter": String(config.filter) } : {}),
        ...(config.orderby !== undefined && config.orderby !== null && config.orderby !== "" ? { "$orderby": String(config.orderby) } : {}),
        ...(config.select !== undefined && config.select !== null && config.select !== "" ? { "$select": String(config.select) } : {}),
      },
    },
  });
}
