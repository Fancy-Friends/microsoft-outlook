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
 * Extend a change-notification subscription before it expires. The host's
 * subscription machinery calls this when the lease is due.
 *
 * PATCH /v1.0/subscriptions/{id} —
 * https://learn.microsoft.com/en-us/graph/api/subscription-update?view=graph-rest-1.0
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

export const SUBSCRIPTION_RENEW_OPERATION = "subscription_renew";

export type SubscriptionRenewOptions = {
  /** The node's resolved config. Keys: id, expirationDateTime. */
  config: Record<string, unknown>;
  credentials?: Record<string, string | undefined>;
  mode?: RequestedMode;
  connectionId?: string | null;
  input?: unknown;
  attempts?: number;
  /** Override the transport. The only way to exercise this without a network. */
  transport?: Transport;
};

export async function microsoftOutlookSubscriptionRenew(options: SubscriptionRenewOptions): Promise<ConnectorResult> {
  const config = options.config ?? {};

  if (config.id === undefined || config.id === null || config.id === "") {
    throw new Error(`subscription_renew: "id" is required (Subscription ID).`);
  }

  if (config.expirationDateTime === undefined || config.expirationDateTime === null || config.expirationDateTime === "") {
    throw new Error(`subscription_renew: "expirationDateTime" is required (Expires at).`);
  }

  return callConnector(MICROSOFT_OUTLOOK, {
    operation: SUBSCRIPTION_RENEW_OPERATION,
    config,
    input: options.input,
    ...(options.credentials === undefined ? {} : { credentials: options.credentials }),
    ...(options.mode === undefined ? {} : { mode: options.mode }),
    ...(options.connectionId === undefined ? {} : { connectionId: options.connectionId }),
    ...(options.attempts === undefined ? {} : { attempts: options.attempts }),
    ...(options.transport === undefined ? {} : { transport: options.transport }),
    request: {
      method: "PATCH",
      path: `/v1.0/subscriptions/${encodeURIComponent(String(config.id))}`,
      json: {
        "expirationDateTime": String(config.expirationDateTime),
      },
    },
  });
}
