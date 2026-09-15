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
 * Subscribe to change notifications on the connected account's calendar
 * events. The host's subscription machinery calls this; it is not a node most
 * workflows need.
 *
 * POST /v1.0/subscriptions —
 * https://learn.microsoft.com/en-us/graph/api/subscription-post-subscriptions?view=graph-rest-1.0
 *
 * Notice what is NOT here: no key, no base URL, no mode check, no retry loop,
 * no fake/real branch. This describes the request; callConnector resolves the
 * connection, picks the estate, and either calls Microsoft Outlook or calls
 * the faker.
 *
 * sideEffects: unsafe-to-replay.
 */

import {
  callConnector,
  type ConnectorResult,
  type RequestedMode,
  type Transport,
} from "@particle-academy/fancy-connector-core";
import { MICROSOFT_OUTLOOK } from "../service.js";

export const SUBSCRIPTION_CREATE_OPERATION = "subscription_create";

export type SubscriptionCreateOptions = {
  /** The node's resolved config. Keys: changeType, resource, notificationUrl, lifecycleNotificationUrl, expirationDateTime, clientState. */
  config: Record<string, unknown>;
  credentials?: Record<string, string | undefined>;
  mode?: RequestedMode;
  connectionId?: string | null;
  input?: unknown;
  attempts?: number;
  /** Override the transport. The only way to exercise this without a network. */
  transport?: Transport;
};

export async function microsoftOutlookSubscriptionCreate(options: SubscriptionCreateOptions): Promise<ConnectorResult> {
  const config = options.config ?? {};

  if (config.resource === undefined || config.resource === null || config.resource === "") {
    throw new Error(`subscription_create: "resource" is required (Resource).`);
  }

  if (config.notificationUrl === undefined || config.notificationUrl === null || config.notificationUrl === "") {
    throw new Error(`subscription_create: "notificationUrl" is required (Notification URL).`);
  }

  if (config.expirationDateTime === undefined || config.expirationDateTime === null || config.expirationDateTime === "") {
    throw new Error(`subscription_create: "expirationDateTime" is required (Expires at).`);
  }

  if (config.clientState === undefined || config.clientState === null || config.clientState === "") {
    throw new Error(`subscription_create: "clientState" is required (Client state).`);
  }

  return callConnector(MICROSOFT_OUTLOOK, {
    operation: SUBSCRIPTION_CREATE_OPERATION,
    config,
    input: options.input,
    ...(options.credentials === undefined ? {} : { credentials: options.credentials }),
    ...(options.mode === undefined ? {} : { mode: options.mode }),
    ...(options.connectionId === undefined ? {} : { connectionId: options.connectionId }),
    ...(options.attempts === undefined ? {} : { attempts: options.attempts }),
    ...(options.transport === undefined ? {} : { transport: options.transport }),
    request: {
      method: "POST",
      path: "/v1.0/subscriptions",
      json: {
        ...(config.changeType !== undefined && config.changeType !== null && config.changeType !== "" ? { "changeType": String(config.changeType) } : {}),
        "resource": String(config.resource),
        "notificationUrl": String(config.notificationUrl),
        ...(config.lifecycleNotificationUrl !== undefined && config.lifecycleNotificationUrl !== null && config.lifecycleNotificationUrl !== "" ? { "lifecycleNotificationUrl": String(config.lifecycleNotificationUrl) } : {}),
        "expirationDateTime": String(config.expirationDateTime),
        "clientState": String(config.clientState),
      },
    },
  });
}
