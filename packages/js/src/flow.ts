/**
 * GENERATED FILE — do not edit.
 *
 * Emitted from provider/actions/ + triggers/ by weaver's generator.
 * A hand-edit here is destroyed by the next protocol sync, which is worse than
 * being rejected, because it works until it silently does not. Fix
 * provider/actions/ + triggers/ (or weaver's template/) and regenerate:
 *
 *     npm run provider -- microsoft_outlook
 */

/**
 * Microsoft Outlook's node kinds with their TypeScript executors attached —
 * for hosts that EXECUTE on TS.
 *
 * The authoring surface in @particle-academy/microsoft-outlook-ui carries no
 * executor: the editor is React on every host, so a PHP or Python project
 * installs the ui package and never this one.
 */

import type { NodeExecutor, NodeKindDefinition } from "@particle-academy/fancy-flow/engine";
import {
  idempotencyKeyFor,
  NO_IDEMPOTENCY_KEY_WARNING,
  resolveConnection,
  triggerEvent,
  type RequestedMode,
} from "@particle-academy/fancy-connector-core";
import { MICROSOFT_OUTLOOK } from "./service.js";

import {
  microsoftOutlookEventGetKind,
  microsoftOutlookEventListKind,
  microsoftOutlookSubscriptionCreateKind,
  microsoftOutlookSubscriptionDeleteKind,
  microsoftOutlookSubscriptionRenewKind,
  microsoftOutlookCalendarChangedTriggerKind,
} from "@particle-academy/microsoft-outlook-ui";

import { microsoftOutlookEventGet } from "./actions/event-get.js";
import { microsoftOutlookEventList } from "./actions/event-list.js";
import { microsoftOutlookSubscriptionCreate } from "./actions/subscription-create.js";
import { microsoftOutlookSubscriptionDelete } from "./actions/subscription-delete.js";
import { microsoftOutlookSubscriptionRenew } from "./actions/subscription-renew.js";
import { MICROSOFT_OUTLOOK_CALENDAR_CHANGED } from "./triggers/calendar-changed.js";

export const microsoftOutlookEventGetExecutor: NodeExecutor = async (ctx) => {
  const config = ((ctx.node.data as { config?: Record<string, unknown> })?.config ?? {});

  const result = await microsoftOutlookEventGet({
    config,
    input: ctx.inputs?.in,
  });

  ctx.emit({
    type: "log",
    level: "info",
    nodeId: ctx.node.id,
    message: `microsoft_outlook event_get ${(result.data as { id?: string })?.id} (${result.mode})`,
  });

  return { __port: "out", value: result };
};

export const microsoftOutlookEventListExecutor: NodeExecutor = async (ctx) => {
  const config = ((ctx.node.data as { config?: Record<string, unknown> })?.config ?? {});

  const result = await microsoftOutlookEventList({
    config,
    input: ctx.inputs?.in,
  });

  ctx.emit({
    type: "log",
    level: "info",
    nodeId: ctx.node.id,
    message: `microsoft_outlook event_list ${(result.data as { id?: string })?.id} (${result.mode})`,
  });

  return { __port: "out", value: result };
};

export const microsoftOutlookSubscriptionCreateExecutor: NodeExecutor = async (ctx) => {
  const config = ((ctx.node.data as { config?: Record<string, unknown> })?.config ?? {});

  const result = await microsoftOutlookSubscriptionCreate({
    config,
    input: ctx.inputs?.in,
  });

  ctx.emit({
    type: "log",
    level: "info",
    nodeId: ctx.node.id,
    message: `microsoft_outlook subscription_create ${(result.data as { id?: string })?.id} (${result.mode})`,
  });

  return { __port: "out", value: result };
};

export const microsoftOutlookSubscriptionDeleteExecutor: NodeExecutor = async (ctx) => {
  const config = ((ctx.node.data as { config?: Record<string, unknown> })?.config ?? {});

  const result = await microsoftOutlookSubscriptionDelete({
    config,
    input: ctx.inputs?.in,
  });

  ctx.emit({
    type: "log",
    level: "info",
    nodeId: ctx.node.id,
    message: `microsoft_outlook subscription_delete ${(result.data as { id?: string })?.id} (${result.mode})`,
  });

  return { __port: "out", value: result };
};

export const microsoftOutlookSubscriptionRenewExecutor: NodeExecutor = async (ctx) => {
  const config = ((ctx.node.data as { config?: Record<string, unknown> })?.config ?? {});

  const result = await microsoftOutlookSubscriptionRenew({
    config,
    input: ctx.inputs?.in,
  });

  ctx.emit({
    type: "log",
    level: "info",
    nodeId: ctx.node.id,
    message: `microsoft_outlook subscription_renew ${(result.data as { id?: string })?.id} (${result.mode})`,
  });

  return { __port: "out", value: result };
};

export const microsoftOutlookCalendarChangedTriggerExecutor: NodeExecutor = async (ctx) => {
  const config = ((ctx.node.data as { config?: Record<string, unknown> })?.config ?? {});
  const connection = resolveConnection({
    service: MICROSOFT_OUTLOOK.service,
    operation: "calendar_changed",
    sandbox: MICROSOFT_OUTLOOK.sandbox,
    baseUrls: MICROSOFT_OUTLOOK.baseUrls,
    requires: MICROSOFT_OUTLOOK.requires,
    connectionId: typeof config.connection === "string" ? config.connection : null,
    requested: typeof config.mode === "string" ? (config.mode as RequestedMode) : null,
  });

  const event = triggerEvent(MICROSOFT_OUTLOOK_CALENDAR_CHANGED, connection, ctx.inputs?.in, config);

  return { __port: "out", value: event };
};

/** The kinds a TypeScript host registers. */
export const MICROSOFT_OUTLOOK_RUNNABLE_KINDS: NodeKindDefinition[] = [
  { ...microsoftOutlookEventGetKind, executor: microsoftOutlookEventGetExecutor },
  { ...microsoftOutlookEventListKind, executor: microsoftOutlookEventListExecutor },
  { ...microsoftOutlookSubscriptionCreateKind, executor: microsoftOutlookSubscriptionCreateExecutor },
  { ...microsoftOutlookSubscriptionDeleteKind, executor: microsoftOutlookSubscriptionDeleteExecutor },
  { ...microsoftOutlookSubscriptionRenewKind, executor: microsoftOutlookSubscriptionRenewExecutor },
  { ...microsoftOutlookCalendarChangedTriggerKind, executor: microsoftOutlookCalendarChangedTriggerExecutor },
];
