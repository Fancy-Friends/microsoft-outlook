# Microsoft Outlook

[![Fancified](art/fancified.svg)](https://particle.academy)

Microsoft Outlook for [fancy-flow][flow] — as **four imported, versioned packages**, one
per runtime. Not vendored source: a copy cannot be upgraded, and third-party APIs
change.

[flow]: https://github.com/Particle-Academy/fancy-flow

| Runtime | Package | Install |
|---|---|---|
| Authoring surface (every host) | `@particle-academy/microsoft-outlook-ui` | `npm install @particle-academy/microsoft-outlook-ui` |
| Node | `@particle-academy/microsoft-outlook-js` | `npm install @particle-academy/microsoft-outlook-js` |
| PHP 8.4+ | `particle-academy/microsoft-outlook-php` | `composer require particle-academy/microsoft-outlook-php` |
| Python 3.11+ | `fancy-microsoft-outlook` | `pip install fancy-microsoft-outlook` |

The `ui` package is the editor surface and is React on every host — a PHP or
Python project installs it *and* its own runtime package, and never the `js` one.

## What it costs you

One dependency: `@particle-academy/fancy-connector-core` (or
`particle-academy/fancy-connector-core` on Composer), which the `js` and `php`
packages pull in themselves. The Python package has **zero** runtime
dependencies.

**No Microsoft Outlook SDK.** Plain HTTP, deliberately: a vendor SDK is third-party code
subject to the kit's full approval bar, and one per provider is hundreds of
dependencies nobody is tracking.

## Setting it up

Everything below is generated from `provider/manifest.json`, so it cannot disagree with what the packages do.

### Credentials

A Microsoft Outlook connection holds 5 values.

**Two kinds of value, and mixing them up matters.** A `provider` credential is ONE value for the whole installation — an OAuth app's client secret serves every connected account. An `account` credential is one per connected account. A host that stores the second where it stores the first lets one account's credentials reach another's.

| Field | Scope | Secret | Where it comes from |
|---|---|---|---|
| **Application (client) ID** | per installation | not secret | From the Microsoft Entra admin center -> App registrations. ONE value for the whole installation, not per connected account. |
| **Client secret** | per installation | **secret** | A client secret of the same app registration. One value for the whole installation; Entra expires them, so record the expiry. |
| **Access token** | per connected account | **secret** | Per connected Microsoft account, and it expires after about ONE HOUR. The host refreshes it with the refresh token. |
| **Refresh token** | per connected account | **secret** | Per connected Microsoft account; issued only when the consent request includes offline_access. Every refresh answers with a NEW one -- persist it -- while the old one keeps working until it expires. |
| **Client state** | per connected account | **secret** | A secret the host chooses per connected account and sends as `clientState` on every subscription. Graph echoes it inside every notification, and that echo is the ONLY thing that authenticates a delivery -- Graph does not sign them. At most 128 characters. |

### Authorising

Microsoft Outlook uses OAuth2 (authorization_code). The package DECLARES the exchange; the HOST performs it — a consent screen needs a browser, a redirect URI and somewhere to persist the result, and all three belong to the host.

- **Authorize URL** — https://login.microsoftonline.com/common/oauth2/v2.0/authorize
- **Token URL** — https://login.microsoftonline.com/common/oauth2/v2.0/token
- **Scopes** — `Calendars.Read`, `offline_access`
- **Access token lifetime** — 3600 seconds (1 hours). A host that never refreshes works all afternoon and is broken by morning.

The refresh tokens do **not** rotate: the same one is reusable, so a refresh may safely be retried and may run concurrently. Stated rather than assumed, because the opposite — a provider that spends the token and revokes the grant on a replay — looks identical until it happens.

### The estate

Microsoft Outlook has a test estate on the same host, reached with credentials from a SEPARATE test account you register. Selecting sandbox mode uses those credentials.

> Microsoft's test estate is a Microsoft 365 Developer Program sandbox: a separate E5 tenant with its own users, its own calendars and its own OAuth consent, on the SAME host. Eligibility is limited -- Visual Studio Professional or Enterprise subscribers, ISV Success / Microsoft AI Cloud Partner Program members, or Premier/Unified Support customers -- and the tenant expires after 90 days of no development activity. Without one, every call reaches a real mailbox.

## What it can do

### Actions

#### `event_get` — Outlook calendar event

Read one event from the connected account's Outlook calendar.

`GET /v1.0/me/events/{id}` · reads only — safe to replay

| Input | Required | What it is |
|---|---|---|
| `id` | yes | The event's id, as published by a notification's resourceData.id or by event_list. |

#### `event_list` — Outlook calendar events

List events from the connected account's Outlook calendar.

`GET /v1.0/me/events` · reads only — safe to replay

| Input | Required | What it is |
|---|---|---|
| `top` | no | Events per page. An @odata.nextLink in the response means there are more. |
| `filter` | no | An OData $filter. Leave blank for everything. |
| `orderby` | no | An OData $orderby. Leave blank for Graph's default order. |
| `select` | no | An OData $select, comma separated. Leave blank for Graph's default set. |

#### `subscription_create` — Outlook subscription

Subscribe to change notifications on the connected account's calendar events. The host's subscription machinery calls this; it is not a node most workflows need.

`POST /v1.0/subscriptions` · **unsafe to replay** — a retried durable run does it TWICE

| Input | Required | What it is |
|---|---|---|
| `changeType` | no | Which changes raise a notification. Graph accepts a comma-separated combination. |
| `resource` | yes | The Graph resource to watch, without the base URL. `me/events` is the connected account's calendar events. |
| `notificationUrl` | yes | The HTTPS URL the host mounts for this trigger. Graph validates it with a validationToken challenge before creating the subscription, then POSTs change notifications to it. |
| `lifecycleNotificationUrl` | no | Where Graph sends `reauthorizationRequired`, `subscriptionRemoved` and `missed` events. May be the notification URL itself. Optional, but without it a subscription Graph removed is one nobody hears about. |
| `expirationDateTime` | yes | When the subscription expires, as an ISO 8601 instant in UTC. The host computes now + the trigger's lifetime; Graph caps an event subscription at 10080 minutes (under seven days) and raises anything under 45 minutes to 45. |
| `clientState` | yes | The connection's clientState. The host fills this from the connection -- never type a value here. Graph echoes it inside every notification, and that echo is how a delivery is verified. At most 128 characters. |

#### `subscription_delete` — Outlook subscription removal

Delete a change-notification subscription. The host's subscription machinery calls this when a trigger is removed.

`DELETE /v1.0/subscriptions/{id}` · idempotent — safe to replay

| Input | Required | What it is |
|---|---|---|
| `id` | yes | The id subscription_create answered with. |

#### `subscription_renew` — Outlook subscription renewal

Extend a change-notification subscription before it expires. The host's subscription machinery calls this when the lease is due.

`PATCH /v1.0/subscriptions/{id}` · idempotent — safe to replay

| Input | Required | What it is |
|---|---|---|
| `id` | yes | The id subscription_create answered with. |
| `expirationDateTime` | yes | The new expiry, as an ISO 8601 instant in UTC. The host computes now + the trigger's lifetime, capped at 10080 minutes for events. |

### Triggers

#### `calendar_changed` — Outlook calendar change

Start a run when an event is created, updated or deleted on a subscribed Outlook calendar.

A SUBSCRIPTION: Microsoft Outlook delivers to your endpoint only while a subscription it issued is alive, at most 168 hours at a time. Its expiry is read from the `subscription_create` response (`expirationDateTime`); 3600 seconds before it, the host renews it (`subscription_renew`). A lease that lapses is re-listed AND re-subscribed, because notifications during the gap are gone. Every delivery: the token Microsoft Outlook echoes back is compared with the connection's `clientState` before anything runs. Microsoft Outlook sends a `validationToken` challenge before it delivers anything, and the endpoint must echo it as plain text.

**You have to set this up with the provider first:**

The host calls subscription_create with its own URL for this trigger, an expiry of now plus this lifetime, and the connection's clientState, and must answer Graph's validationToken challenge on that URL as plain text within 10 seconds or the subscription is refused. It stores the returned id and expirationDateTime. 3600 seconds before expiry it calls subscription_renew; a 404 there means the subscription is gone and it re-lists (event_list) and creates again, because notifications during the gap are gone. Every notification is a batch: the host compares each item's clientState with the connection's clientState, answers 202, and then processes the items. A `missed` lifecycle event means re-list; `reauthorizationRequired` means renew; `subscriptionRemoved` means create again and re-list.

## Run it before you have credentials

Every operation ships a **faker**, whether or not Microsoft Outlook has a sandbox. Set a
node's mode to `fake` and it returns the shape Microsoft Outlook actually publishes — the
same field names, deterministically — so you can wire the downstream nodes before
touching an account, a key, or a network.

## This repository is generated

`provider/` is the source. Everything under `packages/` is emitted from it and
**must not be hand-edited** — CI regenerates and diffs on every push, and the
next protocol sync destroys anything it finds. See [`AGENTS.md`](AGENTS.md).

## Two namespaces, which do not match on purpose

The repo is `github.com/Fancy-Friends/microsoft-outlook`; the packages publish under
`particle-academy`. Nothing derives one from the other — the names come from
weaver's `friends.json` and nowhere else.

## Licence

MIT.
