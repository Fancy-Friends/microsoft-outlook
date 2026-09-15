# Changelog

All notable changes to `@particle-academy/microsoft-outlook-ui`,
`@particle-academy/microsoft-outlook-js`, `particle-academy/microsoft-outlook-php`
and `fancy-microsoft-outlook`.

The four packages share one version, because they are generated from one
`provider/` definition and a version that meant something different in each
would be a version nobody could reason about.

## [0.1.0] — 2026-09-15

### Added

- **First release.** The calendar half of Outlook, on Microsoft Graph v1.0:
  actions `event_get`, `event_list`, `subscription_create`,
  `subscription_renew` and `subscription_delete`; the `calendar_changed`
  trigger with `delivery: "subscription"` — a Graph subscription capped at
  10080 minutes, renewed by `subscription_renew` 3600 seconds before expiry,
  its lease read from `expirationDateTime`. Deliveries are batches, verified
  by the `clientState` Graph echoes in every item, and the endpoint answers
  Graph's `validationToken` challenge before anything is delivered.
- Requires `fancy-connector-core` ≥ 0.8.0 in js and php — the release that
  carries `LeaseDeclaration`, the shared-token verification scheme and the
  challenge handshake.
