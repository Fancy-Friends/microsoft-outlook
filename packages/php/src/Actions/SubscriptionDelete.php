<?php

declare(strict_types=1);

namespace ParticleAcademy\MicrosoftOutlook\Actions;

use ParticleAcademy\MicrosoftOutlook\MicrosoftOutlook;
use ParticleAcademy\Connectors\ConnectorConfigException;

/*
 * GENERATED FILE — do not edit.
 *
 * Emitted from provider/actions/subscription-delete.json by weaver's generator.
 * A hand-edit here is destroyed by the next protocol sync, which is worse than
 * being rejected, because it works until it silently does not. Fix
 * provider/actions/subscription-delete.json (or weaver's template/) and regenerate:
 *
 *     npm run provider -- microsoft_outlook
 */
/**
 * Delete a change-notification subscription. The host's subscription machinery
 * calls this when a trigger is removed.
 *
 * DELETE /v1.0/subscriptions/{id} —
 * https://learn.microsoft.com/en-us/graph/api/subscription-delete?view=graph-rest-1.0
 *
 * This describes the request. The connector client resolves the connection,
 * picks the estate, and either calls Microsoft Outlook or calls the faker.
 */
final class SubscriptionDelete
{
    public const OPERATION = 'subscription_delete';
    public const METHOD = 'DELETE';
    public const PATH = '/v1.0/subscriptions/{id}';
    public const SIDE_EFFECTS = 'idempotent';

    /**
     * Build the form body for one call.
     *
     * Validation fails loudly and specifically here, rather than three frames
     * later as an "invalid request" from Microsoft Outlook.
     *
     * @param array<string,mixed> $config
     * An EMPTY body is `{}`, not `[]` — and PHP cannot tell those apart, because
     * both are `array()` and `json_encode` picks the list. So an empty one is
     * returned as an object. TypeScript and Python have no such ambiguity, which
     * is why this is a difference only the byte-parity suite can see.
     *
     * @return array<string,mixed>|\stdClass
     */
    public static function body(array $config): array|\stdClass
    {
        if (($config['id'] ?? null) === null || ($config['id'] ?? null) === '') {
            throw new ConnectorConfigException('subscription_delete: "id" is required (Subscription ID).');
        }

        $body = [];

        $body = $body === [] ? new \stdClass() : $body;
        return $body;
    }

    /**
     * The request path, with each config value URL-ENCODED into it.
     *
     * `PATH` above is the TEMPLATE, which is what the descriptor advertises;
     * this is what a caller sends. A value interpolated raw changes which URL
     * is called — a range like `Sheet1!A:B` or a sheet named `Q1/Q2` — and the
     * provider answers 404 about the document rather than about the encoding.
     *
     * @param array<string,mixed> $config
     */
    public static function path(array $config): string
    {
        return '/v1.0/subscriptions/'.rawurlencode((string) ($config['id'] ?? ''));
    }
}
