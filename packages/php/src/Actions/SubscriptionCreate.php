<?php

declare(strict_types=1);

namespace ParticleAcademy\MicrosoftOutlook\Actions;

use ParticleAcademy\MicrosoftOutlook\MicrosoftOutlook;
use ParticleAcademy\Connectors\ConnectorConfigException;

/*
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
 * This describes the request. The connector client resolves the connection,
 * picks the estate, and either calls Microsoft Outlook or calls the faker.
 */
final class SubscriptionCreate
{
    public const OPERATION = 'subscription_create';
    public const METHOD = 'POST';
    public const PATH = '/v1.0/subscriptions';
    public const SIDE_EFFECTS = 'unsafe-to-replay';

    /**
     * Build the JSON body for one call.
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
        if (($config['resource'] ?? null) === null || ($config['resource'] ?? null) === '') {
            throw new ConnectorConfigException('subscription_create: "resource" is required (Resource).');
        }

        if (($config['notificationUrl'] ?? null) === null || ($config['notificationUrl'] ?? null) === '') {
            throw new ConnectorConfigException('subscription_create: "notificationUrl" is required (Notification URL).');
        }

        if (($config['expirationDateTime'] ?? null) === null || ($config['expirationDateTime'] ?? null) === '') {
            throw new ConnectorConfigException('subscription_create: "expirationDateTime" is required (Expires at).');
        }

        if (($config['clientState'] ?? null) === null || ($config['clientState'] ?? null) === '') {
            throw new ConnectorConfigException('subscription_create: "clientState" is required (Client state).');
        }

        $body = [];

        $value = $config['changeType'] ?? null;
        if ($value !== null && $value !== '') {
            $body['changeType'] = (string) $value;
        }

        $value = $config['resource'] ?? null;
        $body['resource'] = (string) $value;

        $value = $config['notificationUrl'] ?? null;
        $body['notificationUrl'] = (string) $value;

        $value = $config['lifecycleNotificationUrl'] ?? null;
        if ($value !== null && $value !== '') {
            $body['lifecycleNotificationUrl'] = (string) $value;
        }

        $value = $config['expirationDateTime'] ?? null;
        $body['expirationDateTime'] = (string) $value;

        $value = $config['clientState'] ?? null;
        $body['clientState'] = (string) $value;

        $body = $body === [] ? new \stdClass() : $body;
        return $body;
    }
}
