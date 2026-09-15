<?php

declare(strict_types=1);

namespace ParticleAcademy\MicrosoftOutlook\Actions;

use ParticleAcademy\MicrosoftOutlook\MicrosoftOutlook;
use ParticleAcademy\Connectors\ConnectorConfigException;

/*
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
 * This describes the request. The connector client resolves the connection,
 * picks the estate, and either calls Microsoft Outlook or calls the faker.
 */
final class EventList
{
    public const OPERATION = 'event_list';
    public const METHOD = 'GET';
    public const PATH = '/v1.0/me/events';
    public const SIDE_EFFECTS = 'none';

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
        $top = $config['top'] ?? null;
        if (($top !== null && $top !== '') && ! (is_numeric($top) && (float) $top === floor((float) $top) && (float) $top >= 1 && (float) $top <= 999)) {
            throw new ConnectorConfigException(
                'event_list: "top" must be a integer, got '.json_encode($top).'.'
            );
        }

        $body = [];

        $value = $config['top'] ?? null;
        $body['$top'] = ($value !== null && $value !== '') ? (int) $value : 50;

        $value = $config['filter'] ?? null;
        if ($value !== null && $value !== '') {
            $body['$filter'] = (string) $value;
        }

        $value = $config['orderby'] ?? null;
        if ($value !== null && $value !== '') {
            $body['$orderby'] = (string) $value;
        }

        $value = $config['select'] ?? null;
        if ($value !== null && $value !== '') {
            $body['$select'] = (string) $value;
        }

        $body = $body === [] ? new \stdClass() : $body;
        return $body;
    }
}
