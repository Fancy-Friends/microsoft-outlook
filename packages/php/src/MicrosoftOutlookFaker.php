<?php

declare(strict_types=1);

namespace ParticleAcademy\MicrosoftOutlook;

use ParticleAcademy\Connectors\FakeRequest;

/*
 * GENERATED FILE — do not edit.
 *
 * Emitted from provider/fixtures/ by weaver's generator.
 * A hand-edit here is destroyed by the next protocol sync, which is worse than
 * being rejected, because it works until it silently does not. Fix
 * provider/fixtures/ (or weaver's template/) and regenerate:
 *
 *     npm run provider -- microsoft_outlook
 */
/**
 * The Microsoft Outlook faker — the PHP twin of the js package's
 * `src/faker.ts`.
 *
 * Bit-for-bit identical: the same FNV-1a seed and the same xorshift32
 * sequence, so a golden fixture asserts the exact faked payload and BOTH
 * runtimes have to produce it. That turns the faker into a parity test rather
 * than a convenience.
 */
final class MicrosoftOutlookFaker
{
    /** @param array<string,mixed> $request */
    public static function respond(string $operation, array $request): mixed
    {
        /** @var array<string,mixed> $config */
        $config = $request['config'] ?? [];
        /** @var FakeValuesLike $fake */
        $fake = $request['fake'];

        return match ($operation) {
            'event_get' => self::EventGet($config, $fake),
            'event_list' => self::EventList($config, $fake),
            'subscription_create' => self::SubscriptionCreate($config, $fake),
            'subscription_delete' => self::SubscriptionDelete($config, $fake),
            'subscription_renew' => self::SubscriptionRenew($config, $fake),
            'calendar_changed' => self::CalendarChanged($config, $fake),
            default => throw new \InvalidArgumentException(
                // A faker asked for an operation it has no shape for must SAY so.
                // Making something up would produce a green run whose output
                // silently has none of the fields the author is about to reference.
                'microsoft_outlook: no fake response is defined for "'.$operation.'". '
                    .'Add a fixture under provider/fixtures/ and regenerate — a connector without a faker '
                    .'cannot be developed against, tested, or demonstrated.'
            ),
        };
    }

    /** @param array<string,mixed> $config */
    private static function EventGet(array $config, mixed $fake): array|\stdClass
    {
        $boundId = ((($v = $config['id'] ?? null) !== null && $v !== '') ? (string) $v : $fake->id('AAMkAGI2'));

        return [
        '@odata.context' => 'https://graph.microsoft.com/v1.0/$metadata#users(\'ada%40example.test\')/events/$entity',
        '@odata.etag' => 'W/"ZlnW4RIAV06KYYwlrfNZvQAALfZeRQ=="',
        'id' => $boundId,
        'createdDateTime' => '2026-09-01T09:00:00.0000000Z',
        'lastModifiedDateTime' => '2026-09-10T15:30:00.0000000Z',
        'subject' => 'Design review',
        'bodyPreview' => 'Quarterly design review with the platform team.',
        'isCancelled' => false,
        'isAllDay' => false,
        'showAs' => 'busy',
        'webLink' => 'https://outlook.office365.com/owa/?itemid=AAMkAGI2&exvsurl=1&path=/calendar/item',
        'start' => [
            'dateTime' => '2026-09-22T14:00:00.0000000',
            'timeZone' => 'UTC',
        ],
        'end' => [
            'dateTime' => '2026-09-22T15:00:00.0000000',
            'timeZone' => 'UTC',
        ],
        'location' => [
            'displayName' => 'Room 4',
        ],
        'organizer' => [
            'emailAddress' => [
                'name' => 'Ada Example',
                'address' => 'ada@example.test',
            ],
        ],
        'attendees' => [
            [
                'type' => 'required',
                'status' => [
                    'response' => 'accepted',
                    'time' => '2026-09-02T10:00:00.0000000Z',
                ],
                'emailAddress' => [
                    'name' => 'Grace Example',
                    'address' => 'grace@example.test',
                ],
            ],
        ],
    ];
    }

    /** @param array<string,mixed> $config */
    private static function EventList(array $config, mixed $fake): array|\stdClass
    {
        return [
        '@odata.context' => 'https://graph.microsoft.com/v1.0/$metadata#users(\'ada%40example.test\')/events',
        'value' => [
            [
                '@odata.etag' => 'W/"ZlnW4RIAV06KYYwlrfNZvQAALfZeRQ=="',
                'id' => $fake->id('AAMkAGI2'),
                'createdDateTime' => '2026-09-01T09:00:00.0000000Z',
                'lastModifiedDateTime' => '2026-09-10T15:30:00.0000000Z',
                'subject' => 'Design review',
                'bodyPreview' => 'Quarterly design review with the platform team.',
                'isCancelled' => false,
                'isAllDay' => false,
                'showAs' => 'busy',
                'webLink' => 'https://outlook.office365.com/owa/?itemid=AAMkAGI2&exvsurl=1&path=/calendar/item',
                'start' => [
                    'dateTime' => '2026-09-22T14:00:00.0000000',
                    'timeZone' => 'UTC',
                ],
                'end' => [
                    'dateTime' => '2026-09-22T15:00:00.0000000',
                    'timeZone' => 'UTC',
                ],
                'organizer' => [
                    'emailAddress' => [
                        'name' => 'Ada Example',
                        'address' => 'ada@example.test',
                    ],
                ],
            ],
        ],
    ];
    }

    /** @param array<string,mixed> $config */
    private static function SubscriptionCreate(array $config, mixed $fake): array|\stdClass
    {
        $boundResource = ((($v = $config['resource'] ?? null) !== null && $v !== '') ? (string) $v : 'me/events');
        $boundChangetype = ((($v = $config['changeType'] ?? null) !== null && $v !== '') ? (string) $v : 'created,updated,deleted');
        $boundNotificationurl = ((($v = $config['notificationUrl'] ?? null) !== null && $v !== '') ? (string) $v : 'https://host.example.test/hooks/microsoft-outlook');
        $boundClientstate = ((($v = $config['clientState'] ?? null) !== null && $v !== '') ? (string) $v : $fake->id('cs'));

        return [
        '@odata.context' => 'https://graph.microsoft.com/v1.0/$metadata#subscriptions/$entity',
        'id' => $fake->hex(8).'-'.$fake->hex(4).'-'.$fake->hex(4).'-'.$fake->hex(4).'-'.$fake->hex(12),
        'resource' => $boundResource,
        'applicationId' => $fake->hex(8).'-'.$fake->hex(4).'-'.$fake->hex(4).'-'.$fake->hex(4).'-'.$fake->hex(12),
        'changeType' => $boundChangetype,
        'clientState' => $boundClientstate,
        'notificationUrl' => $boundNotificationurl,
        'expirationDateTime' => '2026-09-22T18:23:45.9356913Z',
        'creatorId' => $fake->hex(8).'-'.$fake->hex(4).'-'.$fake->hex(4).'-'.$fake->hex(4).'-'.$fake->hex(12),
        'latestSupportedTlsVersion' => 'v1_2',
        'notificationContentType' => 'application/json',
    ];
    }

    /** @param array<string,mixed> $config */
    private static function SubscriptionDelete(array $config, mixed $fake): array|\stdClass
    {
        return new \stdClass();
    }

    /** @param array<string,mixed> $config */
    private static function SubscriptionRenew(array $config, mixed $fake): array|\stdClass
    {
        $boundId = ((($v = $config['id'] ?? null) !== null && $v !== '') ? (string) $v : $fake->hex(8).'-'.$fake->hex(4).'-'.$fake->hex(4).'-'.$fake->hex(4).'-'.$fake->hex(12));
        $boundExpirationdatetime = ((($v = $config['expirationDateTime'] ?? null) !== null && $v !== '') ? (string) $v : '2026-09-29T18:23:45.9356913Z');

        return [
        '@odata.context' => 'https://graph.microsoft.com/v1.0/$metadata#subscriptions/$entity',
        'id' => $boundId,
        'resource' => 'me/events',
        'applicationId' => $fake->hex(8).'-'.$fake->hex(4).'-'.$fake->hex(4).'-'.$fake->hex(4).'-'.$fake->hex(12),
        'changeType' => 'created,updated,deleted',
        'notificationUrl' => 'https://host.example.test/hooks/microsoft-outlook',
        'expirationDateTime' => $boundExpirationdatetime,
        'latestSupportedTlsVersion' => 'v1_2',
        'notificationContentType' => 'application/json',
    ];
    }

    /** @param array<string,mixed> $config */
    private static function CalendarChanged(array $config, mixed $fake): array|\stdClass
    {
        $boundSubscriptionid = $fake->hex(8).'-'.$fake->hex(4).'-'.$fake->hex(4).'-'.$fake->hex(4).'-'.$fake->hex(12);
        $boundEventid = $fake->id('AAMkAGI2');

        return [
        'value' => [
            [
                'subscriptionId' => $boundSubscriptionid,
                'subscriptionExpirationDateTime' => '2026-09-22T18:23:45.9356913Z',
                'changeType' => ((($v = $config['sample'] ?? null) !== null && $v !== '') ? (string) $v : 'updated'),
                'resource' => 'Users/ada%40example.test/Events/AAMkAGI2',
                'tenantId' => $fake->hex(8).'-'.$fake->hex(4).'-'.$fake->hex(4).'-'.$fake->hex(4).'-'.$fake->hex(12),
                'clientState' => 'verified-before-injection',
                'resourceData' => [
                    '@odata.type' => '#Microsoft.Graph.Event',
                    '@odata.id' => 'Users/ada%40example.test/Events/AAMkAGI2',
                    '@odata.etag' => 'W/"ZlnW4RIAV06KYYwlrfNZvQAALfZeRQ=="',
                    'id' => $boundEventid,
                ],
            ],
        ],
    ];
    }
}
