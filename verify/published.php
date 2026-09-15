<?php

declare(strict_types=1);

/*
 * Microsoft Outlook — the published Composer package.
 *
 * GENERATED — do not edit. Fix weaver's template/ and regenerate.
 *
 * This runs against the PUBLISHED package, installed by name from the
 * registry into a project that has never seen this repo. Every other test
 * here imports from ../src and therefore cannot see the packaging.
 */

$autoload = getcwd().'/vendor/autoload.php';

if (! is_file($autoload)) {
    fwrite(STDERR, 'No vendor/autoload.php in '.getcwd().PHP_EOL);
    fwrite(STDERR, 'Run this from a project that has composer-required the published package:'.PHP_EOL);
    fwrite(STDERR, '    composer require particle-academy/microsoft-outlook-php'.PHP_EOL);
    exit(2);
}

require $autoload;

use ParticleAcademy\Connectors\FakeValues;
use ParticleAcademy\MicrosoftOutlook\MicrosoftOutlookFaker;

$goldens = [
    [
        'operation' => 'event_get',
        'config' => [],
        'expected' => [
            '@odata.context' => 'https://graph.microsoft.com/v1.0/$metadata#users(\'ada%40example.test\')/events/$entity',
            '@odata.etag' => 'W/"ZlnW4RIAV06KYYwlrfNZvQAALfZeRQ=="',
            'id' => 'AAMkAGI2_fake_4eccb4b0d301',
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
        ],
    ],
    [
        'operation' => 'event_list',
        'config' => [],
        'expected' => [
            '@odata.context' => 'https://graph.microsoft.com/v1.0/$metadata#users(\'ada%40example.test\')/events',
            'value' => [
                [
                    '@odata.etag' => 'W/"ZlnW4RIAV06KYYwlrfNZvQAALfZeRQ=="',
                    'id' => 'AAMkAGI2_fake_9a231b22d1a3',
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
        ],
    ],
    [
        'operation' => 'subscription_create',
        'config' => [],
        'expected' => [
            '@odata.context' => 'https://graph.microsoft.com/v1.0/$metadata#subscriptions/$entity',
            'id' => '4743be5e-2184-10d8-1aa8-c17076db3ac5',
            'resource' => 'me/events',
            'applicationId' => '8b47cca0-28ad-a735-7c8e-eb27866b66b7',
            'changeType' => 'created,updated,deleted',
            'clientState' => 'cs_fake_6621468790c1',
            'notificationUrl' => 'https://host.example.test/hooks/microsoft-outlook',
            'expirationDateTime' => '2026-09-22T18:23:45.9356913Z',
            'creatorId' => '223e4280-a7a8-4fcd-b2a5-bead3a3a24a4',
            'latestSupportedTlsVersion' => 'v1_2',
            'notificationContentType' => 'application/json',
        ],
    ],
    [
        'operation' => 'subscription_delete',
        'config' => [],
        'expected' => [],
    ],
    [
        'operation' => 'subscription_renew',
        'config' => [],
        'expected' => [
            '@odata.context' => 'https://graph.microsoft.com/v1.0/$metadata#subscriptions/$entity',
            'id' => '5ceaa0a0-172a-4973-7579-dcfdbb6519b4',
            'resource' => 'me/events',
            'applicationId' => '5c761487-4231-2970-f5fa-2023840644ce',
            'changeType' => 'created,updated,deleted',
            'notificationUrl' => 'https://host.example.test/hooks/microsoft-outlook',
            'expirationDateTime' => '2026-09-29T18:23:45.9356913Z',
            'latestSupportedTlsVersion' => 'v1_2',
            'notificationContentType' => 'application/json',
        ],
    ],
    [
        'operation' => 'calendar_changed',
        'config' => [
            'sample' => 'updated',
        ],
        'expected' => [
            'value' => [
                [
                    'subscriptionId' => 'a15234ea-1e2f-bdbc-a4dd-0d03f54606ca',
                    'subscriptionExpirationDateTime' => '2026-09-22T18:23:45.9356913Z',
                    'changeType' => 'updated',
                    'resource' => 'Users/ada%40example.test/Events/AAMkAGI2',
                    'tenantId' => '905d610b-b3ff-d562-2f44-4e30083ca9c2',
                    'clientState' => 'verified-before-injection',
                    'resourceData' => [
                        '@odata.type' => '#Microsoft.Graph.Event',
                        '@odata.id' => 'Users/ada%40example.test/Events/AAMkAGI2',
                        '@odata.etag' => 'W/"ZlnW4RIAV06KYYwlrfNZvQAALfZeRQ=="',
                        'id' => 'AAMkAGI2_fake_f78e5b59a049',
                    ],
                ],
            ],
        ],
    ],
];

foreach ($goldens as $golden) {
    $operation = $golden['operation'];
    $config = $golden['config'];

    $fake = new FakeValues(FakeValues::seedForCall('microsoft_outlook', $operation, $config));
    $faked = MicrosoftOutlookFaker::respond($operation, ['config' => $config, 'fake' => $fake]);

    if ($faked !== $golden['expected']) {
        fwrite(STDERR, "the PUBLISHED package produced different bytes for {$operation}\n");
        fwrite(STDERR, '  got:      '.json_encode($faked)."\n");
        fwrite(STDERR, '  expected: '.json_encode($golden['expected'])."\n");
        exit(1);
    }

    echo "  ok   {$operation}\n";
}

echo "\n  ".count($goldens)." operations verified against the published package.\n";
