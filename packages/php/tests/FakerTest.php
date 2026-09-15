<?php

declare(strict_types=1);

use ParticleAcademy\MicrosoftOutlook\MicrosoftOutlookFaker;
use ParticleAcademy\Connectors\FakeValues;

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
 * The golden fixtures — the SAME values the TypeScript and Python packages
 * assert.
 *
 * Bit-for-bit identical is the claim, and this is what checks it.
 * Cross-runtime drift does not fail loudly on its own: it completes, down one
 * path, with no error.
 */

it('event_get fakes the shape Microsoft Outlook publishes', function () {
    $config = [];
    $fake = new FakeValues(FakeValues::seedForCall('microsoft_outlook', 'event_get', $config));

    $faked = MicrosoftOutlookFaker::respond('event_get', ['config' => $config, 'fake' => $fake]);

    expect($faked)->toBe([
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
    ]);
});

it('event_list fakes the shape Microsoft Outlook publishes', function () {
    $config = [];
    $fake = new FakeValues(FakeValues::seedForCall('microsoft_outlook', 'event_list', $config));

    $faked = MicrosoftOutlookFaker::respond('event_list', ['config' => $config, 'fake' => $fake]);

    expect($faked)->toBe([
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
    ]);
});

it('subscription_create fakes the shape Microsoft Outlook publishes', function () {
    $config = [];
    $fake = new FakeValues(FakeValues::seedForCall('microsoft_outlook', 'subscription_create', $config));

    $faked = MicrosoftOutlookFaker::respond('subscription_create', ['config' => $config, 'fake' => $fake]);

    expect($faked)->toBe([
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
    ]);
});

it('subscription_delete fakes the shape Microsoft Outlook publishes', function () {
    $config = [];
    $fake = new FakeValues(FakeValues::seedForCall('microsoft_outlook', 'subscription_delete', $config));

    $faked = MicrosoftOutlookFaker::respond('subscription_delete', ['config' => $config, 'fake' => $fake]);

    expect($faked)->toBe([]);
});

it('subscription_renew fakes the shape Microsoft Outlook publishes', function () {
    $config = [];
    $fake = new FakeValues(FakeValues::seedForCall('microsoft_outlook', 'subscription_renew', $config));

    $faked = MicrosoftOutlookFaker::respond('subscription_renew', ['config' => $config, 'fake' => $fake]);

    expect($faked)->toBe([
        '@odata.context' => 'https://graph.microsoft.com/v1.0/$metadata#subscriptions/$entity',
        'id' => '5ceaa0a0-172a-4973-7579-dcfdbb6519b4',
        'resource' => 'me/events',
        'applicationId' => '5c761487-4231-2970-f5fa-2023840644ce',
        'changeType' => 'created,updated,deleted',
        'notificationUrl' => 'https://host.example.test/hooks/microsoft-outlook',
        'expirationDateTime' => '2026-09-29T18:23:45.9356913Z',
        'latestSupportedTlsVersion' => 'v1_2',
        'notificationContentType' => 'application/json',
    ]);
});

it('calendar_changed fakes the shape Microsoft Outlook publishes', function () {
    $config = [
        'sample' => 'updated',
    ];
    $fake = new FakeValues(FakeValues::seedForCall('microsoft_outlook', 'calendar_changed', $config));

    $faked = MicrosoftOutlookFaker::respond('calendar_changed', ['config' => $config, 'fake' => $fake]);

    expect($faked)->toBe([
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
    ]);
});

it('throws for an operation with no fixture rather than inventing a shape', function () {
    $fake = new FakeValues(FakeValues::seedForCall('microsoft_outlook', 'no_such_operation', []));

    expect(fn () => MicrosoftOutlookFaker::respond('no_such_operation', ['config' => [], 'fake' => $fake]))
        ->toThrow(InvalidArgumentException::class);
});
