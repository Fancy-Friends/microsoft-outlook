<?php

declare(strict_types=1);

namespace ParticleAcademy\MicrosoftOutlook;

use ParticleAcademy\Connectors\FakeValues;
use ParticleAcademy\Connectors\Mode;
use ParticleAcademy\Connectors\PreparedRequest;
use ParticleAcademy\Connectors\SandboxKind;
use ParticleAcademy\Connectors\ServiceDescriptor;

/*
 * GENERATED FILE — do not edit.
 *
 * Emitted from provider/manifest.json by weaver's generator.
 * A hand-edit here is destroyed by the next protocol sync, which is worse than
 * being rejected, because it works until it silently does not. Fix
 * provider/manifest.json (or weaver's template/) and regenerate:
 *
 *     npm run provider -- microsoft_outlook
 */
/**
 * Microsoft Outlook, as one service descriptor shared by every Microsoft
 * Outlook operation.
 *
 * The PHP twin of the js package's `src/service.ts`.
 *
 * ## The sandbox trap, written down where it is used
 *
 * Microsoft's test estate is a Microsoft 365 Developer Program sandbox: a
 * separate E5 tenant with its own users, its own calendars and its own OAuth
 * consent, on the SAME host. Eligibility is limited -- Visual Studio
 * Professional or Enterprise subscribers, ISV Success / Microsoft AI Cloud
 * Partner Program members, or Premier/Unified Support customers -- and the
 * tenant expires after 90 days of no development activity. Without one, every
 * call reaches a real mailbox.
 */
final class MicrosoftOutlook
{
    // The connector API version this package was GENERATED against. A
    // literal, never imported: an imported constant lets an upgrade rewrite
    // the very claim it exists to detect.
    public const CONNECTOR_API_VERSION = 1;

    public const SERVICE = 'microsoft_outlook';

    public const LIVE_URL = 'https://graph.microsoft.com';
    public const SANDBOX_URL = 'https://graph.microsoft.com';

    /** @var list<string> Credential keys a remote call cannot proceed without. */
    public const REQUIRES = [
        'accessToken',
        'refreshToken',
        'clientId',
        'clientSecret',
    ];

    public static function descriptor(): ServiceDescriptor
    {
        return new ServiceDescriptor(
            service: self::SERVICE,
            title: 'Microsoft Outlook',
            sandbox: SandboxKind::SeparateAccount,
            baseUrls: [
                Mode::Live->value => self::LIVE_URL,
                Mode::Sandbox->value => self::SANDBOX_URL,
            ],
            requires: self::REQUIRES,
            authorize: self::authorize(...),
            // The core calls a faker ($operation, $config, $fake, $input); respond()
            // takes TypeScript's FakeRequest shape. This is the translation.
            faker: static fn (string $operation, array $config, FakeValues $fake, mixed $input = null): mixed => MicrosoftOutlookFaker::respond(
                $operation,
                ['config' => $config, 'fake' => $fake, 'input' => $input],
            ),
        );
    }

    /**
     * Apply Microsoft Outlook's auth scheme to an outgoing request.
     *
     *
     *
     * @param array<string,string> $credentials
     */
    public static function authorize(array $credentials, PreparedRequest $request, Mode $mode): void
    {
        $request->withHeader('Authorization', 'Bearer '.($credentials['accessToken'] ?? ''));
    }
}
