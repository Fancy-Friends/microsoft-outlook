<?php

declare(strict_types=1);

namespace ParticleAcademy\MicrosoftOutlook\Triggers;

use ParticleAcademy\Connectors\DeliveryMechanism;
use ParticleAcademy\Connectors\ExpiresAtUnit;
use ParticleAcademy\Connectors\LeaseDeclaration;
use ParticleAcademy\Connectors\WebhookVerifier;

/*
 * GENERATED FILE — do not edit.
 *
 * Emitted from provider/triggers/calendar-changed.json by weaver's generator.
 * A hand-edit here is destroyed by the next protocol sync, which is worse than
 * being rejected, because it works until it silently does not. Fix
 * provider/triggers/calendar-changed.json (or weaver's template/) and regenerate:
 *
 *     npm run provider -- microsoft_outlook
 */
/**
 * Microsoft Outlook's subscription trigger — the delivery contract.
 *
 * Kept beside the service descriptor rather than inside a node, because the
 * way a delivery is verified is a fact about MICROSOFT OUTLOOK. The twin of
 * the js package's trigger module.
 *
 * A SUBSCRIPTION: Microsoft Outlook stops delivering unless somebody renews
 * it, forever, and if nobody does the workflow stops firing with no error
 * anywhere. `lease()` says where the expiry is read from and how early to
 * renew; the host builds the value with `SubscriptionLease::fromResponse()`
 * and runs ONE renewal scheduler for every expiring trigger.
 */
final class CalendarChanged
{
    public const OPERATION = 'calendar_changed';
    public const DELIVERY = DeliveryMechanism::Subscription;

    public const SETUP = 'The host calls subscription_create with its own URL for this trigger, an expiry of now plus this lifetime, and the connection\'s clientState, and must answer Graph\'s validationToken challenge on that URL as plain text within 10 seconds or the subscription is refused. It stores the returned id and expirationDateTime. 3600 seconds before expiry it calls subscription_renew; a 404 there means the subscription is gone and it re-lists (event_list) and creates again, because notifications during the gap are gone. Every notification is a batch: the host compares each item\'s clientState with the connection\'s clientState, answers 202, and then processes the items. A `missed` lifecycle event means re-list; `reauthorizationRequired` means renew; `subscriptionRemoved` means create again and re-list.';

    /** How long Microsoft Outlook keeps a subscription alive at most, in seconds. */
    public const SUBSCRIPTION_TTL = 604800;

    /** Which of this package's actions create, renew and stop the subscription. */
    public const CREATE_OPERATION = 'subscription_create';

    /** The action the host calls when the lease is due. */
    public const RENEW_OPERATION = 'subscription_renew';

    public const STOP_OPERATION = 'subscription_delete';

    /**
     * The lease this subscription carries: where the provider's expiry sits in the
     * `subscription_create` response (`expirationDateTime`, rfc3339), how early
     * the host renews, and what it calls when the lease is due.
     */
    public static function lease(): LeaseDeclaration
    {
        return new LeaseDeclaration(
            'expirationDateTime',
            ExpiresAtUnit::Rfc3339,
            3600,
            'subscription_renew',
        );
    }

    /** Where Microsoft Outlook echoes the token it was given when the subscription was created. */
    public const TOKEN_IN = 'body';

    /** A dotted path into the JSON body; `[]` means every element of the batch, all of which must match. */
    public const TOKEN_NAME = 'value[].clientState';

    /** The credential holding the token. */
    public const SECRET_CREDENTIAL = 'clientState';

    /**
     * Verify one inbound Microsoft Outlook delivery.
     *
     * The host calls this BEFORE starting a run, with the body exactly as
     * received. The token is the connection's `clientState`; `$now` is accepted
     * for symmetry with signed schemes and unused, because an echoed token carries
     * no timestamp.
     *
     * @param array<string,string|list<string>> $headers
     * @return array{ok: bool, reason: ?string}
     */
    public static function verifyDelivery(
        string $raw,
        array $headers,
        ?string $clientState,
        ?int $now = null,
    ): array {
        unset($now);

        return WebhookVerifier::verifySharedToken(
            raw: $raw,
            headers: $headers,
            secret: $clientState,
            in: self::TOKEN_IN,
            name: self::TOKEN_NAME,
        );
    }

    /** The query parameter Microsoft Outlook sends as its challenge before it delivers anything. */
    public const HANDSHAKE_PARAM = 'validationToken';

    /**
     * Answer Microsoft Outlook's challenge — what to send back (200, text/plain,
     * the decoded `validationToken`), or null when the request is not a challenge
     * at all. The host's routing layer asks this before mounting the route.
     *
     * @param array<string,string|list<string>> $query already URL-decoded by the
     * framework
     * @return array{status: int, contentType: string, body: string}|null
     */
    public static function handshakeResponse(array $query): ?array
    {
        return WebhookVerifier::handshakeResponse(self::HANDSHAKE_PARAM, $query);
    }
}
