<?php

declare(strict_types=1);

namespace ParticleAcademy\MicrosoftOutlook\Flow;

use FancyFlow\Attributes\FlowNode;
use FancyFlow\Contracts\NodeExecutor;
use FancyFlow\Runtime\ExecutionContext;
use FancyFlow\Runtime\Port;
use FancyFlow\Runtime\RunEvent;
use ParticleAcademy\Connectors\ConnectorClient;
use ParticleAcademy\MicrosoftOutlook\Actions\SubscriptionRenew;
use ParticleAcademy\MicrosoftOutlook\MicrosoftOutlook;

/*
 * GENERATED FILE — do not edit.
 *
 * Emitted from provider/actions/subscription-renew.json by weaver's generator.
 * A hand-edit here is destroyed by the next protocol sync, which is worse than
 * being rejected, because it works until it silently does not. Fix
 * provider/actions/subscription-renew.json (or weaver's template/) and regenerate:
 *
 *     npm run provider -- microsoft_outlook
 */
/**
 * Outlook subscription renewal, run on a fancy-flow-php host.
 *
 * The PHP twin of `microsoftOutlookSubscriptionRenewExecutor` in
 * @particle-academy/microsoft-outlook-js: the same request, built from the
 * node's config by the same `Actions\SubscriptionRenew` a host would call
 * directly, and the same value on `out` — the client's `{data, mode,
 * connection}`.
 *
 * The client resolves the connection and the estate from the config. With
 * nothing configured that is FAKE, so a node dropped on a canvas runs against
 * the faker rather than Microsoft Outlook. To reach a real estate, pass a
 * `ConnectorClient` that knows the host's connections — or bind one in the
 * container, which resolves the constructor by type.
 */
#[FlowNode(
    name: '@particle-academy/microsoft_outlook_subscription_renew',
    aliases: [
        'microsoft_outlook_subscription_renew',
    ],
    category: 'io',
    label: 'Outlook subscription renewal',
    description: 'Extend a change-notification subscription before it expires. The host\'s subscription machinery calls this when the lease is due.',
    inputs: [
        [
            'id' => 'in',
        ],
    ],
    outputs: [
        [
            'id' => 'out',
        ],
    ],
    sideEffects: 'idempotent',
    outputShape: [
        [
            'path' => 'data.id',
            'type' => 'string',
            'description' => 'The subscription id.',
        ],
        [
            'path' => 'data.expirationDateTime',
            'type' => 'string',
            'description' => 'The new expiry, ISO 8601 UTC. The trigger\'s lease is rebuilt from this.',
        ],
        [
            'path' => 'mode',
            'type' => 'string',
            'description' => 'Which estate this ran against: fake, sandbox or live.',
        ],
    ],
)]
final class SubscriptionRenewExecutor implements NodeExecutor
{
    public function __construct(private readonly ?ConnectorClient $client = null) {}

    public function execute(ExecutionContext $ctx): mixed
    {
        $config = $ctx->config();

        $result = ($this->client ?? new ConnectorClient)->call(
            MicrosoftOutlook::descriptor(),
            SubscriptionRenew::OPERATION,
            $config,
            [
                'method' => SubscriptionRenew::METHOD,
                'path' => SubscriptionRenew::path($config),
                'json' => SubscriptionRenew::body($config),
            ],
            $ctx->input('in'),
        );

        $id = is_array($result->data) ? ($result->data['id'] ?? null) : null;
        $ctx->emit(RunEvent::log(
            'info',
            'microsoft_outlook subscription_renew'.(is_scalar($id) ? ' '.$id : '').' ('.$result->mode->value.')',
            $ctx->node->id,
        ));

        return Port::only('out', $result->toArray());
    }
}
