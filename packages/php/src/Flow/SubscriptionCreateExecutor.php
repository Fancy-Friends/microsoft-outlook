<?php

declare(strict_types=1);

namespace ParticleAcademy\MicrosoftOutlook\Flow;

use FancyFlow\Attributes\FlowNode;
use FancyFlow\Contracts\NodeExecutor;
use FancyFlow\Runtime\ExecutionContext;
use FancyFlow\Runtime\Port;
use FancyFlow\Runtime\RunEvent;
use ParticleAcademy\Connectors\ConnectorClient;
use ParticleAcademy\MicrosoftOutlook\Actions\SubscriptionCreate;
use ParticleAcademy\MicrosoftOutlook\MicrosoftOutlook;

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
 * Outlook subscription, run on a fancy-flow-php host.
 *
 * The PHP twin of `microsoftOutlookSubscriptionCreateExecutor` in
 * @particle-academy/microsoft-outlook-js: the same request, built from the
 * node's config by the same `Actions\SubscriptionCreate` a host would call
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
    name: '@particle-academy/microsoft_outlook_subscription_create',
    aliases: [
        'microsoft_outlook_subscription_create',
    ],
    category: 'io',
    label: 'Outlook subscription',
    description: 'Subscribe to change notifications on the connected account\'s calendar events. The host\'s subscription machinery calls this; it is not a node most workflows need.',
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
    sideEffects: 'unsafe-to-replay',
    outputShape: [
        [
            'path' => 'data.id',
            'type' => 'string',
            'description' => 'The subscription id. subscription_renew and subscription_delete need it.',
        ],
        [
            'path' => 'data.resource',
            'type' => 'string',
            'description' => 'What is watched.',
        ],
        [
            'path' => 'data.changeType',
            'type' => 'string',
            'description' => 'Which changes raise a notification.',
        ],
        [
            'path' => 'data.notificationUrl',
            'type' => 'string',
            'description' => 'Where notifications go.',
        ],
        [
            'path' => 'data.expirationDateTime',
            'type' => 'string',
            'description' => 'When it expires, ISO 8601 UTC with a seven-digit fraction. The subscription trigger\'s lease is built from this.',
        ],
        [
            'path' => 'mode',
            'type' => 'string',
            'description' => 'Which estate this ran against: fake, sandbox or live.',
        ],
    ],
)]
final class SubscriptionCreateExecutor implements NodeExecutor
{
    public function __construct(private readonly ?ConnectorClient $client = null) {}

    public function execute(ExecutionContext $ctx): mixed
    {
        $config = $ctx->config();

        $result = ($this->client ?? new ConnectorClient)->call(
            MicrosoftOutlook::descriptor(),
            SubscriptionCreate::OPERATION,
            $config,
            [
                'method' => SubscriptionCreate::METHOD,
                'path' => SubscriptionCreate::PATH,
                'json' => SubscriptionCreate::body($config),
            ],
            $ctx->input('in'),
        );

        $id = is_array($result->data) ? ($result->data['id'] ?? null) : null;
        $ctx->emit(RunEvent::log(
            'info',
            'microsoft_outlook subscription_create'.(is_scalar($id) ? ' '.$id : '').' ('.$result->mode->value.')',
            $ctx->node->id,
        ));

        return Port::only('out', $result->toArray());
    }
}
