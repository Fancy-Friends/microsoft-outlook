<?php

declare(strict_types=1);

namespace ParticleAcademy\MicrosoftOutlook\Flow;

use FancyFlow\Attributes\FlowNode;
use FancyFlow\Contracts\NodeExecutor;
use FancyFlow\Runtime\ExecutionContext;
use FancyFlow\Runtime\Port;
use FancyFlow\Runtime\RunEvent;
use ParticleAcademy\Connectors\ConnectorClient;
use ParticleAcademy\MicrosoftOutlook\Actions\EventList;
use ParticleAcademy\MicrosoftOutlook\MicrosoftOutlook;

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
 * Outlook calendar events, run on a fancy-flow-php host.
 *
 * The PHP twin of `microsoftOutlookEventListExecutor` in
 * @particle-academy/microsoft-outlook-js: the same request, built from the
 * node's config by the same `Actions\EventList` a host would call directly,
 * and the same value on `out` — the client's `{data, mode, connection}`.
 *
 * The client resolves the connection and the estate from the config. With
 * nothing configured that is FAKE, so a node dropped on a canvas runs against
 * the faker rather than Microsoft Outlook. To reach a real estate, pass a
 * `ConnectorClient` that knows the host's connections — or bind one in the
 * container, which resolves the constructor by type.
 */
#[FlowNode(
    name: '@particle-academy/microsoft_outlook_event_list',
    aliases: [
        'microsoft_outlook_event_list',
    ],
    category: 'io',
    label: 'Outlook calendar events',
    description: 'List events from the connected account\'s Outlook calendar.',
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
    sideEffects: 'none',
    outputShape: [
        [
            'path' => 'data.value',
            'type' => 'array',
            'description' => 'The events on this page, each shaped like event_get\'s output.',
        ],
        [
            'path' => 'data.@odata.nextLink',
            'type' => 'string',
            'description' => 'Present when there is another page: the full URL of the next one.',
        ],
        [
            'path' => 'mode',
            'type' => 'string',
            'description' => 'Which estate this ran against: fake, sandbox or live.',
        ],
    ],
)]
final class EventListExecutor implements NodeExecutor
{
    public function __construct(private readonly ?ConnectorClient $client = null) {}

    public function execute(ExecutionContext $ctx): mixed
    {
        $config = $ctx->config();

        $result = ($this->client ?? new ConnectorClient)->call(
            MicrosoftOutlook::descriptor(),
            EventList::OPERATION,
            $config,
            [
                'method' => EventList::METHOD,
                'path' => EventList::PATH,
                'query' => EventList::body($config),
            ],
            $ctx->input('in'),
        );

        $id = is_array($result->data) ? ($result->data['id'] ?? null) : null;
        $ctx->emit(RunEvent::log(
            'info',
            'microsoft_outlook event_list'.(is_scalar($id) ? ' '.$id : '').' ('.$result->mode->value.')',
            $ctx->node->id,
        ));

        return Port::only('out', $result->toArray());
    }
}
