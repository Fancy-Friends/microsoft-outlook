<?php

declare(strict_types=1);

namespace ParticleAcademy\MicrosoftOutlook\Flow;

use FancyFlow\Attributes\FlowNode;
use FancyFlow\Contracts\NodeExecutor;
use FancyFlow\Runtime\ExecutionContext;
use FancyFlow\Runtime\Port;
use FancyFlow\Runtime\RunEvent;
use ParticleAcademy\Connectors\ConnectorClient;
use ParticleAcademy\MicrosoftOutlook\Actions\EventGet;
use ParticleAcademy\MicrosoftOutlook\MicrosoftOutlook;

/*
 * GENERATED FILE — do not edit.
 *
 * Emitted from provider/actions/event-get.json by weaver's generator.
 * A hand-edit here is destroyed by the next protocol sync, which is worse than
 * being rejected, because it works until it silently does not. Fix
 * provider/actions/event-get.json (or weaver's template/) and regenerate:
 *
 *     npm run provider -- microsoft_outlook
 */
/**
 * Outlook calendar event, run on a fancy-flow-php host.
 *
 * The PHP twin of `microsoftOutlookEventGetExecutor` in
 * @particle-academy/microsoft-outlook-js: the same request, built from the
 * node's config by the same `Actions\EventGet` a host would call directly, and
 * the same value on `out` — the client's `{data, mode, connection}`.
 *
 * The client resolves the connection and the estate from the config. With
 * nothing configured that is FAKE, so a node dropped on a canvas runs against
 * the faker rather than Microsoft Outlook. To reach a real estate, pass a
 * `ConnectorClient` that knows the host's connections — or bind one in the
 * container, which resolves the constructor by type.
 */
#[FlowNode(
    name: '@particle-academy/microsoft_outlook_event_get',
    aliases: [
        'microsoft_outlook_event_get',
    ],
    category: 'io',
    label: 'Outlook calendar event',
    description: 'Read one event from the connected account\'s Outlook calendar.',
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
            'path' => 'data.id',
            'type' => 'string',
            'description' => 'The event id.',
        ],
        [
            'path' => 'data.subject',
            'type' => 'string',
            'description' => 'The title.',
        ],
        [
            'path' => 'data.bodyPreview',
            'type' => 'string',
            'description' => 'The first lines of the body, as text.',
        ],
        [
            'path' => 'data.start.dateTime',
            'type' => 'string',
            'description' => 'Start, in start.timeZone. Graph writes seven fractional digits.',
        ],
        [
            'path' => 'data.end.dateTime',
            'type' => 'string',
            'description' => 'End, in end.timeZone.',
        ],
        [
            'path' => 'data.organizer.emailAddress.address',
            'type' => 'string',
            'description' => 'Who organises it.',
        ],
        [
            'path' => 'data.isCancelled',
            'type' => 'boolean',
            'description' => 'Whether the event was cancelled.',
        ],
        [
            'path' => 'data.webLink',
            'type' => 'string',
            'description' => 'The event in Outlook on the web.',
        ],
        [
            'path' => 'data.lastModifiedDateTime',
            'type' => 'string',
            'description' => 'ISO 8601 last modification, UTC.',
        ],
        [
            'path' => 'mode',
            'type' => 'string',
            'description' => 'Which estate this ran against: fake, sandbox or live.',
        ],
    ],
)]
final class EventGetExecutor implements NodeExecutor
{
    public function __construct(private readonly ?ConnectorClient $client = null) {}

    public function execute(ExecutionContext $ctx): mixed
    {
        $config = $ctx->config();

        $result = ($this->client ?? new ConnectorClient)->call(
            MicrosoftOutlook::descriptor(),
            EventGet::OPERATION,
            $config,
            [
                'method' => EventGet::METHOD,
                'path' => EventGet::path($config),
                'query' => EventGet::body($config),
            ],
            $ctx->input('in'),
        );

        $id = is_array($result->data) ? ($result->data['id'] ?? null) : null;
        $ctx->emit(RunEvent::log(
            'info',
            'microsoft_outlook event_get'.(is_scalar($id) ? ' '.$id : '').' ('.$result->mode->value.')',
            $ctx->node->id,
        ));

        return Port::only('out', $result->toArray());
    }
}
