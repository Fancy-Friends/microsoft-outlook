<?php

declare(strict_types=1);

namespace ParticleAcademy\MicrosoftOutlook\Flow;

use FancyFlow\Attributes\FlowNode;
use FancyFlow\Contracts\NodeExecutor;
use FancyFlow\Runtime\ExecutionContext;
use FancyFlow\Runtime\Port;
use ParticleAcademy\Connectors\ConnectionHost;
use ParticleAcademy\Connectors\TriggerEvent;
use ParticleAcademy\MicrosoftOutlook\MicrosoftOutlook;
use ParticleAcademy\MicrosoftOutlook\Triggers\CalendarChanged;

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
 * Outlook calendar change, run on a fancy-flow-php host.
 *
 * The PHP twin of `microsoftOutlookCalendarChangedTriggerExecutor` in
 * @particle-academy/microsoft-outlook-js. A webhook trigger never calls
 * Microsoft Outlook: it republishes, on `out`, the delivery the HOST received
 * and verified at its own route. With nothing delivered, fake mode publishes
 * the faker's sample event, so a flow can be designed before the endpoint
 * exists; any other mode refuses, and says how to deliver one.
 */
#[FlowNode(
    name: '@particle-academy/microsoft_outlook_calendar_changed_trigger',
    aliases: [
        'microsoft_outlook_calendar_changed_trigger',
    ],
    category: 'trigger',
    label: 'Outlook calendar change',
    description: 'Start a run when an event is created, updated or deleted on a subscribed Outlook calendar.',
    icon: '📅',
    inputs: [],
    outputs: [
        [
            'id' => 'out',
        ],
    ],
    sideEffects: 'none',
    outputShape: [
        [
            'path' => 'value',
            'type' => 'array',
            'description' => 'The notifications in this delivery -- a batch, possibly spanning several subscriptions. Each item is described below.',
        ],
        [
            'path' => 'value[].subscriptionId',
            'type' => 'string',
            'description' => 'Which subscription this item belongs to.',
        ],
        [
            'path' => 'value[].subscriptionExpirationDateTime',
            'type' => 'string',
            'description' => 'When that subscription expires, ISO 8601 UTC -- Graph\'s hint for when to renew.',
        ],
        [
            'path' => 'value[].changeType',
            'type' => 'string',
            'description' => 'created, updated or deleted.',
        ],
        [
            'path' => 'value[].resource',
            'type' => 'string',
            'description' => 'The changed resource\'s path, e.g. Users/{id}/Events/{id}.',
        ],
        [
            'path' => 'value[].resourceData.id',
            'type' => 'string',
            'description' => 'The changed event\'s id. event_get reads the rest.',
        ],
        [
            'path' => 'value[].resourceData.@odata.type',
            'type' => 'string',
            'description' => '#Microsoft.Graph.Event.',
        ],
        [
            'path' => 'value[].tenantId',
            'type' => 'string',
            'description' => 'The tenant the change happened in.',
        ],
        [
            'path' => 'value[].lifecycleEvent',
            'type' => 'string',
            'description' => 'Only on a LIFECYCLE notification: subscriptionRemoved, missed or reauthorizationRequired. Absent on a change notification.',
        ],
    ],
)]
final class CalendarChangedTriggerExecutor implements NodeExecutor
{
    public function __construct(private readonly ?ConnectionHost $host = null) {}

    public function execute(ExecutionContext $ctx): mixed
    {
        $config = $ctx->config();
        $service = MicrosoftOutlook::descriptor();

        $connection = ($this->host ?? new ConnectionHost)->resolve(
            $service->service,
            CalendarChanged::OPERATION,
            $config,
            $service->sandbox,
            $service->requires,
            $service->baseUrls,
        );

        $event = TriggerEvent::resolve(
            $service->service,
            CalendarChanged::OPERATION,
            CalendarChanged::DELIVERY,
            CalendarChanged::SETUP,
            $service->faker,
            $connection,
            $ctx->input('in'),
            $config,
        );

        return Port::only('out', $event);
    }
}
