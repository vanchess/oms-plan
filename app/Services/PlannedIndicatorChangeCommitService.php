<?php
declare(strict_types=1);

namespace App\Services;

use App\Models\PlannedIndicatorChange;

class PlannedIndicatorChangeCommitService
{
    public function __construct(
        private PeriodService $periodService,
        private NodeService $nodeService
    ) { }

    /**
     * Зафиксировать текущее изменения для указанного узла(категории) и его потомков
     */
    public function commitByNodeId(int $nodeId, int $year, int $packageId) {
        $nodeIds = $this->nodeService->nodeWithChildrenIds($nodeId);
        $this->commitByNodeIds($nodeIds, $year, $packageId);
    }

    /**
     * Зафиксировать текущее изменения для указанных узлов(категорий)
     */
    public function commitByNodeIds(array $nodeIds, int $year, int $packageId) {
        $periodIds = $this->periodService->getIdsByYear($year);
        $plannedIndicatorIds = $this->nodeService->plannedIndicatorsForNodeIds($nodeIds);
        PlannedIndicatorChange::whereNull('package_id')
            ->whereIn('period_id',$periodIds)
            ->whereIn('planned_indicator_id', $plannedIndicatorIds)
            ->update(['package_id' => $packageId]);
    }
}
