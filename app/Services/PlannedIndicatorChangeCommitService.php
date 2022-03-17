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
    public function commitByNodeId(int $nodeId, int $year, int $commitId) {
        $nodeIds = $this->nodeService->nodeWithChildrenIds($nodeId);
        $this->commitByNodeIds($nodeIds, $year, $commitId);
    }

    /**
     * Зафиксировать текущее изменения для указанных узлов(категорий)
     */
    public function commitByNodeIds(array $nodeIds, int $year, int $commitId) {
        $periodIds = $this->periodService->getIdsByYear($year);
        $plannedIndicatorIds = $this->nodeService->plannedIndicatorsForNodeIds($nodeIds);
        PlannedIndicatorChange::whereNull('commit_id')
            ->whereIn('period_id',$periodIds)
            ->whereIn('planned_indicator_id', $plannedIndicatorIds)
            ->update(['commit_id' => $commitId]);
    }
}
