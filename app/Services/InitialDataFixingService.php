<?php
declare(strict_types=1);

namespace App\Services;

use App\Exceptions\InitialDataIsFixedException;
use App\Models\DataCommit;
use App\Models\InitialDataLoaded;

class InitialDataFixingService
{
    public function __construct(
        private PlannedIndicatorChangeInitService $plannedIndicatorChangeInitService,
        private PlannedIndicatorChangeCommitService $plannedIndicatorChangeCommitService,
        private NodeService $nodeService
    ) { }

    public function fixed(int $nodeId, int $year): bool
    {
        $f = InitialDataLoaded::where('year',$year)->where('node_id',$nodeId)->first();
        return isset($f);
    }

    public function commit(int $nodeId, int $year, int $userId)
    {
        $nodeIds = $this->nodeService->nodeWithChildrenIds($nodeId);
        foreach ($nodeIds as $key => $id) {
            // Пропускаем зафиксированные ранее подкатегории
            if($this->fixed($id, $year)) {
                unset($nodeIds[$key]);
            }
        }

        if($this->fixed($nodeId, $year)) {
            throw new InitialDataIsFixedException('Начальные данные для данного раздела уже зафиксированны');
        }

        $dataCommit = new DataCommit();
        $dataCommit->user_id = $userId;
        $dataCommit->save();
        $dataCommitId = $dataCommit->id;

        // Помечаем начальные данные для указанного узла(и дочерних) как загруженные
        $this->markLoaded($nodeIds, $year, $dataCommitId, $userId);
        // Формируем "базовые" изменения на основе начальных данных
        $this->plannedIndicatorChangeInitService->fromInitialData($year);
        // Указываем идентификатор фиксации данных для "базовых" изменений
        $this->plannedIndicatorChangeCommitService->commitByNodeIds($nodeIds, $year, $dataCommitId);
    }

    private function markLoaded(array $nodeIds, int $year, int $commitId, int $userId)
    {
        $arr = [];
        foreach($nodeIds as $nodeId) {
            $l = new InitialDataLoaded();
            $l->year = $year;
            $l->node_id = $nodeId;
            $l->user_id= $userId;
            $l->commit_id = $commitId;
            $arr[] = $l->attributesToArray();
        }
        InitialDataLoaded::Insert($arr);
    }
}
