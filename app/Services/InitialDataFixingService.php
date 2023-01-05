<?php
declare(strict_types=1);

namespace App\Services;

use App\Exceptions\InitialDataIsFixedException;
use App\Models\ChangePackage;
use App\Models\InitialDataLoaded;

class InitialDataFixingService
{
    public function __construct(
        private PlannedIndicatorChangeInitService $plannedIndicatorChangeInitService,
        private PlannedIndicatorChangeCommitService $plannedIndicatorChangeCommitService,
        private NodeService $nodeService
    ) { }

    public function getLoadedNodeIdsByYear(int $year): array
    {
        return InitialDataLoaded::where('year',$year)->orderBy('node_id')->get()->pluck('node_id')->toArray();
    }

    public function fixed(int $nodeId, int $year): bool
    {
        $f = InitialDataLoaded::where('year',$year)->where('node_id',$nodeId)->first();
        return isset($f);
    }

    /**
     * Есть зафиксированные начальные данные для указанного года
     */
    public function fixedYear(int $year): bool
    {
        $f = InitialDataLoaded::where('year',$year)->first();
        return isset($f);
    }

    public function commit(int $nodeId, int $year, int $userId)
    {
        $nodeIds = $this->nodeService->nodeWithChildrenIds($nodeId);
        foreach ($nodeIds as $key => $id) {
            // Пропускаем зафиксированные ранее подкатегории
            if ($this->fixed($id, $year)) {
                unset($nodeIds[$key]);
            }
        }

        if($this->fixed($nodeId, $year)) {
            throw new InitialDataIsFixedException('Начальные данные для раздела уже зафиксированны');
        }

        $changePackage = new ChangePackage();
        $changePackage->user_id = $userId;
        $changePackage->save();
        $changePackageId = $changePackage->id;

        // Помечаем начальные данные для указанного узла(и дочерних) как загруженные
        $this->markLoaded($nodeIds, $year, $changePackageId, $userId);
        // Формируем "базовые" изменения на основе начальных данных
        $this->plannedIndicatorChangeInitService->fromInitialData($year);
        // Указываем идентификатор фиксации данных для "базовых" изменений
        $this->plannedIndicatorChangeCommitService->commitByNodeIds($nodeIds, $year, $changePackageId);
    }

    private function markLoaded(array $nodeIds, int $year, int $packageId, int $userId)
    {
        $arr = [];
        foreach($nodeIds as $nodeId) {
            $l = new InitialDataLoaded();
            $l->year = $year;
            $l->node_id = $nodeId;
            $l->user_id= $userId;
            $l->package_id = $packageId;
            $arr[] = $l->attributesToArray();
        }
        InitialDataLoaded::Insert($arr);
    }
}
