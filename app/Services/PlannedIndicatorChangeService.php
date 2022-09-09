<?php
declare(strict_types=1);

namespace App\Services;

use App\Exceptions\InitialDataNotFixedException;
use App\Models\PlannedIndicator;
use App\Models\PlannedIndicatorChange;
use App\Services\Dto\PlannedIndicatorChangeValueCollectionDto;
use App\Services\Dto\PlannedIndicatorChangeValueCollectionQueryResultDto;
use App\Services\Dto\PlannedIndicatorChangeValueDto;
use App\Services\Dto\PlannedIndicatorChangeValueQueryResultDto;
use Illuminate\Support\Facades\DB;

class PlannedIndicatorChangeService
{
    public function __construct(
        private PeriodService $periodService,
        private NodeService $nodeService,
        private VolumeDistributionAlgorithmService $algorithm,
        private InitialDataFixingService $initialDataFixingService,
    ) { }

    public function getByNodeIdAndYear(int $nodeId, int $year): array {
        return $this->getByNodeIdsAndYear([$nodeId],$year);
    }

    public function getByNodeIdsAndYear(array $nodeIds, int $year): array {
        $ids = $this->nodeService->plannedIndicatorsForNodeIds($nodeIds);
        $periodIds = $this->periodService->getIdsByYear($year);

        $data = PlannedIndicatorChange::whereIn('period_id', $periodIds)->whereIn('planned_indicator_id',$ids)->orderBy('id')->get();

        $data = $data->map(function ($e) {
            return (
            new PlannedIndicatorChangeValueDto(
                id: $e->id,
                periodId: $e->period_id,
                moId: $e->mo_id,
                moDepartmentId: $e->mo_department_id,
                plannedIndicatorId: $e->planned_indicator_id,
                value: $e->value,
                userId: $e->user_id,
                commitId: $e->commit_id
            ));
        });
        return $data->all();//->toArray();//
    }

    public function setValue(PlannedIndicatorChangeValueDto $dto)
    {
        bcscale(4);

        $plannedIndicator = PlannedIndicator::findOrFail($dto->plannedIndicatorId);
        // $algorithmId = $this->algorithm->getAlgorithmId($plannedIndicator->indicator_id);

        $plannedIndicatorChange = new PlannedIndicatorChange();
        $plannedIndicatorChange->period_id = $dto->periodId;
        $plannedIndicatorChange->mo_id = $dto->moId;
        $plannedIndicatorChange->planned_indicator_id = $plannedIndicator->id;
        // $initialData->algorithm_id = $algorithmId;
        $plannedIndicatorChange->user_id = $dto->userId;
        $plannedIndicatorChange->mo_department_id = $dto->moDepartmentId;

        // Проверяем, что данные доступны для редактирования.
        $year = $this->periodService->getYearById($dto->periodId);
        if (!$this->initialDataFixingService->fixed($plannedIndicator->node_id, $year)) {
            throw new InitialDataNotFixedException('Начальные данные для текущего раздела не зафиксированны');
        }

        $newValue = $dto->value;
        DB::transaction(function() use ($newValue, $plannedIndicatorChange) {
            $value = null;
            $sql = PlannedIndicatorChange::select('value')
                ->where('period_id', $plannedIndicatorChange->period_id)
                ->where('mo_id', $plannedIndicatorChange->mo_id)
                ->where('planned_indicator_id', $plannedIndicatorChange->planned_indicator_id)
                ->where('mo_department_id', $plannedIndicatorChange->mo_department_id);
            // блокируем строки, содержащие текущее значение планового показателя
            (clone $sql)->lockForUpdate()->get();
            $currentValue = $sql->groupBy('period_id','mo_id','planned_indicator_id','mo_department_id')->sum('value');
            $value = bcsub($newValue, (string)$currentValue);
            if($value === null) {
                /// Exception
            }
            $plannedIndicatorChange->value = $value;
            $plannedIndicatorChange->save();
        });

        $change = new PlannedIndicatorChangeValueDto(
            id: $plannedIndicatorChange->id,
            periodId: $plannedIndicatorChange->period_id,
            moId: $plannedIndicatorChange->mo_id,
            plannedIndicatorId: $plannedIndicatorChange->planned_indicator_id,
            moDepartmentId: $plannedIndicatorChange->mo_department_id,
            value: $plannedIndicatorChange->value,
            userId: $plannedIndicatorChange->user_id
        );

        $result = new PlannedIndicatorChangeValueQueryResultDto(
            operationError: false,
            operationMessage: '',
            hasValue: true,
            value: $change
        );

        return $result;
    }


    public function incrementValues(PlannedIndicatorChangeValueCollectionDto $collection)
    {
        $arr = $collection->toArray();

        // Проверяем, что данные доступны для редактирования.
        foreach($arr as $dto) {
            $plannedIndicator = PlannedIndicator::findOrFail($dto->plannedIndicatorId);
            $year = $this->periodService->getYearById($dto->periodId);
            if (!$this->initialDataFixingService->fixed($plannedIndicator->node_id, $year)) {
                throw new InitialDataNotFixedException('Начальные данные для раздела не зафиксированны');
            }
        }

        $entities = [];
        foreach($arr as $dto) {
            $plannedIndicatorChange = new PlannedIndicatorChange();
            $plannedIndicatorChange->period_id = $dto->periodId;
            $plannedIndicatorChange->mo_id = $dto->moId;
            $plannedIndicatorChange->planned_indicator_id = $plannedIndicator->id;
            // $initialData->algorithm_id = $algorithmId;
            $plannedIndicatorChange->user_id = $dto->userId;
            $plannedIndicatorChange->mo_department_id = $dto->moDepartmentId;
            $plannedIndicatorChange->value = $dto->value;
            $entities[] = $plannedIndicatorChange;
        }
        DB::transaction(function() use ($entities) {
            foreach($entities as $plannedIndicatorChange) {
                $plannedIndicatorChange->save();
            }
        });

        $arr = [];
        foreach($entities as $plannedIndicatorChange) {
            $arr[] = new PlannedIndicatorChangeValueDto(
                id: $plannedIndicatorChange->id,
                periodId: $plannedIndicatorChange->period_id,
                moId: $plannedIndicatorChange->mo_id,
                plannedIndicatorId: $plannedIndicatorChange->planned_indicator_id,
                moDepartmentId: $plannedIndicatorChange->mo_department_id,
                value: $plannedIndicatorChange->value,
                userId: $plannedIndicatorChange->user_id
            );
        }

        $result = new PlannedIndicatorChangeValueCollectionQueryResultDto(
            operationError: false,
            operationMessage: '',
            hasValue: true,
            value: new PlannedIndicatorChangeValueCollectionDto(...$arr)
        );
        return $result;
    }
}
