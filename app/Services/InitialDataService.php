<?php
declare(strict_types=1);

namespace App\Services;

use App\Models\InitialData;
use App\Models\PlannedIndicator;

use App\Exceptions\InitialDataIsFixedException;
use App\Services\Dto\InitialDataValueDto;

use App\Services\Dto\InitialDataValueQueryResultDto;

class InitialDataService
{

    function __construct(
        private NodeService $nodeService,
        private InitialDataFixingService $initialDataFixingService,
        private VolumeDistributionAlgorithmService $algorithm
    ) { }

    public function getByNodeIdAndYear(int $nodeId, int $year) {
        $ids = $this->nodeService->plannedIndicatorsForNodeId($nodeId);
        //::with('plannedIndicator')
        $data = InitialData::where('year', $year)->whereIn('planned_indicator_id',$ids)->orderBy('id')->get();



        $data = $data->map(function ($e) {
            return (
            new InitialDataValueDto(
                id: $e->id,
                year: $e->year,
                moId: $e->mo_id,
                moDepartmentId: $e->mo_department_id,
                plannedIndicatorId: $e->planned_indicator_id,
                value: $e->value,
                userId: $e->user_id
            ));
        });
        return $data->all();//->toArray();//
    }

    /**
     * Добавляет начальные данные в таблицу
     *
     * @return \Illuminate\Http\Response
     */
    public function setValue(InitialDataValueDto $dto)
    {
        $plannedIndicator = PlannedIndicator::findOrFail($dto->plannedIndicatorId);

        $algorithmId = $this->algorithm->getAlgorithmId($plannedIndicator->indicator_id);

        //DB::transaction(function() use ($dto) {
        // Проверяем, что данные доступны для редактирования.
        if ($this->initialDataFixingService->fixed($plannedIndicator->node_id, $dto->year)) {
            throw new InitialDataIsFixedException('Начальные данные для данного раздела зафиксированны');
        }

        $initialData = new InitialData();
        $initialData->year = $dto->year;
        $initialData->mo_id = $dto->moId;
        $initialData->planned_indicator_id = $plannedIndicator->id;
        $initialData->value = $dto->value;
        $initialData->algorithm_id = $algorithmId;
        $initialData->user_id = $dto->userId;
        $initialData->mo_department_id = $dto->moDepartmentId;
        $initialData->save();

        $value = new InitialDataValueDto(
            id: $initialData->id,
            year: $initialData->year,
            moId: $initialData->mo_id,
            plannedIndicatorId: $initialData->planned_indicator_id,
            moDepartmentId: $initialData->mo_department_id,
            value: $initialData->value,
            userId: $initialData->user_id
        );

        $result = new InitialDataValueQueryResultDto(
            operationError: false,
            operationMessage: '',
            hasValue: true,
            value: $value
        );

        return $result;
    }
}
