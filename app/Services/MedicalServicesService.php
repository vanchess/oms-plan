<?php
declare(strict_types=1);

namespace App\Services;

use App\Models\PlannedIndicator;

class MedicalServicesService
{
    public function getIdsByNodeId(int $nodeId)
    {
        $column = 'service_id';
        $usedIds = PlannedIndicator::select($column)->where('node_id', $nodeId)->whereNotNull($column)->groupBy($column)->pluck($column)->toArray();
        if(count($usedIds) === 0) {
            return [];
        }
        sort($usedIds);
        return $usedIds;
    }

    // Возвращает массив Id медицинских (диагностических) услуг для указанного узла(категории)
    // актуальных на начало указанного года (на текущую дату, если не указан год)
    public function getIdsByNodeIdAndYear(int $nodeId, int $year = null)
    {
        $column = 'service_id';

        $date = date('Y-m-d');
        if ($year !== null) {
            $date = ($year.'-01-02');
        }

        $usedIds = PlannedIndicator::select($column)
            ->where('node_id', $nodeId)
            ->whereNotNull($column)
            ->whereRaw("? BETWEEN effective_from AND effective_to", [$date])
            ->groupBy($column)->pluck($column)->toArray();

        if(count($usedIds) === 0) {
            return [];
        }
        sort($usedIds);
        return $usedIds;
    }

    // Возвращает массив Id медицинских (диагностических) услуг актуальных на начало указанного года (на текущую дату, если не указан год)
    public function getIdsByYear(int $year = null)
    {
        $column = 'service_id';

        $date = date('Y-m-d');
        if ($year !== null) {
            $date = ($year.'-01-02');
        }

        $usedIds = PlannedIndicator::select($column)->whereNotNull($column)->whereRaw("? BETWEEN effective_from AND effective_to", [$date])->groupBy($column)->pluck($column)->toArray();
        if(count($usedIds) === 0) {
            return [];
        }
        sort($usedIds);
        return $usedIds;
    }



}
