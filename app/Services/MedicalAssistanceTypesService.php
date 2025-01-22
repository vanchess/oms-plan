<?php
declare(strict_types=1);

namespace App\Services;

use App\Models\PlannedIndicator;

class MedicalAssistanceTypesService {

    public function getIdsByNodeId(int $nodeId)
    {
        $column = 'assistance_type_id';
        $usedIds = PlannedIndicator::select($column)->where('node_id', $nodeId)->whereNotNull($column)->groupBy($column)->pluck($column)->toArray();
        if(count($usedIds) === 0) {
            return [];
        }
        sort($usedIds);
        return $usedIds;
    }

    public function getIdsByNodeIdAndYear(int $nodeId, int $year = null)
    {
        $column = 'assistance_type_id';

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
}
