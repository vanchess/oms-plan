<?php
declare(strict_types=1);

namespace App\Services;

use App\Models\PlannedIndicator;
use App\Models\PlannedIndicatorChange;
use Illuminate\Support\Facades\DB;

class PeopleAssignedInfoForContractService
{
    public function __construct(
        private PeriodService $periodService
    ) {}

    public function GetJson(int $year, array $packageIds = null): string {
        return $this->CreateData($year, $packageIds)->toJson();
    }

    public function GetArray(int $year, array $packageIds = null): array {
        return $this->CreateData($year, $packageIds)->toArray();
    }

    private function CreateData(int $year, array $packageIds = null) {
        $indicatorIds = [10];

        $hospitalNodeIds = [35];
        $ambulanceNodeIds = [39];

        $polyclinicNodeIds = [29,28];
        $polyclinicFapNodeIds = [29];

        $periodIds = $this->periodService->getIdsByYear($year);
        $periodIds = [$periodIds[array_key_last($periodIds)]];

        $dataSql = DB::table((new PlannedIndicator())->getTable().' as pi')
        ->selectRaw('SUM(value) as value, node_id, indicator_id, mo_id, planned_indicator_id, mo_department_id')
        ->leftJoin((new PlannedIndicatorChange())->getTable().' as pic', 'pi.id', '=', 'pic.planned_indicator_id');
        if ($packageIds) {
            $dataSql = $dataSql->whereIn('package_id',$packageIds);
        }
        //
        $dataSql = $dataSql->whereIn('indicator_id', $indicatorIds)
        ->whereIn('period_id', $periodIds)
        ->whereNull('pic.deleted_at')
        ->groupBy('node_id', 'indicator_id', 'mo_id', 'planned_indicator_id', 'mo_department_id');

        $data = $dataSql->get();

        $data = $data->groupBy([
            'mo_id',
            function($item, $key) use  ($hospitalNodeIds, $ambulanceNodeIds, $polyclinicNodeIds) {
                $nodeId = $item->node_id;
                if(in_array($nodeId, $hospitalNodeIds)) {
                    return 'hospital';
                };
                if(in_array($nodeId, $ambulanceNodeIds)) {
                    return 'ambulance';
                };
                if(in_array($nodeId, $polyclinicNodeIds)) {
                    return 'polyclinic';
                };
                return 'none';
            }
        ]);

        foreach ($data as $key => $value) {
            if(isset($value['hospital'])) {
                $value['hospital'] = $value['hospital']->mapWithKeys(
                    function($item, $key) {
                        return ['peopleAssigned' => $item->value];
                    }
                );
            }
            if(isset($value['ambulance'])) {
                $value['ambulance'] = $value['ambulance']->mapWithKeys(
                    function($item, $key) {
                        return ['peopleAssigned' => $item->value];
                    }
                );
            }

            if(isset($value['polyclinic'])) {
                $value['polyclinic'] = $value['polyclinic']->groupBy([
                    function($item, $key) use  ($polyclinicFapNodeIds) {
                        $nodeId = $item->node_id;
                        if(in_array($nodeId, $polyclinicFapNodeIds)) {
                            return 'fap';
                        };
                        return 'mo';
                    },
                    function($item, $key) use  ($polyclinicFapNodeIds) {
                        $nodeId = $item->node_id;
                        $departmentId = $item->mo_department_id;
                        if(in_array($nodeId, $polyclinicFapNodeIds)) {
                            return $departmentId;
                        };
                        return 'all';
                    },
                ]);

                if(isset($value['polyclinic']['fap'])) {
                    $value['polyclinic']['fap']->transform(function($fapData) {
                        return $fapData->mapWithKeys(function($item) {
                            return ['peopleAssigned' => $item->value];
                        });
                    });
                }

                if(isset($value['polyclinic']['mo'])) {
                    $value['polyclinic']['mo'] = $value['polyclinic']['mo']['all']->mapWithKeys(function($moData) {
                        return ['peopleAssigned' => $moData->value];
                    });
                }
            }
        }
        return $data;
    }
}
