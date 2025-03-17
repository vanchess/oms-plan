<?php
declare(strict_types=1);

namespace App\Services;

use App\Models\PumpMonitoringProfiles;
use App\Models\PumpMonitoringProfilesUnit;
use Illuminate\Support\Collection;
use Illuminate\Support\Facades\DB;

class PumpMonitoringProfilesTreeService
{
    public function plannedIndicatorIdsViaChild(int $nodeId, int $unitId) : Array
    {
        $monitoringProfileIds = $this->nodeWithChildrenIds($nodeId);

        $profiles = PumpMonitoringProfiles::whereIn('id', $monitoringProfileIds)->get();
        $arr = [];
        foreach ($profiles as $p) {
            $pUnit = $p->profilesUnits()->where('unit_id', $unitId)->first();
            $pInd = $pUnit?->plannedIndicators?->ToArray() ?? [];
            $pIndIds = array_column($pInd, 'id');
            $arr = array_merge($arr, $pIndIds);
        }
        $bpArr = array_unique($arr, SORT_NUMERIC);
        return $bpArr;
    }

    public function plannedIndicatorIdsViaChildByMonitoringProfilesUnitId(int $monitoringProfilesUnitId)
    {
        $mpu = PumpMonitoringProfilesUnit::find($monitoringProfilesUnitId);
        return $this->plannedIndicatorIdsViaChild($mpu->monitoring_profile_id, $mpu->unit_id);
    }

    private function nodeWithChildren(int $nodeId, \DateTime|null $dt = null)
    {
        $dtStr = date('Y-m-d');
        if ($dt !== null) {
            $dtStr = $dt->format('Y-m-d');
        }
        /*
        $res = DB::select("with RECURSIVE cte as
            (
                select tct.id, tct.parent_id from tbl_category_tree tct where tct.id = ?
                union all
                select tct.id, tct.parent_id from tbl_category_tree tct
                inner join cte as c on c.id = tct.parent_id
            )
            select id from cte;",
            [$nodeId]
            );
        */
        $query = PumpMonitoringProfiles::select('id', 'parent_id')
            ->WhereRaw("? BETWEEN effective_from AND effective_to", [$dtStr])
            ->join('cte', 'parent_id', '=', 'cte.node_id');
        $nodes = PumpMonitoringProfiles::select('id', 'parent_id')->where('id', '=', $nodeId)
            ->WhereRaw("? BETWEEN effective_from AND effective_to", [$dtStr])
            ->unionAll($query);
        $res = DB::select("WITH recursive cte (node_id, node_parent_id) AS ({$nodes->toSql()}) select node_id as id, node_parent_id as parent_id from cte order by node_id;",
                        [$nodeId, $dtStr, $dtStr]);
        return $res;

    }

    public function nodeWithChildrenIds(int $nodeId): array
    {
        $res = $this->nodeWithChildren($nodeId);

        $ids = array_column($res, 'id');
        return $ids;
    }

    private function createTree(Collection $c, $curId)
    {
        $children = $c->where('parent_id', $curId)->toArray();
        if(count($children) === 0) {
            return null;
        }
        $branch = [];
        foreach ($children as $key => $el) {
            $branch[$el->id] = $this->createTree($c, $el->id);
        }
        return $branch;
    }

    public function nodeTree(int $rootNodeId)
    {
        $res = $this->nodeWithChildren($rootNodeId);
        $res = collect($res);

        $tree[$rootNodeId] = $this->createTree($res, $rootNodeId);

        return $tree;
    }
}
