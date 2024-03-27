<?php
declare(strict_types=1);

namespace App\Services;

use App\Models\CategoryTreeNodes;
use App\Models\PlannedIndicator;
use App\Models\Indicator;
use Illuminate\Support\Collection;
use Illuminate\Support\Facades\DB;

class NodeService
{
    /**
     * Возвращает массив id показателей используемых для переданного id узла дерева категорий
     *
     * @return \Illuminate\Http\Response
     */
    public function indicatorsUsedForNodeId(int $nodeId)
    {
        /* TODO:
            Возможно нужно вынести в отдельную таблицу:
            id | node_id | indicator_id | order      unique(node_id, indicator_id)
        */
        $usedId = PlannedIndicator::select('indicator_id')->where('node_id', $nodeId)->groupBy('indicator_id')->pluck('indicator_id');
        return Indicator::whereIn('id',$usedId)->orderBy('order')->pluck('id');
    }

    /**
     * Возвращает массив id плановых показателей для переданного id узла дерева категорий
     * (плановый показатель = показатель,услуга,профиль,вид медпомощи, вид и группа ВМП... )
     *
     * @return \Illuminate\Http\Response
     */
    public function plannedIndicatorsForNodeId(int $nodeId): array
    {
        return PlannedIndicator::select('id')->where('node_id', $nodeId)->pluck('id')->toArray();
    }

    /**
     * Возвращает массив id плановых показателей для переданного массива id узлов дерева категорий
     * (плановый показатель = показатель,услуга,профиль,вид медпомощи, вид и группа ВМП... )
     *
     * @return \Illuminate\Http\Response
     */
    public function plannedIndicatorsForNodeIds(array $nodeIds): array
    {
        return PlannedIndicator::select('id')->whereIn('node_id', $nodeIds)->pluck('id')->toArray();
    }


    /**
     * Возвращает массив id профилей коек для переданного id узла дерева категорий
     *
     * @return \Illuminate\Http\Response
     */
    public function hospitalBedProfilesForNodeId(int $nodeId)
    {
        $usedIds = PlannedIndicator::select('profile_id')->where('node_id', $nodeId)->whereNotNull('profile_id')->groupBy('profile_id')->pluck('profile_id')->toArray();
        if(count($usedIds) === 0) {
            return [];
        }
        sort($usedIds);
        return $usedIds;
    }

    public function careProfilesForNodeId(int $nodeId)
    {
        $usedIds = PlannedIndicator::select('care_profile_id')->where('node_id', $nodeId)->whereNotNull('care_profile_id')->groupBy('care_profile_id')->pluck('care_profile_id')->toArray();

        if(count($usedIds) === 0) {
            return [];
        }
        sort($usedIds);
        return $usedIds;
    }

    public function medicalAssistanceTypesForNodeId(int $nodeId)
    {
        $column = 'assistance_type_id';
        $usedIds = PlannedIndicator::select($column)->where('node_id', $nodeId)->whereNotNull($column)->groupBy($column)->pluck($column)->toArray();
        if(count($usedIds) === 0) {
            return [];
        }
        sort($usedIds);
        return $usedIds;
    }

    public function medicalServicesForNodeId(int $nodeId)
    {
        $column = 'service_id';
        $usedIds = PlannedIndicator::select($column)->where('node_id', $nodeId)->whereNotNull($column)->groupBy($column)->pluck($column)->toArray();
        if(count($usedIds) === 0) {
            return [];
        }
        sort($usedIds);
        return $usedIds;
    }

    private function nodeWithChildren(int $nodeId)
    {
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
        $query = CategoryTreeNodes::select('id', 'parent_id')->join('cte', 'parent_id', '=', 'cte.node_id');
        $nodes = CategoryTreeNodes::select('id', 'parent_id')->where('id', '=', $nodeId)->unionAll($query);
        $res = DB::select("WITH recursive cte (node_id, node_parent_id) AS ({$nodes->toSql()}) select node_id as id, node_parent_id as parent_id from cte order by node_id;",[$nodeId]);
        return $res;

    }

    private function allChildrenNodes(int $nodeId)
    {
        /*
        $res = DB::select("with RECURSIVE cte as
            (
                select tct.id, tct.parent_id from tbl_category_tree tct where tct.parent_id = ?
                union all
                select tct.id, tct.parent_id from tbl_category_tree tct
                inner join cte as c on c.id = tct.parent_id
            )
            select id from cte;",
            [$nodeId]
            );
        */
        $query = CategoryTreeNodes::select('id', 'parent_id')->join('cte', 'parent_id', '=', 'cte.node_id');
        $nodes = CategoryTreeNodes::select('id', 'parent_id')->where('parent_id', '=', $nodeId)->unionAll($query);
        $res = DB::select("WITH recursive cte (node_id, node_parent_id) AS ({$nodes->toSql()}) select node_id as id, node_parent_id as parent_id from cte order by node_id;",[$nodeId]);
        return $res;

    }

    public function allChildrenNodeIds(int $nodeId): array
    {
        $res = $this->allChildrenNodes($nodeId);

        $ids = array_column($res, 'id');
        return $ids;
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
