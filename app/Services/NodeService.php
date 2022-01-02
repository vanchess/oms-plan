<?php
declare(strict_types=1);

namespace App\Services;

use App\Models\PlannedIndicator;
use App\Models\Indicator;
use App\Models\HospitalBedProfiles;
use App\Models\MedicalAssistanceType;
use Illuminate\Http\Request;

//use App\Http\Resources\IndicatorCollection;
//use App\Http\Resources\IndicatorResource;
use Validator;

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
     * (плановый показатель = показатель,услуга,профиль,ФАП )
     *
     * @return \Illuminate\Http\Response
     */
    public function plannedIndicatorsForNodeId(int $nodeId)
    {
        return PlannedIndicator::select('id')->where('node_id', $nodeId)->pluck('id');
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

}
