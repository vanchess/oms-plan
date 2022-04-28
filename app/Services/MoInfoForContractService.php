<?php
declare(strict_types=1);

namespace App\Services;

use App\Models\MedicalInstitution;
use App\Models\MedicalInstitutionDepartment;
use Illuminate\Support\Facades\DB;

class MoInfoForContractService
{
    public function GetJson() {
        return $this->CreateData()->toJson();

    }

    private function CreateData() {
        // MedicalInstitution::select()/

        $data = DB::table((new MedicalInstitution())->getTable().' as mo')
            ->selectRaw('mo.id, mo.code, mo.name, mo.short_name, mo.address, chief, "order", organization_id, SUM(CASE WHEN fap.type_id=49 THEN 1 ELSE 0 END) as fap_count, SUM(CASE WHEN fap.type_id=50 THEN 1 ELSE 0 END) as fp_count')
            ->leftJoin((new MedicalInstitutionDepartment())->getTable().' as fap', function($join) {
                $join->on('mo.id', '=', 'fap.mo_id')->whereIn('fap.type_id',[49,50]);
            })
            ->whereNull('mo.deleted_at')
            ->whereNull('fap.deleted_at')
            ->groupBy('mo.id', 'mo.code', 'mo.name', 'mo.short_name', 'mo.address', 'mo.chief', 'mo.order', 'mo.organization_id')
            ->get()
            ->mapWithKeys(function($item) {return [$item->id => $item];});

        /*
        $faps = MedicalInstitutionDepartment::selectRaw('mo_id, id, name, type_id')
            ->whereIn('type_id',[49,50])
            ->whereNull('deleted_at')
            ->get();

        $faps = $faps->groupBy([
            'mo_id',
            function ($fap) {
                if($fap->type_id === 49) return 'fap';
                if($fap->type_id === 50) return 'fp';
                return 'error';
            },
            'id'
        ]);
        */
        return $data;
    }
}
