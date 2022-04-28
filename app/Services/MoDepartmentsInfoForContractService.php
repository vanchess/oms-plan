<?php
declare(strict_types=1);

namespace App\Services;

use App\Models\MedicalInstitutionDepartment;

class MoDepartmentsInfoForContractService
{
    public function GetJson($typeIds = [49,50]) {
        return $this->CreateData($typeIds)->toJson();

    }

    private function CreateData($typeIds = [49,50]) {
        $data = MedicalInstitutionDepartment::selectRaw('mo_id, id, name, type_id')
            ->whereIn('type_id',$typeIds)
            ->whereNull('deleted_at')
            ->get()
            ->mapWithKeys(function($item) {return [$item->id => $item];});

        return $data;
    }
}
