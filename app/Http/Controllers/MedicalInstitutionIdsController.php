<?php

namespace App\Http\Controllers;

use App\Models\MedicalInstitution;
use App\Models\MedicalInstitutionDepartment;
use Illuminate\Http\Request;
use Validator;

class MedicalInstitutionIdsController extends Controller
{
    /**
     * Display a listing of the resource.
     *
     * @return \Illuminate\Http\Response
     */
    public function havingDepartments(Request $request)
    {
        $validator = Validator::make($request->all(), [
            'departmentTypes'   => 'array',
            'departmentTypes.*' => 'integer|distinct|exists:App\Models\MedicalInstitutionDepartmentType,id',
        ]);

        if ($validator->fails()) {
            return response()->json($validator->errors()->toJson(), 400);
        }
        
        $validated = $validator->validated();
        
        $sqlQuery = MedicalInstitutionDepartment::select('v_mo.id')
            ->join('v_mo', 'v_mo.id', '=', 'mo_id');
        
        if(!empty($validated['departmentTypes'])) {
            $sqlQuery = $sqlQuery->whereIn('type_id', $validated['departmentTypes']);
        }
        
        return $sqlQuery->GroupBy('v_mo.id','v_mo.order')->OrderBy('v_mo.order')->pluck('v_mo.id');
    }

}
