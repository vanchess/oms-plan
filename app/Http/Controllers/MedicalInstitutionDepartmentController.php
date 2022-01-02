<?php

namespace App\Http\Controllers;

use App\Models\MedicalInstitutionDepartment;
use Illuminate\Http\Request;
use Validator;

use App\Http\Resources\MedicalInstitutionDepartmentCollection;

class MedicalInstitutionDepartmentController extends Controller
{
    /**
     * Display a listing of the resource.
     *
     * @return \Illuminate\Http\Response
     */
    public function index(Request $request)
    {
        $validator = Validator::make($request->all(), [
            'departmentTypes'   => 'array',
            'departmentTypes.*' => 'integer|distinct|exists:App\Models\MedicalInstitutionDepartmentType,id',
        ]);

        if ($validator->fails()) {
            return response()->json($validator->errors()->toJson(), 400);
        }
        
        $validated = $validator->validated();
        
        $sqlQuery = null;
        
        if(!empty($validated['departmentTypes'])) {
            $sqlQuery = MedicalInstitutionDepartment::whereIn('type_id', $validated['departmentTypes'])->OrderBy('id');
        } else {
            $sqlQuery = MedicalInstitutionDepartment::OrderBy('id');
        }
        
        return new MedicalInstitutionDepartmentCollection($sqlQuery->get());
    }

    /**
     * Store a newly created resource in storage.
     *
     * @param  \Illuminate\Http\Request  $request
     * @return \Illuminate\Http\Response
     */
    public function store(Request $request)
    {
        //
    }

    /**
     * Display the specified resource.
     *
     * @param  \App\Models\MedicalInstitutionDepartment  $medicalInstitutionDepartment
     * @return \Illuminate\Http\Response
     */
    public function show(MedicalInstitutionDepartment $medicalInstitutionDepartment)
    {
        //
    }

    /**
     * Update the specified resource in storage.
     *
     * @param  \Illuminate\Http\Request  $request
     * @param  \App\Models\MedicalInstitutionDepartment  $medicalInstitutionDepartment
     * @return \Illuminate\Http\Response
     */
    public function update(Request $request, MedicalInstitutionDepartment $medicalInstitutionDepartment)
    {
        //
    }

    /**
     * Remove the specified resource from storage.
     *
     * @param  \App\Models\MedicalInstitutionDepartment  $medicalInstitutionDepartment
     * @return \Illuminate\Http\Response
     */
    public function destroy(MedicalInstitutionDepartment $medicalInstitutionDepartment)
    {
        //
    }
}
