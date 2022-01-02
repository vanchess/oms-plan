<?php

namespace App\Http\Controllers;

use App\Models\MedicalAssistanceType;
use App\Http\Resources\MedicalAssistanceTypeCollection;
use Illuminate\Http\Request;

class MedicalAssistanceTypeController extends Controller
{
    /**
     * Display a listing of the resource.
     *
     * @return \Illuminate\Http\Response
     */
    public function index()
    {
        return new MedicalAssistanceTypeCollection(MedicalAssistanceType::all());
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
     * @param  \App\Models\MedicalAssistanceType  $medicalAssistanceType
     * @return \Illuminate\Http\Response
     */
    public function show(MedicalAssistanceType $medicalAssistanceType)
    {
        //
    }

    /**
     * Update the specified resource in storage.
     *
     * @param  \Illuminate\Http\Request  $request
     * @param  \App\Models\MedicalAssistanceType  $medicalAssistanceType
     * @return \Illuminate\Http\Response
     */
    public function update(Request $request, MedicalAssistanceType $medicalAssistanceType)
    {
        //
    }

    /**
     * Remove the specified resource from storage.
     *
     * @param  \App\Models\MedicalAssistanceType  $medicalAssistanceType
     * @return \Illuminate\Http\Response
     */
    public function destroy(MedicalAssistanceType $medicalAssistanceType)
    {
        //
    }
}
