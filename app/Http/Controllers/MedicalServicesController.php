<?php

namespace App\Http\Controllers;

use App\Models\MedicalServices;
use App\Http\Resources\MedicalServicesCollection;
use Illuminate\Http\Request;

class MedicalServicesController extends Controller
{
    /**
     * Display a listing of the resource.
     *
     * @return \Illuminate\Http\Response
     */
    public function index()
    {
        return new MedicalServicesCollection(MedicalServices::all());
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
     * @param  \App\Models\MedicalServices  $medicalServices
     * @return \Illuminate\Http\Response
     */
    public function show(MedicalServices $medicalServices)
    {
        //
    }

    /**
     * Update the specified resource in storage.
     *
     * @param  \Illuminate\Http\Request  $request
     * @param  \App\Models\MedicalServices  $medicalServices
     * @return \Illuminate\Http\Response
     */
    public function update(Request $request, MedicalServices $medicalServices)
    {
        //
    }

    /**
     * Remove the specified resource from storage.
     *
     * @param  \App\Models\MedicalServices  $medicalServices
     * @return \Illuminate\Http\Response
     */
    public function destroy(MedicalServices $medicalServices)
    {
        //
    }
}
