<?php

namespace App\Http\Controllers;

use App\Models\HospitalBedProfiles;
use Illuminate\Http\Request;

use App\Http\Resources\HospitalBedProfilesCollection;
// use App\Http\Resources\HospitalBedProfilesResource;

class HospitalBedProfilesController extends Controller
{
    /**
     * Display a listing of the resource.
     *
     * @return \Illuminate\Http\Response
     */
    public function index()
    {
        return new HospitalBedProfilesCollection(HospitalBedProfiles::OrderBy('id')->get());//->paginate($perPage));
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
     * @param  \App\Models\HospitalBedProfiles  $hospitalBedProfiles
     * @return \Illuminate\Http\Response
     */
    public function show(HospitalBedProfiles $hospitalBedProfiles)
    {
        //
    }

    /**
     * Update the specified resource in storage.
     *
     * @param  \Illuminate\Http\Request  $request
     * @param  \App\Models\HospitalBedProfiles  $hospitalBedProfiles
     * @return \Illuminate\Http\Response
     */
    public function update(Request $request, HospitalBedProfiles $hospitalBedProfiles)
    {
        //
    }

    /**
     * Remove the specified resource from storage.
     *
     * @param  \App\Models\HospitalBedProfiles  $hospitalBedProfiles
     * @return \Illuminate\Http\Response
     */
    public function destroy(HospitalBedProfiles $hospitalBedProfiles)
    {
        //
    }
}
