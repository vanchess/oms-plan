<?php

namespace App\Http\Controllers;

use App\Http\Resources\CareProfilesCollection;
use App\Models\CareProfiles;
use Illuminate\Http\Request;

class CareProfilesController extends Controller
{
    /**
     * Display a listing of the resource.
     *
     * @return \Illuminate\Http\Response
     */
    public function index()
    {
        return new CareProfilesCollection(CareProfiles::OrderBy('id')->get());
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
     * @param  \App\Models\CareProfiles  $careProfiles
     * @return \Illuminate\Http\Response
     */
    public function show(CareProfiles $careProfiles)
    {
        //
    }

    /**
     * Update the specified resource in storage.
     *
     * @param  \Illuminate\Http\Request  $request
     * @param  \App\Models\CareProfiles  $careProfiles
     * @return \Illuminate\Http\Response
     */
    public function update(Request $request, CareProfiles $careProfiles)
    {
        //
    }

    /**
     * Remove the specified resource from storage.
     *
     * @param  \App\Models\CareProfiles  $careProfiles
     * @return \Illuminate\Http\Response
     */
    public function destroy(CareProfiles $careProfiles)
    {
        //
    }
}
