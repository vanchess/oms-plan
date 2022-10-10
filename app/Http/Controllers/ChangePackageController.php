<?php

namespace App\Http\Controllers;

use App\Http\Resources\ChangePackageCollection;
use App\Models\ChangePackage;
use Illuminate\Http\Request;

class ChangePackageController extends Controller
{
    /**
     * Display a listing of the resource.
     *
     * @return \Illuminate\Http\Response
     */
    public function index()
    {
        return new ChangePackageCollection(ChangePackage::all());
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
     * @param  \App\Models\ChangePackage  $changePackage
     * @return \Illuminate\Http\Response
     */
    public function show(ChangePackage $changePackage)
    {
        //
    }

    /**
     * Update the specified resource in storage.
     *
     * @param  \Illuminate\Http\Request  $request
     * @param  \App\Models\ChangePackage  $changePackage
     * @return \Illuminate\Http\Response
     */
    public function update(Request $request, ChangePackage $changePackage)
    {
        //
    }

    /**
     * Remove the specified resource from storage.
     *
     * @param  \App\Models\ChangePackage  $changePackage
     * @return \Illuminate\Http\Response
     */
    public function destroy(ChangePackage $changePackage)
    {
        //
    }
}
