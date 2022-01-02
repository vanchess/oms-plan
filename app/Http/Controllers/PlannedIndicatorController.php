<?php

namespace App\Http\Controllers;

use App\Models\PlannedIndicator;
use Illuminate\Http\Request;

use App\Http\Resources\PlannedIndicatorCollection;

class PlannedIndicatorController extends Controller
{
    /**
     * Display a listing of the resource.
     *
     * @return \Illuminate\Http\Response
     */
    public function index()
    {
        return new PlannedIndicatorCollection(PlannedIndicator::all());
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
     * @param  \App\Models\PlannedIndicator  $plannedIndicator
     * @return \Illuminate\Http\Response
     */
    public function show(PlannedIndicator $plannedIndicator)
    {
        //
    }

    /**
     * Update the specified resource in storage.
     *
     * @param  \Illuminate\Http\Request  $request
     * @param  \App\Models\PlannedIndicator  $plannedIndicator
     * @return \Illuminate\Http\Response
     */
    public function update(Request $request, PlannedIndicator $plannedIndicator)
    {
        //
    }

    /**
     * Remove the specified resource from storage.
     *
     * @param  \App\Models\PlannedIndicator  $plannedIndicator
     * @return \Illuminate\Http\Response
     */
    public function destroy(PlannedIndicator $plannedIndicator)
    {
        //
    }
}
