<?php

namespace App\Http\Controllers;

use App\Models\PlannedIndicator;
use Illuminate\Http\Request;

use App\Http\Resources\PlannedIndicatorCollection;
use Illuminate\Support\Facades\Validator;

class PlannedIndicatorController extends Controller
{
    /**
     * Display a listing of the resource.
     *
     * @return \Illuminate\Http\Response
     */
    public function index(Request $request)
    {
        $validator = Validator::make($request->all(),[
            'year' => 'required|integer|min:2020|max:2099'
        ]);
        if ($validator->fails()) {
            return response()->json($validator->errors()->toJson(), 400);
        }

        $validated = $validator->validated();
        $date = date('Y-m-d');
        if (isset($validated['year'])) {
            $date = ($validated['year'].'-01-02');
        }

        return new PlannedIndicatorCollection(PlannedIndicator::WhereRaw("? BETWEEN effective_from AND effective_to", [$date])->get());
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
