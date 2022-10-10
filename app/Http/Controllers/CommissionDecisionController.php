<?php

namespace App\Http\Controllers;

use App\Http\Resources\CommissionDecisionCollection;
use App\Http\Resources\CommissionDecisionResource;
use App\Models\ChangePackage;
use App\Models\CommissionDecision;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Auth;
use Illuminate\Support\Facades\DB;
use Illuminate\Support\Facades\Validator;

class CommissionDecisionController extends Controller
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
        $year = (int)$validator->validated()['year'];

        $cd = CommissionDecision::whereYear('date',$year)->get();

        return new CommissionDecisionCollection($cd);
    }

    /**
     * Store a newly created resource in storage.
     *
     * @param  \Illuminate\Http\Request  $request
     * @return \Illuminate\Http\Response
     */
    public function store(Request $request)
    {
        $userId = Auth::id();

        $validator = Validator::make($request->all(), [
            'number' => 'required|string',
            'date' => 'required|date',
            'description' => 'nullable|string',
        ]);
        if ($validator->fails()) {
            return response()->json($validator->errors()->toJson(), 400);
        }

        $validated = $validator->validated();

        $commissionDecision = new CommissionDecision();
        $commissionDecision->number = $validated['number'];
        $commissionDecision->date = $validated['date'];
        if(isset($validated['description'])) {
            $commissionDecision->description = $validated['description'];
        }
        $commissionDecision->created_by = $userId;


        $changePackage = new ChangePackage();
        $changePackage->user_id = $userId;


        DB::transaction(function() use ($commissionDecision, $changePackage) {
            $commissionDecision->save();
            $changePackage->commission_decision_id = $commissionDecision->id;
            $changePackage->save();
        });

        return new CommissionDecisionResource($commissionDecision);
    }

    /**
     * Display the specified resource.
     *
     * @param  \App\Models\CommissionDecision  $commissionDecision
     * @return \Illuminate\Http\Response
     */
    public function show(CommissionDecision $commissionDecision)
    {
        //
    }

    /**
     * Update the specified resource in storage.
     *
     * @param  \Illuminate\Http\Request  $request
     * @param  \App\Models\CommissionDecision  $commissionDecision
     * @return \Illuminate\Http\Response
     */
    public function update(Request $request, CommissionDecision $commissionDecision)
    {
        //
    }

    /**
     * Remove the specified resource from storage.
     *
     * @param  \App\Models\CommissionDecision  $commissionDecision
     * @return \Illuminate\Http\Response
     */
    public function destroy(CommissionDecision $commissionDecision)
    {
        //
    }
}
