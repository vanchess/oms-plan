<?php

namespace App\Http\Controllers;

use App\Http\Resources\PlannedIndicatorChangeCollection;
use App\Http\Resources\PlannedIndicatorChangeResource;
use App\Rules\BCMathString;
use App\Services\Dto\PlannedIndicatorChangeValueDto;
use App\Services\PlannedIndicatorChangeService;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Auth;
use Illuminate\Support\Facades\Validator;

class PlannedIndicatorChangeController extends Controller
{
    /**
     * Display a listing of the resource.
     *
     * @return \Illuminate\Http\Response
     */
    public function index(Request $request, PlannedIndicatorChangeService $plannedIndicatorChangeInitService)
    {
        $validator = Validator::make($request->all(),[
            'node' => 'required|integer|min:1|max:40',
            'year' => 'required|integer|min:2020|max:2099'
        ]);
        if ($validator->fails()) {
            return response()->json($validator->errors()->toJson(), 400);
        }
        $nodeId = (int)$request->node;
        $year = (int)$request->year;

        $data = $plannedIndicatorChangeInitService->getByNodeIdAndYear($nodeId, $year);
        //var_dump($data[0]);
        //return $data;
        //dd($data);
        return new PlannedIndicatorChangeCollection(collect($data));
    }

    /**
     * Store a newly created resource in storage.
     *
     * @param  \Illuminate\Http\Request  $request
     * @return \Illuminate\Http\Response
     */
    public function store(Request $request, PlannedIndicatorChangeService $plannedIndicatorChangeService)
    {
        $user = Auth::user();
        $userId = $user->id;

        $validator = Validator::make($request->all(), [
            'periodId' => 'required|integer|min:1',
            'moId' => 'required|integer',
            'moDepartmentId' => 'integer',
            'plannedIndicatorId' => 'required|integer',
            'value' => ['required', new BCMathString],
        ]);
        if ($validator->fails()) {
            return response()->json($validator->errors()->toJson(), 400);
        }

        $validated = $validator->validated();
        $departmentId = null;
        if(isset($validated['moDepartmentId'])) {
            $departmentId = (int)$validated['moDepartmentId'];
        }
        $dto = new PlannedIndicatorChangeValueDto(
            periodId: (int)$validated['periodId'],
            moId: (int)$validated['moId'],
            moDepartmentId: $departmentId,
            plannedIndicatorId: (int)$validated['plannedIndicatorId'],
            value: (string)$validated['value'],
            userId: $userId
        );

        return new PlannedIndicatorChangeResource($plannedIndicatorChangeService->setValue($dto)->value);
    }

    /**
     * Display the specified resource.
     *
     * @param  \App\Models\PlannedIndicatorChange  $plannedIndicatorChange
     * @return \Illuminate\Http\Response
     */
    public function show()
    {
        //
    }

    /**
     * Update the specified resource in storage.
     *
     * @param  \Illuminate\Http\Request  $request
     * @param  \App\Models\PlannedIndicatorChange  $plannedIndicatorChange
     * @return \Illuminate\Http\Response
     */
    public function update(Request $request)
    {
        //
    }

    /**
     * Remove the specified resource from storage.
     *
     * @param  \App\Models\PlannedIndicatorChange  $plannedIndicatorChange
     * @return \Illuminate\Http\Response
     */
    public function destroy()
    {
        //
    }
}
