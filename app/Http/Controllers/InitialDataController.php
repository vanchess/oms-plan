<?php
declare(strict_types=1);

namespace App\Http\Controllers;

use Illuminate\Http\Request;

use App\Services\InitialDataService;
use App\Services\Dto\InitialDataValueDto;
use App\Http\Resources\InitialDataCollection;
use App\Http\Resources\InitialDataResource;
use Validator;
use App\Rules\BCMathString;
use Illuminate\Support\Facades\Auth;

class InitialDataController extends Controller
{
    /**
     * Display a listing of the resource.
     *
     * @return \Illuminate\Http\Response
     */
    public function index(Request $request, InitialDataService $initialDataService)
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

        $data = $initialDataService->getByNodeIdAndYear($nodeId, $year);
        //var_dump($data[0]);
        //return $data;
        //dd($data);
        return new InitialDataCollection(collect($data));
    }

    /**
     * Store a newly created resource in storage.
     *
     * @param  \Illuminate\Http\Request  $request
     * @return \Illuminate\Http\Response
     */
    public function store(Request $request, InitialDataService $initialDataService)
    {
        $user = Auth::user();
        $userId = $user->id;

        $validator = Validator::make($request->all(), [
            'year' => 'required|integer|min:2020|max:2099',
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
        $dto = new InitialDataValueDto(
            year: (int)$validated['year'],
            moId: (int)$validated['moId'],
            moDepartmentId: $departmentId,
            plannedIndicatorId: (int)$validated['plannedIndicatorId'],
            value: (string)$validated['value'],
            userId: $userId
        );

        return new InitialDataResource($initialDataService->setValue($dto)->value);
        //
        /*
        $initialData = new InitialData();
        $initialData->year = $request->year;
        $initialData->mo_id = $request->mo;
        $initialData->planned_indicator_id = ;
        $initialData->value = $request->value;
        $initialData->algorithm_id =
        $initialData->user_id = $userId;
        */
    }

    /**
     * Display the specified resource.
     *
     * @param  \App\Models\InitialData  $initialData
     * @return \Illuminate\Http\Response
     */
    public function show(InitialData $initialData)
    {
        //
    }

    /**
     * Update the specified resource in storage.
     *
     * @param  \Illuminate\Http\Request  $request
     * @param  \App\Models\InitialData  $initialData
     * @return \Illuminate\Http\Response
     */
    public function update(Request $request, InitialData $initialData)
    {
        //
    }

    /**
     * Remove the specified resource from storage.
     *
     * @param  \App\Models\InitialData  $initialData
     * @return \Illuminate\Http\Response
     */
    public function destroy(InitialData $initialData)
    {
        //
    }
}
