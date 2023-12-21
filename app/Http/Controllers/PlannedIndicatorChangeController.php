<?php

namespace App\Http\Controllers;

use App\Http\Resources\PlannedIndicatorChangeCollection;
use App\Http\Resources\PlannedIndicatorChangeResource;
use App\Models\PlannedIndicator;
use App\Models\PlannedIndicatorChange;
use App\Rules\BCMathString;
use App\Services\Dto\PlannedIndicatorChangeValueCollectionDto;
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
            'nodes'   => 'required|array',
            'nodes.*' => 'required|integer|min:1|max:45',
            'year' => 'required|integer|min:2020|max:2099'
        ]);
        if ($validator->fails()) {
            return response()->json($validator->errors()->toJson(), 400);
        }
        $nodeIds = $request->nodes;
        $year = (int)$request->year;

        $data = $plannedIndicatorChangeInitService->getByNodeIdsAndYear($nodeIds, $year);
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

    public function incrementValues(Request $request, PlannedIndicatorChangeService $plannedIndicatorChangeService)
    {
        bcscale(4);

        $user = Auth::user();
        $userId = $user->id;

        $validator = Validator::make($request->all(), [
            'values' => 'array',
            'values.*.periodId' => 'required|integer|min:1',
            'values.*.moId' => 'required|integer',
            'values.*.moDepartmentId' => 'integer',
            'values.*.plannedIndicatorId' => 'required|integer',
            'values.*.packageId' => 'integer',
            'values.*.value' => ['required', new BCMathString],
            'total' => ['required', new BCMathString],
        ]);
        if ($validator->fails()) {
            return response()->json($validator->errors()->toJson(), 400);
        }
        $validated = $validator->validated();

        $sum = '0';
        for ($i = 0; $i < count($validated['values']); $i++) {
            $sum = bcadd($sum, $validated['values'][$i]['value']);
        }
        if (bccomp($sum, $validated['total']) !== 0)
        {
            return response()->json('Ошибка проверки контрольного значения', 400);
        }

        $arr = [];
        foreach ($validated['values'] as $v) {
            // пропускаем нулевые значения
            if (bccomp((string)$v['value'],'0') === 0) {
                continue;
            }

            $departmentId = null;
            if(isset($v['moDepartmentId'])) {
                $departmentId = (int)$v['moDepartmentId'];
            }
            $arr[] = new PlannedIndicatorChangeValueDto(
                periodId: (int)$v['periodId'],
                moId: (int)$v['moId'],
                moDepartmentId: $departmentId,
                plannedIndicatorId: (int)$v['plannedIndicatorId'],
                packageId: (int)$v['packageId'],
                value: (string)$v['value'],
                userId: $userId
            );
        }
        $collection = new PlannedIndicatorChangeValueCollectionDto(...$arr);

        return new PlannedIndicatorChangeCollection($plannedIndicatorChangeService->incrementValues($collection)->value->toArray());
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
