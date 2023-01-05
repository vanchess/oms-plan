<?php

namespace App\Http\Controllers;

use App\Jobs\InitialDataLoaded;
use App\Services\InitialDataFixingService;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Auth;
use Illuminate\Support\Facades\Validator;

class InitialDataLoadedController extends Controller
{
    /**
     * Display a listing of the resource.
     *
     * @return \Illuminate\Http\Response
     */
    public function index(Request $request, InitialDataFixingService $initialDataFixingService)
    {
        $validator = Validator::make($request->all(),[
            'year' => 'required|integer|min:2020|max:2099'
        ]);
        if ($validator->fails()) {
            return response()->json($validator->errors()->toJson(), 400);
        }
        $year = (int)$request->year;

        $data = $initialDataFixingService->getLoadedNodeIdsByYear($year);
        //var_dump($data[0]);
        //return $data;
        //dd($data);
        return $data;
    }

    /**
     * Зафиксировать начальные данные по заданным разделам(nodes), на заданный год
     *
     * @param  \Illuminate\Http\Request  $request
     * @return \Illuminate\Http\Response
     */
    public function store(Request $request)
    {
        $validator = Validator::make($request->all(),[
            'year' => 'required|integer|min:2020|max:2099',
            'nodes' => 'required|array',
            'nodes.*' => 'required|integer|distinct|min:1|max:40',
        ]);
        if ($validator->fails()) {
            return response()->json($validator->errors()->toJson(), 400);
        }
        $userId = Auth::id();
        $year = (int)$validator->validated()['year'];
        $nodes = $validator->validated()['nodes'];

        foreach ($nodes as $nodeId) {
            InitialDataLoaded::dispatch($nodeId, $year, $userId);
        }
        return "OK";
    }

    /**
     * Display the specified resource.
     *
     * @param  int  $id
     * @return \Illuminate\Http\Response
     */
    public function show($id)
    {
        //
    }

    /**
     * Update the specified resource in storage.
     *
     * @param  \Illuminate\Http\Request  $request
     * @param  int  $id
     * @return \Illuminate\Http\Response
     */
    public function update(Request $request, $id)
    {
        //
    }

    /**
     * Remove the specified resource from storage.
     *
     * @param  int  $id
     * @return \Illuminate\Http\Response
     */
    public function destroy($id)
    {
        //
    }
}
