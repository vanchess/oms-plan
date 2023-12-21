<?php

namespace App\Http\Controllers;

use App\Services\NodeService;
use Illuminate\Http\Request;

//use App\Http\Resources\IndicatorCollection;
//use App\Http\Resources\IndicatorResource;
use Validator;

class UsedIndicatorsController extends Controller
{
    /**
     * Возвращает массив id показателей используемых для переданного id узла дерева категорий
     *
     * @return \Illuminate\Http\Response
     */
    public function indicatorsUsedForNodeId(Request $request, NodeService $nodeService)
    {
        $validator = Validator::make($request->all(),[
            'node' => 'required|integer|min:1|max:45'
        ]);
        if ($validator->fails()) {
            return response()->json($validator->errors()->toJson(), 400);
        }
        $nodeId = $request->node;

        return $nodeService->indicatorsUsedForNodeId($nodeId);


    }

}
