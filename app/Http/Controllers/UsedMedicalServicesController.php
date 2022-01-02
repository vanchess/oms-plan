<?php

namespace App\Http\Controllers;

use App\Services\NodeService;
use Illuminate\Http\Request;

//use App\Http\Resources\IndicatorCollection;
//use App\Http\Resources\IndicatorResource;
use Validator;

class UsedMedicalServicesController extends Controller
{
    /**
     * Возвращает массив id профилей коек для переданного id узла дерева категорий
     *
     * @return \Illuminate\Http\Response
     */
    public function medicalServicesUsedForNodeId(Request $request, NodeService $nodeService)
    {
        $validator = Validator::make($request->all(),[
            'node' => 'required|integer|min:1|max:40'
        ]);
        if ($validator->fails()) {
            return response()->json($validator->errors()->toJson(), 400);
        }
        $nodeId = $request->node;
        
        return $nodeService->medicalServicesForNodeId($nodeId); 
    }

}
