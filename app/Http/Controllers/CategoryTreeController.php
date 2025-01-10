<?php

namespace App\Http\Controllers;

use App\Services\NodeService;
// use Illuminate\Http\Request;
// use Illuminate\Support\Facades\Validator;
use Illuminate\Validation\Factory as Validator;

class CategoryTreeController extends Controller
{
    public function nodeWithChildren(Validator $v, NodeService $nodeService, int $nodeId)
    {
        $validator = $v->make(['node' => $nodeId],[
            'node' => 'required|integer|min:1|max:90'
        ]);
        if ($validator->fails()) {
            return response()->json($validator->errors()->toJson(), 400);
        }

        return $nodeService->nodeWithChildrenIds($nodeId);
    }

    public function getCategoryTree(Validator $v, NodeService $nodeService, int $rootNodeId)
    {
        $validator = $v->make(['node' => $rootNodeId],[
            'node' => 'required|integer|min:1|max:90'
        ]);
        if ($validator->fails()) {
            return response()->json($validator->errors()->toJson(), 400);
        }

        return $nodeService->nodeTree($rootNodeId);
    }
}
