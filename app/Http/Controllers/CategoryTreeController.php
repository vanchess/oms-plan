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
            'node' => 'required|integer|min:1|max:40'
        ]);
        if ($validator->fails()) {
            return response()->json($validator->errors()->toJson(), 400);
        }

        return $nodeService->nodeWithChildrenIds($nodeId);
    }
}
