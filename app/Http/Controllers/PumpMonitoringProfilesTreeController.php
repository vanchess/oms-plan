<?php

namespace App\Http\Controllers;

use App\Services\PumpMonitoringProfilesTreeService;
use Illuminate\Http\Request;
use Illuminate\Validation\Factory as Validator;

class PumpMonitoringProfilesTreeController extends Controller
{
    public function getTree(Validator $v, PumpMonitoringProfilesTreeService $treeService, int $rootNodeId)
    {
        $validator = $v->make(['node' => $rootNodeId],[
            'node' => 'required|integer|min:1|max:10000'
        ]);
        if ($validator->fails()) {
            return response()->json($validator->errors()->toJson(), 400);
        }

        return $treeService->nodeTree($rootNodeId);
    }
}
