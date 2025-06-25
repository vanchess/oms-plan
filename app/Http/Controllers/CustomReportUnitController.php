<?php

namespace App\Http\Controllers;

use App\Http\Controllers\Controller;
use App\Models\CustomReportUnit;
use Illuminate\Http\Request;

class CustomReportUnitController extends Controller
{
    public function index()
    {
        return CustomReportUnit::with('type')
            ->orderBy('order')
            ->get();
    }

    public function store(Request $request)
    {
        $validated = $request->validate([
            'name' => 'required|string|max:128|unique:tbl_custom_report_unit,name',
            'type_id' => 'required|exists:tbl_indicator_types,id',
            'order' => 'nullable|integer',
        ]);

        return CustomReportUnit::create($validated);
    }

    public function update(Request $request, $id)
    {
        $unit = CustomReportUnit::findOrFail($id);

        $validated = $request->validate([
            'name' => 'required|string|max:128|unique:tbl_custom_report_unit,name,' . $id,
            'type_id' => 'required|exists:tbl_indicator_types,id',
            'order' => 'nullable|integer',
        ]);

        $unit->update($validated);

        return $unit;
    }

    public function destroy($id)
    {
        $unit = CustomReportUnit::findOrFail($id);
        $unit->delete();

        return response()->json(['message' => 'Deleted']);
    }
}
