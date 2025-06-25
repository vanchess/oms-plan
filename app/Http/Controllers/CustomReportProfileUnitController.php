<?php

namespace App\Http\Controllers;

use App\Http\Controllers\Controller;
use App\Models\CustomReportProfile;
use App\Models\CustomReportProfileUnit;
use Illuminate\Http\Request;

class CustomReportProfileUnitController extends Controller
{
    public function store(Request $request)
    {
        $validated = $request->validate([
            'profile_id' => 'required|exists:tbl_custom_report_profile,id',
            'unit_id' => 'required|exists:tbl_custom_report_unit,id',
        ]);

        $existing = CustomReportProfileUnit::where($validated)->first();
        if ($existing) {
            return response()->json(['message' => 'Связь уже существует'], 409);
        }

        return CustomReportProfileUnit::create($validated);
    }

    public function destroy($id)
    {
        $unitLink = CustomReportProfileUnit::with('profile')->findOrFail($id);

        $unitLink->delete();

        return response()->json(['message' => 'Deleted']);
    }

    public function attachPlannedIndicator(Request $request, $profileUnitId)
    {
        $validated = $request->validate([
            'planned_indicator_id' => 'required|exists:tbl_planned_indicators,id',
        ]);

        $profileUnit = CustomReportProfileUnit::findOrFail($profileUnitId);

        // Привязка
        $profileUnit->plannedIndicators()->attach($validated['planned_indicator_id']);

        return response()->json(['message' => 'Индикатор успешно привязан']);
    }

    public function detachPlannedIndicator(Request $request, $profileUnitId)
    {
        $validated = $request->validate([
            'planned_indicator_id' => 'required|exists:tbl_planned_indicators,id',
        ]);

        $profileUnit = CustomReportProfileUnit::findOrFail($profileUnitId);

        $profileUnit->plannedIndicators()->detach($validated['planned_indicator_id']);

        return response()->json(['message' => 'Индикатор успешно отвязан']);
    }
}
