<?php

namespace App\Http\Controllers;

use App\Http\Controllers\Controller;
use App\Models\CustomReportProfile;
use Illuminate\Http\Request;

class CustomReportProfileController extends Controller
{
    public function show($id)
    {
        $profile = CustomReportProfile::with([
            //'units.unitType',
            //'units.plannedIndicators',
            'relationType',
            'parent',
        ])->findOrFail($id);

        return response()->json($profile);
    }

    public function store(Request $request)
    {
        $validated = $request->validate([
            'custom_report_id' => 'required|exists:tbl_custom_report,id',
            'parent_id' => 'nullable|exists:tbl_custom_report_profile,id',
            'code' => 'required|string|max:64|unique:tbl_custom_report_profile,code',
            'name' => 'required|string|max:1024',
            'short_name' => 'required|string|max:256',
            'relation_type_id' => [
            'nullable',
            'exists:tbl_custom_report_profile_relation_type,id',
                function ($attribute, $value, $fail) use ($request) {
                    if ($request->filled('parent_id') && !$value) {
                        $fail('Поле "Тип связи с родителем" обязательно, если указан родитель.');
                    }
                },
            ],
            'effective_from' => 'nullable|date',
            'effective_to' => 'nullable|date',
            'order' => 'nullable|integer'
        ]);

        if (!isset($validated['order']) || is_null($validated['order'] ?? null)) {
            $query = CustomReportProfile::where('custom_report_id', $validated['custom_report_id']);

            if (!empty($validated['parent_id'])) {
                $query->where('parent_id', $validated['parent_id']);
            }

            $maxOrder = $query->max('order') ?? 0;
            $validated['order'] = $maxOrder + 1;
        }

        if (is_null($validated['effective_from'] ?? null)) {
            unset($validated['effective_from']);
        }
        if (is_null($validated['effective_to'] ?? null)) {
            unset($validated['effective_to']);
        }

        $validated['user_id'] = auth()->id();

        $profile = CustomReportProfile::create($validated);
        $profile->refresh();

        return response()->json($profile);
    }

    public function update(Request $request, $id)
    {
        $validated = $request->validate([
            'name' => 'required|string|max:1024',
            'short_name' => 'required|string|max:256',
            'code' => 'required|string|max:64|unique:tbl_custom_report_profile,code,' . $id,
            'parent_id' => 'nullable|exists:tbl_custom_report_profile,id|not_in:' . $id,
            'relation_type_id' => [
                'nullable',
                'exists:tbl_custom_report_profile_relation_type,id',
                function ($attribute, $value, $fail) use ($request) {
                    if ($request->filled('parent_id') && !$value) {
                        $fail('Поле "Тип связи с родителем" обязательно, если указан родитель.');
                    }
                },
            ],
            'effective_from' => 'nullable|date',
            'effective_to' => 'nullable|date',
            'order' => 'nullable|integer',
        ]);

        $validated['user_id'] = auth()->id();

        $profile = CustomReportProfile::findOrFail($id);
        $profile->update($validated);

        return response()->json(['message' => 'Profile updated']);
    }

    public function units($id)
    {
        $profile = CustomReportProfile::with('profileUnits')->findOrFail($id);
        return response()->json($profile->profileUnits);
    }

    public function updateUnits(Request $request, $id)
    {
        $validated = $request->validate([
            'unit_ids' => 'array',
            'unit_ids.*' => 'exists:tbl_custom_report_unit,id',
        ]);

        $profile = CustomReportProfile::findOrFail($id);
        $profile->profileUnits()->sync($validated['unit_ids'] ?? []);
        return response()->json(['message' => 'Units updated successfully']);
    }

    public function destroy($id)
    {
        $report = CustomReportProfile::findOrFail($id);
        $report->delete();

        return response()->json(['message' => 'Deleted']);
    }
}
