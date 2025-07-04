<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use App\Http\Controllers\Controller;
use App\Models\CustomReport;

class CustomReportController extends Controller
{
    public function index()
    {
        return CustomReport::with('profiles')->get();
    }

    public function store(Request $request)
    {
        $validated = $request->validate([
            'name' => 'required|string|max:256|unique:tbl_custom_report,name',
            'short_name' => 'nullable|string|max:255',
        ]);

        $validated['user_id'] = auth()->id();

        return CustomReport::create($validated);
    }

    public function show($id)
    {
        return CustomReport::with('profiles.units.unit')
            ->findOrFail($id);
    }

    public function update(Request $request, $id)
    {
        $report = CustomReport::findOrFail($id);

        $validated = $request->validate([
            'name' => 'required|string|max:256|unique:tbl_custom_report,name,' . $id,
            'short_name' => 'nullable|string|max:255',
        ]);

        $report->update($validated);

        return $report;
    }

    public function destroy($id)
    {
        $report = CustomReport::where('user_id', auth()->id())->findOrFail($id);
        $report->delete();

        return response()->json(['message' => 'Deleted']);
    }

    public function profiles($id)
    {
        $report = CustomReport::with('profiles')->findOrFail($id);
        return response()->json($report->profiles);
    }
}
