<?php

namespace App\Http\Controllers;

use App\Http\Controllers\Controller;
use App\Models\CustomReportProfile;
use Illuminate\Http\Request;

class CustomReportProfileController extends Controller
{
    public function store(Request $request)
    {
        $validated = $request->validate([
            'custom_report_id' => 'required|exists:tbl_custom_report,id',
            'parent_id' => 'nullable|exists:tbl_custom_report_profile,id',
            'code' => 'required|string|max:64|unique:tbl_custom_report_profile,code',
            'name' => 'required|string|max:1024',
            'short_name' => 'required|string|max:256',
            'relation_type_id' => 'nullable|exists:tbl_custom_report_profile_relation_type,id',
            'effective_from' => 'nullable|date',
            'effective_to' => 'nullable|date',
            'order' => 'nullable|integer'
        ]);

        $validated['user_id'] = auth()->id();

        return CustomReportProfile::create($validated);
    }

    public function update(Request $request, $id)
    {
        $profile = CustomReportProfile::findOrFail($id);

        $validated = $request->validate([
            'name' => 'sometimes|string|max:1024',
            'short_name' => 'sometimes|string|max:256',
            'effective_to' => 'sometimes|date',
        ]);

        $profile->update($validated);

        return $profile;
    }
}
