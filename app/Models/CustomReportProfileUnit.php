<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;

class CustomReportProfileUnit extends Model
{
    use SoftDeletes;

    protected $table = 'tbl_custom_report_profile_unit';
    protected $fillable = ['profile_id', 'unit_id'];

    public function profile()
    {
        return $this->belongsTo(CustomReportProfile::class, 'profile_id');
    }

    public function unit()
    {
        return $this->belongsTo(CustomReportUnit::class, 'unit_id');
    }

    public function plannedIndicators()
    {
        return $this->belongsToMany(
            PlannedIndicator::class,
            'tbl_custom_report_profile_unit_planned_indicator',
            'profile_unit_id',
            'planned_indicator_id'
        );
    }
}
