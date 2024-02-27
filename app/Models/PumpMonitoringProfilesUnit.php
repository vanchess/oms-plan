<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\Relations\BelongsTo;
use Illuminate\Database\Eloquent\Relations\BelongsToMany;
use Illuminate\Database\Eloquent\SoftDeletes;

class PumpMonitoringProfilesUnit extends Model
{
    use SoftDeletes;

    protected $table = 'tbl_pump_monitoring_profiles_unit';

    public function plannedIndicators(): BelongsToMany
    {
        return $this->belongsToMany(PlannedIndicator::class, 'tbl_pump_monitoring_profiles_unit_planned_indicators',
                      'monitoring_profile_unit_id', 'planned_indicator_id');
    }

    public function pumpMonitoringProfiles(): BelongsTo
    {
        return $this->belongsTo(PumpMonitoringProfiles::class, 'id', 'monitoring_profile_id');
    }

    public function unit(): BelongsTo
    {
        return $this->belongsTo(PumpUnit::class, 'unit_id', 'id');
    }
}
