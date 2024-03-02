<?php

namespace App\Models;

// use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\Relations\BelongsTo;
use Illuminate\Database\Eloquent\Relations\BelongsToMany;
use Illuminate\Database\Eloquent\SoftDeletes;

class PlannedIndicator extends Model
{
    use SoftDeletes;

    protected $table = 'tbl_planned_indicators';

    protected $casts = [
        'indicator_id' => 'integer',
    ];

    public function initialData()
    {
        return $this->hasMany(InitialData::class, 'planned_indicator_id', 'id');
    }

    public function pumpMonitoringProfilesUnits(): BelongsToMany
    {
        return $this->belongsToMany(PlannedIndicator::class, 'tbl_pump_monitoring_profiles_unit_planned_indicators',
                'planned_indicator_id', 'monitoring_profile_unit_id');
    }

    public function assistanceType(): BelongsTo
    {
        return $this->belongsTo(MedicalAssistanceType::class, 'assistance_type_id', 'id');
    }

    public function careProfile(): BelongsTo
    {
        return $this->belongsTo(CareProfiles::class, 'care_profile_id', 'id');
    }

    public function indicator(): BelongsTo
    {
        return $this->belongsTo(Indicator::class, 'indicator_id', 'id');
    }

    public function node(): BelongsTo
    {
        return $this->belongsTo(CategoryTreeNodes::class, 'node_id', 'id');
    }

    public function bedProfile(): BelongsTo
    {
        return $this->belongsTo(HospitalBedProfiles::class, 'profile_id', 'id');
    }

    public function service(): BelongsTo
    {
        return $this->belongsTo(MedicalServices::class, 'service_id', 'id');
    }

    public function vmpGroup(): BelongsTo
    {
        return $this->belongsTo(VmpGroup::class, 'vmp_group_id', 'id');
    }

    public function vmpType(): BelongsTo
    {
        return $this->belongsTo(VmpTypes::class, 'vmp_type_id', 'id');
    }
}
