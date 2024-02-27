<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\Relations\HasMany;
use Illuminate\Database\Eloquent\SoftDeletes;

class PumpMonitoringProfiles extends Model
{
    use SoftDeletes;

    protected $table = "tbl_pump_monitoring_profiles";

    public function profilesUnits(): HasMany
    {
        return $this->hasMany(PumpMonitoringProfilesUnit::class, 'monitoring_profile_id', 'id');
    }
}
