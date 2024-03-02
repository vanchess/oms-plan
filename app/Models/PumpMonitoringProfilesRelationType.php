<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;

class PumpMonitoringProfilesRelationType extends Model
{
    use SoftDeletes;

    protected $table = 'tbl_pump_monitoring_profiles_relation_type';

    protected $fillable = [
        'name',
        'slug'
    ];
}
