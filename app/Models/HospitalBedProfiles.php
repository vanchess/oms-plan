<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class HospitalBedProfiles extends Model
{
    use HasFactory;
    
    protected $dates = ['deleted_at'];
    protected $table = 'tbl_hospital_bed_profiles';
}
