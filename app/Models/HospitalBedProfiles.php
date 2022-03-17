<?php

namespace App\Models;

//use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;

class HospitalBedProfiles extends Model
{
    use SoftDeletes;

    protected $table = 'tbl_hospital_bed_profiles';
}
