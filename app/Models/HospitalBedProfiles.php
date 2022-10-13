<?php

namespace App\Models;

//use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;

class HospitalBedProfiles extends Model
{
    use SoftDeletes;

    protected $table = 'tbl_hospital_bed_profiles';

    public function careProfilesFoms() {
        return $this->belongsToMany(CareProfilesFoms::class, 'tbl_hospital_bed_profile_care_profile_foms', 'hospital_bed_profile_id', 'care_profile_foms_id');
    }
}
