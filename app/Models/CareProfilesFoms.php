<?php

namespace App\Models;

// use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;

class CareProfilesFoms extends Model
{
    use SoftDeletes;

    protected $table = 'tbl_care_profiles_foms';

    public function hospitalBedProfiles() {
        return $this->belongsToMany(HospitalBedProfiles::class, 'tbl_hospital_bed_profile_care_profile_foms', 'care_profile_foms_id', 'hospital_bed_profile_id');
    }

    public function careProfilesMz() {
        return $this->belongsToMany(CareProfiles::class, 'tbl_care_profile_care_profile_foms', 'care_profile_foms_id', 'care_profile_id');
    }
}
