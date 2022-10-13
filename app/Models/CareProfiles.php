<?php

namespace App\Models;

//use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;

class CareProfiles extends Model
{
    use SoftDeletes;

    protected $table = 'tbl_care_profiles';

    public function careProfilesFoms() {
        return $this->belongsToMany(CareProfilesFoms::class, 'tbl_care_profile_care_profile_foms', 'care_profile_id', 'care_profile_foms_id');
    }
}
