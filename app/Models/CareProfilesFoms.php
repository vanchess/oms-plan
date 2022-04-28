<?php

namespace App\Models;

// use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;

class CareProfilesFoms extends Model
{
    use SoftDeletes;

    protected $table = 'tbl_care_profiles_foms';
}
