<?php

namespace App\Models;

//use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;

class MedicalInstitution extends Model
{
    use SoftDeletes;

    protected $dates = ['deleted_at'];
    protected $table = 'v_mo';
}
