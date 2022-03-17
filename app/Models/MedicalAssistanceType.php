<?php

namespace App\Models;

//use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;

class MedicalAssistanceType extends Model
{
    use SoftDeletes;

    protected $table = 'tbl_medical_assistance_types';
}
