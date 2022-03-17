<?php

namespace App\Models;

// use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;

class VmpTypes extends Model
{
    use SoftDeletes;

    protected $table = 'tbl_vmp_types';
}
