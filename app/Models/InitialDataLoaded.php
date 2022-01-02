<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class InitialDataLoaded extends Model
{
    use HasFactory;
    
    protected $dates = ['deleted_at'];
    protected $table = 'tbl_initial_data_loaded';
}
