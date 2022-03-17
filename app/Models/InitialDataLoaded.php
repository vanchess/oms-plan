<?php

namespace App\Models;

// use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;

class InitialDataLoaded extends Model
{
    use SoftDeletes;

    protected $table = 'tbl_initial_data_loaded';

    protected $fillable = [
        'year',
        'node_id',
        'user_id',
        'commit_id',
        'created_at',
        'updated_at',
    ];
}
