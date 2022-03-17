<?php

namespace App\Models;

// use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;

class PlannedIndicatorChange extends Model
{
    use SoftDeletes;

    public $table = 'tbl_indicators_changes';

}
