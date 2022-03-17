<?php

namespace App\Models;

//use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;

class InitialData extends Model
{
    use SoftDeletes;

    protected $table = 'tbl_initial_data';

    public function plannedIndicator() {
        return $this->belongsTo(PlannedIndicator::class, 'planned_indicator_id', 'id');
    }
}
