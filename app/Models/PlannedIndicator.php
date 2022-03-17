<?php

namespace App\Models;

// use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;

class PlannedIndicator extends Model
{
    use SoftDeletes;

    protected $table = 'tbl_planned_indicators';

    protected $casts = [
        'indicator_id' => 'integer',
    ];

    public function initialData()
    {
        return $this->hasMany(InitialData::class, 'planned_indicator_id', 'id');
    }
}
