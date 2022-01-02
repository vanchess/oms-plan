<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class PlannedIndicator extends Model
{
    use HasFactory;
    
    protected $dates = ['deleted_at'];
    protected $table = 'tbl_planned_indicators';
    
    public function initialData()
    {
        return $this->hasMany(InitialData::class, 'planned_indicator_id', 'id');
    }
}
