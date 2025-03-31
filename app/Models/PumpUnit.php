<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\Relations\BelongsTo;
use Illuminate\Database\Eloquent\SoftDeletes;

class PumpUnit extends Model
{
    use SoftDeletes;

    protected $table = 'tbl_pump_unit';

    protected $fillable = [
        'name',
        'type_id'
    ];
/*
    public function unitType(): BelongsTo
    {
        return $this->belongsTo(IndicatorType::class, 'type_id', 'id');
    }
        */
}
