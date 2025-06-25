<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;

class CustomReportUnit extends Model
{
    use SoftDeletes;

    protected $table = 'tbl_custom_report_unit';
    protected $fillable = ['name', 'type_id', 'order'];

    public function type()
    {
        return $this->belongsTo(IndicatorType::class, 'type_id');
    }
}
