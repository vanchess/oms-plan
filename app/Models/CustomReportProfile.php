<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;

class CustomReportProfile extends Model
{
    use SoftDeletes;

    protected $table = 'tbl_custom_report_profile';
    protected $fillable = [
        'custom_report_id', 'parent_id', 'code', 'name', 'short_name',
        'relation_type_id', 'effective_from', 'effective_to', 'order', 'user_id'
    ];

    protected $casts = [
        'effective_from' => 'datetime',
        'effective_to' => 'datetime',
    ];

    public function report()
    {
        return $this->belongsTo(CustomReport::class, 'custom_report_id');
    }

    public function relationType()
    {
        return $this->belongsTo(CustomReportProfileRelationType::class, 'relation_type_id');
    }

    public function units()
    {
        return $this->belongsToMany(
            CustomReportProfileUnit::class,
            'tbl_custom_report_profile_unit_planned_indicator',
            'planned_indicator_id',
            'profile_unit_id'
        );
    }

    public function parent()
    {
        return $this->belongsTo(CustomReportProfile::class, 'parent_id');
    }

    public function children()
    {
        return $this->hasMany(CustomReportProfile::class, 'parent_id');
    }
}
