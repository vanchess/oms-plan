<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;

class CustomReportProfileRelationType extends Model
{
    use SoftDeletes;

    protected $table = 'tbl_custom_report_profile_relation_type';

    protected $fillable = [
        'name',
        'slug',
    ];


}
