<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;

class CustomReport extends Model
{
    use SoftDeletes;

    protected $table = 'tbl_custom_report';
    protected $fillable = ['name', 'short_name', 'user_id'];

    public function profiles()
    {
        return $this->hasMany(CustomReportProfile::class, 'custom_report_id');
    }

    public function user()
    {
        return $this->belongsTo(User::class);
    }
}
