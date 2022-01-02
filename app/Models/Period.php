<?php

namespace App\Models;

// use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;

class Period extends Model
{
    use SoftDeletes;
    
    protected $dates = ['deleted_at'];
    protected $table = 'tbl_period';
    
    protected $fillable = [
        'from',
        'to',
    ];
    
    protected $casts = [
        'from' => 'datetime',
        'to' => 'datetime',
    ];
}
