<?php

namespace App\Models;

// use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;

class CommissionDecision extends Model
{
    //use HasFactory;
    use SoftDeletes;

    protected $table = 'tbl_commission_decisions';

    protected $casts = [
        'date' => 'date',
    ];

    public function changePackage()
    {
        return $this->hasMany(ChangePackage::class, 'commission_decision_id', 'id');
    }
}
