<?php

namespace App\Models;

// use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;

class ChangePackage extends Model
{
    use SoftDeletes;

    protected $table ='tbl_change_package';

    public function commissionDecision() {
        return $this->belongsTo(CommissionDecision::class, 'commission_decision_id', 'id');
    }
}
