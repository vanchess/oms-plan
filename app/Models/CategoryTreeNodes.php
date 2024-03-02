<?php

namespace App\Models;

// use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;
use Illuminate\Support\Facades\DB;

class CategoryTreeNodes extends Model
{
    use SoftDeletes;

    protected $table = 'tbl_category_tree';

    public function nodePath() : string
    {
        return DB::select('SELECT category_node_path(?) AS p', [$this->id])[0]?->p;
    }
}
