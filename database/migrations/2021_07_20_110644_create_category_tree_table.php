<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class CreateCategoryTreeTable extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::create('tbl_category_tree', function (Blueprint $table) {
            $table->id();
            $table->foreignId('category_id')
                   ->constrained('tbl_categories')
                   ->onUpdate('cascade')
                   ->onDelete('restrict');
            $table->foreignId('parent_id')->nullable();
            $table->boolean('is_leaf')->default(false);
            $table->timestamps();
            $table->softDeletes();
            
            $table->foreign('parent_id')->references('id')->on('tbl_category_tree')->onUpdate('cascade')->onDelete('cascade');
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::dropIfExists('tbl_category_tree');
    }
}
