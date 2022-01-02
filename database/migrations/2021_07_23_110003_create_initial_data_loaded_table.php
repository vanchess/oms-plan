<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class CreateInitialDataLoadedTable extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::create('tbl_initial_data_loaded', function (Blueprint $table) {
            $table->id();
            $table->year('year');
            $table->foreignId('root_category_id');
            $table->foreignId('user_id');
            $table->foreignId('commit_id');
            $table->timestamps();
            $table->softDeletes();
            
            $table->foreign('root_category_id')->references('id')->on('tbl_category_tree')->onUpdate('cascade')->onDelete('restrict');
            $table->foreign('user_id')->references('id')->on('users')->onUpdate('cascade')->onDelete('restrict');
            $table->foreign('commit_id')->references('id')->on('tbl_committed_changes')->onUpdate('cascade')->onDelete('restrict');
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::dropIfExists('tbl_initial_data_loaded');
    }
}