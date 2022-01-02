<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class CreateIndicatorsChangesTable extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::create('tbl_indicators_changes', function (Blueprint $table) {
            $table->id();
            $table->foreignId('period_id');
            $table->foreignId('mo_id');
            $table->foreignId('planned_indicator_id');
            $table->foreignId('commit_id')->nullable();
            $table->foreignId('user_id');
            $table->decimal('value', 19, 4);
            $table->timestamps();
            $table->softDeletes();
            
            $table->foreign('period_id')->references('id')->on('tbl_period')->onUpdate('cascade')->onDelete('restrict');
            $table->foreign('mo_id')->references('id')->on('tbl_mo')->onUpdate('cascade')->onDelete('restrict');
            $table->foreign('planned_indicator_id')->references('id')->on('tbl_planned_indicators')->onUpdate('cascade')->onDelete('restrict');
            $table->foreign('commit_id')->references('id')->on('tbl_committed_changes')->onUpdate('cascade')->onDelete('restrict');
            $table->foreign('user_id')->references('id')->on('users')->onUpdate('cascade')->onDelete('restrict');
            
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::dropIfExists('tbl_indicators_changes');
    }
}
