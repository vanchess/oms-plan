<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class AddMoDepartmentColoumnToIndicatorCangesTable extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::table('tbl_indicators_changes', function (Blueprint $table) {
            $table->foreignId('mo_department_id')->nullable();
            $table->foreign('mo_department_id')->references('id')->on('tbl_mo_departments')->onUpdate('cascade')->onDelete('restrict');
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::table('tbl_indicators_changes', function (Blueprint $table) {
            $table->dropColumn(['mo_department_id']);
        });
    }
}
