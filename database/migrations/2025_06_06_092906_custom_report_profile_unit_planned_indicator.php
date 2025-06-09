<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class CustomReportProfileUnitPlannedIndicator extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::create('tbl_custom_report_profile_unit_planned_indicator', function (Blueprint $table) {
            $table->id();
            $table->foreignId('profile_unit_id');
            $table->foreignId('planned_indicator_id');
            $table->timestamps();
            $table->softDeletes();
            $table->foreign('profile_unit_id')->references('id')->on('tbl_custom_report_profile_unit')->onUpdate('cascade')->onDelete('restrict');
            $table->foreign('planned_indicator_id')->references('id')->on('tbl_planned_indicators')->onUpdate('cascade')->onDelete('restrict');
            $table->unique(['profile_unit_id', 'planned_indicator_id'], 'custom_report_profile_unit_planned_indicator_uk');
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::dropIfExists('tbl_custom_report_profile_unit_planned_indicator');
    }
}
