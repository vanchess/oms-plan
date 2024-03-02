<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class PumpMonitoringProfilesUnitPlannedIndicators extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::create('tbl_pump_monitoring_profiles_unit_planned_indicators', function (Blueprint $table) {
            $table->id();
            $table->foreignId('monitoring_profile_unit_id');
            $table->foreignId('planned_indicator_id');
            $table->timestamps();
            $table->softDeletes();
            $table->foreign('monitoring_profile_unit_id')->references('id')->on('tbl_pump_monitoring_profiles_unit')->onUpdate('cascade')->onDelete('restrict');
            $table->foreign('planned_indicator_id')->references('id')->on('tbl_planned_indicators')->onUpdate('cascade')->onDelete('restrict');
            $table->unique(['monitoring_profile_unit_id', 'planned_indicator_id'], 'pump_monitoring_profiles_unit_planned_indicators_uk');
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::dropIfExists('tbl_pump_monitoring_profiles_unit_planned_indicators');
    }
}
