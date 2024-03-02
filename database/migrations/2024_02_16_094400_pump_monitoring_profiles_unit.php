<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class PumpMonitoringProfilesUnit extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::create('tbl_pump_monitoring_profiles_unit', function (Blueprint $table) {
            $table->id();
            $table->foreignId('monitoring_profile_id');
            $table->foreignId('unit_id');
            $table->timestamps();
            $table->softDeletes();
            $table->foreign('unit_id')->references('id')->on('tbl_pump_unit')->onUpdate('cascade')->onDelete('restrict');
            $table->foreign('monitoring_profile_id')->references('id')->on('tbl_pump_monitoring_profiles')->onUpdate('cascade')->onDelete('restrict');
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::dropIfExists('tbl_pump_monitoring_profiles_unit');
    }
}
