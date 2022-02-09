<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class AddCareProfileColumnAndRemoveFapColumnFromPlannedIndicatorsTable extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::table('tbl_planned_indicators', function (Blueprint $table) {
            $table->foreignId('care_profile_id')->nullable();
            $table->foreign('care_profile_id')->references('id')->on('tbl_care_profiles')->onUpdate('cascade')->onDelete('restrict');
            $table->foreignId('vmp_group_id')->nullable();
            $table->foreign('vmp_group_id')->references('id')->on('tbl_vmp_groups')->onUpdate('cascade')->onDelete('restrict');
            $table->foreignId('vmp_type_id')->nullable();
            $table->foreign('vmp_type_id')->references('id')->on('tbl_vmp_types')->onUpdate('cascade')->onDelete('restrict');
            $table->dropColumn(['fap_id']);
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::table('tbl_planned_indicators', function (Blueprint $table) {
            $table->dropColumn(['care_profile_id','vmp_group_id','vmp_type_id']);
            $table->foreignId('fap_id')->nullable();
            $table->foreign('fap_id')->references('id')->on('tbl_mo_departments')->onUpdate('cascade')->onDelete('restrict');
        });
    }
}
