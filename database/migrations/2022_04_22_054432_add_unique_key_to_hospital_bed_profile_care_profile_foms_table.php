<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class AddUniqueKeyToHospitalBedProfileCareProfileFomsTable extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::table('tbl_hospital_bed_profile_care_profile_foms', function (Blueprint $table) {
            $table->unique(['hospital_bed_profile_id', 'care_profile_foms_id'], 'bed_profile_care_profile_uk');
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::table('tbl_hospital_bed_profile_care_profile_foms', function (Blueprint $table) {
            $table->dropUnique('bed_profile_care_profile_uk');
        });
    }
}
