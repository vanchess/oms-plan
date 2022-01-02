<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class AlterHospitalBedProfilesTable extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::table('tbl_hospital_bed_profiles', function (Blueprint $table) {
            $table->integer('code')->nullable();
            $table->string('short_name')->nullable();
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::table('tbl_hospital_bed_profiles', function (Blueprint $table) {
            $table->dropColumn(['code','short_name']);
        });
        
    }
}
