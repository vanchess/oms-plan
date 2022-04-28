<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class RemoveCareProfileIdColoumnFromCareProfilesFomsTable extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::table('tbl_care_profiles_foms', function (Blueprint $table) {
            $table->dropColumn(['care_profile_id']);
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::table('tbl_care_profiles_foms', function (Blueprint $table) {
            $table->foreignId('care_profile_id')->nullable();
            $table->foreign('care_profile_id')->references('id')->on('tbl_care_profiles')->onUpdate('cascade')->onDelete('restrict');
        });
    }
}
