<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class CreateCareProfileCareProfileFomsTable extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::create('tbl_care_profile_care_profile_foms', function (Blueprint $table) {
            $table->id();
            $table->foreignId('care_profile_id');
            $table->foreign('care_profile_id')->references('id')->on('tbl_care_profiles')->onUpdate('cascade')->onDelete('restrict');
            $table->foreignId('care_profile_foms_id');
            $table->foreign('care_profile_foms_id')->references('id')->on('tbl_care_profiles_foms')->onUpdate('cascade')->onDelete('restrict');
            $table->timestamps();
            $table->softDeletes();
            $table->unique(['care_profile_id', 'care_profile_foms_id'], 'care_profile_care_profile_foms_uk');
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::dropIfExists('tbl_care_profile_care_profile_foms');
    }
}
