<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class CreatePlannedIndicatorsTable extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::create('tbl_planned_indicators', function (Blueprint $table) {
            $table->id();
            $table->foreignId('node_id');
            $table->foreignId('indicator_id');
            $table->foreignId('service_id')->nullable();
            $table->foreignId('profile_id')->nullable();
            $table->foreignId('fap_id')->nullable();
            $table->timestamps();
            $table->softDeletes();
            
            $table->foreign('node_id')->references('id')->on('tbl_category_tree')->onUpdate('cascade')->onDelete('cascade');
            $table->foreign('indicator_id')->references('id')->on('tbl_indicators')->onUpdate('cascade')->onDelete('restrict');
            $table->foreign('service_id')->references('id')->on('tbl_medical_services')->onUpdate('cascade')->onDelete('restrict');
            $table->foreign('profile_id')->references('id')->on('tbl_hospital_bed_profiles')->onUpdate('cascade')->onDelete('restrict');
            $table->foreign('fap_id')->references('id')->on('tbl_fap')->onUpdate('cascade')->onDelete('restrict');
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::dropIfExists('tbl_planned_indicators');
    }
}
