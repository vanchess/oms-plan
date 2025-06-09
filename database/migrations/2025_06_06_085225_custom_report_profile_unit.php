<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class CustomReportProfileUnit extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::create('tbl_custom_report_profile_unit', function (Blueprint $table) {
            $table->id();
            $table->foreignId('profile_id');
            $table->foreignId('unit_id');
            $table->timestamps();
            $table->softDeletes();
            $table->foreign('unit_id')->references('id')->on('tbl_custom_report_unit')->onUpdate('cascade')->onDelete('restrict');
            $table->foreign('profile_id')->references('id')->on('tbl_custom_report_profile')->onUpdate('cascade')->onDelete('restrict');
            $table->unique(['profile_id', 'unit_id'], 'custom_report_profile_unit_uk');
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::dropIfExists('tbl_custom_report_profile_unit');
    }
}
