<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class AlterPlannedIndicatorsTable extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::table('tbl_planned_indicators', function (Blueprint $table) {
            $table->foreignId('assistance_type_id')->nullable();
            $table->foreign('assistance_type_id')->references('id')->on('tbl_medical_assistance_types')->onUpdate('cascade')->onDelete('restrict');
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
            $table->dropColumn(['assistance_type_id']);
        });
    }
}
