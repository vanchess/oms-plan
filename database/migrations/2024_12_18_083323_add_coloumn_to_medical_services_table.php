<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class AddColoumnToMedicalServicesTable extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::table('tbl_medical_services', function (Blueprint $table) {
            $table->string('short_name', 255)->nullable();
            $table->integer('order')->default(0);
            $table->string('slug')->nullable();
            $table->boolean('allocateVolumes')->default(false)->comment('отображаются как объемы в приложении 1 к договору');
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::table('tbl_medical_services', function (Blueprint $table) {
            $table->dropColumn(['short_name','order','slug','allocateVolumes']);
        });
    }
}
