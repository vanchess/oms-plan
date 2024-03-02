<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class PumpMonitoringProfiles extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::create('tbl_pump_monitoring_profiles', function (Blueprint $table) {
            $table->id();
            $table->foreignId('oms_program_id');
            $table->foreignId('parent_id')->nullable();
            $table->string('code',16)->unique();
            $table->string('name',1024);
            $table->string('short_name',256);
            $table->foreignId('relation_type_id')->nullable();
            $table->timestamp('effective_from')->useCurrent();
            $table->timestamp('effective_to')->default('9999-12-31 23:59:59');
            $table->boolean('is_leaf');
            $table->timestamps();
            $table->softDeletes();
            $table->foreign('oms_program_id')->references('id')->on('tbl_oms_program')->onUpdate('cascade')->onDelete('restrict');
            $table->foreign('parent_id')->references('id')->on('tbl_pump_monitoring_profiles')->onUpdate('cascade')->onDelete('restrict');
            $table->foreign('relation_type_id')->references('id')->on('tbl_pump_monitoring_profiles_relation_type')->onUpdate('cascade')->onDelete('restrict');
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::dropIfExists('tbl_pump_monitoring_profiles');
    }
}
