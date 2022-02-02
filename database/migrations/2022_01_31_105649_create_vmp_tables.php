<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class CreateVmpTables extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::create('tbl_care_profiles', function (Blueprint $table) {
            $table->id();
            $table->string('name',512);
            $table->string('mz_code',64)->nullable();
            $table->timestamps();
            $table->softDeletes();
        });

        Schema::create('tbl_care_profiles_foms', function (Blueprint $table) {
            $table->id();
            $table->string('name', 512);
            $table->foreignId('care_profile_id')->nullable();
            $table->foreign('care_profile_id')->references('id')->on('tbl_care_profiles')->onUpdate('cascade')->onDelete('restrict');
            $table->string('code_v002',64);
            $table->timestamp('effective_from')->useCurrent();
            $table->timestamp('effective_to')->default('9999-12-31 23:59:59');
            $table->timestamps();
            $table->softDeletes();
        });

        Schema::create('tbl_vmp_methods', function (Blueprint $table) {
            $table->id();
            $table->string('name', 1024);
            $table->string('mz_code', 64)->nullable();
            $table->timestamps();
            $table->softDeletes();
        });

        Schema::create('tbl_vmp_groups', function (Blueprint $table) {
            $table->id();
            $table->string('code', 64);
            $table->timestamps();
            $table->softDeletes();
        });

        Schema::create('tbl_vmp_types', function (Blueprint $table) {
            $table->id();
            $table->string('name', 512);
            $table->string('mz_code', 64)->nullable();
            $table->timestamps();
            $table->softDeletes();
        });

        Schema::create('tbl_vmp_details', function (Blueprint $table) {
            $table->id();
            $table->foreignId('care_profile_id');
            $table->foreign('care_profile_id')->references('id')->on('tbl_care_profiles')->onUpdate('cascade')->onDelete('restrict');
            $table->foreignId('groupe_id');
            $table->foreign('groupe_id')->references('id')->on('tbl_vmp_groups')->onUpdate('cascade')->onDelete('restrict');
            $table->foreignId('method_id');
            $table->foreign('method_id')->references('id')->on('tbl_vmp_methods')->onUpdate('cascade')->onDelete('restrict');
            $table->foreignId('type_id');
            $table->foreign('type_id')->references('id')->on('tbl_vmp_types')->onUpdate('cascade')->onDelete('restrict');
            $table->timestamp('effective_from')->useCurrent();
            $table->timestamp('effective_to')->default('9999-12-31 23:59:59');
            $table->timestamps();
            $table->softDeletes();
        });


    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::dropIfExists('tbl_vmp_details');
        Schema::dropIfExists('tbl_vmp_types');
        Schema::dropIfExists('tbl_vmp_groups');
        Schema::dropIfExists('tbl_vmp_methods');
        Schema::dropIfExists('tbl_care_profiles_foms');
        Schema::dropIfExists('tbl_care_profiles');
    }
}
