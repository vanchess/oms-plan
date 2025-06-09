<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class CustomReportProfile extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::create('tbl_custom_report_profile', function (Blueprint $table) {
            $table->id();
            $table->foreignId('custom_report_id');
            $table->foreignId('parent_id')->nullable();
            $table->string('code',64)->unique();
            $table->string('name',1024);
            $table->string('short_name',256);
            $table->foreignId('relation_type_id')->nullable();
            $table->timestamp('effective_from')->useCurrent();
            $table->timestamp('effective_to')->default('9999-12-31 23:59:59');
            $table->integer('order')->nullable();
            $table->foreignId('user_id');
            $table->timestamps();
            $table->softDeletes();
            $table->foreign('custom_report_id')->references('id')->on('tbl_custom_report')->onUpdate('cascade')->onDelete('restrict');
            $table->foreign('parent_id')->references('id')->on('tbl_custom_report_profile')->onUpdate('cascade')->onDelete('restrict');
            $table->foreign('relation_type_id')->references('id')->on('tbl_custom_report_profile_relation_type')->onUpdate('cascade')->onDelete('restrict');
            $table->foreign('user_id')->references('id')->on('users')->onUpdate('cascade')->onDelete('restrict');
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::dropIfExists('tbl_custom_report_profile');
    }
}
