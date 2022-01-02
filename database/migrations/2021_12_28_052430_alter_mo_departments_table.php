<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class AlterMoDepartmentsTable extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::table('tbl_mo_departments', function (Blueprint $table) {
            $table->string('mz_oid')->nullable();
            $table->foreignId('type_id');
            $table->foreign('type_id')->references('id')->on('tbl_mo_department_types')->onUpdate('cascade')->onDelete('restrict');
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        
        Schema::table('tbl_mo_departments', function (Blueprint $table) {
            $table->dropColumn(['mz_oid', 'type_id']);
        });
    }
}
