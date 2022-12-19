<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class InclusionInRegisterNulableMoTable extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::table('tbl_mo', function (Blueprint $table) {
            $table->timestamp('inclusion_in_register')->nullable()->change();
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::table('tbl_mo', function (Blueprint $table) {
            $table->timestamp('inclusion_in_register')->nullable(false)->change();
        });
    }
}
