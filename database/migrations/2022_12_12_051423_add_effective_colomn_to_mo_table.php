<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class AddEffectiveColomnToMoTable extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::table('tbl_organization', function (Blueprint $table) {
            $table->timestamp('effective_from')->useCurrent();
            $table->timestamp('effective_to')->default('9999-12-31 23:59:59');
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::table('tbl_organization', function (Blueprint $table) {
            $table->dropColumn(['effective_from',  'effective_to']);
        });
    }
}
