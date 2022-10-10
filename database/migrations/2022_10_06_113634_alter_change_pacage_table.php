<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class AlterChangePacageTable extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::table('tbl_change_package', function (Blueprint $table) {
            $table->timestamp('completed_at')->nullable();
            $table->foreignId('commission_decision_id')->nullable();
            $table->foreign('commission_decision_id')->references('id')->on('tbl_commission_decisions')->onUpdate('cascade')->onDelete('restrict');
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::table('tbl_change_package', function (Blueprint $table) {
            $table->dropColumn(['completed_at',  'commission_decision_id']);
        });
    }
}
