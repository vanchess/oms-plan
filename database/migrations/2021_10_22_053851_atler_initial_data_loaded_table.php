<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class AtlerInitialDataLoadedTable extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::table('tbl_initial_data_loaded', function(Blueprint $table) {
            $table->renameColumn('root_category_id', 'node_id');
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::table('tbl_initial_data_loaded', function(Blueprint $table) {
            $table->renameColumn('node_id', 'root_category_id');
        });
    }
}
