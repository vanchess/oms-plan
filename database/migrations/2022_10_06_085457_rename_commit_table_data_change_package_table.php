<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class RenameCommitTableDataChangePackageTable extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {

        Schema::table('tbl_committed_changes', function (Blueprint $table) {
            $table->dropForeign(['user_id']);
         });

        Schema::rename('tbl_committed_changes', 'tbl_change_package');

        Schema::table('tbl_change_package', function (Blueprint $table) {
            $table->foreign('user_id')->references('id')->on('users')->onUpdate('cascade')->onDelete('restrict');
         });


        Schema::table('tbl_indicators_changes', function (Blueprint $table) {
            $table->dropForeign(['commit_id']);
         });
        Schema::table('tbl_indicators_changes', function (Blueprint $table) {
            $table->renameColumn('commit_id', 'package_id');
            $table->foreign('package_id')->references('id')->on('tbl_change_package')->onUpdate('cascade')->onDelete('restrict');
         });

         Schema::table('tbl_initial_data_loaded', function (Blueprint $table) {
            $table->dropForeign(['commit_id']);
         });
        Schema::table('tbl_initial_data_loaded', function (Blueprint $table) {
            $table->renameColumn('commit_id', 'package_id');
            $table->foreign('package_id')->references('id')->on('tbl_change_package')->onUpdate('cascade')->onDelete('restrict');
         });


    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::table('tbl_committed_changes', function (Blueprint $table) {
            $table->dropForeign(['user_id']);
         });

        Schema::rename('tbl_change_package', 'tbl_committed_changes');

        Schema::table('tbl_change_package', function (Blueprint $table) {
            $table->foreign('user_id')->references('id')->on('users')->onUpdate('cascade')->onDelete('restrict');
         });

        Schema::table('tbl_indicators_changes', function (Blueprint $table) {
            $table->dropForeign(['package_id']);
         });
        Schema::table('tbl_indicators_changes', function (Blueprint $table) {
            $table->renameColumn('package_id', 'commit_id');
            $table->foreign('commit_id')->references('id')->on('tbl_committed_changes')->onUpdate('cascade')->onDelete('restrict');
         });

        Schema::table('tbl_initial_data_loaded', function (Blueprint $table) {
            $table->dropForeign(['package_id']);
         });
        Schema::table('tbl_initial_data_loaded', function (Blueprint $table) {
            $table->renameColumn('package_id', 'commit_id');
            $table->foreign('commit_id')->references('id')->on('tbl_committed_changes')->onUpdate('cascade')->onDelete('restrict');
         });


    }
}
