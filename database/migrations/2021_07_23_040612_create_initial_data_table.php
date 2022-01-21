<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class CreateInitialDataTable extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::create('tbl_algorithms', function (Blueprint $table) {
            $table->id();
            $table->string('name',128)->unique();;
            $table->string('description')->nullable();
            $table->timestamp('effective_from')->useCurrent();
            $table->timestamp('effective_to')->default('9999-12-31 23:59:59');
            $table->timestamps();
            $table->softDeletes();
        });

        Schema::create('tbl_initial_data', function (Blueprint $table) {
            $table->id();
            $table->year('year');
            $table->foreignId('mo_id')
                   ->constrained('tbl_mo')
                   ->onUpdate('cascade')
                   ->onDelete('cascade');
            $table->foreignId('planned_indicator_id')
                   ->constrained('tbl_planned_indicators')
                   ->onUpdate('cascade')
                   ->onDelete('cascade');
            $table->decimal('value', 19, 4);
            $table->foreignId('algorithm_id')
                   ->constrained('tbl_algorithms')
                   ->onUpdate('cascade')
                   ->onDelete('restrict');
            $table->foreignId('user_id')
                   ->constrained('users')
                   ->onUpdate('cascade')
                   ->onDelete('restrict');
            $table->timestamp('processed_at')->nullable();
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
        Schema::dropIfExists('tbl_initial_data');
        Schema::dropIfExists('tbl_algorithms');
    }
}
