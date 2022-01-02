<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class CreateMoView extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        DB::statement('CREATE VIEW v_mo AS SELECT 
                        mo.id, 
                        mo.code, 
                        org."name", 
                        org.short_name, 
                        org.description, 
                        org.address, 
                        org.inn, 
                        org.ogrn, 
                        org.kpp, 
                        org.email, 
                        org.phone, 
                        org.chief, 
                        mo.license, 
                        mo."order", 
                        mo.organization_id, 
                        mo.inclusion_in_register, 
                        mo.created_at, 
                        mo.updated_at, 
                        mo.deleted_at 
                        from tbl_mo mo 
                        left join tbl_organization org on mo.organization_id = org.id');
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        DB::statement("DROP VIEW IF EXISTS v_mo");
    }
}
