<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

return new class extends Migration
{
    /**
     * Run the migrations.
     */
    public function up(): void
    {
        Schema::create('projection_toll_roads', function (Blueprint $table) {
            $table->id();
            $table->string('year', 255)->nullable(); // NOT NULL
            $table->string('quartal', 255)->nullable();
            $table->string('month', 255)->nullable();

            $table->string('attribute', 255)->nullable(); // NOT NULL


            $table->string('value')->nullable();

            $table->boolean('is_show')->default(false);
            $table->integer('col_position');
            $table->integer('row_position');
            $table->string('cell_position', 255);

            $table->unsignedBigInteger('realisation_projection_organisation_id')->nullable();
            $table->unsignedBigInteger('realisation_projection_update_id')->nullable();
            $table->foreign('realisation_projection_update_id')
                ->references('id')
                ->on('realisation_projection_updates')                  // ✅ correct table name
                ->onDelete('cascade');



            $table->string('realisation_projection_realisation_id')->nullable();



            $table->boolean('is_fy')->default(false);

            $table->timestamps();
        });
    }

    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::dropIfExists('projection_toll_roads');
    }
};
