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
        Schema::create('realisation_and_projections', function (Blueprint $table) {
            $table->id();
            $table->unsignedBigInteger('sector_id')->nullable();
            $table->unsignedBigInteger('debtor_id')->nullable();
            $table->string('tol_road_section')->nullable();
            $table->date('due_date')->nullable();
            $table->year('due_year')->nullable();
            $table->date('cod_date')->nullable();
            $table->timestamps();
        });
    }

    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::dropIfExists('realisation_and_projections');
    }
};
