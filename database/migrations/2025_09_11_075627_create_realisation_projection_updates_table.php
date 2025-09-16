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
        Schema::create('realisation_projection_updates', function (Blueprint $table) {
            $table->id();

            // FK to parent table
            // $table->foreignId('realisation_projection_id')
            //     ->constrained('realisation_and_projections')
            //     ->cascadeOnDelete();

            $table->foreign('realisation_projection_id')
                ->references('id')
                ->on('realisation_and_projections')                  // ✅ correct table name
                ->onDelete('cascade');

            $table->unsignedBigInteger('realisation_projection_id');
            $table->string('document_type')->nullable();
            $table->string('document_version')->nullable();
            $table->string('reference_document')->nullable();
            $table->date('date')->nullable();
            $table->text('notes')->nullable();
            $table->string('file')->nullable();
            $table->timestamps();
        });
    }

    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::dropIfExists('realisation_projection_updates');
    }
};
