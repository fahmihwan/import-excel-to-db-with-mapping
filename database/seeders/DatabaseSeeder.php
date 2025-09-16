<?php

namespace Database\Seeders;

use App\Models\RealisationAndProjection;
use App\Models\RealisationProjectionUpdate;
use App\Models\User;
use Carbon\Carbon;
// use Illuminate\Database\Console\Seeds\WithoutModelEvents;
use Illuminate\Database\Seeder;

class DatabaseSeeder extends Seeder
{
    /**
     * Seed the application's database.
     */
    public function run(): void
    {
        // User::factory(10)->create();

        User::factory()->create([
            'name' => 'Test User',
            'email' => 'test@example.com',
        ]);

        RealisationAndProjection::create([
            'debtor_id'        => 10,
            'tol_road_section' => 'Section A',
            'due_date'         => Carbon::parse('2025-12-31'),
            'due_year'         => 2025,
            'cod_date'         => Carbon::parse('2026-06-30'),
        ]);

        RealisationProjectionUpdate::create([
            'realisation_projection_id' => 1, // must exist in realisation_and_projections
            'document_type'      => 'Feasibility Report',
            'document_version'   => 'v1.0',
            'reference_document' => 'FR-2025-001',
            'date'               => Carbon::parse('2025-01-10'),
            'notes'              => 'Initial feasibility study document.',
            'file'               => 'feasibility_report_v1.pdf',
        ]);
    }
}
