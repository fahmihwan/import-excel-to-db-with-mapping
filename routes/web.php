<?php

use App\Http\Controllers\TableDinamisController;
use Illuminate\Support\Facades\Route;

Route::get('/', [TableDinamisController::class, 'toll_roads']); //roads
Route::get('/non-plts', [TableDinamisController::class, 'non_plts']); //roads
Route::post('/data', [TableDinamisController::class, 'store_import']);
