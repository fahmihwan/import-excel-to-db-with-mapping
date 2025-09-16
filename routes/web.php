<?php

use App\Http\Controllers\TableDinamisController;
use Illuminate\Support\Facades\Route;

Route::get('/', [TableDinamisController::class, 'home']);
Route::post('/data', [TableDinamisController::class, 'store_import']);
