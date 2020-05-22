<?php

use Illuminate\Support\Facades\Route;

/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider within a group which
| contains the "web" middleware group. Now create something great!
|
*/

/*Route::get('/', function () {
    return view('welcome');
});*/

Route::get('/', 'HomeController@index');

Route::post('/generateppt', 'HomeController@generateppt')->name('generar');

//Ruta que crea la imagen
Route::get('/crearGrafico', 'ChartController')->name("chart");

//Muestra del grafico desde el menu
Route::get('/mostrarGrafico', function () {


    return view('grafico', compact("series"));
})->name('mostrar');


//funcion de shot
Route::get('/generar_jpg', 'ChartController')->name('jpg');
