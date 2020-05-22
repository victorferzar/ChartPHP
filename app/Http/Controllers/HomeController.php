<?php


namespace App\Http\Controllers;
//
//no, no
//

use DateTime;
use Illuminate\Http\Request;
use App;


class HomeController extends Controller
{
    private $servicio;

    public function __construct(App\Services\PptHelper $servicio)
    {
        $this->servicio = $servicio;
    }

    public function index()
    {
        return view('welcome');
    }

    public function generateppt(Request $request)
    {
        //return $request->all();

        $vDesde = $request->input("fecha_des");
        $vHasta = $request->input("fecha_has");


        $this->servicio->generar($vDesde, $vHasta);

    }

}
