<?php


namespace App\Http\Controllers;
header('Content-type: application/vnd.openxmlformats-officedocument.presentationml.presentation');

use DateTime;
use Illuminate\Http\Request;

use Illuminate\Support\Facades\Auth;
use Illuminate\Support\Facades\DB;
use mysql_xdevapi\Table;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Shape\Chart\Axis;
use PhpOffice\PhpPresentation\Shape\Chart\Gridlines;
use PhpOffice\PhpPresentation\Shape\Table\Cell;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Line;

use App;
use PhpOffice\PhpPresentation\Style\Outline;
use PhpOffice\PhpPresentation\Style\Shadow;


class GeneratePPT extends Controller
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
        //RECIBE RANGO DE FECHAS DEL FORMULARIO
        $vDesde = $request->input("fecha_des");
        $vHasta = $request->input("fecha_has");

        $this->servicio->generar($vDesde, $vHasta);

    }
}
