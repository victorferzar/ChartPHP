<?php


namespace App\Http\Controllers;
header('Content-type: application/vnd.openxmlformats-officedocument.presentationml.presentation');


use DateTime;
use Illuminate\Http\Request;

use Illuminate\Support\Facades\Auth;
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


class GeneratePPT extends Controller
{

    public const Datos_Filtro_Blancos_AssayName = array(
        1 => "CuT_CMCCAAS_pct",
        2 => "CuT_pct_2A15tAA",
        3 => "CuS_CMCCAAS_pct",
        4 => "CuS_pct_SCA2pAA",
        5 => "CuCN_pct_CYN6pAA",
        6 => "CuFe_pct_FES7pAA"
    );

    public const Datos_Filtro_Blancos_StandardId = array(
        1 => "BF42",
        2 => "BG4"
    );

    public const Datos_Tabla = array(
        1 => "# of Analyses Above Threshold",
        2 => "# of Outside Warning Limit",
        3 => "# of Outside Error Limit",
        4 => "# of Analyses Bellow Threshold",
        5 => "% Outside Error Limit",
        6 => "Mean",
        7 => "Median",
        8 => "Min",
        9 => "Max",
        10 => "Standard Deviation",
        11 => "Rel. Std. Dev",
        12 => "Standard Error",
        13 => "Rel. Std. Err",
        14 => "Total Bias",
        15 => "% Mean Bias"
    );

    public function index()
    {
        return view('welcome');
    }

    public function filtro($standardID, $assayName, $vDesde, $vHasta)
    {
        $datosBD = App\Dtm_qaqc_blk_std::selectRaw("TRY_CONVERT(date, RETURNDATE,103) as FECHAS, SUM(ASSAYVALUE) as SUMA")
            ->where("STANDARDID", "=", "$standardID")
            ->where("ASSAYNAME", "=", "$assayName")
            ->whereRaw("TRY_CONVERT(date, RETURNDATE, 103) Between ('$vDesde') AND ('$vHasta')")
            ->groupByRaw("TRY_CONVERT(date, RETURNDATE, 103)")
            //->orderByRaw('TRY_CONVERT(date, RETURNDATE, 103) DESC')
            ->get();

        return $datosBD;
    }

    public function generateppt(Request $request)
    {
        //return $request->all();
        //RECIBE RANGO DE FECHAS DEL FORMULARIO
        $vDesde = $request->input("fecha_des");
        $vHasta = $request->input("fecha_has");

        /**
         * se crea una nueva instancia de PowerPoint
         */
        $objPPT = new PhpPresentation();
        $objPPT->getDocumentProperties()
            ->setCreator('Austem');

        //DIBUJAR GRAFICO BLANCOS
        $this->blankDraw($objPPT, $vDesde, $vHasta);


        //GUARDAR EN EL EQUIPO
        $oWriterPPTX = IOFactory::createWriter($objPPT, 'PowerPoint2007');
        $rutaArchivo = __DIR__ . "/sample.pptx";
        $oWriterPPTX->save($rutaArchivo);
        readfile($rutaArchivo);
        exit;
    }

    public function blankDraw(PhpPresentation $objPPT, $vDesde, $vHasta)
    {
        foreach (self::Datos_Filtro_Blancos_StandardId as $vId) {
            foreach (self::Datos_Filtro_Blancos_AssayName as $vAssay) {

                $currentSlide = $objPPT->createSlide();

                //FILTRO BD
                $dataBD = $this->filtro($vId, $vAssay, $vDesde, $vHasta);
                $seriesData = [];

                //DATOS DEL GRAFICO
                foreach ($dataBD as $num) {

                    if (count($seriesData) == 0) {
                        $desdeReal = date('d-m-y', strtotime($num->FECHAS));
                    }
                    $aux = date('d-m', strtotime($num->FECHAS));
                    $seriesData[$aux] = floatval($num->SUMA);
                    $hastaReal = date('d-m-y', strtotime($num->FECHAS));

                }

                //DATOS DE LA TABLA
                $dataA = self::Datos_Tabla;

                //TABLA
                $tableShape = $currentSlide->createTableShape(2);
                $tableShape->setResizeProportional(false)
                    // ->setHeight(200)
                    ->setWidth(330)
                    ->setOffsetX(730)
                    ->setOffsetY(200);

                $row0 = $tableShape->createRow()
                    ->setHeight(25);

                $cell00 = $row0->nextCell();
                $cell00->CreateTextRun("STATISTICS");
                $cell00->getActiveParagraph()
                    ->getAlignment()
                    ->setHorizontal(Alignment::HORIZONTAL_CENTER)
                    ->setVertical(Alignment::VERTICAL_CENTER);

                $cell01 = $row0->getCell(1);
                $cell01->createTextRun("d " . strval($dataBD->count()));
                $cell01->setWidth(60);
                $cell01->getActiveParagraph()
                    ->getAlignment()
                    ->setHorizontal(Alignment::HORIZONTAL_CENTER)
                    ->setVertical(Alignment::VERTICAL_CENTER);

                for ($i = 1; $i <= 15; $i++) {
                    $row[$i] = $tableShape->createRow()
                        ->setHeight(32);
                    $cellAux = $row[$i]->getCell(0);
                    $cellAux->createTextRun($dataA[$i]);
                    $cellAux->getActiveParagraph()
                        ->getAlignment()
                        ->setVertical(Alignment::VERTICAL_CENTER)
                        ->setMarginLeft(2);
                }

                //TITULO
                if ($vId == "BF42") {
                    $vTipo = "Fino";
                } elseif ($vId == "BG4") {
                    $vTipo = "Grueso";
                }
                //LOGO
                $logoImg = $currentSlide->createDrawingShape();
                $logoImg->setName('logo')
                    ->setDescription('Logotipo Corporativa');
                $logoImg->setPath(resource_path() . '/images/logo-bhp.png');
                $logoImg->setOffsetX(60)
                    ->setOffsetY(60);
                $logoImg->setResizeProportional(true)
                    ->setHeight(60);

                //TEXTO
                $txtTitulo = $currentSlide->createRichTextShape()
                    ->setHeight(100)
                    ->setWidth(700)
                    ->setOffsetX(200)
                    ->setOffsetY(30);

                $txtTitulo->getActiveParagraph()
                    ->getAlignment()
                    ->setHorizontal(Alignment::HORIZONTAL_CENTER);

                $txtSubTitulo = $txtTitulo;
                $txtDetalle = $txtTitulo;


                $txtRunT = $txtTitulo->createTextRun("Blancos - " . $vAssay . "\n");
                $txtRunST = $txtSubTitulo->createTextRun("Blancos Aprobados " . $vId . " - " . "Blanco " . $vTipo . "\n");
                $txtRunD = $txtDetalle->createTextRun("Estandares desde: " . $desdeReal . " Hasta: " . $hastaReal . " STD: " . $vId . " Elementos: " . $vAssay);
                $txtRunT->getFont()->setBold(true)
                    ->setSize(35)
                    ->setColor(new Color('FFE06B20'));
                $txtRunST->getFont()->setItalic(true)
                    ->setSize(25);
                $txtRunD->getFont()->setSize(12);


                $oOutLine = new Outline();
                $oOutLine->setWidth(1);
                $oOutLine->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_BLUE));

                //GRAFICO DE BARRAS
                //BARRA
                $barChart = new Bar();
                // $barChart->setGapWidthPercent(158);
                $series1 = new Series('VALORES', $seriesData);
                $series1->setShowSeriesName(false);
                $series1->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFA500'));
                $series1->getFont()->getColor()->setRGB('00FF00');
                $series1->setShowValue(false);
                $series1->setOutline($oOutLine);
                $barChart->addSeries($series1);

                //LINEAS
                $errLine = new Line();
                $warLine = new Line();


                //GRAFICO
                $oGrid = new Gridlines();
                $oGrid->getOutline()->setWidth(1);

                $oGrid->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_BLUE));


                $chartShape = $currentSlide->createChartShape();
                $chartShape->setName("Grafico de Blancos")->setResizeProportional(false)
                    ->setHeight(400)
                    ->setWidth(700)
                    ->setOffsetX(10)
                    ->setOffsetY(200);
                $chartShape->getBorder()->setLineStyle(Border::LINE_SINGLE);
                $chartShape->getTitle()->setText($vId . " - " . $vAssay);
                $chartShape->getTitle()->getFont()->setItalic(true);
                $chartShape->getTitle()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
                $chartShape->getPlotArea()->getAxisX()->setTitle('Fecha de Retorno');
                $chartShape->getPlotArea()->getAxisY()->setTitle('Ley Laboratorio');
                $chartShape->getPlotArea()->getAxisY()->setMajorGridlines($oGrid);
                //  $chartShape->getPlotArea()->getAxisY()->setFormatCode('#,##0');


                //SALTOS ENTRE VALORES EN Y
                //$chartShape->getPlotArea()->getAxisY()->setMajorUnit(0.002);
                //VALOR MAXIMO PARA Y
                //$chartShape->getPlotArea()->getAxisY()->setMaxBounds(0.1);
                //$chartShape->getPlotArea()->getAxisY()->setMinorTickMark(Axis::TICK_MARK_CROSS);

                $chartShape->getPlotArea()->setType($barChart);
                $chartShape->getLegend()->setVisible(false);
            }
        }

    }

}
