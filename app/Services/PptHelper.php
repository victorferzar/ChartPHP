<?php


namespace App\Services;

use App\Dtm_standardsassay;
use Illuminate\Support\Facades\App;
use Illuminate\Support\Facades\DB;
use MongoDB\BSON\Symbol;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Chart\Gridlines;
use PhpOffice\PhpPresentation\Shape\Chart\Legend;
use PhpOffice\PhpPresentation\Shape\Chart\Marker;
use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Line;
use PhpOffice\PhpPresentation\Shape\Media;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Outline;
use PhpOffice\PhpPresentation\Shape\Table;
use PhpOffice\PhpPresentation\Shape\Table\Row;
use PhpOffice\PhpPresentation\Shape\Table\Cell;
use PhpOffice\PhpPresentation\Shape\Chart\PlotArea;
use PhpOffice\PhpPresentation\Shape\Chart\Axis;

class PptHelper
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
        4 => "# of Analyses Bellow Threshold (TOTAL DE REGISTROS ENCONTRADOS)",
        5 => "% Outside Error Limit",
        6 => "Mean",
        7 => "Median",
        8 => "Min",
        9 => "Max",
        10 => "Standard Deviation",
        11 => "% Rel. Std. Dev",
        12 => "Standard Error",
        13 => "% Rel. Std. Err",
        14 => "Total Bias",
        15 => "% Mean Bias"
    );

    public function generar($desde, $hasta)
    {
        /**
         * se crea una nueva instancia de PowerPoint
         */
        $objPPT = new PhpPresentation();
        $objPPT->getDocumentProperties()
            ->setCreator('Austem');

        //DIBUJAR GRAFICO BLANCOS
        $this->blankDraw($objPPT, $desde, $hasta);

        //GUARDAR EN EL EQUIPO
        $oWriterPPTX = IOFactory::createWriter($objPPT, 'PowerPoint2007');
        $rutaArchivo = storage_path("/app") . "/sample" . date('d-m-Y') . ".pptx";
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
                $dataBD = $this->filtroBlancos($vId, $vAssay, $vDesde, $vHasta);

                $seriesData = [];
                $seriesError = [];
                $seriesWarning = [];

                $contWarning = 0;
                $contError = 0;

                /*$cont = 0;
                foreach ($dataBD as $item) {
                    $cont++;
                    $desdeReal = date('d-m-y', strtotime($item->RETURNDATE));
                    $fecha = date('d-m', strtotime($item->RETURNDATE));

                    $seriesData[$cont][$fecha] = floatval($item->ASSAYVALUE);
                   // $fechaAux = $fecha;
                    $hastaReal = date('d-m-y', strtotime($item->RETURNDATE));
                }*/
                $cont = 0;
                foreach ($dataBD as $item) {

                    if (count($seriesData) == 0) {
                        $desdeReal = date('d-m-Y', strtotime($item->RETURNDATE));
                    }
                    $aux = date('j-M', strtotime($item->RETURNDATE));
                    //SE PUEDE USAR ASSAY_PRIORITY >=2
                    if ($item->ASSAYVALUE >= 0.006 && $item->ASSAYVALUE < 0.01) {
                        $contWarning++;
                        $seriesWarning[$cont] = floatval($item->ASSAYVALUE);
                    } elseif ($item->ASSAYVALUE >= 0.01) {
                        $contError++;
                        $seriesError[$cont] = floatval($item->ASSAYVALUE);
                    } else {
                        $seriesData[$cont] = floatval($item->ASSAYVALUE);
                        $seriesWarning[$cont] = 0;
                        $seriesError[$cont] = 0;
                    }
                    $hastaReal = date('d-m-Y', strtotime($item->RETURNDATE));
                    $cont++;
                }

                $contMax = max($seriesData);
                $contMin = min($seriesData);
                $auxM = array_sum($seriesData) / count($seriesData);
                $mean = number_format($auxM, 3);
                $pctError = round(($contError * 100) / count($seriesData), 3);

                $serieAux = $seriesData;
                rsort($serieAux);
                $middle = (count($serieAux) - 1 / 2);
                $median = $serieAux[$middle - 1];

                $devStd = round(stats_standard_deviation($seriesData), 3);
                $pctDevStd = round(floatval($devStd / $mean) * 100, 3);
                $stdError = round($devStd / sqrt(count($seriesData)), 3);
                $pctErrorStd = round(floatval($stdError / $mean) * 100, 3);

                /*   $stdAssays = Dtm_standardsassay::all();

                   foreach ($stdAssays as $assay) {
                       $stdValue = $assay->STANDARDVALUE;
                   }

                   //$bias = $mean / $stdValue[1];
                   //$pctBias = $bias * 100;*/

                //TABLA DATOS
                $valores = array(
                    //TOTAL
                    1 => count($seriesData),
                    //OUTSIDE WARNING
                    2 => $contWarning,
                    3 => $contError,
                    4 => count($dataBD),
                    5 => $pctError,
                    6 => $mean,
                    7 => $median,
                    8 => $contMin,
                    9 => $contMax,
                    10 => $devStd,
                    11 => $pctDevStd,
                    12 => $stdError,
                    13 => $pctErrorStd,
                    14 => "",
                    15 => ""
                );

                $this->crearTablas($currentSlide, $valores);

                $this->crearTitulo($currentSlide, $vId, $vAssay, $desdeReal, $hastaReal);
                $tipoGrafico = "bar";
                $barChart = $this->crearSeries($tipoGrafico, $seriesData, $seriesWarning, $seriesError);

                $this->crearGrafico($currentSlide, $vId, $vAssay, $barChart);
            }
        }

    }

    public function crearSeries($tipoGrafico, $seriesData, $seriesWarning, $seriesError)
    {
        $barChart = new Bar();
        $barLine = new Line();
        $barSerie = new Series("Linea", array(0 => 0.05, 1 => 0.1));
        $barChart->setGapWidthPercent(150);

        $series = new Series('Aprobados', $seriesData);
        $seriesW = new Series('Warning', $seriesWarning);
        $seriesE = new Series('Error', $seriesError);

        $series->setShowSeriesName(false);
        $seriesW->setShowSeriesName(false);
        $seriesE->setShowSeriesName(false);

        $series->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFA500'));
        $seriesW->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFFF00'));
        $seriesE->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF0000'));

        $series->getFont()->getColor()->setRGB('00FF00');
        $seriesW->getFont()->getColor()->setRGB('00FF00');
        $seriesE->getFont()->getColor()->setRGB('00FF00');

        $series->setShowValue(false);
        $seriesW->setShowValue(false);
        $seriesE->setShowValue(false);

        $barChart->addSeries($barSerie);
        $barChart->addSeries($series);
        if ($seriesW->getValues()) {
            $barChart->addSeries($seriesW);
        }
        if ($seriesE->getValues()) {
            $barChart->addSeries($seriesE);
        }
        //$barChart->setBarGrouping(Bar::GROUPING_CLUSTERED);

        return $barChart;
    }

    public function dupDraw()
    {
    }

    public function stdDraw()
    {
    }

    public function filtroBlancos($standardID, $assayName, $pDesde, $pHasta)
    {
        /*  $datosBD = DB::table('DTM_QAQC_BLK_STD')
          ->join('DTM_COLLAR', 'DTM_QAQC_BLK_STD.HOLEID', '=', 'DTM_COLLAR.HOLEID')
          ->where('DTM_COLLAR.PROJECTCODE', '=', "IN-FILL")
          ->whereRaw('DTM_COLLAR.STATUS in (\'Extraible\',\'Modelable\',\'Remapeo\',\'Recodificacion\' )')
          ->where('DTM_QAQC_BLK_STD.STANDARDID', '=', "$standardID")
          ->where('DTM_QAQC_BLK_STD.ASSAYNAME', '=', "$assayName")
          ->where('DTM_QAQC_BLK_STD.ASSAY_PRIORITY', '=', 1)
          ->whereRaw("TRY_CONVERT(date, DTM_QAQC_BLK_STD.RETURNDATE, 103) Between ('$pDesde') AND ('$pHasta')")
          ->select('DTM_QAQC_BLK_STD.*')
          ->get();
      return $datosBD;*/

        return DB::select("select B.*
        from DTM_QAQC_BLK_STD AS B inner join DTM_COLLAR AS C
        on (B.HOLEID = C.HOLEID)
        where B.STANDARDID = '$standardID'
        and B.ASSAYNAME = '$assayName'
        and B.ASSAY_PRIORITY=1
        and (TRY_CONVERT(date, B.RETURNDATE,103))
        between '$pDesde' and '$pHasta'
        and C.PROJECTCODE = 'IN-FILL'
        and C.STATUS in ('Extraible','Modelable','Remapeo','Recodificacion')
        order by (TRY_CONVERT(date, B.RETURNDATE,103)) ASC;
        ");

    }

    public function filtroDup()
    {
    }

    public function filtroStd()
    {
    }

    public function crearTitulo($currentSlide, $vId, $vAssay, $desdeReal, $hastaReal)
    {
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

    }

    public function crearTablas($currentSlide, array $valores)
    {

        $tableShape = $currentSlide->createTableShape(2);
        $tableShape->setResizeProportional(false)
            ->setWidth(330)
            ->setOffsetX(730)
            ->setOffsetY(200);

        $row0 = $tableShape->createRow()
            ->setHeight(25);
        $cell00 = $row0->nextCell();
        $cell00->setColSpan(2);
        $cell00->CreateTextRun("STATISTICS");
        $cell00->getActiveParagraph()
            ->getAlignment()
            ->setHorizontal(Alignment::HORIZONTAL_CENTER)
            ->setVertical(Alignment::VERTICAL_CENTER);


        $cellValor = $row0->getCell(1);
        $cellValor->setWidth(60);


        for ($i = 1; $i <= 15; $i++) {
            $row = $tableShape->createRow()
                ->setHeight(32);
            $cellTitulo = $row->getCell(0);
            $cellTitulo->createTextRun(self::Datos_Tabla[$i]);

            $cellValor = $row->getCell(1);
            $cellValor->createTextRun(strval($valores[$i]));
            $cellValor->getActiveParagraph()
                ->getAlignment()
                ->setHorizontal(Alignment::HORIZONTAL_CENTER)
                ->setVertical(Alignment::VERTICAL_CENTER);

            $cellTitulo->getActiveParagraph()
                ->getAlignment()
                ->setVertical(Alignment::VERTICAL_CENTER)
                ->setMarginLeft(2);

        }


    }

    public function crearGrafico($currentSlide, $vId, $vAssay, $chartType)
    {
        //GRAFICO
        $oGrid = new Gridlines();
        $oGrid->getOutline()->setWidth(1);
        $oGrid->getOutline()->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->setStartColor(new Color(Color::COLOR_BLUE));

        $oOutlineAxisX = new Outline();
        $oOutlineAxisX->setWidth(0.1);
        $oOutlineAxisX->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()->setRGB('000000');


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
        $chartShape->getPlotArea()->getAxisX()->setOutline($oOutlineAxisX);
        //$chartShape->getPlotArea()->getAxisX()->setMinBounds(10);

        $chartShape->getPlotArea()->getAxisY()->setOutline($oOutlineAxisX);


        //$chartShape->getPlotArea()->getAxisY()->setFormatCode('#,##0');


        //   $chartShape->getPlotArea()->getAxisY(2)->setMajorGridlines($oGrid);

        //SALTOS ENTRE VALORES EN Y
        //$chartShape->getPlotArea()->getAxisY()->setMajorUnit(0.002);
        //VALOR MAXIMO PARA Y
        //$chartShape->getPlotArea()->getAxisY()->setMaxBounds(0.1);

        $chartShape->getLegend()->getBorder()->setLineStyle(Border::LINE_SINGLE);


        $chartShape->getPlotArea()->setType($chartType);
        $chartShape->getLegend()->setVisible(true);
        $chartShape->getLegend()->setPosition(Legend::POSITION_BOTTOM);

    }
}
