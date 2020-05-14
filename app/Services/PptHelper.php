<?php


namespace App\Services;

use Illuminate\Support\Facades\DB;

use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Chart\Gridlines;
use PhpOffice\PhpPresentation\Shape\Chart\Legend;

use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Line;

use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Outline;


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

    public const Datos_Filtro_Standard_AssayName = array(
        1 => "CuFe_pct_FES3paAA",
        2 => "Fe_pct_2A15tAA",
        3 => "FeT_pct_TOT6tAA",
        4 => "S-pct_NON1sLE",
        5 => "S2_pct_NAO6hLE",
    );

    public const Datos_Filtro_Blancos_StandardId = array(
        1 => "BF42",
        2 => "BG4"
    );

    public function standardIDActivo($desde, $hasta, $tipo)
    {
        if ($tipo == 'blanco') {
            $sID = 'B%';
        } elseif ($tipo == 'estandar') {
            $sID = 'ST%';
        }
        $standardID_set = DB::select("SELECT distinct B.STANDARDID
        FROM DTM_QAQC_BLK_STD B
        inner join DTM_COLLAR C on B.HOLEID = C.HOLEID
        where (TRY_CONVERT(date, B.RETURNDATE,103))
        between  '$desde' and '$hasta'
        and C.PROJECTCODE = 'IN-FILL'
        and B.STANDARDID like '$sID'
        ;");

        foreach ($standardID_set as $item) {
            $standardID_array[] = $item->STANDARDID;
        }
        return $standardID_array;
    }

    public function assayNameActivos($desde, $hasta, $tipo)
    {

        if ($tipo == 'blanco') {
            $sID = 'B%';
            $aName = "and B.ASSAYNAME like 'Cu%'";
        } elseif ($tipo == 'estandar') {
            $sID = 'ST%';
            $aName = '';
        }
        $assayName_set = DB::select("SELECT distinct B.ASSAYNAME
        FROM DTM_QAQC_BLK_STD B
        inner join DTM_COLLAR C on B.HOLEID = C.HOLEID
        where (TRY_CONVERT(date, B.RETURNDATE,103))
        between  '$desde' and '$hasta'
        and C.PROJECTCODE = 'IN-FILL'
        and B.STANDARDID like '$sID'
		   $aName
        ;");
        foreach ($assayName_set as $item) {
            $assayName_array[] = $item->ASSAYNAME;
        }


        return $assayName_array;

    }

    public function suiteActivo($desde, $hasta)
    {
        $analysisuite_set = DB::select("SELECT distinct B.ANALYSISSUITE
        FROM DTM_QAQC_BLK_STD B
        inner join DTM_COLLAR C on B.HOLEID = C.HOLEID
        where (TRY_CONVERT(date, B.RETURNDATE,103))
        between '$desde' and '$hasta'
        and C.PROJECTCODE = 'IN-FILL';");

        foreach ($analysisuite_set as $item) {
            $analysis_suite_array[] = $item->ANALYSISSUITE;
        }

        return $analysis_suite_array;
    }

    public const Datos_Filtro_Standard_Id = array(
        1 => "ST43",
        2 => "ST45",
        3 => "ST46",
        4 => "ST47",
        5 => "ST48",
        6 => "ST49",
        7 => "ST50",
        8 => "ST52"
    );

    public const Datos_Tabla = array(
        1 => "# of Analyses Above Threshold",
        2 => "# of Outside Warning Limit",
        3 => "# of Outside Error Limit",
        4 => "# of Analyses Bellow Threshold (n° dataSet)",
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
        $this->blancoControl($objPPT, $desde, $hasta);
        $this->estandarControl($objPPT, $desde, $hasta);

        //   GUARDAR EN EL EQUIPO
//        $oWriterPPTX = IOFactory::createWriter($objPPT, 'PowerPoint2007');
//        $rutaArchivo = storage_path("/app") . "/sample" . date('d-m-Y') . ".pptx";
//        $oWriterPPTX->save($rutaArchivo);
//        header('Content-type: application/vnd.openxmlformats-officedocument.presentationml.presentation');
//        readfile($rutaArchivo);
//        exit;
    }

    public function blancoControl(PhpPresentation $objPPT, $vDesde, $vHasta)
    {
        $tipo = 'blanco';
        $blk_standard_id = $this->standardIDActivo($vDesde, $vHasta, $tipo);
        $blk_assayname = $this->assayNameActivos($vDesde, $vHasta, $tipo);
        // $suite_activo = $this->suiteActivo($vDesde, $vHasta);

        foreach ($blk_standard_id as $key => $id) {
            foreach ($blk_assayname as $key => $assay) {

                //FILTRO BD
                $dataBD = $this->filtroBD($tipo, $id, $assay, $vDesde, $vHasta);

                $cont = 0;
                if (empty($dataBD)) {
                    $cont++;
                    continue;
                }

                $currentSlide = $objPPT->createSlide();
                $serieTotal = [];
                $seriesData = [];
                $seriesError = [];
                $seriesWarning = [];

                $contWarning = 0;
                $contError = 0;

                $cont = 0;
                $cont_reg = 0;

                foreach ($dataBD as $item) {
                    $cont++;
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
                        $cont_reg++;
                    }
                    $hastaReal = date('d-m-Y', strtotime($item->RETURNDATE));

                    $serieTotal[$cont] = floatval($item->ASSAYVALUE);

                }

                $total_set = count($dataBD);
                $valores = $this->calculosTabla($total_set, $cont_reg, $serieTotal, $contError, $contWarning);

                $this->crearTablas($currentSlide, $valores);

                $this->crearTitulo($currentSlide, $id, $assay, $desdeReal, $hastaReal);

                $barChart = $this->crearSeries($seriesData, $seriesWarning, $seriesError);
                $this->crearGrafico($currentSlide, $id, $assay, $barChart);
            }

        }

    }

    public function estandarControl(PhpPresentation $objPPT, $vDesde, $vHasta)
    {
        $tipo = "estandar";
        $std_standard_id = $this->standardIDActivo($vDesde, $vHasta, $tipo);
        $std_assayname = $this->assayNameActivos($vDesde, $vHasta, $tipo);

        $cont = 0;

        foreach ($std_standard_id as $key => $id) {

            foreach ($std_assayname as $key => $assay) {
                $cont++;

            }

        }


        $vId = 'ST43';
        $vAssay = 'CuT_CMCCAAS_pct';
        $desdeReal = 'Nol';
        $hastaReal = 'No';


        $filtro_array = [
            'tipo' => $tipo,
            'id' => $vId,
            'assay' => $vAssay,
            'desde' => $vDesde,
            'hasta' => $vHasta
        ];

        $currentSlide = $objPPT->createSlide();

        $data_set = $this->filtroBD($tipo, $vId, $vAssay, $vDesde, $vHasta);
        $datos_grafico = $this->datosGrafico($data_set);

        $cont_error = 0;
        $cont_warn = 0;
        $cont_total = 0;
        $serie_total = [];
        $total_reg = 0;

        foreach ($data_set as $item) {
            if (count($serie_total) == 0) {
                $desdeReal = date('d-m-Y', strtotime($item->RETURNDATE));
            }

            if ((floatval($item->ASSAYVALUE)) >= ($datos_grafico["error_max"]) ||
                floatval($item->ASSAYVALUE) <= ($datos_grafico["error_min"])) {
                $cont_error++;
            } elseif ((floatval($item->ASSAYVALUE)) >= ($datos_grafico["warn_max"]) ||
                floatval($item->ASSAYVALUE) <= ($datos_grafico["warn_min"])) {
                $cont_warn++;
            } else {
                $cont_total++;
                $serie_total[$cont_total] = floatval($item->ASSAYVALUE);
            }
            $hastaReal = date('d-m-Y', strtotime($item->RETURNDATE));
            $total_reg++;
        }


        //llamada a browsershot y le envia los parametros para usar el filtroBD()
        $chartHelp = new ChartHelper();
        $ruta = $chartHelp->generarImagen($filtro_array);

        //TABLA DATOS
        $valores = $this->calculosTabla($total_reg, $cont_total, $serie_total, $cont_error, $cont_warn);

        $this->crearTablas($currentSlide, $valores);
        $this->crearTitulo($currentSlide, $vId, $vAssay, $desdeReal, $hastaReal);

        $this->crearDibujo($currentSlide, $ruta);
    }


    public function stdNorm($data_set)
    {
        $norm = [];
        foreach ($data_set as $item) {
            $aux = date('d-m', strtotime($item->RETURNDATE));

            if (!array_key_exists($aux, $norm)) {
                $norm[$aux][] = round(($item->STANDARDVALUE - $item->ASSAYVALUE) / $item->STANDARDDEVIATION, 3);
            } else {
                $norm[$aux][] = round(($item->STANDARDVALUE - $item->ASSAYVALUE) / $item->STANDARDDEVIATION, 3);
            }
        }
        $error_max = 3;
        $error_min = -3;
        $warn_max = 2;
        $warn_min = -2;
        $acept = 0;

        $datos_grafico = [
            'series' => $norm,
            'error_max' => $error_max,
            'error_min' => $error_min,
            'std_value' => $acept,
            'warn_max' => $warn_max,
            'warn_min' => $warn_min
        ];


        return $norm;
    }

    //retorna las lineas del grafico
    public function datosGrafico($data_set)
    {
        $this->stdNorm($data_set);

        //obtener datos para el grafico
        $cont = 0;
        $series_data = [];
        foreach ($data_set as $item) {
            // Llave para armar el array asociativo
            $cat = "Aprobado"; //date('Y-m-d', strtotime($item->RETURNDATE));
            $fecha = date('Y-m-d', strtotime($item->RETURNDATE));

            // Si no existe la llave, inicializa con un array vacío
            if (!isset($series_data[$cat])) {
                $series_data[$cat] = [];

            }
            // Agrega el valor dentro del array en la posición "$cat"
            $series_data[$cat][] = [strtotime($fecha), floatval($item->ASSAYVALUE)];
           // $series_data[$cat][] = [$fecha, floatval($item->ASSAYVALUE)];
        }
        // Y ahora le damos la foirema final

        foreach ($series_data as $cat => $values) {
            $series[] = ['name' => $cat, 'data' => $values];
        }

       // dd($series[0]["data"][30][0]);

        $std_value = round($item->STANDARDVALUE, 3);
        $error_max = round($std_value + ($item->STANDARDDEVIATION) * 3, 3);
        $error_min = round($std_value - ($item->STANDARDDEVIATION) * 3, 3);
        $warn_max = round($std_value + ($item->STANDARDDEVIATION) * 2, 3);
        $warn_min = round($std_value - ($item->STANDARDDEVIATION) * 2, 3);

        $datos_grafico = [

            'series' => $series,
            'error_max' => $error_max,
            'error_min' => $error_min,
            'std_value' => $std_value,
            'warn_max' => $warn_max,
            'warn_min' => $warn_min,

        ];
        return $datos_grafico;
    }

    public function filtroBD($tipo, $standardID, $assayName, $pDesde, $pHasta)
    {

        $suite = "SUPC_2015";

        $auxSelect = " ";
        $auxInner = " ";
        if ($tipo == 'estandar') {
            $auxInner = "inner join DTM_STANDARDSASSAY S on B.ASSAYNAME = S.NAME and B.STANDARDID = S.STANDARDID";
            $auxSelect = ", C.PROJECTCODE, S.STANDARDVALUE, S.STANDARDDEVIATION, S.ACCEPTABLEMAX, S.ACCEPTABLEMIN";
        }
        return DB::select("SELECT B.* $auxSelect
        FROM DTM_QAQC_BLK_STD as B $auxInner
         inner join DTM_COLLAR as C
          on B.HOLEID = C.HOLEID
        where B.STANDARDID = '$standardID'
         and B.ASSAYNAME = '$assayName'
         and B.ASSAY_PRIORITY=1
         and (TRY_CONVERT(date, B.RETURNDATE,103))
          between '$pDesde' and '$pHasta'
         and C.PROJECTCODE = 'IN-FILL'
         and C.STATUS in ('Extraible','Modelable','Rem  apeo','Recodificacion')
         and (ANALYSISSUITE = '$suite')
        order by (TRY_CONVERT(date, B.RETURNDATE,103)) ASC;");

    }

    public function calculosTabla($count_reg, $seriesData, $serieTotal, $contError, $contWarning)
    {
        //MEAN
        if (count($serieTotal) == 0) {
            $auxM = 0;
        } else {
            $auxM = array_sum($serieTotal) / count($serieTotal);
        }

        $mean = number_format($auxM, 3);

        //MEDIAN
        $serieAux = $serieTotal;

        $midd = ($count_reg - 1) / 2;

        $median = $serieAux[$midd];


        $contMin = min($serieTotal);
        $contMax = max($serieTotal);
        //STANDARD DEVIATION
        $devStd = round(stats_standard_deviation($serieTotal), 3);

        //%OUTSIDE ERROR LIMIT
        $pctError = round(($contError * 100) / count($serieTotal), 3);

        //% REL.STD.DEV
        $pctDevStd = round(floatval($devStd / $mean) * 100, 3);
        //STANDARD ERROR
        $stdError = round($devStd / sqrt(count($serieTotal)), 3);
        //% REL.STD.ERR
        $pctErrorStd = round(floatval($stdError / $mean) * 100, 3);

        $bias = round(doubleval($auxM / 0.001) - 1, 3);
        $pctBias = round($bias * 100, 3);

        $valores = array(
            //TOTAL
            1 => $seriesData,
            //OUTSIDE WARNING
            2 => $contWarning,
            3 => $contError,
            4 => $count_reg,
            5 => $pctError,
            6 => $mean,
            7 => $median,
            8 => $contMin,
            9 => $contMax,
            10 => $devStd,
            11 => $pctDevStd,
            12 => $stdError,
            13 => $pctErrorStd,
            14 => $bias,
            15 => $pctBias
        );
        return $valores;
    }

    public function crearSeries($seriesData, $seriesWarning, $seriesError)
    {
        $barChart = new Bar();

        $barLine = new Line();

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

    public function crearTitulo($currentSlide, $vId, $vAssay, $desdeReal, $hastaReal)
    {
        if ($vId == "BF42") {
            $vTipo = "Fino";
        } elseif ($vId == "BG4") {
            $vTipo = "Grueso";
        } else {
            $vTipo = "Estandares";
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
            ->setWidth(800)
            ->setOffsetX(200)
            ->setOffsetY(30);

        $txtTitulo->getActiveParagraph()
            ->getAlignment()
            ->setHorizontal(Alignment::HORIZONTAL_CENTER);

        $txtSubTitulo = $txtTitulo;
        $txtDetalle = $txtTitulo;

        $txtRunT = $txtTitulo->createTextRun($vTipo . " - " . $vAssay . "\n");
        $txtRunST = $txtSubTitulo->createTextRun($vTipo . " Aprobados " . $vId . " - " . "Blanco " . $vTipo . "\n");
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

    public function crearDibujo($currentSlide, $ruta)
    {
        $chartImage = $currentSlide->createDrawingShape();

        $chartImage->setName('Grafico Standards')->setDescription('Imagen de grafico standards');
        $chartImage->setPath($ruta);
        $chartImage->setResizeProportional(false)
            ->setHeight(250)
            ->setWidth(700)
            ->setOffsetX(10)
            ->setOffsetY(170);

        $chartImage2 = $currentSlide->createDrawingShape();
        $chartImage2->setName('Grafico Standards')->setDescription('Imagen de grafico standards');
        $chartImage2->setPath($ruta);
        $chartImage2->setResizeProportional(false)
            ->setHeight(250)
            ->setWidth(700)
            ->setOffsetX(10)
            ->setOffsetY(420);
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
        $chartShape->getPlotArea()->getAxisY()->setOutline($oOutlineAxisX);

        $chartShape->getLegend()->getBorder()->setLineStyle(Border::LINE_SINGLE);

        $chartShape->getPlotArea()->setType($chartType);
        $chartShape->getLegend()->setVisible(true);
        $chartShape->getLegend()->setPosition(Legend::POSITION_BOTTOM);

    }
}
