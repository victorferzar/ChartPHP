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


    public function standardIDActivos($desde, $hasta, $tipo)
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

    public function suiteActivos($desde, $hasta)
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

    public const Prueba_StandardID = array(
        1 => "ST43",
        2 => "ST45"
    );

    public const Prueba_AssayName = array(
        1 => "CuCN_pct_CYN6pAA",
        2 => "CuT_CMCCAAS_pct",
        3 => "ClT_kgt_SUL2pPT",
        4 => "Fe_pct_2A15tAA"
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
        $oWriterPPTX = IOFactory::createWriter($objPPT, 'PowerPoint2007');
        $rutaArchivo = storage_path("/app") . "/sample" . date('d-m-Y') . ".pptx";
        $oWriterPPTX->save($rutaArchivo);
        header('Content-type: application/vnd.openxmlformats-officedocument.presentationml.presentation');
        readfile($rutaArchivo);
        exit;
    }

    /**
     * Funcion que crea array con fechas en ordenadas por timespam
     * @param $data_set
     * @return array
     */
    public function ordenarFechas($data_set)
    {
        $cont = 0;

        foreach ($data_set as $item) {
            $cont++;
            $array_timestamp[strtotime($item->RETURNDATE)][] = $item;
        }

        //ordena array por timespam
        ksort($array_timestamp);
        //Reorganiza array en orden
        foreach ($array_timestamp as $key => $value) {
            if (count($value) > 1) {
                for ($i = 0; count($value) > $i; $i++) {
                    $array_aux[] = $value[$i];
                }
            } else {
                $array_aux[] = $value[0];
            }
        }

        $array_timestamp = $array_aux;

        return $array_timestamp;
    }


    public function blancoControl(PhpPresentation $objPPT, $vDesde, $vHasta)
    {
        $tipo = 'blanco';
        $blk_standard_id = $this->standardIDActivos($vDesde, $vHasta, $tipo);
        $blk_assayname = $this->assayNameActivos($vDesde, $vHasta, $tipo);

        foreach ($blk_standard_id as $key => $id) {
            foreach ($blk_assayname as $key => $assay) {

                //FILTRO BD
                $data_set = $this->filtroBD($tipo, $id, $assay, $vDesde, $vHasta);
                $cont = 0;
                if (empty($data_set)) {
                    $cont++;
                    continue;
                }

                $cont_set = count($data_set);

                $data_set = $this->ordenarFechas($data_set);

                $currentSlide = $objPPT->createSlide();
                $series_total_value = [];
                $series_data = [];
                $series_error = [];
                $series_warning = [];

                $cont_warning = 0;
                $cont_error = 0;

                $cont = 0;
                $cont_aprob = 0;

                foreach ($data_set as $item) {

                    $cont++;
                    if (count($series_data) == 0) {
                        $fecha_desde_real = date('d-m-Y', strtotime($item->RETURNDATE));
                    }
                    //SE PUEDE USAR ASSAY_PRIORITY >=2
                    if ($item->ASSAYVALUE >= 0.006 && $item->ASSAYVALUE < 0.01) {
                        $cont_warning++;
                        $series_warning[$cont] = floatval($item->ASSAYVALUE);
                    } elseif ($item->ASSAYVALUE >= 0.01) {
                        $cont_error++;
                        $series_error[$cont] = floatval($item->ASSAYVALUE);
                    } else {
                        $series_data[$cont] = floatval($item->ASSAYVALUE);
                        $series_warning[$cont] = 0;
                        $series_error[$cont] = 0;
                        $cont_aprob++;
                    }
                    $fecha_hasta_real = date('d-m-Y', strtotime($item->RETURNDATE));

                    $series_total_value[$cont] = floatval($item->ASSAYVALUE);
                }

                $valores = $this->calculosTabla($series_total_value, $cont_set, $cont_aprob, $cont_warning, $cont_error);

                $this->crearTablas($currentSlide, $valores);

                $this->crearTitulo($currentSlide, $id, $assay, $fecha_desde_real, $fecha_hasta_real);

                $barChart = $this->crearSeries($series_data, $series_warning, $series_error);
                $this->crearGrafico($currentSlide, $id, $assay, $barChart);
            }
        }
    }

    public function estandarControl(PhpPresentation $objPPT, $vDesde, $vHasta)
    {
        $tipo = "estandar";
        $std_standard_id = $this->standardIDActivos($vDesde, $vHasta, $tipo);
        $std_assayname = $this->assayNameActivos($vDesde, $vHasta, $tipo);

        $cont = 0;

        foreach (self::Prueba_StandardID as $key => $id) {
            foreach (self::Prueba_AssayName as $key => $assay) {
//
                // $id = "ST43";
                //    $assay = "CuCN_pct_CYN6pAA";

                $cont++;

                $filtro_array = [
                    'tipo' => $tipo,
                    'id' => $id,
                    'assay' => $assay,
                    'desde' => $vDesde,
                    'hasta' => $vHasta
                ];

                $data_set = $this->filtroBD($tipo, $id, $assay, $vDesde, $vHasta);

                if (empty($data_set)) {
                    $cont++;
                    continue;
                }

                $data_set = $this->ordenarFechas($data_set);

                $currentSlide = $objPPT->createSlide();

                $datos_grafico = $this->datosGrafico($data_set);

                $datos_grafico_norm = $this->datosGraficoNorm($data_set);

                $cont_error = 0;
                $cont_warn = 0;
                $cont_aprob = 0;
                $series_total_values = [];
                $cont_set = 0;


                //OBTENGO FECHAS REALES, CANTIDAD DE ELEMENTOS SOBRE EL ERROR-WARNING,
                //CANTIDAD TOTAL DE ELEMENTOS ACEPTADOS Y CANTIDAD TOTAL DE ELEMENTOS
                foreach ($data_set as $item) {
                    if (count($series_total_values) == 0) {
                        $desdeReal = date('d-m-Y', strtotime($item->RETURNDATE));
                    }

                    if ((floatval($item->ASSAYVALUE)) >= ($datos_grafico["error_max"]) ||
                        floatval($item->ASSAYVALUE) <= ($datos_grafico["error_min"])) {
                        $cont_error++;
                    } elseif ((floatval($item->ASSAYVALUE)) >= ($datos_grafico["warn_max"]) ||
                        floatval($item->ASSAYVALUE) <= ($datos_grafico["warn_min"])) {
                        $cont_warn++;
                    } else {
                        $cont_aprob++;
                        $series_total_values[$cont_aprob] = floatval($item->ASSAYVALUE);
                    }
                    $hastaReal = date('d-m-Y', strtotime($item->RETURNDATE));
                    $cont_set++;
                }

                //llamada a browsershot y le envia los parametros para usar el filtroBD()
                $chartHelp = new ChartHelper();
                $ruta = $chartHelp->generarImagen($filtro_array);
                $filtro_array['tipo'] = 'normalizado';
                $ruta_norm = $chartHelp->generarImagen($filtro_array);

                //TABLA DATOS
                $valores = $this->calculosTabla($series_total_values, $cont_set, $cont_aprob, $cont_warn, $cont_error);

                $this->crearTablas($currentSlide, $valores);
                $this->crearTitulo($currentSlide, $id, $assay, $desdeReal, $hastaReal);

                $this->crearDibujo($currentSlide, $ruta, $ruta_norm);

            }
        }
    }

    public function datosGraficoNorm($data_set)
    {

        $series_data = [];
        foreach ($data_set as $item) {

            $cat = "Aprobado";
            $fecha = strtotime($item->RETURNDATE) * 1000;

            $valor = round(floatval(($item->ASSAYVALUE - $item->STANDARDVALUE) / $item->STANDARDDEVIATION), 3);

//            if (!array_key_exists($aux, $series_data)) {
//                $series_data[$aux][] = round(($item->STANDARDVALUE - $item->ASSAYVALUE) / $item->STANDARDDEVIATION, 3);
//            } else {
//                $series_data[$aux][] = round(($item->STANDARDVALUE - $item->ASSAYVALUE) / $item->STANDARDDEVIATION, 3);
//            }
            $series_data[$cat][] = [$fecha, $valor];
        }

        foreach ($series_data as $cat => $values) {
            $series[] = ['name' => $cat, 'data' => $values];
        }

        $error_max = 3;
        $error_min = -3;
        $warn_max = 2;
        $warn_min = -2;
        $acept = 0;

        $datos_grafico = [
            'series' => $series,
            'error_max' => $error_max,
            'error_min' => $error_min,
            'std_value' => $acept,
            'warn_max' => $warn_max,
            'warn_min' => $warn_min
        ];
        return $datos_grafico;
    }

    //RETORNA ARRAY CON: SERIE LISTA, LINEAS DE GRAFICO: ERROR, WARNING, STANDARD VALUE
    public function datosGrafico($data_set)
    {
        //obtener datos para el grafico
        $cont = 0;
        $series_data = [];
        foreach ($data_set as $item) {

            // Llave para armar el array asociativo
            $cat = "Aprobado"; //date('Y-m-d', strtotime($item->RETURNDATE));
            // $fecha = date('Y-m-d', strtotime($item->RETURNDATE));
            $fecha = strtotime($item->RETURNDATE) * 1000;
            // Si no existe la llave, inicializa con un array vacío
            if (!isset($series_data[$cat])) {
                $series_data[$cat] = [];
            }
            // Agrega el valor dentro del array en la posición "$cat"

            $series_data[$cat][] = [$fecha, floatval($item->ASSAYVALUE)];

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

    public function filtroBD($tipo, $standard_id, $assay_name, $fecha_desde, $fecha_hasta)
    {
        $suite_set = $this->suiteActivos($fecha_desde, $fecha_hasta);
        foreach ($suite_set as $item) {
            $suite_data[] = "'" . $item . "'";
        }
        $suite_query = implode(",", $suite_data);

        $std_select = " ";
        $std_inner_join = " ";
        if ($tipo == 'estandar' || $tipo == 'normalizado') {
            $std_inner_join = "inner join DTM_STANDARDSASSAY S on B.ASSAYNAME = S.NAME and B.STANDARDID = S.STANDARDID";
            $std_select = ", C.PROJECTCODE, S.STANDARDVALUE, S.STANDARDDEVIATION, S.ACCEPTABLEMAX, S.ACCEPTABLEMIN";
        }
        return DB::select("SELECT B.* $std_select
        FROM DTM_QAQC_BLK_STD as B $std_inner_join
         inner join DTM_COLLAR as C
          on B.HOLEID = C.HOLEID
        where B.STANDARDID = '$standard_id'
         and B.ASSAYNAME = '$assay_name'
         and B.ASSAY_PRIORITY=1
         and (TRY_CONVERT(date, B.RETURNDATE,103))
          between '$fecha_desde' and '$fecha_hasta'
         and C.PROJECTCODE = 'IN-FILL'
         and C.STATUS in ('Extraible','Modelable','Remapeo','Recodificacion','Reproceso')
         and B.ANALYSISSUITE IN ($suite_query)
        order by (TRY_CONVERT(date, B.RETURNDATE,103)) ASC;");

    }

    public function calculosTabla($series_total_values, $cont_set, $cont_aprob, $cont_warning, $cont_error)
    {
        $cont = 0;

        //%OUTSIDE ERROR LIMIT
        $pct_outside_error = round(($cont_error * 100) / count($series_total_values), 3);
        //MEAN
        if (empty($series_total_values)) {
            $auxM = 0;

        } else {
            $auxM = array_sum($series_total_values) / count($series_total_values);
        }
        $mean = round($auxM, 3, PHP_ROUND_HALF_DOWN);

        //MEDIAN
        if ($cont_set > 2) {
            $midd = floor(($cont_set - 1) / 2);
            if ($cont_set % 2) {
                $median = $series_total_values[$midd];
            } else {
                $low = $series_total_values[$midd];
                $high = $series_total_values[$midd + 1];
                $median = (($low + $high) / 2);
            }
        } else {
            $median = 0;
        }

        //MIN
        $min_value = min($series_total_values);
        //MAX
        $max_value = max($series_total_values);

        //STANDARD DEVIATION
        $std_deviation = round(stats_standard_deviation($series_total_values), 3);
//        $std_deviation += 0.001;

        //STANDARD ERROR
        $std_error = round(($std_deviation / sqrt($cont_set)), 3);

        //% REL.STD.DEV
        //% REL.STD.ERR
        if ($mean > 0) {
            $pct_rel_std_dev = round(($std_deviation / $mean) * 100, 3);
            $pct_rel_std_err = round(($std_error / $mean) * 100, 3);
        } else {
            $pct_rel_std_dev = 0;
            $pct_rel_std_err = 0;
        }


        //TOTAL BIAS
        $total_bias = round(doubleval($auxM / 0.001) - 1, 3);

        //%MEAN BIAS
        $pct_mean_bias = round($total_bias * 100, 3);

        $valores = array(
            //TOTAL
            1 => $cont_aprob,
            //OUTSIDE WARNING
            2 => $cont_warning,
            3 => $cont_error,
            4 => $cont_set,
            5 => $pct_outside_error,
            6 => $mean,
            7 => $median,
            8 => $min_value,
            9 => $max_value,
            10 => $std_deviation,
            11 => $pct_rel_std_dev,
            12 => $std_error,
            13 => $pct_rel_std_err,
            14 => $total_bias,
            15 => $pct_mean_bias
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

    public function crearDibujo($currentSlide, $ruta, $rutaNorm)
    {
        $chartImage = $currentSlide->createDrawingShape();

        $chartImage->setName('Grafico Standards')->setDescription('Imagen de grafico standards');
        $chartImage->setPath($ruta);
        $chartImage->setResizeProportional(false)
            ->setHeight(270)
            ->setWidth(710)
            ->setOffsetX(10)
            ->setOffsetY(170);

        $chartImage2 = $currentSlide->createDrawingShape();
        $chartImage2->setName('Grafico Normalizado Std')->setDescription('Imagen de grafico standards');
        $chartImage2->setPath($rutaNorm);
        $chartImage2->setResizeProportional(false)
            ->setHeight(270)
            ->setWidth(710)
            ->setOffsetX(10)
            ->setOffsetY(445);
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
