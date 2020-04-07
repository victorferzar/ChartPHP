<?php


namespace App\Http\Controllers;
header('Content-type: application/vnd.openxmlformats-officedocument.presentationml.presentation');


use Illuminate\Http\Request;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Line;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Scatter;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Outline;
use PhpOffice\PhpPresentation\Style\Shadow;
use PhpOffice\PhpPresentation\Shape\Table;
use PhpOffice\PhpPresentation\Shape\Table\Row;
use PhpOffice\PhpPresentation\Shape\Table\Cell;
use PhpOffice\PhpPresentation\Shape\Placeholder;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar;
use PhpOffice\PhpPresentation\Style\Alignment;
use App;

class GeneratePPT extends Controller
{


    public function index()
    {
        return view('welcome');
    }

    public function generateppt(Request $request)
    {
        //return $request->all();

        /**
         * se crea una nueva instancia de PowerPoint
         */
        $objPPT = new PhpPresentation();
        $objPPT->getDocumentProperties()->setCreator('Austem');

        $dataBD = App\Dtm_qaqc_blk_std::where('STANDARDID', 'BF40')
            ->where('ASSAYNAME', 'CuT_CMCCAAS_pct')
            ->get();

        //DATOS DEL GRAFICO
        $cont = 1;
        $series1Data = [];
        foreach ($dataBD as $num) {
            $series1Data[] = floatval($num->ASSAYVALUE);
            $cont++;
        }

        foreach ($series1Data as $enum) {
            return $series1Data+$enum;
        }

        //$blanks = DTM_QAQC_BLK_STD::all();


        $this->blankDraw($objPPT);


        //GUARDAR EN EL EQUIPO
        $oWriterPPTX = IOFactory::createWriter($objPPT, 'PowerPoint2007');

        $rutaArchivo = __DIR__ . "/sample.pptx";
        $oWriterPPTX->save($rutaArchivo);
        readfile($rutaArchivo);
        exit;
    }


    /***
     * CREA LA DIAPOSITIVA Y ESTABLECE EL FONDO DEL OBJETO Y LA SOMBRA QUE DESPERENDE
     * @param PhpPresentation $objPPT ducumento ppt
     * @param $seriesData  valores del grafico
     * @throws \Exception     *
     */
    public function crearSlide(PhpPresentation $objPPT, $seriesData)
    {

        //Llena un objeto Shape de un color solido, en este caso
        $oFill = new Fill();
        $oFill->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ffe6e6'));

        //Sombra del objeto Shape
        $oShadow = new Shadow();
        $oShadow->setVisible(true)->setDirection(45)->setDistance(10);

        //Define Line Chart lines
        $oOutline = new Outline();
        $oOutline->getFill()->setFillType(Fill::FILL_SOLID);
        $oOutline->getFill()->setStartColor(new Color(Color::COLOR_BLACK));
        $oOutline->setWidth(1);

        // Create a line chart (that should be inserted in a shapeLeft)
        $lineChart = new Line();
        $series = new Series('Views', $seriesData);

        $series->setShowSeriesName(false);
        $series->setShowValue(true);
        $series->getFont()->setSize(12);
        $series->setOutline($oOutline);

        $lineChart->addSeries($series);

        //  $this->chartLeft($currentSlide, $oFill, $oShadow, $lineChart);
        // $this->chartRight($currentSlide, $oFill, $oShadow, $seriesData);
    }

    public function chartLeft($currentSlide, $oFill, $oShadow, $lineChart)
    {
        $shapeLeft = $currentSlide->createChartShape();
        $shapeLeft->setName('PHPPRESENTATION DE LA IZQUIERDA')
            ->setResizeProportional(false)
            ->setHeight(400)
            ->setWidth(450)
            ->setOffsetX(10)
            ->setOffsetY(120);
        $shapeLeft->setShadow($oShadow);
        $shapeLeft->setFill($oFill);
        $shapeLeft->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $shapeLeft->getTitle()->setText('PHPPRESENTATION DE LA IZQUIERDA');
        $shapeLeft->getTitle()->getFont()->setItalic(true);
        $shapeLeft->getPlotArea()->setType($lineChart);
        $shapeLeft->getView3D()->setRotationX(30);
        $shapeLeft->getView3D()->setPerspective(30);
        $shapeLeft->getLegend()->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $shapeLeft->getLegend()->getFont()->setItalic(true);
    }

    public function chartRight($currentSlide, $oFill, $oShadow, $seriesData)
    {
        $lineChart = new Scatter();
        $series = new Series('Valores', $seriesData);
        $series->setShowSeriesName(true);
        $series->getMarker()->setSymbol(\PhpOffice\PhpPresentation\Shape\Chart\Marker::SYMBOL_CIRCLE);
        $series->getMarker()->setSize(10);
        $lineChart->addSeries($series);

        $shapeRight = $currentSlide->createChartShape();
        $shapeRight->setName('PHPPRESENTATION DE LA DERECHA')->setResizeProportional(false)->setHeight(400)->setWidth(450)->setOffsetX(500)->setOffsetY(120);
        $shapeRight->setShadow($oShadow);
        $shapeRight->setFill($oFill);
        //$shapeRight->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $shapeRight->getTitle()->setText('PHPPRESENTATION DE LA DERECHA');
        $shapeRight->getTitle()->getFont()->setItalic(true);
        $shapeRight->getPlotArea()->setType($lineChart);
        $shapeRight->getView3D()->setRotationX(30);
        $shapeRight->getView3D()->setPerspective(30);
        $shapeRight->getLegend()->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $shapeRight->getLegend()->getFont()->setItalic(true);
    }

    public function blankDraw(PhpPresentation $objPPT)
    {
        //FILTRO BD
        $dataBD = App\Dtm_qaqc_blk_std::where('STANDARDID', 'BF40')
            ->where('ASSAYNAME', 'CuT_CMCCAAS_pct')
            ->get();

        //DATOS DEL GRAFICO
        $series1Data = array();
        foreach ($dataBD as $num) {
            return $num;
        }
//        $series2Data = array('Jan' => 266, 'Feb' => 198, 'Mar' => 271, 'Apr' => 305, 'May' => 267, 'Jun' => 301, 'Jul' => 340, 'Aug' => 326, 'Sep' => 344, 'Oct' => 364, 'Nov' => 383, 'Dec' => 379);


        //DATOS DE LA TABLA

        $dataA = array(
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

        $dataB = array();

        $currentSlide = $objPPT->createSlide();
        $tableShape = $currentSlide->createTableShape(2);

        $tableShape->setResizeProportional(false)->setHeight(200)->setWidth(250)->setOffsetX(700)->setOffsetY(100);

        $row0 = $tableShape->createRow();

        $cell00 = $row0->nextCell();
        $cell00->CreateTextRun("STATISTICS");

        $cell01 = $row0->getCell(1);
        $cell01->createTextRun(strval($dataBD->count()));
        $cell01->setWidth(50);

        for ($i = 1; $i <= 15; $i++) {
            $row[$i] = $tableShape->createRow();

            $cellAux = $row[$i]->getCell(0);
            $cellAux->CreateTextRun($dataA[$i]);
        }

        //GRAFICO

        $barChart = new Bar();
        $barChart->setGapWidthPercent(158);
        $series1 = new Series('2009', $series1Data);
        $series1->setShowSeriesName(true);
        $series1->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF4F81BD'));
        $series1->getFont()->getColor()->setRGB('00FF00');
        $series1->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFE06B20'));
        /* $series2 = new Series('2010', $series2Data);
         $series2->setShowSeriesName(true);
         $series2->getFont()->getColor()->setRGB('FF0000');
         $series2->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFC0504D'));
         $series2->setLabelPosition(Series::LABEL_INSIDEEND);*/
        $barChart->addSeries($series1);
        // $barChart->addSeries($series2);


        $chartShape = $currentSlide->createChartShape();
        $chartShape->setName("Grafico de Blancos")->setResizeProportional(false)->setHeight(400)
            ->setWidth(450)
            ->setOffsetX(10)
            ->setOffsetY(120);
        $chartShape->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $chartShape->getTitle()->setText('PHPPresentation Monthly Downloads');
        $chartShape->getTitle()->getFont()->setItalic(true);
        $chartShape->getTitle()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        $chartShape->getPlotArea()->getAxisX()->setTitle('Month');
        $chartShape->getPlotArea()->getAxisY()->getFont()->getColor()->setRGB('00FF00');
        $chartShape->getPlotArea()->getAxisY()->setTitle('Downloads');
        $chartShape->getPlotArea()->setType($barChart);
        $chartShape->getLegend()->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $chartShape->getLegend()->getFont()->setItalic(true);


    }
}

