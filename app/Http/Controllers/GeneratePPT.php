<?php


namespace App\Http\Controllers;
header('Content-type: application/vnd.openxmlformats-officedocument.presentationml.presentation');

use Illuminate\Http\Request;
use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Line;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Scatter;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Outline;
use PhpOffice\PhpPresentation\Style\Shadow;


class GeneratePPT extends Controller
{


    public function index()
    {
        return view('welcome');
    }

    public function generateppt(Request $request)
    {
        //return $request->all();
        /*Se rescatan los datos enviados por la vista welcome usando $request
  * se asignan esos datos a una correspondiente variable
  * se insertan dichos datos en un array $seriesData
  * @seriesData = datos a mostrar como puntos en el grafico
  */
        $vLunes = (int)$request->dLunes;
        $vMartes = (int)$request->dMartes;
        $vMiercoles = (int)$request->dMiercoles;
        $vJueves = (int)$request->dJueves;
        $vViernes = (int)$request->dViernes;
        $vSabado = (int)$request->dSabado;
        $vDomingo = (int)$request->dDomingo;

        $seriesData = array(
            'Monday' => $vLunes,
            'Tuesday' => $vMartes,
            'Wednesday' => $vMiercoles,
            'Thursday' => $vJueves,
            'Friday' => $vViernes,
            'Saturday' => $vSabado,
            'Sunday' => $vDomingo
        );

        /**
         * se crea una nueva instancia de PowerPoint
         */
        $objPPT = new PhpPresentation();
        // $oMasterSlide = $objPPT->getAllMasterSlides()[0];
        // $oSlideLayout = $oMasterSlide->getAllSlideLayouts()[0];
        /*
         * Estgablecido el formato de las1 diapositivas
         */
        $objPPT->getLayout()->setDocumentLayout(DocumentLayout::LAYOUT_SCREEN_16X9);
        $objPPT->getDocumentProperties()->setCreator('Austem');

        $this->crearSlide($objPPT, $seriesData);
        //GUARDAR EN EL EQUIPO
        $oWriterPPTX = IOFactory::createWriter($objPPT, 'PowerPoint2007');

        $rutaArchivo = __DIR__ . "/sample.pptx";
        $oWriterPPTX->save($rutaArchivo);
        readfile($rutaArchivo);
        exit;
    }

    public function crearSlide(PhpPresentation $objPPT, $seriesData)
    {
        $currentSlide = $objPPT->createSlide();
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

        $this->chartLeft($currentSlide, $oFill, $oShadow, $lineChart);
        $this->chartRight($currentSlide, $oFill, $oShadow, $seriesData);
    }

    public function chartLeft($currentSlide, $oFill, $oShadow, $lineChart)
    {
        $shapeLeft = $currentSlide->createChartShape();
        $shapeLeft->setName('PHPPRESENTATION DE LA IZQUIERDA')->setResizeProportional(false)->setHeight(400)->setWidth(450)->setOffsetX(10)->setOffsetY(120);
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
}

