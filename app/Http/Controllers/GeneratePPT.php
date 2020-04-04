<?php

namespace App\Http\Controllers;
header('Content-type: application/vnd.openxmlformats-officedocument.presentationml.presentation');


use Illuminate\Http\Request;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Line;
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

//funcion que genera el ppt
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
        /**
         * Creacion de la diapositiva
         */
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

        // Create a line chart (that should be inserted in a shape)
        $lineChart = new Line();
        $series = new Series('Views', $seriesData);

        $series->setShowSeriesName(false);
        $series->setShowValue(true);
        $series->getFont()->setSize(12);
        $series->setOutline($oOutline);

        $lineChart->addSeries($series);
        /**
         * Creacion del Shape (forma)
         * Shape = objetos que pueden ser agregados a una diapositiva
         */
        $shape = $currentSlide->createChartShape();
        $shape->setName('PHPPRESENTATION PRUEBA DE GRAFICO')->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $shape->setShadow($oShadow);
        $shape->setFill($oFill);
        $shape->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $shape->getTitle()->setText('PHPPRESENTATION PRUEBA DE GRAFICO');
        $shape->getTitle()->getFont()->setItalic(true);
        $shape->getPlotArea()->setType($lineChart);
        $shape->getView3D()->setRotationX(30);
        $shape->getView3D()->setPerspective(30);
        $shape->getLegend()->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $shape->getLegend()->getFont()->setItalic(true);

    }

}

