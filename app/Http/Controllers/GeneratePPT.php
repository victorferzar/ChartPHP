<?php

namespace App\Http\Controllers;
header('Content-type: application/vnd.openxmlformats-officedocument.presentationml.presentation');

use Illuminate\Http\Request;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Alignment;


class GeneratePPT extends Controller
{
    public function index()
    {
        return view('welcome');


    }

//funcion que genera el pptsdsdsa
    public function generateppt(Request $request)
    {
        //return $request->all();
        $cantidad = $request->cantidad;


        $objPHPPowerPoint = new PhpPresentation();

        $objPHPPowerPoint->getProperties()->setCreator('Austem')
            ->setLastModifiedBy('Sketch Team')
            ->setTitle('Prueba de Presentacion')
            ->setSubject('Prueba de Presentacion')
            ->setDescription('Generacion automatica de ppts con imagenes')
            ->setKeywords('office 2007 openxml libreoffice odt php')
            ->setCategory('Sample Category');

        $objPHPPowerPoint->removeSlideByIndex(0);

        $this->slide1($objPHPPowerPoint, $cantidad);
     //   $this->slide2($objPHPPowerPoint);

        $oWriterPPTX = IOFactory::createWriter($objPHPPowerPoint, 'PowerPoint2007');

        //return $oWriterPPTX->save(__DIR__ . "/sample.pptx");
        $rutaArchivo = __DIR__ . "/sample.pptx";
        $oWriterPPTX->save(__DIR__ . "/sample.pptx");
        readfile($rutaArchivo);
        exit;

    }

    public function slide1(&$objPHPPowerPoint, $cantidad)
    {
        for ($i = 1; $i <= $cantidad; $i++) {
            // Create slide
            $currentSlide = $objPHPPowerPoint->createSlide();

            // Create a shape (drawing)
            $shape = $currentSlide->createDrawingShape();
            $shape->setName('imageLeft')
                ->setDescription('image of the left')
                ->setPath(public_path() . '/line_graph.jpg')
                ->setHeight(300)
                //  ->setWidht(200)
                ->setOffsetX(10)
                ->setOffsetY(200);

            $shape = $currentSlide->createDrawingShape();
            $shape->setName('imageRight')
                ->setDescription('image of the right')
                ->setPath(public_path() . '/environmental_data.jpg')
                ->setHeight(300)
                //  ->setWidht(200)
                ->setOffsetX(500)
                ->setOffsetY(200);

            //Create a shape (text)
            $shape = $currentSlide->createRichTextShape()
                ->setHeight(300)
                ->setWidth(600)
                ->setOffsetX(170)
                ->setOffsetY(180);
            $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            $textRun = $shape->createTextRun('Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industrys standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. ');
            $textRun->getFont()->setBold(true)
                ->setSize(16)
                ->setColor(new Color('FFE06B20'));

        }
    }
//    public function slide2(&$objPHPPowerPoint)
//    {
//
//        // Create slide
//        $currentSlide = $objPHPPowerPoint->createSlide();
//
//        // Create a shape (drawing)
//        $shape = $currentSlide->createDrawingShape();
//        $shape->setName('image')
//            ->setDescription('image')
//            ->setPath(public_path() . '/environmental_data.jpg')
//            ->setHeight(300)
//            ->setOffsetX(10)
//            ->setOffsetY(10);
//
//        // Create a shape (text)
//        $shape = $currentSlide->createRichTextShape()
//            ->setHeight(300)
//            ->setWidth(600)
//            ->setOffsetX(170)
//            ->setOffsetY(180);
//        $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
//        $textRun = $shape->createTextRun('Lorem Ipsum is simply dummy text of the printing and typesetting industry.');
//        $textRun->getFont()->setBold(true)
//            ->setSize(16)
//            ->setColor(new Color('FFE06B20'));
//
//    }
}

