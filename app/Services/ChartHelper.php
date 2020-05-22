<?php

namespace App\Services;


use Spatie\Browsershot\Browsershot;
use Spatie\Image\Manipulations;

class ChartHelper
{
    /**
     * Handle the incoming request.
     *
     * @param \Illuminate\Http\Request $request
     * @return \Illuminate\Http\Response
     * @throws \Spatie\Browsershot\Exceptions\CouldNotTakeBrowsershot
     */
    public function generarImagen(array $parametros)
    {
        //Genera en segundo plano el grafico y lo guarda como una imagen jpg

        //Establecer cookies, para evitar errores de session
        $sessionCookie = config('session.cookie');
        $cookie = [
            $sessionCookie => $_COOKIE[$sessionCookie],
            'XSRF-TOKEN' => $_COOKIE['XSRF-TOKEN']
        ];
        //lugar y nombre donde se guarda la imagen de chart
        $path = 'app/public/';

        if ($parametros["tipo"] == "estandar") {
            $nombre = 'imgStd_' . time() . '.jpg';
        } elseif ($parametros["tipo"] == "normalizado") {
            $nombre = 'imgNorm_' . time() . '.jpg';
        }
        $rutaImagen = storage_path($path . $nombre);

        $browserShot = Browsershot::url(route('chart', $parametros));

        //binorios ejecutables
        $browserShot->setNodeBinary(env('NODE_BINARY_PATH', '/usr/local/bin/node'))
            ->setNpmBinary(env('NPM_BINARY_PATH', '/usr/local/bin/npm'))
            ->noSandbox()
            ->format('Legal')
            ->emulateMedia('screen')
            ->devicePixelRatio(2)
            ->setScreenshotType("jpeg", 100)
            ->margins(15, 10, 10, 10)
            //->windowSize(1920, 1080)
            //->fit(Manipulations::FIT_CONTAIN, 700, 700)
            ->clip(20, 10, 770, 380)
            ->showBackground()
            ->waitUntilNetworkIdle()
            ->useCookies($cookie)
            ->save($rutaImagen);

        return $rutaImagen;
    }

}
