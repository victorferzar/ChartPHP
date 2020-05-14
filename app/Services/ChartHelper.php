<?php

namespace App\Services;


use Spatie\Browsershot\Browsershot;

class ChartHelper
{
    /**
     * Handle the incoming request.
     *
     * @param \Illuminate\Http\Request $request
     * @return \Illuminate\Http\Response
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
            $nombre = 'imageStandard' . date('Y-m-d') . '.png';
        } elseif($parametros["tipo"]=="normalizado") {
            $nombre = 'imageDoble' . date('Y-m-d') . '.png';
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
            ->windowSize(820, 410)
            ->showBackground()
            ->waitUntilNetworkIdle()
            ->useCookies($cookie)
            ->save($rutaImagen);
        return $rutaImagen;
    }

}
