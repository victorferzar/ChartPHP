<!doctype html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport"
          content="width=device-width, user-scalable=no, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <title>Document</title>
    <link rel="stylesheet" href="{{mix('/css/app.css')}}">
</head>
<body>

<!--ID DEL GRAFICO-->
<div id="graficoStd"></div>


<script src="{{mix('/js/manifest.js')}}"></script>
<script src="{{mix('/js/vendor.js')}}"></script>
<script src="{{mix('/js/app.js')}}"></script>

<!--Funcion de JS, transforma la lista de datos en JSON y se entrega a la funcion graficos.js-->
<script>
    document.addEventListener('DOMContentLoaded', function () {
        //json_encode = sintaxis de blade, imprime esto sin escapar el String
        //array de php a json
        //enconde = transforma un array en formato string json

        var tipo = '{{$tipo}}'
        var series = JSON.parse('{!! json_encode($series) !!}');
        var titulo = '{{ $titulo }}'
        var error_max = '{{$error_max}}'
        var error_min = '{{$error_min}}'
        var std_value = '{{$std_value}}'
        var warn_min = '{{$warn_min}}'
        var warn_max = '{{$warn_max}}'

        for (var i = 0; i<series.length; i+=1){
            @dd(i)

        }

        Graficos.initGraficoStd(tipo, series, titulo, error_max, error_min, std_value, warn_min, warn_max);
    });
</script>

</body>
</html>
