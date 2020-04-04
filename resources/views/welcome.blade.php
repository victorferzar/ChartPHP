<!doctype html>
<html lang="en">
<head>
    <!-- Required meta tags -->
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">

    <!-- Bootstrap CSS -->
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.4.1/css/bootstrap.min.css"
          integrity="sha384-Vkoo8x4CGsO3+Hhxv8T/Q5PaXtkKtu6ug5TOeNV6gBiFeWPGFN9MuhOf23Q9Ifjh" crossorigin="anonymous">

    <title>Generador de PPT</title>
</head>
<body>

<form action="{{route('generar')}}" class="form-group">
    <h1>Generador de PPT</h1>

    <div class="form-inline">
        <input type="date" name="fecha_des" required placeholder="Fecha Desde" value="1991-09-13" disabled>
        <input type="date" name="fecha_has" required placeholder="Fecha Hasta" value="2020-04-03" disabled>
    </div>
    <div class="form-inline">
        <input type="number" name="dLunes" placeholder="Lunes: ">
        <input type="number" name="dMartes" placeholder="Martes: ">
        <input type="number" name="dMiercoles" placeholder="Miercoles: ">
        <input type="number" name="dJueves" placeholder="Jueves: ">
        <input type="number" name="dViernes" placeholder="Viernes: ">
        <input type="number" name="dSabado" placeholder="Sabado: ">
        <input type="number" name="dDomingo" placeholder="Domingo: ">


    </div>

    <input type="number" name="cantidad" required placeholder="Cantidad de diapositivas" class="form-control mb-2"
           min="1"
           max="10" disabled>
    <button type="submit" class="btn btn-primary btn-block">GENERAR</button>

</form>

<!-- Optional JavaScript -->
<!-- jQuery first, then Popper.js, then Bootstrap JS -->
<script src="https://code.jquery.com/jquery-3.4.1.slim.min.js"
        integrity="sha384-J6qa4849blE2+poT4WnyKhv5vZF5SrPo0iEjwBvKU7imGFAV0wwj1yYfoRSJoZ+n"
        crossorigin="anonymous"></script>
<script src="https://cdn.jsdelivr.net/npm/popper.js@1.16.0/dist/umd/popper.min.js"
        integrity="sha384-Q6E9RHvbIyZFJoft+2mJbHaEWldlvI9IOYy5n3zV9zzTtmI3UksdQRVvoxMfooAo"
        crossorigin="anonymous"></script>
<script src="https://stackpath.bootstrapcdn.com/bootstrap/4.4.1/js/bootstrap.min.js"
        integrity="sha384-wfSDF2E50Y2D1uUdj0O3uMBJnjuUD4Ih7YwaYd1iqfktj0Uod8GCExl3Og8ifwB6"
        crossorigin="anonymous"></script>
</body>
</html>
