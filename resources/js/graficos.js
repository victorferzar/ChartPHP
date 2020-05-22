module.exports = {
    initGraficoStd: function (tipo, series, titulo, error_max, error_min, std_value, warn_min, warn_max, standardId, assayName) {

        var auxMax = Number(error_max);
        var auxMin = Number(error_min);

        if (tipo === 'estandar') {
            topeMax = auxMax + 0.02;
            topeMin = auxMin - 0.02;
            auxTick = 0.02;
        } else if (tipo === 'normalizado') {
            topeMin = -4;
            topeMax = 4;
            auxTick = 1;
        }

        Highcharts.chart('graficoStd', {
                chart: {
                    type: 'spline'
                },
                title: {
                    text: tipo,
                    style: {color: 'black'}
                },
                subtitle: {
                    text: standardId +" - "+ assayName

                },
                xAxis: {
                    type: 'datetime',
                    tickInterval: 24 * 3600 * 1000,
                    labels: {
                        style: {color: 'black'}
                    },
                    dateTimeLabelFormats: {
                        month: '%e. %b',
                        year: '%b'
                    },
                    title: {
                        text: 'Fecha de Retorno',
                        style: {color: 'black'}
                    }

                },
                yAxis: {

                    tickInterval: auxTick,
                    min: topeMin,
                    max: topeMax,

                    labels: {
                        style: {
                            color: 'black'
                        }
                    },
                    title: {
                        text: 'Ley Laboratorio',
                        style: {
                            color: 'black'
                        }
                    },
                    plotLines: [{
                        color: 'red',
                        value: error_max,
                        width: 2
                    }, {
                        color: 'red',
                        value: error_min,
                        width: 2
                    }, {
                        color: 'yellow',
                        value: warn_max,
                        width: 2
                    }, {
                        color: 'yellow',
                        value: warn_min,
                        width: 2
                    }, {
                        color: 'green',
                        value: std_value,
                        width: 2
                    }]
                },

                plotOptions: {
                    series: {
                        animation: false,
                        pointInterval: 24 * 3600 * 1000,
                        color: 'rgba(255,165,0, 1.0)',
                        marker: {
                            fillColor: 'black',
                            lineWidth: 1,
                            lineColor: 'white',
                            enabled: true,
                            radius: 1.5,


                        }
                    }
                },

                series: series,

            }
        )
    }
};
