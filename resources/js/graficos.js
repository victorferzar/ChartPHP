module.exports = {
    initGraficoStd: function (tipo, series, titulo, error_max, error_min, std_value, warn_min, warn_max) {

        if (tipo === 'estandar') {
            topeMax = error_max ;
            topeMin = error_min;
        } else {
            topeMin = -5;
            topeMax = 5;
        }


        Highcharts.chart('graficoStd', {
            chart: {
                type: 'spline'
            },
            title: {

                text:  titulo
            },
            subtitle: {
                text: ''
            },
            xAxis: {
                type: 'datetime',
                tickInterval: 3600 * 1000,
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

                min: topeMin,

                max: topeMax,
                tickInterval: 0.02,

                plotLines: [{
                    color: 'red',
                    value: error_max,
                    width: 1
                }, {
                    color: 'red',
                    value: error_min,
                    width: 1
                }, {
                    color: 'yellow',
                    value: warn_max,
                    width: 1
                }, {
                    color: 'yellow',
                    value: warn_min,
                    width: 1
                }, {
                    color: 'green',
                    value: std_value,
                    width: 1
                }],
            },
            // tooltip: {
            //     headerFormat: '<b>{series.name}</b><br>',
            //     pointFormat: '{point.x:%e. %b}: {point.y:.2f} m'
            // },


            plotOptions: {
                series: {
                    color: 'rgba(255,165,0, 1.0)',
                    marker: {
                        fillColor: 'black',
                        enabled: true,
                        radius: 2,

                    }
                }
            },

            colors: ['#6CF', '#39F', '#06C', '#036', '#000'],

            // Define the data points. All series have a dummy year
            // of 1970/71 in order to be compared on the same x axis. Note
            // that in JavaScript, months start at 0 for January, 1 for February etc.
            series: series,


            // responsive: {
            //     rules: [{
            //         condition: {
            //             maxWidth: 500
            //         },
            //         chartOptions: {
            //             plotOptions: {
            //                 series: {
            //                     marker: {
            //                         radius: 2,
            //                         color: 'black'
            //                     }
            //                 }
            //             }
            //         }
            //     }]
            // }
        })
    }
};
