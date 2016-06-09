$(window).load(function () {
    // run code
    
});









    var chart;
    $(document).ready(function () {

        var data1 = $("#tot_1").val() 
        //47500
        //alert($("#tot_1").val())
        var data2 = $("#tot_2").val()
        var data3 = $("#tot_3").val()
        var data4 = $("#tot_4").val()
        var data5 = $("#tot_5").val()

        chart = new Highcharts.Chart({
            
            

            chart: {
                renderTo: 'container',
                defaultSeriesType: 'column'
            },
            title: {
                text: 'Salg & Værdi total i periode'
            },
            xAxis: {
                categories: ['Tilbud', 'Tilbud * Sandsynlighed', 'Aktive job', 'Lukkede job, faktureret', 'Alle job, faktureret']
            },
            yAxis: {
                
                title: {
                    text: '1000 DKK'
                }
            },

            tooltip: {
                formatter: function () {
                    return '' +
								 this.series.name + ': ' + this.y + '';
                }
            },
            credits: {
                enabled: false
            },
            series: [{
            //name: 'John',
            data: [data1, 3, 4, 7, data5]
            
        }]


        });




    $("#loadbar").hide(1000);


});