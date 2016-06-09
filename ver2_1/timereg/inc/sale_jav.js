







$(document).ready(function () {


    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    var options = {
        
        chart: {
            renderTo: 'container',
            defaultSeriesType: 'bar'

        },
        title: {
            text: 'Salg og Værdi aku. pr. medarbejder'
        },
       
        xAxis: {
            categories: [] 
           
        },

      yAxis: {
          title: {
              text: '1000 (cunrrency)'
          }
      },

       plotOptions: {
            series: {
                stacking: 'normal'
            }
        },
       

        credits: {
            enabled: false
            
        },

        
        series: []
    };



    /*
    Load the data from the CSV file. This is the contents of the file:
			 
			
    */
    var filename = $("#FM_filenameA").val()
    $.get('../inc/log/data/' + filename, function (data) {
        // Split the lines
        var lines = data.split('\n');
        $.each(lines, function (lineNo, line) {
            var items = line.split(';');
            //alert(lineNo)

            // header line containes categories
            if (lineNo == 0) {
                $.each(items, function (itemNo, item) {
                    //alert(item)
                	if (itemNo > 0) options.xAxis.categories.push(item);
                });
            }

            // the rest of the lines contain data with their name in the first position
            else {
                var series = {
                    data: []
                };
                $.each(items, function (itemNo, item) {
                    if (itemNo == 0) {
                        series.name = item;
                    } else {
                        //alert(item)
                        series.data.push(parseFloat(item));
                    }
                });

                options.series.push(series);

            }

            

        });


        var chart = new Highcharts.Chart(options);

    });








    //////////////////////////////////////////////////////////////////


    var options2 = {
        chart: {
            renderTo: 'container2',
            defaultSeriesType: 'column'


        },
        title: {
            text: 'Salg og Værdi aku. pr. kolonne'
        },
        //xAxis: {
        //    categories: ['Tilbud', 'Tilbud * Sandsynlighed', 'Aktive job', 'Lukkede job, faktureret', 'Alle job, faktureret']
        //},
        xAxis: {
        	categories: []
        },
        yAxis: {
            title: {
                text: '1000 (currency)'
            }
        },
        legend: {
        //backgroundColor: Highcharts.theme.legendBackgroundColorSolid || '#FFFFFF',
        //reversed: true
    },
    tooltip: {
        formatter: function () {
            return '' +
            this.series.name + ': ' + this.y + '';
        }
    },
    plotOptions: {
        series: {
            stacking: 'normal'
        }
    },
    credits: {
        enabled: false
    },

    series: []
};



/*
Load the data from the CSV file. This is the contents of the file:
			 
			
*/
var filename = $("#FM_filename").val()
$.get('../inc/log/data/' + filename, function (data) {
    // Split the lines
    var lines = data.split('\n');
    $.each(lines, function (lineNo, line) {
        var items = line.split(';');

        // header line containes categories
        if (lineNo == 0) {
            $.each(items, function (itemNo, item) {
                
                //options.xAxis.categories.push(item);
                if (itemNo >= 1) {
                //alert(item)
                options2.xAxis.categories.push(item);
                }
            });
        }

        // the rest of the lines contain data with their name in the first position
        else {
            var series = {
                data: []
            };
            $.each(items, function (itemNo, item) {
                if (itemNo == 0) {
                    series.name = item;
                } else {
                    series.data.push(parseFloat(item));
                }
            });

            options2.series.push(series);

        }

    });

    var chart = new Highcharts.Chart(options2);

});

///////////////////////////////////////////////////////////////////////////////////////////////////////









});
		