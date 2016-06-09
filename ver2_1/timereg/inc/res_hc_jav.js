







$(document).ready(function () {


    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    var options = {
        
        chart: {
            renderTo: 'container',
            defaultSeriesType: 'column'

        },
        title: {
            text: 'Ressource forecast pr. medarb. i valgt periode [beta]'
        },
       
        xAxis: {
            categories: [] 
           
        },

      yAxis: {
          title: {
              text: 'Timer'
          }
      },

       //plotOptions: {
       //     series: {
       //         stacking: 'normal'
       //     }
      // },
       

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








    

});

