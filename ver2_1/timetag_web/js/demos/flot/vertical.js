$(function () {

    var d1, d2, d3, data, chartOptions;


    proc_0 = $('#timefordeling_proc_0').val()
    proc_1 = $('#timefordeling_proc_1').val()
    proc_2 = $('#timefordeling_proc_2').val()

    navn_0 = $('#timefordeling_navn_0').val()
    navn_1 = $('#timefordeling_navn_1').val()
    navn_2 = $('#timefordeling_navn_2').val()

	d1 = [
        [1325376000000, 500], [1328054400000, 700], [1330560000000, 1000], [1333238400000, 600],
        [1335830400000, 350]
    ];
 
    d2 = [
        [1325376000000, 100], [1328054400000, 600], [1330560000000, 300], [1333238400000, 350],
        [1335830400000, 300]
    ];
 
   

    data = [{
        label: navn_0,
    	data: d1
    }, {
        label: navn_1,
    	data: d2
    
    }];

    chartOptions = {
        xaxis: {
            min: (new Date(2011, 11, 15)).getTime(),
            max: (new Date(2012, 04, 18)).getTime(),
            mode: "time",
            tickSize: [2, "month"],
            monthNames: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
            tickLength: 0
        },
        grid: {
            hoverable: true,
            clickable: false,
            borderWidth: 0
        },
        bars: {
	    	show: true,
	    	barWidth: 12*24*60*60*300,
            fill: true,
            lineWidth: 1,
            order: true,
            lineWidth: 0,
            fillColor: { colors: [ { opacity: 1 }, { opacity: 1 } ] }
	    },
        
        tooltip: true,
        tooltipOpts: {
            content: '%s: %y'
        },
        colors: mvpready_core.layoutColors
    }


    var holder = $('#vertical-chart');

    if (holder.length) {
        $.plot(holder, data, chartOptions );
    }


});