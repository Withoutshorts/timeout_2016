$(function () {

	var data, chartOptions


	proc_0 = $('#timefordeling_proc_0').val()
	proc_1 = $('#timefordeling_proc_1').val()
	proc_2 = $('#timefordeling_proc_2').val()
	proc_3 = $('#timefordeling_proc_3').val()
	proc_4 = $('#timefordeling_proc_4').val()
	proc_5 = $('#timefordeling_proc_5').val()
	proc_6 = $('#timefordeling_proc_6').val()
	proc_7 = $('#timefordeling_proc_7').val()

	navn_0 = $('#timefordeling_navn_0').val()
	navn_1 = $('#timefordeling_navn_1').val()
	navn_2 = $('#timefordeling_navn_2').val()
	navn_3 = $('#timefordeling_navn_3').val()
	navn_4 = $('#timefordeling_navn_4').val()
	navn_5 = $('#timefordeling_navn_5').val()
	navn_6 = $('#timefordeling_navn_6').val()
	navn_7 = $('#timefordeling_navn_7').val()

	//fravar = $('#timefordeling_proc_fravar').val()
	//norm = $('#timefordeling_proc_norm').val()

	data = [
		//{ label: "Ferie & Fravær", data: Math.floor(fravar) },
		{ label: "" + navn_0 + " " + proc_0 + " %", data: Math.floor(proc_0) },
		{ label: "" + navn_1 + " " + proc_1 + " %", data: Math.floor(proc_1) },
        { label: "" + navn_2 + " " + proc_2 + " %", data: Math.floor(proc_2) },
        { label: "" + navn_3 + " " + proc_3 + " %", data: Math.floor(proc_3) },
        { label: "" + navn_4 + " " + proc_4 + " %", data: Math.floor(proc_4) },
        { label: "" + navn_5 + " " + proc_5 + " %", data: Math.floor(proc_5) },
		{ label: "" + navn_6 + " " + proc_6 + " %", data: Math.floor(proc_6) },
	    { label: "" + navn_7 + " " + proc_7 + " %", data: Math.floor(proc_7) }
        //{ label: "Norm", data: Math.floor(norm) },
	]

	chartOptions = {		
		series: {
			pie: {
				show: true,  
				innerRadius: .5, 
				stroke: {
					width: 4
				}
			}
		}, 
		legend: {
			position: 'se'
		}, 
		tooltip: true,
		tooltipOpts: {
			content: '%s: %y'
		},
		grid: {
			hoverable: true
		},
		colors: mvpready_core.layoutColors
	}


	var holder = $('#donut-chart')

	if (holder.length) {
		$.plot(holder, data, chartOptions )
	}

})