$(function () {

	var data, chartOptions


	proc_0 = $('#timefordeling_proc_saelger_0').val()
    proc_1 = $('#timefordeling_proc_saelger_1').val()
    proc_2 = $('#timefordeling_proc_saelger_2').val()
    proc_3 = $('#timefordeling_proc_saelger_3').val()

	navn_0 = $('#timefordeling_navn_saelger_0').val()
	navn_1 = $('#timefordeling_navn_saelger_1').val()
	navn_2 = $('#timefordeling_navn_saelger_2').val()
	navn_3 = $('#timefordeling_navn_saelger_3').val()

	//fravar = $('#timefordeling_proc_saelger_fravar').val()
	//norm = $('#timefordeling_proc_norm').val()

	data = [
		//{ label: "Ferie & Fravær", data: Math.floor(fravar) },
		{ label: "" + navn_0 + " " + proc_0 + " %", data: Math.floor(proc_0) },
		{ label: "" + navn_1 + " " + proc_1 + " %", data: Math.floor(proc_1) },
		{ label: "" + navn_2 + " " + proc_2 + " %", data: Math.floor(proc_2) },
	    { label: "" + navn_3 + " " + proc_3 + " %", data: Math.floor(proc_3) }
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
			position: 'ne'
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


	var holder = $('#donut-chart_saelger')

	if (holder.length) {
		$.plot(holder, data, chartOptions )
	}

})