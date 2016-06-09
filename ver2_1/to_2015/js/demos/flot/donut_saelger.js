$(function () {

	var data, chartOptions

	//var util_norm = $('#util_norm').val()
	var util_pre = $('#util_pre').val()
	var util_pre_proc = $('#util_pre_proc').val()

	var util_prod = $('#util_prod').val()
	var util_prod_proc = $('#util_prod_proc').val()

	var util_post = $('#util_post').val()
	var util_post_proc = $('#util_post_proc').val()

	data = [
		{ label: "Pre. " + util_pre_proc + " %", data: Math.floor(util_pre) },
		{ label: "Prod. " + util_prod_proc + " %", data: Math.floor(util_prod) },
		{ label: "Post " + util_post_proc + " %", data: Math.floor(util_post) }
		
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


	var holder = $('#donut-chart')

	if (holder.length) {
		$.plot(holder, data, chartOptions )
	}

})