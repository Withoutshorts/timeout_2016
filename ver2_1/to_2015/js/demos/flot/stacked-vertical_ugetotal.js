$(function () {
 
    var d1, chartOptions, d2

    var felt1 = $('#timerdagman').val()
    var felt2 = $('#timerdagtir').val()
    var felt3 = $('#timerdagons').val()
    var felt4 = $('#timerdagtor').val()
    var felt5 = $('#timerdagfre').val()
    var felt6 = $('#timerdaglor').val()
    var felt7 = $('#timerdagson').val()

    var felt8 = $('#normdagman').val()
    var felt9 = $('#normdagtir').val()
    var felt10 = $('#normdagons').val()
    var felt11 = $('#normdagtor').val()
    var felt12 = $('#normdagfre').val()
    var felt13 = $('#normdaglor').val()
    var felt14 = $('#normdagson').val()

   
  d1 = [
    [1325376000000, felt1], [1328054400000, felt2], [1330560000000, felt3], [1333238400000, felt4],
    [1335830400000, felt5], [1338336000000, felt6], [1340841600000, felt7]
  ]

  d2 = [
    [1325376000002, felt8], [1328054400002, felt9], [1330560000002, felt10], [1333238400002, felt11],
    [1335830400002, felt12], [1338336000002, felt13], [1340841600002, felt14]
  ]

 
 
  lang = $('#lcid_sprog').val()

  if (lang == '2057') {
      mondayTxt = "Monday"
      tuesdayTxt = "Tuesday"
      wednedsdayTxt = "Wed.day"
      thursdayTxt = "Thursday"
      fridayTxt = "Friday"
      saturdayTxt = "Sat.day"
      sundayTxt = "Sunday"

  } else {

      mondayTxt = "Mandag"
      tuesdayTxt = "Tirsdag"
      wednedsdayTxt = "Onsdag"
      thursdayTxt = "Torsdag"
      fridayTxt = "Fredag"
      saturdayTxt = "Lørdag"
      sundayTxt = "Søndag"

  }
  

  //var data = [{
    //  data: d1, d2
    //}]


  var data = [
      { label: "Norm.", data: d2, color: "#CCCCCC"},
      { label: "Real.", data: d1, color: "#207800"}
      
  ];

  chartOptions = {
    xaxis: {
      min: (new Date(2011, 11, 15)).getTime(),
      max: (new Date(2012, 06, 18)).getTime(),
      mode: "time",
      tickSize: [1, "month"],
      monthNames: ["<h6>" + mondayTxt + "</h6>", "<h6>" + tuesdayTxt + "</h6>", "<h6>" + wednedsdayTxt + "</h6>", "<h6>" + thursdayTxt + "</h6>", "<h6>" + fridayTxt + "</h6>", "<h6>" + saturdayTxt + "</h6>", "<h6>" + sundayTxt + "</h6>"],
      tickLength: 0
    },
    grid: {
      hoverable: false,
      clickable: false,
      borderWidth: 0
    },
    series: {
        stack: true

    },
    bars:
      { show: true, barWidth: 36 * 24 * 60 * 60 * 500, align: 'center', lineWidth: 0, fill: 1 }
    
      //{ show: true, barWidth: 36 * 24 * 60 * 60 * 500, fill: true, align: 'center', lineWidth: 1, lineWidth: 0, fillColor: { colors: [{ opacity: 1 }, { opacity: 1 }] } }
      ,
    tooltip: false
    //, tooltipOpts: {
      //  content: 'timer: 1'
      //}
    //, colors: mvpready_core.layoutColors
    //colors: ["#207800", "#CCCCCC"]
  }



  var holder = $('#stacked-vertical-chart')
  
  if (holder.length) {
    $.plot(holder, data, chartOptions )
  }



})