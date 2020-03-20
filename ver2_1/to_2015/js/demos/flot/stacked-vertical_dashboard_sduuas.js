 $(function () {
 
    var d1, chartOptions, d2, d3

    var felt1 = 1 //$('#timerdagman').val()
    var felt2 = 2 //$('#timerdagtir').val()
    var felt3 = 3 //$('#timerdagons').val()
    var felt4 = 4 //$('#timerdagtor').val()
    var felt5 = 5 //$('#timerdagfre').val()
    var felt6 = 6 //$('#timerdaglor').val()
    var felt7 = 7 //$('#timerdagson').val()

    var felt8 = 4 //$('#normdagman').val()
    var felt9 = 3 //$('#normdagtir').val()
    var felt10 = 3 //$('#normdagons').val()
    var felt11 = 2 //$('#normdagtor').val()
    var felt12 = 1 //$('#normdagfre').val()
    var felt13 = 5 // $('#normdaglor').val()
    var felt14 =  0 //$('#normdagson').val()

   
  d1 = [
    [1325376000000, felt1], [1328054400000, felt2], [1330560000000, felt3], [1333238400000, felt4],
    [1335830400000, felt5], [1338336000000, felt6], [1340841600000, felt7]
  ]

  d2 = [
    [1325376000002, felt8], [1328054400002, felt9], [1330560000002, felt10], [1333238400002, felt11],
    [1335830400002, felt12], [1338336000002, felt13], [1340841600002, felt14]
  ]


   d3 = [
         [1325376000003, felt1], [1328054400003, felt9], [1330560000003, felt7], [1333238400003, felt2],
         [1335830400003, felt3], [1338336000003, felt5], [1340841600003, felt14]
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

      mondayTxt = "Hans Jensen"
      tuesdayTxt = "Inge Hansen"
      wednedsdayTxt = "Torben Forsker"
      thursdayTxt = "TimeOut Support"
      fridayTxt = " - "
      saturdayTxt = " - "
      sundayTxt = " - "

  }
  

  //var data = [{
    //  data: d1, d2
    //}]


  var data = [
      { label: "Norm.", data: d2, color: "#CCCCCC"},
      { label: "Real.", data: d1, color: "#207800" },
      { label: "Fomr", data: d3, color: "#FF9911" }
      
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