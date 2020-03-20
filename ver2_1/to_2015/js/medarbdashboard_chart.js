
$(document).ready(function () {

    function round(value, decimals) {
        return Number(Math.round(value + 'e' + decimals) + 'e-' + decimals);
    }

    var fctxt = $("#fc_translated").val() + " ";
    var actxt = $("#ac_translated").val() + " ";

    var fomrnavn1 = $("#fomrnavn_0").val();
    var fomrnavn2 = $("#fomrnavn_1").val();
    var fomrnavn3 = $("#fomrnavn_2").val();
    var fomrnavn4 = $("#fomrnavn_3").val();
    var fomrnavn5 = $("#fomrnavn_4").val();
    var fomrnavn6 = $("#fomrnavn_5").val();
    var fomrnavn7 = $("#fomrnavn_6").val();
    var fomrnavn8 = $("#fomrnavn_7").val();
    var fomrnavn9 = $("#fomrnavn_8").val();
    var fomrnavn10 = $("#fomrnavn_9").val();
    var fomrnavn11 = $("#fomrnavn_10").val();
    var fomrnavn12 = $("#fomrnavn_11").val();
    var fomrnavn13 = $("#fomrnavn_12").val();
    var fomrnavn14 = $("#fomrnavn_13").val();
    var fomrnavn15 = $("#fomrnavn_14").val();

    var color1 = "yellow"
    var color2 = "#ead700"
    var color3 = "orange"
    var color4 = "deepskyblue"
    var color5 = "cornflowerblue"
    var color6 = "dodgerblue"
    var color7 = "cadetblue"
    var color8 = "greenyellow"
    var color9 = "lawngreen"
    var color10 = "grey"
    var color11 = "green"
    var color12 = "red"
    var color13 = "blue"
    var color14 = "purple"
    var color15 = "#f48042"

    var fomrprocent1 = $("#fomrregprocent_0").val(); fomrprocent1 = fomrprocent1.replace(",", "."); fomrprocent1 = parseFloat(fomrprocent1)
    var fomrprocent2 = $("#fomrregprocent_1").val(); fomrprocent2 = fomrprocent2.replace(",", "."); fomrprocent2 = parseFloat(fomrprocent2) 
    var fomrprocent3 = $("#fomrregprocent_2").val(); fomrprocent3 = fomrprocent3.replace(",", "."); fomrprocent3 = parseFloat(fomrprocent3) 
    var fomrprocent4 = $("#fomrregprocent_3").val(); fomrprocent4 = fomrprocent4.replace(",", "."); fomrprocent4 = parseFloat(fomrprocent4) 
    var fomrprocent5 = $("#fomrregprocent_4").val(); fomrprocent5 = fomrprocent5.replace(",", "."); fomrprocent5 = parseFloat(fomrprocent5)
    var fomrprocent6 = $("#fomrregprocent_5").val(); fomrprocent6 = fomrprocent6.replace(",", "."); fomrprocent6 = parseFloat(fomrprocent6)
    var fomrprocent7 = $("#fomrregprocent_6").val(); fomrprocent7 = fomrprocent7.replace(",", "."); fomrprocent7 = parseFloat(fomrprocent7) 
    var fomrprocent8 = $("#fomrregprocent_7").val(); fomrprocent8 = fomrprocent8.replace(",", "."); fomrprocent8 = parseFloat(fomrprocent8) 
    var fomrprocent9 = $("#fomrregprocent_8").val(); fomrprocent9 = fomrprocent9.replace(",", "."); fomrprocent9 = parseFloat(fomrprocent9) 
    var fomrprocent10 = $("#fomrregprocent_9").val(); fomrprocent10 = fomrprocent10.replace(",", "."); fomrprocent10 = parseFloat(fomrprocent10) 
    var fomrprocent11 = $("#fomrregprocent_10").val(); fomrprocent11 = fomrprocent11.replace(",", "."); fomrprocent11 = parseFloat(fomrprocent11) 
    var fomrprocent12 = $("#fomrregprocent_11").val(); fomrprocent12 = fomrprocent12.replace(",", "."); fomrprocent12 = parseFloat(fomrprocent12) 
    var fomrprocent13 = $("#fomrregprocent_12").val(); fomrprocent13 = fomrprocent13.replace(",", "."); fomrprocent13 = parseFloat(fomrprocent13) 
    var fomrprocent14 = $("#fomrregprocent_13").val(); fomrprocent14 = fomrprocent14.replace(",", "."); fomrprocent14 = parseFloat(fomrprocent14) 
    var fomrprocent15 = $("#fomrregprocent_14").val(); fomrprocent15 = fomrprocent15.replace(",", "."); fomrprocent15 = parseFloat(fomrprocent15) 

    var fomrFCprocent1 = $("#fomrfcprocent_0").val(); fomrFCprocent1 = fomrFCprocent1.replace(",", "."); fomrFCprocent1 = parseFloat(fomrFCprocent1)
    var fomrFCprocent2 = $("#fomrfcprocent_1").val(); fomrFCprocent2 = fomrFCprocent2.replace(",", "."); fomrFCprocent2 = parseFloat(fomrFCprocent2)
    var fomrFCprocent3 = $("#fomrfcprocent_2").val(); fomrFCprocent3 = fomrFCprocent3.replace(",", "."); fomrFCprocent3 = parseFloat(fomrFCprocent3)
    var fomrFCprocent4 = $("#fomrfcprocent_3").val(); fomrFCprocent4 = fomrFCprocent4.replace(",", "."); fomrFCprocent4 = parseFloat(fomrFCprocent4)
    var fomrFCprocent5 = $("#fomrfcprocent_4").val(); fomrFCprocent5 = fomrFCprocent5.replace(",", "."); fomrFCprocent5 = parseFloat(fomrFCprocent5)
    var fomrFCprocent6 = $("#fomrfcprocent_5").val(); fomrFCprocent6 = fomrFCprocent6.replace(",", "."); fomrFCprocent6 = parseFloat(fomrFCprocent6)
    var fomrFCprocent7 = $("#fomrfcprocent_6").val(); fomrFCprocent7 = fomrFCprocent7.replace(",", "."); fomrFCprocent7 = parseFloat(fomrFCprocent7)
    var fomrFCprocent8 = $("#fomrfcprocent_7").val(); fomrFCprocent8 = fomrFCprocent8.replace(",", "."); fomrFCprocent8 = parseFloat(fomrFCprocent8)
    var fomrFCprocent9 = $("#fomrfcprocent_8").val(); fomrFCprocent9 = fomrFCprocent9.replace(",", "."); fomrFCprocent9 = parseFloat(fomrFCprocent9)

    var fomrFCprocent10 = $("#fomrfcprocent_9").val(); fomrFCprocent10 = fomrFCprocent10.replace(",", "."); fomrFCprocent10 = parseFloat(fomrFCprocent10)
    
    var fomrFCprocent11 = $("#fomrfcprocent_10").val(); fomrFCprocent11 = fomrFCprocent11.replace(",", "."); fomrFCprocent11 = parseFloat(fomrFCprocent11)
    var fomrFCprocent12 = $("#fomrfcprocent_11").val(); fomrFCprocent12 = fomrFCprocent12.replace(",", "."); fomrFCprocent12 = parseFloat(fomrFCprocent12)
    var fomrFCprocent13 = $("#fomrfcprocent_12").val(); fomrFCprocent13 = fomrFCprocent13.replace(",", "."); fomrFCprocent13 = parseFloat(fomrFCprocent13)
    var fomrFCprocent14 = $("#fomrfcprocent_13").val(); fomrFCprocent14 = fomrFCprocent14.replace(",", "."); fomrFCprocent14 = parseFloat(fomrFCprocent14)
    var fomrFCprocent15 = $("#fomrfcprocent_14").val(); fomrFCprocent15 = fomrFCprocent15.replace(",", "."); fomrFCprocent15 = parseFloat(fomrFCprocent15)
   
    var chart = new CanvasJS.Chart("totalChart", {
        /* title: {
             text: medarbnavn
          }, */
        axisY: {
            title: "Tid i procent %"
        },
        axisY: {
            maximum: 200,
            suffix: " %"
        },
        toolTip: {
            content: "{label} <br/>{name} : {y} % af norm"
        },
        data: [
            {
                type: "stackedColumn",
                color: color1,
                name: fomrnavn1,
                showInLegend: false,
                dataPoints: [
                    { label: fctxt, y: round(fomrFCprocent1,2) },
                    { label: actxt, y: round(fomrprocent1,2) },
                ]
            },
            {
                type: "stackedColumn",
                color: color2,
                name: fomrnavn2,
                showInLegend: false,
                dataPoints: [
                    { label: fctxt, y: round(fomrFCprocent2,2) },
                    { label: actxt, y: round(fomrprocent2,2) }
                ]
            },
            {
                type: "stackedColumn",
                color: color3,
                name: fomrnavn3,
                showInLegend: false,
                dataPoints: [
                    { label: fctxt, y: round(fomrFCprocent3,2) },
                    { label: actxt, y: round(fomrprocent3,2) }
                ]
            },
            {
                type: "stackedColumn",
                color: color4,
                name: fomrnavn4,
                showInLegend: false,
                dataPoints: [
                    { label: fctxt, y: round(fomrFCprocent4,2) },
                    { label: actxt, y: round(fomrprocent4,2) }
                ]
            },
            {
                type: "stackedColumn",
                color: color5,
                name: fomrnavn5,
                showInLegend: false,
                dataPoints: [
                    { label: fctxt, y: round(fomrFCprocent5,2) },
                    { label: actxt, y: round(fomrprocent5,2) }
                ]
            },
            {
                type: "stackedColumn",
                color: color6,
                name: fomrnavn6,
                showInLegend: false,
                dataPoints: [
                    { label: fctxt, y: round(fomrFCprocent6,2) },
                    { label: actxt, y: round(fomrprocent6,2) }
                ]
            },
            {
                type: "stackedColumn",
                color: color7,
                name: fomrnavn7,
                showInLegend: false,
                dataPoints: [
                    { label: fctxt, y: round(fomrFCprocent7,2) },
                    { label: actxt, y: round(fomrprocent7,2) }
                ]
            },
            {
                type: "stackedColumn",
                color: color8,
                name: fomrnavn8,
                showInLegend: false,
                dataPoints: [
                    { label: fctxt, y: round(fomrFCprocent8,2) },
                    { label: actxt, y: round(fomrprocent8,2) }
                ]
            },
            {
                type: "stackedColumn",
                color: color9,
                name: fomrnavn9,
                showInLegend: false,
                dataPoints: [
                    { label: fctxt, y: round(fomrFCprocent9,2) },
                    { label: actxt, y: round(fomrprocent9,2) }
                ]
            },
            {
                type: "stackedColumn",
                color: color10,
                name: fomrnavn10,
                showInLegend: false,
                dataPoints: [
                    { label: fctxt, y: round(fomrFCprocent10,2) },
                    { label: actxt, y: round(fomrprocent10,2) }
                ]
            },
            {
                type: "stackedColumn",
                color: color11,
                name: fomrnavn11,
                showInLegend: false,
                dataPoints: [
                    { label: fctxt, y: round(fomrFCprocent11,2) },
                    { label: actxt, y: round(fomrprocent11,2) }
                ]
            },
            {
                type: "stackedColumn",
                color: color12,
                name: fomrnavn12,
                showInLegend: false,
                dataPoints: [
                    { label: fctxt, y: round(fomrFCprocent12,2) },
                    { label: actxt, y: round(fomrprocent12,2) }
                ]
            },
            {
                type: "stackedColumn",
                color: color13,
                name: fomrnavn13,
                showInLegend: false,
                dataPoints: [
                    { label: fctxt, y: round(fomrFCprocent13,2) },
                    { label: actxt, y: round(fomrprocent13,2) }
                ]
            },
            {
                type: "stackedColumn",
                color: color14,
                name: fomrnavn14,
                showInLegend: false,
                dataPoints: [
                    { label: fctxt, y: round(fomrFCprocent14,2) },
                    { label: actxt, y: round(fomrprocent14,2) }
                ]
            },
            {
                type: "stackedColumn",
                color: color15,
                name: fomrnavn15,
                showInLegend: false,
                dataPoints: [
                    { label: fctxt, y: round(fomrFCprocent15,2) },
                    { label: actxt, y: round(fomrprocent15,2) }
                ]
            }
        ]
    });
    chart.render();


});