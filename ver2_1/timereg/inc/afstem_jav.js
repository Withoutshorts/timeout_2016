

function showgt() {


    //alert("her")

    var media = $("#media").val()
    var jqusemrn = $("#jqusemrn").val()
    var jqstartdato = $("#jqstartdato").val()
    var jqslutdato = $("#jqslutdato").val()
    var jqlic_mansat_dt = $("#jqlic_mansat_dt").val()
    var jqlevel = $("#jqlevel").val()
    var jqstartdatoDK = new Date()

    if ($.cookie('showgt') == '1') {
        vis_gt_print = 1
    } else {
        vis_gt_print = 0
    }


    if (vis_gt_print == 1) {

        $("#div_gt").css("visibility", "visible")
        $("#div_gt").css("display", "")

        $("#dv_udspec").css("top", "140px")

        //alert($("#medarbafstem_aar").css("height"))

        $("#medarbafstem_aar").css("height", "1000px")
   

    } else {

        $("#div_gt").hide(100)
        $("#dv_udspec").css("top", "30px")
        $("#medarbafstem_aar").css("height", "820px")

    }

    //alert(m_names[jqstartdatoDK.getMonth()])
    //jqstartdatoDK = $("#jqstartdatoDK").val()
    //var mthThis = m_names[jqstartdatoDK.getMonth()] + " " + jqstartdatoDK.getFullYear()
    var show = $("#show").val()

    //var mthThis = new Date()
    //mthThis = jqstartdatoDK.getMonth()
    //alert(mthThis)
    //mthThis = mthThis.getMonth()

    if ((show == 12 || show == 77 || show == 7) && vis_gt_print == 1 && media != 'export') {
        $.post("?jqusemrn=" + jqusemrn + "&jqstartdato=" + jqstartdato + "&jqslutdato=" + jqslutdato + "&jqlic_mansat_dt=" + jqlic_mansat_dt + "&jqlevel=" + jqlevel, { control: "grandtotal", AjaxUpdateField: "true" }, function (data) {

            $("#div_gt").html(data);

            if (media == "print") {
                $("#gtdv").css("top", "-200px")
                window.print();
            }

        });

    }


}



$(window).load(function () {
    // run code




    $("#loadbar").hide(1000);

    //showgt();

    var media = $("#media").val()

    //if (show == 4 && media != "export") {


        if (media == "print") {


            window.print();
        }


    //}



});











$(document).ready(function () {



    $("#FM_visallemedarb").click(function () {


        var sesid = $("#FM_sesMid").val()
        $("#usemrn").val(sesid)

        $("#filterkri").submit();


    });

    $("#usemrn").change(function () {

        $("#filterkri").submit();
    });



    $("#showaf").click(function () {
        $.cookie('showakt', '2');
    });


    $("#showst").click(function () {
        $.cookie('showakt', '3');
    });

    $("#showtim").click(function () {
        $.cookie('showakt', '1');
    });


    globalWdt = $("#globalWdt").val()
    //alert(globalWdt)
    $("#medarbafstem_aar").css("width", globalWdt + "px");

    //$("#vis_gt_print").click(function () {

        //alert("her")
    //if ($("#vis_gt_print").is(':checked') == true) {
    //       $.cookie('showgt', '0');
    //  } else {
    //      $.cookie('showgt', '1');
    //   }

        ////showgt();


    //});


    //if ($.cookie('showgt') == '1') {
    //    $("#vis_gt_print").removeAttr("checked");
    //} else {
    //    $("#vis_gt_print").attr("checked", "checked");
    //}

   
    

    
        setTimeout("clwin()", 3000); //millisekunder (er sat til 1 time)
    
        
   

});


function clwin() {
    $("#gt_msg").hide(1000);
}