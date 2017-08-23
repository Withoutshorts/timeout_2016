







$(document).ready(function () {






    $(".closeqmnote").mouseover(function () {
        $(this).css('cursor', 'pointer');
    });

    $(".closeqmnote").click(function () {
        $(".qmarkhelptxt").hide();
    });




    //Dato func jq //
    function IsValidDate(input) {
        var dayfield = input.split(".")[0];
        var monthfield = input.split(".")[1];
        var yearfield = input.split(".")[2];
        var returnval;
        var dayobj = new Date(yearfield, monthfield - 1, dayfield)
        if ((dayobj.getMonth() + 1 != monthfield) || (dayobj.getDate() != dayfield) || (dayobj.getFullYear() != yearfield)) {
            alert("Forkert datoformat, skriv venligst dd.mm.yyyy eller d.m.yyyy");
            returnval = false;
        }
        else {
            returnval = true;
        }
        return returnval;
    }


    //Make datepickers
    $.datepicker.setDefaults($.extend({ showOn: 'button', constrainInput: true, showOtherMonths: true, showWeeks: true, minDate: new Date(2002, 1, 1), firstDay: 1, changeFirstDay: false, buttonImage: '../ill/popupcal.gif', start: 6, buttonImageOnly: true, dateFormat: 'd.m.yy', changeMonth: true, changeYear: true }));

    $("#FM_stdato").datepicker();
    $("#FM_stdato").change(function () { if (IsValidDate($("#FM_stdato").val()) == true) { var datestring = $(this).val(); var datesplit = datestring.split('.'); $("[name=FM_start_dag]").val(datesplit[0]); $("[name=FM_start_mrd]").val(datesplit[1]); $("[name=FM_start_aar]").val(datesplit[2]); } else { alert("dato ikke gemt"); } });
    $("#FM_stdato").css("display", "inline");

    $("#FM_sldato").datepicker();
    $("#FM_sldato").change(function () { if (IsValidDate($("#FM_sldato").val()) == true) { var datestring = $(this).val(); var datesplit = datestring.split('.'); $("[name=FM_slut_dag]").val(datesplit[0]); $("[name=FM_slut_mrd]").val(datesplit[1]); $("[name=FM_slut_aar]").val(datesplit[2]); } else { alert("dato ikke gemt"); } });
    $("#FM_sldato").css("display", "inline");






    $("#sp_med").click(function () {

        if ($("#tr_prog_med").css('display') == "none") {
            $("#tr_prog_med").css("visibility", "visible")
            $("#tr_prog_med").css("display", "")
            $.cookie('tr_medarb', '1');
        } else {
            $("#tr_prog_med").hide('fast')
            $.cookie('tr_medarb', '0');
        }
    });


    $("#sp_med").mouseover(function () {
        $(this).css('cursor', 'pointer');
    });



    $("#sp_kun").click(function () {

        if ($("#tr_kun").css('display') == "none") {
            $("#tr_kun").css("visibility", "visible")
            $("#tr_kun").css("display", "")
            $.cookie('tr_kun', '1');
        } else {
            $("#tr_kun").hide('fast')
            $.cookie('tr_kun', '0');
        }
    });

    $("#sp_kun").mouseover(function () {
        $(this).css('cursor', 'pointer');
    });


    $("#sp_job").click(function () {

        if ($("#tr_job").css('display') == "none") {
            $("#tr_job").css("visibility", "visible")
            $("#tr_job").css("display", "")
            $.cookie('tr_job', '1');
        } else {
            $("#tr_job").hide('fast')
            $.cookie('tr_job', '0');
        }
    });

    $("#sp_job").mouseover(function () {
        $(this).css('cursor', 'pointer');
    });


    $(".qmarkhelp").mouseover(function () {
        $(this).css('cursor', 'pointer');
    });

    $(".qmarkhelp").click(function () {

        var thisid = this.id

        if ($("#qmarkhelptxt_" + thisid).css('display') == "none") {
            $("#qmarkhelptxt_" + thisid).css("visibility", "visible")
            $("#qmarkhelptxt_" + thisid).css("display", "")
        } else {

            $("#qmarkhelptxt_" + thisid).hide(200)

        }

    });






    $("#sp_ava").click(function () {

        if ($("#tr_ava").css('display') == "none") {
            $("#tr_ava").css("visibility", "visible")
            $("#tr_ava").css("display", "")

            $.cookie('tr_ava', '1');
        } else {
            $("#tr_ava").hide('fast')
            $.cookie('tr_ava', '0');
        }
    });

    $("#sp_ava").mouseover(function () {
        $(this).css('cursor', 'pointer');
    });



    $("#sp_pre").click(function () {

        if ($("#tr_pre").css('display') == "none") {
            $("#tr_pre").css("visibility", "visible")
            $("#tr_pre").css("display", "")

            $.cookie('tr_pre', '1');
        } else {
            $("#tr_pre").hide('fast')
            $.cookie('tr_pre', '0');
        }
    });

    $("#sp_pre").mouseover(function () {
        $(this).css('cursor', 'pointer');
    });



    //alert($.cookie('tr_medarb'))
    if ($.cookie('tr_medarb') == '1') {
        $("#tr_prog_med").css("visibility", "visible")
        $("#tr_prog_med").css("display", "")
    }

    if ($.cookie('tr_kun') == '1') {
        $("#tr_kun").css("visibility", "visible")
        $("#tr_kun").css("display", "")
    }

    if ($.cookie('tr_job') == '1') {
        $("#tr_job").css("visibility", "visible")
        $("#tr_job").css("display", "")
    }

    if ($.cookie('tr_ava') == '1') {
        $("#tr_ava").css("visibility", "visible")
        $("#tr_ava").css("display", "")
    }

    if ($.cookie('tr_pre') == '1') {
        $("#tr_pre").css("visibility", "visible")
        $("#tr_pre").css("display", "")
    }



});