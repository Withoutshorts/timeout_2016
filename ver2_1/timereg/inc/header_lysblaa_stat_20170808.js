







$(document).ready(function () {


    //alert("HER 1")


    $("#FM_progrp").change(function () {

        //alert("HER")

        vismedarb();

    });

    $("#FM_vis_medarbejdertyper, #FM_vis_medarbejdertyper_grp").click(function () {



        vismedarb();

    });


    $("#FM_visdeakmed").click(function () {



        if ($("#FM_visdeakmed").is(':checked') == false) {
            $("#FM_visdeakmed12").removeAttr("checked");
            $("#FM_visdeakmed12").attr("disabled", "disabled");

        } else {

            $("#FM_visdeakmed12").attr("disabled", "");
        }


        vismedarb();

    });


    $("#FM_visdeakmed12").click(function () {


        vismedarb();

    });

    $("#FM_vispasmed").click(function () {


        vismedarb();

    });


    $("#FM_medarb").click(function () {

        

        if ($("#FM_medarb option:selected").val() == 0) {

            $("#FM_medarb").each(function () {

                //alert("HER")
                $("#FM_medarb option").prop('selected', true);
                

            });


            $("#FM_medarb option[value='0']").removeAttr("selected");

        }

    });




    function vismedarb() {

        $('#FM_medarb').children().remove();

        jq_medarb = $("#jq_userid").val();


        jq_progrp = $("#FM_progrp").val();

        //alert($("#FM_visdeakmed").is(':checked'))


        if ($("#FM_visdeakmed").is(':checked') == true) {
            jq_mansat = 1
        } else {
            jq_mansat = 0
        }


        if ($("#FM_visdeakmed12").is(':checked') == true) {
            jq_mansat12 = 1
        } else {
            jq_mansat12 = 0
        }


        if ($("#FM_vispasmed").is(':checked') == true) {
            jq_mansatpas = 1
        } else {
            jq_mansatpas = 0
        }


        if ($("#FM_vis_medarbejdertyper").is(':checked') == true) {
            jq_mtyper = 1
        } else {
            jq_mtyper = 0
        }

        if ($("#FM_vis_medarbejdertyper_grp").is(':checked') == true) {
            jq_mtypergrp = 1
        } else {
            jq_mtypergrp = 0
        }



        $.post("?jq_medarb=" + jq_medarb + "&jq_progrp=" + jq_progrp + "&jq_mansat=" + jq_mansat + "&jq_mansatpas=" + jq_mansatpas + "&jq_mansat12=" + jq_mansat12 + "&jq_mtyper=" + jq_mtyper + "&jq_mtypergrp=" + jq_mtypergrp, { control: "FN_medipgrp", AjaxUpdateField: "true" }, function (data) {


            $("#FM_medarb").html(data);



            antalM = $('#FM_medarb option').size();
            antalM = antalM - 1
            $('#antalmedarblist').html(antalM)

            //$("#FM_medarb option[1]").attr('selected', 'selected');
            $("#FM_medarb option[value='0']").removeAttr("selected");

            //alert("her s")

        });





    }




    

    if ($("#FM_visdeakmed").is(':checked') == false) {
        $("#FM_visdeakmed12").removeAttr("checked");
        $("#FM_visdeakmed12").attr("disabled", "disabled");
    }



    antalM = $('#FM_medarb option').size();
    antalM = antalM - 1
    $('#antalmedarblist').html(antalM)

   




});