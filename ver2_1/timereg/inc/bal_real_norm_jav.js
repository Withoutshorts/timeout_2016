



    $(window).load(function () {
        // run code
        $("#loadbar").hide(1000);
    });





$(document).ready(function() {

  

        $("#addDag").click(function () {

            $("#addDagVal").val('1')
            $("#periode").trigger("submit")
        });


        $("#addDag_minus").click(function () {

            $("#addDagVal").val('-1')
            $("#periode").trigger("submit")
        });


    $("#rap_afstem").click(function() {
        $(".akt_afst").attr('checked', true);
        $(".akt_udspec").attr('checked', false);
        $(".akt_ferie").attr('checked', false);
        $(".akt_overarb").attr('checked', false);
        $(".akt_syg").attr('checked', false);
        $(".akt_flex").attr('checked', false);
        $(".akt_afst_saldo").attr('checked', false);
    });

    $("#rap_ferie").click(function() {
        $(".akt_afst").attr('checked', false);
        $(".akt_udspec").attr('checked', false);
        $(".akt_ferie").attr('checked', true);
        $(".akt_overarb").attr('checked', false);
        $(".akt_syg").attr('checked', false);
        $(".akt_flex").attr('checked', false);
        $(".akt_afst_saldo").attr('checked', false);
    });

    $("#rap_lon").click(function() {
        $(".akt_afst").attr('checked', false);
        $(".akt_udspec").attr('checked', false);
        $(".akt_ferie").attr('checked', true);
        $(".akt_overarb").attr('checked', true);
        $(".akt_syg").attr('checked', true);
        $(".akt_flex").attr('checked', false);
        $(".akt_afst_saldo").attr('checked', false);
    });

    $("#rap_syg").click(function() {
        $(".akt_afst").attr('checked', false);
        $(".akt_udspec").attr('checked', false);
        $(".akt_ferie").attr('checked', false);
        $(".akt_overarb").attr('checked', false);
        $(".akt_syg").attr('checked', true);
        $(".akt_flex").attr('checked', false);
        $(".akt_afst_saldo").attr('checked', false);
    });

    $("#rap_udspec").click(function() {
        $(".akt_afst").attr('checked', true);
        $(".akt_udspec").attr('checked', true);
        $(".akt_ferie").attr('checked', false);
        $(".akt_overarb").attr('checked', false);
        $(".akt_syg").attr('checked', false);
        $(".akt_flex").attr('checked', true);
        $(".akt_afst_saldo").attr('checked', false);
    });

    $("#rap_all").click(function() {
        $(".akt_afst").attr('checked', true);
        $(".akt_udspec").attr('checked', true);
        $(".akt_ferie").attr('checked', true);
        $(".akt_overarb").attr('checked', true);
        $(".akt_syg").attr('checked', true);
        $(".akt_flex").attr('checked', true);
        $(".akt_afst_saldo").attr('checked', false);
    });

    $(".akt_afst_saldo").click(function () {

        if ($(this).is(':checked') == true) {
            $(".akt_afst").attr('checked', false);
            $(".akt_udspec").attr('checked', false);
            $(".akt_ferie").attr('checked', false);
            $(".akt_overarb").attr('checked', false);
            $(".akt_syg").attr('checked', false);
            $(".akt_flex").attr('checked', false);

        }

    });

    $("#rap_fk_sdlon").click(function () {
        $(".akt_afst").attr('checked', false);
        $(".akt_udspec").attr('checked', false);
        $(".akt_ferie").attr('checked', false);
        $(".akt_overarb").attr('checked', false);
        $(".akt_syg").attr('checked', false);
        $(".akt_flex").attr('checked', false);
        $(".akt_afst_saldo").attr('checked', false);

        $("#FM_akttype_id_3_4").attr('checked', true); //2
        //$("#FM_akttype_id_3_3").attr('checked', true);
        $("#FM_akttype_id_3_7").attr('checked', true); //5
        $("#FM_akttype_id_3_11").attr('checked', true);
        //$("#FM_akttype_id_4_1").attr('checked', true);
        $("#FM_akttype_id_5_0").attr('checked', true);
        $("#FM_akttype_id_5_1").attr('checked', true);
        //$("#FM_akttype_id_5_2").attr('checked', true);
        //$("#FM_akttype_id_5_3").attr('checked', true);
        $("#FM_akttype_id_5_4").attr('checked', true);
        $("#FM_akttype_id_5_5").attr('checked', true);
        $("#FM_akttype_id_5_6").attr('checked', true);
       
    
    });


    $("#rap_epi_blueg").click(function () {

        $(".akt_afst").attr('checked', false);
        $(".akt_udspec").attr('checked', false);
        $(".akt_ferie").attr('checked', false);
        $(".akt_overarb").attr('checked', false);
        $(".akt_syg").attr('checked', false);
        $(".akt_flex").attr('checked', false);
        $(".akt_afst_saldo").attr('checked', false);

        $("#FM_akttype_id_0_1").attr('checked', true);
        //$("#FM_akttype_id_1_1").attr('checked', true);
        //$("#FM_akttype_id_1_2").attr('checked', true);

        $("#FM_akttype_id_1_3").attr('checked', true); //2
        $("#FM_akttype_id_1_4").attr('checked', true); //5
        $("#FM_akttype_id_1_5").attr('checked', true);
        $("#FM_akttype_id_1_7").attr('checked', true);
       
      


    });



    $("#rap_cflow_huldt").click(function () {

        $(".akt_afst").attr('checked', false);
        $(".akt_udspec").attr('checked', false);
        $(".akt_ferie").attr('checked', false);
        $(".akt_overarb").attr('checked', false);
        $(".akt_syg").attr('checked', false);
        $(".akt_flex").attr('checked', false);
        $(".akt_afst_saldo").attr('checked', false);

        $("#FM_akttype_id_4_0").attr('checked', true);
        $("#FM_akttype_id_4_12").attr('checked', true); 
        $("#FM_akttype_id_4_13").attr('checked', true);
        $("#FM_akttype_id_4_71").attr('checked', true);
        $("#FM_akttype_id_4_72").attr('checked', true);
        $("#FM_akttype_id_4_74").attr('checked', true);
        $("#FM_akttype_id_4_75").attr('checked', true);
        $("#FM_akttype_id_4_8").attr('checked', true);
        $("#FM_akttype_id_4_81").attr('checked', true);
        $("#FM_akttype_id_4_9").attr('checked', true);
        $("#FM_akttype_id_4_91").attr('checked', true);
        $("#FM_akttype_id_4_92").attr('checked', true);
        $("#FM_akttype_id_4_98").attr('checked', true);
        $("#FM_akttype_id_4_99").attr('checked', true);
        //$("#FM_akttype_id_4_8").attr('checked', true);




    });


    $("#chkalle_0").click(function() {
        if ($(this).is(':checked') == true) {
            $(".akt_afst").attr('checked', true);
            $(".akt_afst_saldo").attr('checked', false);
        } else {
            $(".akt_afst").attr('checked', false);
        }

    });

    $("#chkalle_1").click(function() {
        if ($(this).is(':checked') == true) {
            $(".akt_udspec").attr('checked', true);
            $(".akt_afst_saldo").attr('checked', false);
        } else {
            $(".akt_udspec").attr('checked', false);
        }

    });

    $("#chkalle_2").click(function() {
        if ($(this).is(':checked') == true) {
            $(".akt_flex").attr('checked', true);
            $(".akt_afst_saldo").attr('checked', false);
        } else {
            $(".akt_flex").attr('checked', false);
        }

    });

    $("#chkalle_3").click(function() {
        if ($(this).is(':checked') == true) {
            $(".akt_ferie").attr('checked', true);
            $(".akt_afst_saldo").attr('checked', false);
        } else {
            $(".akt_ferie").attr('checked', false);
        }

    });

    $("#chkalle_4").click(function() {
        if ($(this).is(':checked') == true) {
            $(".akt_overarb").attr('checked', true);
            $(".akt_afst_saldo").attr('checked', false);
        } else {
            $(".akt_overarb").attr('checked', false);
        }

    });

    $("#chkalle_5").click(function() {
        if ($(this).is(':checked') == true) {
            $(".akt_syg").attr('checked', true);
            $(".akt_afst_saldo").attr('checked', false);
        } else {
            $(".akt_syg").attr('checked', false);
        }

    });


    $(".akt_afst").click(function() {
        if ($(this).is(':checked') == true) {
            $(".akt_afst_saldo").attr('checked', false);
        }

    });

    $(".akt_udspec").click(function () {
        if ($(this).is(':checked') == true) {
            $(".akt_afst_saldo").attr('checked', false);
        }

    });

    $(".akt_ferie").click(function () {
        if ($(this).is(':checked') == true) {
            $(".akt_afst_saldo").attr('checked', false);
        }

    });

    $(".akt_overarb").click(function () {
        if ($(this).is(':checked') == true) {
            $(".akt_afst_saldo").attr('checked', false);
        }

    });

    $(".akt_syg").click(function () {
        if ($(this).is(':checked') == true) {
            $(".akt_afst_saldo").attr('checked', false);
        }

    });


    $(".akt_flex").click(function () {
        if ($(this).is(':checked') == true) {
            $(".akt_afst_saldo").attr('checked', false);
        }

    });


    globalWdt = $("#globalWdt").val()
    $("#tableD").css("width", globalWdt);



    

});










 
  
    function checkAll(val) {


        //FM_akttype_id_1
        //alert(val)
        antal_v = document.getElementById("antal_v_"+val).value
        //alert(antal_v)
    
        if (document.getElementById("chkalle_"+val).checked == true) {
            tval = true
        } else {
            tval = false
        }
        
        //alert(tval)

        for (i = 0; i < antal_v; i++)
            //alert(i)
            document.getElementById("FM_akttype_id_"+val+"_"+i).checked = tval
    }


    
