

        $(document).ready(function () {


            
         
            
            $("#FM_igndatoi").click(function () {
                
               

                if ($("#FM_igndatoi").is(':checked') == true) {

                    //alert("her" + $("#basis_listdato").val())
                   
                    var listartDag = "" + $("#basis_listdato_dag").val()
                    var listartMrd = "" + $("#basis_listdato_mrd").val()
                    var listartAar = "" + $("#basis_listdato_aar").val()

                    var listartDato = $("#basis_listdato_dag").val() + "." + $("#basis_listdato_mrd").val() + "." + $("#basis_listdato_aar").val()

                    $("#FM_stdato").val(listartDato)

                    $("#FM_start_dag").val(listartDag)
                    $("#FM_start_mrd").val(listartMrd)
                    $("#FM_start_aar").val(listartAar)

                    $("#FM_stdato").attr('disabled', true);
                    
                    
                } else {


                    $("#FM_stdato").removeAttr('disabled');


                }

            });


            if ($("#FM_igndatoi").is(':checked') == true) {
                
                $("#FM_stdato").attr('disabled', true);

            }

        });





    $(window).load(function () {
        // run code
        //$("#loadbar").hide(1000);
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

        

        $(".a_showalert").click(function () {

            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(12, idlngt)

            $("#sp_showalert_" + idtrim).css('visibility', 'visible')
            $("#sp_showalert_" + idtrim).css('display', '')


        });

        $(".sp_showalert").click(function () {

            $(".sp_showalert").hide("fast")

        });


        $(".a_showalert, .sp_showalert").mouseover(function () {

            $(this).css('cursor', 'pointer');

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


    
