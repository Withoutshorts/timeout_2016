// JScript File
$(window).load(function () {
    // run code


    $("#loadbar").hide(1000);
});






$(document).ready(function () {

    /*
    $('.fileupload').click(function (e) {
        $(this).find('input[type="file"]').click();
    });

    $('#myInputText').click(function () {
        $('#myInputText').focus();
        $('#myInputText').select();
    });
    */

   

    $("#sp_fordeltb").click(function () {

        if ($("#dv_fordeltb").css('display') == "none") {

            $("#dv_fordeltb").css("visibility", "visible");
            $("#dv_fordeltb").css("display", "");
            $("#dv_fordeltb").show(2000)

        
        } else {


            $("#dv_fordeltb").css("visibility", "hidden");
            $("#dv_fordeltb").css("display", "none");
            $("#dv_fordeltb").hide(1000)

        }

    });



    $("#sp_fordeltb").mouseover(function () {
        $(this).css('cursor', 'pointer');
    });




    $("#FM_fomr").click(function () {

        chkThis = $("#FM_fomr").val();

        $(".FM_fomr_konto").attr('checked', false);
        $("#FM_fomr_konto_" + chkThis).attr('checked', true);
        

    });




    $("#FM_jobans_proc_1").keyup(function () {

        jobandelmaksval(1)

    });

    $("#FM_jobans_proc_2").keyup(function () {

        jobandelmaksval(2)

    });

    $("#FM_jobans_proc_3").keyup(function () {

        jobandelmaksval(3)

    });

    $("#FM_jobans_proc_4").keyup(function () {

        jobsandelmaksval(4)

    });

    $("#FM_jobans_proc_5").keyup(function () {

        jobandelmaksval(5)

    });


    $("#FM_salgsans_proc_1").keyup(function () {

       
        salgsandelmaksval(1)

    });

    $("#FM_salgsans_proc_2").keyup(function () {

        salgsandelmaksval(2)

    });

    $("#FM_salgsans_proc_3").keyup(function () {

        salgsandelmaksval(3)

    });

    $("#FM_salgsans_proc_4").keyup(function () {

        salgsandelmaksval(4)

    });

    $("#FM_salgsans_proc_5").keyup(function () {

        salgsandelmaksval(5)

    });


    function jobandelmaksval(id) {

        var totalVal = 0;
        totalVal = $("#FM_jobans_proc_1").val().replace(",", ".") / 1 + $("#FM_jobans_proc_2").val().replace(",", ".") / 1 + $("#FM_jobans_proc_3").val().replace(",", ".") / 1 + $("#FM_jobans_proc_4").val().replace(",", ".") / 1 + $("#FM_jobans_proc_5").val().replace(",", ".") / 1

        //alert(totalVal / 1)
        if (totalVal / 1 >= 101) {
            alert("Den samlede % må ikke overstige 100")
            $("#FM_jobans_proc_" + id).val('0')
        }

    }

    function salgsandelmaksval(id) {

      
        var totalVal = 0;
        totalVal = $("#FM_salgsans_proc_1").val().replace(",", ".") / 1 + $("#FM_salgsans_proc_2").val().replace(",", ".") / 1 + $("#FM_salgsans_proc_3").val().replace(",", ".") / 1 + $("#FM_salgsans_proc_4").val().replace(",", ".") / 1 + $("#FM_salgsans_proc_5").val().replace(",", ".") / 1

        //alert(totalVal / 1)
        if (totalVal / 1 >= 101) {
            alert("Den samlede % må ikke overstige 100")
            $("#FM_salgsans_proc_" + id).val('0')
        }

    }

    

    $("#FM_opdmedarbtimepriser").click(function () {
        if ($("#FM_opdmedarbtimepriser").is(':checked') == true) {
            alert("Bemærk at alle eksisterende timer, på aktiviteterne på dette job, bliver opdateret med den nye timepris.")
        }
    });

    $("#FM_usetilbudsnr").click(function () {
        if ($("#FM_usetilbudsnr").is(':checked') == true) {
            $("#FM_tnr").css("color", "#000000");
            nexttbtVal = $("#FM_nexttnr").val();
            $("#FM_tnr").val(nexttbtVal);
            $("#FM_tnr").focus();
            $("#FM_status").val('3');

        } else {
            $("#FM_tnr").css("color", "#999999");
            $("#FM_status").val('1');
        }
    });


    $("#FM_status").change(function () {
        if ($("#FM_status").val() == "1") {
            $("#FM_tnr").css("color", "#999999");
            $("#FM_usetilbudsnr").removeAttr("checked");

        }


        if ($("#FM_status").val() == "2") {
            $("#FM_tnr").css("color", "#999999");
            $("#FM_usetilbudsnr").removeAttr("checked");

        }

        if ($("#FM_status").val() == "0") {
            $("#FM_tnr").css("color", "#999999");
            $("#FM_usetilbudsnr").removeAttr("checked");

        }

        if ($("#FM_status").val() == "4") {
            $("#FM_tnr").css("color", "#999999");
            $("#FM_usetilbudsnr").removeAttr("checked");

        }


        if ($("#FM_status").val() == "3") {
            $("#FM_tnr").css("color", "#000000");
            $("#FM_usetilbudsnr").attr("checked", "checked");
        }




    });

    $("#FM_tnr").focus(function () {
        $("#FM_tnr").css("color", "#000000");
    });


    if ($("#sync_budget_mt").val() == "1") {

        $(".mt_faktor").css("border", "0px");
        $(".mt_faktor").css("color", "#999999");

        $(".mt_timer").css("border", "0px");
        $(".mt_timer").css("color", "#999999");


    }





    /////////////////////////////////////////////////////////////////////////////////////////////////////////
    /////////////////// Budget Medarbtyper /////////////////////////////////////////////////////////////////
    ///////////////////////////////////////////////////////////////////////////////////////////////////////

    
      //$(".ulev").keyup(function () {
    //    belobtilMtypGrp();
    //});


    

   


    $("#FM_budget").keyup(function () {

        //$("#mtyp_msg1").css("visibility", "visible");
        //$("#mtyp_msg1").css("display", "");
        //$("#mtyp_msg1").show(2000)
        
        if ($("#sync_budget_mt").val() == '1') {

            belobtilMtypGrp();
        }


        nytBel = $("#FM_budget").val().replace(".",",")
        $("#SP_budget").html(nytBel)

        settotalbelob();
        


    });

    $(".mt_belob_totgrp").keyup(function () {

        var thisid = this.id
        var idlngt = thisid.length
        var idtrim = thisid.slice(19, idlngt)
        var grpid = idtrim

       
        if ($("#FM_laasbelob_" + grpid).is(':checked') == true) {


          
            genberegnlinjer(grpid);

        }

        belobtilMtypGrp();

    });

    

   
    //$(".FM_laasbelob").click(function () {
        //    alert("ger")
       
    //    var thisid = this.id
    //    var idlngt = thisid.length
    //    var idtrim = thisid.slice(13, idlngt)
    //    var grpid = idtrim
    //    genberegnlinjer(grpid);

    //});
    
   


    function genberegnlinjer(grpid) {

       

        stno = $("#FM_mtype_startno_" + grpid).val()
        slno = $("#FM_mtype_slutno_" + grpid).val()

        var nylsst = stno / 1;
        var nylssl = slno / 1;

        //alert(grpid +"_"+nylsst)

        for (g = nylsst; g < nylssl; g++) {
            nxtid = $("#FM_mtype_ids_" + g).val()
            metypeBelobprlinje(grpid + "_" + nxtid);
        }

    }


    
  


    $("#lukms1").mouseover(function () {
        $(this).css('cursor', 'pointer');
    });

    $("#lukms2").mouseover(function () {
        $(this).css('cursor', 'pointer');
    });

    
    $("#lukms1").click(function () {
        $("#mtyp_msg1").hide('fast')
    });

    $("#lukms2").click(function () {
        $("#mtyp_msg2").hide('fast')
    });


    $(".ulev").keyup(function () {
        if ($("#sync_budget_mt").val() == "1") {
            belobtilMtypGrp();
        }
    });



    // Ved ændring af beløb på gruppe 
    $(".mt_belob_totgrp").keyup(function () {


        var thisid = this.id
        var idlngt = thisid.length
        var idtrim = thisid.slice(19, idlngt)
        var grpid = idtrim

        $("#FM_laasbelob_" + grpid).attr("checked", "checked");

        $(".mt_belob_totgrp").css("background-color", "#FFFFFF");


        //$("#mtyp_msg1").hide(2000)

        //$("#mtyp_msg2").css("visibility", "visible");
        //$("#mtyp_msg2").css("display", "");
        //$("#mtyp_msg2").show(2000)

        //Sæter frave på timefelter //
        //stno = $("#FM_mtype_startno_" + grpid).val()
        //slno = $("#FM_mtype_slutno_" + grpid).val()



        var mt_belob_ff_tot = 0;

        for (g = 0; g < 20; g++) {

            for (i = 0; i < 60; i++) { // Sætter alle time felter gule

                //nxtid = $("#FM_mtype_ids_" + i).val()

                $("#FM_mtype_ti_" + g + "_" + i).css("background-color", "#FFFFE1");

            }
        }


       // setFaktorEqOne(grpid);
       


    });





    /// totalbeløb pr. gruppe til nettoomsætning

    function totprMtypGrp_tilNettoOms() {

        var belob_tot = 0;
        var timer_tot = 0;
        var timepris_tot = 0;
        var faktor_tot = 0;
        //var thisBelobff = 0;
        var belobff_tot = 0;
        var kostpris_tot = 0;
        //alert("her")

        for (i = 1; i < 20; i++) { //Optil 20 grupper


            thisBelobff = $("#FM_mt_belobff_totgrp_" + i).val()
            //alert(thisBelobff)
            thisVal = $("#FM_mt_belob_totgrp_" + i).val()
            thisTimer = $("#FM_mt_timer_totgrp_" + i).val()
            thistimepris = $("#FM_mt_timepris_totgrp_" + i).val()
            thisfaktor = $("#FM_mt_faktor_totgrp_" + i).val()
            thisKostpris = $("#FM_mt_kostbelob_totgrp_" + i).val()

            if (typeof thisVal === 'undefined') {

                belob_tot = belob_tot
                timer_tot = timer_tot
                timepris_tot = timepris_tot
                faktor_tot = faktor_tot
                belobff_tot = belobff_tot
                kostpris_tot = kostpris_tot

            } else {

                //alert("der: " + thisVal)
                thisVal = thisVal.replace(".", "")
                thisVal = thisVal.replace(",", ".")
                thisVal = Math.round(thisVal * 100) / 100
                belob_tot = belob_tot + thisVal


                thisBelobff = thisBelobff.replace(".", "")
                thisBelobff = thisBelobff.replace(",", ".")
                thisBelobff = Math.round(thisBelobff * 100) / 100
                belobff_tot = belobff_tot + thisBelobff
                
                //alert(belobff_tot)

                thisTimer = thisTimer.replace(".", "")
                thisTimer = thisTimer.replace(",", ".")
                thisTimer = Math.round(thisTimer * 100) / 100
                timer_tot = timer_tot + thisTimer


                thistimepris = thistimepris.replace(".", "")
                thistimepris = thistimepris.replace(",", ".")
                thistimepris = Math.round(thistimepris * thisTimer * 100) / 100
                timepris_tot = timepris_tot + thistimepris

                thisfaktor = thisfaktor.replace(".", "")
                thisfaktor = thisfaktor.replace(",", ".")
                thisfaktor = Math.round(thisfaktor * thisTimer * 100) / 100
                faktor_tot = faktor_tot + thisfaktor
             
                thisKostpris = thisKostpris.replace(".", "")
                thisKostpris = thisKostpris.replace(",", ".")
                thisKostpris = Math.round(thisKostpris * 100) / 100
                kostpris_tot = kostpris_tot + thisKostpris
                
                //alert("thisTimer: " + thisTimer + "¤¤" + thisKostpris + "##" + kostpris_tot)

            }
        }


        timer_tot = timer_tot
        faktor_tot = Math.round(faktor_tot / timer_tot * 100) / 100
        timepris_tot = Math.round(timepris_tot / timer_tot * 100) / 100

      

        timer_tot = String(timer_tot)
        timer_tot = timer_tot.replace(".", ",")
        $("#FM_budgettimer").val(timer_tot)

        belob_tot = belob_tot
        belob_tot = String(belob_tot)
        belob_tot = belob_tot.replace(".", ",")

        
        belobff_tot = belobff_tot
        belobff_tot = String(belobff_tot)
        belobff_tot = belobff_tot.replace(".", ",")
      
      

        if (timepris_tot == "NaN") {
            timepris_tot = 0
        } else {
            timepris_tot = String(timepris_tot)
            timepris_tot = timepris_tot.replace(".", ",")
        }

       
        $("#FM_gnsinttpris").val(timepris_tot)


        if (faktor_tot == "NaN") {
            faktor_tot = 1
        } else {
            faktor_tot = String(faktor_tot)
            faktor_tot = faktor_tot.replace(".", ",")
        }


        
        kostpris_tot = kostpris_tot
        kostpris_tot = String(kostpris_tot)
        kostpris_tot = kostpris_tot.replace(".", ",")

        $("#FM_intfaktor").val(faktor_tot)


        $("#FM_interntbelob").val(belob_tot)
        
        //alert(kostpris_tot)
        
        $("#FM_interntomkost").val(kostpris_tot) //belobff_tot

        settotalbelob();

    }






    //// Ved ændring af timeantal eller timepris
    function metypeBelobprlinje(id) {

        //alert("metypeBudget_timepris: " + id)


        $(".mt_belob_totgrp").css("background-color", "#FFFFFF");
        $("#mtyp_msg1").hide(2000)

        var grpid = $("#FM_mtypegrp_id_" + id).val()



        //alert(grpid)

       
        if ($("#FM_laasbelob_" + grpid).is(':checked') == true) {

            totBelGrp = $("#FM_mt_belob_totgrp_" + grpid).val().replace(".", "")
            totBelGrp = totBelGrp.replace(",", ".")
            totBelGrp = Math.round(totBelGrp * 100) / 100

            //alert(totBelGrp)

        } else {
            

            totBelGrp = 0

        }
        // $("#FM_laasbelob_" + grpid).attr("checked", "checked");

        mt_timer = $("#FM_mtype_ti_" + id + "").val().replace(".", "")
        mt_timer = mt_timer.replace(",", ".")

        mt_timepris = $("#FM_mtype_tp_" + id + "").val().replace(".", "")
        mt_timepris = mt_timepris.replace(",", ".")

        mt_timer = Math.round(mt_timer * 100) / 100
        mt_timepris = Math.round(mt_timepris * 100) / 100




        //mt_belob = document.getElementById("FM_mtype_be_" + id + "").value.replace(".", "")
        //mt_belob = mt_belob.replace(",", ".")



        if (mt_timer >= 0 && mt_timepris >= 0) {

            mt_belob = (mt_timer * mt_timepris)
            mt_belob = Math.round(mt_belob * 100) / 100
            mt_belob = String(mt_belob)
            mt_belob = mt_belob.replace(".", ",")
            $("#FM_mtype_ff_" + id + "").val(mt_belob) // samme som beløb da faktor altid er 1



        }

        
      
        //Beregner faktor for gruppe og beløb på hver linje

        stno = $("#FM_mtype_startno_" + grpid).val()
        slno = $("#FM_mtype_slutno_" + grpid).val()

        var nystf = stno / 1;
        var nyslf = slno / 1;
        var z = 0;
        var mt_belob_ff_tot = 0;

        //nxtid = $("#FM_mtype_ids_8").val()
        //alert($("#FM_mtype_ff_" + grpid + "_" + nxtid).val() + " i: " + i)
        
        for (z = nystf; z < nyslf; z++) {

            //alert("her" + z)

            nxtid = $("#FM_mtype_ids_" + z).val()

            mt_belob_ff = $("#FM_mtype_ff_" + grpid + "_" + nxtid).val().replace(".", "")
            mt_belob_ff = mt_belob_ff.replace(",", ".")
            mt_belob_ff = Math.round(mt_belob_ff * 100) / 100


            $("#FM_mtype_ti_" + grpid + "_" + nxtid).css("background-color", "#FFFFFF");

            mt_belob_ff_tot = mt_belob_ff_tot + mt_belob_ff
            //alert(mt_belob_ff_tot)

        }



        //$("#mtyp_msg2").hide(2000)

        if (mt_belob_ff_tot == 0 || mt_belob_ff_tot == "NaN" ) {
            mt_belob_ff_tot = 1
        }



        //alert(totBelGrp)
        if ($("#FM_laasbelob_" + grpid).is(':checked') == true) {


            if (totBelGrp == 0 || totBelGrp == "NaN") {

              
                mt_faktor_beregn = 1
                mt_faktor = 1

            } else {


                mt_faktor_beregn = (totBelGrp / mt_belob_ff_tot)
                //alert(mt_faktor_beregn)
                mt_faktor_beregn = mt_faktor_beregn
                mt_faktor = Math.round(mt_faktor_beregn * 100) / 100

                mt_faktor = String(mt_faktor)
                mt_faktor = mt_faktor.replace(".", ",")
            }

        } else {

            mt_faktor_beregn = 1
            mt_faktor = 1

        }

        //mt_faktor = 1
        //alert(totBelGrp + "/" + mt_belob_ff_tot + " i:" + i + " stno:" + stno + " slno: " + slno + " mt_faktor:" + mt_faktor)
        //alert($("#FM_mtype_ids_8").val())
        //alert($("#FM_mtype_be_2_8").val())

        //for (i = stno; i < slno; i++) {
        var nyst = stno/1;
        var nysl = slno/1;
        for (i = nyst; i < nysl; i++) {
            //alert("her:" + i)
            nxtid = $("#FM_mtype_ids_" + i).val()

            
            if ($("#FM_mtype_ti_" + grpid + "_" + nxtid).val() != 0) {
                $("#FM_mtype_fa_" + grpid + "_" + nxtid).val(mt_faktor)
            } else {
                $("#FM_mtype_fa_" + grpid + "_" + nxtid).val('1')
            }

            mt_belob_ff = $("#FM_mtype_ff_" + grpid + "_" + nxtid).val().replace(".", "")
            mt_belob_ff = mt_belob_ff.replace(",", ".")
            mt_belob_ny = Math.round(mt_belob_ff * mt_faktor_beregn * 100) / 100

            mt_belob_ny = String(mt_belob_ny)
            mt_belob_ny = mt_belob_ny.replace(".", ",")

            $("#FM_mtype_be_" + grpid + "_" + nxtid).val(mt_belob_ny)



        }




        //alert(grpid)
      
            setbelobTotpGrp(grpid);
           

    }







    ///// Beregner Sum af beløb for hver linje på grp og overfører til total for gruppen ////
    function setbelobTotpGrp(grpid) {

        //alert("setToBel: " + grpid)


        var stno = 0;
        var slno = 0;
        var k = 0;

        stno = document.getElementById("FM_mtype_startno_" + grpid).value
        slno = document.getElementById("FM_mtype_slutno_" + grpid).value

        //alert("her 0" + stno + " to " + slno)


        var btStno = stno / 1;
        var btSlno = slno / 1;

        var belob_tot = 0;
        var timer_tot = 0;
        var timepris_tot = 0;
        var faktor_tot = 0;
        var belobkost_tot = 0;

        for (k = btStno; k < btSlno; k++) {

            //alert("her")

            nxtid = $("#FM_mtype_ids_" + k).val()

            thisVal = $("#FM_mtype_be_" + grpid + "_" + nxtid).val()
            thisTimer = $("#FM_mtype_ti_" + grpid + "_" + nxtid).val()
            thisTimepris = $("#FM_mtype_tp_" + grpid + "_" + nxtid).val()
            thisFaktor = $("#FM_mtype_fa_" + grpid + "_" + nxtid).val()
            thisKostpris = $("#FM_mtype_kp_" + grpid + "_" + nxtid).val()
            //alert(thisKostpris)

            if (typeof thisVal === 'undefined' || thisVal == "NaN") {

                belob_tot = belob_tot
                timer_tot = timer_tot
                timepris_tot = timepris_tot
                faktor_tot = faktor_tot
                belobkost_tot = belobkost_tot

            } else {

                thisVal = thisVal.replace(".", "")
                thisVal = thisVal.replace(",", ".")
                thisVal = Math.round(thisVal * 100) / 100
                belob_tot = belob_tot + thisVal

                thisTimer = thisTimer.replace(".", "")
                thisTimer = thisTimer.replace(",", ".")
                thisTimer = Math.round(thisTimer * 100) / 100
                timer_tot = timer_tot + thisTimer

                thisTimepris = thisTimepris.replace(".", "")
                thisTimepris = thisTimepris.replace(",", ".")
                thisTimepris = Math.round(thisTimepris * thisTimer * 100) / 100
                timepris_tot = timepris_tot + thisTimepris

                thisFaktor = thisFaktor.replace(".", "")
                thisFaktor = thisFaktor.replace(",", ".")
                thisFaktor = Math.round(thisFaktor * thisTimer * 100) / 100
                faktor_tot = faktor_tot + thisFaktor

                thisKostpris = thisKostpris.replace(".", "")
                thisKostpris = thisKostpris.replace(",", ".")
                thisKostpris = Math.round(thisKostpris * thisTimer * 100) / 100

                
                belobkost_tot = belobkost_tot + thisKostpris
                //alert(thisKostpris + "##" + belobkost_tot)
                

            }
        }

        belob_tot = belob_tot
        belob_tot = String(belob_tot)
        belob_tot = belob_tot.replace(".", ",")

        timer_tot = timer_tot
        timepris_tot = Math.round(timepris_tot / timer_tot * 100) / 100

        belobff_tot = Math.round(timepris_tot * timer_tot * 100) / 100

        faktor_tot = Math.round(faktor_tot / timer_tot * 100) / 100

        timer_tot = String(timer_tot)
        timer_tot = timer_tot.replace(".", ",")


        timepris_tot = String(timepris_tot)
        timepris_tot = timepris_tot.replace(".", ",")

        
        belobff_tot = String(belobff_tot)
        belobff_tot = belobff_tot.replace(".", ",")


        faktor_tot = String(faktor_tot)
        faktor_tot = faktor_tot.replace(".", ",")

        belobkost_tot = String(belobkost_tot)
        belobkost_tot = belobkost_tot.replace(".", ",")
        

        if ($("#FM_laasbelob_" + grpid).is(':checked') == false) {
            $("#FM_mt_belob_totgrp_" + grpid).val(belob_tot)
            
        }

        $("#span_mtype_totudspec_" + grpid).html(belobff_tot)

        //alert("timepris_tot:" + timepris_tot)
        if (timepris_tot == "NaN") {
            timepris_tot = 0
        }

        if (faktor_tot == "NaN") {
            faktor_tot = 1
        }


        $("#FM_mt_timer_totgrp_" + grpid).val(timer_tot)
        $("#FM_mt_timepris_totgrp_" + grpid).val(timepris_tot)
        $("#FM_mt_faktor_totgrp_" + grpid).val(faktor_tot)

        $("#FM_mt_belobff_totgrp_" + grpid).val(belobff_tot)

        $("#FM_mt_kostbelob_totgrp_" + grpid).val(belobkost_tot)

        if ($("#FM_laasbelob_" + grpid).is(':checked') == false) {
            belobtilMtypGrp();
        }

        
        
        totprMtypGrp_tilNettoOms();

    }






    //// Beregning på linje ///
    function setFaktorEqOne(grpid) {

        //thisVal = $("#FM_mt_belob_totgrp_" + grpid).val()
        //thisVal = String(thisVal)
        //thisVal = thisVal.replace(".", "")

        //alert("fak EQ 1")

        //Sæt faktor = 1 når der ændres bruttobeløb på grp.


        stno = document.getElementById("FM_mtype_startno_" + grpid).value
        slno = document.getElementById("FM_mtype_slutno_" + grpid).value

        antalLinier = (slno - stno)
        nyFaktor = 1;

        belobTotThis = $("#FM_mt_belob_totgrp_" + grpid).val()
        timerTotThis = $("#FM_mt_timer_totgrp_" + grpid).val()
        timeprisTotThis = $("#FM_mt_timepris_totgrp_" + grpid).val()
        
        belobTotThis = belobTotThis.replace(".", "")
        belobTotThis = belobTotThis.replace(",", ".")

        timerTotThis = timerTotThis.replace(".", "")
        timerTotThis = timerTotThis.replace(",", ".")

        timeprisTotThis = timeprisTotThis.replace(".", "")
        timeprisTotThis = timeprisTotThis.replace(",", ".")

        faktorTotThis = Math.round((belobTotThis / (timerTotThis * timeprisTotThis)) * 1000) / 1000
        //alert(faktorTotThis)

        for (i = stno; i < slno; i++) {


            nxtid = $("#FM_mtype_ids_" + i).val()

            timerThis = $("#FM_mtype_ti_" + grpid + "_" + nxtid).val()
            timeprisThis = $("#FM_mtype_tp_" + grpid + "_" + nxtid).val()

            


            if (typeof timerThis === 'undefined' || timerThis == "NaN") {



            } else {




                timerThis = timerThis.replace(".", "")
                timerThis = timerThis.replace(",", ".")

                //alert(timerThis)


                if (timerThis <= 0 && timerThis.length > 0) {
                    timerThis = 0
                    $("#FM_mtype_ti_" + grpid + "_" + nxtid).val(timerThis)
                }
                    


                timeprisThis = timeprisThis.replace(".", "")
                timeprisThis = timeprisThis.replace(",", ".")

               



                if (timerThis != 0) {

                    mt_belob_ff = Math.round((timerThis * timeprisThis) * 100) / 100 //Math.round((nytBelob / nyFaktor) * 100) / 100
                    mt_belob_ff = String(mt_belob_ff)
                    mt_belob_ff = mt_belob_ff.replace(".", ",")
                    $("#FM_mtype_ff_" + grpid + "_" + nxtid).val(mt_belob_ff)

                    //thisVal = (belobThis / (timerThis * timeprisThis))
                    nyFaktor = faktorTotThis //Math.round(thisVal * 1000) / 1000 // 3 dec på ocerførsel
                    nytBelob = Math.round(timerThis * timeprisThis * nyFaktor * 100) / 100

                

                    nyFaktor = String(nyFaktor)
                    nyFaktor = nyFaktor.replace(".", ",")
                    $("#FM_mtype_fa_" + grpid + "_" + nxtid).val(nyFaktor)

                    nytBelob = String(nytBelob)
                    nytBelob = nytBelob.replace(".", ",")
                    $("#FM_mtype_be_" + grpid + "_" + nxtid).val(nytBelob)

                    
                } else {

                    $("#FM_mtype_ff_" + grpid + "_" + nxtid).val('0')
                    $("#FM_mtype_fa_" + grpid + "_" + nxtid).val('1')
                    $("#FM_mtype_be_" + grpid + "_" + nxtid).val('0')

                }

            }

        }


        $("#FM_mt_faktor_totgrp_" + grpid).val(nyFaktor)

        

    }



    
  



    ///////////////////// Tilfører / Overfører beløb til den enkelte gruppe ved indtastning på bruttooms / eller på en anden mtypgrp
    function belobtilMtypGrp() {


        //alert("her")
      

        thisVal = $("#FM_budget").val()
        thisVal = String(thisVal)
        thisVal = thisVal.replace(".", "")
        thisVal = thisVal.replace(",", ".")
        //document.getElementById("SP_budget").innerHTML = 
        $("#SP_budget").html("<b>" + thisVal + "</b>");


        if ($("#sync_budget_mt").val() == '1') { //Medarbejdertpy budget slået til

            antalOpenCalc = $("#antalOpenCalc").val()

            if (antalOpenCalc > 1) { //lukke for indtastning dvs fødes fra brutto
                belob_tot = (thisVal / antalOpenCalc)

                belob_tot = Math.round(belob_tot * 100) / 100
                belob_tot = String(belob_tot)
                belob_tot = belob_tot.replace(".", ",")

            } else {
                belob_tot = thisVal
            }

            var fradrag = 0;
            fradrag = $("#FM_salgspris_ulev").val()
            fradrag = fradrag.replace(".", "")
            fradrag = fradrag.replace(",", ".")
            fradrag = Math.round(fradrag * 100) / 100

            //alert(fradrag)

            var thisBel = 0;

            //alert(belob_tot)
            //Fradrag for de belæb der er indtastet på de andre grupper (ved rediger)
            for (i = 0; i < 20; i++) { //Optil 20 grupper

                //alert($("#FM_mt_opencalc_" + i).val())
                if ($("#FM_mt_opencalc_" + i).val() == '1') {
                    thisBel = $("#FM_mt_belob_totgrp_" + i).val()
                    thisBel = thisBel.replace(".", "")
                    thisBel = thisBel.replace(",", ".")
                    thisBel = Math.round(thisBel * 100) / 100
                    fradrag = (fradrag / 1) + (thisBel)

                    $("#FM_mt_belob_totgrp_" + i).css("background-color", "#FFFFE1");


                }
            }

            for (i = 0; i < 20; i++) { //Optil 20 grupper

                //alert($("#FM_mt_opencalc_" + i).val())
                if ($("#FM_mt_opencalc_" + i).val() == '0') {
                    belob_tot = Math.round((belob_tot - (fradrag)) * 100) / 100
                    belob_tot = String(belob_tot)
                    belob_tot = belob_tot.replace(".", ",")

                    $("#FM_mt_belob_totgrp_" + i).val(belob_tot)
                    //$("#span_mtype_totudspec_" + i).html(belob_tot)

                    // Sætter faktor og beregning på linje på gruppen
                    setFaktorEqOne(i);
                }
            }


        }

        totprMtypGrp_tilNettoOms();

    }




    $("#FM_interntbelob").keyup(function () {



        //$("#nettoomsnote").hide("fast");

    });

    $("#FM_budgettimer").keyup(function () {

        //$("#nettoomsnote").hide("fast");

    });



    $(".mt_faktor").keyup(function () {
        var thisid = this.id
        var idlngt = thisid.length
        var idtrim = thisid.slice(12, idlngt)

        //alert(idtrim)


        metypeBelobprlinje(idtrim);

    });

    $(".mt_faktor").change(function () {
        var thisid = this.id
        var idlngt = thisid.length
        var idtrim = thisid.slice(12, idlngt)

        //alert(idtrim)


        metypeBelobprlinje(idtrim);

    });


    $(".mt_belob").keyup(function () {
        var thisid = this.id
        var idlngt = thisid.length
        var idtrim = thisid.slice(12, idlngt)

        //alert(idtrim)

        return false
        //metypeBelobprlinje(idtrim);


    });



    $(".mt_timepris").keyup(function () {
        var thisid = this.id
        var idlngt = thisid.length
        var idtrim = thisid.slice(12, idlngt)

        //alert(idtrim)


        metypeBelobprlinje(idtrim);


    });

    $(".mt_timepris").change(function () {
        var thisid = this.id
        var idlngt = thisid.length
        var idtrim = thisid.slice(12, idlngt)

        //alert(idtrim)


        metypeBelobprlinje(idtrim);


    });




    $(".mt_timer").keyup(function () {
        var thisid = this.id
        var idlngt = thisid.length
        var idtrim = thisid.slice(12, idlngt)

        //alert(idtrim)


        metypeBelobprlinje(idtrim);



    });


    $(".mt_timer").change(function () {
        var thisid = this.id
        var idlngt = thisid.length
        var idtrim = thisid.slice(12, idlngt)

        //alert(idtrim)


        metypeBelobprlinje(idtrim);



    });




    //onload
    var internKost_tot = 0;
    //var internKost_GrandTot = 0;
    if ($("#sync_budget_mt").val() == '1') {
        for (grpid = 0; grpid < 20; grpid++) { //Optil 20 grupper


        
            internKost_tot = $("#FM_mt_kostbelob_totgrp_" + grpid).val()
            $("#span_mtype_totudspec_" + grpid).html(internKost_tot)

            //internKost_GrandTot = internKost_GrandTot/1 + internKost_tot/1

      }

       
    }


    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



    /// Bruges IKKE ?? ///
    function metypeBudget_timer(grpid) {

        alert("FJENR_ ? meTypeBudget_tiemr")

        stno = document.getElementById("FM_mtype_startno_" + grpid).value
        slno = document.getElementById("FM_mtype_slutno_" + grpid).value

        var mt_belob_grpTot = 0;

        //alert(stno + " - " + slno)

        //i = stno

        for (i = stno; i < slno; i++) {

            nxtid = $("#FM_mtype_ids_" + i).val()

            mt_belob = document.getElementById("FM_mtype_be_" + grpid + "_" + nxtid).value.replace(".", "")
            mt_belob = mt_belob.replace(",", ".")
            mt_belob = Math.round(mt_belob * 100) / 100

            mt_timepris = document.getElementById("FM_mtype_tp_" + grpid + "_" + nxtid).value.replace(".", "")
            mt_timepris = mt_timepris.replace(",", ".")
            mt_timepris = Math.round(mt_timepris * 100) / 100

            if (mt_timepris != 0) {
                mt_timepris = mt_timepris
            } else {
                mt_timepris = 1
            }

            mt_faktor = document.getElementById("FM_mtype_fa_" + grpid + "_" + nxtid).value.replace(".", "")
            mt_faktor = mt_faktor.replace(",", ".")
            mt_faktor = mt_faktor //Math.round(mt_faktor * 100) / 100

            mt_timer = (mt_belob / (mt_faktor * mt_timepris))
            mt_timer = Math.round(mt_timer * 100) / 100

            mt_timer = String(mt_timer)
            mt_timer = mt_timer.replace(".", ",")

            document.getElementById("FM_mtype_ti_" + grpid + "_" + nxtid).value = mt_timer;

            mt_belob_grpTot = mt_belob_grpTot + mt_belob
            mt_timer_grpTot = mt_timer_grpTot + mt_timer
            mt_timepris_grpTot = mt_timepris_grpTot + mt_timer
            mt_faktor_grpTot = mt_faktor_grpTot + mt_timer

        }

        document.getElementById("FM_mt_belob_totgrp_" + grpid).value = mt_belob_grpTot;
       //FM_mt_belob_totgrp_
        document.getElementById("FM_mt_timer_totgrp_" + grpid).value = mt_timer_grpTot;
        document.getElementById("FM_mt_timepris_totgrp_" + grpid).value = mt_timepris_grpTot;
        document.getElementById("FM_mt_faktor_totgrp_" + grpid).value = mt_faktor_grpTot;


        totprMtypGrp_tilNettoOms();

        //totalbelob_ff();



    }





    //// Ved tildeling af budget + ændring af 
    function metypeBudget_belob(id) {

        alert("FJENRN? :_ meTypeBudget_belob: " + id)
        return false

        mt_timer = document.getElementById("FM_mtype_ti_" + id + "").value.replace(".", "")
        mt_timer = mt_timer.replace(",", ".")
        //mt_timer = Math.round(mt_timer) // * 100) / 100

        mt_timepris = document.getElementById("FM_mtype_tp_" + id + "").value.replace(".", "")
        mt_timepris = mt_timepris.replace(",", ".")
        //mt_timepris = Math.round(mt_timepris) // * 100) / 100

        mt_faktor = 1 //document.getElementById("FM_mtype_fa_" + id + "").value.replace(".", "")
        $("#FM_mtype_fa_" + id + "").val('1,00')
        //mt_faktor = mt_faktor.replace(",", ".")
        //mt_faktor = Math.round(mt_faktor) // * 100) / 100

        //alert(mt_timer +"*"+ mt_timepris +"*"+ mt_faktor)

        mt_belob = (mt_timer * mt_timepris * mt_faktor)
        mt_belob = Math.round(mt_belob * 100) / 100

        mt_belob = String(mt_belob)
        mt_belob = mt_belob.replace(".", ",")

        mt_belobff = (mt_timer * mt_timepris)
        mt_belobff = Math.round(mt_belobff * 100) / 100

        mt_belobff = String(mt_belobff)
        mt_belobff = mt_belobff.replace(".", ",")

        document.getElementById("FM_mtype_ff_" + id + "").value = mt_belobff;

        document.getElementById("FM_mtype_be_" + id + "").value = mt_belob;

        //alert(id)
        //totalbelob_ff();


        var grpid = $("#FM_mtypegrp_id_" + id).val()

        //alert(grpid)
        setbelobTotpGrp(grpid);





    }







    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////




    // Skal ikke kunne indttaste beløb på medarbejdertyper. Beløbet skal altid være afledt af brutto oms ****
    //$(".mt_timer").keyup(function () {
    //    var thisid = this.id
    //    var idlngt = thisid.length
    //    var idtrim = thisid.slice(12, idlngt)

    //alert(idtrim)


    //    metypeBudget_timer(idtrim);


    //});



    /// Hvis info iom fastpris //
    //function fn_angfp() {
    //
    //    //alert("her" + $("#angfp0").is(':checked'))
    //    if ($("#angfp0").is(':checked') == true) {
    //        $("#div_angfp").css("display", "");
    //        $("#div_angfp").css("visibility", "visible");
    //        $("#div_angfp").show("fast");


    //   } else {

    //       $("#div_angfp").hide("fast");


    //  }

    //}


    $("#a_indforstamgrp").click(function () {
        if ($("#div_indforstamgrp").css('display') == "none") {
            $("#div_indforstamgrp").css("display", "");
            $("#div_indforstamgrp").css("visibility", "visible");
            $("#div_indforstamgrp").show("fast");
            //$.scrollTo('1400px', 400);
            //$.scrollTo($('#div_indforstamgrp'), 4000);

        } else {

            $("#div_indforstamgrp").hide("fast");
            //$.scrollTo('1000px', 400);
            //$.scrollTo('-=100px', 1500);

        }
    });


    $("#xx_angfp0").click(function () {
        //alert("her")
        $("#vistp").css("background-color", "#8CAAe6");
        //fn_angfp();
    });

    $("#angfp1").click(function () {

        $("#vistp").css("background-color", "#cccccc");
        // fn_angfp();
        //$("#vistp").css("background-color", "#8CAAe6");
    });


    // fn_angfp();


    // Aktlisten //
    function getAktlisten() {
        if ($("#jq_vispasogluk").is(':checked') == true) {
            var visalle = 1
        } else {
            var visalle = 0
        }

        var jid_val = $("#jq_jobid")

        slto = $("#slto").val()
        func = $("#func").val()

        //alert(func)

        $.post("?jq_visalle=" + visalle + "&slto=" + slto + "&func=" + func, { control: "FN_getAktlisten", AjaxUpdateField: "true", cust: jid_val.val() }, function (data) {
            //$("#FM_modtageradr").val(data);
            $("#jq_aktlisten").html(data);
            //$("#jobid").html(data);
            //alert(data)

            varTotTim = $("#jq_akttottimer").val() //String(Math.round($("#jq_akttottimer").val() * 100) / 100).replace(".", ",")
            $("#fasertimertot").html("<b>" + varTotTim + "</b>")

            varTotSum = $("#jq_akttotsum").val() //String(Math.round($("#jq_akttottimer").val() * 100) / 100).replace(".", ",")
            $("#fasersumtot").html("<b>" + varTotSum + "</b>")



            //alert("her " + $("#FM_budgettimer").val() + "# " + varTotTim)
            $("#FM_fasersumtot").val(varTotSum)
            $("#FM_fasertimertot").val(varTotTim)


        });


    }








    $("#FM_navn").focus(function () {
        //alert("her")
        var thisval = $("#FM_navn").val()

        if (thisval == "Jobnavn..") {
            $("#FM_navn").val('')
            $("#FM_navn").css("color", "#000000");
            $("#pb_jobnavn").html('')

        }

    });


    $("#a_budgetmtyp").click(function () {

        if ($(".tr_budgetmtyp").css('display') == "none") {
            $(".tr_budgetmtyp").css("display", "");
            $(".tr_budgetmtyp").css("visibility", "visible");
            $(".tr_budgetmtyp").show("fast");
            //$.scrollTo('1400px', 400);
            $.scrollTo($('#tr_budgetmtyp'), 300, { offset: -100 });

        } else {

            $(".tr_budgetmtyp").hide("fast");
            $.scrollTo($('#tr_budgetmtyp'), 300, { offset: -200 });

        }

    });


    if ($.cookie("showsalgsomk") == '1') {
        $(".tr_salgsomk").css("display", "");
        $(".tr_salgsomk").css("visibility", "visible");
        $(".tr_salgsomk").show("fast");
    }

    if ($.cookie("showsalgsomk") == '0') {
        $(".tr_salgsomk").hide("fast");
    }


    $("#a_salgsomk").click(function () {

        if ($(".tr_salgsomk").css('display') == "none") {
            $(".tr_salgsomk").css("display", "");
            $(".tr_salgsomk").css("visibility", "visible");
            $(".tr_salgsomk").show("fast");
            //$.scrollTo('1400px', 400);

            $.scrollTo($('#tr_salgsomk'), 300, { offset: -100 });

            $.cookie("showsalgsomk", '1');

        } else {

            $(".tr_salgsomk").hide("fast");
            $.scrollTo($('#tr_salgsomk'), 300, { offset: -200 });

            $.cookie("showsalgsomk", '0');

        }

    });




    $("#a_aktlisten").click(function () {

        showhideaktlisten();

    });



    if ($.cookie("showaktfaser") == '1') {
        $(".dv_aktlisten").css("display", "");
        $(".dv_aktlisten").css("visibility", "visible");
        $(".dv_aktlisten").show("fast");
    }

    if ($.cookie("showaktfaser") == '0') {
        $(".dv_aktlisten").hide("fast");
    }


    function showhideaktlisten() {


        if ($(".dv_aktlisten").css('display') == "none") {
            $(".dv_aktlisten").css("display", "");
            $(".dv_aktlisten").css("visibility", "visible");
            $(".dv_aktlisten").show("fast");
            //$.scrollTo('1400px', 400);

            $.scrollTo($('#jq_aktlisten'), 300, { offset: -350 });

            $.cookie("showaktfaser", '1');


        } else {

            $(".dv_aktlisten").hide("fast");
            $.scrollTo($('#jq_aktlisten'), 300, { offset: -200 });

            $.cookie("showaktfaser", '0');

        }

    }





    $("#a_tild_prg").click(function () {

        if ($("#div_tild_prg").css('display') == "none") {
            $("#div_tild_prg").css("display", "");
            $("#div_tild_prg").css("visibility", "visible");
            $("#div_tild_prg").show("fast");
            $.scrollTo('1400px', 400);

        } else {

            $("#div_tild_prg").hide("fast");
            $.scrollTo('1000px', 400);

        }

    });

    $("#a_tild_forr").click(function () {

        if ($("#div_tild_forr").css('display') == "none") {

            $("#div_tild_forr").css("position", "relative");
            $("#div_tild_forr").css("top", "0px");
            $("#div_tild_forr").css("left", "0px");
            $("#div_tild_forr").css("border", "0px");
            $("#div_tild_forr").css("display", "");
            $("#div_tild_forr").css("visibility", "visible");
            $("#div_tild_forr").show("fast");
            $.scrollTo('1800px', 200);

        } else {

            $("#div_tild_forr").hide("fast");
            $.scrollTo('1400px', 200);

        }

    });

    $("#a_tild_forr2").click(function () {

       $("#div_tild_forr").hide("slow");
           
    });


    $("#FM_fomr").click(function () {

        //if ($("#step").val() == 2) {
        $("#div_tild_forr").hide("slow");

        $("#load").css("display", "");
        $("#load").css("visibility", "visible");
        $("#load").show("slow");
        

        autoopret();
        //}

    });


    
    

   




    $("#a_tild_ava").click(function () {

        if ($("#div_tild_ava").css('display') == "none") {
            $("#div_tild_ava").css("display", "");
            $("#div_tild_ava").css("visibility", "visible");
            $("#div_tild_ava").show("fast");
            $.scrollTo('1900px', 400);

        } else {

            $("#div_tild_ava").hide("fast");
            $.scrollTo('1000px', 400);

        }

    });


    $("#FM_tp_jobaktid, #FM_mtype").change(function () {
        //alert("Her")
        $("#rdir").val('redjobcontionue')
        $("#showdiv").val('tpriser')
        $("#jobdata").trigger("submit")
    });


    $("#timepriser_opd").click(function () {
        //alert("Her")
        $("#rdir").val('redjobcontionue')
        $("#showdiv").val('tpriser')
        $("#jobdata").trigger("submit")
    });


    $(".overfortiljob_a").click(function () {
        //alert("Her")
        $("#rdir").val('redjobcontionue')
        $("#showdiv").val('forkalk')
        $("#jobdata").trigger("submit")
    });


    $(".overfortiljob_u").click(function () {
        //alert("Her")
        $("#rdir").val('redjobcontionue')
        $("#showdiv").val('forkalk')
        $("#jobdata").trigger("submit")
    });

    $(".overfortiljob_o").click(function () {
        //alert("Her")
        $("#rdir").val('redjobcontionue')
        $("#showdiv").val('')
        $("#jobdata").trigger("submit")
    });

    $("#overfortiljob_tp").click(function () {
        //alert("Her")
        $("#rdir").val('redjobcontionue')
        $("#showdiv").val('')
        $("#jobdata").trigger("submit")
    });




    $("#tilfoj_ulevgrp").click(function () {

        //alert("her")
        if ($(".jq_ulevgrp").css('display') == "none") {

            $(".jq_ulevgrp").css("display", "");
            $(".jq_ulevgrp").css("visibility", "visible");
            $(".jq_ulevgrp").show("fast");

            //$.scrollTo(this, 300, { offset: -100 });

            $("#jq_stamaktgrp").hide("fast");
            $.scrollTo($('#tilfojulevdiv'), 300, { offset: -100 });

        } else {

            $(".jq_ulevgrp").hide("fast");
            //$.scrollTo('-=100px', 1500);
        }
    });


    $("#luk_ulevgrp").click(function () {



        $(".jq_ulevgrp").hide("fast");

    });


    $("#stgrp_tilfoj").click(function () {
        if ($("#jq_stamaktgrp").css('display') == "none") {

            $("#jq_stamaktgrp").css("display", "");
            $("#jq_stamaktgrp").css("visibility", "visible");
            $("#jq_stamaktgrp").show("fast");

            //$.scrollTo('-=0px', 400);
            //$.scrollTo(this, 3000, { offset: -100 });

            //$('#jq_stamaktgrp').fadeIn('slow', function () {
            // Animation complete
            //});

            $(".jq_ulevgrp").hide("fast");

            $.scrollTo($('#tilfojstamdiv'), 300, { offset: -100 });
            //$.scrollTo('-=100px', 4000);

        } else {

            $("#jq_stamaktgrp").hide("fast");
            //$.scrollTo('-=100px', 300);
        }
    });

    $("#stgrp_luk").click(function () {

        $("#jq_stamaktgrp").hide("fast");
        //$.scrollTo('-=100px', 1500);

    });


    $("#jq_vispasogluk").click(function () {
        getAktlisten()
    });


    $("#jq_visaktogfaser").click(function () {
        getAktlisten()
        //$.cookie("showaktfaser" '1');
    });

    //if ($.cookie("showaktfaser") == '1') {
    //   getAktlisten()
    //}

    getAktlisten();

    $(".drag").click(function () {
        $("select[name*=ajax]").AjaxUpdateField({ parent: "tr", subselector: "td:first > input[name=rowId]" });
        $("#incidentlist").table_sort({ items: 'tbody > tr:gt(1):has(:input[name=rowId])', IdControlNode: "td:first > :input[name=rowId]" });
    });

    $(".drag").mouseover(function () {
        $(this).css('cursor', 'pointer');
    });


    //$("#incidentlist").click(function() {
    //    alert("her")
    //});

    $(".jq_akttimer").keyup(function () {

        var thisid = this.id

        var idlngt = thisid.length
        var idtrim = thisid.slice(3, idlngt)


        //alert(idtrim)


    });


    $(".fs_a").keyup(function () {
        var thisid = this.id
        var idlngt = thisid.length
        var idtrim = thisid.slice(3, idlngt)
        var faseval = this.value
        //alert(faseval + "" + idtrim)
        $(".a_fs_" + idtrim).val(faseval)
    });




    function showhideakt_faser_pb(id, a) {

        //alert("grp:" +id)
        var lastopen = $("#lastopen_" + a).val()
        //$(".akt_pb_" + lastopen + "_" + a).hide(1000);

        if (id != 0) {


            if ($(".akt_pb_" + id + "_" + a).css('display') == "none") {

                $(".akt_pb_" + id + "_" + a).css("display", "");
                $(".akt_pb_" + id + "_" + a).css("visibility", "visible");
                $(".akt_pb_" + id + "_" + a).show('fast');
                //$.scrollTo('250px', 500);

                $("#lastopen_" + a).val(id)

            }
        }


    }


    function showhideUlevLinier(id) {

        //alert("her: " + id)
        var lastopen = $("#lastopen_ulev").val()
        $(".ulev_pb_" + lastopen).hide(1000);

        if (id != 0) {


            if ($(".ulev_pb_" + id).css('display') == "none") {

                $(".ulev_pb_" + id).css("display", "");
                $(".ulev_pb_" + id).css("visibility", "visible");
                $(".ulev_pb_" + id).show('fast');
                //$.scrollTo('250px', 500);



            }
        }

        $("#lastopen_ulev").val(id)


    }



    $("#jq_seljugrp_id").change(function () {

        //var thisid = this.id
        //alert("her")
        //$("#stgrpfs_" + thisid + "").val('')
        var thisval = $("#jq_seljugrp_id option:selected").val()
        //alert(thisval)
        showhideUlevLinier(thisval);

    });






    function showhidekops() {
        if ($("#kontaktoplDiv").css('display') == "none") {

            $("#kontaktoplDiv").css("display", "");
            $("#kontaktoplDiv").css("visibility", "visible");
            $("#kontaktoplDiv").show(4000);

        } else {

            $("#kontaktoplDiv").hide(1000);

        }
    }


    //$("#FM_navn").keyup(function () {
    //    $("#pb_jobnavn").text($("#FM_navn").val())
    //});

    //$("#pb_jobnavn").text($("#FM_navn").val())


    $("#kontaktoplA").click(function () {
        showhidekops();
    });

    $("#kontaktoplA2").click(function () {
        showhidekops();
    });

   

    $(".ulev_ryd").click(function () {
        var thisid = this.id
        var idlngt = thisid.length
        var idtrim = thisid.slice(9, idlngt)


        $("#ulevfase_" + idtrim).val('');
        $("#ulevnavn_" + idtrim).val('');
        $("#ulevstk_" + idtrim).val('0,00');
        $("#ulevstkpris_" + idtrim).val('0,00');
        $("#ulevpris_" + idtrim).val('0,00');
        $("#ulevfaktor_" + idtrim).val('1,00');
        $("#ulevbelob_" + idtrim).val('0,00');

        //document.getElementById("ulevlinie_" + thisid + "").style.visibility = "hidden"
        //document.getElementById("ulevlinie_" + thisid + "").style.display = "none"


        $.scrollTo(this, 300, { offset: -300 });
        $("#ulevlinie_" + idtrim).hide(100);

        settotalbelob();

    });


    $("#ulev_tilfoj_line").click(function () {
        var thisval = $("#antalulev option:selected").val()
        if (thisval <= 49) {
            thisval = Math.round(thisval / 1 + 1)
            $("#antalulev option[value='" + thisval + "']").attr('selected', 'selected');

            for (i = 1; i <= 50; i++) {
                if (i <= thisval) {
                    document.getElementById("ulevlinie_" + i + "").style.visibility = "visible"
                    document.getElementById("ulevlinie_" + i + "").style.display = ""
                } else {
                    document.getElementById("ulevbelob_" + i + "").value = ""
                    document.getElementById("ulevlinie_" + i + "").style.visibility = "hidden"
                    document.getElementById("ulevlinie_" + i + "").style.display = "none"
                }
            }
        }

        $.scrollTo('-=100px', 500);

    });


    function xdiff_timer_sum() {


        /// Diff timer og Diff Sum
        var varBudgetTimer = $("#FM_budgettimer").val().replace(".", "")
        varBudgetTimer = varBudgetTimer.replace(",", ".")
        varBudgetTimer = Math.round(varBudgetTimer * 100) / 100

        varTotTim = $("#FM_fasertimertot").val().replace(".", "")
        varTotTim = varTotTim.replace(",", ".")

        var atkTimTot = Math.round(varTotTim * 100) / 100
        var diff_timer = (atkTimTot - (varBudgetTimer))

        //alert(varBudgetTimer + " # " + atkTimTot)
        diff_timer = String(Math.round(diff_timer * 100) / 100).replace(".", ",")

        //diff_timer = String(Math.round(diff_timer * 100) / 100).replace(".", ",")
        $("#diff_timer").html("<b>" + diff_timer + "</b>")
        $("#FM_diff_timer").val(diff_timer)

        var varBudgetSum = $("#FM_interntbelob").val().replace(".", "")
        varBudgetSum = varBudgetSum.replace(",", ".")
        varBudgetSum = Math.round(varBudgetSum * 100) / 100

        varTotSum = $("#FM_fasersumtot").val().replace(".", "")
        varTotSum = varTotSum.replace(",", ".")
        var atkSumTot = Math.round(varTotSum * 100) / 100
        var diff_sum = (varTotSum - (varBudgetSum))

        diff_sum = String(Math.round(diff_sum * 100) / 100).replace(".", ",")

        $("#diff_sum").html("<b>" + diff_sum + "</b>")
        $("#FM_diff_sum").val(diff_sum)


        if (diff_timer != "0") {

            $("#diff_timer_ok").hide("fast");

        } else {

            $("#diff_timer_ok").css("display", "");
            $("#diff_timer_ok").css("visibility", "visible");
            $("#diff_timer_ok").show("fast");

        }


        if (diff_sum != 0) {
            alert(diff_sum)
            $("#diff_sum_ok").css("display", "");
            $("#diff_sum_ok").css("visibility", "visible");
            $("#diff_sum_ok").hide("fast");

        } else {

            $("#diff_sum_ok").css("display", "");
            $("#diff_sum_ok").css("visibility", "visible");
            $("#diff_sum_ok").show("fast");

        }



    }




    function showhideinternnote() {

        //alert($("#div_internnote").css("visibility"))

        if ($("#div_internnote").css('visibility') == "hidden") {

            $("#div_internnote").css("display", "");
            $("#div_internnote").css("visibility", "visible");
            $("#div_internnote").show(1000);
            $.scrollTo('500px', 500);

        } else {

            $("#div_internnote").css("visibility", "hidden");
            $("#div_internnote").css("display", "none");
            $("#div_internnote").hide(200);
            $.scrollTo('300px', 500);

        }
    }


    

    //function showudv() {
    //    document.getElementById("beregnudviddet").style.display = "";
    //    document.getElementById("beregnudviddet").style.visibility = "visible";
    //    document.getElementById("beregnsimpel").style.display = "none";
    //    document.getElementById("beregnsimpel").style.visibility = "hidden";
    //    document.getElementById("FM_budget").value = document.getElementById("FM_budget_s").value;
    //    document.getElementById("FM_budgettimer").value = document.getElementById("FM_budgettimer_s").value;
    //}

    //$("#a_pb_udv").click(function() {
    //
    //});

    $("#a_internnote").click(function () {
        showhideinternnote();
    });

    $("#a2_internnote").click(function () {
        showhideinternnote();
    });



    function showhidetpris() {
        if ($("#tiltimprisDiv").css('display') == "none") {

            $("#tiltimprisDiv").css("display", "");
            $("#tiltimprisDiv").css("visibility", "visible");
            $("#tiltimprisDiv").show(4000);

        } else {

            $("#tiltimprisDiv").hide(1000);

        }
    }

    $("#tiltimprisA").click(function () {
        showhidetpris();
    });

    $("#tiltimprisA2").click(function () {
        showhidetpris();
    });







    //////////////////////////////////////////////////
    // FANEBLADE  Vis divs //
    //////////////////////////////////////////////////

    $("#vismp_A").click(function () {
        //alert("her")
        milepfunc();

    });

    $("#forkalk_A, #forkalk2_A").click(function () {
        //alert("her")

        $("#tpris").hide(100);
        $("#oblik_div").hide(100);
        $("#milep_div").hide(100);

        $("#pb_div").css("display", "");
        $("#pb_div").css("visibility", "visible");
        $("#pb_div").show(200);

        $("#forkalk").css("background-color", "#5582d2");
        $("#vistp").css("background-color", "#8CAAe6");
        $("#oblik").css("background-color", "#8CAAe6");
        $("#vismp").css("background-color", "#8CAAe6");
        $("#showdiv").val('forkalk')

        // fn_angfp();

    });

    $("#oblik_A").click(function () {

        $("#tpris").hide(100);
        $("#pb_div").hide(100);
        $("#milep_div").hide(100);

        $("#oblik_div").css("display", "");
        $("#oblik_div").css("visibility", "visible");
        $("#oblik_div").show(1000);

        $("#oblik").css("background-color", "#5582d2");
        $("#vistp").css("background-color", "#8CAAe6");
        $("#forkalk").css("background-color", "#8CAAe6");
        $("#vismp").css("background-color", "#8CAAe6");
        $("#showdiv").val('stamdata')



    });

    //$("#vistp_A").click(function () {
    //alert("her")
    //    tpriserfunc();

    //});

    $("#visstd_A").click(function () {

        $("#tpris").hide(100);
        $("#pb_div").hide(100);
        $("#milep_div").hide(100);

        $("#oblik_div").css("display", "");
        $("#oblik_div").css("visibility", "visible");
        $("#oblik_div").show(1000);


        $("#sindhold").css("display", "");
        $("#sindhold").css("visibility", "visible");
        $("#sindhold").show(1000);

        $("#visstd").css("background-color", "#5582d2");
        $("#vistp").css("background-color", "#8CAAe6");
        $("#oblik").css("background-color", "#5582d2");
        $("#forkalk").css("background-color", "#8CAAe6");
        $("#vismp").css("background-color", "#8CAAe6");
        $("#showdiv").val('stamdata')


    });

    $("#vistp_A").click(function () {
        //alert("her")
        tpriserfunc();

    });


    function tpriserfunc() {

        $("#tpris").css("display", "");
        $("#tpris").css("visibility", "visible");
        $("#tpris").show(2000);
        $("#tpris").css("left", "485px"); //357


        $("#vistp").css("background-color", "#5582d2");
        $("#visstd").css("background-color", "#5582d2");
        $("#oblik").css("background-color", "#8CAAe6");
        $("#forkalk").css("background-color", "#8CAAe6");
        $("#vismp").css("background-color", "#8CAAe6");


        $("#pb_div").hide(1000);
        $("#oblik_div").hide(1000);
        $("#milep_div").hide(1000);
        //$("#sindhold").hide(1000);
        $("#showdiv").val('tpriser')

    }

    function milepfunc() {

        $("#tpris").hide(100);
        $("#oblik_div").hide(100);
        $("#pb_div").hide(100);

        $("#milep_div").css("display", "");
        $("#milep_div").css("visibility", "visible");
        $("#milep_div").show(200);

        $("#forkalk").css("background-color", "#8CAAe6");
        $("#vistp").css("background-color", "#8CAAe6");
        $("#oblik").css("background-color", "#8CAAe6");
        $("#vismp").css("background-color", "#5582d2");
        $("#showdiv").val('milep')
    }


    if ($("#showdiv").val() == "tpriser") {
        tpriserfunc();

    } else {

        if ($("#showdiv").val() == "forkalk") {
            $("#forkalk").css("background-color", "#5582d2");
        } else {

            if ($("#showdiv").val() == "milep") {
                milepfunc();
                //$("#vismp").css("background-color", "#8CAAe6");
            } else {
                $("#oblik").css("background-color", "#5582d2");
            }
        }


        $("#visstd").css("background-color", "#5582d2");
        //$("#forkalk").css("background-color", "#5582d2");
    }






    /////////////////////////////////////////////////////////////////




    function kopiertilfase(id) {

        var thisid = id
        var thisval = $("#" + thisid + " option:selected").text()

        var thisvallngt = thisval.length
        var thisvaltrim = thisval.slice(0, thisvallngt - 9)
        thisval = thisvaltrim

        if (thisval != 'I') {
            $("#stgrpfs_" + thisid + "").val(thisval)
        } else {
            $("#stgrpfs_" + thisid + "").val('')
        }

        $(".a_fs_" + thisid).val(thisval)

    }





    $(".selstaktgrp").change(function () {

        var thisid = this.id
        //alert(thisid)
        // fase nulstil //
        $("#stgrpfs_" + thisid + "").val('')

        // antal stam akt grupper //
        var antalaktgrp = $("#aktgrpialt").val()
        //alert(antalaktgrp)


        //var thisval = 0;

        var thisval1 = 0
        //if ($("#selstaktgrp_1 option:selected").length != 0) {
        for (var i = 0, a = 0; i < this.options.length; i++) {

            //thisval1 = $(this).val()
            //alert(thisval1 + " " + this.options[i].selected)
            if (this.options[i].selected == true) {
                thisval1 = this.options[i].value
                //alert(thisval1)
                showhideakt_faser_pb(thisval1, 1);
            } else {
                thisval1 = this.options[i].value
                $(".akt_pb_" + thisval1 + "_1").hide('fast');
            }
        }


    });


    $("#fs_ryd_1").click(function () {
        kopiertilfase(1)
    });

    $("#fs_ryd_2").click(function () {
        kopiertilfase(2)
    });

    $("#fs_ryd_3").click(function () {
        kopiertilfase(3)
    });

    $("#fs_ryd_4").click(function () {
        kopiertilfase(4)
    });

    $("#fs_ryd_5").click(function () {
        kopiertilfase(5)
    });





    //settotalbelob();




}); 
///////////////////////////////////////////////////////////////////////////////
/////////////////////////////// END JQUERY //////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////



























	function show(){
	document.all["stamdata"].style.display = "";
	document.all["stamdata"].style.visibility = "visible";
	
	document.all["logoer"].style.display = "none";
	document.all["logoer"].style.visibility = "hidden";
	
	document.all["dok"].style.display = "none";
	document.all["dok"].style.visibility = "hidden";
	
	document.all["log"].style.display = "none";
	document.all["log"].style.visibility = "hidden";
	
	document.all["kpers"].style.display = "none";
	document.all["kpers"].style.visibility = "hidden";
	document.all["subm"].style.top = 650;
	}
	
	//function showtpris() {
	//document.getElementById("tpris").style.display = "";
	//document.getElementById("tpris").style.visibility = "visible";
	//document.getElementById("sindhold").style.display = "none";
	//document.getElementById("sindhold").style.visibility = "hidden";
	//document.getElementById("vistpris").style.top = 161;
	//document.getElementById("visstamdata").style.top = 159;
	//}
	
	function showstamdata() {
	document.getElementById("tpris").style.display = "none";
	document.getElementById("tpris").style.visibility = "hidden";
	document.getElementById("sindhold").style.display = "";
	document.getElementById("sindhold").style.visibility = "visible";
	document.getElementById("vistpris").style.top = 159;
	document.getElementById("visstamdata").style.top = 161;
	}
	
	function mailtokunde(){
	var kid = 0
	var jnr = 0;
	var jnavn 
	kid = document.getElementById("FM_kunde").value 
	jnr = document.getElementById("FM_jnr").value
	jnavn = document.getElementById("FM_navn").value
		if (document.getElementById("func").value == "dbopr") {
			if (document.jobdata.FM_kundese.checked == true ) {
				window.open("job_mailkunde.asp?kid="+kid+"&jnr="+jnr+"&jnavn="+jnavn+"","TimeOut","width=340,height=300,left=0,top=100,screenX=100,screenY=100")
			}
		} 
	}
	
	function hideoprmed(){
	document.getElementById("visopretmed").style.visibility = "hidden";
	document.getElementById("visopretmed").style.display = "none";
	}
	
	
	function beregnuger() {
	var dato
	var antaldage = 0
	
	uger = document.getElementById("FM_antaluger").value
	
		if (uger < 104) {
			
			dag = document.getElementById("FM_start_dag").value
			mnd = document.getElementById("FM_start_mrd").value - 1
			aar = document.getElementById("FM_start_aar").value 
			dato = new Date (aar, mnd, dag)
			antaldage = parseInt(7*uger)
			dato.setDate(dato.getDate() + antaldage)
			document.getElementById("FM_slut_dag").value = dato.getDate()
			document.getElementById("FM_slut_mrd").value = dato.getMonth() + 1
			
			arr_slut = dato.getFullYear() + "";
			antalchar = arr_slut.length 
			 
			var year = eval(arr_slut.substring(arr_slut.length - 2, arr_slut.length));
			if (year < 10) {
			year = "0"+year
			}
			document.getElementById("FM_slut_aar").value = year
		} else {
		alert("Der kan maks angives 104 uger.(ca. 2 år)")
		}
	}
	
	
	//function showalertallejob(){
	//alert("Er du sikker på at du vil oprette dette job på alle kontakter?")
	//}

	function overfortp(mid, t) {

	//alert(document.getElementById("FM_hd_timepris_" + mid + "_" + t).value)
	document.getElementById("FM_6timepris_"+mid+"").value = document.getElementById("FM_hd_timepris_"+mid+"_"+t).value
	valuta = document.getElementById("FM_hd_valuta_"+mid+"_"+t).value
	
	//alert(document.getElementById("FM_valuta_600"+mid+"").length)
	
	for (i=0; i<document.getElementById("FM_valuta_600"+mid+"").length; i++){
	//alert(i +" -- Val: "+ valuta)
	    if (document.getElementById("FM_valuta_600"+mid+"").options[i].value == valuta) {
	    document.getElementById("FM_valuta_600"+mid+"").options[i].selected=true
	    }
    }
    	
	}




	function showulevlinier() {


	    vislinier = document.getElementById("antalulev").value
	    //alert(vislinier)

	    document.getElementById("ulevopen").value = vislinier
	    

	    if (vislinier < 30) {
	        //document.getElementById("div_tilfoj_ulevgrp").style.visibility = "visible"
	        //document.getElementById("div_tilfoj_ulevgrp").style.display = ""
	        document.getElementById("span_tilfoj_ulevgrp").style.visibility = "hidden"
	        document.getElementById("span_tilfoj_ulevgrp").style.display = "none"
	    } else {
	        //document.getElementById("div_tilfoj_ulevgrp").style.visibility = "hidden"
	        //document.getElementById("div_tilfoj_ulevgrp").style.display = "none"
	        document.getElementById("span_tilfoj_ulevgrp").style.visibility = "visible"
	        document.getElementById("span_tilfoj_ulevgrp").style.display = ""
	    }

        
	    
	    //alert(vislinier)
	    for (i=1; i<=50; i++){
	        if (i <= vislinier) {
	        document.getElementById("ulevlinie_"+i+"").style.visibility = "visible"
	        document.getElementById("ulevlinie_"+i+"").style.display = ""
	        } else {
	        document.getElementById("ulevbelob_"+i+"").value = ""
	        document.getElementById("ulevlinie_"+i+"").style.visibility = "hidden"
	        document.getElementById("ulevlinie_"+i+"").style.display = "none"
	        } 
	    }
	    
	}




	


	
	function settotalbelob(){

       

    /// Åbne for manuel beregning af Netto omsætning og dermed også DB og BruttoOmsætning
	  

	        //alert("her")
	        //alert("her set toal: " + document.getElementById("FM_interntbelob").value)
	        var totalbelob = 0;

            
	        if ($("#FM_ign_tp").is(':checked') == false) {


	            if (document.getElementById("FM_interntbelob").value != "") {
	                totalbelob = document.getElementById("FM_interntbelob").value.replace(".", "")
	                totalbelob = totalbelob.replace(",", ".")

	            } else {
	                totalbelob = 0;
	            }


	            nettooms = Math.round(totalbelob * 100) / 100
	           

	        } else {

	            if (document.getElementById("FM_budget").value != "") {

	                totalbelob = 0;
	                totalbelob = document.getElementById("FM_budget").value.replace(".", "")
	                totalbelob = totalbelob.replace(",", ".")
	                totalbelob = Math.round(totalbelob * 100) / 100
	                //totalbelob = totalbelob / 1

	                //alert(totalbelob)

	            } else {
	                totalbelob = 0;
	            }


	            nettooms = document.getElementById("FM_interntbelob").value.replace(".", "")
	            nettooms = nettooms.replace(",", ".")


	        }
            

	        nettooms = String(nettooms)
	        nettooms = nettooms.replace(".", ",")

	        $("#SP_netto").html("<b>" + nettooms + "</b>")


	     

	        var totalomk = 0;
	        var totalsalg = 0;
	        var totalbelobThis = 0;
	        var totalomkThis = 0;

	        antalUlevlinier = document.getElementById("ulevopen").value
	        for (i = 1; i <= antalUlevlinier; i++) {

	            //alert(document.getElementById("ulevbelob_" + i + "").value)
	            if (document.getElementById("ulevbelob_" + i + "").value != "") {

	                totalbelobThis = document.getElementById("ulevbelob_" + i + "").value.replace(".", "");
	                totalbelobThis = totalbelobThis.replace(",", ".");
	                totalbelobThis = Math.round(totalbelobThis * 100) / 100

	                totalbelob = totalbelob / 1 + (totalbelobThis / 1)
	                totalsalg = totalsalg / 1 + (totalbelobThis / 1)

	                //alert(totalbelob)

	                totalomkThis = document.getElementById("ulevpris_" + i + "").value.replace(".", "");
	                totalomkThis = totalomkThis.replace(",", ".");
	                totalomkThis = Math.round(totalomkThis * 100) / 100
	                totalomk = totalomk / 1 + (totalomkThis / 1)

	                //alert(i +" "+ document.getElementById("ulevbelob_"+i+"").value)
	            }
	        }

	        totalbelob = Math.round(totalbelob * 100) / 100
	        totalomk = Math.round(totalomk * 100) / 100
            

            //NU HER
	        //if ($("#sync_budget_mt").val() == '1') {
	            
	    //internomk = $("#FM_interntbelob").val().replace(".", "")

	    //  } else {
	        internomk = document.getElementById("FM_interntomkost").value.replace(".", "")
	    //  }

	        internomk = internomk.replace(",", ".")
	        internomk = Math.round(internomk * 100) / 100


	        //alert(internomk)

	        udgifter = (totalomk + internomk)
	        udgifter = Math.round(udgifter * 100) / 100

	       



	        if (totalbelob != 0) {
	            db = 100 - (((internomk + totalomk) / totalbelob) * 100)
	            db = Math.round(db * 100) / 100
	        } else {
	            db = 0
	        }




	        //alert(totalbelob / 1 +" - ("+ totalomk / 1 +"+"+ internomk / 1+")")
	        bruttofj = totalbelob / 1 - (totalomk / 1 + internomk / 1)

	        internomk = String(internomk)
	        internomk = internomk.replace(".", ",")

	        document.getElementById("SP_udgifter_intern").innerHTML = internomk
	        document.getElementById("FM_udgifter_intern").value = internomk



	        udgifter = String(udgifter)
	        udgifter = udgifter.replace(".", ",")
	        document.getElementById("SP_udgifter").innerHTML = "<b>" + udgifter + "</b>"
	        document.getElementById("FM_udgifter").value = udgifter

	        totalomk = String(totalomk)
	        totalomk = totalomk.replace(".", ",")
	        document.getElementById("SP_udgifter_ulev").innerHTML = totalomk
	        document.getElementById("FM_udgifter_ulev").value = totalomk



	        totalsalg = String(totalsalg)
	        totalsalg = totalsalg.replace(".", ",")
	        document.getElementById("SP_salgspris_ulev").innerHTML = totalsalg
	        document.getElementById("FM_salgspris_ulev").value = totalsalg

            

	        var internKost_tot = 0;
	        //Opdarer mellemregning på budgetoverblik
	        if ($("#sync_budget_mt").val() == '1') {
	            for (grpid = 0; grpid < 20; grpid++) { //Optil 20 grupper



	                internKost_tot = $("#FM_mt_kostbelob_totgrp_" + grpid).val()
	                $("#span_mtype_totudspec_" + grpid).html(internKost_tot)

	             

	            }


	        }
	       

	        //alert($("#sync_budget_mt").val())
	        //if ($("#sync_budget_mt").is(':checked') == false) {


	        if ($("#sync_budget_mt").val() == '0' && $("#FM_ign_tp").is(':checked') == false) {

	            totalbelob = String(totalbelob)
	            totalbelob = totalbelob.replace(".", ",")
	            document.getElementById("SP_budget").innerHTML = "<b>" + totalbelob + "</b>"
	            document.getElementById("FM_budget").value = totalbelob

	        } else {

	            //if (bruttofj < 0) {
	            //alert("Intern kost + Salgsomk. overstiger brutto beløb")
	            
	            //}
                
            }


        bruttofj = Math.round(bruttofj * 100) / 100
        bruttofj = String(bruttofj)
        bruttofj = bruttofj.replace(".", ",")
        document.getElementById("SP_bruttofortj").innerHTML = "<b>" + bruttofj + "</b>"
        document.getElementById("FM_bruttofortj").value = bruttofj



	        db = String(db)
	        db = db.replace(".", ",")
	        document.getElementById("SP_db").innerHTML = "<b>" + db + "</b>"
	        document.getElementById("FM_db").value = db

	    


	
	}


	function dbManuel() { 
    
    //alert("her")

    bruttoOms = document.getElementById("FM_budget").value.replace(".", "")
	bruttoOms = bruttoOms.replace(",", ".")
    bruttoOms = Math.round(bruttoOms * 100) / 100

 
    var totalomk = 0;
    var totalomkThis = 0;

    antalUlevlinier = document.getElementById("ulevopen").value
    for (i = 1; i <= antalUlevlinier; i++) {

        //alert(document.getElementById("ulevbelob_" + i + "").value)
        if (document.getElementById("ulevbelob_" + i + "").value != "") {
           

            totalomkThis = document.getElementById("ulevpris_" + i + "").value.replace(".", "");
            totalomkThis = totalomkThis.replace(",", ".");
            totalomkThis = Math.round(totalomkThis * 100) / 100
            totalomk = totalomk / 1 + (totalomkThis / 1)

            //alert(i +" "+ document.getElementById("ulevbelob_"+i+"").value)
        }
    }

    
    totalomk = Math.round(totalomk * 100) / 100

    internomk = document.getElementById("FM_interntomkost").value.replace(".", "")
    internomk = internomk.replace(",", ".")
    internomk = Math.round(internomk * 100) / 100

    bruttofj = bruttoOms / 1 - (totalomk / 1 + internomk / 1)

    bruttofj = Math.round(bruttofj * 100) / 100
    bruttofj = String(bruttofj)
    bruttofj = bruttofj.replace(".", ",")
    document.getElementById("SP_bruttofortj").innerHTML = "<b>" + bruttofj + "</b>"
    document.getElementById("FM_bruttofortj").value = bruttofj


    if (bruttoOms != 0) {
        db = 100 - (((internomk + totalomk) / bruttoOms) * 100)
        db = Math.round(db * 100) / 100
    } else {
        db = 0
    }


    db = String(db)
    db = db.replace(".", ",")
    document.getElementById("SP_db").innerHTML = "<b>" + db + "</b>"
    document.getElementById("FM_db").value = db


    }




	function beregnintbelob() {
        
        //alert("her")

	    //if ($("#FM_ign_tp").is(':checked') == false) {

	        timer = document.getElementById("FM_budgettimer").value.replace(".", "")
	        timer = timer.replace(",", ".")
	        timer = Math.round(timer * 100) / 100

	        faktor = document.getElementById("FM_intfaktor").value.replace(".", "")
	        faktor = faktor.replace(",", ".")
	        faktor = Math.round(faktor * 100) / 100

	        gnsinttpris = document.getElementById("FM_gnsinttpris").value.replace(".", "")
	        gnsinttpris = gnsinttpris.replace(",", ".")
	        gnsinttpris = Math.round(gnsinttpris * 100) / 100

	        internOmkost = (timer * gnsinttpris)
	        internOmkost = Math.round(internOmkost * 100) / 100

	        internOmkost = String(internOmkost)
	        internOmkost = internOmkost.replace(".", ",")

	        interntBelob = (timer * gnsinttpris * faktor)
	        interntBelob = Math.round(interntBelob * 100) / 100

	        interntBelob = String(interntBelob)
	        interntBelob = interntBelob.replace(".", ",")

	        document.getElementById("FM_interntbelob").value = interntBelob

	        //alert("1")
	        //if ($("#sync_budget_mt").val() == '0') {
	            document.getElementById("FM_interntomkost").value = internOmkost
	        //}

	        targettpris = (gnsinttpris / 1 * faktor / 1)
	        targettpris = Math.round(targettpris * 100) / 100
	        document.getElementById("pb_tg_timepris").value = targettpris

	    //}

	    settotalbelob();
	
	}


	function beregninttp() {

	//alert("2")

    //alert("her")
	//timer = document.getElementById("FM_budgettimer").value.replace(".", "")
	//timer = timer.replace(",", ".")
	    //timer = Math.round(timer * 100) / 100

	    //showFldHours = document.getElementById("showFldHours").value

	    //if ($("#FM_ign_tp").is(':checked') == false) {

	    //if (showFldHours == 1) {

	        gnsinttpris = document.getElementById("FM_gnsinttpris").value.replace(".", "")
	        gnsinttpris = gnsinttpris.replace(",", ".")
	        //gnsinttpris = Math.round(gnsinttpris * 100) / 100

	        faktor = document.getElementById("FM_intfaktor").value.replace(".", "")
	        faktor = faktor.replace(",", ".")
	        //faktor = Math.round(faktor * 100) / 100

	        interntBelob = document.getElementById("FM_interntbelob").value.replace(".", "")
	        interntBelob = interntBelob.replace(",", ".")
	        //interntBelob = Math.round(interntBelob * 100) / 100


	        if (gnsinttpris != 0 && faktor != 0) {
	            timer = (interntBelob / (gnsinttpris * faktor))
	            timer = Math.round(timer * 100) / 100

	        } else {
	            timer = document.getElementById("FM_budgettimer").value.replace(".", "")
	            timer = timer.replace(",", ".")
	            timer = Math.round(timer * 100) / 100
	        }


	        targettpris = (gnsinttpris / 1 * faktor / 1)
	        targettpris = Math.round(targettpris * 100) / 100
	        document.getElementById("pb_tg_timepris").value = targettpris

	        internOmkost = (timer * gnsinttpris)
	        internOmkost = Math.round(internOmkost * 100) / 100

	        internOmkost = String(internOmkost)
	        internOmkost = internOmkost.replace(".", ",")

	        gnsinttpris = String(gnsinttpris)
	        gnsinttpris = gnsinttpris.replace(".", ",")

	        timer = String(timer)
	        timer = timer.replace(".", ",")
	        document.getElementById("FM_budgettimer").value = timer

	        document.getElementById("FM_gnsinttpris").value = gnsinttpris
	        document.getElementById("FM_interntomkost").value = internOmkost

	     

	        //}
	        //}

	   
	        settotalbelob();
	   
	}


	function beregnulevstkpris(u) {

    //alert("her")


	    ustk = document.getElementById("ulevstk_" + u + "").value.replace(".", "")
	    ustk = ustk.replace(",", ".")
	    ustk = Math.round(ustk * 100) / 100

	    ustkpris = document.getElementById("ulevstkpris_" + u + "").value.replace(".", "")
	    ustkpris = ustkpris.replace(",", ".")
	    ustkpris = Math.round(ustkpris * 100) / 100

	    uPris = (ustk/1) * (ustkpris/1)
	    uPris = Math.round(uPris * 100) / 100


	    ufaktor = document.getElementById("ulevfaktor_" + u + "").value.replace(".", "")
	    ufaktor = ufaktor.replace(",", ".")
	    ufaktor = Math.round(ufaktor * 100) / 100

	    
        uBelob = (uPris * ufaktor)
        uBelob = Math.round(uBelob * 100) / 100

        //alert(uBelob)

	    uPris = String(uPris)
	    uPris = uPris.replace(".", ",")

	    document.getElementById("ulevpris_" + u + "").value = uPris

	   
	    uBelob = String(uBelob)
	    uBelob = uBelob.replace(".", ",")

	    document.getElementById("ulevbelob_" + u + "").value = uBelob

	    //if ($("#sync_budget_mt").val() == "1") {
	    //    totalbelob_ff();
	    //} else {
            
	  

	        settotalbelob();
	        //}


	}
	
	
	function beregnulevbelob(u) {
	
    
    ufaktor = document.getElementById("ulevfaktor_"+u+"").value.replace(",",".")
	ufaktor = Math.round(ufaktor*100)/ 100
	
	upris = document.getElementById("ulevpris_"+u+"").value.replace(",",".")
	upris = Math.round(upris*100)/ 100
	
	uBelob = (upris * ufaktor)
	uBelob = Math.round(uBelob*100)/ 100
	
	ustk = document.getElementById("ulevstk_" + u + "").value.replace(",", ".")
	ustk = Math.round(ustk * 100) / 100

	ustkpris = Math.round(upris / ustk) * 100 / 100
    
    ustkpris = String(ustkpris)
	ustkpris = ustkpris.replace(".", ",")

	document.getElementById("ulevstkpris_" + u + "").value = ustkpris;

	uBelob = String(uBelob)
	uBelob = uBelob.replace(".", ",")
	
    document.getElementById("ulevbelob_" + u + "").value = uBelob

    //if ($("#sync_budget_mt").val() == "1") {
    //    totalbelob_ff();
    //} else {
        settotalbelob();
        //}
	
	}
	
	
    function beregnulevipris(u) {
	
	ufaktor = document.getElementById("ulevfaktor_"+u+"").value.replace(",",".")
	ufaktor = Math.round(ufaktor*100)/ 100
	
	uBelob = document.getElementById("ulevbelob_"+u+"").value.replace(",",".")
	uBelob = Math.round(uBelob*100)/ 100
	
	upris =  (uBelob / ufaktor)
	upris = Math.round(upris*100)/ 100
	
    ustk = document.getElementById("ulevstk_" + u + "").value.replace(",", ".")
	ustk = Math.round(ustk * 100) / 100

	ustkpris = Math.round((upris / ustk) * 100) / 100

	ustkpris = String(ustkpris)
	ustkpris = ustkpris.replace(".", ",")

	document.getElementById("ulevstkpris_" + u + "").value = ustkpris;

	upris = String(upris)
	upris = upris.replace(".", ",")
	document.getElementById("ulevpris_"+u+"").value = upris

	//if ($("#sync_budget_mt").val() == "1") {
	//    totalbelob_ff();
	//} else {
	    settotalbelob();
	    //}

}

////////////////////////// Medarb. typer /////////////////////////////////////////
















	
	
	function validZip(inZip){
	if (inZip == "") {
	return true
	}
	if (isNum(inZip)){
	return true
	}
	return false
	}
	
	function isNum(passedVal){
	invalidChars = " /:;<>abcdefghijklmnopqrstuvwxyzæøå"
	if (passedVal == ""){
	return false
	}
	
	for (i = 0; i < invalidChars.length; i++) {
	badChar = invalidChars.charAt(i)
		if (passedVal.indexOf(badChar, 0) != -1){
		return false
		}
	}
	
		for (i = 0; i < passedVal.length; i++) {
			if (passedVal.charAt(i) == ".") {
			return true
			}
			
			if (passedVal.charAt(i) == ",") {
			return true
			}
			
			else {
				if (passedVal.charAt(i) < "0") {
				return false
				}
					if (passedVal.charAt(i) > "9") {
					return false
					}
				}
			return true
			} 
			
	}
	
	function tjektimer(fld){
		
		
		if (!validZip(document.getElementById(""+fld+"").value)){
		alert("Der er angivet et ugyldigt tegn.")
		return false
		}
		
		return true
	}
	
	
	
	
	
	function xxshowsimpel() {
	document.getElementById("beregnudviddet").style.display = "none";
	document.getElementById("beregnudviddet").style.visibility = "hidden";
	document.getElementById("beregnsimpel").style.display = "";
	document.getElementById("beregnsimpel").style.visibility = "visible";
	document.getElementById("FM_budget_s").value = document.getElementById("FM_budget").value;
	document.getElementById("FM_budgettimer_s").value = document.getElementById("FM_budgettimer").value;

}

function serviceaft(thisid, kid, fm_soeg, forny) {
    if (thisid == 0) {
        window.open("serviceaft.asp?func=opret&kundeid=" + kid + "&id=0&FM_soeg=" + fm_soeg + "&forny=" + forny + "", "TimeOut", "width=800,height=780,left=300,top=20,screenX=100,screenY=100,scrollbars=yes")
    } else {
        window.open("serviceaft.asp?func=red&kundeid=" + kid + "&id=" + thisid + "&FM_soeg=" + fm_soeg + "&forny=" + forny + "", "TimeOut", "width=800,height=780,left=300,top=20,screenX=100,screenY=100,scrollbars=yes")
    }
}


function autoopret() {

    //alert("her")
    $("#rdir").val('redjobcontionue')
    $("#showdiv").val('forkalk')
    $("#jobdata").trigger("submit")

}