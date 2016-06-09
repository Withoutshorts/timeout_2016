


//function mOvr(divId, src, clrOver) {
//    src.bgColor = clrOver;
//}

//function mOut(src, clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn; } }




//$(".dv_besk").hide("fast");

$(document).ready(function() {


    function visbeskfelter() {

        if ($("#visbesk").is(':checked') == true) {

            $(".dv_besk").css("display", "");
            $(".dv_besk").css("visibility", "visible");
            $(".dv_besk").show("fast");

            //$.cookie("ck_visbesk", '1');

        } else {

            $(".dv_besk").hide("fast");
            //$.cookie("ck_visbesk", '0');
        }
    }


    $("#visbesk").click(function () {
        visbeskfelter();
    });


    //Skal besk. felter editor vises ved load ==> ALTDI LUKKET //
    //if ($.cookie("ck_visbesk") == "1") {
    //    $("#visbesk").attr('checked', true);
    //} else {
    //    $("#visbesk").attr('checked', false);
    //}

    // visbeskfelter();



//$(".drag").click(function () {
    //    $("select[name*=ajax]").AjaxUpdateField({ parent: "tr", subselector: "td:first > input[name=rowId]" });
    //$("#incidentlist").table_sort({ items: 'tbody > tr:gt(1):has(:input[name=rowId])', IdControlNode: "td:first > :input[name=rowId]" });
 //});

     //$(".drag").mouseover(function () {
    //    $(this).css('cursor', 'pointer');  
    //});


    $(".faseoskrift_slet").click(function() {

        var thisid = this.id

        var idlngt = thisid.length
        var idtrim = thisid.slice(3, idlngt)

        //alert(idtrim)
        var antalFa = $("#fa_" + idtrim).val();
        //alert("her")
        for (i = 0; i < antalFa; i++) {

            if ($(this).is(':checked') == true) {
                $("#af_" + thisid + "_" + i + "").attr('checked', true);
            } else {
                $("#af_" + thisid + "_" + i + "").attr('checked', false);

            }

        }


    });




    $(".faseoskrift").change(function() {

        var thisid = this.id
        var thisval = this.value
        //alert(this.selectedValues())

        var bgval
        if (thisval == 0) {
            bgval = "crimson"
        }
        if (thisval == 1) {
            bgval = "DarkSeaGreen"
        }
        if (thisval == 2) {
            bgval = "#cccccc"
        }

        var antalFa = $("#fa_" + thisid).val();

        for (i = 0; i < antalFa; i++) {
            $("#af_" + thisid + "_" + i + "").val(thisval)
            $("#af_" + thisid + "_" + i + "").css("background-color", bgval);

        }


    });


    $(".fs_dt_st").keyup(function () {
    //alert("her")

        var thisid = this.id
        var thisval = this.value
        var idlngt = thisid.length
        var idtrim = thisid.slice(6, idlngt)

        thisid = idtrim
        //alert(this.selectedValues())

     
        var antalFa = $("#fa_" + thisid).val();

        for (i = 0; i < antalFa; i++) {
            $("#af_st_dato_" + thisid + "_" + i + "").val(thisval)
           
        }


    });


    $(".fs_dt_sl").keyup(function () {
        //alert("her")

        var thisid = this.id
        var thisval = this.value
        var idlngt = thisid.length
        var idtrim = thisid.slice(6, idlngt)

        thisid = idtrim
        //alert(this.selectedValues())


        var antalFa = $("#fa_" + thisid).val();

        for (i = 0; i < antalFa; i++) {
            $("#af_sl_dato_" + thisid + "_" + i + "").val(thisval)

        }


    });

    
    


    
    $("#vispasluk").click(function () {

        $("#form_aktlist").trigger("submit")
    });



    //$("#synctimertiljob").click(function () {
    //    $("#form_aktlist").trigger("submit")
    //});

    
   


    $(".showhidefase").click(function() {

        //alert("her")
        var thisid = this.id
        var idlngt = thisid.length
        var hd_jobid = $("#hd_jobid").val()
        //alert(idlngt)
        var idtrim = thisid.slice(13, idlngt)
        //alert(idtrim)

        //$(".trfase_" + idtrim + "").hide(1000);

        if ($(".trfase_" + idtrim + "").css('display') == "none") {

            $(".trfase_" + idtrim + "").css("display", "");
            $(".trfase_" + idtrim + "").css("visibility", "visible");
            $(".trfase_" + idtrim + "").show("fast");
            $.scrollTo('+=100px', 1500);
            $.cookie("showfase_" + hd_jobid + "_" + idtrim + "", '1');


        } else {

            $(".trfase_" + idtrim + "").hide("fast");
            $.scrollTo('-=100px', 800);
            $.cookie("showfase_" + hd_jobid + "_" + idtrim + "", '0');

        }



    });




    $(".faseoskrift_navn").keyup(function() {


        //alert("her")
        var thisid = this.id
        var thisval = this.value
        var idlngt = thisid.length
        var idtrim = thisid.slice(8, idlngt)
        thisid = idtrim

        //var n = $("#" + thisid + ":checked").length;
        //alert(thisid + " "+ thisval)
        //var fasen = $("#FM_sum_aid_fs_" + idtrim).val();
        var antalFa = $("#fa_" + thisid).val();
        for (i = 0; i < antalFa; i++) {
            $("#af_n_" + thisid + "_" + i + "").val(thisval)
            //$("#af_n_" + thisid + "_" + i + "").css("background-color", bgval);

        }


    });

    function beregnakttotpris(id, val) {

       

        var thisid = id
        var thisval = val
        var idlngt = id.length
        var idtrim = id.slice(7, idlngt)

        var varstk = $("#FM_stk_" + idtrim).val();
        var vartimer = $("#FM_tim_" + idtrim).val();
        var varpris = $("#FM_pri_" + idtrim).val();

        varpris = varpris.replace(".", "")
        varstk = varstk.replace(".", "")
        vartimer = vartimer.replace(".", "")

        varpris = varpris.replace(",", ".")
        varstk = varstk.replace(",", ".")
        vartimer = vartimer.replace(",", ".")

        //alert(vartimer/1 + " " + varstk/1 + " "+varpris/1)

        if (thisval == 0) {
            akttotpris = (varpris / 1)
        } else {
            if (thisval == 1) {
                akttotpris = (vartimer / 1 * varpris / 1)
            } else {
                akttotpris = (varstk / 1 * varpris / 1)
            }
        }

        akttotpris = String(Math.round(akttotpris * 100) / 100).replace(".", ",")
        vartimer = String(Math.round(vartimer * 100) / 100).replace(".", ",")

        $("#akt_totpris_" + idtrim).html("<b>" + akttotpris + "</b> " & basisValISO)
        $("#FM_akt_totpris_" + idtrim).val("#" + akttotpris)

        //Fasen//

        var fasen = $("#FM_sum_aid_fs_" + idtrim).val();
        var fanum = $("#FM_sum_aid_fa_" + idtrim).val();
        var antalFa = $("#fa_" + fasen).val();
        //alert(fanum)
        $("#af_sum_" + fasen + "_" + fanum).val(akttotpris)
        $("#af_timer_" + fasen + "_" + fanum).val(vartimer)

        //alert(fasen)
        var varFaseTimer = 0;
        var varFaseSum = 0;
        var valThis = 0;
        var valThisTimer = 0;
        for (i = 0; i < antalFa; i++) {
            valThis = $("#af_sum_" + fasen + "_" + i).val().replace(".", "")
            valThis = valThis.replace(",", ".")
            varFaseSum = (varFaseSum / 1) + (valThis / 1);

            valThisTimer = $("#af_timer_" + fasen + "_" + i).val().replace(".", "")
            valThisTimer = valThisTimer.replace(",", ".")

            //alert(valThisTimer)
            varFaseTimer = (varFaseTimer / 1) + (valThisTimer / 1);

        }

        varFaseSum = String(Math.round(varFaseSum * 100) / 100).replace(".", ",")
        $("#slsum_" + fasen).html("<b>" + varFaseSum + "</b> " & basisValISO)

        varFaseTimer = String(Math.round(varFaseTimer * 100) / 100).replace(".", ",")
        $("#sltimer_" + fasen).html("<b>" + varFaseTimer + "</b> timer")

        /// Total //
        var FaID = $("#fatot_" + fasen).val()
        $("#fatot_val_" + FaID).val(varFaseSum)
        $("#fatottimer_val_" + FaID).val(varFaseTimer)
        valThis = 0;
        valThisTimer = 0;
        var varFaseSumTot = 0;
        var varFaseTimerTot = 0;
        var FatotIalt = $("#fatot_ialt").val()
        for (i = 0; i <= FatotIalt; i++) {
            valThis = $("#fatot_val_" + i).val().replace(".", "")
            valThis = valThis.replace(",", ".")
            varFaseSumTot = (varFaseSumTot / 1) + (valThis / 1);

            valThisTimer = $("#fatottimer_val_" + i).val().replace(".", "")
            valThisTimer = valThisTimer.replace(",", ".")
            varFaseTimerTot = (varFaseTimerTot / 1) + (valThisTimer / 1);
        }

        varFaseSumTot = String(Math.round(varFaseSumTot * 100) / 100).replace(".", ",")
        $("#fasersumtot").html("<b>" + varFaseSumTot + "</b> " & basisValISO)

        varFaseTimerTot = String(Math.round(varFaseTimerTot * 100) / 100).replace(".", ",")
        $("#fasertimertot").html("<b>" + varFaseTimerTot + "</b> timer")

    }


    $(".bgr").change(function() {
        var thisid = this.id
        var thisval = this.value
        beregnakttotpris(thisid, thisval)
    });


    $(".timstkpris").keyup(function() {
       // alert(this.id)
        var thisid = this.id
        var idlngt = thisid.length

        //alert(idlngt)
        var idtrim = thisid.slice(7, idlngt)
        var thisval = $("#FM_bgr_" + idtrim).val()
        beregnakttotpris(thisid, thisval)
    });



    // Åbne og lukke faser ved load //
    function showfaservload() {
    var hd_jobid = $("#hd_jobid").val()
    var antalFa = $("#fatot_ialt").val()
    
    for (i = 0; i <= antalFa; i++) {
            //alert(i)
            var fasenavn = $("#fatotbn_"+i+"").val()
            //var idlngt = id.length
            //var idtrim = id.slice(6, idlngt)
            //var fasenavn =
            //alert(fasenavn)
            //alert($.cookie("showfase_" + hd_jobid + "_" + fasenavn + ""))
            if ($.cookie("showfase_" + hd_jobid + "_" + fasenavn + "") == "1") {
                $(".trfase_" + fasenavn + "").css("display", "");
                $(".trfase_" + fasenavn + "").css("visibility", "visible");
                $(".trfase_" + fasenavn + "").show("fast");
            }

    }
   
    }

    
    showfaservload()

   
  
});



