// JScript File



$(document).ready(function () {

    

    $("select[name*=ajax]").AjaxUpdateField({ parent: "tr", subselector: "td:first > input[name=rowId]" });
    $("#incidentlist").table_sort({ items: 'tbody > tr:gt(1):has(:input[name=rowId])', IdControlNode: "td:first > :input[name=rowId]" });


    //$("#incidentlist").click(function() {
    //    alert("her")
    //});

    $(".jq_akttimerstk").keyup(function () {

        var thisid = this.id
        var thisval = this.value
        var idlngt = thisid.length
        var idtrim = thisid.slice(10, idlngt)

        //alert(idtrim)
        var thisval = $("#FM_aktbgr_" + idtrim).val()
        beregnakttotpris(idtrim, thisval, 0)
    });

    $(".jq_akttotal").keyup(function () {

        var thisid = this.id
        var thisval = this.value
        var idlngt = thisid.length
        var idtrim = thisid.slice(14, idlngt)

        //alert(idtrim)
        var thisval = $("#FM_aktbgr_" + idtrim).val()
        beregnakttotpris(idtrim, thisval, 1)
    });


    function beregnakttotpris(id, val, type) {

        //alert("her")


        var thisval = val
        var idtrim = id

        var varstk = $("#FM_aktstk_" + idtrim).val();
        var vartimer = $("#FM_akttim_" + idtrim).val();
        var varpris = $("#FM_aktpri_" + idtrim).val();


        varpris = varpris.replace(".", "")
        varstk = varstk.replace(".", "")
        vartimer = vartimer.replace(".", "")

        varpris = varpris.replace(",", ".")
        varstk = varstk.replace(",", ".")
        vartimer = vartimer.replace(",", ".")

        //alert(vartimer/1 + " " + varstk/1 + " "+varpris/1)


        /// Timer eller totalpris ændret?
        /// 1 = totalpris ellers 0

        //alert("her: " + type)
        if (type == 0) {

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

            $("#FM_akttotpris_" + idtrim).val(akttotpris)

        } else {

            var akttotpris = $("#FM_akttotpris_" + idtrim).val();
            akttotpris = akttotpris.replace(".", "")
            akttotpris = akttotpris.replace(",", ".")

            if (thisval == 0) {
                varpris = akttotpris
            } else {
                if (thisval == 1) { // timer
                    varpris = (akttotpris / 1 / vartimer / 1)
                } else { //stk.
                    varpris = (akttotpris / 1 / varstk / 1)
                }
            }


            vartimer = String(Math.round(vartimer * 100) / 100).replace(".", ",")   //Math.round(vartimer * 100) / 100
            akttotpris = String(Math.round(akttotpris * 100) / 100).replace(".", ",") //Math.round(akttotpris * 100) / 100
            varpris = String(Math.round(varpris * 100) / 100).replace(".", ",")
            $("#FM_aktpri_" + idtrim).val(varpris)

        }





        //Fasen//
        var fasen = $("#FM_aktfas_h_" + idtrim).val();
        //fasen = fasen.replace("  ", "_")
        //alert(fasen)

        var fanum = $("#FM_sum_aid_fa_" + idtrim).val();
        var antalFa = $("#fa_" + fasen).val();
        //alert(vartimer)

        //på den berørte akt. linie i fase//
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
        $("#slsum_" + fasen).html("<b>" + varFaseSum + "</b>")

        varFaseTimer = String(Math.round(varFaseTimer * 100) / 100).replace(".", ",")
        $("#sltimer_" + fasen).html("<b>" + varFaseTimer + "</b>")

        //alert(fasen)


        /// Total //
        var FaID = $("#fatot_" + fasen).val()
        $("#fatot_val_" + FaID).val(varFaseSum)
        $("#fatottimer_val_" + FaID).val(varFaseTimer)
        var valThis = 0;
        var valThisTimer = 0;
        var varFaseSumTot = 0;
        var varFaseTimerTot = 0;
        var FatotIalt = $("#fatot_ialt").val()

        for (i = 0; i <= FatotIalt; i++) {
            valThis = $("#fatot_val_" + i).val().replace(".", "")
            valThis = valThis.replace(",", ".")
            varFaseSumTot = (varFaseSumTot / 1) + (valThis / 1);

            //alert(varFaseSumTot + "#" + valThis)

            valThisTimer = $("#fatottimer_val_" + i).val().replace(".", "")
            valThisTimer = valThisTimer.replace(",", ".")
            varFaseTimerTot = (varFaseTimerTot / 1) + (valThisTimer / 1);
        }


        //alert(valThisTimer + "# "+ varFaseTimerTot)

        varFaseSumTot = String(Math.round(varFaseSumTot * 100) / 100).replace(".", ",")
        $("#FM_fasersumtot").val(varFaseSumTot)
        $("#fasersumtot").html("<b>" + varFaseSumTot + "</b>")


        varFaseTimerTot = String(Math.round(varFaseTimerTot * 100) / 100).replace(".", ",")
        $("#FM_fasertimertot").val(varFaseTimerTot)
        $("#fasertimertot").html("<b>" + varFaseTimerTot + "</b>")

        diff_timer_sum();


    }


    //$("#syncjob").click(function () {
    //
     //   var varAkttimer = $("#akttimer").val()
    //    var varAktsum = $("#aktbudgetsum").val()
    //    $("#FM_budgettimer_s").val(varAkttimer)
    //    $("#FM_budget_s").val(varAktsum)
    //    $("#FM_budgettimer").val(varAkttimer)
    //    $("#FM_budget").val(varAktsum)

    //    
    //});

    $("#sync").click(function () {

        var varFaseSumTot = $("#FM_fasersumtot").val()
        //varFaseSumTot = String(Math.round(varFaseSumTot * 100) / 100).replace(".", ",")
        $("#FM_interntbelob").val(varFaseSumTot)

        var varFaseTimerTot = $("#FM_fasertimertot").val()
        //varFaseTimerTot = String(Math.round(varFaseTimerTot * 100) / 100).replace(".", ",")
        $("#FM_budgettimer").val(varFaseTimerTot)


        varFaseSumTot = varFaseSumTot.replace(".", "")
        varFaseSumTot = varFaseSumTot.replace(",", ".")

        varFaseTimerTot = varFaseTimerTot.replace(".", "")
        varFaseTimerTot = varFaseTimerTot.replace(",", ".")

        var varFaktor = $("#FM_intfaktor").val().replace(".", "")
        varFaktor = varFaktor.replace(",", ".")

        var gnsTp = (varFaseSumTot / (varFaseTimerTot * varFaktor))
        gnsTp = String(Math.round(gnsTp * 100) / 100).replace(".", ",")
        $("#FM_gnsinttpris").val(gnsTp)

        var pb_tg_timepris = (varFaseTimerTot * varFaktor)
        pb_tg_timepris = String(Math.round(pb_tg_timepris * 100) / 100).replace(".", "")
        $("#pb_tg_timepris").val(pb_tg_timepris)

        $.scrollTo('200px', 1200);

        diff_timer_sum();

    });



    function diff_timer_sum() {


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


        if (diff_sum != "0") {

            $("#diff_sum_ok").hide("fast");

        } else {

            $("#diff_sum_ok").css("display", "");
            $("#diff_sum_ok").css("visibility", "visible");
            $("#diff_sum_ok").show("fast");

        }

    }





    $(".faseoskrift_navn").keyup(function () {

        var thisid = this.id
        var thisval = this.value

        var fasen = thisid //$("#FM_aktfas_" + thisid).val();

        var antalFa = $("#fa_" + fasen).val();

        $(".aktFase_" + fasen + "").val(thisval);


    });


    $(".bgr").change(function () {

        var thisid = this.id
        var thisval = this.value
        var idlngt = thisid.length
        var idtrim = thisid.slice(10, idlngt)

        beregnakttotpris(idtrim, thisval, 0)
    });


    $(".faseoskrift").change(function () {

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


    $(".faseoskrift_slet").click(function () {

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





});


   