// JScript File



//$(document).ready(function () {


  


    $(".jq_akttimerstk").keyup(function () {


        var thisid = this.id
        var idlngt = thisid.length
        var idtrim = thisid.slice(0, 10)

        if (idtrim == 'FM_aktpri_' && $("#showtpalert").val() == '0') {

            //alert("Husk at tilrette timeprisen på medarbejderne (faneblad i toppen), så de passer til dit budget på aktivitetslinien.")

            $("#showtpalert").val('1')
        }
    });


    $(".jq_akttimerstk").keyup(function () {

        var thisid = this.id
        var thisval = this.value
        var idlngt = thisid.length
        var idtrim = thisid.slice(10, idlngt)

        var timellertpris = thisid.slice(0, 9)

        if (timellertpris != 'FM_aktpri') {
           var thistype = 0 //beregnerudfra timer
        } else {
           var thistype = 2 //beregnerudfra timepris
        }

        //alert(timellertpris)

        var thisval = $("#FM_aktbgr_" + idtrim).val()

        beregnakttotpris(idtrim, thisval, thistype)
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

        //alert("her be tot")


        var thisval = val
        var idtrim = id

        var varstk = $("#FM_aktstk_" + idtrim).val();
        var vartimer = $("#FM_akttim_" + idtrim).val();
        var varpris = $("#FM_aktpri_" + idtrim).val();

        var lto = $("#lto").val();

        //Dobblet hvis det er mere end 1mio og dermed 2 punktummer
        varpris = varpris.replace(".", "")
        varstk = varstk.replace(".", "")
        vartimer = vartimer.replace(".", "")
        varpris = varpris.replace(".", "")
        varstk = varstk.replace(".", "")
        vartimer = vartimer.replace(".", "")

        varpris = varpris.replace(",", ".")
        varstk = varstk.replace(",", ".")
        vartimer = vartimer.replace(",", ".")

        varpris = Math.round(varpris)
        varstk = Math.round(varstk)
        vartimer = Math.round(vartimer)
       
        //alert(vartimer / 1 + " " + varstk / 1 + " " + varpris / 1 + " thisval : " + thisval + " type: " + type)


        /// Timer eller totalpris ændret?
        /// 1 = totalpris 
        /// 0 = timer
        /// 2 = timepris/stkpris ændret 
        /// ellers 0 (hvis timer ell. timepris ændret)

        //alert("her: " + type + " lto: " + lto)
        // HVIS LTO tottal beløb disabled. KAN ALdrig blive type = 1
        if (type == 0 || type == 2) {

            //alert(lto)

            budgetakt = $("#budgetakt_" + idtrim).val();

            //if (lto == "intranet - local" || lto == "oko") {
                //kontonrslået til - beregn timer pbg. budget

           if (budgetakt == 1) { //Total beløb ghostet, bereng timer pga. af totalbeløb og timepris

               var akttotpris = $("#FM_akttotpris_" + idtrim).val();
               //alert(akttotpris + " idtrim: " + idtrim)

               if (akttotpris == "NaN" || (akttotpris == "-Infinity") || akttotpris == "") {
                   akttotpris = 0
               } else {
                   akttotpris = akttotpris.replace(".", "")
                   akttotpris = akttotpris.replace(".", "")
                   akttotpris = akttotpris.replace(",", ".")
                   //akttotpris = Math.round(akttotpris * 100) / 100
               }

               
               //alert(type)

               if (type == 0) { //beregner tp udfra timer


                   if (vartimer != 0) {
                       varpris = (akttotpris / 1 / vartimer / 1)
                   } else {
                       varpris = (akttotpris / 1)
                   }


                   varpris = String(Math.round(varpris * 100) / 100).replace(".", ",")   //Math.round(vartimer * 100) / 100
                   $("#FM_aktpri_" + idtrim).val(varpris);


               } else { //beregner timer udfra tp

                   if (thisval == 1) { // timer el. stk.

                    if (varpris != 0) {
                        vartimer = (akttotpris / 1 / varpris / 1)
                    } else {
                        vartimer = (akttotpris / 1)
                    }

              
                    vartimer = String(Math.round(vartimer * 100) / 100).replace(".", ",")   //Math.round(vartimer * 100) / 100
                    $("#FM_akttim_" + idtrim).val(vartimer);


                } else { //stk.
                    if (varpris != 0) {
                        varstk = (akttotpris / 1 / varpris / 1)
                    } else {
                        varstk = (akttotpris / 1)
                    }

                    varstk = String(Math.round(varpris * 100) / 100).replace(".", ",")
                    $("#FM_aktstk_" + idtrim).val(varstk);
                }
                  

               }
                

           } else { //budgetakt alm. beregner totalpris



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

           }  //budgetakt

        } else { //type = 1



            var akttotpris = $("#FM_akttotpris_" + idtrim).val();
            akttotpris = akttotpris.replace(".", "")
            akttotpris = akttotpris.replace(".", "")
            akttotpris = akttotpris.replace(",", ".")
            akttotpris = Math.round(akttotpris)
           

                    if (thisval == 0) {
                        varpris = akttotpris
                    } else {
                        if (thisval == 1) { // timer
                            if (vartimer != 0) {
                                varpris = (akttotpris / 1 / vartimer / 1)
                            } else {
                                varpris = (akttotpris / 1)
                            }
                        } else { //stk.
                            if (varstk != 0) {
                                varpris = (akttotpris / 1 / varstk / 1)
                            } else {
                                varpris = (akttotpris / 1)
                            }
                        }
                    }


                    vartimer = String(Math.round(vartimer * 100) / 100).replace(".", ",")   //Math.round(vartimer * 100) / 100
                    akttotpris = String(Math.round(akttotpris * 100) / 100).replace(".", ",") //Math.round(akttotpris * 100) / 100
                    varpris = String(Math.round(varpris * 100) / 100).replace(".", ",")
                    $("#FM_aktpri_" + idtrim).val(varpris)

           



        }


        //alert("lto:"+ lto + ""+ vartimer + " " + akttotpris + " " + varpris + " fs: #FM_aktfas_h_" + idtrim + " ::" + $("#FM_aktfas_h_" + idtrim).val())

       


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

        //alert("fasen: "+ fasen + " antalFa: " + antalFa)
        var varFaseTimer = 0;
        var varFaseSum = 0;
        var valThis = 0;
        var valThisTimer = 0;

       
        for (i = 0; i < antalFa; i++) {

            //alert($("#af_sum_" + fasen + "_" + i).val() + " :: " + i)

            valThis = $("#af_sum_" + fasen + "_" + i).val()
            

            if (valThis == "NaN" || (valThis == "-Infinity") || valThis == "" || String($("#af_sum_" + fasen + "_" + i).val()) == "undefined") {
                valThis = 0
                varFaseSum = (varFaseSum/1)
            } else {

                //alert(valThis)
                valThis = valThis.replace(".", "")
                valThis = valThis.replace(".", "")
                valThis = valThis.replace(",", ".")
                valThis = Math.round(valThis)
                varFaseSum = (varFaseSum/1) + (valThis / 1);
            }

            
            //alert("her OK " + i)

           

            valThisTimer = $("#af_timer_" + fasen + "_" + i).val()
            

            if (valThisTimer == "NaN" || (valThisTimer == "-Infinity") || valThisTimer == "" || String($("#af_timer_" + fasen + "_" + i).val()) == "undefined") {
                valThisTimer = 0
                varFaseTimer = (varFaseTimer/1) 
            } else {
                valThisTimer = valThisTimer.replace(".", "")
                valThisTimer = valThisTimer.replace(".", "")
                valThisTimer = valThisTimer.replace(",", ".")
                valThisTimer = Math.round(valThisTimer)

                varFaseTimer = (varFaseTimer / 1) + (valThisTimer / 1);
            }

            //alert(valThisTimer)
           

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

        //alert("HER:" +FatotIalt)

        for (i = 0; i <= FatotIalt; i++) {

            //alert("raw: "+ $("#fatot_val_" + i).val())

            valThis = $("#fatot_val_" + i).val()
            
            if (valThis == "NaN" || (valThis == "-Infinity") || valThis == "") {
                valThis = 0
            } else {
                valThis = valThis.replace(".", "")
                valThis = valThis.replace(".", "")
                valThis = valThis.replace(",", ".")
                valThis = Math.round(valThis)
                varFaseSumTot = (varFaseSumTot / 1) + (valThis / 1);
            }

            //alert("After: " + valThis)

           
           

            valThisTimer = $("#fatottimer_val_" + i).val()
            if (valThisTimer == "NaN" || (valThisTimer == "-Infinity") || valThisTimer == "") {
                valThisTimer = 0
            } else {
                valThisTimer = valThisTimer.replace(".", "")
                valThisTimer = valThisTimer.replace(".", "")
                valThisTimer = valThisTimer.replace(",", ".")
                valThisTimer = Math.round(valThisTimer)
                varFaseTimerTot = (varFaseTimerTot / 1) + (valThisTimer / 1);
            }
            

            
        }


        //alert(valThisTimer + "# "+ varFaseTimerTot)

        varFaseSumTot = String(Math.round(varFaseSumTot * 100) / 100).replace(".", ",")
        $("#FM_fasersumtot").val(varFaseSumTot)
        $("#fasersumtot").html("<b>" + varFaseSumTot + "</b>")


        varFaseTimerTot = String(Math.round(varFaseTimerTot * 100) / 100).replace(".", ",")
        $("#FM_fasertimertot").val(varFaseTimerTot)
        $("#fasertimertot").html("<b>" + varFaseTimerTot + "</b>")



        // Alert på konto hvis faser overstiger SUM total
        if (lto == 'intranet - local' || lto == 'oko') {

            //KONTO//
            var konto = $("#FM_aktkonto_" + idtrim).val();
            //fasen = fasen.replace("  ", "_")
            //alert(konto)

            //// TIMER, TP SUM TJK /////////////
            $(".FM_kontot").val(0)
            $(".FM_kontotp").val(0)
            $(".FM_kontosum").val(0)
            $(".kontoTclslinje").each(function () {

                //alert("her")

                var thisid = this.id
                var idlngt = thisid.length
                var idtrim = thisid.slice(10, idlngt)

                thisKonto = $("#FM_aktkonto_" + idtrim).val()

                preval = $("#FM_kontot_" + thisKonto).val().replace(",", ".")/1
                newval = preval + $(this).val().replace(",", ".")/1
                newval = String(Math.round(newval * 100) / 100).replace(".", ",")

                //alert("this: " + this + " idtrim:" + idtrim + " Konto: " + thisKonto + " timer: " + $(this).val() + " newval: " + newval);
                
                $("#FM_kontot_" + thisKonto).val(newval)


                prevalTp = $("#FM_kontotp_" + thisKonto).val().replace(",", ".") / 1
                newvalTp = prevalTp + $("#FM_aktpri_" + idtrim).val().replace(",", ".") / 1
                newvalTp = String(Math.round(newvalTp * 100) / 100).replace(".", ",")

                //alert("idtrim:" + idtrim + " Konto: " + thisKonto + " timer: " + $(this).val() + " newval: " + newval);

                $("#FM_kontotp_" + thisKonto).val(newvalTp)

               

                prevalSum = $("#FM_kontosum_" + thisKonto).val().replace(",", ".") / 1
                newvalSum = prevalSum + $("#FM_akttotpris_" + idtrim).val().replace(",", ".") / 1
                newvalSum = String(Math.round(newvalSum * 100) / 100).replace(".", ",")

                //alert("idtrim:" + idtrim + " Konto: " + thisKonto + " timer: " + $(this).val() + " newval: " + newval);

                $("#FM_kontosum_" + thisKonto).val(newvalSum)

                //alert("her")

            });


          

            // Tjek Totaler og Alerts
            baseLineVal = 0;
            baseLineValTp = 0;
            baseLineValSum = 0;
            $(".kontoTclsSum").each(function () {

             
                var thisid = this.id
                var idlngt = thisid.length
                var idtrim = thisid.slice(10, idlngt)


                thisKonto = $("#FM_aktkonto_" + idtrim).val()

                // Timer
                baseLineVal = $("#FM_akttim_" + idtrim).val().replace(".", "")
                baseLineVal = baseLineVal.replace(",", ".") / 1
                totVal = $("#FM_kontot_" + thisKonto).val().replace(".", "")
                totVal = totVal.replace(",", ".") / 1

                if (baseLineVal < totVal) {
                    $("#FM_akttim_" + idtrim).css("background-color", "lightpink")
                } else {
                    $("#FM_akttim_" + idtrim).css("background-color", "")
                }


                //TP
                baseLineValTp = $("#FM_aktpri_" + idtrim).val().replace(".", "")
                baseLineValTp = baseLineValTp.replace(",", ".") / 1
                totValTp = $("#FM_kontotp_" + thisKonto).val().replace(".", "")
                totValTp = totValTp.replace(",", ".") / 1

                if (baseLineValTp < totValTp) {
                    $("#FM_aktpri_" + idtrim).css("background-color", "lightpink")
                } else {
                    $("#FM_aktpri_" + idtrim).css("background-color", "")
                }


                //SUM
                //alert($("#FM_akttotpris_" + idtrim).val())

                baseLineValSum = $("#FM_akttotpris_" + idtrim).val().replace(".", "")
                baseLineValSum = baseLineValSum.replace(",", ".")/1
                totValSum = $("#FM_kontosum_" + thisKonto).val().replace(".", "") 
                totValSum = totValSum.replace(",", ".")/1

                alert(baseLineValSum + " < " + totValSum)

                if (baseLineValSum < totValSum) {
                    $("#FM_akttotpris_" + idtrim).css("background-color", "lightpink")
                } else {
                    $("#FM_akttotpris_" + idtrim).css("background-color", "#999999")
                }


            });


            


        } //LTO OKO


        //diff_timer_sum();
        syncjoga();


    }


   

    $("#sync").click(function () {

        if ($("#sync").is(':checked') == true) {
            $.scrollTo('200px', 1200);
            syncjoga();
        }
        //diff_timer_sum();

    });


    function syncjoga() {

        if ($("#sync").is(':checked') == true) {

            //alert("her")

            var varFaseSumTot = $("#FM_fasersumtot").val()
            //varFaseSumTot = String(Math.round(varFaseSumTot * 100) / 100).replace(".", ",")
            $("#FM_interntbelob").val(varFaseSumTot)

            var varFaseTimerTot = $("#FM_fasertimertot").val()
            //alert(varFaseTimerTot)
            //varFaseTimerTot = String(Math.round(varFaseTimerTot * 100) / 100).replace(".", ",")
            $("#FM_budgettimer").val(varFaseTimerTot)


            varFaseSumTot = varFaseSumTot.replace(".", "")
            varFaseSumTot = varFaseSumTot.replace(".", "")
            varFaseSumTot = varFaseSumTot.replace(",", ".")

            varFaseTimerTot = varFaseTimerTot.replace(".", "")
            varFaseTimerTot = varFaseTimerTot.replace(".", "")
            varFaseTimerTot = varFaseTimerTot.replace(",", ".")

            var varFaktor = $("#FM_intfaktor").val().replace(".", "")
            varFaktor = varFaktor.replace(",", ".")

            if ((varFaseTimerTot * varFaktor) > 0) {
                vsFsFak = varFaseTimerTot * varFaktor

                //alert(vsFsFak)

                var gnsTp = (varFaseSumTot / (vsFsFak))
                gnsTp = String(Math.round(gnsTp * 100) / 100).replace(".", ",")
                $("#FM_gnsinttpris").val(gnsTp)

                var pb_tg_timepris = (vsFsFak)
                pb_tg_timepris = String(Math.round(pb_tg_timepris * 100) / 100).replace(".", "")
                $("#pb_tg_timepris").val(pb_tg_timepris)

            } else {

                $("#FM_gnsinttpris").val('0')
                $("#pb_tg_timepris").val('0')



            }



            //$.scrollTo('200px', 1200);

            //diff_timer_sum();
            //alert("her")
            beregninttp();

            //beregnintbelob()
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
            bgval = "lightpink"
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





//});


   