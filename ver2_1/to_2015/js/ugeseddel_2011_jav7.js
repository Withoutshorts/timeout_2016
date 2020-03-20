



//BRUGES AF UGESEDDEL og TOUCHMONITOR 





$(document).ready(function() {




    mobil_week_reg_job_dd = $("#mobil_week_reg_job_dd").val()
    //alert("her søger: " + mobil_week_reg_job_dd)
    if (mobil_week_reg_job_dd == 1) {


        sogjobogkunde(0);


    }

    $("#dv_akt_0").prop("disabled", false);
    sogakt(0);


    $('.date').datepicker({

    });


    function toSeconds(time_str) {
        // Extract hours, minutes and seconds
        var parts = time_str.split(':');
        // compute  and return total seconds
        return parts[0] * 3600 + // an hour has 3600 seconds
            parts[1] * 60 +   // a minute has 60 seconds
            +"00";        // parts[2] seconds
    }

    $("#FM_sttid, #FM_sltid").keyup(function () {


        sttid = $("#FM_sttid").val()
        sltid = $("#FM_sltid").val()

        //timerThis = 3
        toSecondsA = toSeconds(sttid)
        toSecondsB = toSeconds(sltid)

        var difference = 0
        if (toSecondsA - toSecondsB > 0) {
            difference = 0
        } else {
            difference = Math.abs(toSecondsA - toSecondsB);
        }

        // compute hours, minutes and seconds
        var result = [
            // an hour has 3600 seconds so we have to compute how often 3600 fits
            // into the total number of seconds
            Math.floor(difference / 3600), // HOURS
            // similar for minutes, but we have to "remove" the hours first;
            // this is easy with the modulus operator
            Math.floor((difference % 3600) / 60), // MINUTES
            // the remainder is the number of seconds
            difference % 60 // SECONDS
        ];

        // formatting (0 padding and concatenation)
        result = result.map(function (v) {
            return v < 10 ? '0' + v : v;
        }).join(':');

        $("#FM_timer").val(result);

        var result60 = result //.Tostring()
        var result100Arr = result60.split(':');
        //alert(result100Arr[0])
        var hoursIn100 = result100Arr[0]

        if (hoursIn100 < 10) {
            hoursIn100 = hoursIn100.substr(hoursIn100.length - 1);
        }

        var minIn100 = result100Arr[1] * (100 / 60)
        minIn100 = Math.floor(minIn100)

        if (minIn100 < 10) {
            minIn100 = "0" + minIn100;
        }


        if (hoursIn100 > 24) {
            hoursIn100 = 0
        }

        if (minIn100 > 60) {
            minIn100 = 0
        }

        if (hoursIn100 == 'NaN' || minIn100 == 'NaN') {
            var result100 = 0
        } else {
            var result100 = hoursIn100 + "," + minIn100
        }


        $("#FM_timer").val(result100);

    });



    $(".status_all").click(function () {



        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(11, thisvallngt)
        thisval = thisvaltrim

        //alert("Her:" + thisval)

        if ($("#status_all_" + thisval).is(':checked') == true) {

            $(".status_chk_" + thisval).attr("checked", "checked");

        } else {

            $(".status_chk_" + thisval).removeAttr("checked");

        }


    });





    $("#a_epinote").click(function () {


        $("#d_epinote").hide();

    });




    $("#FM_visallemedarb").click(function () {


        var sesid = $("#FM_sesMid").val()
        $("#usemrn").val(sesid)

        $("#filterkri").submit();


    });

    $("#usemrn").change(function () {

        $("#filterkri").submit();
    });



    $("#dv_akt_0").change(function () {

        jq_aktidc = $("#dv_akt_0").val()
        //alert(jq_jobidc)
        $.post("?jq_aktidc=" + jq_aktidc, { control: "FN_gemaktc", AjaxUpdateField: "true" }, function (data) {

        });

    });


    //Kun ved job = DD
    $("#dv_job_0").change(function () {

        //alert("søger")


        $("#btn_indlas").prop("disabled", false);


        mobil_week_reg_job_dd = $("#mobil_week_reg_job_dd").val()
        mobil_week_reg_akt_dd = $("#mobil_week_reg_akt_dd").val()

        if (mobil_week_reg_job_dd == "1") {

            if (mobil_week_reg_akt_dd != "1") {
                $("#FM_akt_0").val('')
                $("#dv_akt_0").hide();
            } else {

                $("#dv_akt_0").prop("disabled", false);
            }


            sogakt(0);

            if (mobil_week_reg_akt_dd != "1") {

                aktSogVal = $("#FM_akt_0").val()
                //alert(aktSogVal.length)

                if (aktSogVal.length == 0) {
                    $("#FM_aktid_0").val('0');
                    $("#dv_akt_0").hide();
                }
            }

        }


        jq_jobidc = $("#dv_job_0").val()
        //alert("HER2")
        //alert(jq_jobidc)
        $.post("?jq_jobidc=" + jq_jobidc, { control: "FN_gemjobc", AjaxUpdateField: "true" }, function (data) {

            //alert("HER OK")
            //$("#test").val(data);

        });

    });



    $(".FM_job").keyup(function () {


        mobil_week_reg_akt_dd = $("#mobil_week_reg_akt_dd").val()

        if (mobil_week_reg_akt_dd != "1") {
            $("#dv_akt_0").hide();
        }

        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(7, thisvallngt)
        thisval = thisvaltrim

        //alert(thisval)

        sogjobogkunde(thisval);

        if (mobil_week_reg_akt_dd == "1") {
            $("#dv_akt_0").prop("disabled", false);
        }

    });





    $(".FM_akt").keyup(function () {


        $(".dv_job").hide();

        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(7, thisvallngt)
        thisval = thisvaltrim

        sogakt(thisval);

    });


    $("#FM_timer").focus(function () {
        $(".dv_job").hide();
        $(".dv_akt").hide();
    });

    $("#FM_kom").focus(function () {
        $(".dv_job").hide();
        $(".dv_akt").hide();
    });







function sogakt(thisval) {


        mobil_week_reg_akt_dd = $("#mobil_week_reg_akt_dd").val()
        mobil_week_reg_job_dd = $("#mobil_week_reg_job_dd").val()

        //alert(mobil_week_reg_job_dd + " # " + mobil_week_reg_akt_dd)

        if (mobil_week_reg_akt_dd != "1") {
            jq_newfilterval = $("#FM_akt_" + thisval).val()

            if (mobil_week_reg_job_dd != "1") {
                jq_jobid = $("#FM_jobid_" + thisval).val()
            } else {
                jq_jobid = $("#dv_job_0").val()
            }


        } else {
            jq_newfilterval = "-1"
            jq_jobid = $("#dv_job_0").val()
        }



        jq_medid = $("#FM_medid").val()
        jq_aktid = $("#FM_akt_" + thisval).val()
        jq_pa = $("#FM_pa").val()

        varTjDatoUS_man = $("#varTjDatoUS_man").val()

        //$(".chbox_job").hide();
if (mobil_week_reg_akt_dd != "1") {

            $("#dv_akt_" + thisval).css("display", "");
            $("#dv_akt_" + thisval).css("visibility", "visible");
            $("#dv_akt_" + thisval).show(100);

        }

        jq_aktidc = $("#jq_aktidc").val()

        //alert("#" + thisval)
        //alert(jq_jobid + ":: medid: " + jq_medid + " jq_aktid: " + jq_aktid + "## jq_pa" + jq_pa)
        //alert("#" + jq_newfilterval.length + "")

        if (jq_newfilterval.length > 0 || mobil_week_reg_akt_dd == "1") {

            $.post("?jq_newfilterval=" + jq_newfilterval + "&jq_jobid=" + jq_jobid + "&jq_medid=" + jq_medid + "&jq_aktid=" + jq_aktid + "&jq_pa=" + jq_pa + "&varTjDatoUS_man=" + varTjDatoUS_man + "&jq_aktidc=" + jq_aktidc, { control: "FN_sogakt", AjaxUpdateField: "true" }, function (data) {
                //alert("cc")
                //$("#dv_akt_test").html(data);
                $("#dv_akt_" + thisval).html(data);

                //alert("END")


                if (mobil_week_reg_akt_dd != "1") {


                    $(".chbox_akt").bind('keyup', function () {


                        if (window.event.keyCode == "13") {

                            var thisvaltrim = $("#dv_akt_" + thisval).val()
                            thisAktidUse = thisvaltrim
                            //thisJobnavn = $("#hiddn_job_" + thisJobid).val()
                            thisAktnavn = $("#dv_akt_" + thisval + " option:selected").text()

                            $("#FM_akt_" + thisval).val(thisAktnavn)
                            $("#FM_aktid_" + thisval).val(thisAktidUse)
                            //$(".dv_akt").hide();
                            $(".chbox_akt").hide();

                        }

                    });


                    $(".chbox_akt").bind('click', function () {


                        var thisvaltrim = $("#dv_akt_" + thisval).val()
                        thisAktidUse = thisvaltrim
                        //thisJobnavn = $("#hiddn_job_" + thisJobid).val()
                        thisAktnavn = $("#dv_akt_" + thisval + " option:selected").text()

                        $("#FM_akt_" + thisval).val(thisAktnavn)
                        $("#FM_aktid_" + thisval).val(thisAktidUse)
                        //$(".dv_akt").hide();
                        $(".chbox_akt").hide();

                    });

                } else {



                }




            });

        }



    }




    function sogjobogkunde(thisval) {


        mobil_week_reg_job_dd = $("#mobil_week_reg_job_dd").val()
        mobil_week_reg_akt_dd = $("#mobil_week_reg_akt_dd").val()


        jq_newfilterval = $("#FM_job_" + thisval).val()
        jq_medid = $("#FM_medid").val()


        jq_jobidc = $("#jq_jobidc").val()

        if (mobil_week_reg_job_dd != 1) {
            $("#dv_job_" + thisval).css("display", "");
            $("#dv_job_" + thisval).css("visibility", "visible");
            $("#dv_job_" + thisval).show(100);
        }

        varTjDatoUS_man = $("#varTjDatoUS_man").val()
        lto = $("#lto").val()




        //$(".chbox_akt").hide();

        //alert("thisval: " + thisval + " jq_newfilterval: " + jq_newfilterval + " jq_medid: " + jq_medid)

        if (jq_newfilterval.length > 0 || mobil_week_reg_job_dd == "1") {

            //alert("H4")

            $.post("?lto=" + lto + "&jq_newfilterval=" + jq_newfilterval + "&jq_medid=" + jq_medid + "&varTjDatoUS_man=" + varTjDatoUS_man + "&jq_jobidc=" + jq_jobidc, { control: "FN_sogjobogkunde", AjaxUpdateField: "true" }, function (data) {
                //alert("cc")


                $("#dv_job_" + thisval).html(data);



                if (mobil_week_reg_job_dd != "1") {

                    $(".chbox_job").bind('click', function () {



                        var thisvaltrim = $("#dv_job_" + thisval).val()
                        //alert("her: " + thisvaltrim)


                        thisJobid = thisvaltrim
                        //thisJobnavn = $("#hiddn_job_" + thisJobid).val()
                        thisJobnavn = $("#dv_job_" + thisval + " option:selected").text()

                        $("#FM_job_" + thisval).val(thisJobnavn)
                        $("#FM_jobid_" + thisval).val(thisJobid)

                        //$(".dv_job").hide();
                        $(".chbox_job").hide();

                        if (mobil_week_reg_akt_dd == "1") {
                            sogakt(0);
                        }

                    });




                    $(".chbox_job").bind('keyup', function () {

                        if (window.event.keyCode == "13") {

                            var thisvaltrim = $("#dv_job_" + thisval).val()

                            thisJobid = thisvaltrim
                            //thisJobnavn = $("#hiddn_job_" + thisJobid).val()
                            thisJobnavn = $("#dv_job_" + thisval + " option:selected").text()



                            $("#FM_job_" + thisval).val(thisJobnavn)
                            $("#FM_jobid_" + thisval).val(thisJobid)

                            //$(".dv_job").hide();
                            $(".chbox_job").hide();

                        }

                        if (mobil_week_reg_akt_dd == "1") {
                            sogakt(0);
                        }


                    });



                } // mobil_week_reg_job_dd


                //alert("jq_jobidc" + jq_jobidc)
                if (jq_jobidc != "-1") {
                    $("#dv_job_0").val(jq_jobidc)
                    sogakt(0);
                    $("#dv_akt_0").prop("disabled", false);
                }

            });

        }
    }





    dagsval1 = $("#sumtimer_dag_1").val();
    dagsval2 = $("#sumtimer_dag_2").val();
    dagsval3 = $("#sumtimer_dag_3").val();
    dagsval4 = $("#sumtimer_dag_4").val();
    dagsval5 = $("#sumtimer_dag_5").val();
    dagsval6 = $("#sumtimer_dag_6").val();
    dagsval7 = $("#sumtimer_dag_7").val();

    stempelurOn = $("#stempelurOn").val();

    if (stempelurOn = 1) {
        dagsvallon1 = $("#sumlontimer_dag_1").val();
        dagsvallon2 = $("#sumlontimer_dag_2").val();
        dagsvallon3 = $("#sumlontimer_dag_3").val();
        dagsvallon4 = $("#sumlontimer_dag_4").val();
        dagsvallon5 = $("#sumlontimer_dag_5").val();
        dagsvallon6 = $("#sumlontimer_dag_6").val();
        dagsvallon7 = $("#sumlontimer_dag_7").val();

        $("#sp_sumlontimer_dag_1").html(dagsvallon1);
        $("#sp_sumlontimer_dag_2").html(dagsvallon2);
        $("#sp_sumlontimer_dag_3").html(dagsvallon3);
        $("#sp_sumlontimer_dag_4").html(dagsvallon4);
        $("#sp_sumlontimer_dag_5").html(dagsvallon5);
        $("#sp_sumlontimer_dag_6").html(dagsvallon6);
        $("#sp_sumlontimer_dag_7").html(dagsvallon7);


        //alert(dagsvallon4 * 1 + ">" + dagsval4 * 1)

        //if (dagsvallon1*1 > dagsval1*1) {
        // $("#sp_sumlontimer_dag_4").css('color', 'yellowgreen')
        //}
    }


    //alert(dagsval1)

    //if (dagsval1 != "0,00") {
    $("#sp_sumtimer_dag_1").html(dagsval1);
    //}

    //if (dagsval2 != "0,00") {
    //alert(dagsval2)
    $("#sp_sumtimer_dag_2").html(dagsval2);
    //}

    //if (dagsval3 != "0,00") {
    $("#sp_sumtimer_dag_3").html(dagsval3);
    //}

    //if (dagsval4 != "0,00") {
    $("#sp_sumtimer_dag_4").html(dagsval4);
    //}

    //if (dagsval5 != "0,00") {
    $("#sp_sumtimer_dag_5").html(dagsval5);
    //}

    //if (dagsval6 != "0,00") {
    $("#sp_sumtimer_dag_6").html(dagsval6);
    //}

    //if (dagsval7 != "0,00") {
    $("#sp_sumtimer_dag_7").html(dagsval7);
    //}




    $("#jq_dato").change(function () {

        var d = $("#jq_dato").val();
        var month = d.substring(3, 5);

        var date = d.substring(0, 2);
        var year = d.substring(6, 10);

        $("#regskabsaarUseAar").val(year);

        $("#regskabsaarUseMd").val(month)

    });





    /// Tjekker timer IKKE er over maks forecast.
    $("#FM_timer").keyup(function () {

        if (!validZip($("#FM_timer").val())) {
            alert("Der er angivet et ugyldigt tegn.\nYou have typed an invalid character")
            $("#FM_timer").val('');
            $("#FM_timer").focus()
            $("#FM_timer").select()
            return false
        }


        tjektimerForecast();



    });


    function tjektimerForecast() {


        var aktid = $("#dv_akt_0").val()
        var jobid = $("#dv_job_0").val()

        akttype = 1 //Skal have en værdi 

        //ER ALERT VED MAKS BUDGET SLÅET TIL
        akt_maksforecast_treg = $("#akt_maksforecast_treg").val()

        //alert("Aktid:" + aktid +" jobid: " + jobid )

        // MAKS forecast på aktivitet
        if (akt_maksforecast_treg == 1 && akttype == 1) {


            //Timertastet på linjen - minus timerOpr
            timer_tastet_this = 0;



            timer_tastet_this = $("#FM_timer").val();
            //alert(timer_tastet_this)

            if (timer_tastet_this == "NaN" || (timer_tastet_this == "-Infinity") || timer_tastet_this == "") {
                timer_tastet_this = 0
            } else {
                timer_tastet_this = timer_tastet_this.replace(",", ".")
                //timer_tastet = timer_tastet / 1 + timer_tastet_this / 1
            }




            timer_tastet = timer_tastet_this


            //alert(timer_tastet)


            //indenfor budgetår
            ibudgetaar = $("#aktBudgettjkOn_afgr").val()
            ibudgetaarMd = $("#regskabsaarStMd").val()
            ibudgetUseMd = $("#regskabsaarUseMd").val()
            ibudgetUseAar = $("#regskabsaarUseAar").val()
            treg_usemrn = $("#treg_usemrn").val()


            $.post("?aktid=" + aktid + "&timer_tastet=" + timer_tastet + "&treg_usemrn=" + treg_usemrn + "&ibudgetaar=" + ibudgetaar + "&ibudgetaarMd=" + ibudgetaarMd + "&ibudgetUseMd=" + ibudgetUseMd + "&ibudgetUseAar=" + ibudgetUseAar, { control: "FN_tjktimer_forecast", AjaxUpdateField: "true", cust: thisval }, function (data) {



                $("#aktnotificer_fc").val(data);

                //alert(data)

                var fcoverskreddet = 0;
                fcoverskreddet = $("#aktnotificer_fc").val().replace(",", ".")
                fcoverskreddet = fcoverskreddet * 1


                //fcoverskreddet = -2
                if (fcoverskreddet < 0) {

                    timer_opr = $("#FM_timer").val('0');
                    alert("Forecast er opbrugt!\nDit timeforecast er opbrugt på denne aktivitet.")


                }

            });

        } //akt_maksforecast_treg

    } // END FUNCTION



    function validZip(inZip) {
        if (inZip == "") {
            return true
        }
        if (isNum_treg(inZip)) {
            return true
        }
        return false
    }



    function isNum_treg(passedVal) {
        invalidChars = " /:;<>abcdefghijklmnopqrstuvwxyzæøå"

        //alert("her")

        if (passedVal == "") {
            return false
        }

        for (i = 0; i < invalidChars.length; i++) {
            badChar = invalidChars.charAt(i)
            if (passedVal.indexOf(badChar, 0) != -1) {
                return false
            }
        }

        for (i = 0; i < passedVal.length; i++) {
            if (passedVal.charAt(i) == "." || passedVal.charAt(i) == "-") {
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



    $("#bt_indlaspaajob").click(function () {
        $("#indlaspaajob").val('1')

        $("#bt_indlaspaajob").hide('fast')

        //$("#bt_indlaspaajob").css('display', 'none')
        //$("#bt_indlaspaajob").css('visibility', 'hidden')

        //alert("HER: " + $("#FM_akt_0").val())
        
        //FM_aktivitetid

        $("#monitorform").submit();
    });


    $("#sb_monitor_stempleud").click(function () {

        $("#sb_monitor_stempleud").hide('fast')

    });


}); // Doc ready


$(window).load(function () {
    // run code
    $("#loadbar").hide(1000);





});