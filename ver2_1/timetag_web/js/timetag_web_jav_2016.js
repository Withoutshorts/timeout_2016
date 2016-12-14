



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






$(document).ready(function() {


   

    mobil_week_reg_job_dd = $("#mobil_week_reg_job_dd").val()

    //alert("HER:" + mobil_week_reg_job_dd)
    if (mobil_week_reg_job_dd == "1") {
        sogjobogkunde();
    }


    //Kun ved job = DD
    $("#dv_job").change(function () {

        mobil_week_reg_job_dd = $("#mobil_week_reg_job_dd").val()
        mobil_week_reg_akt_dd = $("#mobil_week_reg_akt_dd").val()
        //alert("søger: j:" + mobil_week_reg_job_dd + " a:" + mobil_week_reg_akt_dd)

        if (mobil_week_reg_job_dd == "1") {

            if (mobil_week_reg_akt_dd != "1") {

                $("#FM_akt").val('');
                $("#dv_akt").hide();

            } else {

                $("#dv_akt").attr('disabled', false);
            }


            sogakt();

            if (mobil_week_reg_akt_dd != "1") {

                aktSogVal = $("#FM_akt").val()
                //alert(aktSogVal.length)

                if (aktSogVal.length == 0) {
                    $("#FM_aktid").val('0');
                    $("#dv_akt").hide();
                }
            }

       
        }

    });


  

    //approved
    $("#sbm_timer").click(function () {

        // Skal ikke tilføje logud tid, da det skal kunne indtastes manuelt //
        //if ($("#showstop").val() == "1") {
        //    ststop();
        //}

        //$("#dvindlaes_msg").attr('class', 'approved');
        $("#dv_mat_container").hide(10)

        $("#dvindlaes_msg").css("display", "");
        $("#dvindlaes_msg").css("visibility", "visible");
        $("#dvindlaes_msg").show(100);

    });


    
    $("#FM_sltid").keyup(function () {

        ststop(1);

    });


    
    
    $("#bt_stst").click(function () {

        ststop(0);
       
    });


    function ststop(io) {

        //var sttid = new Date();
        //var sltid = new Date();

        sttid = $("#FM_sttid").val()

        var idtrim = sttid.slice(1, 2)
        if (idtrim == ":") {
            $("#FM_sttid").val("0" + sttid)
            sttid = $("#FM_sttid").val()
        }


        sltid = $("#FM_sltid").val()

        /*var idtrim = sltid.slice(1, 2)
        if (idtrim == ":") {
            $("#FM_sltid").val("0" + sltid)
            sltid = $("#FM_sltid").val()
        }*/


        if (sttid.length == 0 || sttid == "00:00") {

           
           

            var dt = new Date();

            var minutter = dt.getMinutes()
            if (minutter < 10) {
                minutter = 0 + "" + minutter
            }

            var sttimeNow = dt.getHours() + ":" + minutter // + ":" + dt.getSeconds();

            $("#FM_sttid").val(sttimeNow)
            $("#FM_sttid").css("color", "#000000")

            $("#FM_sltid").val(sttimeNow)


            sttid = $("#FM_sttid").val()


        } else {

            if (io == 0) {

                var dt = new Date();

                var minutter = dt.getMinutes()
                if (minutter < 10) {
                    minutter = 0 + "" + minutter
                }

                var sltimeNow = dt.getHours() + ":" + minutter // + ":" + dt.getSeconds();

                $("#FM_sltid").val(sltimeNow)
                $("#FM_sltid").css("color", "#000000")
                sltid = $("#FM_sltid").val()

            }


        }

        //var dt = new Date();
        //dt.getDay() + ":" + minutter
        //alert(sttid)

        //dt1 = new Date(2001, 1, 1); //UTC time
        //alert(dt1.getTime()); //gives back timestamp in ms

        dt1 = new Date("1/1/2001 "+ sttid + ":00")
        dt2 = new Date("1/1/2001 "+ sltid + ":00")


        //alert(dt1 +"//"+ dt2)

        //var beginTime = stringToDate(dt1);
        //var endTime = stringToDate(dt1);

        //alert(beginTime)
        if (sttid < sltid) {
            var sp = (((dt2 - dt1) / 60000));
        } else {
            var sp = 0
        }


       
        sp = sp / 60
      
        //alert("timeuse：" + sp.hour + " hour " + sp.minute + " minute " + sp.second + " second")

        if (sp > 0) {
            $("#FM_timer").css("color", "#000000")
        }

        sp = Math.round((parseFloat(sp)) * 100) / 100
        var spTxt = String(sp).replace(".",",")
        $("#FM_timerlbl").html(spTxt)

        if (spTxt != "0") {
            $("#FM_timer").val(spTxt)
       }
      


    }


    $("#sbmmat").click(function () {

        nytMatantal = $("#FM_matantal").val()

       
      

        if (isNum_treg(nytMatantal)) {
            //return true
        } else {

            nytMatantal = 0
            $("#FM_matantal").val(0)
            return false 
        }
       
       
        //alert("HER: " + nytMatantal)
        

        if (nytMatantal.length > 0 && nytMatantal > 0 && $("#FM_matnavn").val() != 'Tilføj Materiale' && $("#FM_matnavn").val() != 'flere?' && $("#FM_matnavn").val() != '') {

            nytMatnavn = $("#FM_matantal").val() + " " + $("#FM_matnavn").val()

            prevIndlast = $("#dv_mat_sbm").html();

            if (prevIndlast.length > 0) {
                $("#dv_mat_sbm").html(nytMatnavn + ", " + prevIndlast);
            } else {
                $("#dv_mat_sbm").html(nytMatnavn);
            }


            prevMatids = $("#FM_matids").val();
            nytMatid = $("#FM_matid").val();

            if (prevMatids.length > 0) {
                $("#FM_matids").val(prevMatids + ", " + nytMatid);
            } else {
                $("#FM_matids").val(nytMatid);
            }


            prevMatAntals = $("#FM_matantals").val();


            //alert(nytMatantal)

            if (prevMatAntals.length > 0) {
                $("#FM_matantals").val(prevMatAntals + ", " + nytMatantal);
            } else {
                $("#FM_matantals").val(nytMatantal);
            }

            $("#FM_matantal").val('')
            $("#FM_matnavn").val('flere?')

        }
    
    });

    

    /* 
    $("#a_mat").click(function () {


        if ($("#dv_mat_container").css('display') == "none") {
            $("#dv_mat_container").css("display", "");
            $("#dv_mat_container").css("visibility", "visible");
            $("#dv_mat_container").show(100);
        } else {
            $("#dv_mat_container").css("display", "none");
            $("#dv_mat_container").css("visibility", "hidden");
            $("#dv_mat_container").hidden(100);
        }

    });
    */

   
    $("#FM_matnavn").focus(function () {

        if ($("#FM_matnavn").val() == 'Tilføj Materiale' || $("#FM_matnavn").val() == 'flere?') {
            $("#FM_matnavn").val('')
        }
    });
   

    $("#FM_job").focus(function () {

        //if ($("#FM_job").val() == 'Kunde/job') {
            $("#FM_job").val('')
        //}
    });


    $("#FM_akt").focus(function () {

        //if ($("#FM_akt").val() == 'Aktivitet') {
            $("#FM_akt").val('')
        //}
    });


    //$("#FM_timer").focus(function () {
    //
    //    if ($("#FM_timer").val() == 'Antal timer') {
    //        $("#FM_timer").val('')
    //    }
    //});


    $("#FM_kom").focus(function () {

        if ($("#FM_kom").val() == 'Kommentar') {
            $("#FM_kom").val('')
        }
    });


    $("#FM_job").keyup(function () {


        mobil_week_reg_job_dd = $("#mobil_week_reg_job_dd").val()
        mobil_week_reg_akt_dd = $("#mobil_week_reg_akt_dd").val()

        if (mobil_week_reg_akt_dd != "1") {
            $("#dv_akt").hide();
        } 

        $("#dv_mat").hide();
        

       
        sogjobogkunde();



    });


    $("#FM_matnavn").keyup(function () {

        //alert("gdgf")

        $("#dv_job").hide();
        $("#dv_akt").hide();
        
        sogMat();



    });
    

    $("#FM_akt").keyup(function () {

        $("#dv_job").hide();
        $("#dv_mat").hide();

        sogakt();

    });



    function sogMat() {

        

        jq_newfilterval = $("#FM_matnavn").val()
       

       
        if (jq_newfilterval.length > 0) {



            //alert("cc")

            $.post("?jq_newfilterval=" + jq_newfilterval, { control: "FN_sogmat", AjaxUpdateField: "true" }, function (data) {
                //alert("cc")


                $("#dv_mat").html(data);

                //alert("her")
                //$("#dv_akt").attr('class', 'dv-open');

                if ($("#dv_mat").css('display') == "none") {
                    $("#dv_mat").css("display", "");
                    $("#dv_mat").css("visibility", "visible");
                    $("#dv_mat").show(100);
                }



                    /*

                    var thismatid = this.id
                    var thisvallngt = thismatid.length
                    var thisvaltrim = thismatid.slice(10, thisvallngt)


                    thismatid = thisvaltrim
                    thismatnavn = $("#hiddn_mat_" + thismatid).val()

                   
                    $("#FM_matnavn").val(thismatnavn)
                    $("#FM_matid").val(thismatid)
                    $("#dv_mat").hide();
                    */
                    



                    $("#dv_mat").bind('click', function () {


                        var thisvaltrim = $("#dv_mat").val()
                        thisMatidUse = thisvaltrim
                        thisMatnavn = $("#dv_mat option:selected").text()

                        $("#FM_matnavn").val(thisMatnavn)
                        $("#FM_matid").val(thisMatidUse)
                        
                        $("#dv_mat").hide();

                    });



               


            });

        } 

    }



    function sogakt() {

       


        mobil_week_reg_akt_dd = $("#mobil_week_reg_akt_dd").val()
        mobil_week_reg_job_dd = $("#mobil_week_reg_job_dd").val()

        //alert(mobil_week_reg_job_dd + " # " + mobil_week_reg_akt_dd)

        if (mobil_week_reg_akt_dd != "1") {

            jq_newfilterval = $("#FM_akt").val()

            if (mobil_week_reg_job_dd != "1") {
                jq_jobid = $("#FM_jobid").val()
            } else {
                jq_jobid = $("#dv_job").val()
            }


        } else {
            jq_newfilterval = "-1"

            if (mobil_week_reg_job_dd != "1") {
                jq_jobid = $("#FM_jobid").val()
            } else {
                jq_jobid = $("#dv_job").val()
            }
        }


        jq_lto = $("#jq_lto").val()


        if (jq_lto == 'hestia' || jq_lto == 'intranet - local') {
            //alert("søog ak" + jq_lto)
            sogjobbesk(jq_jobid);
        }



        //alert("jobid: " + jq_jobid)

        //jq_newfilterval = $("#FM_akt").val()
        //jq_jobid = $("#FM_jobid").val()
        jq_medid = $("#FM_medid").val()
        jq_aktid = $("#FM_akt").val()
        jq_pa = $("#FM_pa").val()
        
       

        varTjDatoUS_man = $("#varTjDatoUS_man").val()

        if (jq_newfilterval.length > 0 || mobil_week_reg_akt_dd == "1") {
            
            $.post("?jq_newfilterval=" + jq_newfilterval + "&jq_jobid=" + jq_jobid + "&jq_medid=" + jq_medid + "&jq_aktid=" + jq_aktid + "&jq_pa=" + jq_pa + "&varTjDatoUS_man=" + varTjDatoUS_man, { control: "FN_sogakt", AjaxUpdateField: "true" }, function (data) {
               


                $("#dv_akt").html(data);


               


                if (mobil_week_reg_akt_dd != "1") {

                    if ($("#dv_akt").css('display') == "none") {
                        $("#dv_akt").css("display", "");
                        $("#dv_akt").css("visibility", "visible");
                        $("#dv_akt").show(100);
                    }


                    $("#dv_akt").bind('keyup', function () {


                        if (window.event.keyCode == "13") {

                            var thisvaltrim = $("#dv_akt").val()
                            thisAktidUse = thisvaltrim
                            //thisJobnavn = $("#hiddn_job_" + thisJobid).val()
                            thisAktnavn = $("#dv_akt option:selected").text()

                            $("#FM_akt").val(thisAktnavn)
                            $("#FM_aktid").val(thisAktidUse)
                            //$(".dv_akt").hide();
                            $("#dv_akt").hide();

                        }

                    });


                    $("#dv_akt").bind('click', function () {


                        var thisvaltrim = $("#dv_akt").val()
                        thisAktidUse = thisvaltrim
                        //thisJobnavn = $("#hiddn_job_" + thisJobid).val()
                        thisAktnavn = $("#dv_akt option:selected").text()

                        $("#FM_akt").val(thisAktnavn)
                        $("#FM_aktid").val(thisAktidUse)
                        //$(".dv_akt").hide();
                        $("#dv_akt").hide();

                    });

                }



            });

        } 

    }



    function sogjobbesk(jq_jobid) {

        $("#dv_jobbesk").html('');
        
       
        //alert("jobbesk" + jq_jobid)

        $.post("?id=" + jq_jobid, { control: "FN_jobbesk", AjaxUpdateField: "true" }, function (data) {

            $("#dv_jobbesk").html(data);

            });

        }      


  
    function sogjobogkunde() {

        
        mobil_week_reg_job_dd = $("#mobil_week_reg_job_dd").val()
        mobil_week_reg_akt_dd = $("#mobil_week_reg_akt_dd").val()
        
        varTjDatoUS_man = $("#varTjDatoUS_man").val()
        
        jq_newfilterval = $("#FM_job").val()
        jq_medid = $("#FM_medid_k").val()
       
        jq_lto = $("#jq_lto").val()

        if (mobil_week_reg_job_dd != 1) {

            if ($("#dv_job").css('display') == "none") {

                $("#dv_job").css("display", "");
                $("#dv_job").css("visibility", "visible");
                $("#dv_job").show(100);

            }
        }

        //alert(jq_newfilterval.length)
        //alert("cc" + jq_newfilterval)
        if (jq_newfilterval.length > 0 || mobil_week_reg_job_dd == "1") {

            $.post("?jq_newfilterval=" + jq_newfilterval + "&jq_medid=" + jq_medid + "&varTjDatoUS_man=" + varTjDatoUS_man + "&jq_lto=" + jq_lto, { control: "FN_sogjobogkunde", AjaxUpdateField: "true" }, function (data) {

                

                $("#dv_job").html(data);

              
                if (mobil_week_reg_job_dd != "1") {

                    $(".chbox_job").bind('click', function () {



                        var thisvaltrim = $("#dv_job").val()
                        //alert("her: " + thisvaltrim)


                        thisJobid = thisvaltrim
                        //thisJobnavn = $("#hiddn_job_" + thisJobid).val()
                        thisJobnavn = $("#dv_job option:selected").text()

                        $("#FM_job").val(thisJobnavn)
                        $("#FM_jobid").val(thisJobid)
                       
                        $("#dv_jobbesk").html("timereg_web_jobbesk.asp?id=" + thisJobid)
                        
                        //$(".dv_job").hide();
                        $(".chbox_job").hide();

                        if (mobil_week_reg_akt_dd == "1") {
                            sogakt();

                        }

                    });




                    $(".chbox_job").bind('keyup', function () {

                        if (window.event.keyCode == "13") {

                            var thisvaltrim = $("#dv_job").val()

                            thisJobid = thisvaltrim
                            //thisJobnavn = $("#hiddn_job_" + thisJobid).val()
                            thisJobnavn = $("#dv_job option:selected").text()


                            $("#FM_job").val(thisJobnavn)
                            $("#FM_jobid").val(thisJobid)

                            //$(".dv_job").hide();
                            $(".chbox_job").hide();

                        }

                        if (mobil_week_reg_akt_dd == "1") {
                            sogakt();
                        }


                    });


                } // mobil_week_reg_job_dd


                ststop(1);

             

            });

        } 
    }



    


    setTimeout(function () {
        // Do something after 5 seconds
        $("#timer_indlast").hide(1000);
    }, 2500);
    


    if ($("#tomobjid").val() != 0) {

        

        
        thisJobid = $("#tomobjid").val()
        $("#FM_jobid").val(thisJobid)

        $.post("?jq_jobid=" + thisJobid, { control: "FN_tomobjid", AjaxUpdateField: "true" }, function (data) {




            $("#FM_job").val(data)


        });

           

    }




    $("#jq_dato").change(function () {

        var d = $("#jq_dato").val();
        var month = d.substring(3, 5);

        var date = d.substring(0, 2);
        var year = d.substring(6, 10);

        $("#regskabsaarUseAar").val(year);
        $("#regskabsaarUseMd").val(month)

    });



    /// Tjekker timer IKKE er over maks forecast.
    $("#FM_timer, #FM_sttid, #FM_sltid").keyup(function () {

       

        if (!validZip($("#FM_timer").val())) {
            alert("Der er angivet et ugyldigt tegn.\nYou have typed an invalid character")
            $("#FM_timer").val('');
            $("#FM_timer").focus()
            $("#FM_timer").select()
            return false
        }


        tjektimerForecast();



    });

   

    $("#bt_stst").click(function () {

        tjektimerForecast();
    });


   


    function tjektimerForecast() {

        var aktid = $("#dv_akt").val()
        var jobid = $("#dv_job").val()

        //alert("HER aktid: " + aktid + " jobid: " + jobid)

        akttype = 1 //Skal have en værdi // hardcoaded til fakturerbar

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


            //indenfor budgetår
            ibudgetaar = $("#aktBudgettjkOn_afgr").val()
            ibudgetaarMd = $("#regskabsaarStMd").val()
            ibudgetUseMd = $("#regskabsaarUseMd").val()
            ibudgetUseAar = $("#regskabsaarUseAar").val()
            treg_usemrn = $("#FM_medid").val()

            thisval = 0

            //alert("timer_tastet: " + timer_tastet + " treg_usemrn: " + treg_usemrn + " aktid: " + aktid + " ibudgetaar" + ibudgetaar + "ibudgetaarMd: " + ibudgetaarMd + "ibudgetUseMd: " + ibudgetUseMd + " ibudgetUseAar: " + ibudgetUseAar)

            $.post("?aktid=" + aktid + "&timer_tastet=" + timer_tastet + "&treg_usemrn=" + treg_usemrn + "&ibudgetaar=" + ibudgetaar + "&ibudgetaarMd=" + ibudgetaarMd + "&ibudgetUseMd=" + ibudgetUseMd + "&ibudgetUseAar=" + ibudgetUseAar, { control: "FN_tjktimer_forecast_tt", AjaxUpdateField: "true", cust: thisval }, function (data) {
            
                
                //alert("HEJK")
                $("#aktnotificer_fc").val(data);

                var fcoverskreddet = 0;
                fcoverskreddet = $("#aktnotificer_fc").val().replace(",", ".")
                fcoverskreddet = fcoverskreddet * 1

                //alert(fcoverskreddet)

                //fcoverskreddet = -2
                if (fcoverskreddet < 0) {

                    timer_opr = $("#FM_timer").val('');
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




});

$(window).load(function() {
    // run code
    
});