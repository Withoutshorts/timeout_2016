





$(document).ready(function() {


    $('.date').datepicker({

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

 








    $("#FM_visallemedarb").click(function () {

       
        var sesid = $("#FM_sesMid").val()
        $("#usemrn").val(sesid)

        $("#filterkri").submit();


    });

    $("#usemrn").change(function () {
       
        $("#filterkri").submit();
    });

    

   

    $(".FM_job").keyup(function () {

       

        $(".dv_akt").hide();

        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(7, thisvallngt)
        thisval = thisvaltrim

        //alert(thisval)
        
        sogjobogkunde(thisval);

    });

    //$(".FM_job").blur()(function () {

    
    //    $(".dv_job").hide();
    //});

    $(".FM_akt").keyup(function () {

    

        $(".dv_job").hide();

       var thisid = this.id
       var thisvallngt = thisid.length
       var thisvaltrim = thisid.slice(7, thisvallngt)
       thisval = thisvaltrim

        sogakt(thisval);

    });


    $("#FM_timer").focus(function () {
        $(".chbox_job").hide();
        $(".chbox_akt").hide();
    });

    $("#FM_kom").focus(function () {
        $(".chbox_job").hide();
        $(".chbox_akt").hide();
    });


    function sogakt(thisval) {

        //alert(thisval)

        jq_newfilterval = $("#FM_akt_" + thisval).val()
        jq_jobid = $("#FM_jobid_" + thisval).val()
        jq_medid = $("#FM_medid").val()
        jq_aktid = $("#FM_akt_" + thisval).val()
        jq_pa = $("#FM_pa").val()

        varTjDatoUS_man = $("#varTjDatoUS_man").val()

        $(".chbox_job").hide();

        $("#dv_akt_" + thisval).css("display", "");
        $("#dv_akt_" + thisval).css("visibility", "visible");
        $("#dv_akt_" + thisval).show(100);

        //alert(jq_jobid + ":: medid: " + jq_medid + "#" + jq_newfilterval.length + " jq_aktid: " + jq_aktid + "## jq_pa" + jq_pa)

        if (jq_newfilterval.length > 0) {

            $.post("?jq_newfilterval=" + jq_newfilterval + "&jq_jobid=" + jq_jobid + "&jq_medid=" + jq_medid + "&jq_aktid=" + jq_aktid + "&jq_pa=" + jq_pa + "&varTjDatoUS_man=" + varTjDatoUS_man, { control: "FN_sogakt", AjaxUpdateField: "true" }, function (data) {
            //alert("cc")
            $("#dv_akt_" + thisval).html(data);

           



          

            $(".chbox_akt").bind('keyup', function () {

              
                if (window.event.keyCode == "13") {

                    var thisvaltrim = $("#dv_akt_" + thisval).val()
                    thisAktidUse = thisvaltrim
                    //thisJobnavn = $("#hiddn_job_" + thisJobid).val()
                    thisAktnavn = $("#dv_akt_" + thisval +" option:selected").text()



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
                thisAktnavn = $("#dv_akt_"+ thisval + " option:selected").text()

                

                $("#FM_akt_" + thisval).val(thisAktnavn)
                $("#FM_aktid_" + thisval).val(thisAktidUse)
                //$(".dv_akt").hide();
                $(".chbox_akt").hide();

            });


     

        });

        }

            

    }



  
    function sogjobogkunde(thisval) {

      

        jq_newfilterval = $("#FM_job_" + thisval).val()
        jq_medid = $("#FM_medid").val()


        $("#dv_job_" + thisval).css("display", "");
        $("#dv_job_" + thisval).css("visibility", "visible");
        $("#dv_job_" + thisval).show(100);

        varTjDatoUS_man = $("#varTjDatoUS_man").val()
       
        $(".chbox_akt").hide();

        //alert("thisval: " + thisval + " jq_newfilterval: " + jq_newfilterval + " jq_medid: " + jq_medid)

        if (jq_newfilterval.length > 0) {

            $.post("?jq_newfilterval=" + jq_newfilterval + "&jq_medid=" + jq_medid +"&varTjDatoUS_man=" + varTjDatoUS_man, { control: "FN_sogjobogkunde", AjaxUpdateField: "true" }, function (data) {
            //alert("cc")
            $("#dv_job_" + thisval).html(data);

           
            

          


                
            $(".chbox_job").bind('click', function () {

             
                    
                var thisvaltrim = $("#dv_job_" + thisval).val()
                    //alert("her: " + thisvaltrim)
                   
                   
                   

                    thisJobid = thisvaltrim
                    //thisJobnavn = $("#hiddn_job_" + thisJobid).val()
                    thisJobnavn = $("#dv_job_"+ thisval +" option:selected").text()
                    
                   

                    $("#FM_job_" + thisval).val(thisJobnavn)
                    $("#FM_jobid_" + thisval).val(thisJobid)

                    //$(".dv_job").hide();
                    $(".chbox_job").hide();

            });
                    



                $(".chbox_job").bind('keyup', function () {

                    if (window.event.keyCode == "13") {
                    
                        var thisvaltrim = $("#dv_job_" + thisval).val()
                
                    thisJobid = thisvaltrim
                    //thisJobnavn = $("#hiddn_job_" + thisJobid).val()
                    thisJobnavn = $("#dv_job_"+ thisval +" option:selected").text()
                    
                   

                    $("#FM_job_" + thisval).val(thisJobnavn)
                    $("#FM_jobid_" + thisval).val(thisJobid)

                    //$(".dv_job").hide();
                    $(".chbox_job").hide();

                    }
                    

                });


        
                
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

    if (dagsval1 != "0,00") {
        $("#sp_sumtimer_dag_1").html(dagsval1);
    }

    if (dagsval2 != "0,00") {
        //alert(dagsval2)
        $("#sp_sumtimer_dag_2").html(dagsval2);
    }

    if (dagsval3 != "0,00") {
        $("#sp_sumtimer_dag_3").html(dagsval3);
    }

    if (dagsval4 != "0,00") {
        $("#sp_sumtimer_dag_4").html(dagsval4);
    }

    if (dagsval5 != "0,00") {
        $("#sp_sumtimer_dag_5").html(dagsval5);
    }

    if (dagsval6 != "0,00") {
        $("#sp_sumtimer_dag_6").html(dagsval6);
    }

    if (dagsval7 != "0,00") {
    $("#sp_sumtimer_dag_7").html(dagsval7);
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


 }); // Doc ready


$(window).load(function() {
    // run code
    $("#loadbar").hide(1000);
});