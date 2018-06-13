

//var dataString = navigator.userAgent
//var index = dataString.indexOf("Chrome");
////alert(navigator.userAgent)
//if (index != -1) {
//    alert("Der er iøjeblikket et problem med faktura oprettelsen i denne browser.\n\nSkift til Internet Explorer.\n\nVi arbejder på at løse problemet.")
    
//}




// run code
//GetMatTilFak();

    


$(window).load(function () {


    
    

   

    // run code
    //GetMatTilFak();

    if ($("#tjekantalakt_all").val() < 100) {
        $("#sideload").hide("fast");
    }

    

    $("#menu").css("display", "");
    $("#menu").css("visibility", "visible");


    $("#d_sideombr1").css('display', 'none');
    $("#d_sideombr1").css('visibility', 'hidden');
    $("#d_sideombr2").css('display', 'none');
    $("#d_sideombr2").css('visibility', 'hidden');
    $("#d_sideombr3").css('display', 'none');
    $("#d_sideombr3").css('visibility', 'hidden');
    $("#d_sideombr4").css('display', 'none');
    $("#d_sideombr4").css('visibility', 'hidden');
    $("#d_sideombr5").css('display', 'none');
    $("#d_sideombr5").css('visibility', 'hidden');
    $("#d_sideombr6").css('display', 'none');
    $("#d_sideombr6").css('visibility', 'hidden');
    $("#d_sideombr7").css('display', 'none');
    $("#d_sideombr7").css('visibility', 'hidden');




    /// Slå aktiviteter til / fra i oversigten ///
    //alert("her")
    //$("#ch_akt_" + x).attr('checked', true);


    antal_x = (document.getElementById("antal_x").value)
    //alert(antal_x)

    //antal_x = 0
    for (x = 0; x < antal_x; x++) {



        for (i = 0; i < 15; i++) { // maks 15 linier sum aktiviteter pr. aktivitet


            //alert($("#FM_show_akt_" + x + "_" + i))

            if ($("#FM_show_akt_" + x + "_" + i).is(':checked') == true) {
                $("#ch_akt_" + x).attr('checked', true);
                break;
            } else {
                $("#ch_akt_" + x).attr('checked', false);
            }

        }
    }



});







$(document).ready(function () {

   

    //ch_akt_x

    //setTimeout("cl_spfakdato()", 5000); //millisekunder (er sat til 1 time)

    
    $("#sp_editfak").mouseover(function () {
        $(this).css('cursor', 'pointer');
    });


    $("#sp_editfak").click(function () {


        
        $("#visikkefak").css("display", "none");
        $("#visikkefak").css("visibility", "hidden");
        $("#visikkefak").hide(100);
       
        //alert("her")

        $("#fidiv_2").css("display", "");
        $("#fidiv_2").css("visibility", "visible");
        $("#fidiv_2").show(200);

        
        $("#knap_fidiv").css("display", "");
        $("#knap_fidiv").css("visibility", "visible");
        $("#knap_fidiv").show(200);

        $("#knap_joblogdiv").css("display", "");
        $("#knap_joblogdiv").css("visibility", "visible");
        $("#knap_joblogdiv").show(200);

        $("#knap_matlogdiv").css("display", "");
        $("#knap_matlogdiv").css("visibility", "visible");
        $("#knap_matlogdiv").show(200);

        $("#knap_faksogdiv").css("display", "");
        $("#knap_faksogdiv").css("visibility", "visible");
        $("#knap_faksogdiv").show(200);

        $("#knap_modtagdiv").css("display", "");
        $("#knap_modtagdiv").css("visibility", "visible");
        $("#knap_modtagdiv").show(200);

        $("#knap_jobbesk").css("display", "");
        $("#knap_jobbesk").css("visibility", "visible");
        $("#knap_jobbesk").show(200);

        $("#knap_aktdiv").css("display", "");
        $("#knap_aktdiv").css("visibility", "visible");
        $("#knap_aktdiv").show(200);

        $("#knap_matdiv").css("display", "");
        $("#knap_matdiv").css("visibility", "visible");
        $("#knap_matdiv").show(200);

        $("#knap_betdiv").css("display", "");
        $("#knap_betdiv").css("visibility", "visible");
        $("#knap_betdiv").show(200);

        $("#genindlaes").css("display", "");
        $("#genindlaes").css("visibility", "visible");
        $("#genindlaes").show(200);

        $("#fidiv").css("display", "");
        $("#fidiv").css("visibility", "visible");
        $("#fidiv").show(200);

        $("#kontodiv").css("display", "");
        $("#kontodiv").css("visibility", "visible");
        $("#kontodiv").show(200);
        
        

        //knap_fidiv;

        
        //if ($("#knap_joblogdiv").css('display') == "none") {

        

    });


    $("#showfakind, #luk_fakpre").click(function () {

     

        if ($("#fakind_d").css('display') == "none") {

            $("#fakind_d").css("display", "");
            $("#fakind_d").css("visibility", "visible");
            $("#fakind_d").show(2000);

        } else {

            $("#fakind_d").hide();

        }

        //alert("her" + $("#fakind_d").css('display'))

    });

    $("#luk_fakpre").mouseover(function () {
        $(this).css('cursor', 'pointer');
    });


    




    $("#FM_sog").focus(function () {
        $("#FM_sog").val('')
    });

    $("#FM_globalfaktor").keyup(function () {
        kommaTrail = $("#FM_globalfaktor").val().replace(".", ",")
        $("#FM_globalfaktor").val(kommaTrail)
        if ($("#FM_job").val() != 0) {
            berengFaktorBelobUdgor();
        }
    });

    
    //Globalfaktor beregning//
    $("#bt_beregn_globalfaktor").click(function () {
       
        var glbfaktor = $("#FM_globalfaktor").val().replace(",", ".")
        glbfaktor = Math.round(glbfaktor * 100) / 100
        
        //$(".FM_medarbtimer").val('20')
        
        for (xval = 0; xval < 20 ; xval++) {
            //alert(xval)
            // Medarbejder felter
            //alert(document.getElementById("aktId_n_" + xval + "").value)
            antal_nx = $("#antal_n_" + xval + "").val()

            //alert("her " + antal_nx)

            if (antal_nx >= 1) {

                
                for (e = 1; e < antal_nx; e++) {
                    fak = $("#FM_m_fak_" + e + "_" + xval).val().replace(",", ".")
                    newfak = Math.round((parseFloat(fak * glbfaktor)) * 100) / 100
                    $("#FM_m_fak_" + e + "_" + xval).val(newfak)

                    //alert("x:" + xval + " e:" + e)
                    if ($("#FM_show_akt_"+xval+"_"+e).is(':checked') == true) {
                    
                        enhedsprismedarb(xval, e)


                        newfak = String(newfak).replace(".", ",")
                        $("#FM_m_fak_" + e + "_" + xval).val(newfak)
                    }


                }
            }
        }

      
        $("#bt_beregn_globalfaktor").hide("fast");
   
        $("#bt_beregn_globalfaktor_fortryd").css("display", "");
        $("#bt_beregn_globalfaktor_fortryd").css("visibility", "visible");
        $("#bt_beregn_globalfaktor_fortryd").show(100);

        
         

    });



    //Globalfaktor Reciprokke værdi / Fortryd//
    $("#bt_beregn_globalfaktor_fortryd").click(function () {

        var glbfaktor = $("#FM_globalfaktor").val().replace(",", ".")
        glbfaktor = Math.round(glbfaktor * 100) / 100

        //$(".FM_medarbtimer").val('20')

        for (xval = 0; xval < 20 ; xval++) {
            //alert(xval)
            // Medarbejder felter
            //alert(document.getElementById("aktId_n_" + xval + "").value)
            antal_nx = $("#antal_n_" + xval + "").val()

            //alert("her " + antal_nx)

            if (antal_nx >= 1) {


                for (e = 1; e < antal_nx; e++) {
                    fak = $("#FM_m_fak_" + e + "_" + xval).val().replace(",", ".")
                    newfak = Math.round((parseFloat(fak / glbfaktor)) * 100) / 100
                    $("#FM_m_fak_" + e + "_" + xval).val(newfak)

                    //alert("x:" + xval + " e:" + e)
                    if ($("#FM_show_akt_" + xval + "_" + e).is(':checked') == true) {

                        enhedsprismedarb(xval, e)


                        newfak = String(newfak).replace(".", ",")
                        $("#FM_m_fak_" + e + "_" + xval).val(newfak)
                    }


                }
            }
        }


        $("#bt_beregn_globalfaktor_fortryd").hide("fast");

        $("#bt_beregn_globalfaktor").css("display", "");
        $("#bt_beregn_globalfaktor").css("visibility", "visible");
        $("#bt_beregn_globalfaktor").show(100);



    });




    //Start faktor beløb hvis den tilføjes
    //alert($("#FM_job").val())
    if ($("#FM_job").val() != 0) {
        berengFaktorBelobUdgor();
    }

    


    $("#FM_afsender").change(function () {
        //alert("her")
        thisSelVal = $("#FM_afsender").val()


        $.post("?kidsel=" + thisSelVal, { control: "FN_bankkonto", AjaxUpdateField: "true" }, function (data) {
            //alert(thisSelVal)
            $("#FM_afs_bankkonto").html(data);

        });



    });



    $("#showinternnote").click(function () {
        //alert("her" + $("#FM_job_internbesk").val())
        $("#internbesk_tweet").html($("#internbesk_hd").html())
        $("#showinternnote").html("<u>Intern Note</u>")
        $("#showjobtweet").html("Job Tweet")
        
    });

    $("#showjobtweet").click(function () {
        //alert("her" + $("#tweet_hd").html())
        $("#internbesk_tweet").html($("#tweet_hd").html())
        $("#showinternnote").html("Intern Note")
        $("#showjobtweet").html("<u>Job Tweet</u>")
    });



    // Til fra aktiviterter alle
    $("#akttilfra").click(function () {


        antal_x = (document.getElementById("antal_x").value)
        var aTimer = 0;
        var nulAktTjk = 0;
        // Hvis tjekket kun sum samle akt. bliver slået til. 

        for (i = 0; i < antal_x; i++) {
        nulAktTjk = 0;

            if ($("#akttilfra").is(':checked') == true) {

                // tjekker kun de sum akt. til der er timer på
                for (e = 0; e < 15; e++) { // maks 15 linier sum aktiviteter pr. aktivitet

                    if (e == 0) { //Sum akt
                        aTimerSum = $("#FM_timerthisSum_" + i + "_" + e + "").val()
                        aTimer = $("#FM_timerthis_" + i + "_" + e + "").val()

                      
                    } else {
                        aTimer = $("#FM_timerthis_" + i + "_" + e + "").val()
                        aTimerSum = 1 //altid til
                    }

                    //if (i == 0) {
                    //    alert(aTimer + " aTimerSum:" + aTimerSum + "nulAktTjk: " + nulAktTjk)
                    //}

                    if (aTimer > 0 && aTimerSum > 0 && nulAktTjk == 0) {
                        $("#FM_show_akt_" + i + "_" + e).attr('checked', true);
                        $("#ch_akt_" + i).attr('checked', true);

                        //Lukker for de alm. sum akt, hvis Sum samle akt er chk.
                        if (e == 0) { //Sum samle akt
                          
                            if (aTimer > 0 && aTimerSum > 0) {
                                nulAktTjk = 1
                            }

                        }

                    }
                }


            } else {
                $(".visakt_" + i).attr('checked', false);
                $(".ch_akt_x").attr('checked', false);
            }


        }




        opdatervalutAllelinier(1,0,1)

    });


    /// Til fra aktivitet enkel
    $(".ch_akt_x").click(function () {

        
        //alert(this.id)

        var thisid = this.id
        var idlngt = thisid.length
        var idtrim = thisid.slice(7, idlngt)

        checkAkt(idtrim);
        opdatervalutAllelinier(1,0,1);


    });



    function checkAkt(id) {

        //alert(id)

        //var thisid = this.id
        //var idlngt = thisid.length
        //var idtrim = thisid.slice(7, idlngt)
        idtrim = id

        var aTimer = 0;
        //alert(idtrim)
        //alert($(this).is(':checked'))

        if ($("#ch_akt_" + idtrim).is(':checked') == true) {

            // tjekker kun de sum akt. til der er timer på
            for (e = 0; e < 15; e++) { // maks 15 linier sum aktiviteter pr. aktivitet

                aTimer = $("#FM_timerthis_" + idtrim + "_" + e + "").val()
                if (aTimer > 0) {
                    $("#FM_show_akt_" + idtrim + "_" + e).attr('checked', true);

                }
            }


        } else {
            $(".visakt_" + idtrim).attr('checked', false);
        }


    }


 
    function berengFaktorBelobUdgor() {
        //alert("berenger")
        var glbfaktor = $("#FM_globalfaktor").val().replace(",", ".")
        glbfaktor = Math.round(glbfaktor * 100) / 100
        var aktbelob = $("#FM_timer_beloeb").val().replace(",", ".")
        aktbelob = Math.round(aktbelob * 100) / 100
        var aktbelobInclFaktor = Math.round((parseFloat((aktbelob * glbfaktor) - aktbelob)) * 100) / 100
        aktbelobInclFaktor = String(aktbelobInclFaktor).replace(".", ",")


        $("#gblfaktorbelob_udgor").html("<b>" + aktbelobInclFaktor + "</b>");
    }



    //////////////////////// INLINE FAKDATO VÆLGER ///////////////////////////////
    //Dato func jq //
    function IsValidDate(input) {
        var dayfield = input.split(".")[0];
        var monthfield = input.split(".")[1];
        var yearfield = input.split(".")[2];
        var returnval;
        var dayobj = new Date(yearfield, monthfield - 1, dayfield)
        if ((dayobj.getMonth() + 1 != monthfield) || (dayobj.getDate() != dayfield) || (dayobj.getFullYear() != yearfield)) {
            alert("Forkert datoformat, skriv venligst dd.mm.yyyy eller d.m.yyyy");
            returnval = false;
        }
        else {
            returnval = true;
        }
        return returnval;
    }


    
    //Make datepickers
    $.datepicker.setDefaults($.extend({ showOn: 'button', constrainInput: true, showOtherMonths: true, showWeeks: true, minDate: new Date(2002, 1, 1), firstDay: 1, changeFirstDay: false, buttonImage: '../ill/popupcal.gif', start: 6, buttonImageOnly: true, dateFormat: 'd.m.yy', changeMonth: true, changeYear: true }));

    $("#FM_fakdato").datepicker();
    $("#FM_fakdato").change(function () { if (IsValidDate($("#FM_fakdato").val()) == true) { var datestring = $(this).val(); var datesplit = datestring.split('.'); $("[name=FM_fakdato_dag]").val(datesplit[0]); $("[name=FM_fakdato_mrd]").val(datesplit[1]); $("[name=FM_fakdato_aar]").val(datesplit[2]); GetBetIntVal(); } else { alert("dato ikke gemt"); } });
    $("#FM_fakdato").css("display", "inline");

















    ///////////////////////////
    // Forfaldsdato beregner //

    //$.datepicker.setDefaults($.extend({ showOn: 'button', constrainInput: true, showOtherMonths: true, showWeeks: true, minDate: new Date(2002, 1, 1), firstDay: 1, changeFirstDay: false, buttonImage: '../ill/popupcal.gif', start: 6, buttonImageOnly: true, dateFormat: 'd.m.yy', changeMonth: true, changeYear: true }));
    //alert($("#hideffdato").val())
    $("#FM_forfaldsdato").datepicker()

    //if ($("#hideffdato").val() != "1") {
    //}
    $("#FM_forfaldsdato").css("display", "inline");
    //{ buttonImage: '../ill/popupcal.gif' }

    $("#FM_forfaldsdato").change(function () {
        //Sætter select til "IKKE ANGIVET"
        $("#FM_betbetint").val("-1");
    });

    //hide unused fields and show relevant
    $(".popupcalImg").hide();

    //FM_betbetint//#FM_istdato, #FM_istdato2"//
    $("#FM_betbetint, #FM_labeldato").change(function () { GetBetIntVal(); });
    $("[name=FM_brugfakdatolabel]").click(function () {
        GetBetIntVal();
    });





    function daydiff(first, second) {
        return (second.getTime() - first.getTime()) / (1000 * 60 * 60 * 24);
    }

    function daysInMonth(month, year) {
        var dd = new Date(year, month, 0);
        //alert(dd.getDate() + " m:" + month);
        return dd.getDate();

    }


    //$(".tr_fase").mouseover(function () {

    //    alert("her")

    //});
   

    $("#FM_hidefasesum").click(function () {

        showhideFaser();

    });

    showhideFaser();



    function showhideFaser() {

        //alert("faser:" + $("#FM_hidefasesum").is(':checked'))

        if ($("#FM_hidefasesum").is(':checked') == true) {
            $(".tr_fase").css("visibility", "hidden");
            $(".tr_fase").css("display", "none");
            $(".tr_fase").hide(500);
        } else {

            $(".tr_fase").css("visibility", "visible");
            $(".tr_fase").css("display", "");
            $(".tr_fase").show(500);
        }

    }


    $(".sortFaseCHK").click(function () {

        if ($(this).is(':checked') == true) {
            $(".sortFaseCHK").attr("checked", "checked");
        } else {
            $(".sortFaseCHK").removeAttr("checked");
        }

    });


    $(".sortFase").keyup(function () {

        
       

        sorterOpdFn(this.id);

    });

  

    function sorterOpdFn(idThis) {
        
        //alert("her" + idThis)
        var thisid = idThis
        var idlngt = thisid.length
        var idtrim = thisid.slice(8, idlngt)
        var xval = idtrim

        var faseval = $("#fs_" + xval + "").val()
        var sortVal = $("#aktsort_" + xval + "").val()

        $("#sort_" + xval + "").html(sortVal)
      

        if (faseval != "") {

            if ($("#aktsort_chk_" + xval).is(':checked') == true) {
                antal_x = (document.getElementById("antal_x").value)
                for (i = 0; i < antal_x; i++) {
                    if ($("#fs_" + i + "").val() == faseval) {
                        $("#aktsort_" + i + "").val(sortVal)
                        $("#sort_" + i + "").html(sortVal)
                    }
                }

  
                
            }



        }
    
    }








    function GetBetIntVal() {
        var numdays;
        var intval = $("#FM_betbetint").val();

        (intval == "0") ? $("#FM_forfaldsdato").attr("disabled", true) : $("#FM_forfaldsdato").removeAttr("disabled");


        if (intval != "-100") {


            //var concatdate = $("#FM_start_dag").val() + "." + $("#FM_start_mrd").val() + "." + $("#FM_start_aar").val();
            var orgdate;
            $("[name=FM_brugfakdatolabel]").attr("checked") ? orgdate = $.datepicker.parseDate('d.m.yy', $("#FM_labeldato").val()) : orgdate = $.datepicker.parseDate('d.m.yy', $("#FM_fakdato").val()) //$.datepicker.parseDate('d.m.yy', concatdate)

            //alert(orgdate)
            //var newdate;
            lto = $("#lto").val()
            transport = $("#transport").val()

            if (intval == "-1") numdays = 0;
            if (intval == "20") numdays = 0;
            if (intval == "21") numdays = 0;
            if (intval == "1") numdays = 8;
            if (intval == "2") numdays = 14;
            if (intval == "5") numdays = 21;
            if (intval == "3") numdays = 30;
            if (intval == "6") numdays = 45;
            if (intval == "9") numdays = 60;
            if (intval == "14") numdays = 90;
            if (intval == "15") numdays = 120;

            if (lto == "intranet - local" || lto == "nt") {

                if (intval != "21") {

                    if (transport == "By Sea") {
                        numdays = numdays + 35
                    }

                    if (transport == "By Truck") {
                        numdays = numdays + 7
                    }

                }

            }

            var adval = 1

            //lbn + 15 dage
            if (intval == "4") {
                numdays = 0
                //var dd = orgdate.getDate()
                //alert(orgdate.getMonth() + 2)
                //alert(daysInMonth(orgdate.getMonth() + 2, orgdate.getYear()))
                orgdate.setDate(15);
                orgdate.setMonth(orgdate.getMonth() + 1);
                //orgdate.setDate(30)
                //alert(orgdate.setDate('30'));
                //$("#FM_forfaldsdato").datepicker('setDate', orgdate);
            }

            //lbn + 20 dage
            if (intval == "22") {
                numdays = 0
                //var dd = orgdate.getDate()
                //alert(orgdate.getMonth() + 2)
                //alert(daysInMonth(orgdate.getMonth() + 2, orgdate.getYear()))
                orgdate.setDate(20);
                orgdate.setMonth(orgdate.getMonth() + 1);
                //orgdate.setDate(30)
                //alert(orgdate.setDate('30'));
                //$("#FM_forfaldsdato").datepicker('setDate', orgdate);
            }

            //Lbn + 45 dage
            if (intval == "7") {
                numdays = 0
                //var dd = orgdate.getDate()
                //alert(orgdate.getMonth() + 2)
                //alert(daysInMonth(orgdate.getMonth() + 2, orgdate.getYear()))
                orgdate.setDate(15);
                orgdate.setMonth(orgdate.getMonth() + 2);
            }

            //lbn + 30 dage. Altid sidste bankdag. Aldrig ind i ny md.
            if (intval == "8") {
                if (orgdate.getMonth() == 0) {
                    var lastdayimt = 28;
                } else {
                    var lastdayimt = 30;
                }

                numdays = 0;
                //var dd = orgdate.getDate()
                //alert(orgdate.getMonth())
                //lastdayimt = daysInMonth(orgdate.getMonth() + 1, orgdate.getFullYear())
                //alert(lastdayimt)
                orgdate.setDate(lastdayimt);
                orgdate.setMonth(orgdate.getMonth() + 1);
                var adval = -1
                //alert(orgdate.getDate() + " " + (orgdate.getMonth() + 1) + " " + orgdate.getFullYear())
            }


            //lbn + 63 dage. Altid sidste bankdag. Aldrig ind i ny md.
            if (intval == "10") {
                if (orgdate.getMonth() == 0) {
                    var lastdayimt = 28;
                } else {
                    var lastdayimt = 30;
                }

                numdays = 0;
                //var dd = orgdate.getDate()
                //alert(orgdate.getMonth())
                //lastdayimt = daysInMonth(orgdate.getMonth() + 1, orgdate.getFullYear())
                //alert(lastdayimt)
                orgdate.setDate(3);
                orgdate.setMonth(orgdate.getMonth() + 3);
                var adval = -1
                //alert(orgdate.getDate() + " " + (orgdate.getMonth() + 1) + " " + orgdate.getFullYear())
            }


            //lbn + 60 dage. Altid sidste bankdag. Aldrig ind i ny md.
            if (intval == "11") {
                if (orgdate.getMonth() == 0) {
                    var lastdayimt = 28;
                } else {
                    var lastdayimt = 30;
                }

                numdays = 0;
                //var dd = orgdate.getDate()
                //alert(orgdate.getMonth())
                //lastdayimt = daysInMonth(orgdate.getMonth() + 1, orgdate.getFullYear())
                //alert(lastdayimt)
                orgdate.setDate(lastdayimt);
                orgdate.setMonth(orgdate.getMonth() + 2);
                var adval = -1
                //alert(orgdate.getDate() + " " + (orgdate.getMonth() + 1) + " " + orgdate.getFullYear())
            }


            //lbn + 90 dage. Altid sidste bankdag. Aldrig ind i ny md.
            if (intval == "12") {
                if (orgdate.getMonth() == 0) {
                    var lastdayimt = 28;
                } else {
                    var lastdayimt = 30;
                }

                numdays = 0;
                //var dd = orgdate.getDate()
                //alert(orgdate.getMonth())
                //lastdayimt = daysInMonth(orgdate.getMonth() + 1, orgdate.getFullYear())
                //alert(lastdayimt)
                orgdate.setDate(lastdayimt);
                orgdate.setMonth(orgdate.getMonth() + 3);
                var adval = -1
                //alert(orgdate.getDate() + " " + (orgdate.getMonth() + 1) + " " + orgdate.getFullYear())
            }

            //lbn + 120 dage. Altid sidste bankdag. Aldrig ind i ny md.
            if (intval == "13") {
                if (orgdate.getMonth() == 0) {
                    var lastdayimt = 28;
                } else {
                    var lastdayimt = 30;
                }

                numdays = 0;
                //var dd = orgdate.getDate()
                //alert(orgdate.getMonth())
                //lastdayimt = daysInMonth(orgdate.getMonth() + 1, orgdate.getFullYear())
                //alert(lastdayimt)
                orgdate.setDate(lastdayimt);
                orgdate.setMonth(orgdate.getMonth() + 4);
                var adval = -1
                //alert(orgdate.getDate() + " " + (orgdate.getMonth() + 1) + " " + orgdate.getFullYear())
            }


            //lbn + 62 dage. Altid sidste bankdag. Aldrig ind i ny md.
            if (intval == "16") {
                if (orgdate.getMonth() == 0) {
                    var lastdayimt = 28;
                } else {
                    var lastdayimt = 30;
                }

                numdays = 0;
                //var dd = orgdate.getDate()
                //alert(orgdate.getMonth())
                //lastdayimt = daysInMonth(orgdate.getMonth() + 1, orgdate.getFullYear())
                //alert(lastdayimt)
                orgdate.setDate(2);
                orgdate.setMonth(orgdate.getMonth() + 3);
                var adval = -1
                //alert(orgdate.getDate() + " " + (orgdate.getMonth() + 1) + " " + orgdate.getFullYear())
            }



            //alert(orgdate.getDate() + " " + numdays)
            orgdate.setDate(orgdate.getDate() + numdays);
            while (orgdate.getDay() == 6 || orgdate.getDay() == 7) {
                orgdate.setDate(orgdate.getDate() + (adval));
            }

            //alert("ger:" + intval + "#" + numdays)
            //alert(orgdate)
            $("#FM_forfaldsdato").datepicker('setDate', orgdate);



            //if (intval == "4") {
            //var dd = orgdate.getDate()
            //alert(orgdate.getDate())
            //$("#FM_forfaldsdato").datepicker('setDate', orgdate);
            //}

            //numdays = (daysInMonth(orgdate.getMonth()+1, orgdate.getYear())) + 15;
            //if (intval == "7") numdays = (daysInMonth(orgdate.getMonth(), orgdate.getYear()) - orgdate.getDate()) + 45;
            //if (intval == "8") numdays = daysInMonth(orgdate.getMonth(), orgdate.getYear()) + 30;



        }



    }

    GetBetIntVal();
    /// End ///







    // MaTERIALER //
    function GetMatTilFak() {

       

        if ($("#FM_mat_viskuneks1").is(':checked') == true) {
            var viskuneks = 1
        } else {
            var viskuneks = 0
        }

        if ($("#FM_mat_ignper").is(':checked') == true) {
            var ignper = 1
        } else {
            var ignper = 0
        }

        var valutaValId = $("#FM_valuta_all_3")
        var valutaValLabel = $("#FM_valuta_all_3:select option:selected").text() //'select option:selected').text()  
        //alert(valutaValLabel)
        var valutaValkurs = $("#valutakurs_" + valutaValId.val() + "")
        var kid_val = $("#FM_kid")
        var jobid_val = $("#FM_job")

        var jq_lto = $("#lto")

        //if (jq_lto.val() == 'essens') {
        //    alert(jobid_val.val())
        //}

        var aftid_val = $("#FM_aftale")
        var jq_func = $("#jq_func")

        var jq_stdato = $("#jq_stdato")
        var jq_sldato = $("#jq_sldato")
        var jq_id = $("#jq_id")
        var jq_intType = $("#jq_inttype")

        var jq_fastpris = $("#fastpris").val()
       

        $.post("?jq_lto=" + jq_lto.val() + "&jq_inttype=" + jq_intType.val() + "&jq_id=" + jq_id.val() + "&jq_stdato=" + jq_stdato.val() + "&jq_sldato=" + jq_sldato.val() + "&jq_func=" + jq_func.val() + "&jq_valutaValkurs=" + valutaValkurs.val() + "&jq_valutafaklabel=" + valutaValLabel + "&jq_kuneks=" + viskuneks + "&jq_ignper=" + ignper + "&jq_jobid=" + jobid_val.val() + "&jq_aftid=" + aftid_val.val() + "&jq_fastpris=" + jq_fastpris, { control: "FN_getMatTilFak", AjaxUpdateField: "true", cust: kid_val.val() }, function (data) {
            //$("#FM_modtageradr").val(data);
            $("#divmatreg").html(data);
            //$("#jobid").html(data);
            //alert($("#jq_matbeltot").val())

            //mat beløb tot.
            if ($("#jq_matbeltot").val() != "") {
                var mattot = $("#jq_matbeltot").val().replace(",", ".")
                mattot = Math.round(mattot * 100) / 100
            } else {
                var mattot = 0
            }
            var mattotTxt = String(mattot).replace(".", ",")
            $("#FM_materialer_beloeb").val(mattotTxt);
            $("#divmatbelobtot").html("<b>" + mattotTxt + " " + valutaValLabel + "</b>");

            var akttot = $("#FM_timer_beloeb").val().replace(",", ".")
            var totbel = Math.round((mattot / 1 + akttot / 1) * 100) / 100

            totbel = String(totbel).replace(".",",")


            //ialtbeløb tot
            $("#FM_beloeb").val(totbel);
            $("#divbelobtot").html("<b>" + valutaValLabel + "</b>");

            //totbel + " " +


            //mat beløb tot umoms.
            var mattot_umoms = Math.round($("#jq_matbeltot_umoms").val().replace(",", ".") * 100) / 100
            var akttot_umoms = $("#FM_timer_beloeb_umoms").val().replace(",", ".")
            var totbel_umoms = Math.round((mattot_umoms / 1 + akttot_umoms / 1) * 100) / 100
            totbel_umoms = String(totbel_umoms).replace(".", ",")
            //alert(totbel_umoms)
            //alert(akttot + " " + mattot + " " + totbel + " " + mattot_umoms)

            //divbelobtot_umoms
            //ialtbeløb tot
            $("#FM_beloeb_umoms").val(totbel_umoms);
            $("#divbelobtot_umoms").html("" & erp_txt_182 & ": <br> (" + totbel_umoms + " " + valutaValLabel + ") ");

        });

      

    }





    /// Sætter monsbeløb, hvis der sktives direkte i beløbs feltet
    $("#FM_beloeb").keyup(function () {

        var valutaValLabel = $("#FM_valuta_all_3:select option:selected").text()

        var totBelnoDot = $("#FM_beloeb").val() //.replace(".",",");

        var kCode = 0;

        if (totBelnoDot != 0) {
            //kCode = window.event.keyCode
            var e
            if (window.event) kCode = window.event.keyCode; // IE
            else if (e) kCode = e.which;
        } else {
            kCode = 0
        }


        if (kCode == 37) {
        } else {
            if (kCode == 39) {
            } else {

                passedVal = totBelnoDot

                invalidChars = "/+:;<>abcdefghijklmnopqrstuvwxyzæøå"
                if (passedVal == "") {
                    return false
                }

                for (i = 0; i < invalidChars.length; i++) {
                    badChar = invalidChars.charAt(i)
                    if (passedVal.indexOf(badChar, 0) != -1) {
                        alert("Der er angivet et ugyldigt tegn.")
                        return false
                    }
                }

                //$("#FM_beloeb").val(totBelnoDot)

                var totBel = Math.round($("#FM_beloeb").val().replace(",", ".") * 100) / 100;

                //var momsSats = $("#FM_momssats").val() / 100;
                //var momsBel = Math.round((totBel / 1 * momsSats / 1) * 100) / 100
                //totbel_umoms = String(momsBel).replace(".", ",")

                //totbel_umoms = altid = 0 når nettobeløb angive manuelt, timeOut beregner momsens ved indlæs opret/rediger i db
                $("#FM_beloeb_umoms").val(0);
                $("#divbelobtot_umoms").html("" & erp_txt_182 & ": <br> (0 " + valutaValLabel + ") ");

            }
        }

    });




    // MaTERIALER opdater IKKE total >= Ved rediger faktura //
    function GetMatTilFak_dontupdatetotal() {



        if ($("#FM_mat_viskuneks1").is(':checked') == true) {
            var viskuneks = 1
        } else {
            var viskuneks = 0
        }

        if ($("#FM_mat_ignper").is(':checked') == true) {
            var ignper = 1
        } else {
            var ignper = 0
        }

        var valutaValId = $("#FM_valuta_all_3")
        var valutaValLabel = $("#FM_valuta_all_3:select option:selected").text() //'select option:selected').text()  
        //alert(valutaValLabel)
        var valutaValkurs = $("#valutakurs_" + valutaValId.val() + "")
        var kid_val = $("#FM_kid")
        var jobid_val = $("#FM_job")
        var aftid_val = $("#FM_aftale")
        var jq_func = $("#jq_func")

        var jq_stdato = $("#jq_stdato")
        var jq_sldato = $("#jq_sldato")
        var jq_id = $("#jq_id")
        var jq_intType = $("#jq_inttype")
        var jq_lto = $("#lto")

        $.post("?jq_lto=" + jq_lto.val() + "&jq_inttype=" + jq_intType.val() + "&jq_id=" + jq_id.val() + "&jq_stdato=" + jq_stdato.val() + "&jq_sldato=" + jq_sldato.val() + "&jq_func=" + jq_func.val() + "&jq_valutaValkurs=" + valutaValkurs.val() + "&jq_valutafaklabel=" + valutaValLabel + "&jq_kuneks=" + viskuneks + "&jq_ignper=" + ignper + "&jq_jobid=" + jobid_val.val() + "&jq_aftid=" + aftid_val.val(), { control: "FN_getMatTilFak", AjaxUpdateField: "true", cust: kid_val.val() }, function (data) {
            //$("#FM_modtageradr").val(data);
            $("#divmatreg").html(data);
            //$("#jobid").html(data);
            //alert($("#jq_matbeltot").val())

            //mat beløb tot.
            if ($("#jq_matbeltot").val() != "") {
                var mattot = $("#jq_matbeltot").val().replace(",", ".")
                mattot = Math.round(mattot * 100) / 100
            } else {
                var mattot = 0
            }
            var mattotTxt = String(mattot).replace(".", ",")
            $("#FM_materialer_beloeb").val(mattotTxt);
            $("#divmatbelobtot").html("<b>" + mattotTxt + " " + valutaValLabel + "</b>");



        });



    }


    /// Beregner forbrug hvis det er opret og job ikke er fastpris //
    if ($("#func").val() != "red" && $("#fastpris").val() != "1" && $("#FM_aftale").val() == "0") {
        GetMatTilFak();
    } else {
        GetMatTilFak_dontupdatetotal();
    }











    $("#FM_mat_hentnye").click(function () {
        GetMatTilFak();
    });

    $("#FM_mat_viskuneks0").click(function () {
        GetMatTilFak();
    });

    $("#FM_mat_viskuneks1").click(function () {
        GetMatTilFak();
    });

    $("#FM_mat_ignper").click(function () {
        GetMatTilFak();
    });


    $("#nyekpers").click(function () {
        GetNyeKpers();
    });

    function GetNyeKpers() {

        var kid_val = $("#FM_kid")
        var att_val = $("#FM_altadr")

        $.post("?att_val=" + att_val.val(), { control: "FN_getKpers", AjaxUpdateField: "true", cust: kid_val.val() }, function (data) {
            //$("#fajl").val(data);
            $("#FM_altadr").html(data);
            $("#FM_att").html(data);
        });
    }




    function GetCustDesc(alertmoms) {

        //alert("HER" + lto)

        if ((lto != "intranet - local" && lto != "bf") || (alertmoms == 0 )) {

        if (document.getElementById("FM_usealtadr").checked == true) {
            var attkid = "att"
        } else {
            var attkid = "kid"
        }

        if (document.getElementById("FM_cvr_vis").checked == true) {
            var vis_cvr = 1
        } else {
            var vis_cvr = 0
        }

        if (document.getElementById("FM_land_vis").checked == true) {
            var vis_land = 1
        } else {
            var vis_land = 0
        }

        //alert(vis_cvr)
        var att_val = $("#FM_altadr")
        var kid_val = $("#FM_kid")

        //alert(attkid)
        //alert(att_val.val())



        //$("#BtnCustDescUpdate").data("cust", thisC.val());
        $.post("?reset=0&vis_land=" + vis_land + "&vis_cvr=" + vis_cvr + "&attkid=" + attkid + "&att_val=" + att_val.val(), { control: "FN_getCustDesc", AjaxUpdateField: "true", cust: kid_val.val() }, function (data) {
            $("#FM_modtageradr").val(data);
            $("#DIV_modtageradr").html(data);
            //$("#jobid").html(data);
        });



        if (attkid == 'XXX') { //att

            //kid_val.val()
            //alert("her:" + attkid + " att_val.val(): " + att_val.val())

            //Brug ikke denne, da det altid er hovedkunde der bestemmer om der skal moms på og ikke en evt. modtager adresse. /co Adresse.

            $.post("?att_val=" + att_val.val(), { control: "FN_getCustDesc_land", AjaxUpdateField: "true", cust: kid_val.val() }, function (data) {
                $("#momsland").val(data);

                setMomsSats();

            });
        } else {


            if (alertmoms == 1) {
                setMomsSats();
            }

        }

      
        } // lto
       


    }


   

    function setMomsSats() {

        //alert($("#momsland").val() + " : " + $("#afsmomsland").val())

        if ($("#momsland").val() != '') {

            //selAfsenderLand = $("#afsmomsland option:selected").val();
            //selAfsenderLand = $("#afsmomsland").find(":selected").val();
            //alert(selAfsenderLand)
            //$('#txtEntry2').text($(this).find(":selected").text());

            //alert($("#momsland").val() + " : " + $("#afsmomsland").val())


            if (($("#momsland").val() != $("#afsmomsland").val())) {

                if ($("#lto").val() != "epi2017") {
                    $("#FM_momssats").val(0)
                    //alert("Modtager land afviger fra faktura afsender land.\nMomssats er sat = 0")
                }

                
            }
        }
    }


    $("#FM_medregnikkeioms").change(function () {

        if ($("#FM_medregnikkeioms").val() == 2) {
            $("#FM_momssats").val(0)
        } else {


            $("#FM_momssats").val(1) //bør kunne variere hvis det er et UK land med andne momssats
            setMomsSats();

        }
    });



    $("#FM_att").change(function () {
        GetCustDesc(0);
    });

    $("#FM_cvr_vis").click(function () {
        GetCustDesc(0);
    });

    $("#FM_land_vis").click(function () {
        GetCustDesc(0);
    });

    $("#FM_usealtadr").click(function () {
        GetCustDesc(0);
    });

    $("#FM_altadr").change(function () {
        GetCustDesc(0);
    });


    // Opdater alt adr, ved opret fak og forvalgt alt modtager
    if ($("#altadrFunc").val() != "red") {
        //alert("her")
        GetCustDesc(1);
    }


    $("#red_adr").click(function () {
        $("#DIV_modtageradr").hide("fast");
        $("#FM_modtageradr").css("display", "");
        $("#FM_modtageradr").css("visibility", "visible");
        $("#FM_modtageradr").show(4000);
        $("#gem_adr").css("display", "");
        $("#gem_adr").css("visibility", "visible");
        $("#gem_adr").show(1000);
        $("#red_adr").hide("fast");

    });


    $("#gem_adr").click(function () {
        //alert("her")
        $("#DIV_modtageradr").html($("#FM_modtageradr").val())
        $("#DIV_modtageradr").css("display", "");
        $("#DIV_modtageradr").css("visibility", "visible");
        $("#DIV_modtageradr").show(4000);
        $("#FM_modtageradr").hide("fast");
        $("#gem_adr").hide("fast");
        $("#red_adr").show("fast");
    });



    $(".showemedarbLinier").click(function () {

        var thisid = this.id
        //alert($("." + thisid).css('display'))
        if ($("." + thisid).css('display') == "none") {

            $("." + thisid).css("display", "");
            $("." + thisid).css("visibility", "visible");
            $("." + thisid).show(4000);
            $.scrollTo("." + thisid, 500, { axis: 'y' });

        } else {

            $("." + thisid).hide(1000);


        }



    });

    $("#FM_fakdato").focus(function () {

        if ($("#showfakmsg").val() == 0) {
            $("#fakdatoinfo").css("display", "");
            $("#fakdatoinfo").css("visibility", "visible");
            $("#fakdatoinfo").show(1000);

            $("#showfakmsg").val(1);
        }


    });


    $(".qlinks").click(function () {

        
        $(".sumaktdiv").css("display", "none");
        $(".sumaktdiv").css("visibility", "hidden");
        $(".sumaktdiv").hide(100);

        $(".traktlink").css("background-color", "#FFFFFF");

        var thisid = this.id
        //alert(thisid)
        //alert($("." + thisid).css('display'))
        if ($("#sumakt_" + thisid).css('display') == "none") {

            $("#sumakt_" + thisid).css("display", "");
            $("#sumakt_" + thisid).css("visibility", "visible");
            $("#sumakt_" + thisid).show(1000);
            //alert("her")
            $("#tr_akt_" + thisid).css("background-color", "#FFFF99");

            
            //$.scrollTo("#sumakt_" + thisid, 300);
            //$.scrollTo("#sumakt_" + thisid, 500, { axis: 'y' });
            //$.scrollTo("#sumakt_" + thisid, 500, { axis: 'y' });
            $.scrollTo("#qlinkstop", 400);
            //$.scrollTo('720px', 500);

        } else {

            $("." + thisid).hide(500);
        }



    });

    $(".tiltoppen").click(function () {
        $(".sumaktdiv").hide(100);
        $.scrollTo('300px', 500); //'785px'
    });

    $(".naeste").click(function () {
        //var thisid = this.id
        $.scrollTo("#sumakt_" + thisid, '785px', { axis: 'x' });
    });

    //$("#menushowakt").click(function () {
    //    $.scrollTo('785px', 3000);
    //});


    lt0 = $("#lto").val()
    //checkAkt(0);
    if(lto == "nt" || lto == "intranet - local") {
        //opdatervalutAllelinier(1,0,0);
    }
   

});


///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////// END JQUERY /////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



function setmTxt(aktid, n, x) {
    //alert(n + " "+ x)
    medarbval = document.getElementById("FM_mid_" + n + "_" + x).value
    valeribrug = 0
    prevSelIndex = document.getElementById("prevSelIndex_" + n + "_" + x).value
    var antal_medarblinier = parseInt(document.getElementById("antal_n_" + x).value);
    //alert(n + " X " + antal_medarblinier)

    for (i = 1; i < antal_medarblinier; i++) {
        //alert(i + " x " + x + "medarbval " + medarbval)
        if (i != n) {

            if (document.getElementById("FM_mid_" + i + "_" + x).value == medarbval) {
                valeribrug = 1
                alert("Den valgte medarbejder er allerede ibrug på denne aktivitet.")
                document.getElementById("FM_mid_" + n + "_" + x).selectedIndex = prevSelIndex
                //document.getElementById("FM_show_medspec_" + n + "_" + x).checked = false
                //document.getElementById("FM_m_fak_" + n + "_" + x).value = 0
            }

        }
    }

    // Hvis ikke ibrug opdaer felter
    if (valeribrug == 0) {
        document.getElementById("FM_m_tekst_" + n + "_" + x).value = document.getElementById("a_" + aktid + "_" + medarbval).value
        document.getElementById("prevSelIndex_" + n + "_" + x).value = document.getElementById("FM_mid_" + n + "_" + x).selectedIndex 
    }

}
    
    
    //function showmenu(){
    //document.getElementById("menu").style.visibility = "visible"
    //document.getElementById("menu").style.display = ""
    
    //document.getElementById("sideload").style.visibility = "hidden"
    //document.getElementById("sideload").style.display = "none"
    //}
    
    
    function showvalutaload() {
    document.getElementById("valutaload").style.visibility = "visible"
    document.getElementById("valutaload").style.display = ""
    }
    
    function hidevalutaload() {
    document.getElementById("valutaload").style.visibility = "hidden"
    document.getElementById("valutaload").style.display = "none"
    }




    



    
    function opdatervalutAllelinier(a,x,m) {
    
        //alert("valuta funktion")
        //alert("her x:"+ x + " a:"+a )

      
   
        var valutaIdFak = document.getElementById("FM_valuta_all_"+a).value
	    var valutaKodeFak = document.getElementById("FM_valuta_all_"+a).options[document.getElementById("FM_valuta_all_"+a).selectedIndex].text
			    
	    document.getElementById("FM_valuta_all_1").value = valutaIdFak
	    document.getElementById("FM_valuta_all_2").value = valutaIdFak
	                
	            var valutaIdFak = document.getElementById("FM_valuta_all_1").value
	            var valutaKodeFak = document.getElementById("FM_valuta_all_1").options[document.getElementById("FM_valuta_all_1").selectedIndex].text
		        var valutaKursFak = document.getElementById("valutakurs_"+valutaIdFak).value
				var valutaKursBeregn = 0;
				var belobsubakt = 0;
				var belobsubakt_umoms = 0;
				   
				   //alert("a")     
				        
		                var antal_x = parseInt(document.getElementById("antal_x").value);


		                as = document.getElementById("antal_subtotal_akt_" + x + "").value
		                /// Aftale eller job

                        

	                    for (i=0;i<antal_x;i++){

	                        if (as != -1) {

	                            // Medarbejder felter
	                            antal_nx = (document.getElementById("antal_n_" + i + "").value)
	                            for (e = 1; e < antal_nx; e++) {

	                                valutaKode = document.getElementById("FM_mvaluta_" + e + "_" + i + "").options[document.getElementById("FM_mvaluta_" + e + "_" + i + "").selectedIndex].text
	                                valutaId = document.getElementById("FM_mvaluta_" + e + "_" + i + "").value
	                                valutaKurs = document.getElementById("valutakurs_" + valutaId + "").value.replace(",", ".")


	                                this_mtp = document.getElementById("FM_mtimepris_" + e + "_" + i + "").value.replace(",", ".")
	                                this_mtimer = document.getElementById("FM_m_fak_" + e + "_" + i + "").value.replace(",", ".")
	                                rabat = document.getElementById("FM_mrabat_" + e + "_" + i + "").value

	                                // Beregn beløb på medarb ved Valuta ændring i forhold til valuta på FAK //
	                                mBelobValBeregn(valutaKode, valutaKodeFak, valutaKurs, valutaKursFak, this_mtp, this_mtimer, rabat, e, i)

	                            }


	                        } else { //job eller aftale fak


	                            valutaKode = document.getElementById("FM_valuta_" + i + "_1").options[document.getElementById("FM_valuta_" + i + "_1").selectedIndex].text
	                            valutaId = document.getElementById("FM_valuta_" + i + "_1").value
	                            valutaKurs = document.getElementById("valutakurs_" + valutaId + "").value.replace(",", ".")


	                            this_mtp = document.getElementById("FM_enhedspris_" + i + "_1").value.replace(",", ".")
	                            this_mtimer = document.getElementById("FM_timerthis_" + i + "_1").value.replace(",", ".")
	                            rabat = document.getElementById("FM_rabat_" + i + "_1").value

	                            // Beregn beløb på medarb ved Valuta ændring i forhold til valuta på FAK //
	                            var e = 1;
	                            mBelobValBeregn(valutaKode, valutaKodeFak, valutaKurs, valutaKursFak, this_mtp, this_mtimer, rabat, e, i)

	                        }

                	        
	                        // Sum aktivitets felter
                            belobsubakt = 0;
                            belobsubakt_umoms = 0;
                            as_this = parseInt(document.getElementById("antal_subtotal_akt_"+i+"").value/1 + 1)
                            antal_n_x_this = parseInt(document.getElementById("antal_n_"+i+"").value/1 - 1)

                            //alert(as_this + "#" + antal_n_x_this)
                             
                            for (e=-(antal_n_x_this);e<as_this;e++){
                                    
                                //alert(e)    
                                if (as != -1 || e >= 0) { // job eller e > 0

                                //alert(i + "" + e)

                                //if (document.getElementById("FM_show_akt_"+i+"_"+e+"").checked == true) {
					
						        //aBelob = document.getElementById("FM_beloebthis_"+i+"_"+e+"").value.replace(",",".") 
						        valutaIda = document.getElementById("FM_valuta_"+i+"_"+e).value
                                
                                if (e != 0) {
                                valutaKodea = document.getElementById("valutadiv_"+i+"_"+e).innerText
                                } else {
                                valutaKodea = document.getElementById("FM_valuta_"+i+"_"+e).options[document.getElementById("FM_valuta_"+i+"_"+e).selectedIndex].text
		                        }
                                
                                valutaKursa = document.getElementById("valutakurs_"+valutaIda).value 
                                aTimer = document.getElementById("FM_timerthis_"+i+"_"+e+"").value.replace(",",".")
                                
                                if (e != 0) {
                                document.getElementById("timeprisdiv_" +i+"_"+e+"").innerHTML = "<b>"+document.getElementById("FM_timerthis_"+i+"_"+e+"").value+"</b>"
							    }   
							            
                                aEnhPris = document.getElementById("FM_enhedspris_"+i+"_"+e+"").value.replace(",",".")
                                aRabat = document.getElementById("FM_rabat_"+i+"_" +e+"").value.replace(",",".")
                                
                                
                                aBelob = parseFloat(aTimer) * parseFloat(aEnhPris) - (parseFloat(aTimer) * parseFloat(aEnhPris) * aRabat)
                                
                                               	                   
                	                        if (valutaKodea != valutaKodeFak) {
			                                 valutaKursBeregn = aBelob/1 * (valutaKursa.replace(",",".")/1 / valutaKursFak.replace(",",".")/1)
			                                //alert(valutaKursBeregn)
			                                //valutaKursBeregn = parseFloat(aBelob) * (parseFloat(valutaKursa) / parseFloat(valutaKursFak))
			                                } else {
			                                        if (parseFloat(valutaKursa) == parseFloat(valutaKursFak)) {
			                                        valutaKursBeregn = aBelob
			                                        } else {
			                                        valutaKursBeregn = ((aBelob/1) / (valutaKursFak/1)) * 100
			                                        }
			                                }



			                                valutaKursBeregn = Math.round(valutaKursBeregn * 100) / 100

			                                //alert("FM_show_akt_" + i + "_" + e + "")

			                                if (document.getElementById("FM_show_akt_" + i + "_" + e + "").checked == true) {
			                                    belobsubakt = belobsubakt / 1 + valutaKursBeregn / 1

                                                

			                                    if (document.getElementById("FM_momsfri_" + i + "_" + e + "").checked == true) {
			                                        belobsubakt_umoms = belobsubakt_umoms / 1 + valutaKursBeregn / 1
			                                    } else {
			                                    belobsubakt_umoms = belobsubakt_umoms
			                                    }
			                                        
			                                    
			                                } else {
			                                    belobsubakt = belobsubakt
			                                    belobsubakt_umoms = belobsubakt_umoms
			                                }
			                                
			                                valutaKursBeregn = String(valutaKursBeregn)
                        			        
                        			        
			                                document.getElementById("FM_beloebthis_"+i+"_"+e+"").value = valutaKursBeregn.replace(".",",") 
                        			        document.getElementById("belobdiv_"+i+"_"+e+"").innerHTML = "<b>"+valutaKursBeregn.replace(".",",")+" "+ valutaKodeFak +"</b>" 
                        			        
                        			      
                               } //job OR e > 0 
                                
                            } //for
                            
                            

                            //alert(i + " " + belobsubakt)
                                    

                            document.getElementById("belob_subtotal_akt_"+i+"").value = Math.round(belobsubakt*100)/ 100
                            document.getElementById("belob_subtotal_akt_"+i+"").value = document.getElementById("belob_subtotal_akt_"+i+"").value.replace(".",",")

                            document.getElementById("belob_subtotal_akt_umoms_"+i+"").value = Math.round(belobsubakt_umoms * 100) / 100
                            document.getElementById("belob_subtotal_akt_umoms_"+i+"").value = document.getElementById("belob_subtotal_akt_umoms_"+i+"").value.replace(".", ",")
                            
                            //document.getElementById("valutasubtdiv_"+i+"").innerHTML = valutaKodeFak  


                       
	                   
	                   } // For




                  

	   

        if (m != 0) {
				        
				        // Materialer 
	                    antal_mat = parseInt(document.getElementById("FM_antal_materialer_ialt").value/1)
	                    
	                    for (m=0;m<antal_mat;m++){
	                    //document.getElementById("FM_matvaluta_"+m).value = valutaIdFak
	                    beregnmatpris(m,0)
	                    document.getElementById("matvalutadiv_"+m).innerHTML = valutaKodeFak 
	                    }

        }

	                   
 
       

	                    ////setTimerTot2(x)
    setBeloebTot2(x,m)
    }	
    
    
    
    
    
    
    
	
	function overtilsum(x){
    var valutaKodeFak = document.getElementById("FM_valuta_all_1").options[document.getElementById("FM_valuta_all_1").selectedIndex].text
	var valutaIdFak = document.getElementById("FM_valuta_all_1").value
	            	
	var timer = 0;
	var beloeb = 0;
	var enhedspris = 0;
	var faklinjeTxt = "";

	faklinjeTxt = document.getElementById("txt_subtotal_akt_" + x).value
	timer = document.getElementById("timer_subtotal_akt_" + x).value.replace(",", ".") - (document.getElementById("FM_timerthis_" + x + "_0").value.replace(",", ".")) / 1
	beloeb = document.getElementById("belob_subtotal_akt_" + x).value.replace(",", ".") - (document.getElementById("FM_beloebthis_" + x + "_0").value.replace(",", ".")) / 1
    
    if (timer != 0 && beloeb != 0) {
        enhedspris = parseFloat(beloeb / timer)
        timer = Math.round(timer * 100) / 100
        beloeb = Math.round(beloeb * 100) / 100
	} else {
	enhedspris = 0
	timer = 0
	beloeb = 0
	}
	//enhedspris = 300

	
	enhedspris = Math.round(enhedspris * 100) / 100

	//faklinjeTxt
	document.getElementById("FM_aktbesk_" + x + "_0").value = faklinjeTxt
	
	document.getElementById("FM_timerthis_"+x+"_0").value = timer
	document.getElementById("FM_timerthis_"+x+"_0").value = document.getElementById("FM_timerthis_"+x+"_0").value.replace(".",",")
	
    ////document.getElementById("FM_hidden_timerthis"+x+"_0").value = timer //document.getElementById("FM_timerthis_"+x+"_0").value
	
	//alert(""+document.getElementById("FM_hidden_timerthis"+x+"_0").value+"")
	
	document.getElementById("FM_beloebthis_"+x+"_0").value = beloeb
	document.getElementById("FM_beloebthis_"+x+"_0").value = document.getElementById("FM_beloebthis_"+x+"_0").value.replace(".",",")
	document.getElementById("belobdiv_"+x+"_0").innerHTML = "<b>"+document.getElementById("FM_beloebthis_"+x+"_0").value+" "+ valutaKodeFak +"</b>"	
								
	
	document.getElementById("FM_enhedspris_"+x+"_0").value = enhedspris
	document.getElementById("FM_enhedspris_"+x+"_0").value = document.getElementById("FM_enhedspris_"+x+"_0").value.replace(".",",")
	
				    offSet_this = String(document.getElementById("FM_enhedspris_"+x+"_0").value);
					offSetL_this = String(document.getElementById("FM_enhedspris_"+x+"_0").length);
					pkt_this = offSet_this.indexOf(",");
					if (pkt_this > 0){
						document.getElementById("FM_enhedspris_"+x+"_0").value = offSet_this.substring(0, pkt_this + 3)
					}
		
	
	//// Opdaterer det samle beløb og time antalt på denne aktivitet ///
    //faklinjeTxt = document.getElementById("txt_subtotal_akt_" + x).value
					document.getElementById("timer_subtotal_akt_" + x).value = document.getElementById("FM_timerthis_" + x + "_0").value
					document.getElementById("belob_subtotal_akt_" + x).value = document.getElementById("FM_beloebthis_" + x + "_0").value
					//belob_subtotal_akt_umoms
					
					
	/// Sætter Andre sum-aktivitet aktiv 				
	document.getElementById("FM_show_akt_"+x+"_0").checked = true
	
	    /// Sætter sum-aktiviteter inaktive
	//antal_x = (document.getElementById("antal_x").value)
	//alert(antal_x)
	    //for (i=0;i<antal_x;i++){
	  
	  
	  
	        antal_highest_aval_m = parseInt(document.getElementById("highest_aval_m_"+x+"").value/1)
	        antal_highest_aval = parseInt(document.getElementById("highest_aval_" + x + "").value/1)
	        
	        
	        //if (antal_highest_aval > 1) {
	        //st = -antal_highest_aval-1
	        //} else {
	        st = -antal_highest_aval_m
	        //}

	        //alert(antal_highest_aval + "st: "+ st)

	        for (e = st; e <= antal_highest_aval; e++) {
	            if (e != 0) {
	                //alert("e: "+ e)
	                document.getElementById("FM_show_akt_" + x + "_" + e).checked = false
	            }
            }
            
	    //}
	    
	    
	        // Sætter medarbejdere inaktive 
	        antal_n_x = parseInt(document.getElementById("antal_n_"+x+"").value)
		    for (i=1;i<antal_n_x;i++){
				    if (i != 0){
				    document.getElementById("FM_show_medspec_"+i+"_"+x+"").checked = false
				    }
	         }
    	     
    	     
    	     document.getElementById("FM_valuta_"+x+"_0").value = valutaIdFak
    	     document.getElementById("FM_valuta_opr_"+x+"_0").value = valutaIdFak
	         document.getElementById("FM_rabat_"+x+"_0").value = "0"
	         //document.getElementById("FM_rabat_"+x+"_0").style.backgroundColor = "#ffff99"
    	     
	         //document.getElementById("angivtxt_"+x+"_0").style.visibility = "visible"
	         //document.getElementById("angivtxt_"+x+"_0").style.display = ""
	        
	     
	}
	
	
	
	
	//===//
	//function setbgcol(x){
	//document.getElementById("FM_rabat_"+x+"_0").style.backgroundColor = "#ffffff"
	//}
	
	
	
	
	function offsetSumTtimer(x){
    if (window.event.keyCode == 37){ 
			} else {
				if (window.event.keyCode == 39){
				} else {
				offSet_this = String(document.getElementById("timer_subtotal_akt_"+x).value.replace(".",","));
				offSetL_this = String(document.getElementById("timer_subtotal_akt_"+x).length);
				pkt_this = offSet_this.indexOf(",");
				if (pkt_this > 0){
					document.getElementById("timer_subtotal_akt_"+x).value = offSet_this.substring(0, pkt_this + 3)
				}
			}
		}
	}
	
	function offsetSumTbeloeb(x){
	if (window.event.keyCode == 37){ 
			} else {
				if (window.event.keyCode == 39){
				} else {
				offSet_this = String(document.getElementById("belob_subtotal_akt_"+x).value.replace(".",","));
				offSetL_this = String(document.getElementById("belob_subtotal_akt_"+x).length);
				pkt_this = offSet_this.indexOf(",");
				if (pkt_this > 0){
					document.getElementById("belob_subtotal_akt_"+x).value = offSet_this.substring(0, pkt_this + 3)
				}
			}
		}
	}
	
	
	function showmodtagerafsender() {
	document.getElementById("modtagerafsender").style.visibility = "visible"
	document.getElementById("modtagerafsender").style.display = ""
	}
	
	function hidemodtagerafsender() {
	document.getElementById("modtagerafsender").style.visibility = "hidden"
	document.getElementById("modtagerafsender").style.display = "none"
	}
	
	
	
	function opd_akt_endhed(nyenhed,val) {
	    var enh_val = 0;
	    enh_val = val
	    antal_x = (document.getElementById("antal_x").value)
	    //ehLabel = document.getElementById("FM_med_enh_" + n + "_" + x +"").options[document.getElementById("FM_med_enh_" + n + "_" + x +"").selectedIndex].text


	    as = document.getElementById("antal_subtotal_akt_0").value
	    /// Aftale eller job
	    
	        
	    
      

	    for (i=0;i<antal_x;i++){

	        if (as != -1) { //job

	            // Medarbejder felter
	            antal_nx = (document.getElementById("antal_n_" + i + "").value)
	            for (e = 1; e < antal_nx; e++) {
	                // Skal ikke kunne ændre enhed på Kørsels aktiviteter på medarb., fra den globale enhedsvælger
	                if (document.getElementById("FM_med_enh_" + e + "_" + i).value != '3') {
	                    document.getElementById("FM_med_enh_" + e + "_" + i).value = enh_val
	                }
	            }


	            // Aktiviteter
	            antal_highest_aval = parseInt(document.getElementById("highest_aval_" + i + "").value / 1 + 1)
	            for (e = 0; e < antal_highest_aval; e++) {


	                // Skal ikke kunne ændre enhed på Kørsels aktiviteter, fra den globale enhedsvælger
	                if (document.getElementById("FM_akt_enh_" + i + "_" + e).value != '3') {
	                    document.getElementById("FM_akt_enh_" + i + "_" + e).value = enh_val

	                    if (e != 0) {
	                        document.getElementById("ehdiv_" + i + "_" + e + "").innerHTML = nyenhed
	                    }

	                }
	                //nyenhed
	            }


	        } else {

	            for (e = 0; e < 21; e++) {


	                // Skal ikke kunne ændre enhed på Kørsels aktiviteter, fra den globale enhedsvælger
	                if (document.getElementById("FM_akt_enh_" + e + "_1").value != '3') {
	                    document.getElementById("FM_akt_enh_" + e + "_1").value = enh_val

	                   
	                }
	                //nyenhed
	            }

            
            }


            
	    }
	}
	


    //Gang timer med faktor på hver enkelt linje//
	function opd_timer_faktor(xval) {
	    var fak = 0;
	    var newfak = 0;
	    var faktorval = document.getElementById("aktfaktor_" + xval).value.replace(",", ".")
	        
       
	         
	         
	         
	        // Medarbejder felter
	        antal_nx = (document.getElementById("antal_n_"+xval+"").value) 
            for (e=1;e<antal_nx;e++){
                fak = document.getElementById("FM_m_fak_" + e + "_" + xval).value.replace(",", ".")
                newfak = Math.round((parseFloat(fak * faktorval))* 100) / 100
                document.getElementById("FM_m_fak_" + e + "_" + xval).value = newfak
                //document.getElementById("FM_m_fak_"+e+"_"+xval).value = document.getElementById("FM_m_fak_"+e+"_"+xval).value.replace(",",".")
            
                enhedsprismedarb(xval, e)

                document.getElementById("FM_m_fak_" + e + "_" + xval).value = String(newfak).replace(".", ",")
                
            }

            
            //alert("her 1")

	    //fak = document.getElementById("FM_timerthis_" + xval + "_0").value.replace(",", ".")
	    //   newfak = Math.round((parseFloat(fak * faktorval)) * 100) / 100
	    //    document.getElementById("FM_timerthis_" + xval + "_0").value = newfak
            
	    //   andreEnhprs(xval, 0);


	    
            
	}
	
	
	
	
	
	
	
	
	function opd_rabatrall() {
	    val = document.getElementById("FM_rabat_all").value
	    antal_x = (document.getElementById("antal_x").value)

	    as = document.getElementById("antal_subtotal_akt_0").value
	    /// Aftale eller job
	    
	    for (i=0;i<antal_x;i++){


	        if (as != -1) { //job

	            // Sum aktivitets felter        
	            as_this = parseInt(document.getElementById("antal_subtotal_akt_" + i + "").value / 1 + 1)
	            antal_n_x_this = parseInt(document.getElementById("antal_n_" + i + "").value / 1 - 1)

	            for (e = -(antal_n_x_this); e < as_this; e++) {

	                //   //alert("ok")
	                document.getElementById("FM_rabat_" + i + "_" + e).value = val
	                if (e != 0) {
	                    //if (document.getElementById("FM_show_akt_"+i+"_"+e+"").checked == true) {
	                    document.getElementById("rabatdiv_" + i + "_" + e).innerHTML = (val * 100) + " %"
	                    //}
	                }

	            }

	        } else {


	            for (e = 0; e < 21; e++) {
	                document.getElementById("FM_rabat_" + e + "_1").value = val
	            }

	        }
            
	    }
	    
	    // Materialer 
	    antal_mat = parseInt(document.getElementById("FM_antal_materialer_ialt").value/1)
	    for (m=0;m<antal_mat;m++){
	    document.getElementById("FM_matrabat_"+m).value = val
	    }
	    
	    //alert("Medarbejder linier opdateret!")
	    
	    opdatervalutAllelinier(1,0,1)
	    
	}
	
	
	
	

	
	
	function genberegntimeprissumakt(x){
	var timer = 0;
	var beloeb = 0;
	var enhedspris = 0;
	var kvotient = 0;
	
	beloeb = document.getElementById("FM_beloebthis_"+x+"_0").value.replace(",",".")
	timer = document.getElementById("FM_timerthis_"+x+"_0").value.replace(",",".")
	
	var rbt = 0;
	rbt = document.getElementById("FM_rabat_"+x+"_0").value 
	enhedspris = (beloeb/timer) / (1-rbt)
	
	
	
	
				    offSet_this = String(enhedspris);
					offSetL_this = String(enhedspris.length);
					pkt_this = offSet_this.indexOf(".");
					if (pkt_this > 0){
					document.getElementById("FM_enhedspris_"+x+"_0").value = offSet_this.substring(0, pkt_this + 3)
					} else {
					document.getElementById("FM_enhedspris_"+x+"_0").value = enhedspris
					}
	
	document.getElementById("FM_enhedspris_"+x+"_0").value = document.getElementById("FM_enhedspris_"+x+"_0").value.replace(".",",")
	
	}
	
	
	
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	
	
	function decimaler(pm, n, x){
	document.getElementById("FM_mbelob_"+n+"_"+x+"").value = document.getElementById("FM_mbelob_"+n+"_"+x+"").value.replace(",",".")
	
		if (pm == 'plus'){
		document.getElementById("FM_mbelob_"+n+"_"+x+"").value = parseFloat(document.getElementById("FM_mbelob_"+n+"_"+x+"").value) + parseFloat("0.01")
		}else{
		document.getElementById("FM_mbelob_"+n+"_"+x+"").value = parseFloat(document.getElementById("FM_mbelob_"+n+"_"+x+"").value) - parseFloat("0.01")
		}
	
		document.getElementById("FM_mbelob_"+n+"_"+x+"").value = document.getElementById("FM_mbelob_"+n+"_"+x+"").value.replace(".",",")
	
			
							//afrunder decimaler
							offSet_this = String(document.getElementById("FM_mbelob_"+n+"_"+x+"").value);
							offSetL_this = String(document.getElementById("FM_mbelob_"+n+"_"+x+"").length);
							pkt_this = offSet_this.indexOf(",");
							if (pkt_this > 0){
								document.getElementById("FM_mbelob_"+n+"_"+x+"").value = offSet_this.substring(0, pkt_this + 3)
							}
			
	
		        document.getElementById("medarbbelobdiv_"+n+"_"+x+"").innerHTML = "<b>"+ document.getElementById("FM_mbelob_"+n+"_"+x+"").value +"</b>"
	
	
				antal_n_x = parseInt(document.getElementById("antal_n_"+x+"").value)
				highest_aval = parseInt((document.getElementById("highest_aval_"+x+"").value/1) + 1)
				antal_n_x2 = parseInt(document.getElementById("antal_n_"+x+"").value/1 - 1)
				
				//Finder medarbejder timepriser
				this_mtp = document.getElementById("FM_mtimepris_" + n + "_" + x +"").value.replace(",",".")
				
					var tmpbel = 0;
					var nyBeloebtot = 0;
					//Samler det totale antal timer på den enhedspris der er valgt
					for (i=1;i<antal_n_x;i++){
						if (i != 0) {
							if (parseFloat(this_mtp) == parseFloat(document.getElementById("FM_mtimepris_" + i + "_" + x +"").value.replace(",","."))) {
							tmpbel = document.getElementById("FM_mbelob_" + i + "_" + x +"").value.replace(",",".")
							nyBeloebtot = nyBeloebtot + parseFloat(tmpbel)
							}
						}
					}
					
					//Opdaterer den valgte sum-aktivitet (Den der passer til timeprisen)
					for (i=-(antal_n_x2);i<highest_aval;i++){
							
							if (i != 0) {
								if (parseFloat(this_mtp) == parseFloat(document.getElementById("FM_enhedspris_"+x+"_"+i+"").value.replace(",","."))) {
								//Totaler på den valgte sum-aktivitet
								document.getElementById("FM_beloebthis_"+x+"_"+i+"").value = nyBeloebtot 
								document.getElementById("FM_beloebthis_"+x+"_"+i+"").value = document.getElementById("FM_beloebthis_"+x+"_"+i+"").value.replace(".",",")
								
								document.getElementById("belobdiv_"+x+"_"+i+"").innerHTML = "<b>"+document.getElementById("FM_beloebthis_"+x+"_"+i+"").value+"</b>"	
								}
							}	
					}
	
	setBeloebTot2(x,1)
	}
	
	
	
	
	
	//Beregning af Andre? Sum aktiviteten
	function setBeloebThis2(x,a){


	   
	
	var valutaKode = document.getElementById("FM_valuta_"+x+"_"+a+"").options[document.getElementById("FM_valuta_"+x+"_"+a+"").selectedIndex].text
    var valutaId = document.getElementById("FM_valuta_"+x+"_"+a+"").value
    var valutaKurs = document.getElementById("valutakurs_"+valutaId+"").value.replace(",",".")
    
    
    if (a == 0) {

        
        valutaIdOpr = document.getElementById("FM_valuta_opr_" + x + "_" + a + "").value

        //alert("her x:" + x + "a:" + a + " valutaIdOpr: " + valutaIdOpr)
        valutaKursOpr = document.getElementById("valutakurs_" + valutaIdOpr + "").value.replace(",", ".")

    

    } else {
    valutaIdOpr = valutaId
    valutaKursOpr = valutaKurs
    }
	
	var belobThis = 0;
	
	
	
	
	timerThis = document.getElementById("FM_timerthis_"+x+"_"+a+"").value.replace(",",".")
	timeprisThis = document.getElementById("FM_enhedspris_"+x+"_"+a+"").value.replace(",",".")
	this_mtp = timeprisThis
	
    rabat = document.getElementById("FM_rabat_"+x+"_"+a+"").value
    
                // Bregn timepris ved Valuta ændring //
			    if (valutaId != valutaIdOpr) {
			    //alert(valutaKode +"!="+ valutaKodeOpr)
			    var this_mtp_valutakursberegn = 0;
			    this_mtp_valutakursberegn = this_mtp/1 * (valutaKursOpr.replace(",",".")/1 / valutaKurs.replace(",",".")/1)
			    this_mtp = Math.round(this_mtp_valutakursberegn*100)/ 100
			    }
    
    timeprisThis = this_mtp
    
    //alert(timerThis +" "+timeprisThis + " Rb: "+ rabat)
	belobThis = timerThis * timeprisThis - (timerThis * timeprisThis * rabat)
	
	//alert(belobThis)
	
								    
	
						
						    //afrunder decimaler
							timerThis = Math.round(timerThis*100)/100
				            document.getElementById("FM_timerthis_"+x+"_"+a+"").value = timerThis
				            document.getElementById("FM_timerthis_"+x+"_"+a+"").value = document.getElementById("FM_timerthis_"+x+"_"+a+"").value.replace(".",",")
            			
							
							//afrunder decimaler
							belobThis = Math.round(belobThis*100)/100
				            document.getElementById("FM_beloebthis_"+x+"_"+a+"").value = belobThis
				            document.getElementById("FM_beloebthis_"+x+"_"+a+"").value = document.getElementById("FM_beloebthis_"+x+"_"+a+"").value.replace(".",",")
            			    
            			    document.getElementById("belobdiv_"+x+"_"+a+"").innerHTML = "<b>"+document.getElementById("FM_beloebthis_"+x+"_"+a+"").value+" "+ valutaKode +"</b>"	
	
							
							
							
					
	
	document.getElementById("FM_valuta_opr_"+x+"_"+a+"").value = valutaId


	

	setBeloebTot2(x,1)				
	//opdatervalutAllelinier(1,x)	
	}







	//// Main medarbejder timer til sum aktiviteter /////
	function setTimerTot2(x) {

	}
	
	




	function setAllmomsFri(x, a) {

	    as = document.getElementById("antal_subtotal_akt_" + x + "").value
	    antal_n_x = parseInt(document.getElementById("antal_n_" + x + "").value / 1 - 1)
        thisCHK = document.getElementById("FM_momsfri_" + x + "_" + a + "").checked

	    //alert("as: " + as + " antal_n_x: " + antal_n_x + "x:" +x)


	    for (i = -(antal_n_x); i <= as; i++) {
	            
	            if (thisCHK == true) {
	                document.getElementById("FM_momsfri_" + x + "_" + i + "").checked = true
	                } else {
	                document.getElementById("FM_momsfri_" + x + "_" + i + "").checked = false
	                }
	            


	        }
	}
	
	
	
	
	
	
	//// Main medarbejder timepris og beløb og materialer til sum aktiviteter /////
	function setBeloebTot2(x,m) {

	    

	var tmpmatbelob = 0;
	var belobmattot = 0;
	var belobsubakt = 0;
	var belobakttot = 0;
	var tmpmatbelob = 0;
	var belobmattot = 0;
	var aktbelIaltUmoms = 0;
	var belobakttot_umoms = 0;
	var tmpbeluMoms = 0;
	var kCode = 0;
	var aTimerTot = 0;
	//var aTimerSubTot = 0;
	var timerakttot = 0;
	var valutaKodeFak = document.getElementById("FM_valuta_all_1").options[document.getElementById("FM_valuta_all_1").selectedIndex].text
	
	if (x != 0) {
	    //kCode = window.event.keyCode
	    var e
	    if (window.event) kCode = window.event.keyCode; // IE
	    else if (e) kCode = e.which;

	} else {
	kCode = 0
	} 
	
    


    if (kCode == 37){ 
	} else {
		if (kCode == 39){
		} else {
				as = document.getElementById("antal_subtotal_akt_"+x+"").value
				antal_n_x = parseInt(document.getElementById("antal_n_"+x+"").value/1 - 1)


				//alert("as: " + as + " antal_n_x: " + antal_n_x + "x:" +x)
                /// Aftale eller job 
				if (as == -1) { //aftale
                        
				        i = 1
				        
                        //alert(x + " " + i)
                        if (document.getElementById("FM_show_akt_" + x + "_" + i + "").checked == true) {

				            aTimer = document.getElementById("FM_timerthis_" + x + "_" + i + "").value.replace(",", ".")
				            aTimerTot = aTimerTot + aTimer / 1

                            //alert(aTimerTot)

				            aBelob = document.getElementById("FM_beloebthis_" + x + "_" + i + "").value.replace(",", ".")
				            belobsubakt = belobsubakt + aBelob / 1

				            //alert(aBelob)

				            // Beløb u moms //
				            if (document.getElementById("FM_momsfri_" + x + "_" + i + "").checked == true) {
				                aktbelIaltUmoms = (aktbelIaltUmoms / 1) + (document.getElementById("FM_beloebthis_" + x + "_" + i + "").value.replace(",", ".") / 1)
				                //alert("her: " + aktbelIaltUmoms)
				            }


				        }
				   

				} else {

				       
				    for (i = -(antal_n_x); i <= as; i++) {
				       
                        if (document.getElementById("FM_show_akt_" + x + "_" + i + "").checked == true) {


                          
				            aTimer = document.getElementById("FM_timerthis_" + x + "_" + i + "").value.replace(",", ".")
				            aTimerTot = aTimerTot + aTimer / 1

				            aBelob = document.getElementById("FM_beloebthis_" + x + "_" + i + "").value.replace(",", ".")
				            belobsubakt = belobsubakt + aBelob / 1

				            //alert(aBelob)

				            // Beløb u moms //
				            if (document.getElementById("FM_momsfri_" + x + "_" + i + "").checked == true) {
				                aktbelIaltUmoms = (aktbelIaltUmoms / 1) + (document.getElementById("FM_beloebthis_" + x + "_" + i + "").value.replace(",", ".") / 1)
				                //alert("her: " + aktbelIaltUmoms)
				            }


				        }
				    }

				}


				

                /////////////////////////////////
				//** Total beløb på faktura ***//
				////////////////////////////////

				

				        //Sub-total på akt.
				        belobsubakt = Math.round(belobsubakt * 100) / 100
				        //alert(belobsubakt)

				

				            // Timer tot på akt ///
				            aTimerTot = Math.round(aTimerTot * 100) / 100
				            document.getElementById("timer_subtotal_akt_" + x + "").value = aTimerTot
				            document.getElementById("timer_subtotal_akt_" + x + "").value = document.getElementById("timer_subtotal_akt_" + x + "").value.replace(".", ",")

				            if (as != -1) { //job

				            var thisAktId = document.getElementById("aktId_n_" + x + "").value
				            document.getElementById("timeae_akt_" + thisAktId + "").innerHTML = " => " + document.getElementById("timer_subtotal_akt_" + x + "").value

				            }
				

				        document.getElementById("belob_subtotal_akt_"+x+"").value = belobsubakt
				        document.getElementById("belob_subtotal_akt_"+x+"").value = document.getElementById("belob_subtotal_akt_"+x+"").value.replace(".",",")

				        //alert(aktbelIaltUmoms)
				        belobsubaktUmoms = Math.round(aktbelIaltUmoms * 100) / 100
				        //alert(belobsubaktUmoms)
				        document.getElementById("belob_subtotal_akt_umoms_" + x + "").value = belobsubaktUmoms
				        document.getElementById("belob_subtotal_akt_umoms_" + x + "").value = document.getElementById("belob_subtotal_akt_umoms_" + x + "").value.replace(".", ",")
						
				        
				        


				        // Sum af sum aktivter pr. aktivitet ///

				        //if (as != -1) { //job
				            antalx = document.getElementById("antal_x").value
                            
				            //alert(antalx)
            
				            for (i = 0; i < antalx; i++) {

				                //tmptim = document.getElementById("timer_subtotal_akt_" + i + "").value.replace(",", ".")
				                //timerakttot = timerakttot + parseFloat(tmptim)

				                //if (document.getElementById("FM_show_akt_" + x + "_" + i + "").checked == true) {
				                tmptim2 = document.getElementById("belob_subtotal_akt_" + i + "").value.replace(",", ".")
				                belobakttot = belobakttot + parseFloat(tmptim2)

				                tmpbeluMoms = document.getElementById("belob_subtotal_akt_umoms_" + i + "").value.replace(",", ".")
				                belobakttot_umoms = belobakttot_umoms + parseFloat(tmpbeluMoms)
				                //}

				            }

				      

				           

				            if (m != 0) {


				                // MATERIALER
				                antalm = document.getElementById("FM_antal_materialer_ialt").value
				                antalm = antalm

				                var matBelobialt_umoms = 0;

				                for (i = 0; i < antalm; i++) {

				                    if (document.getElementById("FM_vis_" + i + "").checked == true) {
				                        tmpmatbelob = document.getElementById("FM_matbeloeb_" + i + "").value.replace(",", ".")
				                        belobmattot = belobmattot + tmpmatbelob / 1


				                        if (document.getElementById("FM_matikkemoms_" + i + "").checked == true) {
				                            matBelobialt_umoms = (matBelobialt_umoms / 1) + (document.getElementById("FM_matbeloeb_" + i).value.replace(",", ".") / 1)
				                        }

				                    }
				                }


				                belobmattot = Math.round(belobmattot * 100) / 100
				                document.getElementById("FM_materialer_beloeb").value = belobmattot
				                document.getElementById("FM_materialer_beloeb").value = document.getElementById("FM_materialer_beloeb").value.replace(".", ",")
				                matBelobialtHTML = document.getElementById("FM_materialer_beloeb").value.replace(".", ",")
				                document.getElementById("divmatbelobtot").innerHTML = "<b>" + matBelobialtHTML + " " + valutaKodeFak + "</b>"


				            }

          

				    // Total beløb //
				    //alert(matBelobialt_umoms)
				    matBelobialt_umoms = Math.round(matBelobialt_umoms * 100) / 100
				    belobakttot_umoms = Math.round(belobakttot_umoms * 100) / 100
				    document.getElementById("FM_beloeb_umoms").value = (matBelobialt_umoms / 1) + (belobakttot_umoms / 1)
				    document.getElementById("FM_beloeb_umoms").value = document.getElementById("FM_beloeb_umoms").value.replace(".", ",")

				    
				    belobakttot = Math.round(belobakttot * 100) / 100
				    document.getElementById("FM_timer_beloeb").value = belobakttot
				    document.getElementById("FM_timer_beloeb").value = document.getElementById("FM_timer_beloeb").value.replace(".", ",")
				    document.getElementById("FM_timer_beloeb_umoms").value = belobsubaktUmoms
				    document.getElementById("FM_timer_beloeb_umoms").value = document.getElementById("FM_timer_beloeb_umoms").value.replace(".", ",")    
				    

                    
				    belobtotal = belobakttot + belobmattot
				    belobtotal = Math.round(belobtotal * 100) / 100
				    document.getElementById("FM_beloeb").value = belobtotal
				    document.getElementById("FM_beloeb").value = document.getElementById("FM_beloeb").value.replace(".", ",")


				    document.getElementById("divbelobtot").innerHTML = "<b>" + valutaKodeFak + "</b>"
				    //+ document.getElementById("FM_beloeb").value + " " 
				    document.getElementById("divtimerbelobtot").innerHTML = "<b>" + document.getElementById("FM_timer_beloeb").value + " " + valutaKodeFak + "</b>"
				    document.getElementById("divbelobtot_umoms").innerHTML = "" & erp_txt_182 & ": <br> (" + document.getElementById("FM_beloeb_umoms").value + " " + valutaKodeFak + ") "
				    
				    
				    if (lto != 'nt' && lto != 'intranet - local') {

				        if ($("#FM_job").val() != 0) {
				            berengFaktorBelobUdgor();
				        }
				    }
                
				    //alert("her 2")

				}
	  		}
	}











	
	
	//Bruges af alle sumaktiviteteterne til at tjekke for dublet timepriser //	
	function andreEnhprs(x, a) {

    

	var kCode = 0;
    

	if (x != 0) {

	    var e
        if (window.event) kCode = window.event.keyCode; // IE
	    else if (e) kCode = e.which;
	  
	
	} else {
    kCode = 0
	} 
	
    
    if (kCode == 37){ 
	} else {
		    if (kCode == 39){
		    } else {
		    
		    //Faktura valuta 
				
				
				
				//Tjekker om timepris allerede eksisterer
				//antal_n_x = parseInt(document.getElementById("antal_n_"+x+"").value)
				highest_aval = parseInt((document.getElementById("highest_aval_"+x+"").value/1) + 1)
				antal_n_x2 = parseInt(document.getElementById("antal_n_"+x+"").value/1 - 1)
				
		        
		        for (i=-(antal_n_x2);i<highest_aval;i++){
					//if (i != 0) {
					//alert(parseFloat(document.getElementById("FM_enhedspris_"+x+"_"+a+"").value)+" == "+ parseFloat(document.getElementById("FM_enhedspris_"+x+"_"+i+"").value.replace(",",".")))


		            if (parseFloat(document.getElementById("FM_enhedspris_" + x + "_" + a + "").value.replace(",", ".")) == parseFloat(document.getElementById("FM_enhedspris_" + x + "_" + i + "").value.replace(",", ".")) && (document.getElementById("FM_valuta_" + x + "_" + a + "").value == document.getElementById("FM_valuta_" + x + "_" + i + "").value) && (document.getElementById("FM_show_akt_" + x + "_" + i + "").checked == true) && (i != a)) {
							
							alert ("Denne timepris er allerede ibrug på en sum-aktivitet. Sum-aktvitets linien nulstilles.")
							//document.getElementById("FM_enhedspris_" + x + "_"+a+"").value = -2
							//document.getElementById("FM_timerthis_" + x + "_"+a+"").value = 0
							//document.getElementById("FM_beloebthis_" + x + "_"+a+"").value = 0
							//document.getElementById("belobdiv_"+x+"_"+a+"").innerHTML = "<b>"+document.getElementById("FM_beloebthis_"+x+"_"+a+"").value+" "+ valutaKodeFak +"</b>"	
	                        document.getElementById("FM_show_akt_"+x+"_"+a+"").checked = false
							//document.getElementById("FM_rabat_" + x + "_"+a+"").value = 0
							
							}
					//}	
				}
				
				           
							
						   
	       
    	
	        }
	    }
		
    //alert("andre ehnpris 2")

	}	
		
		
	
	
	
	
	
	
	
	/// Denne funktion er slået fra 30112009 ///
	function nulstilfaktimer(x,a){
	if (a != 0){
		
			if (document.getElementById("FM_show_akt_"+x+"_"+a+"").checked == false) {
				
				if (document.getElementById("FM_showalert").value == 0){
				var t = confirm("Når en sum-aktivitet ikke faktureres, bliver alle fakturerbare medarbejder-timer der hører til denne sum-aktivitet automatisk overført til 'vente' timer.\nDenne meddelelse bliver kun vist en gang pr. faktura oprettelse.")
				document.getElementById("FM_showalert").value = 1
				
				//Hvis alert vises.	
					if (t) {
					document.getElementById("FM_timerthis_" +x+"_"+a+"").value = 0,00
					document.getElementById("FM_beloebthis_" +x+"_"+a+"").value = 0,00
					
					document.getElementById("timeprisdiv_" +x+"_"+a+"").innerHTML = "<b>0,00</b>"
							    
					
					document.getElementById("belobdiv_"+x+"_"+a+"").innerHTML = "<b>0,00</b>"	
					
					//rbt
					document.getElementById("rabatdiv_"+x+"_"+a+"").innerHTML = "<b>0 %</b>"					
					
					this_akttp = document.getElementById("FM_enhedspris_"+x+"_"+a+"").value.replace(",",".")
					antal_n_x = parseInt(document.getElementById("antal_n_"+x+"").value)
								
								for (i=1;i<antal_n_x;i++){
									//alert(parseFloat(this_akttp) +" == "+ parseFloat(document.getElementById("FM_mtimepris_" + i + "_" + x +"").value.replace(",",".")))
									if (i != 0){
										if (parseFloat(this_akttp) == parseFloat(document.getElementById("FM_mtimepris_" + i + "_" + x +"").value.replace(",","."))) {
										document.getElementById("FM_m_vent_" +i+ "_"+x+"").value = document.getElementById("FM_m_fak_" +i+"_"+x+"").value
										document.getElementById("FM_mbelob_" +i+ "_"+x+"").value = 0,00
										document.getElementById("medarbbelobdiv_"+i+"_"+x+"").innerHTML = "<b>0,00</b>"	
										document.getElementById("FM_m_fak_" +i+"_"+x+"").value = 0,00
										document.getElementById("FM_show_medspec_"+i+"_"+x+"").checked = false
										}
									}
								}
					
					} else {
					document.getElementById("FM_show_akt_"+x+"_"+a+"").checked = true
					}
				
				
				//Hvis alert er vist 1 gang!
				} else {
					
				 	document.getElementById("FM_timerthis_" +x+"_"+a+"").value = 0,00
					document.getElementById("FM_beloebthis_" +x+"_"+a+"").value = 0,00
					
					document.getElementById("timeprisdiv_" +x+"_"+a+"").innerHTML = "<b>0,00</b>"
				    document.getElementById("belobdiv_"+x+"_"+a+"").innerHTML = "<b>0,00</b>"	
					//rbt
					document.getElementById("rabatdiv_"+x+"_"+a+"").innerHTML = "<b>0 %</b>"
							
					
										
					this_akttp = document.getElementById("FM_enhedspris_"+x+"_"+a+"").value.replace(",",".")
					antal_n_x = parseInt(document.getElementById("antal_n_"+x+"").value)
								
								for (i=1;i<antal_n_x;i++){
									//alert(parseFloat(this_akttp) +" == "+ parseFloat(document.getElementById("FM_mtimepris_" + i + "_" + x +"").value.replace(",",".")))
										if (i != 0){
										    if (parseFloat(this_akttp) == parseFloat(document.getElementById("FM_mtimepris_" + i + "_" + x +"").value.replace(",","."))) {
										    document.getElementById("FM_m_vent_" +i+ "_"+x+"").value = document.getElementById("FM_m_fak_" +i+"_"+x+"").value
										    document.getElementById("FM_mbelob_" +i+ "_"+x+"").value = 0,00
										    document.getElementById("medarbbelobdiv_"+i+"_"+x+"").innerHTML = "<b>0,00</b>"	
										    document.getElementById("FM_m_fak_" +i+"_"+x+"").value = 0,00
										    document.getElementById("FM_show_medspec_"+i+"_"+x+"").checked = false
										    }
										}
								}
					
					}
			}
	}
}
	
	
	
	
	
	//////////////////////////////////////
	////Medarbejder beløb funktion //////
	// Beregn beløb på medarb ved Valuta ændring i forhold til valuta på FAK //
	//////////////////////////////////////
    function mBelobValBeregn(valutaKode, valutaKodeFak, valutaKurs, valutaKursFak, this_mtp, this_mtimer, rabat, n, x) {
	            
	            if (valutaKode != valutaKodeFak) {
			    
			    var this_belob_valutakursberegn = 0;
			    this_belob_valBeregn = this_mtp/1 * (valutaKurs.replace(",",".")/1 / valutaKursFak.replace(",",".")/1)
			    this_valuta_mtp = this_belob_valBeregn
			    } else {
			    this_valuta_mtp = this_mtp
			    }
				
				
				// Medarbejderbeløb //
				var this_mbelob = (this_valuta_mtp * this_mtimer -(this_valuta_mtp * this_mtimer * rabat))
				this_mbelob = Math.round(this_mbelob*100)/ 100
				if (this_mbelob != 0) {
				this_mbelob = String(this_mbelob)
				} else {
				this_mbelob = "0.00"
				}

                as = document.getElementById("antal_subtotal_akt_" + x + "").value
                if (as != -1) { //job

                    document.getElementById("FM_mbelob_" + n + "_" + x + "").value = this_mbelob.replace(".", ",")
                    document.getElementById("medarbbelobdiv_" + n + "_" + x + "").innerHTML = "<b>" + document.getElementById("FM_mbelob_" + n + "_" + x + "").value + " " + valutaKodeFak + "</b>"

                } else {
                 
                    //alert(n + " "+ x)
                    document.getElementById("FM_beloebthis_" + x + "_1").value = this_mbelob.replace(".", ",")
                    document.getElementById("belobdiv_" + x + "_1").innerHTML = "<b>" + document.getElementById("FM_beloebthis_" + x + "_1").value + " " + valutaKodeFak + "</b>"

                } 
					
	}
	
	
	
	
	
	
	
	//////////////////////////////////////////////////////////////
	/// Main funktion 
	//////////////////////////////////////////////////////////////
    function enhedsprismedarb(x, n) {

        //alert("starter")

	    var kCode = 0;

	    if (x != 0) {
	        //kCode = window.event.keyCode
	        var e
	        if (window.event) kCode = window.event.keyCode; // IE
	        else if (e) kCode = e.which;
	    } else {
	        kCode = 0
	    }


	    if (kCode == 37) {
	    } else {
	        if (kCode == 39) {
	        } else {
				
				
				
				
				///// Finder værdier og variable ////
				
				//Finder a og n værdier
				antal_n_x = parseInt(document.getElementById("antal_n_"+x+"").value)
				highest_aval = parseInt((document.getElementById("highest_aval_"+x+"").value/1) + 1)
				antal_n_x2 = parseInt(document.getElementById("antal_n_"+x+"").value/1 - 1)
				
				
				
				//Valuta ID FakturA 
				var valutaKodeFak = document.getElementById("FM_valuta_all_1").options[document.getElementById("FM_valuta_all_1").selectedIndex].text
	            var valutaIdFak = document.getElementById("FM_valuta_all_1").value
	            var valutaKursFak = document.getElementById("valutakurs_"+valutaIdFak).value
				
				//Valuta ID Medarb. linie
				
				
				valutaKode = document.getElementById("FM_mvaluta_" + n + "_" + x +"").options[document.getElementById("FM_mvaluta_" + n + "_" + x +"").selectedIndex].text
			    valutaKodeOpr = document.getElementById("FM_mvalutaKode_" + n + "_" + x +"").value
				valutaId = document.getElementById("FM_mvaluta_" + n + "_" + x +"").value 				
				valutaIdOpr = document.getElementById("FM_mvaluta_opr_" + n + "_" + x +"").value 			
				document.getElementById("FM_mvalutaKode_" + n + "_" + x +"").value = valutaKode
				valutaKurs = document.getElementById("valutakurs_"+valutaId+"").value.replace(",",".")
                valutaKursOpr = document.getElementById("valutakurs_"+valutaIdOpr+"").value.replace(",",".")
                
				
				//Finder medarbejder timepriser
				var this_mtp = 0; 
				this_mtp = document.getElementById("FM_mtimepris_" + n + "_" + x +"").value.replace(",",".")
			    this_mtp_compare = this_mtp
				
			    
			    
			    // Beregn timepris på medar. linie ved Valuta ændring //
			    if (valutaKode != valutaKodeOpr) {
			    //alert(valutaKode +"!="+ valutaKodeOpr)
			    
			    var this_mtp_valutakursberegn = 0;
			    this_mtp_valutakursberegn = this_mtp/1 * (valutaKurs.replace(",",".")/1 / valutaKursOpr.replace(",",".")/1)
			    this_mtp = Math.round(this_mtp_valutakursberegn*100)/ 100
			    this_mtp = String(this_mtp)
			    }
				
				//this_mtp_opr = document.getElementById("FM_mtimepris_opr_" + n + "_" + x +"").value
				this_mtp_opr_compare = String(document.getElementById("FM_mtimepris_opr_" + n + "_" + x +"").value.replace(",","."));
				
				
				// Antal timer / enheder
				this_mtimer = document.getElementById("FM_m_fak_" + n + "_" + x +"").value.replace(",",".")
				
				
				//Sætter den nye timepris på medarbejderen efter replace
				document.getElementById("FM_mtimepris_" + n + "_" + x +"").value = this_mtp.replace(".",",")
				
				
				//Rabat
				rabat = document.getElementById("FM_mrabat_"+n+"_"+x+"").value
				rabatopr = document.getElementById("FM_mrabat_opr_"+n+"_"+x+"").value
				
			    // Enheds angivelse //
			    enhedsang = document.getElementById("FM_med_enh_" + n + "_" + x +"").value
			    enhedsangopr = document.getElementById("FM_med_enh_opr_" + n + "_" + x +"").value
			    ehLabel = document.getElementById("FM_med_enh_" + n + "_" + x +"").options[document.getElementById("FM_med_enh_" + n + "_" + x +"").selectedIndex].text
			    
				// Beregn beløb på medarb ved Valuta ændring i forhold til valuta på FAK //
			    mBelobValBeregn(valutaKode, valutaKodeFak, valutaKurs, valutaKursFak, this_mtp, this_mtimer, rabat, n, x)
				
				
			    var nyBeloebtot = 0;
				var tmpbel = 0;
				var antalmedarbtimertot = 0;
				var tmptim = 0;
				var isSumAktWrt = "##,";
				
				
				//Samler det totale antal timer og beløb på den enhedspris der er valgt
				//Gennemløber medarbejder linier
				for (i=1;i<antal_n_x;i++){
						
						//alert(parseFloat(this_mtp) +" == "+ parseFloat(document.getElementById("FM_mtimepris_" + i + "_" + x +"").value.replace(",",".")))
						if (i != 0) {
						
						    
							   // +"&&"+ rabat +"=="+ document.getElementById("FM_mrabat_"+i+"_"+x+"").value +"&&"+ valutaId +"=="+ document.getElementById("FM_mvaluta_"+i+"_"+x+"").value +"&&"+ parseFloat(this_mtp) +"=="+ parseFloat(document.getElementById("FM_mtimepris_" + i + "_" + x +"").value.replace(",","."))))
				                if (enhedsang == document.getElementById("FM_med_enh_"+i+"_"+x+"").value && rabat == document.getElementById("FM_mrabat_"+i+"_"+x+"").value && valutaId == document.getElementById("FM_mvaluta_"+i+"_"+x+"").value && parseFloat(this_mtp) == parseFloat(document.getElementById("FM_mtimepris_" + i + "_" + x +"").value.replace(",","."))) {
						        
						        //alert("Den valgte x/i: "+x+"/"+i+" Enh: "+ enhedsang +"=="+ document.getElementById("FM_med_enh_"+i+"_"+x+"").value +" Rabat: "+ rabat +"=="+ document.getElementById("FM_mrabat_"+i+"_"+x+"").value +" Valuta: "+ valutaId +"=="+ document.getElementById("FM_mvaluta_"+i+"_"+x+"").value +" tpris: "+ parseFloat(this_mtp) +"=="+ document.getElementById("FM_mtimepris_" + i + "_" + x +"").value)
							   
						        
						        tmpbel = document.getElementById("FM_mbelob_" + i + "_" + x +"").value.replace(",",".")/1
						        nyBeloebtot = nyBeloebtot + tmpbel
						        tmptim = document.getElementById("FM_m_fak_" + i + "_" + x +"").value.replace(",",".")/1 
						        antalmedarbtimertot = antalmedarbtimertot + tmptim
        						
						        } // timepris
						    
						} // i
				} // for
					
					
					
					
					
					ifHighestA = -1000
					foundone = 0
					var oprSumAktTpris = 0;
					//var oprSumAktValutaId = 0;
					
					////////////////////////////////////////////////////////////////////////
					// Opdaterer den valgte sum-aktivitet (Den der passer til timeprisen)
					// Og til valuta
					// highest_aval
					////////////////////////////////////////////////////////////////////////
					
					for (i=-(antal_n_x2);i<highest_aval;i++){
					//alert(isSumAktWrt) 
					        
					        oprSumAktTpris = parseFloat(document.getElementById("FM_enhedspris_"+x+"_"+i+"").value.replace(",","."))
					        
							// Er akt. med samme kombination af timepris og valuta allerede skrevet ///
							if (isSumAktWrt.indexOf("#"+ rabat +"_"+ enhedsang +"_"+ valutaId +"_"+ oprSumAktTpris +"#,") == -1) {
        				    
							if (i != 0) {
							
							// Sum akt //
							
							//alert("her"+ enhedsangopr +"=="+ document.getElementById("FM_akt_enh_" +x+"_"+i+"").value)
							if (enhedsang == document.getElementById("FM_akt_enh_" +x+"_"+i+"").value && rabat == document.getElementById("FM_rabat_"+x+"_"+i+"").value && valutaId == document.getElementById("FM_valuta_"+x+"_"+i+"").value && parseFloat(this_mtp) == oprSumAktTpris) {
							//alert("Aktuel AKT - timer: "+ antalmedarbtimertot +" x/i: "+x+"/"+i+" Enh: "+ enhedsang +"=="+ document.getElementById("FM_akt_enh_"+x+"_"+i+"").value +" Rabat: "+ rabat +"=="+ document.getElementById("FM_rabat_"+x+"_"+i+"").value +" Valuta: "+ valutaId +"=="+ document.getElementById("FM_valuta_"+x+"_"+i+"").value +" tpris: "+ parseFloat(this_mtp) +"=="+ oprSumAktTpris)
							   
							
							       
        							    if (foundone == 0){
								        foundone = 1
								        // ifHighestA bruges hvis der er flere medarbejdere med samme 
								        // timepris og sidste medarbejder's a værdi er højere end den
								        // sidste sum-aktivitets.
								        ifHighestA = i
								        }
        							
        							
        							
							        //Totaler på den valgte sum-aktivitet
							        document.getElementById("FM_timerthis_"+x+"_"+i+"").value = antalmedarbtimertot
							        document.getElementById("FM_timerthis_"+x+"_"+i+"").value = document.getElementById("FM_timerthis_"+x+"_"+i+"").value.replace(".",",") //document.getElementById("FM_m_fak_"+n+"_"+x+"").value
							        document.getElementById("FM_beloebthis_"+x+"_"+i+"").value = nyBeloebtot 
							        document.getElementById("FM_beloebthis_"+x+"_"+i+"").value = document.getElementById("FM_beloebthis_"+x+"_"+i+"").value.replace(".",",")
        							
        							// Enhedspris
        							document.getElementById("FM_enhedspris_"+x+"_"+i+"").value = this_mtp.replace(".",",")
        							document.getElementById("enhprisdiv_" +x+"_"+i+"").innerHTML = "<b>"+document.getElementById("FM_enhedspris_"+x+"_"+i+"").value+"</b>"
								    
        							
        							
							        // Rabat
							        document.getElementById("FM_rabat_"+x+"_"+i+"").value = rabat
        							document.getElementById("rabatdiv_"+x+"_"+i+"").innerHTML = (100 * document.getElementById("FM_rabat_"+x+"_"+i+"").value)+" %"
        							
        							// Antal
        							document.getElementById("timeprisdiv_" +x+"_"+i+"").innerHTML = "<b>"+document.getElementById("FM_timerthis_"+x+"_"+i+"").value+"</b>"
								    
        							
        							// Beløb
							        document.getElementById("belobdiv_"+x+"_"+i+"").innerHTML = "<b>"+document.getElementById("FM_beloebthis_"+x+"_"+i+"").value+" "+ valutaKodeFak +"</b>"	
								    
								    // Enhedsang
							        document.getElementById("FM_akt_enh_"+x+"_"+i).value = enhedsang
							        document.getElementById("ehdiv_"+x+"_"+i+"").innerHTML = ehLabel
            						    
								        
							        // Valuta 
							        document.getElementById("FM_valuta_"+x+"_"+i).value = valutaId
							        document.getElementById("valutadiv_"+x+"_"+i+"").innerHTML = "<b>"+valutaKode+"</b>"
            							
        							
							        // Flueben til / fra
							        document.getElementById("FM_show_akt_" + x + "_" + i + "").checked = true
							       
							        // Opdaterer isSumAktWrt
							        isSumAktWrt = isSumAktWrt +"#"+ rabat +"_"+ enhedsang +"_"+ valutaId +"_"+ oprSumAktTpris +"#,"
							        //alert(isSumAktWrt)
							        
							        
						       } // SLUT if timepris og valutaId 
							
						 } //i = 0	
						 
						  } else {  
						           
						            //alert("ja tak")
						            // Antal //
						            document.getElementById("FM_timerthis_"+x+"_"+i+"").value = 0
						            
						            document.getElementById("timeprisdiv_" +x+"_"+i+"").innerHTML = "<b>"+document.getElementById("FM_timerthis_"+x+"_"+i+"").value+"</b>"
								     
								    document.getElementById("belobdiv_"+x+"_"+i+"").innerHTML = "<b>"+document.getElementById("FM_beloebthis_"+x+"_"+i+"").value+"</b>"	
								    // Rabat
							        document.getElementById("rabatdiv_"+x+"_"+i+"").innerHTML = (100 * document.getElementById("FM_rabat_"+x+"_"+i+"").value)+" %"
        							
						           
        							
						            // Flueben til / fra
						            document.getElementById("FM_show_akt_"+x+"_"+i+"").checked = false
						            } // SLUT iswrt
						 
				    } //for
							
				
				
				
					
				var nyBeloebtot2 = 0;
				var tmpbel2 = 0;
				var antalmedarbtimertot2 = 0;
				var tmptim2 = 0;
				var showbesk = "";
				var showbesk1 = "";
				
				//alert("her "+ this_mtp_compare + " - " + this_mtp_opr_compare)
				
				
				
				
				/////////////////////////////////////////////////////////////////
				// Genberegner summen/antallet af timer 
				// på den oprindelige. medarbejdertimepris
				////////////////////////////////////////////////////////////////
				
				
				// Hvis der er skiftet timepris eller valuta  ///
				//if (parseFloat(this_mtp_compare) != parseFloat(this_mtp_opr_compare || valutaIdOpr != valutaId)) {
						
						
						
						for (i=1;i<antal_n_x;i++){
							//alert(parseFloat(this_mtp_opr_compare) +"=="+ parseFloat(document.getElementById("FM_mtimepris_" + i + "_" + x +"").value.replace(",",".")))
							if (i != 0) {
							        
							        //alert(valutaIdOpr +"=="+ document.getElementById("FM_mvaluta_"+i+"_"+x+"").value +"&&"+ parseFloat(this_mtp_opr_compare) +"=="+ parseFloat(document.getElementById("FM_mtimepris_"+i+"_"+x+"").value.replace(",","."))
							       
							        // Hvis valutaID og Timepris passer til den oprindelige //
							       
							        if (enhedsangopr == document.getElementById("FM_med_enh_" +i+"_"+x+"").value && rabatopr == document.getElementById("FM_mrabat_"+i+"_"+x+"").value && valutaIdOpr == document.getElementById("FM_mvaluta_"+i+"_"+x+"").value && parseFloat(this_mtp_opr_compare) == parseFloat(document.getElementById("FM_mtimepris_"+i+"_"+x+"").value.replace(",","."))) {
							        //alert("Oprindelig: x/i: "+x+"/"+i+" Enh: "+ enhedsangopr +"=="+ document.getElementById("FM_med_enh_"+i+"_"+x+"").value +" Rabat: "+ rabatopr +"=="+ document.getElementById("FM_mrabat_"+i+"_"+x+"").value +" Valuta: "+ valutaIdOpr +"=="+ document.getElementById("FM_mvaluta_"+i+"_"+x+"").value +" tpris: "+ parseFloat(this_mtp_opr_compare) +"=="+ document.getElementById("FM_mtimepris_" + i + "_" + x +"").value)
							   
							        
							        tmpbel2 = document.getElementById("FM_mbelob_" + i + "_" + x +"").value.replace(",",".")
							        nyBeloebtot2 = nyBeloebtot2 + parseFloat(tmpbel2)
        							
						            tmptim2 = document.getElementById("FM_m_fak_" + i + "_" + x +"").value.replace(",",".") 
							        antalmedarbtimertot2 = antalmedarbtimertot2 + parseFloat(tmptim2)
							        
							        } // timepris & Valuta
							        
							   
						    } // i
					    } // for
						
						
						
						
						//Opdaterer sum-aktiviteten på den oprindelige timepris (den forladte)
						for (i=-(antal_n_x2);i<highest_aval;i++){
							   
							    if (i != 0) {
							        
							            
							            if (enhedsangopr == document.getElementById("FM_akt_enh_"+x+"_"+i+"").value && rabatopr == document.getElementById("FM_rabat_"+x+"_"+i+"").value && valutaIdOpr == document.getElementById("FM_valuta_"+x+"_"+i+"").value && parseFloat(this_mtp_opr_compare) == parseFloat(document.getElementById("FM_enhedspris_"+x+"_"+i+"").value.replace(",","."))) {
							            
							            //alert("Oprindelig AKT - timer: "+ antalmedarbtimertot2 +" x/i: "+x+"/"+i+" Enh: "+ enhedsangopr +"=="+ document.getElementById("FM_akt_enh_"+x+"_"+i+"").value +" Rabat: "+ rabatopr +"=="+ document.getElementById("FM_rabat_"+x+"_"+i+"").value +" Valuta: "+ valutaIdOpr +"=="+ document.getElementById("FM_valuta_"+x+"_"+i+"").value +" tpris: "+ parseFloat(this_mtp_opr_compare) +"=="+ document.getElementById("FM_enhedspris_"+x+"_"+i+"").value)
							   
							            
							            //Totaler på forladte sum-aktivitet
							            document.getElementById("FM_timerthis_"+x+"_"+i+"").value = antalmedarbtimertot2 //document.getElementById("FM_m_fak_"+n+"_"+x+"").value
							            document.getElementById("FM_beloebthis_"+x+"_"+i+"").value = nyBeloebtot2
							            
							            //Rabat
							            //alert(rabat)
							            //document.getElementById("FM_rabat_"+x+"_"+i+"").value = rabat
            							//document.getElementById("rabatdiv_"+x+"_"+i+"").innerHTML = "<b>"+(100 * document.getElementById("FM_rabat_"+x+"_"+i+"").value)+" %</b>"
            							
							            
							            // Timer 
						                var aktTimer = antalmedarbtimertot2 
						                //antalmedarbtimertot2 //document.getElementById("FM_timerthis_"+x+"_"+i+"").value
			                            //alert(antalmedarbtimertot)
			                            aktTimer = Math.round(aktTimer*100)/ 100
			                            if (aktTimer != 0) {
			                            aktTimer = String(aktTimer)
			                            } else {
			                            aktTimer = "0.00"
			                            }
            							
            							document.getElementById("FM_timerthis_"+x+"_"+i+"").value = aktTimer.replace(".",",")
							            
							            document.getElementById("timeprisdiv_" +x+"_"+i+"").innerHTML = "<b>"+document.getElementById("FM_timerthis_"+x+"_"+i+"").value+"</b>"
							            
							            
							            // Beløb
							            var aktBelob = document.getElementById("FM_beloebthis_"+x+"_"+i+"").value
			                            aktBelob = Math.round(aktBelob*100)/ 100
			                            
			                            if (aktBelob != 0) {
			                            aktBelob = String(aktBelob)
			                            } else {
			                            aktBelob = "0.00"
			                            }
            							
            							document.getElementById("FM_beloebthis_"+x+"_"+i+"").value = aktBelob.replace(".",",")
							            document.getElementById("belobdiv_"+x+"_"+i+"").innerHTML = "<b>"+document.getElementById("FM_beloebthis_"+x+"_"+i+"").value+" "+ valutaKodeFak +"</b>"
							            
							            
							            // Enhedsang
							            //document.getElementById("FM_akt_enh_"+x+"_"+i).value = enhedsangopr
							            //document.getElementById("ehdiv_"+x+"_"+i+"").innerHTML = "<b>"+ehLabel+"</b>"
            						
							           
							           
            							
            							//Valuta 
							            document.getElementById("FM_valuta_"+x+"_"+i).value = valutaIdOpr
							            document.getElementById("valutadiv_"+x+"_"+i+"").innerHTML = "<b>"+ valutaKodeOpr +"</b>"
            							
            							
            							
            							    if (antalmedarbtimertot2 == 0) {
								                document.getElementById("FM_show_akt_" + x + "_" + i +"").checked = false
								            } else {
								                document.getElementById("FM_show_akt_" + x + "_" + i +"").checked = true
								            }
								            
								            
							            } // if timepris
							            
							       
							        } // if	
							
						} // for
					
					
					
					
			    //} // IF genBeregn hvis timepris = oprtimepris
			    
			    
			    /// SLUT Genberegner sum akt. på den akt. der 
			    /// er blevet forladt ved. valuta eller timepris skift 
			    //////// Slut opdater timepris på oprindelig akt ////////
    			
					
					
					
					
					
					
					                //Viser sumaktivitet div. (bruges ved redigering af fak.)
							        if (parseInt(foundone) == 1) {
								        //alert("her")
								        //er n større end det højest brugte A på denne sum-aktivitet?
								        if (n > ifHighestA) {
								        useThisN = ifHighestA
								        } else {
								        useThisN = n
								        }
        							
								        document.getElementById("sumaktdiv_"+x+"_"+useThisN+"").style.display = "";
								        document.getElementById("sumaktdiv_" + x + "_" + useThisN + "").style.visibility = "visible";

								        
        								
							        }
							        ////// end if ////
							        
							        
							        
							
							
					            ////////////////////////////////////////////////////////////////////////
			                    //Hvis der ikke findes en åben sumaktivitet med den valgte timepris
			                    ////////////////////////////////////////////////////////////////////////
			                    
			                    //alert("foundone: "+foundone)
			                    
			                    if (parseInt(foundone) == 0) {
			                    
			                    //alert("åbner ny div")
			                    
			                    //Åbner ny div
			                    document.getElementById("sumaktdiv_"+x+"_"+-(n)+"").style.display = "";
			                    document.getElementById("sumaktdiv_"+x+"_"+-(n)+"").style.visibility = "visible";
            					
            					 
			                    //Sætter værdier i de nye div
			                    document.getElementById("FM_enhedspris_" + x + "_"+ -(n) +"").value = document.getElementById("FM_mtimepris_" + n + "_" + x +"").value.replace(".",",")
			                    document.getElementById("enhprisdiv_" + x + "_"+ -(n) +"").innerHTML = "<b>"+ document.getElementById("FM_enhedspris_" + x + "_"+ -(n) +"").value +"</b>"
            					
			                        //alert for samme pris som 0 akt.
			                        if (parseFloat(document.getElementById("FM_enhedspris_" + x + "_"+ -(n) +"").value) == parseFloat(document.getElementById("FM_enhedspris_" + x + "_0").value)) {
			                        alert ("Du har angivet samme timepris som på Sum-aktivitets samle linien, som derfor nulstilles.\nDet kan være nødvendigt at klikke på 'Calc' knappen endnu engang for at opdatere timer og beløb.")
			                        document.getElementById("FM_enhedspris_" + x + "_0").value = -2
			                        document.getElementById("FM_timerthis_" + x + "_0").value = 0
			                        document.getElementById("FM_beloebthis_" + x + "_0").value = 0
			                        document.getElementById("FM_show_akt_"+x+"_0").checked = false
			                        }
            					
            					
			                            // Timer 
						                var mTimer = this_mtimer //document.getElementById("FM_timerthis_"+ x +"_"+ -(n) +"").value
						                mTimer = Math.round(mTimer*100)/ 100
			                            if (mTimer != 0) {
			                            mTimer = String(mTimer)
			                            } else {
			                            mTimer = "0.00"
			                            }
            							
            							document.getElementById("FM_timerthis_"+ x +"_"+ -(n) +"").value = mTimer.replace(".",",")
							            document.getElementById("timeprisdiv_" + x +"_"+ -(n) +"").innerHTML = "<b>"+ document.getElementById("FM_timerthis_"+ x +"_"+ -(n) +"").value +"</b>"
            						    
            					       
            					        // Rabat beregning
			                            rabat = document.getElementById("FM_mrabat_" + n + "_" + x +"").value
                                        document.getElementById("FM_beloebthis_" + x + "_"+ -(n) +"").value = (this_mtp * this_mtimer) - (this_mtp * this_mtimer * rabat)
                                       
			                            document.getElementById("FM_rabat_"+x+"_"+-(n)+"").value = rabat
            					        document.getElementById("rabatdiv_" + x + "_"+ -(n) +"").innerHTML = (100 * document.getElementById("FM_rabat_" + x + "_"+ -(n) +"").value) +" %"
                						
            		                
			                       
            						    // Beløb
							            var mBelob = document.getElementById("FM_beloebthis_"+ x +"_"+ -(n) +"").value
			                            mBelob = Math.round(mBelob*100)/ 100
			                            
			                            if (mBelob != 0) {
			                            mBelob = String(mBelob)
			                            } else {
			                            mBelob = "0.00"
			                            }
            							
            							document.getElementById("FM_beloebthis_"+x+"_"+ -(n) +"").value = mBelob.replace(".",",")
							            document.getElementById("belobdiv_" + x + "_"+ -(n) +"").innerHTML = "<b>"+ document.getElementById("FM_beloebthis_" + x + "_"+ -(n) +"").value +" "+ valutaKodeFak +"</b>"
            						
							            
							            // Enhedsang
							            document.getElementById("FM_akt_enh_"+x+"_"+ -(n)).value = enhedsang
							            document.getElementById("ehdiv_"+x+"_"+ -(n)).innerHTML = ehLabel
            						
							            
							            
							            // Valuta 
							            document.getElementById("FM_valuta_"+x+"_"+ -(n)).value = valutaId
							            document.getElementById("valutadiv_"+x+"_"+ -(n)+"").innerHTML = "<b>"+valutaKode+"</b>"
            						
            						
			                       
			                        
			                        document.getElementById("FM_show_akt_" + x + "_" + -(n) +"").checked = true
            						
			                    //Tekst (Fjernet, da det er irriterende at den tekst man har skrevet bilver slettet med tryk på "calc")
			                    //showbesk = document.getElementById("sumaktbesk_"+x+"").value
			                    //document.getElementById("FM_aktbesk_" + x + "_"+ -(n) +"").value = showbesk
			                    } 
							    ///////////////////////////////////////////////
							
							
							
							
							// Opdaterer oprindelig (hiddden) valuta på medarbejder linje
				            document.getElementById("FM_mvaluta_opr_" + n + "_" + x +"").value = valutaId
						
							//Sætter timepris opr
							document.getElementById("FM_mtimepris_opr_" + n + "_" + x +"").value = this_mtp.replace(".",",")
							
							//Opdaterer hidden Rabat 
							document.getElementById("FM_mrabat_opr_" + n + "_" + x +"").value = rabat
							
							//Enhedaangivelse 
							document.getElementById("FM_med_enh_opr_" + n + "_" + x +"").value = enhedsang
							
							//Afrunder
							offSet_this = String(document.getElementById("FM_mtimepris_opr_" + n + "_" + x +"").value);
							offSetL_this = String(document.getElementById("FM_mtimepris_opr_" + n + "_" + x +"").length);
							pkt_this = offSet_this.indexOf(",");
							if (pkt_this > 0){
								document.getElementById("FM_mtimepris_opr_" + n + "_" + x +"").value = offSet_this.substring(0, pkt_this + 3)
							}
							/////////////////////////////////////////////////



							//setTimerTot2(x)

				//*** Tjekker om sum-samle akt har samme timepris///

                //alert("her")
				        andreEnhprs(x,0)
                        setBeloebTot2(x,1)	
				
				//opdatervalutAllelinier(1,x)			
				
				}
	    }

        
	}
	
	
	function offsetventer(x,n){
	    var kCode = 0;

	    if (x != 0) {
	        //kCode = window.event.keyCode
	        var e
	        if (window.event) kCode = window.event.keyCode; // IE
	        else if (e) kCode = e.which;
	    } else {
	        kCode = 0
	    }


	    if (kCode == 37) {
	    } else {
	        if (kCode == 39) {
	        } else {
				offSet_this = String(document.getElementById("FM_m_vent_" + n + "_" + x +"").value.replace(".",","));
				offSetL_this = String(document.getElementById("FM_m_vent_" + n + "_" + x +"").length);
				pkt_this = offSet_this.indexOf(",");
				if (pkt_this > 0){
					document.getElementById("FM_m_vent_" + n + "_" + x +"").value = offSet_this.substring(0, pkt_this + 3)
				}
			}
		}
	}
	
	
	function offsetfak(x,n){
	    var kCode = 0;

	    if (x != 0) {
	        //kCode = window.event.keyCode
	        var e
	        if (window.event) kCode = window.event.keyCode; // IE
	        else if (e) kCode = e.which;
	    } else {
	        kCode = 0
	    }


	    if (kCode == 37) {
	    } else {
	        if (kCode == 39) {
	        } else {
					offSet_this = String(document.getElementById("FM_m_fak_" + n + "_" + x +"").value.replace(".",","));
					offSetL_this = String(document.getElementById("FM_m_fak_" + n + "_" + x +"").length);
					pkt_this = offSet_this.indexOf(",");
					if (pkt_this > 0){
						document.getElementById("FM_m_fak_" + n + "_" + x +"").value = offSet_this.substring(0, pkt_this + 3)
					}
				}
		}
	}			
	
	function offsetmtp(x,n){
	    var kCode = 0;

	    if (x != 0) {
	        //kCode = window.event.keyCode
	        var e
	        if (window.event) kCode = window.event.keyCode; // IE
	        else if (e) kCode = e.which;
	    } else {
	        kCode = 0
	    }


	    if (kCode == 37) {
	    } else {
	        if (kCode == 39) {
	        } else {
				offSet_this = String(document.getElementById("FM_mtimepris_" + n + "_" + x +"").value.replace(".",","));
				offSetL_this = String(document.getElementById("FM_mtimepris_" + n + "_" + x +"").length);
				pkt_this = offSet_this.indexOf(",");
				if (pkt_this > 0){
					document.getElementById("FM_mtimepris_" + n + "_" + x +"").value = offSet_this.substring(0, pkt_this + 3)
				}
			}
		}
	}
	
	
	
	
	
	function isNum(passedVal){
	invalidChars = " /:;+<>abcdefghijklmnopqrstuvwxyzæøå"
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

                if (passedVal.charAt(i) == "-") {
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
	
	function validZip(inZip){
	if (inZip == "") {
	return true
	}
	if (isNum(inZip)){
	return true
	}
	return false
	}
	
	
	function tjektimer(x, a) {

	
	var oprTimerThisTj = 0;
	oprTimerThisTj = (document.getElementById("FM_hidden_timerthis_"+x+"_"+a+"").value.replace(",",".") / 1);
		
	
		
		if (!validZip(document.getElementById("FM_timerthis_"+x+"_"+a+"").value)){
		alert("Der er angivet et ugyldigt tegn.")
		document.getElementById("FM_timerthis_"+x+"_"+a+"").value = oprTimerThisTj;
		document.getElementById("FM_timerthis_"+x+"_"+a+"").focus()
		document.getElementById("FM_timerthis_"+x+"_"+a+"").select()
		return false
		}
	
        
		
		return true

	}
	
	function tjektimer2(x,n){
	    var kCode = 0;

	    if (x != 0) {
	        //kCode = window.event.keyCode
	        var e
	        if (window.event) kCode = window.event.keyCode; // IE
	        else if (e) kCode = e.which;
	    } else {
	        kCode = 0
	    }


	    if (kCode == 37) {
	    } else {
	        if (kCode == 39) {
	        } else {
				var oprTimerThisTj = 0;
				oprTimerThisTj = (document.getElementById("FM_m_fak_opr_"+n+"_"+x+"").value.replace(",",".") / 1);
					if (!validZip(document.getElementById("FM_m_fak_"+n+"_"+x+"").value)){
					alert("Der er angivet et ugyldigt tegn.")
					document.getElementById("FM_m_fak_"+n+"_"+x+"").value = oprTimerThisTj;
					document.getElementById("FM_m_fak_"+n+"_"+x+"").focus()
					document.getElementById("FM_m_fak_"+n+"_"+x+"").select()
					return false
					}
				return true
				}
			}
	}
	
	function tjektimer4(x,n){
	    var kCode = 0;

	    if (x != 0) {
	        //kCode = window.event.keyCode
	        var e
	        if (window.event) kCode = window.event.keyCode; // IE
	        else if (e) kCode = e.which;
	    } else {
	        kCode = 0
	    }


	    if (kCode == 37) {
	    } else {
	        if (kCode == 39) {
	        } else {
				var oprTimerThisTj = 0;
				oprTimerThisTj = (document.getElementById("FM_m_vent_opr_"+n+"_"+x+"").value.replace(",",".") / 1);
					if (!validZip(document.getElementById("FM_m_vent_"+n+"_"+x+"").value)){
					alert("Der er angivet et ugyldigt tegn.")
					document.getElementById("FM_m_vent_"+n+"_"+x+"").value = oprTimerThisTj;
					document.getElementById("FM_m_vent_"+n+"_"+x+"").focus()
					document.getElementById("FM_m_vent_"+n+"_"+x+"").select()
					return false
					}
				return true
				}
			}
	}
	
	function tjektimepris(x){
	    var kCode = 0;

	    if (x != 0) {
	        //kCode = window.event.keyCode
	        var e
	        if (window.event) kCode = window.event.keyCode; // IE
	        else if (e) kCode = e.which;
	    } else {
	        kCode = 0
	    }


	    if (kCode == 37) {
	    } else {
	        if (kCode == 39) {
	        } else {
				    

                    as = document.getElementById("antal_subtotal_akt_"+x+"").value
				    /// Aftale eller job
                    if (as == -1) { //aftale
                        a = 1
                    } else {
                        a = 0
                    }

				    var oprTimerThisTj = 0;
				    oprTimerThisTj = (document.getElementById("FM_hidden_timepristhis_" + x + "_" + a + "").value.replace(",", ".") / 1);
				    if (!validZip(document.getElementById("FM_enhedspris_" + x + "_" + a + "").value)) {
					alert("Der er angivet et ugyldigt tegn.")
					document.getElementById("FM_enhedspris_" + x + "_" + a + "").value = oprTimerThisTj;
					document.getElementById("FM_enhedspris_" + x + "_" + a + "").focus()
					document.getElementById("FM_enhedspris_" + x + "_" + a + "").select()
					return false
					}
				return true
				}
			}
	}





	function tjektimer3(x,n){
	    var kCode = 0;

	    if (x != 0) {
	        //kCode = window.event.keyCode
	        var e
	        if (window.event) kCode = window.event.keyCode; // IE
	        else if (e) kCode = e.which;
	    } else {
	        kCode = 0
	    }


	    if (kCode == 37) {
	    } else {
	        if (kCode == 39) {
	        } else {
	                    var oprTimerThisTj = 0;
	                    oprTimerThisTj = (document.getElementById("FM_mtimepris_opr_"+n+"_"+x+"").value.replace(",",".") / 1);
		                    if (!validZip(document.getElementById("FM_mtimepris_"+n+"_"+x+"").value)){
		                    alert("Der er angivet et ugyldigt tegn.")
		                    document.getElementById("FM_mtimepris_"+n+"_"+x+"").value = oprTimerThisTj;
		                    document.getElementById("FM_mtimepris_"+n+"_"+x+"").focus()
		                    document.getElementById("FM_mtimepris_"+n+"_"+x+"").select()
		                    return false
		                    }
	                    return true
	                    }
			}
	}
	
	
	
	function tjekBeloeb(x,a){
		if (!validZip(document.getElementById("FM_beloebthis_"+x+"_"+a+"").value)){
		alert("Der er angivet et ugyldigt tegn.")
		setBeloebTot2(x,1)
		return false
		}
	return true
	}
	
	
	function showborder(x, n) {

      
	document.getElementById("beregn_"+n+"_"+x+"_a").style.borderColor = "red"
	document.getElementById("FM_mvaluta_"+n+"_"+x).disabled = true
	document.getElementById("FM_med_enh_"+n+"_"+x).disabled = true
	document.getElementById("FM_mrabat_"+n+"_"+x).disabled = true
	}
	
	function hideborder(x, n) {

	   
	document.getElementById("beregn_"+n+"_"+x+"_a").style.borderColor = ""
	document.getElementById("FM_mvaluta_"+n+"_"+x).disabled = false
	document.getElementById("FM_mtimepris_"+n+"_"+x+"").disabled = false
	document.getElementById("FM_med_enh_"+n+"_"+x).disabled = false
	document.getElementById("FM_mrabat_"+n+"_"+x).disabled = false
	}
	
	function showborder2(x,a) {
	//Faktura valuta 
    var valutaKodeFak = document.getElementById("FM_valuta_all_1").options[document.getElementById("FM_valuta_all_1").selectedIndex].text
	            
	document.getElementById("beregn2_"+x+"_"+a+"_a").style.borderColor = "red"
	document.getElementById("FM_valuta_"+x+"_"+a).disabled = true
	document.getElementById("FM_akt_enh_"+x+"_"+a).disabled = true
	document.getElementById("FM_rabat_"+x+"_"+a).disabled = true
	
	document.getElementById("belobdiv_"+x+"_"+a+"").innerHTML = "<b>"+document.getElementById("FM_beloebthis_"+x+"_"+a+"").value+" "+ valutaKodeFak +"</b>" 
   
    }
	
	function hideborder2(x,a) {
	document.getElementById("beregn2_"+x+"_"+a+"_a").style.borderColor = ""
	document.getElementById("FM_valuta_"+x+"_"+a).disabled = false
	document.getElementById("FM_akt_enh_"+x+"_"+a).disabled = false
	document.getElementById("FM_rabat_"+x+"_"+a).disabled = false
	}
	
	
	/// Opret fakturaer på en aftale
	function aft_stk_pris() {

  

	//var bel = 0;
	//bel = document.getElementById("FM_beloeb").value.replace(",",".");
	var tim = 0;
	tim = document.getElementById("FM_timer").value.replace(",",".");
	var rabat = 0;
	rbt = document.getElementById("FM_rabat").value
	var stkpris = 0;
	stkpris = document.getElementById("FM_stkpris").value.replace(",",".")
	
	var belob = 0;
	belob = parseFloat(tim * stkpris - (tim * stkpris * rbt))

    //alert(belob)
	
	document.getElementById("FM_beloeb").value = belob
	document.getElementById("FM_beloeb").value = document.getElementById("FM_beloeb").value.replace(".",",")
	
	//document.getElementById("FM_stkpris").value = tmpstkprisval
    //document.getElementById("FM_stkpris").value = document.getElementById("FM_stkpris").value.replace(".",",");
	
	offSet_this = String(document.getElementById("FM_beloeb").value);
	offSetL_this = String(document.getElementById("FM_beloeb").length);
	pkt_this = offSet_this.indexOf(",");
	if (pkt_this > 0){
		document.getElementById("FM_beloeb").value = offSet_this.substring(0, pkt_this + 3)
	}
	
	}
	
	
	
	function opdaterFakdato(){
	    document.getElementById("gen_FM_start_dag").value = document.getElementById("FM_istdato_dag").value
	    document.getElementById("gen_FM_start_mrd").value = document.getElementById("FM_istdato_mrd").value
	    document.getElementById("gen_FM_start_aar").value = document.getElementById("FM_istdato_aar").value
	    document.getElementById("gen_FM_slut_dag").value = document.getElementById("FM_istdato2_dag").value
	    document.getElementById("gen_FM_slut_mrd").value = document.getElementById("FM_istdato2_mrd").value
	    document.getElementById("gen_FM_slut_aar").value = document.getElementById("FM_istdato2_aar").value
	
	}
	
	
	
	function addbr(x,a) {
	
	//var selText=document.getElementById("FM_aktbesk_"+x+"_"+a).selection.createRange();
	//alert(selText)
	document.getElementById("FM_aktbesk_"+x+"_"+a).value = document.getElementById("FM_aktbesk_"+x+"_"+a).value + "<br>"
	}
	
	
	
	/////////////////////////////////////////////////////////
	// Materiale linier //
	/////////////////////////////////////////////////////////

	function beregnmatpris(m, beregntot) {
	antal = document.getElementById("FM_matantal_" + m).value.replace(".", "")
	antal = antal.replace(",", ".")

	enhedspris = document.getElementById("FM_matenhedspris_" + m).value.replace(".", "")
	enhedspris = enhedspris.replace(",", ".")

	rabat = document.getElementById("FM_matrabat_"+m).value
	
	//Valuta 
	var valutaIdFak = document.getElementById("FM_valuta_all_1").value
	var valutaKodeFak = document.getElementById("FM_valuta_all_1").options[document.getElementById("FM_valuta_all_1").selectedIndex].text
    var valutaKursFak = document.getElementById("valutakurs_"+valutaIdFak).value
				
				
	var valutaMatId = document.getElementById("FM_matvaluta_"+m).value 	
	var valutaMatKurs = document.getElementById("valutakurs_"+valutaMatId).value
				
				
	//alert(valutaMatId +"!="+ valutaIdFak)
	
	    if (valutaMatId != valutaIdFak) {
        var this_mat_valutakursberegn = 0;
        this_mat_valutakursberegn = enhedspris/1 * (valutaMatKurs.replace(",",".")/1 / valutaKursFak.replace(",",".")/1)
        enhedspris = Math.round(this_mat_valutakursberegn*100)/ 100
        //enhedspris = String(enhedspris)
        }
	
	pris = parseFloat(enhedspris*antal) - parseFloat(enhedspris * antal * rabat)
	pris = Math.round(pris*100)/ 100
	
	document.getElementById("FM_matbeloeb_"+m).value = pris 
	document.getElementById("FM_matbeloeb_"+m).value = document.getElementById("FM_matbeloeb_"+m).value.replace(".",",")
	document.getElementById("matbelobdiv_"+m).innerHTML = document.getElementById("FM_matbeloeb_"+m).value
	
	//alert(enhedspris +" % "+ rabat)            
	    if (beregntot == 1){        
	    materialerTot()
	    }            
	}



	
	
	/// Materialer total /////
	function materialerTot() {


	var matBelobialt = 0;
	var matBelobialt_umoms = 0;
	var matBelobialtHTML = 0;
	var aktBelobialt_umoms = 0;
	antal_m = document.getElementById("FM_antal_materialer_ialt").value
	
	var valutaKodeFak = document.getElementById("FM_valuta_all_1").options[document.getElementById("FM_valuta_all_1").selectedIndex].text
		        
	            
	            for (m=0;m<antal_m;m++){
	                if (document.getElementById("FM_vis_"+m+"").checked == true){
	                matBelobialt = parseFloat(matBelobialt) + parseFloat(document.getElementById("FM_matbeloeb_"+m).value.replace(",","."))
	                
	                if (document.getElementById("FM_matikkemoms_"+m+"").checked == true){
	                matBelobialt_umoms = parseFloat(matBelobialt_umoms) + parseFloat(document.getElementById("FM_matbeloeb_"+m).value.replace(",","."))
	                }
	                
	                }
	            }

	           // alert(matBelobialt) // + " " + matBelobialt_umoms)
	         
	            document.getElementById("FM_materialer_beloeb").value = matBelobialt
	            document.getElementById("FM_materialer_beloeb").value = document.getElementById("FM_materialer_beloeb").value.replace(".",",")

	            //alert(document.getElementById("FM_timer_beloeb").value)

	            
	            timerBelobialt = document.getElementById("FM_timer_beloeb").value.replace(",",".")
	            timerBelobialt = Math.round(timerBelobialt * 100) / 100
	            //alert("matbel: "+ matBelobialt + " timerbel: " + timerBelobialt)


	            document.getElementById("FM_beloeb").value = Math.round((matBelobialt + timerBelobialt) * 100) / 100
	            document.getElementById("FM_beloeb").value = document.getElementById("FM_beloeb").value.replace(".",",")
	           

	            

	          
	            //+belobialtHTML+" "

	          
	            //document.getElementById("FM_materialer_beloeb_umoms").value = matBelobialt_umoms

	            aktBelobialt_umoms = document.getElementById("FM_timer_beloeb_umoms").value.replace(",",".")
	            aktBelobialt_umoms = Math.round(aktBelobialt_umoms * 100) / 100
	            //aktbelIaltUmoms = document.getElementById("FM_beloeb_umoms").value
	            document.getElementById("FM_beloeb_umoms").value = (matBelobialt_umoms / 1 + aktBelobialt_umoms / 1)
	            document.getElementById("FM_beloeb_umoms").value = document.getElementById("FM_beloeb_umoms").value.replace(".",",")

	            beloeb_umoms = document.getElementById("FM_beloeb_umoms").value

	            matBelobialtHTML = document.getElementById("FM_materialer_beloeb").value.replace(".", ",")
	            document.getElementById("divmatbelobtot").innerHTML = "<b>" + matBelobialtHTML + " " + valutaKodeFak + "</b>"

	            belobialtHTML = document.getElementById("FM_beloeb").value
	            document.getElementById("divbelobtot").innerHTML = "<b>"+ valutaKodeFak + "</b>"
	            document.getElementById("divbelobtot_umoms").innerHTML = "" & erp_txt_182 & ": <br> (" + beloeb_umoms + " " + valutaKodeFak + ") "
	            
	}



	
	function hidebordermat(m) {
	document.getElementById("beregn_"+m+"_m").style.borderColor = ""
	}
	
	function showbordermat(m) {
	document.getElementById("beregn_"+m+"_m").style.borderColor = "red"
	}
	
	                
	                function offsetmatant(m){
	                    var kCode = 0;

	                    if (m != 0) {
	                        //kCode = window.event.keyCode
	                        var e
	                        if (window.event) kCode = window.event.keyCode; // IE
	                        else if (e) kCode = e.which;
	                    } else {
	                        kCode = 0
	                    }


	                    if (kCode == 37) {
	                    } else {
	                        if (kCode == 39) {
	                        } else {
	                                    offSet_this = String(document.getElementById("FM_matantal_"+m+"").value.replace(".",","));
					                    offSetL_this = String(document.getElementById("FM_matantal_"+m+"").length);
					                    pkt_this = offSet_this.indexOf(",");
					                    if (pkt_this > 0){
						                    document.getElementById("FM_matantal_"+m+"").value = offSet_this.substring(0, pkt_this + 3)
					                    }
					                }
	                            }
					}
					
					
					function offsetmatenhpris(m){
					    var kCode = 0;

					    if (m != 0) {
					        //kCode = window.event.keyCode
					        var e
					        if (window.event) kCode = window.event.keyCode; // IE
					        else if (e) kCode = e.which;
					    } else {
					        kCode = 0
					    }


					    if (kCode == 37) {
					    } else {
					        if (kCode == 39) {
					        } else {
	                                    offSet_this = String(document.getElementById("FM_matenhedspris_"+m+"").value.replace(".",","));
					                    offSetL_this = String(document.getElementById("FM_matenhedspris_"+m+"").length);
					                    pkt_this = offSet_this.indexOf(",");
					                    if (pkt_this > 0){
						                    document.getElementById("FM_matenhedspris_"+m+"").value = offSet_this.substring(0, pkt_this + 5)
					                    }
					                }
	                            }
					}
					
					
	
	function tjekmatantal(m){
	    var kCode = 0;

	    if (m != 0) {
	        //kCode = window.event.keyCode
	        var e
	        if (window.event) kCode = window.event.keyCode; // IE
	        else if (e) kCode = e.which;
	    } else {
	        kCode = 0
	    }


	    if (kCode == 37) {
	    } else {
	        if (kCode == 39) {
	        } else {
	                    var oprAntalthisTj = 0;
	                    oprAntalthisTj = (document.getElementById("FM_matantal_opr_"+m+"").value.replace(",",".") / 1);
		                    if (!validZip(document.getElementById("FM_matantal_"+m+"").value)){
		                    alert("Der er angivet et ugyldigt tegn.")
		                    document.getElementById("FM_matantal_"+m+"").value = oprAntalthisTj;
		                    document.getElementById("FM_matantal_"+m+"").focus()
		                    document.getElementById("FM_matantal_"+m+"").select()
		                    return false
		                    }
		                    document.getElementById("FM_matantal_opr_"+m+"").value = document.getElementById("FM_matantal_"+m+"").value
	                    return true
	            }
	        }
	}
	
	function tjekmatehpris(m){
	    var kCode = 0;

	    if (m != 0) {
	        //kCode = window.event.keyCode
	        var e
	        if (window.event) kCode = window.event.keyCode; // IE
	        else if (e) kCode = e.which;
	    } else {
	        kCode = 0
	    }


	    if (kCode == 37) {
	    } else {
	        if (kCode == 39) {
	        } else {
	                    var oprAntalthisTj = 0;
	                    oprAntalthisTj = (document.getElementById("FM_matenhedspris_opr_"+m+"").value.replace(",",".") / 1);
		                    if (!validZip(document.getElementById("FM_matenhedspris_"+m+"").value)){
		                    alert("Der er angivet et ugyldigt tegn.")
		                    document.getElementById("FM_matenhedspris_"+m+"").value = oprAntalthisTj;
		                    document.getElementById("FM_matenhedspris_"+m+"").focus()
		                    document.getElementById("FM_matenhedspris_"+m+"").select()
		                    return false
		                    }
		                    document.getElementById("FM_matenhedspris_opr_"+m+"").value = document.getElementById("FM_matenhedspris_"+m+"").value
	                    return true
	            }
	        }
	}
	
	
	
	function showdiv(dennediv){
	
	    
	    lastopendiv = document.getElementById("lastopendiv").value

	    

	if (lastopendiv != 0) {
	document.getElementById(lastopendiv).style.display = "none";
	document.getElementById(lastopendiv).style.visibility = "hidden";
	document.getElementById(lastopendiv+"_2").style.display = "none";
	document.getElementById(lastopendiv+"_2").style.visibility = "hidden";
	//if (lastopendiv != "betdiv") {
	document.getElementById("knap_"+lastopendiv).style.background = "#ffffff";
	    //}



	if (dennediv != "fidiv") {
	    document.getElementById("genindlaes").style.display = "none";
	    document.getElementById("genindlaes").style.visibility = "hidden";
	} else {
	    document.getElementById("genindlaes").style.display = "";
	    document.getElementById("genindlaes").style.visibility = "visible";
	}


	//alert(dennediv + "_" + lastopendiv)

	if (dennediv != "jobbesk") {
	    document.getElementById("d_sideombr1").style.display = "none";
	    document.getElementById("d_sideombr1").style.visibility = "hidden";
	    document.getElementById("d_sideombr2").style.display = "none";
	    document.getElementById("d_sideombr2").style.visibility = "hidden";
	    document.getElementById("d_sideombr3").style.display = "none";
	    document.getElementById("d_sideombr3").style.visibility = "hidden";
	    document.getElementById("d_sideombr4").style.display = "none";
	    document.getElementById("d_sideombr4").style.visibility = "hidden";
	    document.getElementById("d_sideombr5").style.display = "none";
	    document.getElementById("d_sideombr5").style.visibility = "hidden";
	    document.getElementById("d_sideombr6").style.display = "none";
	    document.getElementById("d_sideombr6").style.visibility = "hidden";
	    document.getElementById("d_sideombr7").style.display = "none";
	    document.getElementById("d_sideombr7").style.visibility = "hidden";
	} else {
	    document.getElementById("d_sideombr1").style.display = "";
	    document.getElementById("d_sideombr1").style.visibility = "visible";
	    document.getElementById("d_sideombr2").style.display = "";
	    document.getElementById("d_sideombr2").style.visibility = "visible";
	    document.getElementById("d_sideombr3").style.display = "";
	    document.getElementById("d_sideombr3").style.visibility = "visible";
	    document.getElementById("d_sideombr4").style.display = "";
	    document.getElementById("d_sideombr4").style.visibility = "visible";
	    document.getElementById("d_sideombr5").style.display = "";
	    document.getElementById("d_sideombr5").style.visibility = "visible";
	    document.getElementById("d_sideombr6").style.display = "";
	    document.getElementById("d_sideombr6").style.visibility = "visible";
	    document.getElementById("d_sideombr7").style.display = "";
	    document.getElementById("d_sideombr7").style.visibility = "visible";
	}
	
	} //lasopendiv

	
	
	document.getElementById(dennediv).style.display = "";
	document.getElementById(dennediv).style.visibility = "visible";
	
	document.getElementById(dennediv+"_2").style.display = "";
	document.getElementById(dennediv+"_2").style.visibility = "visible";
	
	//if (dennediv != "betdiv") {
	document.getElementById("knap_"+dennediv).style.background = "#ffff99";
	//}

	
	
	if (dennediv == "fidiv") {
	document.getElementById("kontodiv").style.display = "";
	document.getElementById("kontodiv").style.visibility = "visible";
	} else {
	document.getElementById("kontodiv").style.display = "none";
	document.getElementById("kontodiv").style.visibility = "hidden";
	
	}
	
	document.getElementById("lastopendiv").value = dennediv

	
	
	
	}
	
	



function insertTag(what_type, field) 
{ 
var totalString; //store original string so that we can manipulated 
var selectedString; //the selected string 
var totalLength; //total length of the string 
var selectedLength; //length of the selected area 
var indexStart; //index of the string selected 
var indexFinished; //last index of the string selected 
var firstTextValue; //the first part of the string before index 
var lastTextValue; //the last part of the string after index 

totalString = document.getElementById(field).innerText; 
selectedString = document.selection.createRange().text; 
totalLength = document.getElementById(field).innerText.length; 
selectedLength = document.selection.createRange().text.length;
if (selectedLength == 0){
indexStart = totalLength
} else {
indexStart = totalString.indexOf(selectedString);
}

indexFinished = indexStart + selectedLength; 
firstTextValue = totalString.substring(0, indexStart); 
lastTextValue = totalString.substring(indexFinished, totalLength); 
if(what_type == 'b') 
{ 
selectedString = "<b>"+selectedString+"</b>"; 
} 

if(what_type == 'br') 
{ 
selectedString = selectedString+"<br>"; 
} 

if(what_type == 'i') 
{ 
selectedString = "<i>"+selectedString+"</i>"; 
}

if(what_type == 'u') 
{ 
selectedString = "<u>"+selectedString+"</u>"; 
}

document.getElementById(field).innerText = firstTextValue + selectedString + lastTextValue; 

//document.form1.text1.value = firstTextValue + selectedString + lastTextValue; 

//document.form1.text1.value = document.all("BodySpan").innerText; 
//document.form1.submit(); 

} 



function showForfaldsDato(){
antaldage = parseInt(document.getElementById("FM_forfald").value)

            dag = document.getElementById("FM_interval_slutdag").value
			mnd = document.getElementById("FM_interval_slutmrd").value 
			aar = document.getElementById("FM_interval_slutaar").value 
			var dato = new Date(aar, mnd, dag);
			
			
			
            dato.setDate(dato.getDate()+antaldage);
			
			//dato.setDate(dato.getDate() + antaldage)
			//nyDato = dato.setDate(dato.getDate() + antaldage)
			//document.getElementById("FM_slut_dag").value = dato.getDate()
			//document.getElementById("FM_slut_mrd").value = dato.getMonth() + 1

document.getElementById("showforfaldsdato").value = dato.getDate() +"."+ dato.getMonth() +"."+ parseInt(2000 + dato.getYear());
}
 



//// Funktion til forkalkuleret fakturering ////
function beregnBelob(x) {
timer = document.getElementById("FM_timerthis_"+x+"_0").value.replace(",",".")
enhpris = document.getElementById("FM_enhedspris_" + x + "_0").value.replace(",", ".") 

//if (document.getElementById("FM_show_akt_" + x + "_0").checked == true) {
    belob = (timer / 1) * (enhpris / 1)
//} else {
//    belob = 0
//}

rbt = document.getElementById("FM_rabat_"+x+"_0").value

belob = belob - (belob * (rbt/1))
belob = Math.round(belob*100)/100
belob = String(belob)

document.getElementById("FM_beloebthis_"+x+"_0").value = belob.replace(".",",") 
document.getElementById("belob_subtotal_akt_"+x+"").value = belob.replace(".",",") 
document.getElementById("timer_subtotal_akt_"+x+"").value = document.getElementById("FM_timerthis_"+x+"_0").value.replace(".",",")

//setTimerTot2(x)
setBeloebTot2_forkalk(x)

//opdvalAlleForkalk(1,x)

}


function beregnEnhedspris(x) {

timer = document.getElementById("FM_timerthis_"+x+"_0").value.replace(",",".") 
belob = document.getElementById("FM_beloebthis_"+x+"_0").value.replace(",",".")

var fenhp = 0



if (timer/1 == 0) {
    enhpris = 0
    fenhp = 1
}

if (belob/1 == 0) {
    enhpris = 0
    fenhp = 1
}

if (fenhp == 0) {
    enhpris = ((belob / 1) / (timer / 1))
    enhpris = Math.round(enhpris * 100) / 100
}

//alert("tim: " + timer + " bel: " + belob + " enhpris: " + enhpris)
enhpris = String(enhpris)
belob = String(belob)

document.getElementById("FM_enhedspris_"+x+"_0").value = enhpris.replace(".",",") 
//alert("ok")
document.getElementById("belob_subtotal_akt_"+x+"").value = belob.replace(".",",") 
document.getElementById("timer_subtotal_akt_"+x+"").value = document.getElementById("FM_timerthis_"+x+"_0").value.replace(".",",")


//setTimerTot2(x)
setBeloebTot2_forkalk(x)

//opdvalAlleForkalk(1,x)
}



function cl_spfakdato(){

    $("#sp_perstdato").hide();

}

	
	
	
    function XopdvalAlleForkalk(a,x,typ) {
        
        
        var valutaIdFak = document.getElementById("FM_valuta_all_"+a).value
	    var valutaKodeFak = document.getElementById("FM_valuta_all_"+a).options[document.getElementById("FM_valuta_all_"+a).selectedIndex].text
		
		//valutaIdopr1 = document.getElementById("FM_valuta_all_1").value
		//valutaIdopr2 = document.getElementById("FM_valuta_all_2").value
			    
	    document.getElementById("FM_valuta_all_1").value = valutaIdFak
	    document.getElementById("FM_valuta_all_2").value = valutaIdFak
	                
	            var valutaIdFak = document.getElementById("FM_valuta_all_1").value
	            var valutaKodeFak = document.getElementById("FM_valuta_all_1").options[document.getElementById("FM_valuta_all_1").selectedIndex].text
		        var valutaKursFak = document.getElementById("valutakurs_"+valutaIdFak).value
				var valutaKursBeregn = 0;
				//var belobsubakt = 0;
				var totalbelob = 0; 
				
				aRabat = document.getElementById("FM_rabat_all").value
				
                               
				   
				   //alert(aRabat)     
				        
		                var antal_x = parseInt(document.getElementById("antal_x").value);
                	    
                	    
	                    for (i=0;i<antal_x;i++){
	                           
	                          
                               
                               valutaIda = document.getElementById("FM_valuta_"+i+"_0").value
                               
                                //alert(valutaIda)
                                //valutaKodea = document.getElementById("FM_valuta_"+i+"_0").options[document.getElementById("FM_valuta_"+i+"_0").selectedIndex].text
		                        valutaKursa = document.getElementById("valutakurs_"+valutaIda).value 
                                
                                aTimer = document.getElementById("FM_timerthis_"+i+"_0").value.replace(",",".")
                                
                                if (aRabat == 0){
                                aEnhPris = document.getElementById("FM_enhedspris_opr_"+i+"_0").value.replace(",",".")
                                } else {
                                aEnhPris = document.getElementById("FM_enhedspris_"+i+"_0").value.replace(",",".")
                                }
                                
                                aBelob = (aTimer * aEnhPris) - (aTimer * aEnhPris * aRabat)
                                aBelob = Math.round(aBelob*100)/ 100
                                
                              
                                
                                //aBelob = document.getElementById("FM_beloebthis_"+i+"_0").value.replace(",",".")
                                //alert(aBelob)
                                               	                   
                	                        if (valutaIda != valutaIdFak) {
			                                var valutaKursBeregn = 0;
			                                valutaKursBeregn = aBelob/1 * (valutaKursa.replace(",",".")/1 / valutaKursFak.replace(",",".")/1)
			                                } else {
			                                valutaKursBeregn = aBelob
			                                }
			                                
			                                //alert(valutaKursBeregn)
			                              
                        			     
			                                valutaKursBeregn = Math.round(valutaKursBeregn*100)/ 100
			                                valutaKursBeregn = String(valutaKursBeregn)
			                                
			                                 
                                
                                
                        		if (typ == 'val') {
                        		aEnhPrisNy = ((valutaKursBeregn/1)/(aTimer/1))
                                aEnhPrisNy = String(Math.round(aEnhPrisNy*100)/ 100)
                        		document.getElementById("FM_enhedspris_"+i+"_0").value = aEnhPrisNy.replace(".",",")	       
                        		document.getElementById("FM_enhedspris_opr_"+i+"_0").value = aEnhPrisNy.replace(".",",")	       
                        		
                        		}
                        		
                        		document.getElementById("FM_rabat_"+i+"_0").value = aRabat
                        		
                        		document.getElementById("FM_beloebthis_"+i+"_0").value = valutaKursBeregn.replace(".",",")
                                document.getElementById("belob_subtotal_akt_"+i+"").value = valutaKursBeregn.replace(".",",")
                        		document.getElementById("FM_valuta_"+i+"_0").value = valutaIdFak	
                        		
                        		
                        		document.getElementById("valutadiv_"+i+"_0").innerHTML = "<b>"+ valutaKodeFak +"</b>"
                                document.getElementById("divbelobtot").innerHTML = "<b>"+ valutaKodeFak +"</b>"
	                 
                        			      
                        			      
                        } // For
	                    
	                   
                        // Materialer 
	                    antal_mat = parseInt(document.getElementById("FM_antal_materialer_ialt").value/1)
	                    
	                    for (m=0;m<antal_mat;m++){
	                    //document.getElementById("FM_matvaluta_"+m).value = valutaIdFak
	                    beregnmatpris(m,0)
	                    document.getElementById("matvalutadiv_"+m).innerHTML = valutaKodeFak 
	                    }



	                    //setTimerTot2(x)
    setBeloebTot2(x,1)
	                    
}	


function opd_akt_endhed_forkalk(nyenhed,val) {
	    var enh_val = 0;
	    enh_val = val
	    antal_x = (document.getElementById("antal_x").value)
	    
	    for (i=0;i<antal_x;i++){



	        // Aktiviteter
            document.getElementById("FM_akt_enh_" + i + "_0").value = enh_val
	        
            
            
	    }
	}
	
	
	
	
	function showeditor(x) {
	    //alert("her")
	 document.getElementById("editor").style.visibility = "visible"
	 document.getElementById("editor").style.display = ""
	}
	
	function hideeditor(x) {
	    //alert("her")
	 document.getElementById("editor").style.visibility = "hidden"
	 document.getElementById("editor").style.display = "none"
	}
	
    

	

	


   
	