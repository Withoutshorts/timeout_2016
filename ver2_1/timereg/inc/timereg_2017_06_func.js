

/////////// Søg job og kunder hoved søge filter ///////////////







function GetCustDesc(visning) {
    //visning tidligere = focusofoff



    var thisC = $("#FM_kontakt")
    var mrn = document.getElementById("usemrn").value;
    var ignprg = 0;

    var seljobid = document.getElementById("seljobid").value;



    //alert(seljobid)
    var sortby = 0;


    //if (document.getElementById("sort1").checked == true) {
    //    sortby = 1
    //}

    if (document.getElementById("sort2").checked == true) {
        sortby = 2
    }

    //if (document.getElementById("sort3").checked == true) {
    //    sortby = 3
    // }


    //$.scrollTo('650px', 4000);

    //$.scrollTo($("#timereg"), 2000, { offset: -300 });


    $.cookie('treg_sort', sortby);
    //alert(sortby)

    if (document.getElementById("FM_ignorer_projektgrupper").checked == true) {
        ignprg = 1
    } else {
        ignprg = 0
    }

    if (document.getElementById("FM_nyeste").checked == true) {
        visnyeste = 1
        $.cookie('nyeste', '1');
    } else {
        visnyeste = 0
        $.cookie('nyeste', '0');
    }

    if ($("#FM_easyreg").is(':checked') == true) {
        viseasyreg = 1
        $.cookie('easyreg', '1');

    } else {
        viseasyreg = 0
        $.cookie('easyreg', '0');
    }


    if ($("#FM_hr").is(':checked') == true) {
        vishr = 1
        $.cookie('hr', '1');

    } else {
        vishr = 0
        $.cookie('hr', '0');
    }


    var jobsog = document.getElementById("FM_sog_job_navn_nr").value
    var ugeStDato = $("#datoSonSQL").val()




    if ($("#FM_smartreg").is(':checked') == true) {
        vissmartreg = 1
        $.cookie('smartreg', '1');

    } else {
        vissmartreg = 0
        $.cookie('smartreg', '0');
    }


    // Sætter valgt kunde til nul ved søgning så den søger på alle //
    if (jobsog != "" && visning != 2) {
        //if ($("#FM_sog_kun_valgte").is(':checked') == false) {
        //$("#FM_kontakt").val(0)
        //}
    }

    hentjava = $("#FM_hentjava").val()
    $("#FM_hentjava").val('0')


    if ($("#FM_jobansv").is(':checked') == true) {
        jobansv = 1
        $.cookie('jobansv', '1');

    } else {
        jobansv = 0
        $.cookie('jobansv', '0');
    }

    //alert(thisC.val())

    //alert("her " + ignprg)
    //$("#BtnCustDescUpdate").data("cust", thisC.val());
    $.post("?FM_hentjava=" + hentjava + "&sortby=" + sortby + "&usemrn=" + mrn + "&ignprg=" + ignprg + "&seljobid=" + seljobid + "&visnyeste=" + visnyeste + "&jobsog=" + jobsog + "&viseasyreg=" + viseasyreg + "&vissmartreg=" + vissmartreg + "&ugeStDato=" + ugeStDato + "&jobansv=" + jobansv + "&vishr=" + vishr, { control: "FN_getCustDesc", AjaxUpdateField: "true", cust: thisC.val() }, function (data) {
        //$("#fajl").val(data);

        $("#div_jobid").html(data);
        //if (focusofoff == 1) {
        //    $("#jobid").focus();
        //}



        $(".jobid").bind('click', function () {


            var thisid = this.id
            var thisvallngt = thisid.length
            var thisvaltrim = thisid.slice(6, thisvallngt)
            thisval = thisvaltrim

            //alert(thisval)

            str = $("#seljobid").val()
            str = String(str).replace("," + thisval + ",", ",")
            $("#seljobid").val(str)

            $.post("?usemrn=" + mrn + "&jobid=" + thisval, { control: "FN_fjernjob", AjaxUpdateField: "true" }, function (data) {

            });


        });



    });


    return true;

}





$(window).load(function () {
    // run code





});










$(document).ready(function () {

    
    function tjkNorm(thisid) {

        //var thisid = this.id
        //alert(thisid)
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(9, thisvallngt)
        //alert(thisvaltrim)
        indtastTimer = $("#" + thisid).val()
        indtastTimer = indtastTimer.replace(",", ".")
        if (indtastTimer != "") {
            indtastTimer = indtastTimer
        } else {
            indtastTimer = 0
        }
        
        //alert(indtastTimer)
        dagstimer = parseFloat(dagstimer) + parseFloat(indtastTimer)
        //alert(thisid)
        timer_tastet_this_opr = $("#FM_timer_opr_" + thisvaltrim).val()
        timer_tastet_this_opr = timer_tastet_this_opr.replace(",", ".")
        
        if (timer_tastet_this_opr != "") {
            timer_tastet_this_opr = timer_tastet_this_opr
        } else {
            timer_tastet_this_opr = 0
        }
        
        timer_tastet_this_opr_dagstimer_tot = parseFloat(timer_tastet_this_opr_dagstimer_tot) + parseFloat(timer_tastet_this_opr)
               
    }

    function erTimerMereEndNorm(thisid, normdag, dagnavn, datoTxt) {

        dagsTimer_tot_new = 0
        dagsTimer_tot = 0
        // Vi skal minus de allerede oprettede timer, med de totalte timer for dagen, og lægge de nytastede timer til. 
        dagsTimer_tot = $("#" + thisid).val()
        dagsTimer_tot = dagsTimer_tot.replace(",", ".")
        if (dagsTimer_tot != 0) {
            dagsTimer_tot = dagsTimer_tot
        }
        else {
            dagsTimer_tot = 0
        }

        dagsTimer_tot_new = (parseFloat(dagsTimer_tot) - parseFloat(timer_tastet_this_opr_dagstimer_tot)) + parseFloat(dagstimer)

        if (dagsTimer_tot_new > parseFloat(normdag)) {
            $("#normTxt").append("<br>" + dagnavn + " " + datoTxt)
            visPopup = 1
        }
    }


    // tjek norm
    $(".tjknorm").click(function () {

       
        tjkNorm_btnclick();
       
    });
    // Tjek norm slut



    $("#a_ignproj_info").mouseover(function () {
        $(this).css("cursor", "pointer");
    });

    $("#al_ignproj_info").mouseover(function () {
        $(this).css("cursor", "pointer");
    });


    $("#sp_closekom").mouseover(function () {
        $(this).css("cursor", "pointer");
    });


    $("#sp_closekom").click(function () {

        closekom_x();

    });



    function closekom_x() {

        $("#matreg_indlast").html('')
        $("#dv_mat_lager").hide()
        $("#dv_mat_lager_list").hide()
        $("#dv_mat_otf").hide()
        $("#dv_mat_otf_sb").hide()
        $("#matreg_indlast_allerede").hide();
        $("#sp_mat_forbrug").html('[+]')

        // Nulstiller ecvt. err msg

        $("#matreg_indlast_err").html('')
        $("#matreg_indlast_err").css("background-color", "#ffffff")
        $("#matreg_indlast").css("background-color", "#ffffff")

        $("#kom").hide()

    }









    var isScrolled = 0;



    $("#a_filterTreg").click(function () {



        var thisid = this.id

        //alert(thisid)

        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(2, thisvallngt)
        thisval = thisvaltrim

        filterOnOff(thisval);


        //alert(thisval)

    });


    $("#a_ignproj_info").click(function () {
        if ($("#sp_ignproj_info").css('display') == "none") {


            $("#sp_ignproj_info").css("display", "");
            $("#sp_ignproj_info").css("visibility", "visible");
            $("#sp_ignproj_info").show(1000);


        } else {

            $("#sp_ignproj_info").hide(100);

        }
    });

    $("#al_ignproj_info").click(function () {


        $("#sp_ignproj_info").hide(100);


    });



    function filterOnOff(thisval) {

        var headerTxt = $("#sprog_filterheader").val()



        var jq_newfilterval = 0;

        if ($("#t_" + thisval).css('display') == "none") {

            //$.cookie('showfilter', '0');
            $("#t_" + thisval).css("display", "");
            $("#t_" + thisval).css("visibility", "visible");
            $("#t_" + thisval).show(1000);

            $("#a_" + thisval).html("<img src='ill/minus2.gif' border=0> " + headerTxt);

            jq_newfilterval = 0;
            //document.getElementById("showjobinfo").value = 1

        } else {

            //$.cookie('showfilter', '1');
            $("#t_" + thisval).hide(100);

            $("#a_" + thisval).html("<img src=ill/plus.gif border=0> " + headerTxt);

            //$("#jinftd").css("background-color", "#EFF3FF");
            //document.getElementById("showjobinfo").value = 0

            jq_newfilterval = 1;

        }


        $.post("?jq_newfilterval=" + jq_newfilterval, { control: "FN_showfilter", AjaxUpdateField: "true" }, function (data) {

        });
    }



    var headerTxt = $("#sprog_filterheader").val()

    // Autohide filter
    if ($("#FM_autohidefilter").val() == "1") {

        $("#t_filterTreg").hide(100);
        $("#a_filterTreg").html("<img src=ill/plus.gif border=0> " + headerTxt);

    } else {


        //if ($.cookie('showfilter') == "0") {
        $("#t_filterTreg").css("display", "");
        $("#t_filterTreg").css("visibility", "visible");
        $("#t_filterTreg").show(1000);

        $("#a_filterTreg").html("<img src='ill/minus2.gif' border=0> " + headerTxt);

        //} else {
        //    $("#t_filterTreg").hide(200);
        //    $("#a_filterTreg").html("<img src=ill/plus.gif border=0> Jobbank (tilføj job til aktiv-jobliste)");
        //}

    }
    ////filterOnOff('filterTreg');


    $("#a_luk_velkommen").click(function () {
        $("#div_velkommen").hide(1000);

        $.cookie('denneuge_jobans', '1');
        $.cookie('denneuge_ignrper', '1');

    });




    $("#FM_viskunetjob").click(function () {

        var mrn = document.getElementById("usemrn").value;

        var cf = confirm("Du er ved at rydde din aktive jobliste liste, ok?")

        if (cf == true) {
            $(".jobid").attr('checked', false);
            $("#seljobid").val('');


            $.post("?usemrn=" + mrn, { control: "FN_fjernAllejob", AjaxUpdateField: "true" }, function (data) {

            });


            $("#filterkri").submit();

        } else {


            $("#FM_viskunetjob").attr('checked', false);
        }

    });


    $("#a_todos").click(function () {


        show_div_todo();

    });


    function show_div_todo() {

        $("#div_todolist").css("display", "");
        $("#div_todolist").css("visibility", "visible");
        $("#a_todos").css("background-color", "#FFFF99");

        $("#a_denneuge").css("background-color", "#EFF3FF");
        $("#a_tildel").css("background-color", "#EFF3FF");
        $("#a_todos").css("background-color", "#FFFF99");

        $("#div_tildel").hide(1000);
        $("#div_denneuge").hide(1000);

        $.cookie('showhinfo', '1');

    }


    $("#a_tildel").click(function () {


        show_div_tildel();

    });


    function show_div_tildel() {

        $("#div_tildel").css("display", "");
        $("#div_tildel").css("visibility", "visible");

        $("#a_denneuge").css("background-color", "#EFF3FF");
        $("#a_tildel").css("background-color", "#FFFF99");
        $("#a_todos").css("background-color", "#EFF3FF");


        $("#div_todolist").hide(1000);
        $("#div_denneuge").hide(1000);

        $.cookie('showhinfo', '2');



    }



    $("#a_denneuge").click(function () {

        show_div_denneuge();

    });




    $("#abnlukallejob").click(function () {


        $("#timeregfm").submit()

        //for (i = 0; i < 50; i++) {


        //    jobid = $("#FM_jobid_" + i).val()

        //    hdVal = $.cookie("job_" + jobid);
        //    if (hdVal == 'visible') {
        //        $.cookie("job_" + jobid, 'hidden');
        //        $("#div_timereg_" + jobid).hide(1000);
        //    }



        //}




    });



    $("#s_luk_igv").click(function () {

        $("#jobstatusigv").hide(500);

    });

    $("#s_luk_igv").mouseover(function () {

        $(this).css("cursor", "pointer");

    });


    function show_div_denneuge() {

        $("#div_denneuge").css("display", "");
        $("#div_denneuge").css("visibility", "visible");

        $("#a_denneuge").css("background-color", "#FFFF99");
        $("#a_tildel").css("background-color", "#EFF3FF");
        $("#a_todos").css("background-color", "#EFF3FF");


        $("#div_tildel").hide(1000);
        $("#div_todolist").hide(1000);

        $.cookie('showhinfo', '0');



    }



    if ($.cookie('showhinfo') == '0') {
        show_div_denneuge();
    }

    if ($.cookie('showhinfo') == '2') {
        show_div_tildel();
    }

    if ($.cookie('showhinfo') == '1') {
        show_div_todo();
    }








    ////////////////// Tildel job //////////////////////
    $("#tildel_sel_pgrp").change(function () {

        var vlgtPrgrp = $("#tildel_sel_pgrp").val()


        tildelaktivejobliste()

    });


    tildelaktivejobliste()


    function tildelaktivejobliste() {

        //alert("her")
        var vlgtPrgrp = $("#tildel_sel_pgrp")

        if ($.cookie("tildel_sel_medarb_c") != '') {
            var medCookie = $.cookie("tildel_sel_medarb_c")
        } else {
            var medCookie = 0
        }

        //alert(medCookie)
        $.post("?vlgtmedarb=0&medCookie=" + medCookie, { control: "FN_getTildelJob_step_1", AjaxUpdateField: "true", prgrp: vlgtPrgrp.val() }, function (data) {
            //$("#fajl").val(data);

            $("#div_tildeljoblisten_msel").html(data);


            $("#tildel_sel_medarb_bt").bind('click', function () {

                tildelaktivejobliste_load();

            });



        });






    }




    function tildelaktivejobliste_load() {

        //alert("gfh")

        var vlgtPrgrp = $("#tildel_sel_pgrp")
        vlgtmedarb = $("#tildel_sel_medarb").val()


        //var medCookie = $.cookie("tildel_sel_medarb_c")
        $.cookie("tildel_sel_medarb_c", vlgtmedarb)

        //alert(vlgtmedarb)

        //$("#BtnCustDescUpdate").data("cust", thisC.val());
        $.post("?vlgtmedarb=" + vlgtmedarb, { control: "FN_getTildelJob", AjaxUpdateField: "true", prgrp: vlgtPrgrp.val() }, function (data) {
            //$("#fajl").val(data);

            $("#div_tildeljoblisten").html(data);

        });





    }



    function tjkNorm_btnclick() {


       


        norm_man = $("#norm_man").val()
        norm_tir = $("#norm_tir").val()
        norm_ons = $("#norm_ons").val()
        norm_tor = $("#norm_tor").val()
        norm_fre = $("#norm_fre").val()
        norm_lor = $("#norm_lor").val()
        norm_son = $("#norm_son").val()

        norm_man = norm_man.replace(",", ".")
        norm_tir = norm_tir.replace(",", ".")
        norm_ons = norm_ons.replace(",", ".")
        norm_tor = norm_tor.replace(",", ".")
        norm_fre = norm_fre.replace(",", ".")
        norm_lor = norm_lor.replace(",", ".")
        norm_son = norm_son.replace(",", ".")

        /*
        alert(norm_man)
        alert(norm_tir)
        alert(norm_ons)
        alert(norm_tor)
        alert(norm_fre)
        alert(norm_lor)
        alert(norm_son) */

        mandagtimer = 0
        tirsdagtimer = 0
        onsdagtimer = 0
        torsdagtimer = 0
        fredagtimer = 0
        lordagtimer = 0
        sondagtimer = 0

        timer_tastet_this_opr_mandag_tot = 0
        timer_tastet_this_opr_tirsdag_tot = 0
        timer_tastet_this_opr_onsdag_tot = 0
        timer_tastet_this_opr_torsdag_tot = 0
        timer_tastet_this_opr_fredag_tot = 0
        timer_tastet_this_opr_lordag_tot = 0
        timer_tastet_this_opr_sondag_tot = 0

        visPopup = 0

        // Mandag
        dagstimer = 0
        timer_tastet_this_opr_dagstimer_tot = 0
        timer_tastet_this_opr = 0
        $('.dcls_1').each(function () {
            var thisid = this.id
            tjkNorm(thisid);
        });

        Timer_totID = "manTimer_tot"
        erTimerMereEndNorm(Timer_totID, norm_man, "Mandag", $("#datoMan").val());


        // Tirsdag
        dagstimer = 0
        timer_tastet_this_opr_dagstimer_tot = 0
        timer_tastet_this_opr = 0
        $('.dcls_2').each(function () {

            var thisid = this.id
            tjkNorm(thisid);

        });

        Timer_totID = "tirTimer_tot"
        erTimerMereEndNorm(Timer_totID, norm_tir, "Tirsdag", $("#datoTir").val());


        //Onsdag
        dagstimer = 0
        timer_tastet_this_opr_dagstimer_tot = 0
        timer_tastet_this_opr = 0
        $('.dcls_3').each(function () {

            var thisid = this.id
            tjkNorm(thisid);
        });

        Timer_totID = "onsTimer_tot"
        erTimerMereEndNorm(Timer_totID, norm_ons, "Onsdag", $("#datoOns").val());



        // Torsdag
        dagstimer = 0
        timer_tastet_this_opr_dagstimer_tot = 0
        timer_tastet_this_opr = 0
        $('.dcls_4').each(function () {
            var thisid = this.id
            tjkNorm(thisid);
        });

        Timer_totID = "torTimer_tot"
        erTimerMereEndNorm(Timer_totID, norm_tor, "Torsdag", $("#datoTor").val());




        // Fredag
        dagstimer = 0
        timer_tastet_this_opr_dagstimer_tot = 0
        timer_tastet_this_opr = 0
        $('.dcls_5').each(function () {

            var thisid = this.id
            tjkNorm(thisid);

        });

        Timer_totID = "freTimer_tot"
        erTimerMereEndNorm(Timer_totID, norm_fre, "Fredag", $("#datoFre").val());

        //Lørdag
        dagstimer = 0
        timer_tastet_this_opr_dagstimer_tot = 0
        timer_tastet_this_opr = 0
        $('.dcls_6').each(function () {

            var thisid = this.id
            tjkNorm(thisid);

        });

        Timer_totID = "lorTimer_tot"
        erTimerMereEndNorm(Timer_totID, norm_lor, "Lørdag", $("#datoLor").val());


        // Søndag
        dagstimer = 0
        timer_tastet_this_opr_dagstimer_tot = 0
        timer_tastet_this_opr = 0
        $('.dcls_7').each(function () {

            var thisid = this.id
            tjkNorm(thisid);

        });

        Timer_totID = "sonTimer_tot"
        erTimerMereEndNorm(Timer_totID, norm_son, "Søndag", $("#datoSon").val());



        // Tjekker om vinduet skal poppe up
        if (visPopup == 1) {
            var modal = document.getElementById('normtimerModal');

            modal.style.display = "block";

            $("#opdater_timer").click(function () {
                //alert("Opdater alligevel")
                $("#timeregfm").submit();
            });

            $("#ret_timer").click(function () {
                modal.style.display = "none";
                $("#normTxt").html("")
            });
        }
        else {
            $("#timeregfm").submit();
        }

    } // Slut Tjk Norm



    //////////////////// Send besked om luk job ////////////////////



    $(".FM_lukjob").click(function () {

        //var jobid = $(this.id).val()
        //alert(this.id + " jobid: " + jobid)
        var thisid = this.id;

        var idlngt = thisid.length;
        var jobid = thisid.slice(10, idlngt);

        //alert(this.id + " jobid: " + jobid)

        sbesklukjobFunc(jobid);

        // Fjern job fra aktiv liste
        fjernjob(jobid);



        //alert("Der bliver afsendt en mail")

        $("#timeregfm").submit();

    });






    function sbesklukjobFunc(jobid) {




        var jobid = jobid


        //Sender mail til jobans
        vlgtmedarb = $("#treg_usemrn").val()
        jstatus = $("#FM_lukjobstatus_" + jobid).val()
        //alert(jstatus)

        //alert(medCookie)
        $.post("?vlgtmedarb=" + vlgtmedarb + "&jstatus=" + jstatus, { control: "FN_notify_jobluk", AjaxUpdateField: "true", jq_lukjob: jobid }, function (data) {

            GetCustDesc(0);
        });


        TONotifie("Der er afsendt email til jobansvarlige.", true)

        //$("#timeregfm").submit();


    }



    $("#FM_visallemedarb").click(function () {

        $("#seljobid").val('')
        $(".jobid").val('')
        $("#jobid_nul").val('')
        $("#FM_ch_medarb").val('1')

        var sesid = $("#FM_sesMid").val()
        $("#usemrn").val(sesid)

        $("#filterkri").submit();


    });








    /// Job afsluttet denne uge //
    function GetJobDenneUge() {

        //var thisC = $("#FM_kontakt")


        stDato = $("#datoMan").val()
        slDato = $("#datoSon").val()
        //$.cookie('treg_sort', sortby);
        //alert(stDato)

        //alert($("#denneuge_jobans").is(':checked'))

        var jdu_jobans = 0;

        if ($("#denneuge_jobans").is(':checked') == true) {
            jdu_jobans = 1
            $.cookie('denneuge_jobans', '1');
        } else {
            $.cookie('denneuge_jobans', '0');
        }


        //var jdu_ignrper = 0;

        //if ($("#denneuge_ignrper").is(':checked') == true) {
        //    jdu_ignrper = 1
        //    $.cookie('denneuge_ignrper', '1');
        //} else {
        //    $.cookie('denneuge_ignrper', '0');
        //}

        jdu_ignrper = $("#denneuge_ignrper").val()
        $.cookie('denneuge_ignrper', jdu_ignrper);

        jdu_limit = $("#denneuge_limit").val()



        var jdu_alfa = 0;

        if ($("#denneuge_alfa").is(':checked') == true) {
            jdu_alfa = 1
            $.cookie('denneuge_alfa', '1');
        } else {
            $.cookie('denneuge_alfa', '0');
        }


        lto = $("#dennuge_lto").val()

        $.post("?sortby=0&lto=" + lto + "&stDato=" + stDato + "&slDato=" + slDato + "&jobans=" + jdu_jobans + "&ignrper=" + jdu_ignrper + "&alfa=" + jdu_alfa + "&limit=" + jdu_limit, { control: "FN_jobidenneuge", AjaxUpdateField: "true" }, function (data) {
            //$("#fajl").val(data);
            $("#jobidenneuge").html(data);

        });




    }






    $("#usemrn").change(function () {
        $("#seljobid").val('')
        $(".jobid").val('')
        $("#jobid_nul").val('')
        $("#FM_ch_medarb").val('1')
        //$("#FM_viskunetjob").attr('checked', true);
        $("#filterkri").submit();
    });





    //jobid = $("#FM_jobid_1").val()
    //aktlist(jobid);



    //$(".imgDrag").mousedown(function () {
    //    $("select[name*=ajax]").AjaxUpdateField({ parent: "tr", subselector: "td:first > input[name=rowId]" });
    //    $("#incidentlist").table_sort({ items: 'tbody > tr:gt(1):has(:input[name=rowId])', IdControlNode: "td:first > :input[name=rowId]" });
    //});





    /// Job denne uge ///
    $("#denneuge_jobans").click(function () {
        GetJobDenneUge();
    });

    $("#denneuge_ignrper").change(function () {
        GetJobDenneUge();
    });

    $("#denneuge_alfa").click(function () {
        GetJobDenneUge();
    });



    if ($.cookie('denneuge_jobans') == "1") {
        $("#denneuge_jobans").attr('checked', true);

    } else {
        $("#denneuge_jobans").attr('checked', false);
    }


    if ($.cookie('denneuge_ignrper') == "1") {
        $("#denneuge_ignrper").val('1');
    }

    if ($.cookie('denneuge_ignrper') == "2") {
        $("#denneuge_ignrper").val('2');
    }

    if ($.cookie('denneuge_ignrper') == "3") {
        $("#denneuge_ignrper").val('3');
    }

    if ($.cookie('denneuge_ignrper') == "4") {
        $("#denneuge_ignrper").val('5');
    }

    if ($.cookie('denneuge_alfa') == "1") {
        $("#denneuge_alfa").attr('checked', true);

    } else {
        $("#denneuge_alfa").attr('checked', false);
    }


    if ($.cookie('c_visallemedarb_blandetliste') == "1") {
        $("#visallemedarb_blandetliste").attr('checked', true);
        //alert($.cookie('c_jobnr_blandetliste'))
        c_jobnr_blandetliste_val = $.cookie('c_jobnr_blandetliste')
        $("#jobnr_blandetliste").val(c_jobnr_blandetliste_val);
    } else {
        $("#visallemedarb_blandetliste").attr('checked', false);
    }



    GetJobDenneUge();







    //$("#jobid").bind('click', function () { alert($(this).find('option:selected').text()); });



    //var thisid = this.id
    //alert(thisid)









    $(".xxjobid").click(function () {

        var seloptions = String($(this).val())
        alert(seloptions);


        //var thisid = this.id
        //alert(thisid)
        var alloptions = "";
        var optionVal = "";
        $(".jobid").each(function () {






            optionVal = String($("#jobid option[value='" + $(this).val() + "']").val())

            instr32 = seloptions.indexOf(optionVal, 0)

            //alert(instr32)
            //instr32 = 0;
            if (instr32 == -1) {
                alloptions = alloptions + ", " + $("#jobid option[value='" + $(this).val() + "']").val()
            }

            $("#sogsubmit").focus()
        });




        //alert(alloptions)
        var trimrStr = "";
        var arr = new Array();
        arr = alloptions.split(",")

        //alert(arr[1])

        //var i = 0;
        str = $("#seljobid").val()
        for (i = 0; i < arr.length; i++) {


            trimrStr = jQuery.trim(arr[i]);
            //alert(trimrStr)

            if (trimrStr.length != 0) {
                str = String(str).replace("" + trimrStr + ",", "")
                $("#seljobid").val(str)
            }

        }

        tempstr = String($("#seljobid").val())
        //tempstr = tempstr.replace("0,", "")
        tempstr = jQuery.trim(tempstr)
        //if (tempstr.length == 0) {
        $("#seljobid").val(tempstr)
        //}

        //alert($(this).val())

    });


    ///// Job Tweet ////
    $(".aa_job_komm").click(function () {

        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(12, thisvallngt)
        thisval = thisvaltrim

        //alert(thisval)

        var jobkom = $("#FM_job_komm_" + thisval).val()
        var session_user = $("#FM_session_user").val()
        var dt = $("#FM_now").val()

        if (jobkom.length >= 100) {
            alert("Tweet er over 100 karakterer lang! \nDen er på " + jobkom.length + " karakterer.")
            return false
        }

        jobkom = jQuery.trim(jobkom)

        //alert(jobkom.length)

        if (jobkom != "Job tweet..(åben for alle)" && jobkom.length > 0) {

            $.post("?jobid=" + thisval + "&jobkom=" + jobkom + "&session_user=" + session_user + "&dt=" + dt, { control: "FN_updatejobkom", AjaxUpdateField: "true" }, function (data) {

                $("#aa_job_komm_" + thisval).text('  Ok!')
                $("#aa_job_komm_" + thisval).css("color", "#5582d2");

                $("#FM_job_komm_" + thisval).css("color", "#999999");
                $("#FM_job_komm_" + thisval).val('Job tweet..(åben for alle)');

                var forTxt = $("#dv_job_komm_" + thisval).html()
                //alert(jobkom + "<br>" + forTxt)
                $("#dv_job_komm_" + thisval).html("<span style='color:#999999;'>" + dt + ", " + session_user + ":</span> " + jobkom + "<br>" + forTxt);



            });

        }

    });

    $(".FM_job_komm").keyup(function () {

        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(12, thisvallngt)
        thisval = thisvaltrim

        //alert(thisval)

        var jobkom = $("#FM_job_komm_" + thisval).val()

        if (jobkom.length >= 100) {
            alert("Tweet er over 100 karakterer lang! \nDen er på " + jobkom.length + " karakterer.")
            return false
        }

    });

    $(".FM_job_komm").focus(function () {

        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(12, thisvallngt)
        thisval = thisvaltrim

        //alert(thisval)

        var jobkom = $("#FM_job_komm_" + thisval).val()

        //alert(jobkom)
        var newVal = ""

        if (jQuery.trim(jobkom) == "Job tweet..(åben for alle)") {
            $("#FM_job_komm_" + thisval).val(newVal)
            $("#FM_job_komm_" + thisval).css("color", "#000000");
        }

        $("#aa_job_komm_" + thisval).text('Gem')
        $("#aa_job_komm_" + thisval).css("color", "yellowgreen");

    });







    $(".fjern_job").click(function () {


        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(10, thisvallngt)
        thisval = thisvaltrim

        fjernjob(thisval);

    });





    function fjernjob(thisval) {

        //alert(thisval)



        $("#ad_timereg_" + thisval).hide(1000);
        $("#div_det_" + thisval).hide(1000);
        $("#div_timereg_" + thisval).hide(1000);

        str = $("#seljobid").val()
        str = String(str).replace("," + thisval + ",", ",")
        $("#seljobid").val(str)


        mrn = document.getElementById("usemrn").value;

        //$("#BtnCustDescUpdate").data("cust", thisC.val());
        $.post("?usemrn=" + mrn + "&jobid=" + thisval, { control: "FN_fjernjob", AjaxUpdateField: "true" }, function (data) {

        });


        $(".jobid").each(function () {
            if ($(this).val() == thisval) {
                //$(this).val()
                //alert($(this).val())
                $("#jobid_" + $(this).val()).attr('checked', false);
                //$("#jobid option[value='" + $(this).val() + "']").removeAttr("selected")

            }

        });


    }







    $(".showjobbesk").click(function () {



        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(12, thisvallngt)
        thisval = thisvaltrim

        if ($("#jobbeskdiv_" + thisval).css('display') == "none") {

            $("#jobbeskdiv_" + thisval).css("display", "");
            $("#jobbeskdiv_" + thisval).css("visibility", "visible");
            $("#jobbeskdiv_" + thisval).show(2000);
            $.scrollTo(this, 300, { offset: -200 });


        } else {

            $("#jobbeskdiv_" + thisval).hide(0);
            $.scrollTo(this, 300, { offset: -300 });

        }
    });


    function showhideJob(thisval) {

    }


    function showhidesorter() {

        if ($("#sorterdiv").css('display') == "none") {

            $("#sorterdiv").css("display", "");
            $("#sorterdiv").css("visibility", "visible");
            $("#sorterdiv").show(4000);


        } else {

            $("#sorterdiv").hide(1000);


        }
    }

    $(".todoid").click(function () {

        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(7, thisvallngt)
        thisval = thisvaltrim

        //var thisC = thisval
        //alert(thisval)

        $.post("?", { control: "FN_opdtodo", AjaxUpdateField: "true", cust: thisval }, function (data) {

            window.location.reload();

        });


    });







    /// Tjekker timer IKKE er over maks budget, maks ferie mv
    //function tjktimer_overbudget() {
    function tjektimer(thisid) {




        if (!validZip(document.getElementById(thisid).value)) {
            alert("Der er angivet et ugyldigt tegn.\nYou have typed an invalid character")
            document.getElementById(thisid).value = '';
            document.getElementById(thisid).focus()
            document.getElementById(thisid).select()
            return false
        }

        var arr = thisid.split('_');

        //TRIM til at bruge til row og ugedage på andre type felter, f.eks Timer_opr og Akttype
        var idlngt = thisid.length
        var idtrim = thisid.slice(9, idlngt)

        //Akttyper 
        var akttype = $("#FM_akttp_" + idtrim).val()

        //LIDT DET SAMME IGEN FRA GL: KODE
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(9, thisvallngt)
        thisval = thisvaltrim

        //alert(thisval)
        var res = thisval.split("_");
        aktrow = res[0]
        //alert(aktrow)
        var aktid = $("#FM_mat_aktid_" + aktrow).val()
        var jobid = $("#FM_jobid_" + aktrow).val()


        //ER ALERT VED MAKS BUDGET SLÅET TIL
        akt_maksbudget_treg = $("#akt_maksbudget_treg").val()
        akt_maksforecast_treg = $("#akt_maksforecast_treg").val()


        if (akt_maksbudget_treg == 1 || (akt_maksforecast_treg == 1 && akttype == 1)) {

            $("#aktnotificer").val('0')




            //Timertastet på linjen - minus timerOpr
            timer_tastet = 0;
            timer_tastet_this = 0;
            for (i = 1; i < 7; i++) {


                timer_tastet_this = $("#FM_timer_" + arr[2] + "_" + i).val();
                //alert(timer_tastet_this)

                if (timer_tastet_this == "NaN" || (timer_tastet_this == "-Infinity") || timer_tastet_this == "") {
                    timer_tastet_this = 0
                } else {
                    timer_tastet_this = timer_tastet_this.replace(",", ".")
                    timer_tastet = timer_tastet / 1 + timer_tastet_this / 1
                }



            }

            //alert("HER OK")

            //timer Opr
            timer_tastet_opr = 0;
            timer_tastet_this_opr = 0;
            for (i = 1; i < 7; i++) {


                timer_tastet_this_opr = $("#FM_timer_opr_" + arr[2] + "_" + i).val();
                //alert(timer_tastet_this_opr)

                if (timer_tastet_this_opr == "NaN" || (timer_tastet_this_opr == "-Infinity") || timer_tastet_this_opr == "") {
                    timer_tastet_this_opr = 0
                } else {
                    timer_tastet_this_opr = timer_tastet_this_opr.replace(",", ".")
                    timer_tastet_opr = timer_tastet_opr / 1 + timer_tastet_this_opr / 1

                }



            }


            //Total tastet diff mellem opr og indtastet
            timer_tastet = (timer_tastet / 1 - timer_tastet_opr / 1)

            //alert(timer_tastet)





        }







        //MAKS forkalk på aktivitet
        //Også interne Hvis der er angivet BUDGET (tjekkes på jquere del tim timereg.asp)
        if (akt_maksbudget_treg == 1) {

            $.post("?aktid=" + aktid + "&timer_tastet=" + timer_tastet, { control: "FN_tjktimer_budget", AjaxUpdateField: "true", cust: thisval }, function (data) {

                $("#aktnotificer").val(data);


                var budgetoverskreddet = 0;
                budgetoverskreddet = $("#aktnotificer").val().replace(",", ".")
                budgetoverskreddet = budgetoverskreddet * 1

                //alert(budgetoverskreddet)

                if (budgetoverskreddet < 0) {


                    makstimertilindtastning = ((timer_tastet * 1) - (budgetoverskreddet * -1))

                    timer_opr = $("#FM_timer_opr_" + idtrim).val();
                    $("#" + thisid).val(timer_opr);

                    alert("Forkalkulation overskreddet! \nAntal timer er over forkalkulation (budget) på denne aktivitet.\nDer kan maks tastes " + (makstimertilindtastning) + " timer.")



                }

            });

        } //akt_maksbudget_treg






        //MAKS forecast på aktivitet
        if (akt_maksforecast_treg == 1 && akttype == 1) { //kun fakturerbare

            //indenfor budgetår
            ibudgetaar = $("#aktBudgettjkOn_afgr").val()
            ibudgetaarMd = $("#regskabsaarStMd").val()
            ibudgetUseMd = $("#regskabsaarUseMd").val()
            ibudgetUseAar = $("#regskabsaarUseAar").val()
            treg_usemrn = $("#treg_usemrn").val()

            //alert("HER")



            $.post("?aktid=" + aktid + "&timer_tastet=" + timer_tastet + "&treg_usemrn=" + treg_usemrn + "&ibudgetaar=" + ibudgetaar + "&ibudgetaarMd=" + ibudgetaarMd + "&ibudgetUseMd=" + ibudgetUseMd + "&ibudgetUseAar=" + ibudgetUseAar, { control: "FN_tjktimer_forecast", AjaxUpdateField: "true", cust: thisval }, function (data) {

                //alert(data)


                $("#aktnotificer_fc").val(data);


                var fcoverskreddet = 0;
                fcoverskreddet = $("#aktnotificer_fc").val().replace(",", ".")
                fcoverskreddet = fcoverskreddet * 1

                //alert("HER 3" + akt_maksforecast_treg + " == type: " + akttype + " fcoverskreddet: " + fcoverskreddet)

                //fcoverskreddet = -2
                if (fcoverskreddet < 0) {


                    makstimertilindtastning = ((timer_tastet * 1) - (fcoverskreddet * -1))

                    timer_opr = $("#FM_timer_opr_" + idtrim).val();
                    $("#" + thisid).val(timer_opr);

                    //$("#aktnotificer_fc_" + jobid).css("padding", "20px 0px 20px 260px");

                    //$("#aktnotificer_fc_" + jobid).css("background-color", "pink");
                    //$("#aktnotificer_fc_" + jobid).html('<b>Forecast er opbrugt!</b><br>Dit timeforecast er opbrugt p&aring; denne aktivitet.<br>Du kan derfor maks. indtaste <b>' + (makstimertilindtastning) + " time(r)</b> yderligere på denne aktivitet.&nbsp;&nbsp;<span style='color:red;' id='sp_aktnotificer_fc'><b>[X]</b></span>") //"<b>Timebudget opbrugt!</b><br>Timebudget er opbrugt p&aring; denne aktivitet. <br>Du har indtastet <b>"+ formatnumber(data, 2) +" timer for meget.</b>");

                    alert("Forecast er opbrugt!\nDit timeforecast er opbrugt på denne aktivitet.")
                    //return false



                    //$('#sp_aktnotificer_fc').unbind('mouseover').bind('mouseover', function () {

                    //    $(this).css("cursor", "pointer");

                    //});


                    //$('#sp_aktnotificer_fc').unbind('click').bind('click', function () {

                    //    nulstilnofikation();

                    //});

                    //setTimeout(nulstilnofikation, 7000);

                    //function nulstilnofikation() {
                    //    $("#aktnotificer_fc_" + jobid).css("background-color", "#ffffff");
                    //    $("#aktnotificer_fc_" + jobid).html("&nbsp;");
                    //    $("#aktnotificer_fc_" + jobid).css("padding", "0px 0px 0px 0px");
                    //}

                }

            });

        } //akt_maksforecast_treg





        ///////////////////////////////////////////////////
        /////       Tjekker rejsedage - diæter          ///
        ///////////////////////////////////////////////////
        traveldietexp_on = $("#traveldietexp_on").val()


        if (traveldietexp_on == 1) {


            //Timertastet på linjen - minus timerOpr
            timer_tastet = 0;
            timer_tastet_this = 0;

            $(".aty_real_cls_1_" + arr[3]).each(function () {
                //Tjekker summen for dne ankelte dag for alle de typer der tæller med i daglig timregskab


                timer_tastet_this = $(this).val()

                if (timer_tastet_this == "NaN" || (timer_tastet_this == "-Infinity") || timer_tastet_this == "") {
                    timer_tastet_this = 0
                } else {
                    timer_tastet_this = timer_tastet_this.replace(",", ".")
                    timer_tastet = timer_tastet / 1 + timer_tastet_this / 1
                }


            });


            timer_tastet_opr = 0;
            timer_tastet_this_opr = 0;
            $(".aty_real_cls_1_opr_" + arr[3]).each(function () {
                //Tjekker summen for dne ankelte dag for alle de typer der tæller med i daglig timregskab


                timer_tastet_this_opr = $(this).val()

                if (timer_tastet_this_opr == "NaN" || (timer_tastet_this_opr == "-Infinity") || timer_tastet_this_opr == "") {
                    timer_tastet_this_opr = 0
                } else {
                    timer_tastet_this_opr = timer_tastet_this_opr.replace(",", ".")
                    timer_tastet_opr = timer_tastet_opr / 1 + timer_tastet_this_opr / 1
                }


            });


            //Total tastet diff mellem opr og indtastet
            timer_tastet = (timer_tastet / 1 - timer_tastet_opr / 1)

            //alert(timer_tastet)

            treg_usemrn = $("#treg_usemrn").val()
            //treg_dag = $("#treg_dag").val()
            //treg_md = $("#treg_md").val()
            //treg_aar = $("#treg_aar").val()
            tdatoSQL = $("#FM_mat_dato_" + idtrim).val()  //treg_aar + "-" + treg_md + "-" + treg_dag



            $.post("?timer_tastet=" + timer_tastet + "&treg_usemrn=" + treg_usemrn + "&tdato=" + tdatoSQL, { control: "FN_tjktimer_traveldietexp", AjaxUpdateField: "true", cust: thisval }, function (data) {


                //alert(data)
                $("#aktnotificer_trvl").val(data);


                var trvloverskreddet = 0;
                trvloverskreddet = $("#aktnotificer_trvl").val().replace(",", ".")
                trvloverskreddet = trvloverskreddet * 1

                //alert(trvloverskreddet)

                //trvloverskreddet = -2
                if (trvloverskreddet < 0) {


                    makstimertilindtastning = ((timer_tastet * 1) - (trvloverskreddet * -1))

                    timer_opr = $("#FM_timer_opr_" + idtrim).val();
                    $("#" + thisid).val(timer_opr);

                    $("#aktnotificer_trvl_" + jobid).css("padding", "20px 0px 20px 260px");
                    //$("#aktnotificer_" + jobid).css("top", "-420px");
                    $("#aktnotificer_trvl_" + jobid).css("background-color", "lightgray");
                    $("#aktnotificer_trvl_" + jobid).html('<b>Maks. timeantal overskreddet!</b><br>Dit maksimale timeantal på dage med rejseraktivitet er overskreddet<br>Du kan maks. indtaste <b>10 time(r)</b> på dage med rejseaktivitet.&nbsp;&nbsp;<span style="color:red;" id="sp_aktnotificer_trvl"><b>[X]</b></span>');
                    //Du kan maks indtaste <b>'+ makstimertilindtastning +'</b> på denne dag.<br>


                    $('#sp_aktnotificer_trvl').unbind('mouseover').bind('mouseover', function () {

                        $(this).css("cursor", "pointer");

                    });


                    $('#sp_aktnotificer_trvl').unbind('click').bind('click', function () {

                        nulstilnofikation();

                    });

                    setTimeout(nulstilnofikation, 7000);

                    function nulstilnofikation() {
                        $("#aktnotificer_trvl_" + jobid).css("background-color", "#ffffff");
                        $("#aktnotificer_trvl_" + jobid).html("&nbsp;");
                        $("#aktnotificer_trvl_" + jobid).css("padding", "0px 0px 0px 0px");
                    }

                }

            });




        } //traveldietexp_on


        return true



    };





    $(".a_treg").click(function () {

        //alert("ger")

        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(10, thisvallngt)
        thisval = thisvaltrim

        //alert(thisval)

        timOnOff = $("#FM_jobid_timerOn_" + thisval).val()

        if (timOnOff == 1) {
            alert("Der er indtastet / ændret timer på dette job.\nTimeOut indlæser timer..")
            $("#timeregfm").submit();
        } else {



            //alert("HER 3 er: thisid: " + thisid + " ;;" + document.getElementById("" + thisid + "").offsetTop)
            //$("#sidediv").animate({ scrollTop: "0px" }, "slow");

            //document.getElementById("" + thisid + "").scrollTop = 200;
            //$("div_timereg_" + thisval).css("top", "200px");
            //document.getElementById("div_timereg_" + thisval).scrollTop = 200;

            //alert(thisid)

            //var one = document.querySelector("#" + thisid);

            //var rect = one.getBoundingClientRect();

            //one.style.top = rect.top + "px";
            //one.style.left = rect.left + "px";
            //alert("User clicked at position (" + x + "," + rect.top + ")")


            //document.documentElement.pageYOffset = 300;
            $('html, body').scrollTo($('#div_timereg_' + thisval), { duration: 300, offset: -200 });
            //$('#sidediv').scrollTo($('#div_timereg_' + thisval), { duration: 300, offset: -200 });

            //$.scroll("#sp_a_timereg_" + thisval + "", 300, { offset: -100 }); //
            //$.scroll($("#a_timereg_" + thisval + ""), 800);

            //$.scrollTo('#sp_a_timereg_" + thisval + "', 800, { easing: 'elasout' });


            aktlist(thisval, 1);



        }


    });



    $("#bn_visallemedarb_blandetliste").click(function () {
        //alert("her")
        aktlist(0, 0);
    });

    $("#visallemedarb_blandetliste").click(function () {
        //alert("her")
        aktlist(0, 0);
    });


    //$("#jobnr_blandetliste").keyup(function () {
    //alert("her")
    //    aktlist(0,0);
    //});



    function aktlist(thisval, srcl) {


        if ($("#sp_a_timereg_" + thisval).html() == '[+]') {
            $("#sp_a_timereg_" + thisval).html('[-]')
        } else {
            $("#sp_a_timereg_" + thisval).html('[+]')
        }




        //$(".afase").bind('click', function () {

        $("#div_timereg_" + thisval).html("<table width=100%><tr><td style='padding:20px 20px 20px 20px;'>Henter aktiviteter 1-4 sek pr. 20 aktiviteter ...<br /><span style='color:darkred; font-size:10px;'>Maks. 100 - 300 aktiviteter pr. job. (lto settings)</span></td></tr></table>")

        //alert("der")
        usemrn = $("#usemrn").val()

        //alert("" + usemrn)

        stDato = $("#datoMan").val()
        slDato = $("#datoSon").val()

        lastJobstatus = $("#FM_job_laststatus_" + thisval).val()

        lastFakDato = $("#FM_job_lastfakdato_" + thisval).val()

        job_aktids = $("#job_aktids_" + thisval).val()
        job_akt_faser = $("#job_fasenavne_" + thisval).val()
        //alert(job_akt_faser)
        //job_aktnavne = $("#job_aktnavne_" + thisval).val()

        job_iRowLoops = $("#job_iRowLoop_" + thisval).val()

        if (job_iRowLoops == '0') {
            job_iRowLoops = "0,1"
        }

        sortByval = $("#sortByval").val()

        //alert(job_aktids)
        //alert(job_iRowLoops)

        if ($("#FM_vistimereltid_1").is(':checked') == true) {
            vistimereltid = 1
        } else {
            vistimereltid = 0
        }

        if ($("#FM_visSimpelAktLinje").is(':checked') == true) {
            visSimpelAktLinje = 1
        } else {
            visSimpelAktLinje = 0
        }




        //alert(visSimpelAktLinje)

        if ($("#FM_igtlaas").is(':checked') == true) {
            ingTlaas = 1
        } else {
            ingTlaas = 0
        }

        if ($("#FM_igakttype").is(':checked') == true) {
            ignakttype = 1
        } else {
            ignakttype = 0
        }

        if ($("#FM_easyreg").is(':checked') == true) {
            intEasyreg = 1
        } else {
            intEasyreg = 0
        }


        if ($("#FM_hr").is(':checked') == true) {
            intHR = 1
        } else {
            intHR = 0
        }

        if ($("#FM_ignJobogAktper").is(':checked') == true) {
            ignJobogAktper = 1
        } else {
            ignJobogAktper = 0
        }

        // tilføj linej ved blandet visning //
        //nyakt = $("#nyaktivitet").val()
        //+ "&nyakt=" + nyakt


        //if ($("#viskunejobmbudget").is(':checked') == true) {
        //    viskunForecast = 1
        //} else {
        //    viskunForecast = 0
        //}



        visallemedarb_bl = 0


        if ($("#visallemedarb_blandetliste").is(':checked') == true) { //&& (sortByval == 7)
            visallemedarb_bl = 1
            $.cookie('c_visallemedarb_blandetliste', '1');
            //$("#visallemedarb_blandetliste").attr('checked', false);
        } else {
            $.cookie('c_visallemedarb_blandetliste', '0');
        }

        jobnr_bl = $("#jobnr_blandetliste").val()
        $.cookie('c_jobnr_blandetliste', jobnr_bl);

        //alert(sortByval)



        $.post("?usemrn=" + usemrn + "&job_iRowLoops=" + job_iRowLoops + "&jobid=" + thisval + "&FM_easyreg=" + intEasyreg + "&FM_hr=" + intHR + "&ingTlaas=" + ingTlaas + "&ignakttype=" + ignakttype + "&job_aktids=" + job_aktids + "&lastFakDato=" + lastFakDato + "&job_akt_faser=" + job_akt_faser + "&stDato=" + stDato + "&slDato=" + slDato + "&FM_vistimereltid=" + vistimereltid + "&FM_ignJobogAktper=" + ignJobogAktper + "&FM_visSimpelAktLinje=" + visSimpelAktLinje + "&FM_jobstatus=" + lastJobstatus + "&sortByval=" + sortByval + "&visallemedarb_bl=" + visallemedarb_bl + "&jobnr_bl=" + jobnr_bl, { control: "FN_showakt", AjaxUpdateField: "true" }, function (data) {
            //$("#fajl").val(data);


           


            if ($("#FM_easyreg").is(':checked') == true || sortByval == 5 || $("#sortByval").val() == 6) { // Easyreg eller blandet liste
                $("#div_timereg_0").html(data);

                addHeight = 750;

                if ($("#FM_easyreg").is(':checked') == true) {
                    rowhgt = 35
                } else {

                    if ($("#jq_lto").val() == "oko" || $("#jq_lto").val() == "xintranet - local") {
                        addHeight = 150;
                        rowhgt = 70 //105

                    } else {
                        rowhgt = 105 //105
                        addHeight = 350;
                    }
                }

                var antalAkt = $('#jq_antalakt_0').val()
                //var antalAkt = $('#FM_akt_' + thisval).val()
                //alert(antalAkt)
                //if (antalAkt <= 100) {
                var dvheigt = (antalAkt / 1 * rowhgt / 1) + addHeight //750
                //} else {
                //    $("#div_timereg_" + thisval).css("height", 6000);
                //}

                //alert(dvheigt)
                $("#div_timereg_0").css("height", dvheigt);

                //alert("her")
            } else {
                $("#div_timereg_" + thisval).html(data); //html

                var antalAkt = $('#jq_antalakt_' + thisval).val()
                //var antalAkt = $('#FM_akt_' + thisval).val()
                //alert(antalAkt + "thisval " + thisval)


                if ($("#FM_visSimpelAktLinje").is(':checked') == true) {
                    rwhgt = 35
                } else {
                    rwhgt = 65
                }

                //alert(antalAkt)
                if (antalAkt <= 100) {
                    var dvheigt = (antalAkt / 1 * rwhgt / 1) + 150 //400

                    //if ($("#treg_usemrn").val() == 1) {
                    //    alert("antalAkt: " + antalAkt + "rwhgt: " + rwhgt + " = dvheigt: " + dvheigt)
                    //}


                    $("#div_timereg_" + thisval).css("height", dvheigt);
                } else {
                    $("#div_timereg_" + thisval).css("height", 750); //1200
                }

            }



       

            $(".tjknorm2").bind('click', function () {

               

                tjkNorm_btnclick();


            });



            $(".dagecls").bind('keyup', function () {

                browstype = $("#browstype").val()
                if (browstype == "mz") {
                    //alert("mz")
                    isWordcharacter = ""

                } else {

                    //var kCode = event.keyCode;
                    //alert("else")

                    var c = String.fromCharCode(event.keyCode); //event.keyCode
                    var isWordcharacter = c.match(/\w/);
                }


                if (isWordcharacter != null) {

                    //alert(isWordcharacter)

                    var thisid = this.id
                    var thisvallngt = thisid.length
                    var thisvaltrim = thisid.slice(9, thisvallngt)
                    thisval = thisvaltrim

                    var dage = $("#FM_dager_" + thisval).val().replace(",", ".")
                    var normTimerDag = $("#FM_normt_" + thisval).val().replace(",", ".")


                    //alert(normTimerDag)
                    if (normTimerDag > 0) {
                        nyTimerDag = ((dage / 1) * (normTimerDag / 1))
                        nyTimerDag = String(Math.round(nyTimerDag * 100) / 100).replace(".", ",")
                    } else {
                        nyTimerDag = 0
                    }

                    //alert(nyTimerDag)

                    $("#FM_timer_" + thisval).val(nyTimerDag)

                    //tjektimer(thisval)
                }

                //var x = event.which || event.keyCode;

                //alert(x)

                //beregnTimer(x);



            });


            function beregnTimer(e) {
                // alert(window.event.keyCode)
            }




            $('.a_showkom2').bind('mouseover', function () {

                $(this).css("cursor", "pointer");

            });


            //$('.timerflt').unbind('keyup').bind('keyup', function () {

            //    var thisid = this.id

            //tjektimer(thisid);

            //});




            $('#clskom').bind('mouseover', function () {

                $(this).css("cursor", "pointer");

            });


            $(".afase").unbind('click').bind('click', function () {

                //alert("her")

                //$("#lastclkfase").val('0')
                //afaselnk_
                var thisid = this.id
                var thisvallngt = thisid.length
                var thisvaltrim = thisid.slice(9, thisvallngt)
                idval = thisvaltrim
                showhidefase(idval);

            });


            $(".fasesog").bind('keyup', function () {



                var thisid = this.id
                var thisvallngt = thisid.length
                var thisvaltrim = thisid.slice(3, thisvallngt)
                idval = thisvaltrim

                //alert(idval)
                sogShowHideFase(idval);

            });

            $("#nyaktivitet_but").bind('click', function () {
                //aktlist(0, 0);

                mrn = document.getElementById("usemrn").value;
                aktids = $("#nyaktivitet").val();

                //alert(aktids)

                //$("#BtnCustDescUpdate").data("cust", thisC.val());
                $.post("?usemrn=" + mrn + "&aktids=" + aktids, { control: "FN_tilfojakt", AjaxUpdateField: "true" }, function (data) {

                    //alert("her")
                    //$("#div_timereg_0").html(data);

                });


                $("#timeregfm").submit();
            });


            $(".fjeasyregakt").bind('click', function () {

                var thisid = this.id

                fjernEasyregAkt(thisid);


            });

            $(".fjern_akt_bl").bind('click', function () {


                var thisid = this.id
                var thisvallngt = thisid.length
                var thisvaltrim = thisid.slice(10, thisvallngt)
                thisval = thisvaltrim


                fjernakt(thisval);

            });


            $(".las_akt_bl").bind('click', function () {



                var thisid = this.id
                var thisvallngt = thisid.length
                var thisvaltrim = thisid.slice(8, thisvallngt)
                thisval = thisvaltrim

                //alert(thisval)

                lasakt(thisval);

            });

            $(".sli_akt_bl").bind('click', function () {



                var thisid = this.id
                var thisvallngt = thisid.length
                var thisvaltrim = thisid.slice(8, thisvallngt)
                thisval = thisvaltrim

                //alert(thisval)

                sliakt(thisval);

            });




            $('.a_showkom2').bind('click', function () {
                //alert('User clicked on "foo."');


                divid = this.name;
                var thisid = this.id;

                var idlngt = thisid.length;
                var nummer = thisid.slice(4, idlngt - 2);
                var dagtype = thisid.slice(idlngt - 1, idlngt);
                //var dagtype = 1;
                //alert(nummer + "_" + dagtype)

                expandkomm(nummer, dagtype, divid);

            });








            $(".dcls_1").unbind('keyup').bind('keyup', function () {


                alertDiv(this.id)
                //dagstotal(1)
                tjektimer(this.id)
                //alert("her")
                beregnDage(this.id)



            });

            $(".dcls_2").unbind('keyup').bind('keyup', function () {

                alertDiv(this.id)
                //dagstotal(2)
                tjektimer(this.id)

                beregnDage(this.id)

            });

            $(".dcls_3").unbind('keyup').bind('keyup', function () {

                alertDiv(this.id)
                //dagstotal(3)
                tjektimer(this.id)

                beregnDage(this.id)
            });

            $(".dcls_4").unbind('keyup').bind('keyup', function () {

                alertDiv(this.id)
                //dagstotal(4)
                tjektimer(this.id)

                beregnDage(this.id)
            });

            $(".dcls_5").unbind('keyup').bind('keyup', function () {

                alertDiv(this.id)
                //dagstotal(5)
                tjektimer(this.id)

                beregnDage(this.id)
            });

            $(".dcls_6").unbind('keyup').bind('keyup', function () {

                alertDiv(this.id)
                //dagstotal(6)
                tjektimer(this.id)

                beregnDage(this.id)
            });

            $(".dcls_7").unbind('keyup').bind('keyup', function () {

                alertDiv(this.id)
                //dagstotal(7)
                tjektimer(this.id)

                beregnDage(this.id)
            });





            $(".dagecls, .dcls_1, .dcls_2, .dcls_3, .dcls_4, .dcls_5, .dcls_6, .dcls_7").bind('focus', function () {



                var thisid = this.id
                var idlngt = thisid.length;
                var thisidslice = thisid.slice(9, idlngt);
                thisval = thisidslice


                var akttype = $("#FM_akttp_" + thisval).val()


                jq_lto = $("#jq_lto").val()




                //Viser KM dialog + fylder Kontaadresser i selectbokse

                if (akttype == 5) {


                    GetFiliogKpers(thisval);

                }


            });



            $(".dagecls, .dcls_1, .dcls_2, .dcls_3, .dcls_4, .dcls_5, .dcls_6, .dcls_7").bind('keyup', function () {



                var thisid = this.id
                var idlngt = thisid.length;
                var thisidslice = thisid.slice(9, idlngt);
                thisval = thisidslice


                var akttype = $("#FM_akttp_" + thisval).val()

                // Nulstiller frokost på dagen //
                jq_lto = $("#jq_lto").val()







                //alert(jq_lto)

                doFrokostopd = $("#jq_frokostalert").val()





                if (jq_lto == "kejd_pb" || jq_lto == "fk" || jq_lto == "intranet - local") {

                    if (akttype == 7 || akttype == 8 || akttype == 20 || akttype == 21 || akttype == 24 || akttype == 25 || akttype == 31 || akttype == 115) {

                        var str = thisval
                        var dagnr = str.substring(str.length - 1, str.length);

                        // Spørg ALTID på disse typer 
                        if (doFrokostopd == 0) {
                            var r = confirm("Du er ved at indtaste fravær. \nSkal frokost nulstilles på denne dag?")

                            if (r == true) {
                                for (rw = 0; rw < 300; rw++) {

                                    akttype = $("#FM_akttp_" + rw + "_" + dagnr).val()
                                    if (akttype == '10') {
                                        $("#FM_timer_" + rw + "_" + dagnr).val(0)
                                    }
                                }
                            }



                        } else { //Dofrokost 

                            for (rw = 0; rw < 300; rw++) {

                                akttype = $("#FM_akttp_" + rw + "_" + dagnr).val()
                                if (akttype == '10') {
                                    $("#FM_timer_" + rw + "_" + dagnr).val(0)
                                }
                            }


                        }

                    } // akttype


                    // Altid frokost = 0 på ferie, ferie fri og barselsdage og omsorgsdage da de er defineret ved en hel dag.
                    if (akttype == 11 || akttype == 13 || akttype == 14 || akttype == 22 || akttype == 23 || akttype == 125) {

                        var str = thisval
                        var dagnr = str.substring(str.length - 1, str.length);

                        if (doFrokostopd == 0) {
                            alert("Du er ved at indtaste Ferie, Rejsedage eller Barsel. \nFrokost nulstilles altid på dage med fravær af denne type.")
                        }


                        for (rw = 0; rw < 300; rw++) {

                            akttype = $("#FM_akttp_" + rw + "_" + dagnr).val()
                            if (akttype == '10') {
                                $("#FM_timer_" + rw + "_" + dagnr).val(0)
                            }
                        }


                    } // akttype


                } // Lto



                // Skal kun vise ved første tast, ikke 3 gang ved f.eks 1,5
                $("#jq_frokostalert").val(1)



            });








            $(".a_showhideaktbesk").bind('click', function () {

                thisid = this.id
                var idlngt = thisid.length
                var idtrim = thisid.slice(8, idlngt)
                aktid = idtrim

                showhideaktbesk(aktid)

            });


            return true;

        });



        function beregnDage(thisid) {

            //alert("her")

            var thisid = thisid
            var idlngt = thisid.length;
            var thisidslice = thisid.slice(9, idlngt);
            thisval = thisidslice


            var akttype = $("#FM_akttp_" + thisval).val()

            if (akttype >= 11 && akttype <= 23 || akttype == 91 || akttype == 111 || akttype == 112 || akttype == 115 || akttype >= 50 && akttype <= 52 || akttype == 8 || akttype >= 120 && akttype <= 122) {

                //alert("her" + thisval)

                var timerThis = $("#FM_timer_" + thisval).val().replace(",", ".")
                var normTimerDag = $("#FM_normt_" + thisval).val().replace(",", ".")

                if (normTimerDag > 0) {
                    nyTimerDag = ((timerThis / 1) / (normTimerDag / 1))
                    nyTimerDag = String(Math.round(nyTimerDag * 100) / 100).replace(".", ",")
                } else {
                    nyTimerDag = 0
                }



                //alert(nyTimerDag)
                //nyTimerDag = "2,5"
                $("#FM_dager_" + thisval).val(nyTimerDag)


            } // akttype



        } // function



        // ******** Fjern akt **************

        function fjernakt(thisval) {

            var usemrn = $("#usemrn").val()


            //alert("her" + thisval + " m:" + usemrn)


            $.post("?aktid=" + thisval + "&usemrn=" + usemrn, { control: "FN_fjernakt", AjaxUpdateField: "true" }, function (data) {


                //$("#td_" + thisval).hide(300);
                aktlist(0, 0);

            });


        }


        // lås akt ///
        function lasakt(thisval) {

            var usemrn = $("#usemrn").val()


            //alert("her" + thisval + " m:" + usemrn)


            $.post("?aktid=" + thisval + "&usemrn=" + usemrn, { control: "FN_lasakt", AjaxUpdateField: "true" }, function (data) {

                aktlist(0, 0);
                //$("#div_timereg_0").html(data);

            });


        }

        // slip akt ///
        function sliakt(thisval) {

            var usemrn = $("#usemrn").val()


            //alert("her" + thisval + " m:" + usemrn)


            $.post("?aktid=" + thisval + "&usemrn=" + usemrn, { control: "FN_sliakt", AjaxUpdateField: "true" }, function (data) {

                aktlist(0, 0);
                //$("#div_timereg_0").html(data);

            });


        }





        function showhideaktbesk(aktid) {


            //alert(aktid)

            hgt = $("#div_aktbesk_" + aktid + "").css("height")

            if (hgt == "100px") {
                $("#div_aktbesk_" + aktid + "").css("height", "14px");
                $("#div_aktbesk_" + aktid + "").css("overflow", "hidden");
                //document.getElementById(aktid).style.height = "14px";
                //document.getElementById(aktid).style.overflow = "hidden";

                $.scrollTo("#div_aktbesk_" + aktid + "", 300, { offset: -300 }); //

            } else {
                //document.getElementById(aktid).style.height = "100px";
                //document.getElementById(aktid).style.overflow = "auto";
                $("#div_aktbesk_" + aktid + "").css("height", "100px");
                $("#div_aktbesk_" + aktid + "").css("overflow", "auto");

                $.scrollTo("#div_aktbesk_" + aktid + "", 300, { offset: -300 }); //
            }


        }










        $("#ea_kn_1").click(function () {
            fordelaesyregtimer(1);
        });

        $("#ea_kn_2").click(function () {
            fordelaesyregtimer(2);
        });

        $("#ea_kn_3").click(function () {
            fordelaesyregtimer(3);
        });

        $("#ea_kn_4").click(function () {
            fordelaesyregtimer(4);
        });

        $("#ea_kn_5").click(function () {
            fordelaesyregtimer(5);
        });

        $("#ea_kn_6").click(function () {
            fordelaesyregtimer(6);
        });

        $("#ea_kn_7").click(function () {
            fordelaesyregtimer(7);
        });



        function fordelaesyregtimer(id) {

            var antallinier = $("#antalaktlinier").val();
            //alert(antallinier)
            var antaltimer = $("#ea_" + id + "").val();

            if (antaltimer.length > 0) {

                antaltimer = antaltimer.replace(",", ".")
                if (antallinier == 0) {
                    antallinier = 1;
                }


                var res = 0;

                res = Math.round(antaltimer / antallinier * 100) / 100
                alert("Der bliver tildelt: " + res + " time til hver Easyreg. akt.\nFindes der timer i forvejen bliver de nye timer lagt til.")

                var preval = 0;
                var newval = 0;
                var oprVerdi = 0;
                var i = 1; //iRowloop

                for (i = 1; i < 500; i++) { //bør svare til maks antal EA linier




                    if (res == 0) {

                        //oprVerdi = $("#FM_timer_opr_" + i + "_" + id).val()
                        //alert("her: " + oprVerdi)
                        //(document.getElementById("FM_" + dagtype + "_opr_" + nummer + "").value / 1);
                        //$("#FM_timer_" + i + "_" + id).val(oprVerdi)
                        $("#FM_timer_" + i + "_" + id).val(0)

                    } else {

                        //if ($("#FM_timer_" + i + "_" + id).exists()) {
                        //var thisid_len = $("#FM_timer_" + i + "_" + id).val()

                        if ($("#FM_timer_" + i + "_" + id).length > 0) {

                            if (!validZip($("#FM_timer_" + i + "_" + id).val())) {
                                //alert("Der er angivet et ugyldigt tegn.")
                                $(".dcls_" + id).val(0)

                                return false

                            } else {

                                preval = $("#FM_timer_" + i + "_" + id).val()
                                preval = jQuery.trim(preval)


                                if (preval != 0 && preval.length > 0) {
                                    preval = preval.replace(",", ".")
                                    preval = preval
                                    newval = String(Math.round((preval / 1 + res / 1) * 100) / 100)
                                    newval = newval.replace(".", ",")
                                } else {
                                    //alert(res.length + "  res: " + res)
                                    if (res != 0) {
                                        preval = 0
                                        newval = String(Math.round((res / 1) * 100) / 100)
                                    } else {
                                        newval = 0
                                    }
                                }



                                $("#FM_timer_" + i + "_" + id).val(newval)
                                //$("#FM_timer_opr_" + i + "_" + id).val(newval)


                            }

                        } else {
                            //alert(res+" i: " + i)
                            res = String(res)
                            res = res.replace(".", ",")
                            $("#FM_timer_" + i + "_" + id).val(res)
                            //$("#FM_timer_opr_" + i + "_" + id).val(res)
                        }


                    } // nulstil


                } // for

            } //length ea
        }
        // end func //



        function fjernEasyregAkt(thisid) {

            var thisid = thisid

            $.post("?fjerneasyakt=1", { control: "FN_fjeasyregakt", AjaxUpdateField: "true", cust: thisid }, function (data) {

                //$.scrollTo('+=250px', 200);
                $("#" + thisid).hide(200);

                //alert("ok")
            });


        }










        function alertDiv(rloopThis) {
            var thisid = rloopThis
            var thisvallngt = thisid.length
            var thisvaltrim = thisid.slice(9, thisvallngt - 2)
            thisval = thisvaltrim


            //alert(thisval)
            //alert($("#FM_timer_" + this.id + "_1").val())
            jobid = $("#FM_jobid_" + thisval).val()
            $("#FM_jobid_timerOn_" + jobid).val('1')

        }




        /// Når siden loader ///////////////////////////////////////////////////////////////////////////////

        if ($("#FM_easyreg").is(':checked') == true || $("#sortByval").val() == 5 || $("#sortByval").val() == 6) {

        } else {





            if ($("#div_timereg_" + thisval).css('display') == "none") { //|| ($("#antaljob").val() == 1)

                $("#div_timereg_" + thisval).css("display", "");
                $("#div_timereg_" + thisval).css("visibility", "visible");
                $("#div_timereg_" + thisval).show(300);


                $.cookie("job_" + thisval, 'visible');

                if ($("#antaljob").val() == 1) {
                    ofs = -400
                } else {
                    ofs = -200
                }


                //alert(srcl)

                if (srcl == 1) {
                    $.scrollTo("#div_timereg_" + thisval, 300, { offset: ofs });
                    isScrolled = 1
                }

                //alert("ser")



            } else {


                $.cookie("job_" + thisval, 'hidden');

                //$("#timeregForm").submit();

                if (srcl == 1) {
                    $.scrollTo("#div_timereg_" + thisval, 300, { offset: -200 });
                    isScrolled = 1
                }

                $("#div_timereg_" + thisval).hide(1000);


            }



        }

        //////////////////////////////////////////////////////////////////////////////////



    }



    function sogShowHideFase(idval) {



        sogVal = $("#fs_" + idval).val().toLowerCase()

        $(".an_" + idval).each(function () {
            str = $(this).val();

            var thisid = this.id
            var thisvallngt = thisid.length
            var thisvaltrim = thisid.slice(3, thisvallngt)
            var thisJobAid = thisvaltrim
            //alert(thisJobAid)

            //alert(str.indexOf(sogVal))

            if (str.toLowerCase().indexOf(sogVal) >= 0 && sogVal.length > 0) {

                $("#" + thisJobAid + "").css("display", "");
                $("#" + thisJobAid + "").css("visibility", "visible");





            } else {


                $("#" + thisJobAid + "").css("display", "none");
                $("#" + thisJobAid + "").css("visibility", "hidden");



            }
        });



    };





    function showhidefase(idval) {



        //fasesoeg_"&jobid&"_"&lcase(job_fase)
        //vzb = $(".td_" + idval + "").css("visibility")
        vzb = $("#faseshow_" + idval + "").val()
        thisJobid = $("#fasejobid_" + idval + "").val()


        //alert(thisJobid)
        if (vzb == "hidden") {
            $(".td_" + idval + "").css("display", "");
            $(".td_" + idval + "").css("visibility", "visible");

            $(".td_" + idval + "").show(300)
            // $("#faseshow_" + idval + "").css("visibility", "visible");

            $("#faseshow_" + idval + "").val("visible")

            //document.getElementById("faseshow_" + idval + "").value =
            oldHgt = $("#div_timereg_" + thisJobid).height() //css("height", dvheigt);
            oldHgt = $("#div_timereg_" + thisJobid).css("height", oldHgt + 200);

            $.scrollTo("#faseshow_" + idval + "", 300, { offset: -20 });

        } else {
            $(".td_" + idval + "").css("display", "none");
            $(".td_" + idval + "").css("visibility", "hidden");

            $(".td_" + idval + "").hide(300)
            //$("#faseshow_" + idval + "").css("visibility", "hidden");

            //document.getElementById("faseshow_" + idval + "").value = "hidden"

            $("#faseshow_" + idval + "").val("hidden")

            oldHgt = $("#div_timereg_" + thisJobid).height() //css("height", dvheigt);
            oldHgt = $("#div_timereg_" + thisJobid).css("height", oldHgt - 200);

            $.scrollTo("#faseshow_" + idval + "", 300, { offset: -20 }); //
            //$.scrollTo(this, 800, { offset: 200 });
            //$.scrollTo('400px', 800);
        }



    }

    function expandkomm(nummer, dagtype, divid) {


        thisnum = nummer / 1

        $("#kom").css("display", "");
        $("#kom").css("visibility", "visible");

        //alert($("#t_filterTreg").css("display"))
        sortByval = $("#sortByval").val()

        if (sortByval == 5 || sortByval == 6) {

            //var left = 0
            var top = 200

        } else {
            var a = $("#ad_timereg_" + divid) //$(this);
            var pos = a.position()

            //alert(pos.top)
            //var left = pos.left
            var top = pos.top
        }


        $("#kom").css("top", top + 100);
        $("#kom").css("left", 25);
        $("#kom").show(1000);

        $.scrollTo('#kom', 200, { offset: -100 });



        thisval = $("#FM_kom_" + nummer + "" + dagtype + "").val();


        if ($("#sdskKomm").val() == "") {
            thisval = thisval
        } else {
            thisval = thisval + $("#sdskKomm").val()
        }

        thisvaloff = $("#FM_off_" + nummer + "" + dagtype + "").val()
        thisbopalvaloff = $("#FM_bopal_" + nummer + "" + dagtype + "").val()

        $("#rowcounter").val(nummer);
        $("#daytype").val(dagtype);



        $("#FM_kommentar").val(thisval);

        $("#FM_off").val(thisvaloff);
        $("#FM_bopalstur").val(thisbopalvaloff);


        $("#FM_kommentar").focus();


        thisstring = thisval.length;

        $("#antch").val(thisstring);


        //alert(nummer + "" + dagtype)
        //FM_aktivitetid
        //alert($("#FM_jobid_" + nummer).val())



        $("#matreg_regdato_0").val($("#FM_mat_dato_" + nummer + "_" + dagtype + "").val())
        $("#matreg_jobid").val($("#FM_jobid_" + nummer).val())
        $("#matreg_aktid").val($("#FM_mat_aktid_" + nummer).val())


    }








    $(".xa_treg").click(function () {

        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(10, thisvallngt)
        thisval = thisvaltrim

        //var thisC = thisval
        //alert(thisval)

        //showhideJob(thisval);
        if ($("#div_timereg_" + thisval).css('display') == "none") {

            $("#div_timereg_" + thisval).css("display", "");
            $("#div_timereg_" + thisval).css("visibility", "visible");
            $("#div_timereg_" + thisval).show(300);
            $.scrollTo(this, 300, { offset: -200 });

            $.cookie("job_" + thisval, 'visible');

        } else {

            $("#div_timereg_" + thisval).hide(1000);
            $.scrollTo(this, 300, { offset: -200 });

            $.cookie("job_" + thisval, 'hidden');

        }

    });


    $(".a_det").mouseover(function () {
        $(this).css('cursor', 'pointer');
    });

    $(".a_det").click(function () {

        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(6, thisvallngt)
        thisval = thisvaltrim

        //var thisC = thisval
        //alert(thisval)

        if ($("#div_det_" + thisval).css('display') == "none") {

            $("#div_det_" + thisval).css("display", "");
            $("#div_det_" + thisval).css("visibility", "visible");
            $("#div_det_" + thisval).show(300);

            //$.scrollTo(this, 300, { offset: -200 });



        } else {

            $("#div_det_" + thisval).hide(200);
            //$.scrollTo(this, 300, { offset: -200 });



        }


    });









    $("#ko0chk").focus(function () {
        $.scrollTo('450px', 1000);
    });

    $("#jq_hideKpersdiv").click(function () {
        //$.scrollTo('650px', 500);
        $("#korseldiv").hide('fast')

    });


    $('#jq_hideKpersdiv').mouseover(function () {

        $(this).css("cursor", "pointer");

    });

    // Kpers og filialer til Kørsel //
    //$("#FM_sog_kpers_but").click(function () {
    //GetFiliogKpers(1);
    //});

    //$("#FM_sog_kpers_dist_all").click(function() {
    //    //alert("her")
    //    GetFiliogKpers();
    //});


    $("#indlaes_koadr_2").click(function () {
        //alert("her")
        koadr_2013g();

    });




    function koadr_2013g() {
        //alert("her")
        //xval = xval/1



        var xrow = $("#koFlt").val();
        var dagn = $("#koFltx").val();

        flt = xrow + "" + dagn

        var destination = 0;
        destination = document.getElementById("ko1").value;
        //alert(destination)
        $("#FM_destination_" + flt).val(destination)


        //alert(flt)

        varTurTypeTxtFra = document.getElementById("ko0")
        varTurTypeTxtFra = varTurTypeTxtFra.options[varTurTypeTxtFra.selectedIndex].text.replace("..........", "")


        varTurTypeTxtTil = document.getElementById("ko1")
        varTurTypeTxtTil = varTurTypeTxtTil.options[varTurTypeTxtTil.selectedIndex].text.replace("..........", "")



        var kotxt = "Fra:\n" + varTurTypeTxtFra + "\n\nTil:\n" + varTurTypeTxtTil

        //alert("#FM_kom_" + flt + "")
        $("#FM_kom_" + flt + "").val(kotxt)
        $("#FM_kommentar").val(kotxt)



        //.replace("#br#", vbCrLf)
        //.replace("<BR>"," vbcrlf ")

        $("#korseldiv").hide('fast')

        // Call Lei GoogleMap function



        Lei.Indlaes(kotxt, flt);



        // 


    }




    var Lei = {
        Indlaes: function (fmKom, flt) {
            var kom = "";
            if (fmKom !== null && fmKom !== undefined)
                kom = fmKom;
            kom = encodeURIComponent(kom);
            window.open("CalculateDistance.aspx?comments=" + kom + "&flt=" + flt, "_blank", "width=500,height=200,fullscreen=no", "true");
        }


    }



    // Kørsel Henter kunder (desitnationer)
    function GetFiliogKpers(id) {

        //alert(id)

        x = $("#FM_timer_" + id).offset();


        var thisid = id
        var thisvallngt = thisid.length

        var dagn = thisid.slice(thisvallngt - 1, thisvallngt)
        var xrow = thisid.slice(0, thisvallngt - 2)


        // var strSQLKundKri;
        //var thisC = $("#FM_sog_kpers_dist").val();
        //var visalle = 0;



        var kid = $("#FM_sog_kpers_dist_kid").val();
        //alert(kid.val())

        /*if (document.getElementById("FM_sog_kpers_dist_all").checked == true) {
        if ($("#FM_sog_kpers_dist_all").is(':checked') == true) {
            visalle = 1;
        } else {
            visalle = 0;
        }
        */


        var jqmid = $("#usemrn").val();

        var visalle = 1;

        //alert("her")

        koKmDialog = document.getElementById("koKmDialog").value

        if (koKmDialog == 1) {

            //$("#BtnCustDescUpdate").data("cust", thisC.val());
            $.post("?kid=" + kid + "&visalle=" + visalle + "&jqmid=" + jqmid, { control: "FM_get_destinations", AjaxUpdateField: "true", cust: 0 }, function (data) {
                //$("#koTEST").html(data);
                //$("#fajl").html(data);
                $("#ko0").html(data);
                $("#ko1").html(data);
                //alert("her")

                /*
                if (n == 1) {
                    $("#filialerogkontakter").html("<br><hr><span style='background-color:#FFFF99;'><b>Ny søgning:</b> " + thisC.val() + "</span><br>" + data + $("#filialerogkontakter").html());
    
                } else {
                    //$("#filialerogkontakter").html("<br><hr><span style='background-color:#FFFF99;'><b>Kontakter på valgte job:</b></span><br>" + data);
    
                }
                
                */
                $("#koFlt").val(xrow);
                $("#koFltx").val(dagn);
                //document.getElementById("koFltx").value = (document.getElementById("kperfil_fundet").value / 1) * 1



            });
        }



        topPx = parseInt(x.top)

        //alert(topPx)

        $("#korseldiv").css("top", topPx);
        $("#korseldiv").css("left", 145);

        $.scrollTo("#korseldiv", 300, { offset: -400 });


    }






    /// Easyreg funktioner //
    $(".ea_kom").click(function () {



        var thisid = this.id;

        var idlngt = thisid.length;
        var idtrim = thisid.slice(5, idlngt);

        //alert(idtrim)

        showEasyKom(idtrim);


    });

    function showEasyKom(idthis) {

        $("#kom_easy").css("visibility", "visible");
        $("#kom_easy").css("display", "");
        //$.scrollTo('+=100px', 800);
        $("#FM_kommentar_easy").focus();
        $("#FM_kom_dagtype").val(idthis)

    }


    $("#close_kom_easy").click(function () {

        $("#kom_easy").css("visibility", "hidden");
        $("#kom_easy").css("display", "none");
        //$("#kom").show(4000);
        var dagtype = $("#FM_kom_dagtype").val()
        //alert(dagtype)
        var EasyKom = $("#FM_kommentar_easy").val()
        //alert(EasyKom)
        var i = 1;
        var preval = "";

        for (i = 1; i < 500; i++) {

            //$(".d_kom_cls_" + dagtype + "").val(EasyKom)
            preval = $("#FM_kom_" + i + "" + dagtype + "").val()
            $("#FM_kom_" + i + "" + dagtype + "").val(preval + "\n" + "\n" + EasyKom)

        }

        $.scrollTo('+=50px', 800);


    });



    //document.getElementById("kom").style.display = "";
    //document.getElementById("kom").style.visibility = "visible";
    //setKomFocus();












    //$("#FM_easyreg").click(function () {


    $("#FM_easyreg").click(function () {

        if ($("#FM_easyreg").is(':checked') == true) {

            $.cookie('easyreg', '1');

            $.cookie('smartreg', '0');
            $("#FM_smartreg").attr('checked', false);

            $.cookie('classicreg', '0');
            $("#FM_classicreg").attr('checked', false);

            $.cookie('nyeste', '0');
            $("#FM_nyeste").attr('checked', false);

            $.cookie('hr', '0');
            $("#FM_hr").attr('checked', false);



            $("#FM_ignorer_projektgrupper").attr('checked', false);



        }

        GetCustDesc(0);
        $("#filterkri").submit();

    });


    $("#FM_hr").click(function () {

        if ($("#FM_hr").is(':checked') == true) {


            $.cookie('hr', '1');

            $.cookie('easyreg', '0');
            $("#FM_easyreg").attr('checked', false);

            $.cookie('smartreg', '0');
            $("#FM_smartreg").attr('checked', false);

            $.cookie('classicreg', '0');
            $("#FM_classicreg").attr('checked', false);

            $.cookie('nyeste', '0');
            $("#FM_nyeste").attr('checked', false);


            $("#FM_ignorer_projektgrupper").attr('checked', false);



        }

        GetCustDesc(0);
        $("#filterkri").submit();

    });



    $("#FM_smartreg").click(function () {

        if ($("#FM_smartreg").is(':checked') == true) {

            $.cookie('smartreg', '1');

            $.cookie('easyreg', '0');
            $("#FM_easyreg").attr('checked', false);

            $.cookie('hr', '0');
            $("#FM_hr").attr('checked', false);

            $.cookie('classicreg', '0');
            $("#FM_classicreg").attr('checked', false);

            $.cookie('nyeste', '0');
            $("#FM_nyeste").attr('checked', false);


            $("#FM_ignorer_projektgrupper").attr('checked', false);



        }

        GetCustDesc(0);
        $("#filterkri").submit();

    });


    $("#FM_classicreg").click(function () {

        if ($("#FM_classicreg").is(':checked') == true) {

            $.cookie('classicreg', '1');

            $.cookie('smartreg', '0');
            $("#FM_smartreg").attr('checked', false);

            $.cookie('easyreg', '0');
            $("#FM_easyreg").attr('checked', false);

            $.cookie('nyeste', '0');
            $("#FM_nyeste").attr('checked', false);

            $.cookie('hr', '0');
            $("#FM_hr").attr('checked', false);



            $("#FM_ignorer_projektgrupper").attr('checked', false);



        }

        GetCustDesc(0);
        $("#filterkri").submit();

    });

    ////////////////////////////////////////////




    if ($.cookie('showakt') == "1") {


        //GetCustDesc();
        jobinfoshowhide();

        $("#FM_kontakt").change(function () {
            GetCustDesc(2);
        });


        $("#FM_ignorer_projektgrupper").click(function () {

            if ($("#FM_ignorer_projektgrupper").is(':checked') == true) {

                $("#FM_easyreg").attr('checked', false);
                $.cookie('easyreg', '0');

                $("#FM_hr").attr('checked', false);
                $.cookie('hr', '0');

                $.cookie('smartreg', '0');
                $("#FM_smartreg").attr('checked', false);
            }

            GetCustDesc(1);
        });

        $("#FM_nyeste").click(function () {

            if ($("#FM_nyeste").is(':checked') == true) {

                $("#FM_easyreg").attr('checked', false);
                $.cookie('easyreg', '0');

                $("#FM_hr").attr('checked', false);
                $.cookie('hr', '0');

                $.cookie('smartreg', '0');
                $("#FM_smartreg").attr('checked', false);

                $.cookie('classicreg', '0');
                $("#FM_classicreg").attr('checked', false);
            }

            //GetCustDesc(1);
            //showhidesorter();
            GetCustDesc(0);
            $("#filterkri").submit();

        });


        $("#FM_jobansv").click(function () {
            //alert("her")
            GetCustDesc(0);
        });


        $("#FM_sog_job_navn_nr").keyup(function () {
            GetCustDesc(0);
        });

        //$("#FM_sog_kun_valgte").click(function () {
        //alert("her")
        //    GetCustDesc(0);
        //});

        $("#sort0").click(function () {
            GetCustDesc(1);
        });

        $("#sort1").click(function () {
            GetCustDesc(1);
        });

        $("#sort2").click(function () {
            GetCustDesc(1);
        });

        $("#sort3").click(function () {
            GetCustDesc(1);
        });

        $("#helpandfaq").css("display", "");
        $("#helpandfaq").css("visibility", "visible");

    }




    function validZip(inZip) {
        if (inZip == "") {
            return true
        }
        if (isNum_treg(inZip)) {
            return true
        }
        return false
    }





    function tjekkm(dagtype, nummer) {

        if (!validZip(document.getElementById(thisid).value)) {
            alert("Der er angivet et ugyldigt tegn.")
            document.getElementById(thisid).value = '';
            document.getElementById(thisid).focus()
            document.getElementById(thisid).select()
            return false
        }
        return true
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





    function jobinfoshowhide() {

        //if ($.cookie('showjobinfo') == "1") {

        $("#jobinfo").css("display", "");
        $("#jobinfo").css("visibility", "visible");
        //$("#jinftd").css("background-color", "#FFFF99");
        $("#jobinfo").show(4000);



        //} else {

        //    $("#jobinfo").hide(1000);


        //}
    }


    //if (screen.width > 1024) {

    gblDivWdt = 800
    gblDivWdt_1 = 745
    pagehlpTop = 330
    //}
    //else {
    //    gblDivWdt = 745
    //    gblDivWdt_1 = 525
    //    pagehlpTop = 285
    //}



    //$("#jobinfo").width(gblDivWdt_1);
    $("#stempelur").width(gblDivWdt);
    //$("#afstem").width(gblDivWdt);

    //$("#jobbeskdiv").width(gblDivWdt_1 / 1 - 150);

    //document.getElementById("kalender").style.left = (gblDivWdt_1 / 1 + 30)
    //document.getElementById("pagehelp").style.left = (gblDivWdt_1 / 1 + 30)
    //document.getElementById("pagehelp").style.top = pagehlpTop + "px"

    //var jpos = $("#kalender")
    //var position = jpos.position()
    //position.left = position.left + 1000
    //alert(position.left) 
    // = (gblDivWdt_1 + 20)

    //$("#kalender").scrollLeft(1000);

    //$("#loadbar").hide(1000);

    var afstBGCol = "#EFF3FF";
    var tregBGCol = "#EFF3FF";
    var stempBGCol = "#EFF3FF";
    var s1BGCol = "#EFF3FF";
    var s2BGCol = "#EFF3FF";



    //$("#treg").corner("tl dog 5px tr 5px");
    //$("#afst").corner("tl 5px tr 5px");
    //$("#stemp").corner("tl 5px tr 5px");












    $("#showtim").click(function () {


        GetCustDesc(0);
        jobinfoshowhide();


        $.cookie('showakt', '1');

        afstBGCol = "#EFF3FF";
        tregBGCol = "#FFFF99";
        stempBGCol = "#EFF3FF"


        //$("#afstem").hide("fast");
        $("#stempelur").hide("fast");



        $("#treg").css("background-color", "#FFFF99");
        $("#afst").css("background-color", "#EFF3FF");
        $("#stemp").css("background-color", "#EFF3FF");


        $("#jinf_knap").css("display", "");
        $("#jinf_knap").css("visibility", "visible");
        $("#jinf_knap").show("fast");

        $("#div_todolist").css("display", "");
        $("#div_todolist").css("visibility", "visible");
        $("#div_todolist").show("fast");

        $("#kalender").css("display", "");
        $("#kalender").css("visibility", "visible");
        $("#kalender").show("fast", function () {
            // use callee so don't have to name the function
            //$(this).next().show("fast", arguments.callee);
        });

        $("#filter").css("display", "");
        $("#filter").css("visibility", "visible");
        $("#filter").show("fast");

        $("#pagehelp").css("display", "");
        $("#pagehelp").css("visibility", "visible");
        $("#pagehelp").show("fast", function () {
            // use callee so don't have to name the function
            //$(this).next().show("fast", arguments.callee);
        });





        $("#s").css("display", "");
        $("#s").css("visibility", "visible");
        $("#s").show("fast", function () {
            // use callee so don't have to name the function
            //$(this).next().show("fast", arguments.callee);
        });


        $("#timereg").css("display", "");
        $("#timereg").css("visibility", "visible");
        $("#timereg").show("fast", function () {
            // use callee so don't have to name the function
            //$(this).next().show("fast", arguments.callee);
        });

        $("#helpandfaq").css("display", "");
        $("#helpandfaq").css("visibility", "visible");


    });


    $("#afstemtot").click(function () {
        $.cookie('showakt', '5');
    });

    function smileystatshowhide() {

        s2BGCol = "#FFFF99";
        s1BGCol = "#EFF3FF";

        $("#s1").hide(1000);

        $("#s2_td").css("background-color", "#FFFF99");
        $("#s1_td").css("background-color", "#EFF3FF");


        if ($("#s2").css('display') == "none") {

            $("#s2").css("display", "");
            $("#s2").css("visibility", "visible");
            $("#s2").show(4000);

        } else {

            $("#s2").hide(1000);

        }

    }


    $("#s2_k").click(function () {

        smileystatshowhide()

    });

    $("#sA2_k").click(function () {

        smileystatshowhide()

    });


    $("#pagesetadv_but").click(function () {


        $("#pagesetadv").css("display", "");
        $("#pagesetadv").css("visibility", "visible");

        //alert($("#t_filterTreg").css("display"))

        if ($("#t_filterTreg").css("display") != "none") {
            $("#pagesetadv").css("top", "800px");
        } else {
            $("#pagesetadv").css("top", "400px");
        }

        $("#pagesetadv").show(1000);

        $.scrollTo('#pagesetadv', 300, { offset: -200 });

        $("#multivmed").hide("fast");
        $("#tregaktmsg1").hide("fast");

    });


    $("#pagesetadv_close").click(function () {


        $("#pagesetadv").hide("fast");


    });










    $("#ajobinfo").click(function () {
        if ($("#jobinfo").css('display') == "none") {

            $.cookie('showjobinfo', '1');
            $("#jobinfo").css("display", "");
            $("#jobinfo").css("visibility", "visible");
            //$("#jinftd").css("background-color", "#FFFF99");
            $("#jobinfo").show(4000);

            //document.getElementById("showjobinfo").value = 1

        } else {

            $.cookie('showjobinfo', '0');
            $("#jobinfo").hide(1000);
            //$("#jinftd").css("background-color", "#EFF3FF");
            //document.getElementById("showjobinfo").value = 0

        }
    });




    $("#afst").mouseover(function () {
        $("#afst").css("background-color", "#FFFFFF");
    });

    $("#afst").mouseout(function () {
        $("#afst").css("background-color", afstBGCol);
    });

    $("#treg").mouseover(function () {
        $("#treg").css("background-color", "#FFFFFF");
    });

    $("#treg").mouseout(function () {
        $("#treg").css("background-color", tregBGCol);
    });

    $("#stemp").mouseover(function () {
        $("#stemp").css("background-color", "#FFFFFF");
    });

    $("#stemp").mouseout(function () {
        $("#stemp").css("background-color", stempBGCol);
    });

    $("#s1_td").mouseover(function () {
        $("#s1_td").css("background-color", "#FFFFFF");
    });

    $("#s1_td").mouseout(function () {
        $("#s1_td").css("background-color", s1BGCol);
    });

    $("#s2_td").mouseover(function () {
        $("#s2_td").css("background-color", "#FFFFFF");
    });

    $("#s2_td").mouseout(function () {
        $("#s2_td").css("background-color", s2BGCol);
    });










    GetCustDesc(0);

    //alert("GetCustDesc(")

    $("#loadbar").hide(1000);


    /// Hvis der kun er et job på listen vises det altid åben
    /*
    if ($("#FM_easyreg").is(':checked') == true || $("#sortByval").val() == 5 || $("#sortByval").val() == 6) {

    } else {

    if ($("#antaljob").val() == 1) {
    thisval = $("#first_jobid").val()
    aktlist(thisval, 1);
    }
    }


    // Autoscroll by pageload //
    if (isScrolled != 1) {
    //alert("her")
    //$.scrollTo('160px', 300, { offset: 0 });

    }
    */




    if ($("#FM_easyreg").is(':checked') == true || $("#sortByval").val() == 5 || $("#sortByval").val() == 6) {

        //alert("Easureg ok")
        aktlist(0, 0);

    } else {

        hideallbut_first = $("#hideallbut_first").val()

        if (hideallbut_first != '') {
            hideallbut_first = hideallbut_first
        } else {
            hideallbut_first = 0
        }

        // ALTID = 0 12.12.2012 den fejler ved 2 javafejl og taber session. (synergi bl.a)
        //hideallbut_first = 0
        if (hideallbut_first == 0) {

            for (i = 0; i < 20; i++) {

                jobid = $("#FM_jobid_" + i).val()

                hdVal = $.cookie("job_" + jobid);
                if (hdVal == 'visible') {
                    aktlist(jobid, 0);
                }

            }

        }

        // ALTID alle akt. linier åben ///
        if (hideallbut_first == 3) {
            for (i = 0; i < 5; i++) { //50

                jobid = $("#FM_jobid_" + i).val()

                antalaktlinier = $("#job_antalaktids_" + jobid).val()

                //alert("jobid:" + jobid + "antalaktlinier:" + antalaktlinier)

                if (antalaktlinier < 15) {
                    aktlist(jobid, 0);
                }



            }
        }


    }
    // Akt linjer Åben  SLUT
    ///////////////////////////////////////////////////////

});






function showKMdailog(val, flt, chk, rk, kid) {

    //alert("her")
    document.getElementById("FM_sog_kpers_dist_kid").value = kid
    komLengthVal = document.getElementById("FM_kom_" + flt + "").value
    komLength = komLengthVal.length
    //alert(komLength)

    lastKmdiv = document.getElementById("lastkmdiv").value


    koKmDialog = document.getElementById("koKmDialog").value

    if (koKmDialog == 1) {

        if (komLength == 0) {

            if (val == 5) {

                //document.getElementById("korseldiv").style.top = 700 //+ (flt*2.5) 

                //document.getElementById("korseldiv").style.left = 225

                document.getElementById("korseldiv").style.visibility = "visible"
                document.getElementById("korseldiv").style.display = ""
                document.getElementById("kpersdiv").style.visibility = "hidden"
                document.getElementById("kpersdiv").style.display = "none"



                //document.getElementById("koVisible").selected = 1
                document.getElementById("ko0chk").focus()
                //document.getElementById("koVisible").selectedIndex = 1
                //varTurTypeTxt.options[varTurTypeTxt.selectedIndex].selected = 1

                document.getElementById("koFlt").value = flt




                document.getElementById("FM_sog_kpers_but").focus()



            }
        }
    }



}






function xkoadr() {
    //alert("her")
    //xval = xval/1

    xval = document.getElementById("koFlt").value;
    flt = document.getElementById("koFltx").value;

    flt = xval + "_" + xval

    //alert(flt)

    //alert(xval + " "+ flt)
    var kotxt = "";
    //Fra:\n
    var varTurTypeTxt = "";
    //nulstiller bopælstur ///
    document.getElementById("FM_bopal_" + flt).value = 0
    //nulstiller destination ///
    document.getElementById("FM_destination_" + flt).value = ""

    for (i = 0; i <= xval; i++) {
        //alert("i:" + i + " "+ document.getElementById("ko" + i + "chk").checked)
        if (document.getElementById("ko" + i + "chk").checked == true) {

            //Bopæls tur //
            if (i == 0) {
                document.getElementById("FM_bopal_" + flt).value = 1
            }

            //if (i > 0) {
            varTurTypeTxt = document.getElementById("ko" + i + "sel")
            varTurTypeTxt = varTurTypeTxt.options[varTurTypeTxt.selectedIndex].text
            if (i > 0) {
                kotxt = kotxt + "\n\n" + varTurTypeTxt + ":\n" + document.getElementById("ko" + i + "").innerText
            } else {
                kotxt = kotxt + varTurTypeTxt + ":\n" + document.getElementById("ko" + i + "").innerText
            }


            // Destination //
            if (i > 0) {

                //alert("her" + document.getElementById("ko" + i + "sel").value)
                if (document.getElementById("ko" + i + "sel").value == "2") {
                    document.getElementById("FM_destination_" + flt).value = document.getElementById("ko" + i + "kid").value
                }
            }



            //} else {
            //    kotxt = kotxt + document.getElementById("ko" + i + "").innerText
            //}

        }
    }

    document.getElementById("FM_kom_" + flt + "").value = kotxt
    //.replace("#br#", vbCrLf)
    //.replace("<BR>"," vbcrlf ")

    document.getElementById("korseldiv").style.visibility = "hidden"
    document.getElementById("korseldiv").style.display = "none"

    // Call Lei GoogleMap function


    Lei.Indlaes(kotxt, flt);



    // 
}


var xLei = {
    Indlaes: function (fmKom, flt) {
        var kom = "";
        if (fmKom !== null && fmKom !== undefined)
            kom = fmKom;
        kom = encodeURIComponent(kom);
        window.open("CalculateDistance.aspx?comments=" + kom + "&flt=" + flt, "_blank", "width=500,height=200,fullscreen=no", "true");
    }


}

function SetValueInParent() {



    var flt = window.location.search;

    var index = flt.indexOf("&");
    if (index != -1)
        flt = flt.substr(index + 5);
    flt = flt.substr(0, flt.length - 1) + "_" + flt.substr(flt.length - 1);
    //alert("flt: " + flt)
    var fmTimer = window.opener.document.getElementById("FM_timer_" + flt);
    if (fmTimer !== null)
        fmTimer.value = document.getElementById("lblKM").innerHTML.replace("km", "").replace(" ", "");

    if (fmTimer.value.indexOf("m") > -1) {
        var vm = fmTimer.value.replace("m", "");
        fmTimer.value = vm / 1000;
    }


    window.close();

    return false;
}


function showmultiDiv() {

    if (document.getElementById("multivmed").style.visibility == "hidden") {
        document.getElementById("multivmed").style.visibility = "visible"
        document.getElementById("multivmed").style.display = ""
        $("#tildelalle").val(1)

        alert("Vær opmærksom på, at:\n\nAlle felter med timer i, bliver tildelt til alle de valgte medarbejdere.\n\nDage indtastet på ferie, fravær mm. vil blive beregnet forskelligt hvis de valgte medarbejdere ikke har samme nomrtid.");

        //$.scrollTo('350px', 1000);

    } else {
        document.getElementById("multivmed").style.visibility = "hidden"
        document.getElementById("multivmed").style.display = "none"
        document.getElementById("tildelalle").checked = false
        $("#tildelalle").val(0)
    }


}




function visstempelur() {
    document.getElementById("stempelur").style.visibility = "visible"
    document.getElementById("stempelur").style.display = ""
}

function lukstempelur() {
    document.getElementById("stempelur").style.visibility = "hidden"
    document.getElementById("stempelur").style.display = "none"
}

function lukaftalealert() {
    document.getElementById("aftalealert").style.visibility = "hidden"
    document.getElementById("aftalealert").style.display = "none"
    document.getElementById("jobinfo").style.visibility = "visible"
    document.getElementById("jobinfo").style.display = ""
}

function lukalert() {
    document.getElementById("alert").style.visibility = "hidden"
    document.getElementById("alert").style.display = "none"
}

function checkAll(field) {
    for (i = 0; i < field.length; i++)
        field[i].checked = true;
}

function uncheckAll(field) {
    for (i = 0; i < field.length; i++)
        field[i].checked = false;
}





function setTimerTot(nummer, dagtype, job, aktivitet) {

    //piletaster
    if (window.event.keyCode == 37) {
    } else {
        if (window.event.keyCode == 39) {
            //alert(window.event.keyCode)
        } else {


            var varValue = 0;
            var varValue_total = 0;

            document.getElementById("FM_" + dagtype + "_opr_" + nummer + "").value = document.getElementById("FM_" + dagtype + "_opr_" + nummer + "").value.replace(",", ".")
            oprVerdi = (document.getElementById("FM_" + dagtype + "_opr_" + nummer + "").value / 1);

            document.getElementById("Timer_" + dagtype + "_" + nummer + "").value = document.getElementById("Timer_" + dagtype + "_" + nummer + "").value.replace(",", ".")
            varValue = (document.getElementById("Timer_" + dagtype + "_" + nummer + "").value / 1);
            document.getElementById("Timer_" + dagtype + "_" + nummer + "").value = document.getElementById("Timer_" + dagtype + "_" + nummer + "").value.replace(".", ",")

            document.getElementById("FM_" + dagtype + "_total").value = document.getElementById("FM_" + dagtype + "_total").value.replace(",", ".")
            varValue_total = (document.getElementById("FM_" + dagtype + "_total").value / 1);

            //if (varValue > 24) {
            //alert("En time-indtastning må ikke overstige 24 timer. \n Der er angivet: " + varValue + " timer.")
            //document.getElementById("Timer_"+ dagtype +"_" + nummer + "").value = document.getElementById("FM_"+ dagtype +"_opr_" + nummer + "").value.replace(".",",") //oprVerdi;
            //document.getElementById("FM_"+ dagtype +"_total").value = document.getElementById("FM_"+ dagtype +"_total").value.replace(".",",")
            //}
            //else {
            varTillaeg = (varValue - oprVerdi);
            varTotal_dag_beg = (varValue_total + varTillaeg);

            //if (varTotal_dag_beg > 24) {
            //alert("Et døgn indeholder kun 24 timer!!")
            //document.getElementById("Timer_"+ dagtype +"_" + nummer + "").value = parseFloat(oprVerdi);
            //}

            //if (varTotal_dag_beg <= 24){
            document.getElementById("FM_" + dagtype + "_total").value = parseFloat(varTotal_dag_beg);
            document.getElementById("FM_" + dagtype + "_opr_" + nummer + "").value = parseFloat(varValue);
            sonValue = document.getElementById("FM_son_total").value.replace(",", ".") / 1;
            manValue = document.getElementById("FM_man_total").value.replace(",", ".") / 1;
            tirValue = document.getElementById("FM_tir_total").value.replace(",", ".") / 1;
            onsValue = document.getElementById("FM_ons_total").value.replace(",", ".") / 1;
            torValue = document.getElementById("FM_tor_total").value.replace(",", ".") / 1;
            freValue = document.getElementById("FM_fre_total").value.replace(",", ".") / 1;
            lorValue = document.getElementById("FM_lor_total").value.replace(",", ".") / 1;
            document.getElementById("FM_week_total").value = parseFloat(sonValue + manValue + tirValue + onsValue + torValue + freValue + lorValue)
            document.getElementById("FM_week_total").value = document.getElementById("FM_week_total").value.replace(".", ",")
            document.getElementById("FM_" + dagtype + "_total").value = document.getElementById("FM_" + dagtype + "_total").value.replace(".", ",")
            //}
            //}
        }
    } // piletaster
}








//Aktiviteter expand
if (document.images) {
    plus = new Image(200, 200);
    plus.src = "ill/plus.gif";
    minus = new Image(200, 200);
    minus.src = "ill/minus2.gif";
}

function expand(de) {
    //alert(navigator.appVersion.indexOf("MSIE"))
    if (navigator.appVersion.indexOf("MSIE") == -1) {
        //alert("hej")
        //if (document.getElementById("Menu" + de)){
        if (document.getElementById("Menu" + de).style.display == "none") {
            document.getElementById("Menu" + de).style.display = "";
            document.images["Menub" + de].src = minus.src;
        } else {
            document.getElementById("Menu" + de).style.display = "none";
            document.images["Menub" + de].src = plus.src;
        } // else
    } else {
        //alert("pc")
        if (document.all("Menu" + de)) {
            if (document.all("Menu" + de).style.display == "none") {
                document.all("Menu" + de).style.display = "";
                document.images["Menub" + de].src = minus.src;
            } else {
                document.all("Menu" + de).style.display = "none";
                document.images["Menub" + de].src = plus.src;
            } //else
        } //else
    } //function
} //images




function antalchar() {
    //if (navigator.appVersion.indexOf("MSIE")==-1){
    thisstring = document.getElementById("FM_kom").value.length
    if (thisstring > 255) {
        alert("Der er ikke tilladt mere end 255 karakterer i en kommentar! \n Du har brugt: " + thisstring + "")
    } else {
        document.getElementById("antch").value = thisstring;
    }
}


function antalakt(jid) {
    newval = document.all["FM_akt_" + jid].value;
    document.all["FM_antal_akt_" + jid].value = newval;
}

function closetomanyjob() {
    document.all["tomanyjob"].style.visibility = "hidden";
}

///  ressource timer tildelt ///////
function showrestimer() {
    document.getElementById("ressourcetimer").style.display = "";
    document.getElementById("ressourcetimer").style.visibility = "visible";
}

function hiderestimer() {
    document.getElementById("ressourcetimer").style.display = "none";
    document.getElementById("ressourcetimer").style.visibility = "hidden";
}

///  viser indtastede timer på den enkelte dag ///////
function showtimedetail(thisnameid) {
    document.getElementById("ressourcetimer").style.display = "none";
    document.getElementById("ressourcetimer").style.visibility = "hidden";

    document.getElementById("timedetailson").style.display = "none";
    document.getElementById("timedetailson").style.visibility = "hidden";

    document.getElementById("timedetailman").style.display = "none";
    document.getElementById("timedetailman").style.visibility = "hidden";

    document.getElementById("timedetailtir").style.display = "none";
    document.getElementById("timedetailtir").style.visibility = "hidden";

    document.getElementById("timedetailons").style.display = "none";
    document.getElementById("timedetailons").style.visibility = "hidden";

    document.getElementById("timedetailtor").style.display = "none";
    document.getElementById("timedetailtor").style.visibility = "hidden";

    document.getElementById("timedetailfre").style.display = "none";
    document.getElementById("timedetailfre").style.visibility = "hidden";

    document.getElementById("timedetaillor").style.display = "none";
    document.getElementById("timedetaillor").style.visibility = "hidden";

    document.getElementById(thisnameid).style.display = "";
    document.getElementById(thisnameid).style.visibility = "visible";
}

function hidetimedetail(thisid) {
    document.getElementById("timedetailson").style.display = "none";
    document.getElementById("timedetailson").style.visibility = "hidden";

    document.getElementById("timedetailman").style.display = "none";
    document.getElementById("timedetailman").style.visibility = "hidden";

    document.getElementById("timedetailtir").style.display = "none";
    document.getElementById("timedetailtir").style.visibility = "hidden";

    document.getElementById("timedetailons").style.display = "none";
    document.getElementById("timedetailons").style.visibility = "hidden";

    document.getElementById("timedetailtor").style.display = "none";
    document.getElementById("timedetailtor").style.visibility = "hidden";

    document.getElementById("timedetailfre").style.display = "none";
    document.getElementById("timedetailfre").style.visibility = "hidden";

    document.getElementById("timedetaillor").style.display = "none";
    document.getElementById("timedetaillor").style.visibility = "hidden";
}


//function showmilepal(){
//document.getElementById("divmilepal").style.display = "";
//document.getElementById("divmilepal").style.visibility = "visible";

//document.getElementById("divopgave").style.display = "none";
//document.getElementById("divopgave").style.visibility = "hidden";

//document.getElementById("vistimerformedarb").style.display = "none";
//document.getElementById("vistimerformedarb").style.visibility = "hidden";


//document.getElementById("knapopgliste").src = "../ill/knap_opglisten_off.gif";
//document.getElementById("knapmilepal").src = "../ill/knap_milepal_on.gif";
//document.getElementById("knaptimer").src = "../ill/knap_timer_off.gif";
//}

//function showopgave(){
//document.getElementById("divmilepal").style.display = "none";
//document.getElementById("divmilepal").style.visibility = "hidden";

//document.getElementById("knapopgliste").src = "../ill/knap_opglisten_on.gif";
//document.getElementById("knapmilepal").src = "../ill/knap_milepal_off.gif";
//document.getElementById("knaptimer").src = "../ill/knap_timer_off.gif";

//document.getElementById("divopgave").style.display = "";
//document.getElementById("divopgave").style.visibility = "visible";

//document.getElementById("vistimerformedarb").style.display = "none";
//document.getElementById("vistimerformedarb").style.visibility = "hidden";
//}

//function showtimer(){
//document.getElementById("divmilepal").style.display = "none";
//document.getElementById("divmilepal").style.visibility = "hidden";

//document.getElementById("divopgave").style.display = "none";
//document.getElementById("divopgave").style.visibility = "hidden";

//document.getElementById("knapopgliste").src = "../ill/knap_opglisten_off.gif";
//document.getElementById("knapmilepal").src = "../ill/knap_milepal_off.gif";
//document.getElementById("knaptimer").src = "../ill/knap_timer_on.gif";

//document.getElementById("vistimerformedarb").style.display = "";
//document.getElementById("vistimerformedarb").style.visibility = "visible";
//}








//// NEW 2006 ///









var boolCtrlDown = false;
var curField = null;


var boolCtrlDown = false;
var curField = null;



function doKeyDown() {

    //alert("her")

    var elm = event.srcElement;
    curField = (elm.tagName.toUpperCase() == 'INPUT' && elm.type == 'text') ? elm : null;
    var kCode = event.keyCode;
    var boolNoCheck = false;
    //alert (kCode) 
    switch (kCode) {
        case 90: if (boolCtrlDown) navVer('previous'); event.returnValue = false; event.cancelBubble = true; break; //Ctrl + venstre pil (37)
        case 38: navHor('previous'); event.returnValue = false; event.cancelBubble = true; break; //pil op
        case 88: if (boolCtrlDown) navVer('next'); event.returnValue = false; event.cancelBubble = true; break; // Ctrl + højre pil
        case 40: navHor('next'); event.returnValue = false; event.cancelBubble = true; break; //pil ned
    }
    boolCtrlDown = (kCode == 17);
}



function navVer(dir) {
    var elm = eval('curField.parentElement.' + dir + 'Sibling');
    if (elm != null) {
        elm = elm.firstChild;
        if (elm != null && elm.tagName.toUpperCase() == 'INPUT' && elm.type == 'text') elm.focus();
    }
}


function navHor(dir) {
    var elm = eval('curField.parentElement.parentElement.' + dir + 'Sibling');
    var ix = 0;
    var tmpElm = curField.parentElement;
    while (tmpElm.previousSibling) {
        ix++;
        tmpElm = tmpElm.previousSibling;
    }
    if (elm != null) {
        elm = elm.firstChild;
        if (elm != null) {
            i = 0;
            while (ix > i) {
                if (elm.nextSibling) {
                    elm = elm.nextSibling;
                }
                else return false;
                i++;
            }
            if (elm.firstChild != null) elm.firstChild.focus();
        }
    }
}















function closekomm() {
    //if (navigator.appVersion.indexOf("MSIE")==-1){

    thisstring = document.getElementById("FM_kommentar").value.length
    thisvalue = document.getElementById("FM_kommentar").value
    thisrow = document.getElementById("rowcounter").value
    thisdaytype = document.getElementById("daytype").value
    thisoffentlig = document.getElementById("FM_off").value
    thisbopalvaloff = document.getElementById("FM_bopalstur").value

    //alert(document.getElementById("anc_" + thisrow + "" + thisdaytype + "").name)


    // Markerer felt
    lasttd = document.getElementById("lasttd").value
    lasttddaytype = document.getElementById("lasttddaytype").value

    if (lasttd != "0") {
        if (lasttddaytype >= 6) {
            document.getElementById("td_" + lasttd + "").style.background = "#dcdcdf";
        } else {
            document.getElementById("td_" + lasttd + "").style.background = "#ffffe1";
        }
    }

    document.getElementById("td_" + thisrow + "" + thisdaytype + "").style.background = "#FFff99";
    document.getElementById("lasttd").value = "" + thisrow + "" + thisdaytype + "";
    document.getElementById("lasttddaytype").value = "" + thisdaytype + "";

    if (thisstring > 1255) {
        alert("Der er ikke tilladt mere end 1255 karakterer i en kommentar! \n Du har brugt: " + thisstring + "")
    } else {
        document.getElementById("FM_kom_" + thisrow + "" + thisdaytype + "").value = thisvalue;
        document.getElementById("FM_off_" + thisrow + "" + thisdaytype + "").value = thisoffentlig;
        document.getElementById("FM_off").value = 0

        document.getElementById("FM_bopal_" + thisrow + "" + thisdaytype + "").value = thisbopalvaloff
        document.getElementById("FM_bopalstur").value = 0

        document.getElementById("kom").style.display = "none";
        document.getElementById("kom").style.visibility = "hidden";




        /// finder tilbage på siden ///
        //document.getElementById("anc_" + thisrow + "_" + thisdaytype + "").select();
        //document.getElementById("anc_" + thisrow + "_" + thisdaytype + "").focus();
        //document.getElementById("anc_" + thisrow + "_" + thisdaytype + "").scrollIntoView(true);
        //window."td_" + thisrow + "" + thisdaytype + "".focus;
    }
}

function antalchar() {
    //if (navigator.appVersion.indexOf("MSIE")==-1){
    thisstring = document.getElementById("FM_kommentar").value.length
    if (thisstring > 1255) {
        alert("Der er ikke tilladt mere end 1255 karakterer i en kommentar! \n Du har brugt: " + thisstring + "")
    } else {
        document.getElementById("antch").value = thisstring;
    }
}

function popUp(URL, width, height, left, top) {
    window.open(URL, 'navn', 'left=' + left + ',top=' + top + ',toolbar=0,scrollbars=1,location=0,statusbar=1,menubar=0,resizable=1,width=' + width + ',height=' + height + '');
}


function showtimeregdetail(divid) {


    for (i = 0; i < 8; i++) {
        document.getElementById("tr_" + i + "").style.display = "none";
        document.getElementById("tr_" + i + "").style.visibility = "hidden";
    }

    //document.getElementById("smiley").style.display = "none";
    //document.getElementById("smiley").style.visibility = "hidden";


    document.getElementById("tr_" + divid + "").style.display = "";
    document.getElementById("tr_" + divid + "").style.visibility = "visible";
}


function visSmileystatus() {
    document.getElementById("smileydiv").style.display = "";
    document.getElementById("smileydiv").style.visibility = "visible";
}

function gemSmileystatus() {
    document.getElementById("smileydiv").style.display = "none";
    document.getElementById("smileydiv").style.visibility = "hidden";
}



function markerjob(tdid) {
    lastjob = document.getElementById("lastjob").value

    if (lastjob != 0) {
        document.getElementById("tr_" + lastjob + "").style.background = "#FFffff";
    }

    document.getElementById("tr_" + tdid + "").style.background = "#FFff99";
    document.getElementById("lastjob").value = tdid
}


function markerfelt(feltid, thisdaytype) {
    // Markerer felt
    lasttd = document.getElementById("lasttd").value
    lasttddaytype = document.getElementById("lasttddaytype").value

    if (lasttd != "0") {
        if (lasttddaytype >= 6) {
            document.getElementById("td_" + lasttd + "").style.background = "#dcdcd0";
        } else {
            document.getElementById("td_" + lasttd + "").style.background = "#ffffe1";
        }
    }

    document.getElementById("td_" + feltid + "").style.background = "#FFff99";
    document.getElementById("lasttd").value = "" + feltid + "";
    document.getElementById("lasttddaytype").value = "" + thisdaytype + "";
}

function renssog() {
    document.getElementById("FM_sog_job_navn_nr").value = ""
}

//function showJobBesk(){
//document.getElementById("jobbeskdiv").style.display = "";
//document.getElementById("jobbeskdiv").style.visibility = "visible";
//document.getElementById("kpersdiv").style.display = "none";
//document.getElementById("kpersdiv").style.visibility = "hidden";
//}

//function hideJobBesk(){
//document.getElementById("jobbeskdiv").style.display = "none";
//document.getElementById("jobbeskdiv").style.visibility = "hidden";
//}

function curserType(tis) {
    document.getElementById(tis).style.cursor = 'hand'
}

//function showAktBesk(aid){
//document.getElementById("aktbesk_"+aid).style.display = "";
//document.getElementById("aktbesk_"+aid).style.visibility = "visible";
//}


function showKpersdiv() {
    document.getElementById("kpersdiv").style.top = 445
    document.getElementById("kpersdiv").style.left = 480

    document.getElementById("kpersdiv").style.display = "";
    document.getElementById("kpersdiv").style.visibility = "visible";
    //document.getElementById("jobbeskdiv").style.display = "none";
    //document.getElementById("jobbeskdiv").style.visibility = "hidden";
}

function hideKpersdiv() {

    //document.getElementById("korseldiv").style.display = "none";
    //document.getElementById("korseldiv").style.visibility = "hidden";


    document.getElementById("kpersdiv").style.display = "none";
    document.getElementById("kpersdiv").style.visibility = "hidden";
}

function showdiv(div) {
    //alert(div)
    lastopendiv = document.getElementById("lastopendiv").value

    document.getElementById(lastopendiv).style.display = "none";
    document.getElementById(lastopendiv).style.visibility = "hidden";

    document.getElementById("knap_" + lastopendiv).style.background = "#FFFFFF";
    document.getElementById("knap_" + lastopendiv).style.border = 0;


    document.getElementById(div).style.display = "";
    document.getElementById(div).style.visibility = "visible";

    document.getElementById("knap_" + div).style.background = "#FFFF99";
    document.getElementById("knap_" + div).style.border = "1px orange solid";

    // Sørger for at udløbet / lukket aftale alert blive lukket.
    document.getElementById("aftalealert").style.visibility = "hidden"
    document.getElementById("aftalealert").style.display = "none"

    document.getElementById("lastopendiv").value = div
}



var ns, ns6, ie, newlayer;

ns4 = (document.layers) ? true : false;
ie4 = (document.all) ? true : false
ie5 = (document.getElementById) ? true : false
ns6 = (document.getElementById && !document.all) ? true : false;

function getLayerStyle(lyr) {
    if (ns4) {
        return document.layers[lyr];
    } else if (ie4) {
        return document.all[lyr].style;
    } else if (ie5) {
        return document.all[lyr].style;
    } else if (ns6) {
        return document.getElementById(lyr).style;
    }
}

function ShowHide(layer) {
    newlayer = getLayerStyle(layer)

    var styleObj = (ns4) ? document.layers[layer] : (ie4) ? document.all[layer].style : document.getElementById(layer).style;

    if (newlayer.visibility == "hidden") {
        newlayer.visibility = "visible";
        styleObj.display = ""
    }
    else if (newlayer.visibility == "visible") {
        newlayer.visibility = "hidden";
        styleObj.display = "none"
    }
}









