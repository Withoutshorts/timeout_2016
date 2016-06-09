

/////////// Søg job og kunder hoved søge filter ///////////////
function GetCustDesc(focusofoff) {

    var thisC = $("#FM_kontakt")
    var mrn = document.getElementById("usemrn").value;
    var ignprg = 0;
    var seljobid = document.getElementById("seljobid").value;
    var sortby = 0;


    if (document.getElementById("sort1").checked == true) {
        sortby = 1
    }

    if (document.getElementById("sort2").checked == true) {
        sortby = 2
    }

    if (document.getElementById("sort3").checked == true) {
        sortby = 3
    }


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

    var jobsog = document.getElementById("FM_sog_job_navn_nr").value
    var ugeStDato = $("#datoSonSQL").val()

    //if (ugeStDato != "") {
    //    ugeStDato = ugeStDato
    //    } else {
    //    ugeStDato = "2002-1-1"
    //    }

    //alert(ugeStDato)


    // Sætter valgt kunde til nul ved søgning så den søger på alle //
    if (jobsog != "") {
        $("#FM_kontakt").val(0)
    }

    //$("#BtnCustDescUpdate").data("cust", thisC.val());
    $.post("?sortby=" + sortby + "&usemrn=" + mrn + "&ignprg=" + ignprg + "&seljobid=" + seljobid + "&visnyeste=" + visnyeste + "&jobsog=" + jobsog + "&viseasyreg=" + viseasyreg + "&ugeStDato=" + ugeStDato, { control: "FN_getCustDesc", AjaxUpdateField: "true", cust: thisC.val() }, function (data) {
        //$("#fajl").val(data);
        $("#jobid").html(data);
        if (focusofoff == 1) {
            $("#jobid").focus();
        }
    });




}




/// Job afsluttet denne uge //
function GetJobDenneUge() {

    //var thisC = $("#FM_kontakt")


    stDato = $("#datoMan").val()
    slDato = $("#datoSon").val()
    //$.cookie('treg_sort', sortby);
    //alert(stDato)

   

    

    //if ($("#FM_easyreg").is(':checked') == true) {
    //    viseasyreg = 1
    //    $.cookie('easyreg', '1');

    //} else {
    //    viseasyreg = 0
    //    $.cookie('easyreg', '0');
    //}

   

  
    $.post("?sortby=0&stDato="+stDato+"&slDato="+slDato, { control: "FN_jobidenneuge", AjaxUpdateField: "true" }, function (data) {
        //$("#fajl").val(data);
        $("#jobidenneuge").html(data);
       
    });




}



$(document).ready(function () {


    $("select[name*=ajax]").AjaxUpdateField({ parent: "tr", subselector: "td:first > input[name=rowId]" });
    $("#incidentlist").table_sort({ items: 'tbody > tr:gt(1):has(:input[name=rowId])', IdControlNode: "td:first > :input[name=rowId]" });


    GetJobDenneUge();


    $(".dcls_1").keyup(function () {

        //alert(this.id)
        //alert($("#FM_timer_" + this.id + "_1").val())
        dagstotal(1)

    });

    $(".dcls_2").keyup(function () {

        dagstotal(2)

    });

    $(".dcls_3").keyup(function () {

        dagstotal(3)

    });

    $(".dcls_4").keyup(function () {

        dagstotal(4)

    });

    $(".dcls_5").keyup(function () {

        dagstotal(5)

    });

    $(".dcls_6").keyup(function () {

        dagstotal(6)

    });

    $(".dcls_7").keyup(function () {

        dagstotal(7)

    });



    function dagstotal(dag) {

        var timerval = 0;
        var tfltVal = 0;
        var loops = $("#loops").val();


        for (i = 1; i <= loops; i++) {
            tfltVal = jQuery.trim($("#FM_timer_" + i + "_" + dag + "").val())
            //$("#FM_timer_" + i + "_" + dag + "").val(tfltVal.replace(".", ","))
            if (tfltVal.length > 0) {
                //alert($("#FM_timer_" + i + "_1").val() + " # " + tfltVal.length)
                if (tfltVal / 1 > 0) {
                    tfltVal = tfltVal.replace(",", ".")
                    timerval = timerval + (tfltVal / 1)
                }
            }

        }

        timerval = String(Math.round(timerval * 100) / 100).replace(".", ",")
        $("#timer_" + dag).val(timerval)

        //var p = $("div:dagstotaler");
        //var position = p.position();
        //alert(position.left + " # " + position.top +" ## "+  position.bottom)
        //var top = screen.height;
        //top = top - 300
        //$("div:dagstotaler").css("top", top);
    }


    $(".a_showkom").click(function () {

        var a = $(this);
        var pos = a.position()
        var left = pos.left
        var top = pos.top
        //alert(top)
        $("#kom").css("top", top + 600);
        $("#kom").css("left", "200");
        $(".FM_kommentar").blur();
        $(".FM_kommentar").focus();
        $.scrollTo("#kom", 300, { offset: -200 });

    });


    //$("#jobid").bind('click', function () { alert($(this).find('option:selected').text()); });



    //var thisid = this.id
    //alert(thisid)

    $("#xjobid").bind('click', function () {
        if ($('.jobid :selected').size() == 0)
        { alert('not selected'); }
    });



    $("#jobid").click(function () {



        var seloptions = String($(this).val())
        //alert(seloptions);


        //var thisid = this.id
        //alert(thisid)
        var alloptions = "";
        var optionVal = "";
        $("#jobid option").each(function () {




            optionVal = String($("#jobid option[value='" + $(this).val() + "']").val())

            instr32 = seloptions.indexOf(optionVal, 0)

            //alert(instr32)
            //instr32 = 0;
            if (instr32 == -1) {
                alloptions = alloptions + ", " + $("#jobid option[value='" + $(this).val() + "']").val()
            }


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


        //alert("her")
        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(10, thisvallngt)
        thisval = thisvaltrim

        //alert(thisval)

        //showhideJob(thisval)

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


        $("#jobid option:selected").each(function () {
            if ($(this).val() == thisval) {
                //$(this).val()
                //alert($(this).val())
                $("#jobid option[value='" + $(this).val() + "']").removeAttr("selected")
            }

        });






    });


    $(".showjobbesk").click(function () {

        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(12, thisvallngt)
        thisval = thisvaltrim

        if ($("#jobbeskdiv_" + thisval).css('display') == "none") {

            $("#jobbeskdiv_" + thisval).css("display", "");
            $("#jobbeskdiv_" + thisval).css("visibility", "visible");
            $("#jobbeskdiv_" + thisval).show(4000);
            $.scrollTo(this, 300, { offset: -200 });


        } else {

            $("#jobbeskdiv_" + thisval).hide(1000);
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



    $(".a_treg").click(function () {

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
            $.scrollTo(this, 300, { offset: -200 });



        } else {

            $("#div_det_" + thisval).hide(1000);
            $.scrollTo(this, 300, { offset: -200 });



        }


    });







    $("#tildelalle").click(function () {
        alert("Vær opmærksom på, at alle felter med timer i, bliver tildelt til alle de valgte medarbejdere.");
    });

    $("#ko0chk").focus(function () {
        $.scrollTo('450px', 1000);
    });

    $("#jq_hideKpersdiv").click(function () {
        $.scrollTo('650px', 500);
    });

    // Kpers og filialer til Kørsel //
    $("#FM_sog_kpers_but").click(function () {
        GetFiliogKpers(1);
    });

    //$("#FM_sog_kpers_dist_all").click(function() {
    //    //alert("her")
    //    GetFiliogKpers();
    //});



    function GetFiliogKpers(n) {
        //alert("her")
        var strSQLKundKri;
        var thisC = $("#FM_sog_kpers_dist");
        //var visalle = 0;

        var kid = $("#FM_sog_kpers_dist_kid");

        //if (document.getElementById("FM_sog_kpers_dist_all").checked == true) {
        if ($("#FM_sog_kpers_dist_all").is(':checked') == true) {
            visalle = 1;
        } else {
            visalle = 0;
        }

        //if (document.getElementById("FM_sog_kpers_dist_more").checked == true) {
        //if ($("#FM_sog_kpers_dist_more").is(':checked') == true) {
        //if ($("FM_sog_kpers_dist_all").is(':checked') == true) {
        var xvalbegin = $("#koFltx").val()
        if (n == 1) {
            more = 1;
        } else {
            more = 0;
        }




        //$("#BtnCustDescUpdate").data("cust", thisC.val());
        $.post("?kid=" + kid.val() + "&visalle=" + visalle + "&xvalbegin=" + xvalbegin + "&more=" + more, { control: "FM_get_destinations", AjaxUpdateField: "true", cust: thisC.val() }, function (data) {
            //$("#filialerogkontakter").val(data);
            if (n == 1) {
                $("#filialerogkontakter").html("<br><hr><span style='background-color:#FFFF99;'><b>Ny søgning:</b> " + thisC.val() + "</span><br>" + data + $("#filialerogkontakter").html());

            } else {
                $("#filialerogkontakter").html("<br><hr><span style='background-color:#FFFF99;'><b>Kontakter på valgte job:</b></span><br>" + data);

            }
            //$("#koFltx").val(xval);
            document.getElementById("koFltx").value = (document.getElementById("kperfil_fundet").value / 1) * 1

        });
    }

    GetFiliogKpers(0);
    //////////////////////////////////



    /// Easyreg funktioner //
    $(".ea_kom").click(function () {

        var thisid = this.id;

        var idlngt = thisid.length;
        var idtrim = thisid.slice(5, idlngt);
        showEasyKom(idtrim);
    });

    function showEasyKom(idthis) {
        $("#kom_easy").css("visibility", "visible");
        $("#kom_easy").css("display", "");
        $.scrollTo('+=100px', 800);
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

        for (i = 1; i < 10; i++) {

            //$(".d_kom_cls_" + dagtype + "").val(EasyKom)
            preval = $("#FM_kom_" + i + "" + dagtype + "").val()
            $("#FM_kom_" + i + "" + dagtype + "").val(preval + "\n" + "\n" + EasyKom)

        }

        $.scrollTo('+=50px', 800);


    });



    //document.getElementById("kom").style.display = "";
    //document.getElementById("kom").style.visibility = "visible";
    //setKomFocus();



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

            //alert(res + " id:" + id)

            var preval = 0;
            var newval = 0;
            var oprVerdi = 0;
            var i = 1; //iRowloop

            for (i = 1; i < 500; i++) { //bør svare til maks antal EA linier


                res = Math.round(antaltimer / antallinier * 100) / 100

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


            $("#FM_ignorer_projektgrupper").attr('checked', false);



        }

        GetCustDesc();
        $("#filterkri").submit();

    });



    $("#FM_smartreg").click(function () {

        if ($("#FM_smartreg").is(':checked') == true) {

            $.cookie('smartreg', '1');

            $.cookie('easyreg', '0');
            $("#FM_easyreg").attr('checked', false);

            $.cookie('classicreg', '0');
            $("#FM_classicreg").attr('checked', false);

            $.cookie('nyeste', '0');
            $("#FM_nyeste").attr('checked', false);


            $("#FM_ignorer_projektgrupper").attr('checked', false);



        }

        GetCustDesc();
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


            $("#FM_ignorer_projektgrupper").attr('checked', false);



        }

        GetCustDesc();
        $("#filterkri").submit();

    });

    ////////////////////////////////////////////




    if ($.cookie('showakt') == "1") {


        //GetCustDesc();
        jobinfoshowhide();

        $("#FM_kontakt").change(function () {
            GetCustDesc();
        });

        $("#FM_ignorer_projektgrupper").click(function () {

            if ($("#FM_ignorer_projektgrupper").is(':checked') == true) {

                $("#FM_easyreg").attr('checked', false);
                $.cookie('easyreg', '0');

                $.cookie('smartreg', '0');
                $("#FM_smartreg").attr('checked', false);
            }

            GetCustDesc();
        });

        $("#FM_nyeste").click(function () {

            if ($("#FM_nyeste").is(':checked') == true) {

                $("#FM_easyreg").attr('checked', false);
                $.cookie('easyreg', '0');

                $.cookie('smartreg', '0');
                $("#FM_smartreg").attr('checked', false);

                $.cookie('classicreg', '0');
                $("#FM_classicreg").attr('checked', false);
            }

            GetCustDesc(1);
            showhidesorter();

        });


        $("#FM_sog_job_navn_nr").keyup(function () {
            GetCustDesc();
        });

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


        GetCustDesc();
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
        $("#pagesetadv").show(1000);
        $("#tregaktmsg1").hide("fast");

    });


    $("#pagesetadv_close").click(function () {


        $("#pagesetadv").hide("fast");


    });


    $(".afase").click(function () {
        idval = this.id
        //alert(idval)
        vzb = $(".td_" + idval + "").css("visibility")
        if (vzb == "hidden") {
            $(".td_" + idval + "").css("display", "");
            $(".td_" + idval + "").css("visibility", "visible");
            document.getElementById("faseshow_" + idval + "").value = "visible"

            $.scrollTo(this, 800, { offset: -100 });

        } else {
            $(".td_" + idval + "").css("display", "none");
            $(".td_" + idval + "").css("visibility", "hidden");
            document.getElementById("faseshow_" + idval + "").value = "hidden"

            $.scrollTo(this, 800, { offset: -400 });
            //$.scrollTo('400px', 800);
        }


    });


    function smileyshowhide() {

        $("#s2").hide(1000);

        $("#s2_td").css("background-color", "#EFF3FF");


        if ($("#s1").css('display') == "none") {

            $("#s1").css("display", "");
            $("#s1").css("visibility", "visible");
            $("#s1").show(4000);
            $("#s1_td").css("background-color", "#FFFF99");

        } else {

            $("#s1").hide(1000);
            $("#s1_td").css("background-color", "#EFf3FF");

        }

    }

    $("#sA1_k").click(function () {
        smileyshowhide()
    });

    $("#s1_k").click(function () {
        smileyshowhide()
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










});



$(window).load(function() {
    // run code
    GetCustDesc(0);
    $("#loadbar").hide(1000);
});







function showKMdailog(val, flt, chk, rk, topad) {


    komLengthVal = document.getElementById("FM_kom_" + flt + "").value
    komLength = komLengthVal.length
    //alert(komLength)

    lastKmdiv = document.getElementById("lastkmdiv").value


    koKmDialog = document.getElementById("koKmDialog").value

    //(komLength + "#" + lastKmdiv + "#" + koKmDialog)

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



            }
        }
    }



}




function koadr() {
   //alert("her")
   //xval = xval/1

   xval = document.getElementById("koFltx").value;
   flt = document.getElementById("koFlt").value;

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
}



function showmultiDiv(){
    
    if (document.getElementById("tildelalle").checked == true) {
	document.getElementById("multivmed").style.visibility = "visible"
	document.getElementById("multivmed").style.display = ""
	} else {
	document.getElementById("multivmed").style.visibility = "hidden"
	document.getElementById("multivmed").style.display = "none"
	document.getElementById("tildelalle").checked = false
	}


}





function showlukalleuger(){
	if (document.getElementById("FM_afslutuge").checked == true) {
	document.getElementById("lukalleuger").style.visibility = "visible"
	document.getElementById("lukalleuger").style.display = ""
	} else {
	document.getElementById("lukalleuger").style.visibility = "hidden"
	document.getElementById("lukalleuger").style.display = "none"
	document.getElementById("FM_alleuger").checked = false
	}
}
		
function visstempelur()
{
document.getElementById("stempelur").style.visibility = "visible"
document.getElementById("stempelur").style.display = ""
}		
		
function lukstempelur(){
document.getElementById("stempelur").style.visibility = "hidden"
document.getElementById("stempelur").style.display = "none"
}		

function lukaftalealert(){
document.getElementById("aftalealert").style.visibility = "hidden"
document.getElementById("aftalealert").style.display = "none"
document.getElementById("jobinfo").style.visibility = "visible"
document.getElementById("jobinfo").style.display = ""
}				

function lukalert(){
document.getElementById("alert").style.visibility = "hidden"
document.getElementById("alert").style.display = "none"
}	

function checkAll(field)
	{
	for (i = 0; i < field.length; i++)
		field[i].checked = true ;
	}
	
	function uncheckAll(field)
	{
	for (i = 0; i < field.length; i++)
		field[i].checked = false ;
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


function setTimerTot(nummer, dagtype, job, aktivitet) {
	
	//piletaster
	if (window.event.keyCode == 37){ 
	} else {
		if (window.event.keyCode == 39){
		//alert(window.event.keyCode)
		} else {
		
	
		var varValue = 0;
		var varValue_total = 0;
		
		document.getElementById("FM_"+ dagtype +"_opr_" + nummer + "").value = document.getElementById("FM_"+ dagtype +"_opr_" + nummer + "").value.replace(",",".")
		oprVerdi = (document.getElementById("FM_"+ dagtype +"_opr_" + nummer + "").value / 1);
		
		document.getElementById("Timer_"+ dagtype +"_" + nummer + "").value = document.getElementById("Timer_"+ dagtype +"_" + nummer + "").value.replace(",",".")
		varValue = (document.getElementById("Timer_"+ dagtype +"_" + nummer + "").value / 1);
		document.getElementById("Timer_"+ dagtype +"_" + nummer + "").value = document.getElementById("Timer_"+ dagtype +"_" + nummer + "").value.replace(".",",")
		
		document.getElementById("FM_"+ dagtype +"_total").value = document.getElementById("FM_"+ dagtype +"_total").value.replace(",",".")
		varValue_total = (document.getElementById("FM_"+ dagtype +"_total").value / 1); 
	 	
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
			document.getElementById("FM_"+ dagtype +"_total").value = parseFloat(varTotal_dag_beg);
			document.getElementById("FM_"+ dagtype +"_opr_" + nummer + "").value = parseFloat(varValue);
			sonValue = document.getElementById("FM_son_total").value.replace(",",".")/1;
			manValue = document.getElementById("FM_man_total").value.replace(",",".")/1;
			tirValue = document.getElementById("FM_tir_total").value.replace(",",".")/1;
			onsValue = document.getElementById("FM_ons_total").value.replace(",",".")/1;
			torValue = document.getElementById("FM_tor_total").value.replace(",",".")/1;
			freValue = document.getElementById("FM_fre_total").value.replace(",",".")/1;
			lorValue = document.getElementById("FM_lor_total").value.replace(",",".")/1;
			document.getElementById("FM_week_total").value = parseFloat(sonValue + manValue + tirValue + onsValue + torValue + freValue + lorValue) 
			document.getElementById("FM_week_total").value = document.getElementById("FM_week_total").value.replace(".",",")
			document.getElementById("FM_"+ dagtype +"_total").value = document.getElementById("FM_"+ dagtype +"_total").value.replace(".",",")
			//}
		//}
	}} // piletaster
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

function tjektimer(dagtype, nummer){
	if (!validZip(document.getElementById("Timer_"+ dagtype +"_" + nummer + "").value)){
	alert("Der er angivet et ugyldigt tegn.")
	document.getElementById("Timer_"+ dagtype +"_" + nummer + "").value = oprVerdi;
	document.getElementById("Timer_"+ dagtype +"_" + nummer + "").focus()
	document.getElementById("Timer_"+ dagtype +"_" + nummer + "").select()
	return false
	}
return true
}

function tjekkm(dagtype, nummer){
	oprVerdi = (document.getElementById("FM_"+ dagtype +"_opr_" + nummer + "").value / 1);
	if (!validZip(document.getElementById("Timer_"+ dagtype +"_" + nummer + "").value)){
	alert("Der er angivet et ugyldigt tegn.")
	document.getElementById("Timer_"+ dagtype +"_" + nummer + "").value = oprVerdi;
	document.getElementById("Timer_"+ dagtype +"_" + nummer + "").focus()
	document.getElementById("Timer_"+ dagtype +"_" + nummer + "").select()
	return false
	}
return true
}

	
	
	//Aktiviteter expand
 	if (document.images){
		plus = new Image(200, 200);
		plus.src = "ill/plus.gif";
		minus = new Image(200, 200);
		minus.src = "ill/minus2.gif";
		}

		function expand(de){
		//alert(navigator.appVersion.indexOf("MSIE"))
		    if (navigator.appVersion.indexOf("MSIE")==-1){
			//alert("hej")
				//if (document.getElementById("Menu" + de)){
				if (document.getElementById("Menu" + de).style.display == "none"){
					document.getElementById("Menu" + de).style.display = "";
					document.images["Menub" + de].src = minus.src;
				}else{
					document.getElementById("Menu" + de).style.display = "none";
					document.images["Menub" + de].src = plus.src;
				} // else
			}else{
				//alert("pc")
				if (document.all("Menu" + de)){
				if (document.all("Menu" + de).style.display == "none"){
					document.all("Menu" + de).style.display = "";
					document.images["Menub" + de].src = minus.src;
				}else{
					document.all("Menu" + de).style.display = "none";
					document.images["Menub" + de].src = plus.src;
				} //else
			} //else
		} //function
	} //images
	
	
	
	
	function antalchar(){
	//if (navigator.appVersion.indexOf("MSIE")==-1){
		thisstring = document.getElementById("FM_kom").value.length
		if (thisstring > 255){
			alert("Der er ikke tilladt mere end 255 karakterer i en kommentar! \n Du har brugt: "+ thisstring +"")
		}else{
			document.getElementById("antch").value = thisstring;
		}
	}
	

 function antalakt(jid){
 newval = document.all["FM_akt_" + jid].value;
 document.all["FM_antal_akt_" + jid].value = newval;
 }

 function closetomanyjob() {
 document.all["tomanyjob"].style.visibility = "hidden";
 }
 
  ///  ressource timer tildelt ///////
 	function showrestimer(){
	document.getElementById("ressourcetimer").style.display = "";
	document.getElementById("ressourcetimer").style.visibility = "visible";
	}
	
	function hiderestimer(){
	document.getElementById("ressourcetimer").style.display = "none";
	document.getElementById("ressourcetimer").style.visibility = "hidden";
	}
	
 ///  viser indtastede timer på den enkelte dag ///////
 	function showtimedetail(thisnameid){
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
	
	function hidetimedetail(thisid){
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
		
		
		
	    //Kommentar expand
		function expandkomm(nummer, dagtype) {
		    //xpos = 50 //window.event.clientX - 200
		    //ypos = -300 //window.event.clientY - 20


		    
		        thisnum = nummer / 1

		        if (thisnum >= 9) {
		            document.getElementById("clskom").href = "#anc_" + (thisnum - 7) + "_" + dagtype + ""
		        } else {
		            document.getElementById("clskom").href = "#anc_s0"
		        }



		        thisval = document.getElementById("FM_kom_" + nummer + "" + dagtype + "").value;


		        if (document.getElementById("sdskKomm").value == "") {
		            thisval = thisval
		        } else {
		            thisval = thisval + document.getElementById("sdskKomm").value
		        }

		        thisvaloff = document.getElementById("FM_off_" + nummer + "" + dagtype + "").value
		        thisbopalvaloff = document.getElementById("FM_bopal_" + nummer + "" + dagtype + "").value

		        document.getElementById("rowcounter").value = nummer;
		        document.getElementById("daytype").value = dagtype;

		        //alert(document.body.scrollTop)
		        document.getElementById("kom").style.display = "";
		        document.getElementById("kom").style.visibility = "visible";


		        document.getElementById("FM_kommentar").value = thisval;
		        //document.getElementById("FM_kommentar").value = "her der"

		        //$("#FM_kommentar").val("her komer jeg");
		        //$(".komfocus").focus();

		        document.getElementById("FM_off").value = thisvaloff;
		        document.getElementById("FM_bopalstur").value = thisbopalvaloff;
		        document.getElementById("kom").style.top = "550px"
		        //document.body.scrollTop - 200;
		        document.getElementById("kom").style.left = "60px";
		        thisstring = document.getElementById("FM_kommentar").value.length;
		        document.getElementById("antch").value = thisstring;
		        //window.frames['a'].document.getElementById("FM_kommentar").focus();


		        //document.getElementById("FM_kommentar").scrollIntoView(true);
		        //window.FM_kommentar.focus;
		        //document.getElementById("FM_kommentar").select();
		        //document.getElementById("FM_kommentar").focus();
		        //setKomFocus();
		    	
			}
			
			function closekomm(){
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
					
					if (lasttd != "0"){
						if (lasttddaytype >= 6) {
						document.getElementById("td_"+lasttd+"").style.background = "#dcdcdf";
						} else {
						document.getElementById("td_"+lasttd+"").style.background = "#ffffe1";
						}
					}
					
					document.getElementById("td_"+thisrow+""+thisdaytype+"").style.background = "#FFff99"; 
					document.getElementById("lasttd").value = ""+thisrow+""+thisdaytype+"";
					document.getElementById("lasttddaytype").value = ""+thisdaytype+"";
					
					if (thisstring > 1255){
					alert("Der er ikke tilladt mere end 1255 karakterer i en kommentar! \n Du har brugt: "+ thisstring +"")
					}else{
					document.getElementById("FM_kom_"+thisrow+""+thisdaytype+"").value = thisvalue;
					document.getElementById("FM_off_"+thisrow+""+thisdaytype+"").value = thisoffentlig;
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
			
			function antalchar(){
			//if (navigator.appVersion.indexOf("MSIE")==-1){
			thisstring = document.getElementById("FM_kommentar").value.length
				if (thisstring > 1255){
					alert("Der er ikke tilladt mere end 1255 karakterer i en kommentar! \n Du har brugt: "+ thisstring +"")
				}else{
					document.getElementById("antch").value = thisstring;
				}
			}
			
			function popUp(URL,width,height,left,top) {
				window.open(URL, 'navn', 'left='+left+',top='+top+',toolbar=0,scrollbars=1,location=0,statusbar=1,menubar=0,resizable=1,width=' + width + ',height=' + height + '');
			}
			
			
			function showtimeregdetail(divid){
			
			
			for (i = 0; i < 8; i++) {
			document.getElementById("tr_"+i+"").style.display = "none";
			document.getElementById("tr_"+i+"").style.visibility = "hidden";
			}
			
			//document.getElementById("smiley").style.display = "none";
			//document.getElementById("smiley").style.visibility = "hidden";
			
			
			document.getElementById("tr_"+ divid +"").style.display = "";
			document.getElementById("tr_"+ divid +"").style.visibility = "visible";
			}
			
			
			function visSmileystatus(){
			document.getElementById("smileydiv").style.display = "";
			document.getElementById("smileydiv").style.visibility = "visible";
			}
			
			function gemSmileystatus(){
			document.getElementById("smileydiv").style.display = "none";
			document.getElementById("smileydiv").style.visibility = "hidden";
			}
			
			
			
			function markerjob(tdid){
			lastjob = document.getElementById("lastjob").value
			
			if (lastjob != 0){
			document.getElementById("tr_"+lastjob+"").style.background = "#FFffff";
			}
			
			document.getElementById("tr_"+tdid+"").style.background = "#FFff99";
			document.getElementById("lastjob").value = tdid
			}
			
			
			function markerfelt(feltid, thisdaytype){
			// Markerer felt
			lasttd = document.getElementById("lasttd").value
			lasttddaytype = document.getElementById("lasttddaytype").value
			
			if (lasttd != "0"){
				if (lasttddaytype >= 6) {
				document.getElementById("td_"+lasttd+"").style.background = "#dcdcd0";
				} else {
				document.getElementById("td_"+lasttd+"").style.background = "#ffffe1";
				}
			}
			
			document.getElementById("td_"+feltid+"").style.background = "#FFff99"; 
			document.getElementById("lasttd").value = ""+feltid+"";
			document.getElementById("lasttddaytype").value = ""+thisdaytype+"";
			}
			
			function renssog(){
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
			
			function curserType(tis){
			document.getElementById(tis).style.cursor='hand'
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
			document.getElementById("korseldiv").style.display = "none";
			document.getElementById("korseldiv").style.visibility = "hidden";
			document.getElementById("kpersdiv").style.display = "none";
			document.getElementById("kpersdiv").style.visibility = "hidden";
			}
			
	function showdiv(div) {
	//alert(div)
	lastopendiv = document.getElementById("lastopendiv").value
	
	document.getElementById(lastopendiv).style.display = "none";
	document.getElementById(lastopendiv).style.visibility = "hidden";
	
	document.getElementById("knap_"+lastopendiv).style.background = "#FFFFFF";
	document.getElementById("knap_"+lastopendiv).style.border = 0;
	
	
	document.getElementById(div).style.display = "";
	document.getElementById(div).style.visibility = "visible";
	
	document.getElementById("knap_"+div).style.background = "#FFFF99";
	document.getElementById("knap_"+div).style.border = "1px orange solid";
	
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
    
    
    function ShowHideAktBesk(aktid) {
    if (document.getElementById(aktid).style.height == "100px") {
    document.getElementById(aktid).style.height = "14px";
    document.getElementById(aktid).style.overflow = "hidden";
    } else {
    document.getElementById(aktid).style.height = "100px";
    document.getElementById(aktid).style.overflow = "auto";
	}
    }