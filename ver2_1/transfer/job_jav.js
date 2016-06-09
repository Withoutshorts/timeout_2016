// JScript File








$(document).ready(function () {

//alert("Så er siden klar..")

    // Kundesøg //
    function getKundelisten() {

        var visalle = 1 //bruges ikke
        var sog_val = $("#kunde_sog_step1")

        //alert("her: " + sog_val.val())
        if (sog_val.val() != "") {

        $.post("?jq_sogkunde=" + visalle, { control: "FN_getKundelisten", AjaxUpdateField: "true", cust: sog_val.val() }, function (data) {
            //$("#FM_modtageradr").val(data);

            //var idlngt = sog_val.val()
            //alert(idlngt.lenght)
            //if (idlngt.lenght == 1) {
            //    alert("Du er ved at foretege en søgning..")
            //}
            //alert("Der søges...")
            $("#jq_kunde_sel").html(data);
            $("#jq_kunde_sel").focus();

            //$("#jobid").html(data);

        });

        }


    }



    $("#kunde_sog_step1").keyup(function () {
        getKundelisten();
    });

    $("#kunde_sog_step1_but").click(function () {
        getKundelisten();
    });


    //alert("her")


    // This will hold our timer

    //var myTimer = {};

    // Set the timer for 2 seconds
    //myTimer = $.timer(12000, function () {

    // Display hello message when timer goes off
    //alert('A Delayed Hello!');
    //   $('#huskbudget').hide(1000)

    // Optional function to call when timer is canceled


    //});



    //$("#alukhuskbudget").click(function () {
    //    $('#huskbudget').hide(1000)
    //});


    /// Hvis info iom fastpris //
    function fn_angfp() {

        //alert("her" + $("#angfp0").is(':checked'))
        if ($("#angfp0").is(':checked') == true) {
            $("#div_angfp").css("display", "");
            $("#div_angfp").css("visibility", "visible");
            $("#div_angfp").show("fast");


        } else {

            $("#div_angfp").hide("fast");


        }

    }


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


    $("#angfp0").click(function () {
        //alert("her")
        fn_angfp();
    });

    $("#angfp1").click(function () {
        fn_angfp();
    });


    fn_angfp();


    // Aktlisten //
    function getAktlisten() {
        if ($("#jq_vispasogluk").is(':checked') == true) {
            var visalle = 1
        } else {
            var visalle = 0
        }

        var jid_val = $("#jq_jobid")

        $.post("?jq_visalle=" + visalle, { control: "FN_getAktlisten", AjaxUpdateField: "true", cust: jid_val.val() }, function (data) {
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
            $("#div_tild_forr").css("display", "");
            $("#div_tild_forr").css("visibility", "visible");
            $("#div_tild_forr").show("fast");
            $.scrollTo('1400px', 400);

        } else {

            $("#div_tild_forr").hide("fast");
            $.scrollTo('1000px', 400);

        }

    });




    $("#a_tild_ava").click(function () {

        if ($("#div_tild_ava").css('display') == "none") {
            $("#div_tild_ava").css("display", "");
            $("#div_tild_ava").css("visibility", "visible");
            $("#div_tild_ava").show("fast");
            $.scrollTo('1400px', 400);

        } else {

            $("#div_tild_ava").hide("fast");
            $.scrollTo('1000px', 400);

        }

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



    $("#tilfoj_ulevgrp").click(function () {

        //alert("her")
        if ($(".jq_ulevgrp").css('display') == "none") {

            $(".jq_ulevgrp").css("display", "");
            $(".jq_ulevgrp").css("visibility", "visible");
            $(".jq_ulevgrp").show("fast");
            $.scrollTo('-=0px', 400);

        } else {

            $(".jq_ulevgrp").hide("fast");
            $.scrollTo('-=100px', 1500);
        }
    });

    $("#stgrp_tilfoj").click(function () {
        if ($("#jq_stamaktgrp").css('display') == "none") {

            $("#jq_stamaktgrp").css("display", "");
            $("#jq_stamaktgrp").css("visibility", "visible");
            $("#jq_stamaktgrp").show("fast");


            //$('#jq_stamaktgrp').fadeIn('slow', function () {
            // Animation complete
            //});

            //$.scrollTo($('#jq_stamaktgrp'), 2000);
            //$.scrollTo('-=100px', 4000);

        } else {

            $("#jq_stamaktgrp").hide("fast");
            //$.scrollTo('-=100px', 1500);
        }
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

    getAktlisten()

    $("select[name*=ajax]").AjaxUpdateField({ parent: "tr", subselector: "td:first > input[name=rowId]" });
    $("#incidentlist").table_sort({ items: 'tbody > tr:gt(1):has(:input[name=rowId])', IdControlNode: "td:first > :input[name=rowId]" });


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
        var idtrim = thisid.slice(8, idlngt)
        var faseval = this.value
        //alert(faseval)
        $(".a_fs_" + idtrim).val(faseval)
    });




    function showhideakt_faser_pb(id, a) {

        var lastopen = $("#lastopen_" + a).val()
        $(".akt_pb_" + lastopen + "_" + a).hide(1000);

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

                $("#lastopen_ulev").val(id)

            }
        }


    }



    $("#jq_seljugrp_id").click(function () {

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


    $("#FM_navn").keyup(function () {
        $("#pb_jobnavn").text($("#FM_navn").val())
    });

    $("#pb_jobnavn").text($("#FM_navn").val())


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
        $("#ulevpris_" + idtrim).val('0,00');
        $("#ulevfaktor_" + idtrim).val('1,00');
        $("#ulevbelob_" + idtrim).val('0,00');


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
        if ($("#div_internnote").css('display') == "none") {

            $("#div_internnote").css("display", "");
            $("#div_internnote").css("visibility", "visible");
            $("#div_internnote").show(4000);
            $.scrollTo('400px', 1500);

        } else {

            $("#div_internnote").hide(1000);
            $.scrollTo('300px', 1500);

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



    //function xxshowhidefastpris() {
    //    if ($("#FM_fastpris1").is(':checked') == true) {

    //        $("#fastprisDiv").css("display", "");
    //        $("#fastprisDiv").css("visibility", "visible");
    //        $("#fastprisDiv").show(4000);

    //    } else {

    //        $("#fastprisDiv").hide(1000);

    //    }
    //}


    //function hidefastpris() {
    //    $("#fastprisDiv").hide(1000);
    //}

    //$("#FM_fastpris1").click(function () {
    //    showhidefastpris();
    //});

    //$("#FM_fastpris2").click(function () {
    //    hidefastpris();
    //});

    //$("#fp_job_or_akt_Luk").click(function () {
    //    hidefastpris();
    //});


    //showhidefastpris();




    //$(".jobmenu").mouseout(function() {
    //var thisid = this.id
    //$("#" + thisid).css("background-color", "#8CAAe6");
    //});

    //$(".jobmenu").mouseover(function() {
    //    //alert("her" + this.id)
    //    var thisid = this.id
    //    $("#" + thisid).css("background-color", "#5582d2");
    //});



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

        fn_angfp();

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

    $("#vistp_A").click(function () {
        //alert("her")
        tpriserfunc();

    });

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
        $("#stgrpfs_" + thisid + "").val('')
        var thisval = $("#" + thisid + " option:selected").val()
        //alert(thisval)
        showhideakt_faser_pb(thisval, thisid);

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









});



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
	
	function overfortp(mid,t){
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
	
	


	function showulevlinier(){

	    vislinier = document.getElementById("antalulev").value
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
	
	//alert("her set toal: " + document.getElementById("FM_interntbelob").value)
	
	if (document.getElementById("FM_interntbelob").value != "") {
	totalbelob = document.getElementById("FM_interntbelob").value.replace(",",".")/1
	} else {
	totalbelob = 0
	}
	
	totalomk = 0


	antalUlevlinier = document.getElementById("ulevopen").value
	for (i = 1; i <= antalUlevlinier; i++) {

	    //alert(document.getElementById("ulevbelob_" + i + "").value)
	    if (document.getElementById("ulevbelob_"+i+"").value != "") {
	    totalbelob = totalbelob/1 + document.getElementById("ulevbelob_"+i+"").value.replace(",",".")/1;
	    totalomk = totalomk/1 + document.getElementById("ulevpris_"+i+"").value.replace(",",".")/1;
	    
	    //alert(i +" "+ document.getElementById("ulevbelob_"+i+"").value)
	    }
	}
	
	totalomk = Math.round(totalomk*100)/ 100
	
    internomk = document.getElementById("FM_interntomkost").value.replace(",",".")
    internomk = Math.round(internomk * 100) / 100
	
    udgifter = (totalomk + internomk)
    udgifter = Math.round(udgifter*100)/ 100

    totalbelob = Math.round(totalbelob*100)/ 100
	
	db = 100 - (((internomk+totalomk) / totalbelob) * 100)
	db = Math.round(db*100)/ 100
	

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
	

	
	bruttofj = Math.round(bruttofj * 100) / 100
	bruttofj = String(bruttofj)
	bruttofj = bruttofj.replace(".", ",")
	document.getElementById("SP_bruttofortj").innerHTML = "<b>" + bruttofj + "</b>"
	document.getElementById("FM_bruttofortj").value = bruttofj 

  
    
    totalbelob = String(totalbelob)
	totalbelob = totalbelob.replace(".", ",")
	document.getElementById("SP_budget").innerHTML = "<b>" + totalbelob + "</b>"
	document.getElementById("FM_budget").value = totalbelob

    db = String(db)
	db = db.replace(".", ",")
	document.getElementById("SP_db").innerHTML = "<b>" + db + "</b>"
	document.getElementById("FM_db").value = db


	/// Diff timer og Diff Sum
	var varBudgetTimer = document.getElementById("FM_budgettimer").value.replace(".", "")
	varBudgetTimer = varBudgetTimer.replace(",", ".")
	varBudgetTimer = Math.round(varBudgetTimer * 100) / 100

	varTotTim = document.getElementById("FM_fasertimertot").value.replace(".", "")
	varTotTim = varTotTim.replace(",", ".")

	var atkTimTot = Math.round(varTotTim * 100) / 100
	var diff_timer = (atkTimTot - (varBudgetTimer))

	//alert(varBudgetTimer + " # " + atkTimTot)
	diff_timer = String(Math.round(diff_timer * 100) / 100).replace(".", ",")

	//diff_timer = String(Math.round(diff_timer * 100) / 100).replace(".", ",")
	document.getElementById("diff_timer").innerHTML = "<b>" + diff_timer + "</b>"
	document.getElementById("FM_diff_timer").value = diff_timer

	var varBudgetSum = document.getElementById("FM_interntbelob").value.replace(".", "")
	varBudgetSum = varBudgetSum.replace(",", ".")
	varBudgetSum = Math.round(varBudgetSum * 100) / 100

	varTotSum = document.getElementById("FM_fasersumtot").value.replace(".", "")
	varTotSum = varTotSum.replace(",", ".")
	var atkSumTot = Math.round(varTotSum * 100) / 100
	var diff_sum = (varTotSum - (varBudgetSum))

	//alert(varBudgetSum + " # " + varTotSum)

	diff_sum = String(Math.round(diff_sum * 100) / 100).replace(".", ",")

	document.getElementById("diff_sum").innerHTML = "<b>" + diff_sum + "</b>"
	document.getElementById("FM_diff_sum").value = diff_sum

	if (diff_timer != "0") {

	    document.getElementById("diff_timer_ok").style.visibility = "hidden";
	    document.getElementById("diff_timer_ok").style.display = "none";

	} else {

	    document.getElementById("diff_timer_ok").style.visibility = "visible";
	    document.getElementById("diff_timer_ok").style.display = "";

	}


	if (diff_sum != "0") {

	    document.getElementById("diff_sum_ok").style.visibility = "hidden";
	    document.getElementById("diff_sum_ok").style.display = "none";

	} else {

	    document.getElementById("diff_sum_ok").style.visibility = "visible";
	    document.getElementById("diff_sum_ok").style.display = "";

	}

	
	}




	function beregnintbelob() {

    
	timer = document.getElementById("FM_budgettimer").value.replace(".", "")
	timer = timer.replace(",", ".")
	timer = Math.round(timer*100)/ 100

	faktor = document.getElementById("FM_intfaktor").value.replace(".", "")
	faktor = faktor.replace(",", ".")
	faktor = Math.round(faktor*100)/ 100

	gnsinttpris = document.getElementById("FM_gnsinttpris").value.replace(".", "")
	gnsinttpris = gnsinttpris.replace(",", ".")
	gnsinttpris = Math.round(gnsinttpris*100)/ 100
	
	internOmkost = (timer * gnsinttpris)
	internOmkost = Math.round(internOmkost*100)/ 100
	
	internOmkost = String(internOmkost)
	internOmkost = internOmkost.replace(".",",")
	
	interntBelob = (timer * gnsinttpris * faktor)
	interntBelob = Math.round(interntBelob*100)/ 100
	
	interntBelob = String(interntBelob)
	interntBelob = interntBelob.replace(".",",")
	
	document.getElementById("FM_interntbelob").value = interntBelob
	document.getElementById("FM_interntomkost").value = internOmkost

	targettpris = (gnsinttpris/1 * faktor/1)
	targettpris = Math.round(targettpris * 100) / 100
	document.getElementById("pb_tg_timepris").value = targettpris
	
	settotalbelob()
	
	}
	
	
	function beregninttp() {
	timer = document.getElementById("FM_budgettimer").value.replace(".", "")
	timer = timer.replace(",", ".")
	timer = Math.round(timer*100)/ 100

	faktor = document.getElementById("FM_intfaktor").value.replace(".", "")
	faktor = faktor.replace(",", ".")
	faktor = Math.round(faktor*100)/ 100

	interntBelob = document.getElementById("FM_interntbelob").value.replace(".", "")
	interntBelob = interntBelob.replace(",", ".")
	interntBelob = Math.round(interntBelob*100)/ 100
	
   
    if (timer != 0 && faktor != 0) {
	gnsinttpris = (interntBelob/(timer * faktor))
    } else {
    gnsinttpris = document.getElementById("FM_gnsinttpris").value.replace(".", "")
    gnsinttpris = gnsinttpris.replace(",", ".")
    }
	gnsinttpris = Math.round(gnsinttpris*100)/ 100

	targettpris = (gnsinttpris / 1 * faktor / 1)
	targettpris = Math.round(targettpris * 100) / 100
	document.getElementById("pb_tg_timepris").value = targettpris

	internOmkost = (timer * gnsinttpris)
	internOmkost = Math.round(internOmkost*100)/ 100
	
	internOmkost = String(internOmkost)
	internOmkost = internOmkost.replace(".",",")
	
	gnsinttpris = String(gnsinttpris)
	gnsinttpris = gnsinttpris.replace(".",",")
	
	document.getElementById("FM_gnsinttpris").value = gnsinttpris
	document.getElementById("FM_interntomkost").value = internOmkost

	

	settotalbelob()
	
	}
	
	
	
	
	function beregnulevbelob(u) {
	
	ufaktor = document.getElementById("ulevfaktor_"+u+"").value.replace(",",".")
	ufaktor = Math.round(ufaktor*100)/ 100
	
	upris = document.getElementById("ulevpris_"+u+"").value.replace(",",".")
	upris = Math.round(upris*100)/ 100
	
	uBelob = (upris * ufaktor)
	uBelob = Math.round(uBelob*100)/ 100
	
	uBelob = String(uBelob)
	uBelob = uBelob.replace(".",",")
	
	document.getElementById("ulevbelob_"+u+"").value = uBelob
	
	
	settotalbelob()
	
	}
	
	function beregnulevipris(u) {
	
	ufaktor = document.getElementById("ulevfaktor_"+u+"").value.replace(",",".")
	ufaktor = Math.round(ufaktor*100)/ 100
	
	uBelob = document.getElementById("ulevbelob_"+u+"").value.replace(",",".")
	uBelob = Math.round(uBelob*100)/ 100
	
	upris =  (uBelob / ufaktor)
	upris = Math.round(upris*100)/ 100
	upris = String(upris)
	upris = upris.replace(".",",")
	
	document.getElementById("ulevpris_"+u+"").value = upris
	
	
	settotalbelob()
	
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