
$(window).load(function () {

    // run code
    $("#loadbar").hide('fast');
  
});



$(document).ready(function () {

    //$("select[name*=ajax]").AjaxUpdateField({ parent: "tr", subselector: "td:first > input[name=rowId]" });
    // $("#incidentlist").table_sort({ items: 'tbody > tr:gt(1):has(:input[name=rowId])', IdControlNode: "td:first > :input[name=rowId]" });


    //$.cookie("showavanceret" '1');


    if ($.cookie("showavanceret") == '1') {
        $(".dv_avanceret_1").css("display", "");
        $(".dv_avanceret_1").css("visibility", "visible");
        $(".dv_avanceret_1").show("fast");
    }


    //alert($.browser.mozilla +" && "+ parseFloat($.browser.version))


    $(".aa_job_komm").mouseover(function () {
      
        $(this).css('cursor', 'pointer');
    
    });

    $("#sp_avanceret_1").mouseover(function () {

        $(this).css('cursor', 'pointer');

    });


    $("#sp_avanceret_1").click(function () {

    if ($(".dv_avanceret_1").css('display') == "none") {
        $(".dv_avanceret_1").css("display", "");
        $(".dv_avanceret_1").css("visibility", "visible");
        $(".dv_avanceret_1").show("fast");

        $.cookie("showavanceret",'1');
        
    } else {

        $(".dv_avanceret_1").hide("fast");
        $.cookie("showavanceret", '0');
    }
 

    });

  

    $(".aa_job_komm_ryd").mouseover(function () {

        $(this).css('cursor', 'pointer');

    });

    
   

    $("#FM_sorter").change(function () {

       

        if ($("#FM_sorter").val() == "7") {
            $("#dv_projektgrp").css("display", "");
            $("#dv_projektgrp").css("visibility", "visible");
            $("#dv_projektgrp").show("fast");
            $.scrollTo('600px', 1000);
            //$.scrollTo($('#dv_projektgrp'), -300, 400);

        } else {

            $("#dv_projektgrp").hide("fast");
            //$.scrollTo('1000px', 400);
            //$.scrollTo('-=100px', 1500);

            $("#joblist_filter").trigger("submit");

        }

    });

    $("#progrp_all").change(function () {
       
        if ($("#progrp_all").is(':checked') == true) {
            $(".FM_sorter_visprogrp").attr("checked", "checked");
        } else {
            $(".FM_sorter_visprogrp").removeAttr("checked");
        }
    });
    

    


    $(".aa_job_komm").click(function () {

        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(12, thisvallngt)
        thisval = thisvaltrim

        jobid = $("#jobid_" + thisval).val()

        //alert(thisval)
        var session_user = $("#FM_session_user").val()
        var dt = $("#FM_now").val()
        var jobkom = $("#FM_job_komm_" + thisval).val()

        if (jobkom.length >= 100) {
            alert("Tweet er over 100 karakterer lang! \nDen er på " + jobkom.length + " karakterer.")
            return false
        }

        //alert(jobkom)

        jobkom = jQuery.trim(jobkom)

        //alert(jobkom.length)

        if (jobkom != "Job tweet..(åben for alle)" && jobkom.length > 0) {

            $.post("?jobid=" + jobid + "&jobkom=" + jobkom, { control: "FN_updatejobkom", AjaxUpdateField: "true" }, function (data) {

                //$("#aa_job_komm_" + thisval).text('  Ok!')
                //$("#aa_job_komm_" + thisval).css("color", "#5582d2");

                $("#FM_job_komm_" + thisval).css("color", "#999999");
                $("#FM_job_komm_" + thisval).val('Job tweet..(åben for alle)');

                var forTxt = $("#dv_job_komm_" + thisval).html()
                //$("#dv_job_komm_" + thisval).html(jobkom + "<br>" + forTxt);
                $("#dv_job_komm_" + thisval).html("<span style='color:#999999;'>" + dt + ", " + session_user + ":</span> " + jobkom + "<br>" + forTxt);

            });


        }
    });

    $(".aa_job_komm_ryd").click(function () {

        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(16, thisvallngt)
        thisval = thisvaltrim

        jobid = $("#jobid_" + thisval).val()
      

      
        $.post("?jobid=" + jobid, { control: "FN_updatejobkom_ryd", AjaxUpdateField: "true" }, function (data) {

              
                //$("#dv_job_komm_" + thisval).html(jobkom + "<br>" + forTxt);
                $("#dv_job_komm_" + thisval).html('');

            });


    });


    $(".s_jobstatus").change(function () {


        //alert("fer")

        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(13, thisvallngt)
        thisval = thisvaltrim

        jobid = $("#jobid_" + thisval).val()

        jobstatus = $(this).val()

       

        $.post("?jobid=" + jobid + "&jobstatus=" + jobstatus, { control: "FN_updatejobstatus", AjaxUpdateField: "true" }, function (data) {
        });

        $("#sp_stopd_" + thisval).css("visibility", "visible")


        setTimeout(function () {
            // Do something after 5 seconds

            $("#sp_stopd_" + thisval).css("visibility", "hidden")
        }, 1000);

    });


    $(".s_prio").keyup(function () {


        //alert("fer")

        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(10, thisvallngt)
        thisval = thisvaltrim

        jobid = $("#jobid_" + thisval).val()

        jobprioitet = $(this).val()



        $.post("?jobid=" + jobid + "&jobprioitet=" + jobprioitet, { control: "FN_updatejobprio", AjaxUpdateField: "true" }, function (data) {
        });

        $("#sp_propd_" + thisval).css("visibility", "visible")


        setTimeout(function () {
            // Do something after 5 seconds

            $("#sp_propd_" + thisval).css("visibility", "hidden")
        }, 1000);

    });


    $(".s_forkalk").keyup(function () { 

        //alert("forkalktimer");

        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(11, thisvallngt)
        thisval = thisvaltrim
        
        jobid = $("#jobid_" + thisval).val()

        //alert(jobid)
     
        jobforklak = $(this).val()
        //alert(jobforklak)

        $.post("?jobid=" + jobid + "&budgettimer=" + jobforklak, { control: "FN_updatejovforkalktimer", AjaxUpdateField: "true" }, function (data) {
        });

        $("#sp_forkalk_" + thisval).css("visibility", "visible")


        setTimeout(function () {
            // Do something after 5 seconds

            $("#sp_forkalk_" + thisval).css("visibility", "hidden")
        }, 1000);

    });


    $(".s_brutoms").keyup(function () {

        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(11, thisvallngt)
        thisval = thisvaltrim

        jobid = $("#jobid_" + thisval).val()
        //alert(jobid)
        jobbruttooms = $(this).val()



        //alert(jobbruttooms)

        $.post("?jobid=" + jobid + "&jo_bruttooms=" + jobbruttooms, { control: "FN_updatejovbrutoms", AjaxUpdateField: "true" }, function (data) {
        });


        $("#sp_forkalk_" + thisval).css("visibility", "visible")


        setTimeout(function () {
            // Do something after 5 seconds

            $("#sp_forkalk_" + thisval).css("visibility", "hidden")
        }, 1000);


    });


    $(".s_jobdato").change(function () {


        //alert("fer")

        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(13, thisvallngt)
        thisval = thisvaltrim

        jobid = $("#jobid_" + thisval).val()

        jobstDato = $("#FM_start_aar_" + thisval).val() + "-" + $("#FM_start_mrd_" + thisval).val() + "-" + $("#FM_start_dag_" + thisval).val()
        jobslDato = $("#FM_slutt_aar_" + thisval).val() + "-" + $("#FM_slutt_mrd_" + thisval).val() + "-" + $("#FM_slutt_dag_" + thisval).val()

        //alert(jobslDato)


        $.post("?jobid=" + jobid + "&jobstDato=" + jobstDato + "&jobslDato=" + jobslDato, { control: "FN_updatejobdato", AjaxUpdateField: "true" }, function (data) {
        });

        $("#sp_dtopd_" + thisval).css("visibility", "visible")
        

        setTimeout(function () {
            // Do something after 5 seconds
          
            $("#sp_dtopd_" + thisval).css("visibility", "hidden")
        }, 1000);

    });


    $(".rest").click(function () {


        opdrest(this.id)

    });

    $(".stade").change(function () {

     
        opdrest(this.id)

    });

   

    function opdrest(thisid){
        
        //alert("fer")

        var thisid = thisid
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(15, thisvallngt)
        thisval = thisvaltrim

        jobid = $("#jobid_" + thisval).val()

        rest = $("#FM_restestimat_" + thisval).val() 
        timerproc = $("#FM_stade_ti_pr_" + thisval).val()
        //alert(rest + " timerproc:" + timerproc)


        $.post("?jobid=" + jobid + "&rest=" + rest + "&timerproc=" + timerproc, { control: "FN_updateest", AjaxUpdateField: "true" }, function (data) {
        });


        $("#bn_restestimat_" + thisval).css("background-color", "#FFFFFF")
        $("#bn_restestimat_" + thisval).css("color", "green")
        $("#bn_restestimat_" + thisval).html('<i>V</i>')
        //$("#sp_esopd_" + thisval).css("visibility", "visible")


        setTimeout(function () {
            // Do something after 5 seconds

            //$("#sp_esopd_" + thisval).css("visibility", "hidden")
            //$("#bn_restestimat_" + thisval).css("visibility", "visible")

            $("#bn_restestimat_" + thisval).css("background-color", "#999999")
            $("#bn_restestimat_" + thisval).css("color", "#000000")
            $("#bn_restestimat_" + thisval).html('>>')

        }, 1000);
        
    }


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

        //$("#aa_job_komm_" + thisval).text('[ Gem >> ]')
        //$("#aa_job_komm_" + thisval).css("color", "yellowgreen");

    });


    $(".stade").change(function () {
        var thisid = this.id;

        var idlngt = thisid.length;
        var idtrim = thisid.slice(15, idlngt);
        restestimat(idtrim);
    });


    $(".rest").mouseover(function () {
        $(this).css('cursor', 'pointer');
    });


    $(".rest").click(function () {
        var thisid = this.id;

        var idlngt = thisid.length;
        var idtrim = thisid.slice(15, idlngt);
        restestimat(idtrim);

      
    });

    function restestimat(idtrim) {
   

        $("#divindikator_" + idtrim).css("background-color", "#999999");

        var thisval = 0
        thisval = $("#FM_restestimat_" + idtrim).val()
        thisval = thisval.replace(",", ".")

       

       
        var thisreal = 0
        thisreal = $("#FM_timerreal_" + idtrim).val()
        thisreal = thisreal.replace(",", ".")

        
        //alert(thisreal)

        var thisforkalk = 0
        thisforkalk = $("#FM_forkalk_" + idtrim).val()
        
        thisforkalk = thisforkalk.replace(",", ".")


        //alert(idtrim)
        var timer_proc = 0;
        timer_proc = $("#FM_stade_ti_pr_" + idtrim).val()

      

        var balance = 0;
        var stade = 0;
        var afv = 0;
        var res = 0;

        //alert(timer_proc)

        //Rest angivet i timer
        if (timer_proc == 0) {
            res = (thisval / 1) + (thisreal / 1)

            //alert(thisforkalk)
            if (res != 0) {
                stade = (thisreal / res) * 100
            } else {
                stade = 0
            }

            //rest angivet i proc
        } else {

            res = (thisreal / 1) * (100 / (thisval / 1))


            stade = thisval

        }

        balance = thisforkalk - res

        //alert(balance)
        if (thisforkalk < res) {
            afv = 100 - (thisforkalk / res) * 100
        } else {
            afv = 100 - (res / thisforkalk) * 100
        }



        if (stade < 0 || thisreal == 0) {
            stade = 0;
        }

        if (stade > 100) {
            stade = 100;
        }


        afv = Math.round(afv)

        stade = Math.round(stade)

        res = Math.round(res * 100) / 100
        res = String(res)
        res = res.replace(".", ",")

        balance = Math.round(balance * 100) / 100

        var bgthis = "#cccccc";
        if (balance >= 0) {
            bgthis = "#DCF5BD";
        } else {
            bgthis = "crimson";
        }

        balance = String(balance)
        balance = balance.replace(".", ",")

       

        $("#totalfb_" + idtrim).text("Forv. samlet tidsforbr.: " + res);
        $("#totalbal_" + idtrim).text("Balance: " + balance + " ");
        $("#totalafv_" + idtrim).text(" (Afv: " + afv + " %)");
        $("#totalstade_" + idtrim).text("Stade: ~ " + stade + " % afsluttet");
        //$("#forvdb_" + idtrim).text("Forv. DB: " + stade + " % ");



        $("#divindikator_" + idtrim).css("background-color", bgthis);

        
    }


});