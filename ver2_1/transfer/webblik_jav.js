$(document).ready(function () {

    $("select[name*=ajax]").AjaxUpdateField({ parent: "tr", subselector: "td:first > input[name=rowId]" });
    $("#incidentlist").table_sort({ items: 'tbody > tr:gt(1):has(:input[name=rowId])', IdControlNode: "td:first > :input[name=rowId]" });



    $(".aa_job_komm").click(function () {

        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(12, thisvallngt)
        thisval = thisvaltrim


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

            $.post("?jobid=" + thisval + "&jobkom=" + jobkom, { control: "FN_updatejobkom", AjaxUpdateField: "true" }, function (data) {

                $("#aa_job_komm_" + thisval).text('  Ok!')
                $("#aa_job_komm_" + thisval).css("color", "#5582d2");

                $("#FM_job_komm_" + thisval).css("color", "#999999");
                $("#FM_job_komm_" + thisval).val('Job tweet..(åben for alle)');

                var forTxt = $("#dv_job_komm_" + thisval).html()
                //$("#dv_job_komm_" + thisval).html(jobkom + "<br>" + forTxt);
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


    $(".stade").change(function () {
        var thisid = this.id;

        var idlngt = thisid.length;
        var idtrim = thisid.slice(18, idlngt);
        restestimat(idtrim);
    });


    $(".rest").keyup(function () {
        var thisid = this.id;

        var idlngt = thisid.length;
        var idtrim = thisid.slice(15, idlngt);
        restestimat(idtrim);
    });

    function restestimat(idtrim) {

        //alert(idtrim)

        $("#divindikator_" + idtrim).css("background-color", "#999999");

        var thisval = 0
        thisval = $("#FM_restestimat_" + idtrim).val()
        thisval = thisval.replace(",", ".")

        var thisreal = 0
        thisreal = $("#FM_timerreal_" + idtrim).val()
        thisreal = thisreal.replace(",", ".")

        var thisforkalk = 0
        thisforkalk = $("#FM_forkalk_" + idtrim).val()
        //alert(thisforkalk)
        thisforkalk = thisforkalk.replace(",", ".")

        var timer_proc = 0;
        timer_proc = $("#FM_stade_tim_proc_" + idtrim).val()

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