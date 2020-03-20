

$(document).ready(function() {


    $('.date').datepicker({

    });

    //alert("Click3")

    $("#sbmExpence").click(function () {

       // alert("Click3")


        jq_jobid = $("#FM_jobid").val();
        aktid = $("#FM_aktid").val();
        FM_udlaeg_navn = $("#FM_udlaeg_navn").val();

        FM_udlaeg_belob = $("#FM_udlaeg_belob").val();
        FM_udlaeg_belob = FM_udlaeg_belob.replace(",",".")

        FM_udlaeg_valuta = $("#FM_udlaeg_valuta").val();
        FM_udlaeg_gruppe = $("#FM_udlaeg_gruppe").val();
        FM_udlaeg_faktbar = $("#FM_udlaeg_faktbar").val();
        FM_udlaeg_form = $("#FM_udlaeg_form").val();

        error = 0

        if (jq_jobid == 0)
        {
            error = 1
            $("#errorMessage").text("Der mangler at blive valgt et job")           
        }

        if (FM_udlaeg_navn == "")
        {
            error = 1
            $("#errorMessage").text("Der mangler at blive angivet et navn")
        }

        if (FM_udlaeg_belob == "")
        {
            error = 1
            $("#errorMessage").text("Der mangler at blive angivet et beløb")
        }

        //$("#jobidholderMatreg").val(jq_jobid);

        //alert("jobid " + jq_jobid + " aktid " + aktid + " navn " + FM_udlaeg_navn + " belob " + FM_udlaeg_belob + " valuta " + FM_udlaeg_valuta + " gruppe " + FM_udlaeg_gruppe + " faktbar " + FM_udlaeg_faktbar + " form " + FM_udlaeg_form)


        if (error == 0)
        {        
        $.post("?jq_jobid=" + jq_jobid + "&aktid=" + aktid + "&FM_udlaeg_valuta=" + FM_udlaeg_valuta + "&FM_udlaeg_form=" + FM_udlaeg_form + "&FM_udlaeg_belob=" + FM_udlaeg_belob + "&FM_udlaeg_navn=" + FM_udlaeg_navn + "&FM_udlaeg_gruppe=" + FM_udlaeg_gruppe, { control: "FN_createOutlay", AjaxUpdateField: "true" }, function (data) {

           // alert("Done!")

            $("#image_upload").submit();

            });
        }


    });


    $("#FM_akt").click(function () {

        $("#FM_aktid").val(0)

    });

    $("#FM_akt").keyup(function () {

        $("#dv_job").hide();

        sogakt(0);

    });

    function sogakt(tomobjid) {
        
        tomobjid = tomobjid

        //mobil_week_reg_akt_dd = $("#mobil_week_reg_akt_dd").val()
        //mobil_week_reg_job_dd = $("#mobil_week_reg_job_dd").val()       

        if (tomobjid == 0) {

            jq_newfilterval = $("#FM_akt").val()

            jq_jobid = $("#FM_jobid").val()


        } else {
            jq_jobid = tomobjid
        }
        
        //alert(jq_jobid)
        //alert(mobil_week_reg_job_dd + " # " + mobil_week_reg_akt_dd + " tomobjid: " + tomobjid + " jq_jobid: " + jq_jobid)

        //jq_newfilterval = $("#FM_akt").val()
        //jq_jobid = $("#FM_jobid").val()
        jq_medid = $("#FM_medid_k").val()
        jq_aktid = $("#FM_akt").val()
        jq_pa = $("#FM_pa").val()

        

        varTjDatoUS_man = $("#varTjDatoUS_man").val()

        if (jq_newfilterval.length > 0) {


            $.post("?jq_newfilterval=" + jq_newfilterval + "&jq_jobid=" + jq_jobid + "&jq_medid=" + jq_medid + "&jq_aktid=" + jq_aktid + "&jq_pa=" + jq_pa + "&varTjDatoUS_man=" + varTjDatoUS_man, { control: "FN_sogakt", AjaxUpdateField: "true" }, function (data) {
                //alert("cc")
                
                $("#dv_akt").html(data);

                //alert("her")
                //$("#dv_akt").attr('class', 'dv-open');

                //if (mobil_week_reg_akt_dd != "1") {

                    if ($("#dv_akt").css('display') == "none") {
                        $("#dv_akt").css("display", "");
                        $("#dv_akt").css("visibility", "visible");
                        $("#dv_akt").show(100);
                    }



                    $(".luk_aktsog").bind('mouseover', function () {

                        $(this).css("cursor", "pointer");

                    });



                    $(".luk_aktsog").bind('click', function () {

                        $("#dv_akt").hide();
                        //$("#dv_akt").attr('class', 'dv-closed');

                    });


                    $('.span_akt').bind('mouseover', function () {

                        $(this).css("background-color", "#CCCCCC");
                        $(this).css("cursor", "pointer");

                    });

                    $('.span_akt').bind('mouseout', function () {

                        $(this).css("background-color", "#FFFFFF");

                    });


                    $(".span_akt").bind('click', function () {


                        var thisaktid = this.id
                        var thisvallngt = thisaktid.length
                        var thisvaltrim = thisaktid.slice(9, thisvallngt)


                        thisaktid = thisvaltrim
                        thisAktnavn = $("#hiddn_akt_" + thisaktid).val()

                        //alert(thisAktnavn + " " + thisaktid)

                        $("#FM_akt").val(thisAktnavn)
                        $("#FM_aktid").val(thisaktid)
                        $("#dv_akt").hide();
                        //$("#dv_akt").attr('class', 'dv-closed');


                    });


               // }




            });

        } else {

            //if (mobil_week_reg_akt_dd != "1") {
                $("#dv_akt").hide();
            //}
        }



    }


    $("#FM_job").click(function () {

        $("#FM_jobid").val(0)

    });

    $("#FM_job").keyup(function () {

     $("#dv_akt").hide();
     sogjobogkunde();

    });


    function sogjobogkunde() {

        //mobil_week_reg_job_dd = $("#mobil_week_reg_job_dd").val()
        //mobil_week_reg_akt_dd = $("#mobil_week_reg_akt_dd").val()

        varTjDatoUS_man = $("#varTjDatoUS_man").val()

        jq_newfilterval = $("#FM_job").val()

        jq_medid = $("#FM_medid_k").val()

        jq_lto = $("#jq_lto").val()
        thisJobid = $("#tomobjid").val()

        // SÆTER prvalgt job til når der er indlæst. DD/Jobsøg starter med dette job
        // $("#tomobjid").val(jq_newfilterval)

        //alert(thisJobid)


        if ($("#dv_job").css('display') == "none") {

            $("#dv_job").css("display", "");
            $("#dv_job").css("visibility", "visible");
            $("#dv_job").show(100);

        }


        //alert(jq_newfilterval.length)
        //alert("cc" + jq_newfilterval)
        if (jq_newfilterval.length > 0) {

            if (jq_newfilterval == "%") {
                jq_newfilterval = "all"
            }

            $.post("?jq_newfilterval=" + jq_newfilterval + "&jq_medid=" + jq_medid + "&varTjDatoUS_man=" + varTjDatoUS_man + "&jq_lto=" + jq_lto + "&thisJobid=" + thisJobid, { control: "FN_sogjobogkunde", AjaxUpdateField: "true" }, function (data) {



                $("#dv_job").html(data);

                //$("#dv_job").attr('class', 'dv-open');



                $('#luk_jobsog').bind('mouseover', function () {


                    $(this).css("cursor", "pointer");

                });




                $("#luk_jobsog").bind('click', function () {


                    $("#dv_job").hide();
                    //$("#dv_job").attr('class', 'dv-closed');

                });


                $(".span_job").bind('click', function () {



                    var thisjobid = this.id
                    var thisvallngt = thisjobid.length
                    var thisvaltrim = thisjobid.slice(10, thisvallngt)

                    thisJobid = thisvaltrim

                    thisJobnavn = $("#hiddn_job_" + thisJobid).val()

                    $("#FM_job").val(thisJobnavn)
                    $("#FM_jobid").val(thisJobid)

                    $("#dv_job").hide();

                    // Jobbeskrivelse //
                    jq_lto = $("#jq_lto").val()

                    if (jq_lto == 'hestia' || jq_lto == 'intranet - local') {
                        sogjobbesk(thisJobid);
                    }

                });


                $('.span_job').bind('mouseover', function () {

                    $(this).css("background-color", "#CCCCCC");
                    $(this).css("cursor", "pointer");

                });

                $('.span_job').bind('mouseout', function () {

                    $(this).css("background-color", "#FFFFFF");

                });



            });

        } else {
            $("#dv_job").hide();
        }
    }


});

