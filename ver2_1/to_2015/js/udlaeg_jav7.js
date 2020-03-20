$(document).ready(function () {

    var delta = 90;
    $(".rotateImage").click(function () {
        thisid = $(this).data('id');

        document.getElementById("image_" + thisid).style.webkitTransform = "rotate(" + delta + "deg)";
        delta += 90;

    });


    $("#afregnAlle").click(function () {

        if ($("#afregnAlle").prop("checked") == true) {
            ($(".afregnMat").prop("checked", true))
        } else {
            ($(".afregnMat").prop("checked", false))
        }
    });

    $("#matsub").click(function () {

        //alert("click2")

        jobid = $("#FM_jobid").val();
        aktid = $("#FM_aktid_0").val();
        mid = $("#matreg_medid").val();
        navn = $("#FM_navn").val();

        FM_pris = $("#FM_belob").val();
        FM_pris = FM_pris.replace(",", ".")

        valuta = $("#FM_valuta").val();
        gruppe = $("#FM_gruppe").val();
        personlig = $("#FM_personlig").val();
        intkode = $("#FM_faktbar").val();
        matregid = $("#matregid").val();

        mat_func = $("#func").val();

        if (mat_func == "red") {
            mat_func = "dbred"
        }
        else {
            mat_func = "dbopr"
        }

        error = 0

        if (jobid == 0) {
            error = 1
            $("#errorMessageRow").show();
            $("#errorMessage").text("Der mangler at blive valgt et job")
        }

        if (navn == "") {
            error = 1
            $("#errorMessageRow").show();
            $("#errorMessage").text("Der mangler at blive angivet et navn")
        }

        if (FM_pris == "") {
            error = 1
            $("#errorMessageRow").show();
            $("#errorMessage").text("Der mangler at blive angivet et beløb")
        }


        if (error == 0) {

            var from = "NA"

            $("#FM_valuta > option").each(function () {
                if ($(this).is(':selected')) {
                    from = $(this).data('valutakode');
                }
            });

            basic_valuta = $("#FM_basic_valuta").val();
            to = basic_valuta

            endpoint = 'convert'
            access_key = 'de1b777d2882c4fe895b0ade03dbb001';

            amount = FM_pris;
            amount = amount.replace(/\,/g, '.')

            forbrugsdato = $("#FM_forbrugsdato").val();
            
            currencyDate = $("#FM_kobsdato").val(); // Købtsdato skal være her i tilfælde af at man ændre pris eller valuta
           // alert("from " + from)
           // alert(currencyDate)
           // alert(currencyDate)

            if (from != "NA") {

                // set endpoint and your access key
                endpoint = 'historical'
                access_key = 'de1b777d2882c4fe895b0ade03dbb001';
             //   alert(from)
                source = from
                date = currencyDate

                codestr = 'json.quotes.' + source + to
               // alert(codestr)
                // get the most recent exchange rates via the "live" endpoint:
                $.ajax({
                    url: 'https://apilayer.net/api/' + endpoint + '?access_key=' + access_key + '&date=' + date + "&source=" + source,
                    success: function (json) {

                        basic_kurs = eval(codestr)
                       // alert("Herp " + basic_kurs)
                        basic_valuta = to
                        basic_belob = amount * basic_kurs

                      //  alert(basic_kurs + ' ' + basic_valuta + ' ' + basic_belob);

                        $.post("?jobid=" + jobid + "&aktid=" + aktid + "&mid=" + mid + "&navn=" + navn + "&FM_pris=" + FM_pris + "&valuta=" + valuta + "&gruppe=" + gruppe + "&personlig=" + personlig + "&intkode=" + intkode + "&matregid=" + matregid + "&mat_func=" + mat_func + "&basic_valuta=" + basic_valuta + "&basic_kurs=" + basic_kurs + "&basic_belob=" + basic_belob + "&forbrugsdato=" + forbrugsdato, { control: "FN_indlasmat_matudlaeg2015", AjaxUpdateField: "true" }, function (data) {

                            // alert(data)

                            $("#imageUpload").submit();

                        })

                    }
                });

            }
            //alert("txt")
            //alert("jobid = " + jobid + "\n & aktid=" + aktid + " \n& mid=" + mid + " \n& navn=" + navn + " \n&FM_pris=" + FM_pris + " \n & valuta=" + valuta + " \n & gruppe=" + gruppe + " \n& personlig=" + personlig + " \n& intkode=" + intkode + " \n& matregid=" + matregid + " \n& mat_func=" + mat_func)

        }

    });

    $(".FM_job").keyup(function () {

        //alert("awd")
        sogjobogkunde(0)

    });


    $("#FM_vis_passive").click(function () {

        if ($("#FM_vis_passive").is(':checked') == true) {
            vispassive_medarb = 1
        } else {
            vispassive_medarb = 0
        }

        matreg_medid = $("#matreg_medid").val()

        $.post("?matreg_medid=" + matreg_medid + "&vispassive_medarb=" + vispassive_medarb, { control: "FN_vis_pas_medarb_matudlaeg2015", AjaxUpdateField: "true" }, function (data) {

            $("#td_selmedarb").html(data)


        });

    });




    $("#godkendAll").click(function () {

        $(".godkendCHB").prop('checked', true)

    });

    $("#afvisAll").click(function () {

        $(".afvisCHB").prop('checked', true)

    });



    $(".FM_akt").keyup(function () {

        // alert("key up")
        $(".dv_job").hide();

        sogakt(0);

    });


    function sogakt(thisval) {

        //alert("Sog akt")

        // mobil_week_reg_akt_dd = $("#mobil_week_reg_akt_dd").val()
        // mobil_week_reg_job_dd = $("#mobil_week_reg_job_dd").val()

        //alert(mobil_week_reg_job_dd + " # " + mobil_week_reg_akt_dd)


        jq_newfilterval = $("#FM_akt_" + thisval).val()

        jq_jobid = $("#FM_jobid").val()




        jq_medid = $("#matreg_medid").val()
        jq_aktid = $("#FM_akt_" + thisval).val()
        jq_pa = $("#FM_pa").val()

        varTjDatoUS_man = $("#varTjDatoUS_man").val()



        //$(".chbox_job").hide();

        //alert("jq_newfilterval" + jq_newfilterval + " jq_jobid" + jq_jobid + " jq_medid " + jq_medid + " jq_aktid " + jq_aktid + " jq_pa " + jq_pa + " varTjDatoUS_man " + varTjDatoUS_man + " jq_aktidc " + jq_aktidc)



        $("#dv_akt_" + thisval).css("display", "");
        $("#dv_akt_" + thisval).css("visibility", "visible");
        $("#dv_akt_" + thisval).show(100);

        jq_aktidc = $("#jq_aktidc").val()

        //alert("#" + thisval)
        //alert(jq_jobid + ":: medid: " + jq_medid + " jq_aktid: " + jq_aktid + "## jq_pa" + jq_pa)
        //alert("#" + jq_newfilterval.length + "")

        if (jq_newfilterval.length > 0 || mobil_week_reg_akt_dd == "1") {

            $.post("?jq_newfilterval=" + jq_newfilterval + "&jq_jobid=" + jq_jobid + "&jq_medid=" + jq_medid + "&jq_aktid=" + jq_aktid + "&jq_pa=" + jq_pa + "&varTjDatoUS_man=" + varTjDatoUS_man + "&jq_aktidc=" + jq_aktidc, { control: "FN_sogakt_matudlaeg2015", AjaxUpdateField: "true" }, function (data) {
                //alert("cc akt")
                //$("#dv_akt_test").html(data);
                $("#dv_akt_" + thisval).html(data);

                //alert("END")

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

            });

        }



    }



    function sogjobogkunde(thisval) {


        mobil_week_reg_job_dd = $("#mobil_week_reg_job_dd").val()
        mobil_week_reg_akt_dd = $("#mobil_week_reg_akt_dd").val()


        jq_newfilterval = $("#FM_job").val()
        jq_medid = $("#matreg_medid").val()


        jq_jobidc = $("#jq_jobidc").val()

        $("#dv_job").css("display", "");
        $("#dv_job").css("visibility", "visible");
        $("#dv_job").show(100);


        varTjDatoUS_man = $("#varTjDatoUS_man").val()

        lto = $("#lto").val()




        //$(".chbox_akt").hide();

        //alert("thisval: " + thisval + " jq_newfilterval: " + jq_newfilterval + " jq_medid: " + jq_medid)

        if (jq_newfilterval.length > 0) {

            //alert("H4")

            $.post("?lto=" + lto + "&jq_newfilterval=" + jq_newfilterval + "&jq_medid=" + jq_medid + "&varTjDatoUS_man=" + varTjDatoUS_man + "&jq_jobidc=" + jq_jobidc, { control: "FN_sogjobogkunde_matudlaeg2015", AjaxUpdateField: "true" }, function (data) {
                //alert("cc")


                $("#dv_job").html(data);



                // if (mobil_week_reg_job_dd != "1") {

                $(".chbox_job").bind('click', function () {



                    var thisvaltrim = $("#dv_job").val()
                    //alert("her: " + thisvaltrim)


                    thisJobid = thisvaltrim
                    //thisJobnavn = $("#hiddn_job_" + thisJobid).val()
                    thisJobnavn = $("#dv_job option:selected").text()

                    $("#FM_job").val(thisJobnavn)
                    $("#FM_jobid").val(thisJobid)

                    //$(".dv_job").hide();
                    $(".chbox_job").hide();


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

                });



                // } // mobil_week_reg_job_dd


                //alert("jq_jobidc" + jq_jobidc)
                if (jq_jobidc != "-1") {
                    $("#dv_job").val(jq_jobidc)
                    //sogakt(0);
                    $("#dv_akt_0").prop("disabled", false);
                }

            });

        }
    }






});


