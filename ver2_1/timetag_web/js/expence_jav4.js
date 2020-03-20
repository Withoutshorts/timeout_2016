

$(document).ready(function () {




    $("#sbmExpence").click(function () {

       // alert("HEPHEP")

        endpoint = 'convert'
        access_key = 'de1b777d2882c4fe895b0ade03dbb001';

        var from = "NA"

        $("#FM_udlaeg_valuta > option").each(function () {
            if ($(this).is(':selected')) {
                from = $(this).data('valutakode');
            }          
        });

       // alert(from)

       // from = 'USD';
        to = 'DKK'; // Basic valuta
        amount = $("#FM_udlaeg_belob").val();
        //amount = amount.replace(".", ",")

        if (from != "NA") {
            $.ajax({
                url: 'https://apilayer.net/api/' + endpoint + '?access_key=' + access_key + '&from=' + from + '&to=' + to + '&amount=' + amount + "&format=1",
                dataType: 'jsonp',
                success: function (json) {
                    // access the conversion result in json.result
                   // alert(json.result);
                    var kurs = parseFloat(json.result) / parseFloat(amount)
                   // alert(kurs)
                   // alert(kurs)
                    IndlaesExpence(to, kurs, json.result);
                },
                error: function (json) {

                    alert("error")

                }
            });
        }

    });

    function IndlaesExpence(basic_valuta, basic_kurs, basic_belob)
    {

       // alert(basic_valuta + " " + basic_kurs + " " + basic_belob)

        mobil_week_reg_job_dd = $("#mobil_week_reg_job_dd").val();
        mobil_week_reg_akt_dd = $("#mobil_week_reg_akt_dd").val();

        if (mobil_week_reg_job_dd != 1) {
            jq_jobid = $("#FM_jobid").val();
        } else {
            jq_jobid = $("#dv_job").val();
        }

        if (mobil_week_reg_akt_dd != 1) {
            aktid = $("#FM_aktid").val();
        } else {
            aktid = $("#dv_akt").val();
        }

        FM_udlaeg_navn = $("#FM_udlaeg_navn").val();

        FM_udlaeg_belob = $("#FM_udlaeg_belob").val();
        FM_udlaeg_belob = FM_udlaeg_belob.replace(",", ".")

        FM_udlaeg_valuta = $("#FM_udlaeg_valuta").val();
        FM_udlaeg_gruppe = $("#FM_udlaeg_gruppe").val();
        FM_udlaeg_faktbar = $("#FM_udlaeg_faktbar").val();
        FM_udlaeg_form = $("#FM_udlaeg_form").val();

        FM_datoer = $("#FM_datoer").val();

        error = 0

        if (jq_jobid == 0 || jq_jobid == -1) {
            error = 1
            $("#errorMessage").text("Der mangler at blive valgt et job")
        }

        if (FM_udlaeg_navn == "") {
            error = 1
            $("#errorMessage").text("Der mangler at blive angivet et navn")
        }

        if (FM_udlaeg_belob == "") {
            error = 1
            $("#errorMessage").text("Der mangler at blive angivet et beløb")
        }

        //$("#jobidholderMatreg").val(jq_jobid);

        //alert("jobid " + jq_jobid + " aktid " + aktid + " navn " + FM_udlaeg_navn + " belob " + FM_udlaeg_belob + " valuta " + FM_udlaeg_valuta + " gruppe " + FM_udlaeg_gruppe + " faktbar " + FM_udlaeg_faktbar + " form " + FM_udlaeg_form)


        if (error == 0) {

            $.post("?jq_jobid=" + jq_jobid + "&aktid=" + aktid + "&FM_udlaeg_valuta=" + FM_udlaeg_valuta + "&FM_udlaeg_form=" + FM_udlaeg_form + "&FM_udlaeg_belob=" + FM_udlaeg_belob + "&FM_udlaeg_navn=" + FM_udlaeg_navn + "&FM_udlaeg_gruppe=" + FM_udlaeg_gruppe + "&FM_datoer=" + FM_datoer + "&basic_valuta=" + basic_valuta + "&basic_kurs=" + basic_kurs + "&basic_belob=" + basic_belob, { control: "FN_createOutlay", AjaxUpdateField: "true" }, function (data) {

               // alert("data " + data)


                $("#image_upload").submit();

            });
        }
    }



});

