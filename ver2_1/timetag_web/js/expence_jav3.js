

$(document).ready(function() {


  
    
    $("#sbmExpence").click(function () {

        

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
        FM_udlaeg_belob = FM_udlaeg_belob.replace(",",".")

        FM_udlaeg_valuta = $("#FM_udlaeg_valuta").val();
        FM_udlaeg_gruppe = $("#FM_udlaeg_gruppe").val();
        FM_udlaeg_faktbar = $("#FM_udlaeg_faktbar").val();
        FM_udlaeg_form = $("#FM_udlaeg_form").val();

        FM_datoer = $("#FM_datoer").val();

        error = 0

        if (jq_jobid == 0 || jq_jobid == -1)
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
            //alert("Done1")

            $.post("?jq_jobid=" + jq_jobid + "&aktid=" + aktid + "&FM_udlaeg_valuta=" + FM_udlaeg_valuta + "&FM_udlaeg_form=" + FM_udlaeg_form + "&FM_udlaeg_belob=" + FM_udlaeg_belob + "&FM_udlaeg_navn=" + FM_udlaeg_navn + "&FM_udlaeg_gruppe=" + FM_udlaeg_gruppe + "&FM_datoer=" + FM_datoer, { control: "FN_createOutlay", AjaxUpdateField: "true" }, function (data) {

             //alert("Done 2")
          

            $("#image_upload").submit();

            });
        }


    });

});

