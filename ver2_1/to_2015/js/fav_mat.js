





$(document).ready(function () {

    $(".kommodal").click(function () {

        /*thisid = this.id
      
        var idlngt = thisid.length
        var idtrim = thisid.slice(6, idlngt)

        alert(idtrim) */

    });

    $(".matreg_sb").click(function () {

        thisid = this.id
        //alert(thisid)

       
        mat_antal = $("#mat_antal_" + thisid).val()
        mat_navn = $("#mat_navn_" + thisid).val()
        mat_kobpris = $("#mat_kobpris_" + thisid).val()
        mat_enhed = $("#mat_endhed_" + thisid).val()
        mat_jobid = $("#mat_jobid_" + thisid).val()
        mat_aktid = $("#mat_aktid_" + thisid).val()
        mat_dato = $("#mat_dato_" + thisid).val()
        mat_forbrugsdato = $("#mat_forbrugsdato_" + thisid).val()
        mat_editor = $("#mat_editor").val()
        mat_usrid = $("#mat_userid").val()
        mat_gruppe = $("#mat_gruppe_" + thisid).val()
        mat_salgspris = $("#mat_salgspris_" + thisid).val()
        mat_bilagsnr = $("#mat_bilagsnr").val()
        mat_valuta = $("#mat_valuta" + thisid).val()
        mat_varenr = $("#mat_varenr").val()
      
        //alert(mat_varenr)

        if (mat_navn == "")
        {
           // alert("fejl")
            $("#error_felt_" + thisid).css("display", "");
            $("#error_felt_" + thisid).css("visibility", "visible");
            $("#error_felt_" + thisid).show(1000);

            $("#error_txt_" + thisid).html("Du mangler at angive et navn")
        } else { 

        $("#error_felt_" + thisid).css("display", "none");
        $("#error_felt_" + thisid).css("visibility", "hidden");
        $("#error_felt_" + thisid).show();
        }
                

        //document.getElementById('error_felt').style.display = "none"
            
       // alert(mat_antal)
       // alert(mat_navn)
       // alert(mat_kobpris)
       // alert(mat_enhed)
       // alert(mat_jobid)
       // alert(mat_aktid)

        $.post("?mat_jobid=" + mat_jobid + "&mat_aktid=" + mat_aktid + "&mat_antal=" + mat_antal + "&mat_navn=" + mat_navn + "&mat_kobpris=" + mat_kobpris + "&mat_enhed=" + mat_enhed + "&mat_dato=" + mat_dato + "&mat_forbrugsdato=" + mat_forbrugsdato + "&mat_editor=" + mat_editor + "&mat_usrid=" + mat_usrid + "&mat_gruppe=" + mat_gruppe + "&mat_salgspris=" + mat_salgspris + "&mat_bilagsnr=" + mat_bilagsnr + "&mat_valuta=" + mat_valuta + "&mat_varenr=" + mat_varenr, { control: "tilfoj_mat", AjaxUpdateField: "true" }, function (data) {

            alert("post") 

        });

    });




  /*  $(".matreg_sb").click(function () {

        //alert("her")

       indlaes_mat(0);

    });


    function indlaes_mat(otf_lager) {

        alert("indladats")

        matreg_lto = $("#matreg_lto").val()
        //oftlager = oftlager

        matreg_func = $("#matreg_func").val()         
        if (matreg_func.length == 0) {
            matreg_func = "dbopr"
        } else {
            matreg_func = matreg_func
        }

        //error skal måske ind her og resettes

        matregid = $("#matregid").val()
        matreg_jobid = $("#matreg_jobid").val()

        matreg_aftid = $("#matreg_aftid").val()
        matreg_medid = $("#matreg_medid").val()


        if (otf_lager == 0) {
            matreg_onthefly = 1 //OnTheFly   $("#matreg_onthefly").val()
        } else {
            matreg_onthefly = 0 //Fra lager
        }


        matreg_aktid = $("#matreg_aktid").val()
        if (matreg_aktid == null) {
            matreg_aktid = 0
        }

        matreg_regdato = $("#matreg_regdato_" + otf_lager).val()

        // Error felter mangler at blive sat ind
        if (matreg_regdato.length == 0) {
            $("#matreg_indlast_err").html("Dato er ikke gyldig.")
            $("#matreg_indlast_err").css("background-color", "#FFC0CB")
            return false;
        }


        matreg_matid = $("#matreg_matid_" + otf_lager).val()
        matreg_antal = $("#matreg_antal_" + otf_lager).val()










        alert(matreg_regdato)


    } */

});


