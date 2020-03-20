
$(document).ready(function () {

    
    //$(".dagecls").bind('keyup', function () {

    $(".beregnsalgspris_txt").keyup(function () {

        var thisid = this.id
        thisVal = $("#" + this.id).val().replace(".", "")
        $("#" + this.id).val(thisVal)


    });


    $(".salgspris_txt").keyup(function () {

        var thisid = this.id
        thisVal = $("#" + this.id).val().replace(".", "")
        $("#" + this.id).val(thisVal)
    });

    
    $("#FM_vis_passive").click(function () {

        if ($("#FM_vis_passive").is(':checked') == true) {
            vispassive_medarb = 1
        } else {
            vispassive_medarb = 0
        }

            matreg_medid = $("#matreg_medid").val()

            $.post("?matreg_medid=" + matreg_medid + "&vispassive_medarb=" + vispassive_medarb, { control: "FN_vis_pas_medarb", AjaxUpdateField: "true" }, function (data) {

                $("#td_selmedarb").html(data)


            });

      

    });


    vispasogluk = $.cookie("vispasogluk");
    ignprojekgrp = $.cookie("ignprojekgrp");
   
    
        
    if (vispasogluk == '1') {
        $("#vispasogluk").attr('checked', true);
    }
   
    if (ignprojekgrp == '1') {
        $("#ignprojekgrp").attr('checked', true);
    }

    //jobliste(0);

    
    $("#matreg_jobid").change(function () {

        matreg_jobid = $("#matreg_jobid").val()

        $.post("?matreg_jobid=" + matreg_jobid, { control: "FN_sog_akt", AjaxUpdateField: "true" }, function (data) {

            $("#matreg_aktid").html(data)

          
        });

    });

    $("#a_lager").click(function () {
        indlaes_allerede_mat();
        lagershow();
        $("#sp_mat_forbrug").html('[-]')

    });


    function lagershow() {

        $("#dv_mat_lager").css("display", "");
        $("#dv_mat_lager").css("visibility", "visible");
        $("#dv_mat_lager").show(1000);

        $("#dv_mat_lager_list").css("display", "");
        $("#dv_mat_lager_list").css("visibility", "visible");
        $("#dv_mat_lager_list").show(1000);

        $("#dv_mat_otf").hide()
        $("#dv_mat_otf_sb").hide()
    }



    $("#sp_mat_forbrug").mouseover(function () {
        $(this).css("cursor", "pointer");
    });

    $("#sp_komm").mouseover(function () {
        $(this).css("cursor", "pointer");
    });

    $("#sp_komm").click(function () {

        pls_minus = $("#sp_komm").html()

        if (pls_minus == "[-]") {
            $("#sp_komm").html('[+]')
        } else {
            $("#sp_komm").html('[-]')
        }

        if (pls_minus == "[-]") {
            $(".tr_kom").hide()
        } else {

            $(".tr_kom").css("visibility", "visible")
            $(".tr_kom").css("display", "")

        }

    });



    $("#sp_mat_forbrug").click(function () {

        pls_minus = $("#sp_mat_forbrug").html()

        if (pls_minus == "[-]") {
            $("#sp_mat_forbrug").html('[+]')
        } else {
            $("#sp_mat_forbrug").html('[-]')
        }

        if (pls_minus == "[-]") {

            $("#dv_mat_lager").hide()
            $("#dv_mat_lager_list").hide()
            $("#dv_mat_otf").hide()
            $("#dv_mat_otf_sb").hide()
            $("#matreg_indlast_allerede").hide()

        } else {



            lto = $("#matreg_lto").val()
            //alert(lto)

            indlaes_allerede_mat();

            if (lto == 'dencker' || lto == "jttek" || lto == "intranet - local") {
                lagershow();
            } else {
                otfshow();
            }


        }

    });


    

    $("#a_lager").mouseover(function () {
        $(this).css("cursor", "pointer");
    });

    $("#a_otf").mouseover(function () {
        $(this).css("cursor", "pointer");
    });


    $("#a_otf").click(function () {

        indlaes_allerede_mat();
        oftshow();
        $("#sp_mat_forbrug").html('[-]')


    });


    function oftshow() {

        $("#dv_mat_otf").css("display", "");
        $("#dv_mat_otf").css("visibility", "visible");
        $("#dv_mat_otf").show(1000);

        $("#dv_mat_otf_sb").css("display", "");
        $("#dv_mat_otf_sb").css("visibility", "visible");
        $("#dv_mat_otf_sb").show(1000);

        $("#dv_mat_lager").hide()
        $("#dv_mat_lager_list").hide()


    }



    $("#matreg_sb").click(function () {

        //alert("her")
     
        indlaes_mat(0);

    });




    $("#sogmatgrp, #soglev").change(function () {

        sog_lager();

    });

    $("#FM_sog").keyup(function () {

        sog_lager();

    });




    function sog_lager() {


        matreg_grp = $("#sogmatgrp").val()
        matreg_lev = $("#soglev").val()
        matreg_sog = $("#FM_sog").val()

        $.cookie('matreg_grp', matreg_grp);
        $.cookie('matreg_lev', matreg_lev);
        $.cookie('matreg_sog', matreg_sog);

     
        //alert(matreg_lev + " " + matreg_grp + " " + matreg_sog)
        //$.post("?matreg_jobid=" + matreg_jobid, { control: "FN_indlaes_mat", AjaxUpdateField: "true" }, function (data) {
        $.post("?matreg_grp=" + matreg_grp + "&matreg_lev=" + matreg_lev + "&matreg_sog=" + matreg_sog, { control: "FN_sog_lager", AjaxUpdateField: "true" }, function (data) {

            

            $("#dv_mat_lager_list").html(data)

            $(".matreglg_sb").bind('click', function () {

                var thisid = this.id
                //alert(this.id)
                indlaes_mat(thisid);

            });

            $(".beregnsalgspris_txt").unbind('keyup').bind('keyup', function () {

                var thisid = this.id

                //alert(thisid)
                thisVal = $("#" + this.id).val().replace(".", "")
                $("#" + this.id).val(thisVal)

                var idlngt = thisid.length
                var idtrim = thisid.slice(5, idlngt)

                beregnsalgsprisOTF_jq(idtrim)

            });

            
           $(".s_matgrp").bind('change', function () {

                //alert("her")

                var thisid = this.id

                var idlngt = thisid.length
                var idtrim = thisid.slice(7, idlngt)

                beregnsalgsprisOTF_jq(idtrim)

            });
            

         
         


          
           $(".salgspris_txt").unbind('keyup').bind('keyup', function () {

               var thisid = this.id
               thisVal = $("#" + this.id).val().replace(".", "")
               $("#" + this.id).val(thisVal)
           });



        });

    }



    $("#FM_jobsog").keyup(function () {

        jobliste(1);

    });

    $("#vispasogluk").click(function () {
        jobliste(1);
    });

    $("#ignprojekgrp").click(function () {
        jobliste(1);
    });

    
    

    function jobliste(vis) {

        if (vis == 1) {
            jobsogVal = $("#FM_jobsog").val()
        } else {
            jobsogVal = "a"
        }

        if ($("#vispasogluk").is(':checked') == true) {
            vispasogluk = 1
            $.cookie('vispasogluk', '1');
        } else {
            vispasogluk = 0
            $.cookie('vispasogluk', '0');
        }


        if ($("#ignprojekgrp").is(':checked') == true) {
            ignprojekgrp = 1
            $.cookie('ignprojekgrp', '1');
        } else {
            ignprojekgrp = 0
            $.cookie('ignprojekgrp', '0');
        }

        matreg_jobid = $("#matreg_jobid").val()


        $.post("?matreg_jobid = " + matreg_jobid + "&jobsogVal=" + jobsogVal + "&vispasogluk=" + vispasogluk + "&ignprojekgrp=" + ignprojekgrp, { control: "FN_sogjob_mat", AjaxUpdateField: "true" }, function (data) {




            $("#matreg_jobid").html(data)


        });

    }


    // ******** Materiale indlæste på aktivitet på dagen allerede ////
    function indlaes_allerede_mat() {



        matreg_medid = $("#matreg_medid").val()
        matreg_aktid = $("#matreg_aktid").val()
        matreg_regdato = $("#matreg_regdato_0").val()


        //alert("her" + matreg_aktid)

        $.post("?matreg_medid=" + matreg_medid + "&matreg_aktid=" + matreg_aktid + "&matreg_regdato=" + matreg_regdato, { control: "FN_indlaest_allerede_mat", AjaxUpdateField: "true" }, function (data) {


            //alert(data.length)

            if (data.length > 10) {
                $("#matreg_indlast_allerede").html(data)

                $("#matreg_indlast_allerede").css("visibility", "visible")
                $("#matreg_indlast_allerede").css("display", "")
            }


        });



        setTimeout(function () {
            // Do something after 5 seconds
            $("#matreg_indlast").html('');
            $("#matreg_indlast").css("background-color", "#FFFFFF")
        }, 2500);



    }

    

    

    /// **** INDLÆS MATERIALER ******//
    function indlaes_mat(otf_lager) {

        matreg_lto = $("#matreg_lto").val()
     

        otf_lager = otf_lager

       
        matreg_func = $("#matreg_func").val()
        if (matreg_func.length == 0) {
            matreg_func = "dbopr"
        } else {
            matreg_func = matreg_func
        }

        //alert("her2")

        // Nulstiller ecvt. err msg
        $("#matreg_indlast_err").html('')
        $("#matreg_indlast_err").css("background-color", "#ffffff")

        matregid = $("#matregid").val()

        // *** Værdier **/
        // ** STAM DATA **/
        matreg_jobid = $("#matreg_jobid").val()

     

        if (matreg_jobid == null){
            $("#matreg_indlast_err").html("Der mangler at blive valgt et job.")
            $("#matreg_indlast_err").css("background-color", "#FFC0CB")
            return false;
        }

        if (matreg_jobid == 0 || matreg_jobid.length == 0) {
            $("#matreg_indlast_err").html("Der mangler at blive valgt et job.")
            $("#matreg_indlast_err").css("background-color", "#FFC0CB")
            return false;
        }

        matreg_aftid = $("#matreg_aftid").val()
        matreg_medid = $("#matreg_medid").val()

        if (otf_lager == 0) {
            matreg_onthefly = 1 //OnTheFly   $("#matreg_onthefly").val()
        } else {
            matreg_onthefly = 0 //Fra lager
        }
        //alert(matreg_onthefly)
        matreg_aktid = $("#matreg_aktid").val()
        if (matreg_aktid == null) {
            matreg_aktid = 0
        }

       

        matreg_regdato = $("#matreg_regdato_" + otf_lager).val()
        //alert(matreg_regdato)
       
        if (matreg_regdato.length == 0) {
           $("#matreg_indlast_err").html("Dato er ikke gyldig.")
           $("#matreg_indlast_err").css("background-color", "#FFC0CB")
            return false;
        }


       

        //** Pr linje ***//
        matreg_matid = $("#matreg_matid_" + otf_lager).val()
        matreg_antal = $("#matreg_antal_" + otf_lager).val()


        if (matreg_antal == 0 || matreg_antal.length == 0) {
            $("#matreg_indlast_err").html("Der mangler at blive indtastet antal.")
            $("#matreg_indlast_err").css("background-color", "#FFC0CB")
            return false;
        }

        matreg_valuta = $("#matreg_valuta_" + otf_lager).val()
        matreg_intkode = $("#intkode_" + otf_lager).val()



        if ($("#matreg_personlig_" + otf_lager).is(':checked') == true) {
            matreg_personlig = 1
        } else {
            matreg_personlig = 0
        }



        matreg_bilagsnr = $("#matreg_bilagsnr_" + otf_lager).val()



        matreg_pris = $("#pris_" + otf_lager).val()
        if (matreg_pris.length == 0) {
            $("#matreg_indlast_err").html("Der mangler at blive indtastet pris.")
            $("#matreg_indlast_err").css("background-color", "#FFC0CB")
            return false;
        }


        //if (matreg_lto == 'epi' || matreg_lto == 'epi_no' || matreg_lto == 'epi_uk' || matreg_lto == 'epi_ab') {

        //    matreg_salgspris = matreg_pris
           
        //} else {

            matreg_salgspris = $("#FM_salgspris_" + otf_lager).val()
            if (matreg_salgspris.length == 0) {
                $("#matreg_indlast_err").html("Der mangler at blive indtastet salgspris.")
                $("#matreg_indlast_err").css("background-color", "#FFC0CB")
                return false;
            }

        //}

            matreg_gruppe = $("#gruppe_" + otf_lager).val()

            //alert(matreg_gruppe)

        matreg_navn = $("#matreg_navn_" + otf_lager).val()
        if (matreg_navn.length == 0) {
            $("#matreg_indlast_err").html("Der mangler at blive indtastet navn.")
            $("#matreg_indlast_err").css("background-color", "#FFC0CB")
            return false;
        }


      
        matreg_varenr = $("#matreg_varenr_" + otf_lager).val()

        if ($("#matreg_opretlager").is(':checked') == true) {
            matreg_opretlager = 1
        } else {
            matreg_opretlager = 0
        }

        matreg_betegn = $("#matreg_betegn_" + otf_lager).val()

        //alert(matreg_betegn)

        if ($("#FM_opdaterpris_" + otf_lager).is(':checked') == true) {
            matreg_opdaterpris = 1
        } else {
            matreg_opdaterpris = 0
        }

        strPrevHtml = $("#matreg_indlast").html()
        //alert(strPrevHtml.length)

        //if (matreg_lto == 'epi' || matreg_lto == 'epi_no' || matreg_lto == 'epi_uk' || matreg_lto == 'epi_ab') {
        //    alert("her")
        //}

        //$.post("?matreg_jobid=" + matreg_jobid, { control: "FN_indlaes_mat", AjaxUpdateField: "true" }, function (data) {

        amount = matreg_pris
        amount = amount.replace(/\,/g, '.')

        basic_valuta = $("#basic_valuta").val()
        to = basic_valuta
        from = "NA"
        $("#matreg_valuta_" + otf_lager +" > option").each(function () {
            if ($(this).is(':selected')) {
                from = $(this).attr("data-valutakode");
            }
        });
    
        endpoint = 'historical'
        access_key = 'de1b777d2882c4fe895b0ade03dbb001';
        
        kobsdatostr = $("#matreg_regdato_" + otf_lager).val();
        kobsdatostr = kobsdatostr.replace(/\//g, '-')

        kobsdato = kobsdatostr.split('-')
        kobsdato = kobsdato[2] + '-' + kobsdato[1] + '-' + kobsdato[0]

        source = from

        codestr = 'json.quotes.' + source + to

        if (matreg_pris == 0) {
            amount = 1
        }

        $.ajax({
            url: 'https://apilayer.net/api/' + endpoint + '?access_key=' + access_key + '&source=' + source + '&date=' + kobsdato,
            dataType: 'jsonp',
            success: function (json) {
                // access the conversion result in json.result
                // alert(json.result);
                basic_kurs = eval(codestr)
                basic_belob = amount * basic_kurs

               // alert(basic_kurs)
               // alert(basic_belob)
               // alert(basic_valuta)

                if (matreg_pris == 0) {
                    basic_belob = 0
                }

                $.post("?matregid=" + matregid + "&matreg_func=" + matreg_func + "&matreg_jobid=" + matreg_jobid + "&matreg_aftid=" + matreg_aftid + "&matreg_medid=" + matreg_medid + "&matreg_onthefly=" + matreg_onthefly + "&matreg_matid=" + matreg_matid + "&matreg_aktid=" + matreg_aktid + "&matreg_regdato=" + matreg_regdato + "&matreg_antal=" + matreg_antal + "&matreg_valuta=" + matreg_valuta + "&matreg_intkode=" + matreg_intkode + "&matreg_personlig=" + matreg_personlig + "&matreg_bilagsnr=" + matreg_bilagsnr + "&matreg_pris=" + matreg_pris + "&matreg_salgspris=" + matreg_salgspris + "&matreg_gruppe=" + matreg_gruppe + "&matreg_navn=" + matreg_navn + "&matreg_varenr=" + matreg_varenr + "&matreg_opretlager=" + matreg_opretlager + "&matreg_betegn=" + matreg_betegn + "&matreg_opdaterpris=" + matreg_opdaterpris + "&basic_kurs=" + basic_kurs + "&basic_valuta=" + basic_valuta + "&basic_belob=" + basic_belob, { control: "FN_indlaes_mat", AjaxUpdateField: "true" }, function (data) {

                    //alert($("#matreg_indlast").val())
                    //$("#matreg_test").val(data)
                    //strPrevHTML = ("#matreg_indlast").val()
                    if (strPrevHtml.length > 0) {
                        $("#matreg_indlast").html(strPrevHtml + " " + data)
                    } else {
                        $("#matreg_indlast").html("Har indlæst:<br> " + data)

                    }



                    $("#matreg_indlast").css("background-color", "#DCF5BD")

                    indlaes_allerede_mat() //matreg_regdato

                });
              
            },
            error: function (json) {

                alert("error")

            }
        });


        ///**  Nulstiller værdier **///
        $("#matreg_antal_" + otf_lager).val('')
        $("#matreg_personlig_" + otf_lager).attr('checked', false);


        //$("#matreg_bilagsnr").val('')

       
        //$("#gruppe_" + otf_lager).val('')

        if (otf_lager == 0) {
        $("#matreg_navn_" + otf_lager).val('')
        $("#pris_" + otf_lager).val('')
        $("#FM_salgspris_" + otf_lager).val('')
        }

        $("#matreg_varenr_" + otf_lager).val('')

        $("#matreg_opretlager_" + otf_lager).attr('checked', false);

        $("#matreg_betegn_" + otf_lager).val('')


        $("#FM_opdaterpris_" + otf_lager).attr('checked', false);

        t = $("#thisfile").val()
        //alert(t)
        if (t == "matreg") {
              
          senesteMatRegList()

       
        }



    }





    $("#sogliste").keyup(function () {

        //alert("her")
        senesteMatRegList();

    });

    $("#showonlypers").click(function () {

        //alert("her")
        senesteMatRegList();

    });

    $("#vasallemed").click(function () {

        //alert("her")
        senesteMatRegList();

    });



    function senesteMatRegList() {

        matreg_medid = $("#matreg_medid").val()
        sogliste = $("#sogliste").val()
        //sogBilagOrJob = //$("#sogBilagOrJob").val()
        if ($("#showonlypers").is(':checked') == true) {
            matreg_personlig = 1
        } else {
            matreg_personlig = 0
        }

        if ($("#vasallemed").is(':checked') == true) {
            matreg_visallemed = 1
        } else {
            matreg_visallemed = 0
        }

        //alert(matreg_visallemed)

        $.post("?matreg_medid=" + matreg_medid + "&matreg_personlig=" + matreg_personlig + "&matreg_visallemed=" + matreg_visallemed  + "&sogliste=" + sogliste, { control: "FN_seneste_mat", AjaxUpdateField: "true" }, function (data) {

            //alert($("#matreg_indlast").val())

            $("#mhi").html(data)


        });


    }




    
    $(".s_matgrp").change(function () {

        //alert("her")

        var thisid = this.id

        var idlngt = thisid.length
        var idtrim = thisid.slice(7, idlngt)

        beregnsalgsprisOTF_jq(idtrim)

    });

    $(".beregnsalgspris_txt").keyup(function () {

        var thisid = this.id

        //alert(thisid)

        var idlngt = thisid.length
        var idtrim = thisid.slice(5, idlngt)

        beregnsalgsprisOTF_jq(idtrim)

      

    });


    function beregnsalgsprisOTF_jq(id) {

        //alert("her: "+ id) // + id + "avaId " + avaId + " avaVal " + avaVal)

        avaId = $("#gruppe_" + id).val()

        //alert("her" + id + "avaId " + avaId)

        avaVal = $("#avagrpval_" + avaId).val().replace(",", ".")

      

        //$("#FM_salgspris_" + id).val(200)

        kobspris = $("#pris_" + id).val()
        kobspris = kobspris.replace(",", ".")

        //alert("kp: " + kobspris)

        nysalgspris = ((kobspris / 1 * avaVal) / 100 + kobspris / 1)
        nysalgspris = Math.round(nysalgspris * 100) / 100
        nysalgspris =  String(nysalgspris).replace(".", ",")

        //alert("sp: " + nysalgspris)

        $("#FM_salgspris_" + id).val(nysalgspris)


    }



    
   

    matreg_grp = $.cookie('matreg_grp');
    matreg_lev = $.cookie('matreg_lev');
    matreg_sog = $.cookie('matreg_sog');


    $("#sogmatgrp").val(matreg_grp)
    $("#soglev").val(matreg_lev)
    $("#FM_sog").val(matreg_sog)


}); //document.ready




function XberegnsalgsprisOTF(id) {

    alert("her") // + id + "avaId " + avaId + " avaVal " + avaVal)

    avaId = $("#gruppe_" + id).val()
    avaVal =  $("#avagrpval_" + avaId).val().replace(",", ".")

    kobspris = $("#pris_" + id).val().replace(",", ".")
    nysalgspris = ((kobspris / 1 * avaVal) / 100 + kobspris / 1)
    nysalgspris = Math.round(nysalgspris * 100) / 100

    nysalgspris = nysalgspris.replace(".", ",")
   

    $("#FM_salgspris_" + id).val(nysalgspris)


}




