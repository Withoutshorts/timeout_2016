




$(document).ready(function () {


    

    $(".xtimer_flt").click(function () {

        //alert("HER")
        //tfaktim, felt, chk?, række, kid

        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(9, thisvallngt)
        thisval = thisvaltrim

        jq_flt = $("#jq_flt_" + thisval).val() 
        jq_row = $("#jq_row_" + thisval).val()
        jq_kid = $("#jq_kid_" + thisval).val() 
        jq_tfa = $("#jq_tfa_" + thisval).val() 
        $("#FM_sog_kpers_dist_kid").val(jq_kid) 
        //alert("HER: " + this.id + " jq_kid: " + jq_kid)

       

        showKMdailog(jq_tfa, jq_flt, 0, jq_row, jq_kid)

        

    });

    /// Kørsel //
    $("#jq_hideKpersdiv").click(function () {
        //$.scrollTo('650px', 500);
        $("#korseldiv").hide('fast')

    });


    $('#jq_hideKpersdiv').mouseover(function () {

        $(this).css("cursor", "pointer");

    });

    // Kpers og filialer til Kørsel //
    //$("#FM_sog_kpers_but").click(function () {
    //GetFiliogKpers(1);
    //});

    //$("#FM_sog_kpers_dist_all").click(function() {
    //    //alert("her")
    //    GetFiliogKpers();
    //});


    $("#indlaes_koadr_2").click(function () {
        //alert("her")
        koadr_2013g();

    });




    function koadr_2013g() {
        //alert("her")
        //xval = xval/1



        var xrow = $("#koFlt").val();
        var dagn = $("#koFltx").val();

        flt = xrow + "" + dagn

        var destination = 0;
        destination = document.getElementById("ko1").value;
        //alert(destination)
        $("#FM_destination_" + flt).val(destination)


        //alert(flt)

        varTurTypeTxtFra = document.getElementById("ko0")
        varTurTypeTxtFra = varTurTypeTxtFra.options[varTurTypeTxtFra.selectedIndex].text.replace("..........", "")


        varTurTypeTxtTil = document.getElementById("ko1")
        varTurTypeTxtTil = varTurTypeTxtTil.options[varTurTypeTxtTil.selectedIndex].text.replace("..........", "")



        var kotxt = "Fra:\n" + varTurTypeTxtFra + "\n\nTil:\n" + varTurTypeTxtTil

        //alert("#FM_kom_" + flt + "")
        $("#FM_kom_" + flt + "").val(kotxt)
        $("#FM_kommentar").val(kotxt)



        //.replace("#br#", vbCrLf)
        //.replace("<BR>"," vbcrlf ")

        $("#korseldiv").hide('fast')

        // Call Lei GoogleMap function



        Lei.Indlaes(kotxt, flt);



        // 


    }




    var Lei = {
        Indlaes: function (fmKom, flt) {
            var kom = "";
            if (fmKom !== null && fmKom !== undefined)
                kom = fmKom;
            kom = encodeURIComponent(kom);
            window.open("../timereg/CalculateDistance.aspx?comments=" + kom + "&thisfile=favorit&flt=" + flt, "_blank", "width=500,height=200,fullscreen=no", "true");
        }


    }



    // Kørsel Henter kunder (desitnationer)
    function GetFiliogKpers(id) {

        //alert(id)

        x = 100 //$("#FM_timer_" + id).offset();


        //var thisid = id
        //var thisvallngt = thisid.length

        //var dagn = thisid.slice(thisvallngt - 1, thisvallngt)
        //var xrow = thisid.slice(0, thisvallngt - 2)

        //alert(xrow +"_"+ dagn)

        var kid = $("#FM_sog_kpers_dist_kid").val();
        //alert(kid.val())

        var jqmid = $("#usemrn").val();
        var visalle = 1;

       

        koKmDialog = document.getElementById("koKmDialog").value

        //alert("her" + jqmid + " kid: " + kid + " koKmDialog:" + koKmDialog)

        if (koKmDialog == 1) {

            //$("#BtnCustDescUpdate").data("cust", thisC.val());
            $.post("?kid=" + kid + "&visalle=" + visalle + "&jqmid=" + jqmid, { control: "FM_get_destinations", AjaxUpdateField: "true", cust: 0 }, function (data) {
                //$("#koTEST").html(data);

                //$("#fajl").html(data);


                $("#ko0").html(data);
                $("#ko1").html(data);
             
                $("#koFlt").val(id);

                //$("#koFlt").val(xrow);
                //$("#koFltx").val(dagn);
                
            });
        }



        topPx = parseInt(x.top)

        //alert(topPx)

        $("#korseldiv").css("top", topPx);
        $("#korseldiv").css("left", 145);

        $.scrollTo("#korseldiv", 300, { offset: -400 });


    }



    function showKMdailog(val, flt, chk, rk, kid) {

        //alert("her")
        document.getElementById("FM_sog_kpers_dist_kid").value = kid
        komLengthVal = "" //document.getElementById("FM_kom_" + flt + "").value
        komLength = komLengthVal.length
      
        lastKmdiv = 0 //document.getElementById("lastkmdiv").value
        koKmDialog = document.getElementById("koKmDialog").value

        if (koKmDialog == 1) {

            if (komLength == 0) {

                if (val == 5) {



                  
                   

                    //document.getElementById("korseldiv").style.top = 700 //+ (flt*2.5) 

                    //document.getElementById("korseldiv").style.left = 225

                    //document.getElementById("korseldiv").style.visibility = "visible"
                    //document.getElementById("korseldiv").style.display = ""
                    //document.getElementById("kpersdiv").style.visibility = "hidden"
                    //document.getElementById("kpersdiv").style.display = "none"


                    $("#korseldiv").css("visibility", "visible");
                    $("#korseldiv").css("display", "");

                    //var feltid = rk + flt;

                    //alert(feltid)
                    //GetFiliogKpers(feltid);
                    GetFiliogKpers(rk);


                    //document.getElementById("koVisible").selected = 1
                    //document.getElementById("ko0chk").focus()

                    //$("#ko0chk").focus();
                    //document.getElementById("koVisible").selectedIndex = 1
                    //varTurTypeTxt.options[varTurTypeTxt.selectedIndex].selected = 1

                    //document.getElementById("koFlt").value = flt
                    //$("#koFlt").val(flt);



                    //document.getElementById("FM_sog_kpers_but").focus()

                 



                }
            }
        }



    }


  


    /// SLUT kørsel






    

    $('.date').datepicker({
        
    });


    $(".FM_job").keyup(function () {

        //alert("keupup")
        $(".aktivitet_sog").val("")

        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(7, thisvallngt)
        thisval = thisvaltrim

        //alert(thisid)

        sogjobogkunde(thisid);



    });




    function sogjobogkunde(thisid) {

        //alert("søg")

        jq_newfilterval = $("#FM_job").val()
        jq_medid = $("#FM_medid").val()

        ign_projgrp_hd = $("#ign_projgrp_hd").val()

        if (ign_projgrp_hd == 1) {

            if ($('#ign_projgrp').is(':checked')) {
                jq_ign_projgrp = 1
            } else {
                jq_ign_projgrp = 0
            }//$("#ign_projgrp").val()

        } else {
            jq_ign_projgrp = 0
        }

        $("#dv_job").css("display", "");
        $("#dv_job").css("visibility", "visible");
        $("#dv_job").show(100);

       //alert("medid" + jq_newfilterval)

        $.post("?jq_newfilterval=" + jq_newfilterval + "&jq_ign_projgrp=" + jq_ign_projgrp  +"&jq_medid=" + jq_medid, { control: "FN_sogjobogkunde", AjaxUpdateField: "true" }, function (data) {
            //alert("cc")
            $("#dv_job").html(data);


            $(".chbox_job").bind('click', function () {



                var thisvaltrim = $("#dv_job").val()
                //alert("her: " + thisvaltrim)


                thisJobid = thisvaltrim
                //thisJobnavn = $("#hiddn_job_" + thisJobid).val()
                thisJobnavn = $("#dv_job" + " option:selected").text()

                $("#FM_job").val(thisJobnavn)
                $("#FM_jobid").val(thisJobid)

                //$(".dv_job").hide();
                $(".chbox_job").hide();

                sogakt();
                
            });

           

        });
    }


    



    function sogakt() {

       
        
        jq_newfilterval = $("#FM_akt").val()
              
        jq_medid = $("#FM_medid").val()
        FM_medid = $("#FM_medid").val()
        jq_aktid = $("#FM_akt").val()
        jq_pa = $("#FM_pa").val()

        jq_jobid = $("#FM_jobid").val()
        FM_jobid = $("#FM_jobid").val()

        jq_ign_projgrp = $("#ign_projgrp").val()

        //alert(jq_jobid)
        //alert(FM_jobid)
        /*
        $("#dv_akt").css("display", "");
        $("#dv_akt").css("visibility", "visible");
        $("#dv_akt").show(100);
        */
        
       // alert(jq_jobid)
        //alert("medid: " + jq_medid + "jobid: " + jq_jobid)

        $.post("?jq_newfilterval=" + jq_newfilterval + "&jq_ign_projgrp=" + jq_ign_projgrp +"&jq_jobid=" + jq_jobid + "&jq_medid=" + jq_medid + "&jq_aktid=" + jq_aktid + "&jq_pa=" + jq_pa + "&FM_medid=" + FM_medid + "&FM_jobid=" + FM_jobid, { control: "FN_sogakt", AjaxUpdateField: "true" }, function (data) {


            $("#dv_akt").html(data);

          





        });


    }

    var clicks = 1;

    $(".tilfoj_akt").click(function () {

        //thisid = this.id
       
        jobid = $("#FM_jobid").val()
        aktid = $("#dv_akt").val()//$("#FM_aktid").val()
        medid = $("#FM_medid_id").val()

        jobnavn = $("#FM_job").val()
        aktnavn = $("#dv_akt").text()

        next_akt_id = $("#next_akt_id").val()

       

        //alert(medid + " j:" + jobid + " a:"+ aktid )

        $.post("?FM_jobid=" + jobid + "&FM_aktid=" + aktid + "&FM_medid_id=" + medid, { control: "tilfoj_akt", AjaxUpdateField: "true" }, function (data) {
            //alert("fe")
            $("#dv_akttil").html(data);

            
            $("#favorit").submit();


         
            
        });


    });


    $(".picmodal").click(function () {

        var modalid = this.id
        var idlngt = modalid.length
        var idtrim = modalid.slice(6, idlngt)

        //var modalidtxt = $("#myModal_" + idtrim);
        var modal = document.getElementById('myModal_' + idtrim);


        if (modal.style.display !== 'none') {
            modal.style.display = 'none';
        }
        else {
            modal.style.display = 'block';
        }


    });


    $(".kommodal").click(function () {

        //alert("klik")

        var modalid = this.id
        var idlngt = modalid.length
        var idtrim = modalid.slice(6, idlngt)

        //var modalidtxt = $("#myModal_" + idtrim);
        var modal = document.getElementById('kommentarmodal_' + idtrim);

        modal.style.display = "block";

        window.onclick = function (event) {
            if (event.target == modal) {
                modal.style.display = "none";
            }
        }

    });



    $(".jobinfo").click(function () {

        var modalid = this.id
        var idlngt = modalid.length
        var idtrim = modalid.slice(8, idlngt)

        //var modalidtxt = $("#myModal_" + idtrim);
        var jobmodal = document.getElementById('jobmodal_' + idtrim);

        jobmodal.style.display = "block";

        window.onclick = function (event) {
            if (event.target == jobmodal) {
                jobmodal.style.display = "none";
            }
        }


    });



});


function SetValueInParent() {

   

    var flt = window.location.search;

   

    var index = flt.indexOf("&flt");
    if (index != -1)
        flt = flt.substr(index + 5);
    flt = flt.substr(0, flt.length - 1) //+ "_" + flt.substr(flt.length - 1);
    //alert("flt: " + flt)
    var fmTimer = window.opener.document.getElementById("FM_timer_" + flt);
    if (fmTimer !== null)
        fmTimer.value = document.getElementById("lblKM").innerHTML.replace("km", "").replace(" ", "");

    if (fmTimer.value.indexOf("m") > -1) {
        var vm = fmTimer.value.replace("m", "");
        fmTimer.value = vm / 1000;
    }


    window.close();

    return false;
}



