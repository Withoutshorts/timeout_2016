



function isNum_treg(passedVal) {
    invalidChars = " /:;<>abcdefghijklmnopqrstuvwxyzæøå"

    //alert("her")

    if (passedVal == "") {
        return false
    }

    for (i = 0; i < invalidChars.length; i++) {
        badChar = invalidChars.charAt(i)
        if (passedVal.indexOf(badChar, 0) != -1) {
            return false
        }
    }

    for (i = 0; i < passedVal.length; i++) {
        if (passedVal.charAt(i) == "." || passedVal.charAt(i) == "-") {
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






$(document).ready(function() {


    //approved
    $("#sbm_timer").click(function () {

        // Skal ikke tilføje logud tid, da det skal kunne indtastes manuelt //
        //if ($("#showstop").val() == "1") {
        //    ststop();
        //}

        //$("#dvindlaes_msg").attr('class', 'approved');
        $("#dv_mat_container").hide(10)

        $("#dvindlaes_msg").css("display", "");
        $("#dvindlaes_msg").css("visibility", "visible");
        $("#dvindlaes_msg").show(100);

    });


    
    $("#FM_sltid").keyup(function () {

        ststop(1);

    });
  

    
    $("#bt_stst").click(function () {

        ststop(0);
       
    });


    function ststop(io) {

        //var sttid = new Date();
        //var sltid = new Date();

        sttid = $("#FM_sttid").val()

        var idtrim = sttid.slice(1, 2)
        if (idtrim == ":") {
            $("#FM_sttid").val("0" + sttid)
            sttid = $("#FM_sttid").val()
        }


        sltid = $("#FM_sltid").val()

        /*var idtrim = sltid.slice(1, 2)
        if (idtrim == ":") {
            $("#FM_sltid").val("0" + sltid)
            sltid = $("#FM_sltid").val()
        }*/


        if (sttid.length == 0 || sttid == "00:00") {

           
           

            var dt = new Date();

            var minutter = dt.getMinutes()
            if (minutter < 10) {
                minutter = 0 + "" + minutter
            }

            var sttimeNow = dt.getHours() + ":" + minutter // + ":" + dt.getSeconds();

            $("#FM_sttid").val(sttimeNow)
            $("#FM_sttid").css("color", "#000000")

            $("#FM_sltid").val(sttimeNow)


            sttid = $("#FM_sttid").val()


        } else {

            if (io == 0) {

                var dt = new Date();

                var minutter = dt.getMinutes()
                if (minutter < 10) {
                    minutter = 0 + "" + minutter
                }

                var sltimeNow = dt.getHours() + ":" + minutter // + ":" + dt.getSeconds();

                $("#FM_sltid").val(sltimeNow)
                $("#FM_sltid").css("color", "#000000")
                sltid = $("#FM_sltid").val()

            }


        }

        //var dt = new Date();
        //dt.getDay() + ":" + minutter
        //alert(sttid)

        //dt1 = new Date(2001, 1, 1); //UTC time
        //alert(dt1.getTime()); //gives back timestamp in ms

        dt1 = new Date("1/1/2001 "+ sttid + ":00")
        dt2 = new Date("1/1/2001 "+ sltid + ":00")


        //alert(dt1 +"//"+ dt2)

        //var beginTime = stringToDate(dt1);
        //var endTime = stringToDate(dt1);

        //alert(beginTime)
        if (sttid < sltid) {
            var sp = (((dt2 - dt1) / 60000));
        } else {
            var sp = 0
        }


       
        sp = sp / 60
      
     

        if (sp > 0) {
            $("#FM_timer").css("color", "#000000")
        }

        sp = Math.round((parseFloat(sp)) * 100) / 100
        var spTxt = String(sp).replace(".",",")

        $("#FM_timerlbl").html(spTxt)
        $("#FM_timer").val(spTxt)
       
        //alert("timeuse：" + sp.hour + " hour " + sp.minute + " minute " + sp.second + " second 


    }


    $("#sbmmat").click(function () {

        nytMatantal = $("#FM_matantal").val()

       
        if (isNum_treg(nytMatantal)) {
            //return true
        } else {

            nytMatantal = 0
            $("#FM_matantal").val(0)
            return false 
        }
       
       

        

        if (nytMatantal.length > 0 && nytMatantal > 0 && $("#FM_matnavn").val() != 'Tilføj Materiale' && $("#FM_matnavn").val() != 'flere?' && $("#FM_matnavn").val() != '') {

            nytMatnavn = $("#FM_matantal").val() + " " + $("#FM_matnavn").val()

            prevIndlast = $("#dv_mat_sbm").html();

            if (prevIndlast.length > 0) {
                $("#dv_mat_sbm").html(nytMatnavn + ", " + prevIndlast);
            } else {
                $("#dv_mat_sbm").html(nytMatnavn);
            }


            prevMatids = $("#FM_matids").val();
            nytMatid = $("#FM_matid").val();

            if (prevMatids.length > 0) {
                $("#FM_matids").val(prevMatids + ", " + nytMatid);
            } else {
                $("#FM_matids").val(nytMatid);
            }


            prevMatAntals = $("#FM_matantals").val();


            //alert(nytMatantal)

            if (prevMatAntals.length > 0) {
                $("#FM_matantals").val(prevMatAntals + ", " + nytMatantal);
            } else {
                $("#FM_matantals").val(nytMatantal);
            }

            $("#FM_matantal").val('')
            $("#FM_matnavn").val('flere?')

        }
    
    });

    

    $("#a_mat").click(function () {


        if ($("#dv_mat_container").css('display') == "none") {
            $("#dv_mat_container").css("display", "");
            $("#dv_mat_container").css("visibility", "visible");
            $("#dv_mat_container").show(100);
        } else {
            $("#dv_mat_container").css("display", "none");
            $("#dv_mat_container").css("visibility", "hidden");
            $("#dv_mat_container").hidden(100);
        }

    });


    $("#FM_matnavn").focus(function () {

        if ($("#FM_matnavn").val() == 'Tilføj Materiale' || $("#FM_matnavn").val() == 'flere?') {
            $("#FM_matnavn").val('')
        }
    });
   

    $("#FM_matantal").focus(function () {

        if ($("#FM_matantal").val() == 'Ant.') {
            $("#FM_matantal").val('')
        }
    });


    $("#FM_job").focus(function () {

        //if ($("#FM_job").val() == 'Kunde/job') {
            $("#FM_job").val('')
        //}
    });


    $("#FM_akt").focus(function () {

        //if ($("#FM_akt").val() == 'Aktivitet') {
            $("#FM_akt").val('')
        //}
    });


    $("#FM_timer").focus(function () {

        if ($("#FM_timer").val() == 'Antal timer') {
            $("#FM_timer").val('')
        }
    });


    $("#FM_kom").focus(function () {

        if ($("#FM_kom").val() == 'Kommentar') {
            $("#FM_kom").val('')
        }
    });


    $("#FM_job").keyup(function () {


       // if ($("#showstop").val() == "1") {

        //     ststop();

        //     timerThis = $("#FM_timer").val().replace(",", ".")

        //    if (timerThis > 0) {
        //        $("#container").submit();
        //    }
        //}

        $("#dv_akt").hide();
        $("#dv_mat").hide();
        
        sogjobogkunde();



    });


    $("#FM_matnavn").keyup(function () {

        //alert("gdgf")

        $("#dv_job").hide();
        $("#dv_akt").hide();
        
        sogMat();



    });
    

    $("#FM_akt").keyup(function () {

        $("#dv_job").hide();
        $("#dv_mat").hide();

        sogakt();

    });



    function sogMat() {

        

        jq_newfilterval = $("#FM_matnavn").val()
       

       
        if (jq_newfilterval.length > 0) {



            //alert("cc")

            $.post("?jq_newfilterval=" + jq_newfilterval, { control: "FN_sogmat", AjaxUpdateField: "true" }, function (data) {
                //alert("cc")


                $("#dv_mat").html(data);

                //alert("her")
                //$("#dv_akt").attr('class', 'dv-open');

                if ($("#dv_mat").css('display') == "none") {
                    $("#dv_mat").css("display", "");
                    $("#dv_mat").css("visibility", "visible");
                    $("#dv_mat").show(100);
                }



                $('#luk_matsog').bind('mouseover', function () {

                    $(this).css("cursor", "pointer");

                });



                $("#luk_matsog").bind('click', function () {


                    $("#dv_mat").hide();
                    //$("#dv_akt").attr('class', 'dv-closed');

                });





                $('.span_mat').bind('mouseover', function () {

                    $(this).css("background-color", "#CCCCCC");
                    $(this).css("cursor", "pointer");

                });

                $('.span_mat').bind('mouseout', function () {

                    $(this).css("background-color", "#FFFFFF");

                });


                $(".span_mat").bind('click', function () {



                    var thismatid = this.id
                    var thisvallngt = thismatid.length
                    var thisvaltrim = thismatid.slice(10, thisvallngt)


                    thismatid = thisvaltrim
                    thismatnavn = $("#hiddn_mat_" + thismatid).val()

                    //alert(thismatnavn + " " + thismatid)

                    $("#FM_matnavn").val(thismatnavn)
                    $("#FM_matid").val(thismatid)
                    $("#dv_mat").hide();
                    //$("#dv_mat").attr('class', 'dv-closed');

                });


            });

        } else {
            $("#dv_mat").hide();
        }

    }



    function sogakt() {

        

        jq_newfilterval = $("#FM_akt").val()
        jq_jobid = $("#FM_jobid").val()
        jq_medid = $("#FM_medid").val()
        jq_aktid = $("#FM_akt").val()
        jq_pa = $("#FM_pa").val()
        //alert(jq_medid)

       
        if (jq_newfilterval.length > 0) {



            //alert("cc")

            $.post("?jq_newfilterval=" + jq_newfilterval + "&jq_jobid=" + jq_jobid + "&jq_medid=" + jq_medid + "&jq_aktid=" + jq_aktid + "&jq_pa=" + jq_pa, { control: "FN_sogakt", AjaxUpdateField: "true" }, function (data) {
                //alert("cc")


                $("#dv_akt").html(data);

                //alert("her")
                //$("#dv_akt").attr('class', 'dv-open');

                if ($("#dv_akt").css('display') == "none") {
                    $("#dv_akt").css("display", "");
                    $("#dv_akt").css("visibility", "visible");
                    $("#dv_akt").show(100);
                }




                $('#luk_aktsog').bind('mouseover', function () {

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




            });

        } else {

            $("#dv_akt").hide();
        }

            

    }



  
    function sogjobogkunde() {

        
        

        

        jq_newfilterval = $("#FM_job").val()
        jq_medid = $("#FM_medid_k").val()
       
        //alert(jq_newfilterval.length)
        //alert("cc" + jq_newfilterval)
        if (jq_newfilterval.length > 0) {

            $.post("?jq_newfilterval=" + jq_newfilterval + "&jq_medid=" + jq_medid, { control: "FN_sogjobogkunde", AjaxUpdateField: "true" }, function (data) {








                $("#dv_job").html(data);

                //$("#dv_job").attr('class', 'dv-open');

                if ($("#dv_job").css('display') == "none") {

                    $("#dv_job").css("display", "");
                    $("#dv_job").css("visibility", "visible");
                    $("#dv_job").show(100);

                }




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
                    //$("#dv_job").attr('class', 'dv-closed');


                    ststop(1);


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



    


    setTimeout(function () {
        // Do something after 5 seconds
        $("#timer_indlast").hide(1000);
    }, 2500);
    


    if ($("#tomobjid").val() != 0) {

        

        
        thisJobid = $("#tomobjid").val()
        $("#FM_jobid").val(thisJobid)

        $.post("?jq_jobid=" + thisJobid, { control: "FN_tomobjid", AjaxUpdateField: "true" }, function (data) {




            $("#FM_job").val(data)


        });

           

    }



});

$(window).load(function() {
    // run code
    
});