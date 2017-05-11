




$(document).ready(function () {


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

       
        $("#dv_job").css("display", "");
        $("#dv_job").css("visibility", "visible");
        $("#dv_job").show(100);

       //alert("medid" + jq_newfilterval)

        $.post("?jq_newfilterval=" + jq_newfilterval + "&jq_medid=" + jq_medid, { control: "FN_sogjobogkunde", AjaxUpdateField: "true" }, function (data) {
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
                
            });


            $(".chbox_job").bind('keyup', function () {

                if (window.event.keyCode == "13") {

                    var thisvaltrim = $("#dv_job").val()

                    thisJobid = thisvaltrim
                    //thisJobnavn = $("#hiddn_job_" + thisJobid).val()
                    thisJobnavn = $("#dv_job" + " option:selected").text()



                    $("#FM_job").val(thisJobnavn)
                    $("#FM_jobid").val(thisJobid)

                    //$(".dv_job").hide();
                    $(".chbox_job").hide();

                }

               
            });

        });
    }


    $(".aktivitet_sog").keyup(function () {

        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(7, thisvallngt)
        thisval = thisvaltrim
        //alert(thisid)
        sogakt(thisval);

    });



    function sogakt(thisid) {

       
        jq_newfilterval = $("#FM_akt").val()
              
        jq_medid = $("#FM_medid").val()
        jq_aktid = $("#FM_akt").val()
        jq_pa = $("#FM_pa").val()

        jq_jobid = $("#FM_jobid").val()

        $("#dv_akt").css("display", "");
        $("#dv_akt").css("visibility", "visible");
        $("#dv_akt").show(100);
        
       // alert(jq_jobid)
        //alert("medid: " + jq_medid + "jobid: " + jq_jobid)

        $.post("?jq_newfilterval=" + jq_newfilterval + "&jq_jobid=" + jq_jobid + "&jq_medid=" + jq_medid + "&jq_aktid=" + jq_aktid + "&jq_pa=" + jq_pa, { control: "FN_sogakt", AjaxUpdateField: "true" }, function (data) {
           // alert("cc")
            //$("#dv_akt_test").html(data);
            $("#dv_akt").html(data);

            //alert("END")


                $(".chbox_akt").bind('keyup', function () {


                    if (window.event.keyCode == "13") {

                        var thisvaltrim = $("#dv_akt").val()
                        thisAktidUse = thisvaltrim
                        //thisJobnavn = $("#hiddn_job_" + thisJobid).val()
                        thisAktnavn = $("#dv_akt" + " option:selected").text()

                        $("#FM_akt").val(thisAktnavn)
                        $("#FM_aktid").val(thisAktidUse)
                        //$(".dv_akt").hide();
                        $(".chbox_akt").hide();

                    }

                });


                $(".chbox_akt").bind('click', function () {


                    var thisvaltrim = $("#dv_akt").val()
                    thisAktidUse = thisvaltrim
                    //thisJobnavn = $("#hiddn_job_" + thisJobid).val()
                    thisAktnavn = $("#dv_akt" + " option:selected").text()

                    $("#FM_akt").val(thisAktnavn)
                    $("#FM_aktid").val(thisAktidUse)
                    //$(".dv_akt").hide();
                    $(".chbox_akt").hide();

                });

            




        });


    }

    var clicks = 1;

    $(".tilfoj_akt").click(function () {

        //thisid = this.id
       
        jobid = $("#FM_jobid").val()
        aktid = $("#FM_aktid").val()
        medid = $("#FM_medid_id").val()

        jobnavn = $("#FM_job").val()
        aktnavn = $("#FM_akt").val()

        next_akt_id = $("#next_akt_id").val()

        //alert (medid)
        
        //alert(thisid)

       // next_akt_id = $(".next_akt_id").attr('id')
       // var akt_id_lngt = next_akt_id.length
       // var akt_id_trim = next_akt_id.slice(16, akt_id_lngt)
       // this_next_akt_id = akt_id_trim

       // alert("id " + next_akt_id)

       
        

        //alert(medid)

        $.post("?FM_jobid=" + jobid + "&FM_aktid=" + aktid + "&FM_medid_id=" + medid, { control: "tilfoj_akt", AjaxUpdateField: "true" }, function (data) {
            //alert("fe")
            $("#dv_akttil").html(data);

            
            //alert("cc")

            //next_akt_id = $(".next_akt_id").val()

            $("#FN_akt_tilfojed_" + next_akt_id).css("display", "");
            $("#FN_akt_tilfojed_" + next_akt_id).css("visibility", "visible");

            $('input[id="next_akt_jobid_' + next_akt_id + '"]').val(jobnavn);
            $('input[id="next_akt_aktid_' + next_akt_id + '"]').val(aktnavn);
            
            clicks++;

            if (clicks > 11)
            {
                alert("Klik opdater inden du opretter flere aktiviteter")
            }

            //alert(clicks)

            $("#next_akt_id").val(clicks)
            
            $("#FM_job").val("")
            $("#FM_akt").val("")
            
            
        });


    });



});



