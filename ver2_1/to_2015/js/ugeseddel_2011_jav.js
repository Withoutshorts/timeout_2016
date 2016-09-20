





$(document).ready(function() {


    $('.date').datepicker({

    });



    $(".status_all").click(function () {

        

        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(11, thisvallngt)
        thisval = thisvaltrim

        //alert("Her:" + thisval)

        if ($("#status_all_" + thisval).is(':checked') == true) {

            $(".status_chk_" + thisval).attr("checked", "checked");

            } else {

            $(".status_chk_" + thisval).removeAttr("checked");
           
        }


    });

 








    $("#FM_visallemedarb").click(function () {

       
        var sesid = $("#FM_sesMid").val()
        $("#usemrn").val(sesid)

        $("#filterkri").submit();


    });

    $("#usemrn").change(function () {
       
        $("#filterkri").submit();
    });

    

   

    $(".FM_job").keyup(function () {

       

        $(".dv_akt").hide();

        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(7, thisvallngt)
        thisval = thisvaltrim

        //alert(thisval)
        
        sogjobogkunde(thisval);

    });

    //$(".FM_job").blur()(function () {

    
    //    $(".dv_job").hide();
    //});

    $(".FM_akt").keyup(function () {

    

        $(".dv_job").hide();

       var thisid = this.id
       var thisvallngt = thisid.length
       var thisvaltrim = thisid.slice(7, thisvallngt)
       thisval = thisvaltrim

        sogakt(thisval);

    });


    $("#FM_timer").focus(function () {
        $(".chbox_job").hide();
        $(".chbox_akt").hide();
    });

    $("#FM_kom").focus(function () {
        $(".chbox_job").hide();
        $(".chbox_akt").hide();
    });


    function sogakt(thisval) {

        //alert(thisval)

        jq_newfilterval = $("#FM_akt_" + thisval).val()
        jq_jobid = $("#FM_jobid_" + thisval).val()
        jq_medid = $("#FM_medid").val()
        jq_aktid = $("#FM_akt_" + thisval).val()
        jq_pa = $("#FM_pa").val()

        varTjDatoUS_man = $("#varTjDatoUS_man").val()

        $(".chbox_job").hide();

        $("#dv_akt_" + thisval).css("display", "");
        $("#dv_akt_" + thisval).css("visibility", "visible");
        $("#dv_akt_" + thisval).show(100);

        //alert(jq_jobid + ":: medid: " + jq_medid + "#" + jq_newfilterval.length + " jq_aktid: " + jq_aktid + "## jq_pa" + jq_pa)

        if (jq_newfilterval.length > 0) {

            $.post("?jq_newfilterval=" + jq_newfilterval + "&jq_jobid=" + jq_jobid + "&jq_medid=" + jq_medid + "&jq_aktid=" + jq_aktid + "&jq_pa=" + jq_pa + "&varTjDatoUS_man=" + varTjDatoUS_man, { control: "FN_sogakt", AjaxUpdateField: "true" }, function (data) {
            //alert("cc")
            $("#dv_akt_" + thisval).html(data);

           



           // $('.luk_aktsog').bind('mouseover', function () {

            //     $(this).css("cursor", "pointer");



            //});

       

            // $(".luk_aktsog").bind('click', function () {


            //   $(".dv_akt").hide();

            // });


            // $('.chbox_akt').bind('mouseover', function () {

            //   $(this).css("cursor", "pointer");

            //  });

            $(".chbox_akt").bind('keyup', function () {

              
                if (window.event.keyCode == "13") {

                    var thisvaltrim = $("#dv_akt_" + thisval).val()
                    thisAktidUse = thisvaltrim
                    //thisJobnavn = $("#hiddn_job_" + thisJobid).val()
                    thisAktnavn = $("#dv_akt_" + thisval +" option:selected").text()



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
                thisAktnavn = $("#dv_akt_"+ thisval + " option:selected").text()

                

                $("#FM_akt_" + thisval).val(thisAktnavn)
                $("#FM_aktid_" + thisval).val(thisAktidUse)
                //$(".dv_akt").hide();
                $(".chbox_akt").hide();

            });


     

        });

        }

            

    }



  
    function sogjobogkunde(thisval) {

      

        jq_newfilterval = $("#FM_job_" + thisval).val()
        jq_medid = $("#FM_medid").val()


        $("#dv_job_" + thisval).css("display", "");
        $("#dv_job_" + thisval).css("visibility", "visible");
        $("#dv_job_" + thisval).show(100);

        varTjDatoUS_man = $("#varTjDatoUS_man").val()
       
        $(".chbox_akt").hide();

        //alert("thisval: " + thisval + " jq_newfilterval: " + jq_newfilterval + " jq_medid: " + jq_medid)

        if (jq_newfilterval.length > 0) {

            $.post("?jq_newfilterval=" + jq_newfilterval + "&jq_medid=" + jq_medid +"&varTjDatoUS_man=" + varTjDatoUS_man, { control: "FN_sogjobogkunde", AjaxUpdateField: "true" }, function (data) {
            //alert("cc")
            $("#dv_job_" + thisval).html(data);

           
            

            //$('.luk_jobsog').bind('mouseover', function () {

            //    $(this).css("cursor", "pointer");

            //});


            //$('.chbox_job').bind('mouseover', function () {

            //    $(this).css("cursor", "pointer");

            //});


            

               // $(".luk_jobsog").bind('click', function () {

              
            //     $(".dv_job").hide();

            //  });


            //  $(".luk_jobsog").bind('click', function () {


            //      $(".dv_job").hide();

            //  });



                
            $(".chbox_job").bind('click', function () {

             
                    
                var thisvaltrim = $("#dv_job_" + thisval).val()
                    //alert("her: " + thisvaltrim)
                   
                    //var thisjobid = this.id
                    //var thisvallngt = thisjobid.length
                    //var thisvaltrim = thisjobid.slice(10, thisvallngt)
                    //var thisvaltrim = thisvaltrim
                   

                    thisJobid = thisvaltrim
                    //thisJobnavn = $("#hiddn_job_" + thisJobid).val()
                    thisJobnavn = $("#dv_job_"+ thisval +" option:selected").text()
                    
                   

                    $("#FM_job_" + thisval).val(thisJobnavn)
                    $("#FM_jobid_" + thisval).val(thisJobid)

                    //$(".dv_job").hide();
                    $(".chbox_job").hide();

            });
                    



                $(".chbox_job").bind('keyup', function () {

                    if (window.event.keyCode == "13") {
                    
                        var thisvaltrim = $("#dv_job_" + thisval).val()
                    //alert("her: " + thisvaltrim)
                   
                    //var thisjobid = this.id
                    //var thisvallngt = thisjobid.length
                    //var thisvaltrim = thisjobid.slice(10, thisvallngt)
                    //var thisvaltrim = thisvaltrim
                   

                    thisJobid = thisvaltrim
                    //thisJobnavn = $("#hiddn_job_" + thisJobid).val()
                    thisJobnavn = $("#dv_job_"+ thisval +" option:selected").text()
                    
                   

                    $("#FM_job_" + thisval).val(thisJobnavn)
                    $("#FM_jobid_" + thisval).val(thisJobid)

                    //$(".dv_job").hide();
                    $(".chbox_job").hide();

                    }
                    

                });


        
                
        });

        }
    }

    
   
    

    dagsval1 = $("#sumtimer_dag_1").val();
    dagsval2 = $("#sumtimer_dag_2").val();
    dagsval3 = $("#sumtimer_dag_3").val();
    dagsval4 = $("#sumtimer_dag_4").val();
    dagsval5 = $("#sumtimer_dag_5").val();
    dagsval6 = $("#sumtimer_dag_6").val();
    dagsval7 = $("#sumtimer_dag_7").val();

    if (dagsval1 != "0,00") {
        $("#sp_sumtimer_dag_1").html(dagsval1);
    }

    if (dagsval2 != "0,00") {
        //alert(dagsval2)
        $("#sp_sumtimer_dag_2").html(dagsval2);
    }

    if (dagsval3 != "0,00") {
        $("#sp_sumtimer_dag_3").html(dagsval3);
    }
    if (dagsval4 != "0,00") {
        $("#sp_sumtimer_dag_4").html(dagsval4);
    }
        if (dagsval5 != "0,00") {
            $("#sp_sumtimer_dag_5").html(dagsval5);
        }
            if (dagsval6 != "0,00") {
                $("#sp_sumtimer_dag_6").html(dagsval6);
            }
                if (dagsval7 != "0,00") {
                $("#sp_sumtimer_dag_7").html(dagsval7);
                }
   

});

$(window).load(function() {
    // run code
    $("#loadbar").hide(1000);
});