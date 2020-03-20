


function visSubtotJobmed() {

    if ($.cookie("vissimpel") == '1') {
        $(".jobtotal").hide(1000);
        $(".medtotal").hide(1000);

        $("#FM_vis_simpel").attr("checked", "checked");

    } else {

        $(".jobtotal").css("display", "");
        $(".jobtotal").css("visibility", "visible");
        $(".jobtotal").show(2000);

        $(".medtotal").css("display", "");
        $(".medtotal").css("visibility", "visible");
        $(".medtotal").show(2000);

        $("#FM_vis_simpel").removeAttr("checked");
    }



  

}


function disVisjobuforecast() {
    
    ialt = 0
    $("#FM_medarb option").each(function () {
        ialt += 1;
    });

    var count = $("#FM_medarb :selected").length;
    var selectedVal = $("#FM_medarb").val();

    if (count > 10 || (ialt > 10 && selectedVal == 0)) {
        $("#FM_vis_job_u_fc").attr('checked', false);
        $("#FM_vis_job_u_fc").attr("disabled", true);
    } else {
        $("#FM_vis_job_u_fc").attr("disabled", false);
    }


}

function DisJobSel() {

    var SELcount = $("#FM_jobsel :selected").length;
    SELcount = parseInt(SELcount);

    $("#sogtxt").val('')
    if ($("#FM_jobsel").val() == '0' || SELcount > 5) {
        $("#FM_vis_job_u_fc").prop("checked", false);
        $("#FM_vis_job_u_fc").attr('disabled', true)
    }
    else {
        $("#FM_vis_job_u_fc").attr('disabled', false)
    }
}

function DisMedSel()
{
    ialt = 0
    $("#FM_medarb2 option").each(function () {
        ialt += 1;
    });

    var count = $("#FM_medarb2 :selected").length;
    var selectedVal = $("#FM_medarb2").val();

    if (count > 10 || (ialt > 10 && selectedVal == 0)) {
        $("#FM_vis_job_u_fc").attr('checked', false);
        $("#FM_vis_job_u_fc").attr("disabled", true);
    } else {
        $("#FM_vis_job_u_fc").attr("disabled", false);
    }
}

$(window).load(function () {

    //visSubtotJobmed();

});


    $(document).ready(function () {

        disVisjobuforecast();

        var ddcView = $("#ddcView").val();
        var sortering = $("#FM_sortering").val();
        
        if (ddcView == 1 && sortering == 0) {
            DisJobSel();
        }

        if (ddcView == 1 && sortering == 1)
        {
            DisMedSel();
        }

        $("#loadfilter").click(function () {

            var jobsel = $("#FM_jobsel").val()
            var medsel = $("#FM_medarb2").val()
          //  alert(medsel)
            if (jobsel == "0" && medsel == "0") {
                $('#FM_medarb2 option').prop('selected', true);
                $("#FM_medarb2 option:first").prop('selected', false);
            }

            $("#mdselect").submit();
        
        });


        // Tilføj nyt medarb felt start
        var strNewline = "";
        numoffdaysorweeksinperiode = $("#numoffdaysorweeksinperiode").val();
        numoffdaysorweeksinperiode = parseInt(numoffdaysorweeksinperiode)

        var startMD = $("#startdatoMD").val();
        startMD = parseInt(startMD)
        var startYY = $("#startdatoYY").val();
        startYY = parseInt(startYY)
        
        var slutMD = $("#slutdatoMD").val();
        slutMD = parseInt(slutMD);
        var slutYY = $("#slutdatoYY").val();
        slutYY = parseInt(slutYY);

        var feltnrTilfojet = 0;
        var btnReady = 1;        
        $(".addField").click(function () {
            if (btnReady == 1) {
                btnReady = 0;
                thisid = this.id
                strNewLine = "<tr class='fieldfor_" + thisid + " columns'>"
                feltnrTilfojet += 1
                               
                loopYY = startYY
                loopMD = startMD

                var jobid = $(this).data('jobid');
                var aktid = -1 // Henter først aktiviteter på job
                $.post("?jobid=" + jobid, { control: "HentAktiviteter", AjaxUpdateField: "true" }, function (data) {

                    aktid = data
                    $(".aktid_" + feltnrTilfojet).val(aktid)
                    $(".aktid_old_" + feltnrTilfojet).val(aktid)

                    // Henter Medarbejdere
                    $.post("?jobid=" + jobid + "&aktid=" + aktid + "&startmd=" + startMD + "&startYY=" + startYY + "&slutMD= " + slutMD + "&slutYY=" + slutYY, { control: "HentMedarbejdere", AjaxUpdateField: "true" }, function (data) {
                       // alert(data)
                        $("#medarbSEL_" + feltnrTilfojet).html(data)
                        btnReady = 1;
                    });

                });


                strNewLine += "<td><select class='form-control input-small medarbSEL' id='medarbSEL_" + feltnrTilfojet + "' data-feltnr='" + feltnrTilfojet + "' ><option value='-1'>Henter</option></select></td>"

                var i = 0
                while (i < numoffdaysorweeksinperiode + 1) {
                    if (i != 0) {
                        if (loopMD == 12) {
                            loopMD = 1
                            loopMD = loopYY + 1
                        }
                        else {
                            loopMD += 1
                        }
                    }

                    strNewLine += "<td>"
                    strNewLine += "<input type='hidden' name='FM_dato' value='1-" + loopMD + "-" + loopYY + "' />"
                    strNewLine += "<input type='hidden' name='FM_medarbid' class='medarb_for_" + feltnrTilfojet + "' value='0' />"
                    strNewLine += "<input type='hidden' name='FM_jobid' value='" + jobid + "' />"
                    strNewLine += "<input type='hidden' name='FM_aktid' class='aktid_" + feltnrTilfojet + "' value='' />"
                    strNewLine += "<input type='hidden' name='FM_aktid_old' class='aktid_old_" + feltnrTilfojet + "' value='' />"
                    strNewLine += "<input type='text' name='FM_timer' class='form-control input-small' style='text-align:right;' value='0' />"
                    strNewLine += "<input type='hidden' name='FM_timer' id='FM_timer' class='form-control input-small' value='#' />"
                    strNewLine += "</td>"
                    strNewLine += "<td style='text-align:center'>-</td>"

                    i++;
                }

                strNewLine += "<td style='text-align:center'></td><td style='text-align:center'>-</td><td style='text-align:center'>-</td>"

                strNewLine += "</tr>"

                var lastIndex = 0
                /*   $(".fieldfor_" + thisid).each(function (index) {
                       lastIndex = index
                   });
       
                   $(".fieldfor_" + thisid).each(function (index) {
                       if (parseInt(lastIndex) == parseInt(index))
                       {
                         //  $(this).after(strNewLine);
                       }
                   }); */
                           
                if ($(".fieldfor_" + thisid).is(':hidden')) {
                    $(".fieldfor_" + thisid).show();
                    $("#expandsign_" + thisid).html("-");
                }

                $("#fieldfor_" + thisid).after(strNewLine);
               
            }
        });

        $(document).on('change', '.medarbSEL', function () {
            thisfelt = $(this).data('feltnr')
            newMedarbVal = $(this).val();
            $(".medarb_for_" + thisfelt).val(newMedarbVal);
            preFM_medarb = $("#FM_medarb_new").val();
            $("#FM_medarb_new").val(preFM_medarb + ", " + newMedarbVal);
            
        });
        // Tilføj nyt medarb felt slut

        $(".addField_job").click(function () {
            if (btnReady == 1) {
                btnReady = 0;
                thisid = this.id
                strNewline = "<tr class='tr_medarb tr_medarb_"+ thisid +"'>"
                feltnrTilfojet += 1
                medid = thisid;

                $.post("?medid=" + medid + "&startmd=" + startMD + "&startYY=" + startYY + "&slutMD=" + slutMD + "&slutYY=" + slutYY, { control: "HentJobs", AjaxUpdateField: "true" }, function (data) {
                   // alert(data)
                    $("#jobSEL_" + feltnrTilfojet).html(data);
                    btnReady = 1;
                });



                strNewline += "<td><select class='form-control input-small jobSEL' id='jobSEL_" + feltnrTilfojet + "' data-feltnr='" + feltnrTilfojet + "'><option value='0'>Henter</option></select></td>"


                loopYY = startYY
                loopMD = startMD

                var i = 0
                while (i < numoffdaysorweeksinperiode + 1) {
                    if (i != 0) {
                        if (loopMD == 12) {
                            loopMD = 1
                            loopMD = loopYY + 1
                        }
                        else {
                            loopMD += 1
                        }
                    }

                    strNewline += "<td>"
                    strNewline += "<input type='hidden' name='FM_dato' value='1-" + loopMD + "-" + loopYY + "' />"
                    strNewline += "<input type='hidden' name='FM_medarbid' value='" + medid + "' />"
                    strNewline += "<input type='hidden' name='FM_jobid' class='FM_jobid_" + feltnrTilfojet + "' value='0' />"
                    strNewline += "<input type='hidden' name='FM_aktid' class='aktid_" + feltnrTilfojet + "' value='0' />"
                    strNewline += "<input type='hidden' name='FM_aktid_old' class='aktid_old_" + feltnrTilfojet + "' value='0' />"
                    strNewline += "<input type='text' name='FM_timer' class='form-control input-small' style='text-align:right;' value='0' />"
                    strNewline += "<input type='hidden' name='FM_timer' id='FM_timer' class='form-control input-small' value='#' />"
                    strNewline += "</td>"
                    strNewline += "<td style='text-align:center'>-</td>"
                    i++;
                }

                strNewline += "<td style='text-align:center'>-</td><td style='text-align:center'>-</td><td style='text-align:center'>-</td>"
                strNewline += "</tr>"

                if ($(".tr_medarb_" + thisid).is(':hidden')) {
                    $(".tr_medarb_" + thisid).show();
                    $("#expandsign_" + thisid).html("-");
                }


                $("#tr_medarb_" + thisid).after(strNewline)
            }
        });

        $(document).on('change', '.jobSEL', function () {
            thisfelt = $(this).data('feltnr')
            newJobVal = $(this).val();
            $(".FM_jobid_" + thisfelt).val(newJobVal);

            $.post("?jobid=" + newJobVal, { control: "HentAktiviteter", AjaxUpdateField: "true" }, function (data) {
                aktid = data;
                $(".aktid_" + thisfelt).val(aktid)
                $(".aktid_old_" + thisfelt).val(aktid)
            });
            preFM_job = $("#FM_job_new").val()
            $("#FM_job_new").val(preFM_job + ", " + newJobVal)

        });


        $("#createnewbtn").click(function () {
            $("#createnew").show();
        });

        $("#createnewclose").click(function () {
            $("#createnew").hide();
        });

        var closed = 1
        $(".closeexpandall").click(function () {
            if (closed == 1) {
                $(".columns").show();
                $(".tr_medarb").show();
                $("#expandallsign").html("-");
                $(".expandsign").html("-");
                closed = 0
            } else {
                $(".columns").hide();
                $(".tr_medarb").hide();
                $("#expandallsign").html("+");
                $(".expandsign").html("+");
                closed = 1
            }                       
        });

        
        $(".expandjob").click(function () {
            
            thisid = this.id

            if ($(".fieldfor_" + thisid).is(":not(':hidden')")) {
                $(".fieldfor_" + thisid).hide();
                $("#expandsign_" + thisid).html("+");
            } else {
                $(".fieldfor_" + thisid).show();
                $("#expandsign_" + thisid).html("-");
            }
            
        });

        $(".totFCPer").each(function () {
            thisid = this.id;
            thisval = $(this).val()

            $("#totFC_" + thisid).html(thisval)
        });

        $(".totRealPer").each(function () {
            thisid = this.id;
            thisval = $(this).val()

            $("#totReal_" + thisid).html(thisval)
        });

        
        $(".jobFCtot").each(function () {
            thisid = this.id;
            thisval = $(this).val()

            $("#jobtotFC_" + thisid).html(thisval)
        });

        $(".jobRLtot").each(function () {
            thisid = this.id;
            thisval = $(this).val()

            $("#jobtotRL_" + thisid).html(thisval)
        });

        $(".jobtotSaldo").each(function () {
            thisid = this.id;
            thisval = $(this).val()

            $("#jobtotSaldo_" + thisid).html(thisval)
        });

        


        $("#FM_medarb").change(function () {
            //var selectedValues = $('#FM_medarb').val();
            disVisjobuforecast();

        });

        
        
        $("#FM_medarb2").change(function () {
        //var selectedValues = $('#FM_medarb').val();
            var sortering = $("#FM_sortering").val();
            if (ddcView == 1 && sortering == 1) {
                DisMedSel();
            }
        });
        


   

        $(".btn_timer_kopy").mouseover(function () {
            $(this).css('cursor', 'pointer');
        });

        $(".bt").mouseover(function () {
            $(this).css('cursor', 'pointer');
        });


        $(".btpush").mouseover(function () {
            $(this).css('cursor', 'pointer');
        });

   
        $("#load").hide(1000);

        if ($("#FM_jobsel").val() == 0) {

            $("#tr_job").hide('fast')
            //$.cookie('tr_job', '0');
        }

  

        $(".sp_medarbjoblist").mouseover(function () {
            $(this).css('cursor', 'pointer');
        });
    

        $(".sp_medarbjoblist").click(function () {
      
            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(17, idlngt)

            //alert(idtrim + ".tr_medarb_" + idtrim)

            if ($(".tr_medarb_" + idtrim).css('display') == "none") {

                $(".tr_medarb_" + idtrim).css("display", "");
                $(".tr_medarb_" + idtrim).css("visibility", "visible");
                $(".tr_medarb_" + idtrim).show(1000);

                str = idtrim
                var n = str.indexOf("_");
                if (n == -1) {
            
                    $("#sp_medarbjoblist_" + idtrim).html('<b>[-] Vis alle</b>');
                }

          

            } else {

                $(".tr_medarb_" + idtrim).hide(100);

                str = idtrim
                var n = str.indexOf("_");
            
                if (n == -1) {
                    $("#sp_medarbjoblist_" + idtrim).html('<b>[+]  Vis alle</b>');
                }

            }
      

        });


        if ($("#FM_vis_simpel").val() == 2) {
            $(".tr_medarb").hide(100)
        } else {

            $(".tr_medarb").css("display", "");
            $(".tr_medarb").css("visibility", "visible");
            $(".tr_medarb").show(200);

        }


        
        if (ddcView == 1 || ddcView == "1") {
            $(".tr_medarb").hide();
            $(".medarbjobtotal").hide();
           
            $(".monthtotal_fc").each(function () {
                thisid = this.id

                $("#monthTot_" + thisid).html($(this).val().toString())
            });
            
            $(".monthtotal_real").each(function () {
                thisid = this.id

                $("#monthTot_real_" + thisid).html($(this).val().toString())
            });
            
            $(".pertotal_fc").each(function () {
                thisid = this.id
                totforecast = $("#total_fc_per_" + thisid).val();

                $(this).html("<b>" + totforecast.toString() + "</b>")
            });
            
            $(".pertotal_rl").each(function () {
                thisid = this.id 
                totreal = $("#total_rl_per_" + thisid).val();

                $(this).html("<b>" + totreal.toString() + "</b>")
            });
            
            $(".pertotal_saldo").each(function () {
                thisid = this.id
                totsaldo = $("#total_saldo_per_" + thisid).val();

                $(this).html("<b>" + totsaldo.toString() + "</b>")
            });


        }
       
        $(".tr_job_medarb").click(function () {
            thisid = this.id

            if ($(".tr_medarb_" + thisid).is(":not(':hidden')")) {
                $(".tr_medarb_" + thisid).hide();
                $(".medarbjobtotal_" + thisid).hide();
                $("#expandsign_" + thisid).html("+");
            }
            else {
                $(".tr_medarb_" + thisid).show();
                if (ddcView != 1 || ddcView != "1") {
                   $(".medarbjobtotal_" + thisid).show();
                }
                $("#expandsign_" + thisid).html("-");
            }


        });





        //$("#FM_vis_simpel").click(function () {
        //
        //    if ($("#FM_vis_simpel").is(':checked') == true) {
        //        $.cookie("vissimpel", '1');
        //    } else {
        //        $.cookie("vissimpel", '0');
        //    }

        //    visSubtotJobmed();
        //});


   



        function showhideinternnote(usemrn) {
            if ($("#newline_" + usemrn).css('display') == "none") {

                $("#newline_" + usemrn).css("display", "");
                $("#newline_" + usemrn).css("visibility", "visible");
                $("#newline_" + usemrn).show(2000);
                $.scrollTo($("#newline_" + usemrn), 1500, { offset: { top: -100, left: -30} });

            } else {

                $("#newline_" + usemrn).hide(1000);
                $.scrollTo($("#newline_" + usemrn), 1500, { offset: { top: -150, left: -30} });
            }
        }


        function showhidenylinie(jobid_medid) {
            //alert(jobid_medid)

            if ($("#nylinjepaajob_" + jobid_medid).css('display') == "none") {

                $("#nylinjepaajob_" + jobid_medid).css("display", "");
                $("#nylinjepaajob_" + jobid_medid).css("visibility", "visible");
                $("#nylinjepaajob_" + jobid_medid).show(2000);
                $.scrollTo($("#nylinjepaajob_" + jobid_medid), 1500, { offset: { top: -100, left: -30} });

            } else {

                $("#nylinjepaajob_" + jobid_medid).hide(1000);
                $.scrollTo($("#nylinjepaajob_" + jobid_medid), 1500, { offset: { top: -150, left: -30} });
            }
        }


        $(".rodstor").click(function () {

            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(4, idlngt)

            showhideinternnote(idtrim);
        });

        $(".rodlille").click(function () {

            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(6, idlngt)


            //var medid = $("#nl_medid_" + idtrim).val();

            //alert(idtrim + "_" + $("#nl_medid_" + idtrim).val())

            // + "_" + medid

            showhidenylinie(idtrim);
        });

        $(".red").click(function () {

            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(4, idlngt)

            showhideinternnote(idtrim);
        });


        $(".red_j").click(function () {

            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(4, idlngt)

            //var medid = $("#nl_medid_" + idtrim).val();

            //alert(idtrim)
            showhidenylinie(idtrim);
        });



                 
        $("#FM_jobsel").change(function () {
            var sortering = $("#FM_sortering").val();

            if (ddcView == 1 && sortering == 0) {
                DisJobSel();
            }
        });

        $("#FM_sortering").change(function () {
            thisval = $(this).val();
            if (parseInt(thisval) == 0) {
                DisJobSel();
            }
            else {
                DisMedSel();
            }
            
        });
        

        $("#sogtxt").focus(function () {

            //alert("her")
            $("#FM_jobsel").val('0')




        });


        $(".nlFM_jobid").change(function () {

            var thisid = this.id
            var thislnght = thisid.length
            thismedid = thisid.slice(11, thislnght)

           // alert(thismedid)
            var thisval = $("#" + thisid + " option:selected").val()
            

            $(".hdFM_jobid_" + thismedid + "").val(thisval);


        });

        $(".aaFM_jobid").change(function () {

            var thisid = this.id
            var thislnght = thisid.length
            thismedid = thisid.slice(11, thislnght)


            var thisval = $("#" + thisid + " option:selected").val()
            //alert(thisval)

            $(".ahFM_jobid_" + thismedid + "").val(thisval);


        });


        $(".aaNLFM_jobid").change(function () {

            var thisid = this.id
            var thislnght = thisid.length
            thismedid = thisid.slice(13, thislnght)


            var thisval = $("#" + thisid + " option:selected").val()
            //alert(thisval)

            $(".ahNLFM_jobid_" + thismedid + "").val(thisval);


        });


        $(".bt").click(function () {

            var thisid = this.id
            //alert(thisid)
            var thislnght = thisid.length
            var thisxid = thisid.slice(3, thislnght)

            var thisyid = thisid.slice(thislnght - 2, thislnght)
            var rtrm = 1

            var thisyid_first = thisyid.substring(0, 1)
            if (thisyid_first == "_") {
                //alert("her")
                thisyid = thisyid.substring(1, 2)
                rtrm = 1
            } else {
                rtrm = 2
            }

            var thisid_uden_y = thisid.slice(3, thislnght - 1 - rtrm)

            //alert(thisyid + "_" + thisid_uden_y + "")

            var thisval = $("#FM_timer_" + thisxid + "").val()
            //alert(thisval)

            //thisval = 500
            if (thisval != "") {
                for (i = thisyid; i <= 15; i++) {
                    $("#FM_timer_" + thisid_uden_y + "_" + i).val(thisval) 
                    //alert("#FM_timer_" + thisid_uden_y + "_" + i)
                }
            }


        });



        $(".btpush").click(function () {

            var thisid = this.id
     
            var thislnght = thisid.length
            var thisxid = thisid.slice(7, thislnght)

            //alert(thisxid)

            var thisyid = thisid.slice(thislnght - 2, thislnght)
            var rtrm = 1

            var thisyid_first = thisyid.substring(0, 1)
            if (thisyid_first == "_") {
                //alert("her")
                thisyid = thisyid.substring(1, 2)
                rtrm = 1
            } else {
                rtrm = 2
            }

            var thisid_uden_y = thisid.slice(7, thislnght - 1 - rtrm)

            //alert(thisyid + "_" + thisid_uden_y + "")

       
            //alert(thisval)
            var cnt = 0;
            var thisval = 0;
            var nextVal = 0;
            var nextnextVal = 0;
            //thisval = 500
            //if (thisval != "") {
            for (i = 0; i < 12; i++) { //må ikke opdatere sidste felt

                cnt = (i + 1) / 1

                //alert("her")
           
                var nextnextval = $("#FM_timer_" + thisid_uden_y + "_" + cnt).val()

                if (i == 0) {
             
                    thisval = $("#FM_timer_" + thisid_uden_y + "_" + i).val()
                    $("#FM_timer_" + thisid_uden_y + "_" + i).val(0)
               
                } else {
                    if (nextval != '') {
                        thisval = nextval
                    } else {
                        thisval = 0
                    }
                }

                //alert(thisid_uden_y + "_" + i + " nextval " + nextval + " nextnextval" + nextnextval)
                $("#FM_timer_" + thisid_uden_y + "_" + cnt).val(thisval)

                nextval = nextnextval
            }
        


        });



        $(".tilfoj_ny_job").change(function () {



            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(3, idlngt)

            var medarbid = thisid.slice(12, idlngt);
            var sog_val = $("#" + thisid + "")

            //alert(idtrim + " " + sog_val.val() + " medarbid:" + medarbid)

            $.post("?jq_sogkunde=1&medarbid=" + medarbid, { control: "FN_getAkt", AjaxUpdateField: "true", cust: sog_val.val() }, function (data) {
                                                      
   

                //alert("her")
                //$("#tdtest" + idtrim + "").html(data);
                $("#td" + idtrim + "").html(data);
                $("#td" + idtrim + "").focus();

            });



        });



   


    });


        
         



    function setjobidval(usemrn) {
        ymax = document.getElementById("ymax").value

        for (i = 0; i <= ymax; i++) {
            //salert(document.getElementById("selFM_jobid").value)
            document.getElementById("FM_jobid_" + usemrn + "_" + i).value = document.getElementById("selFM_jobid_" + usemrn).value
        }
    }


    function copyTimer(usemrn, jobid, yval) {
        ymax = document.getElementById("ymax").value
        var valuethis = 0;
        var yvalkri = 0;
        yvalkri = yval / 1;


        //alert(usemrn + " " + jobid + " " + yval)

        //if (document.getElementById("FM_timer_" + usemrn + "_" + jobid + "_" + yval).lenght =! 0) {
        valuethis = document.getElementById("FM_timer_" + usemrn + "_" + jobid + "_" + yval).value
        //}else{
        //valuethis = 0
        //}

        //alert(valuethis)

        //if (valuethis =! 0) {
        //valuethis = valuethis/1

        for (i = yvalkri; i <= ymax; i++) {
            //salert(document.getElementById("selFM_jobid").value)
            document.getElementById("FM_timer_" + usemrn + "_" + jobid + "_" + i).value = valuethis
        }

        //}
    }

    function clearJnavn() {
        document.getElementById("jobselect").value = ""
    }


    function seladd() {
        document.getElementById("selectplus").value = 1
        $("#indlas").submit()
    }

    function selminus() {
        document.getElementById("selectminus").value = 1
        $("#indlas").submit()
    }

    function showtildeltimer(thisMid, dato, jobid, persel) {
        document.getElementById("tildeltimer").style.visibility = "visible"
        document.getElementById("tildeltimer").style.display = ""
        document.getElementById("FM_tildeltimer_mid").value = thisMid
        //document.getElementById("FM_dato_sel").value = dato

        document.getElementById("chkboxes").innerHTML = dato


        //document.getElementById("FM_dato").style.visibility = "hidden"
        //document.getElementById("FM_dato").style.display = "none"

        //if (persel =! 1) {
        //document.getElementById("ugetxt").style.visibility = "visible"
        //document.getElementById("ugetxt").style.display = ""
        //}


        for (i = 0; i < document.forms["timertildel"]["FM_tildeltimer_job"].length; i++) {
            if (document.forms["timertildel"]["FM_tildeltimer_job"][i].value == jobid) {
                document.forms["timertildel"]["FM_tildeltimer_job"][i].selected = true;
            }
        }

    }


    function checkAll(field) {
        field.checked = true;
        for (i = 0; i < field.length; i++)
            field[i].checked = true;
    }

    function unCheckAll(field) {
        field.checked = true;
        for (i = 0; i < field.length; i++)
            field[i].checked = false;
    }

    function hidetildeltimer() {
        document.getElementById("tildeltimer").style.visibility = "hidden"
        document.getElementById("tildeltimer").style.display = "none"
    }


    function BreakItUp() {
        //Set the limit for field size.
        var FormLimit = 102399
        //102399

        //Get the value of the large input object.
        var TempVar = new String
        TempVar = document.theForm.BigTextArea.value

        //If the length of the object is greater than the limit, break it
        //into multiple objects.
        if (TempVar.length > FormLimit) {
            document.theForm.BigTextArea.value = TempVar.substr(0, FormLimit)
            TempVar = TempVar.substr(FormLimit)

            while (TempVar.length > 0) {
                var objTEXTAREA = document.createElement("TEXTAREA")
                objTEXTAREA.name = "BigTextArea"
                objTEXTAREA.value = TempVar.substr(0, FormLimit)
                document.theForm.appendChild(objTEXTAREA)

                TempVar = TempVar.substr(FormLimit)
            }
        }
    }


       
      