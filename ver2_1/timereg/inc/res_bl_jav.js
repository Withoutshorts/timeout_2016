


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

var count = $("#FM_medarb :selected").length;
if (count > 10) {
    $("#FM_vis_job_u_fc").attr('checked', false);
    $("#FM_vis_job_u_fc").attr("disabled", "disabled");
} else {
    $("#FM_vis_job_u_fc").attr("disabled", "");
}

}



$(window).load(function () {

    //visSubtotJobmed();

});


    $(document).ready(function () {

        disVisjobuforecast();
    
        $("#FM_medarb").change(function () {
            //var selectedValues = $('#FM_medarb').val();
            disVisjobuforecast();

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
            $.cookie('tr_job', '0');
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

            //alert("her")
            $("#sogtxt").val('')

            $("#mdselect").submit()


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


       
      