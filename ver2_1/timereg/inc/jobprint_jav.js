

function print_pdf() {

    document.getElementById("FM_vis_indledning_pdf").value = document.getElementById("FM_vis_indledning").value
    document.getElementById("FM_vis_afslutning_pdf").value = document.getElementById("FM_vis_afslutning").value
    document.getElementById("pdf_skabelonnavn").value = document.getElementById("skabelonnavn").value
    document.getElementById("pdf_doktype").value = document.getElementById("doktype").value
    document.getElementById("pdf_foid").value = document.getElementById("foid").value

    $("#mk_pdf").submit()

}


function print_pnt() {

    document.getElementById("print_skabelonnavn").value = document.getElementById("skabelonnavn").value
    document.getElementById("print_foid").value = document.getElementById("foid").value
    document.getElementById("print_doktype").value = document.getElementById("doktype").value

    $("#mk_pnt").submit()
}


$(window).load(function () {
    // run code


    


  
});











$(document).ready(function () {


    $("#gempdf").click(function () {
        
        var vedrinfoVal = $("#vedrinfo").val()

        //var idlngt = vedrinfoVal.length
        var idtrim = vedrinfoVal.slice(10, 12)

        //alert(idtrim)

        if (idtrim == "l<") {

        $("#dv_jobst").css("visibility", "visible")
        $("#dv_jobst").css("display", "")

        $("#jobs_pdf_pnt").val('1')

        } else {

            print_pdf();

        }

    });

    $("#gempnt").click(function () {

        var vedrinfoVal = $("#vedrinfo").val()

        //var idlngt = vedrinfoVal.length
        var idtrim = vedrinfoVal.slice(10, 12)

        //alert(idtrim)

        if (idtrim == "l<") {

            $("#dv_jobst").css("visibility", "visible")
            $("#dv_jobst").css("display", "")

            $("#jobs_pdf_pnt").val('2')

        } else {

            print_pnt();

        }

    });


    //alert("ok")

   
    $("#jobstatus").change(function () {

       
        if ($("#jobstatus option:first").is(':selected') == true) {
        skiftjobstatus(1)
        }

    });

    $("#jobstatus_bt").click(function () {

        delayval = 1

        if ($("#statushilsenTxtSet_0").is(':checked') == true) {
            
        } else {
        
            

        if ($("#statushilsenTxtSet_1").is(':checked') == true) {
            $("#td_hilsSt_1").html('[X]')
        }

        if ($("#statushilsenTxtSet_2").is(':checked') == true) {
            $("#td_hilsSt_2").html("[X]")
        }

        if ($("#statushilsenTxtSet_3").is(':checked') == true) {
            $("#td_hilsSt_3").html("[X]")
        }

        if ($("#statushilsenTxtSet_4").is(':checked') == true) {
            $("#td_hilsSt_4").html("[X]")
            //alert($("#statushilsenTxtTxt").val())
            statushilsenTxtTxt = $("#statushilsenTxtTxt").val()
            $("#td_hilsSt_4_txt").html("Andet:<br>" + statushilsenTxtTxt)
        }

        txt = $("#dv_flgseddel_footer").html()
        $("#FM_vis_afslutning").val(txt) 

        }
       

        skiftjobstatus(delayval)

    });


    function skiftjobstatus(delay) {


        //alert("her")
   
        var jobid = $("#jobs_jobid").val()
        var jobstatus = $("#jobstatus").val()
      
        //alert("jobid: "+ jobid + "# jobstatus" + jobstatus)
        
        
        $.post("?jobid=" + jobid + "&jobstatus=" + jobstatus, { control: "FN_jobstatus", AjaxUpdateField: "true" }, function (data) {


            //$("#lbs_jobst").css("visibility");
            //$("#lbs_jobst").html(">");

            //$("#lbs_jobst").css("visibility", "visible")

            //alert("her")
          
            if (delay == 0) {

                setTimeout(function () {
                    // Do something after 5 seconds

                    $("#lbs_jobst").css("visibility", "hidden")

                    pdf_pnt = $("#jobs_pdf_pnt").val()
                    if (pdf_pnt == 1) {
                        print_pdf();
                    } else {
                        print_pnt();
                    }

                }, 500);

            } else {

                $("#lbs_jobst").css("visibility", "hidden")

                pdf_pnt = $("#jobs_pdf_pnt").val()
                if (pdf_pnt == 1) {
                    print_pdf();
                } else {
                    print_pnt();
                }

            }


            $("#dv_jobst").hide(100)
            $("#FM_vis_afslutning ").val('')
            
        });
    }

    
        





    });

   




