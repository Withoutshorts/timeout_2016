// JScript File

$(window).load(function() {
    // run code


//if jq_func
//getAktlisten()

   

});

$(document).ready(function () {
    //alert("her")


    
   


    // Medarbsøg //
    function getMedlisten() {


        var visalle = 0 //bruges ikke

        if ($("#jq_vispasogluk").is(':checked') == true) {
            var visalle = 1
        } else {
            var visalle = 0
        }

        
        $.cookie('visalleCHK', visalle);

        var sog_val = $("#m_sog")
        $.cookie('msogval', sog_val.val());


        lastMid = $("#jq_lastmid").val()

        //sogKri = $("#m_sog").val()


        $.post("?jq_visalle=" + visalle + "&lastMid=" + lastMid, { control: "FN_getMedlisten", AjaxUpdateField: "true", cust: sog_val.val() }, function (data) {
            //$("#FM_modtageradr").val(data);
            $("#div_medarb_list").html(data);
            //$("#jobid").html(data);
            //alert(data)

            $("#antalm").val($("#jq_antal").val())
            $("#antalmialt").val($("#jq_antal_tot").val())
            $("#ekspTxt").val($("#jq_ekspTxt").val())

        });

    }

     $("#m_sog").focus(function () {
         if ($("#m_sog").val() == "Søg på medarbejder her..") {
        $("#m_sog").val('')
    }
    });

    $("#m_sog").keyup(function () {
        getMedlisten();
    });

    $("#jq_vispasogluk").click(function () {
        getMedlisten();
    });

    $("#Submit1").click(function () {
        getMedlisten();
    });


    if ($.cookie('msogval') != "null") {
        $("#m_sog").val($.cookie('msogval'))
        if ($("#m_sog").val() == "null") {
            $("#m_sog").val('Søg på medarbejder her..')
        }
    }

    if ($.cookie('visalleCHK') == "1") {
       $("#jq_vispasogluk").attr('checked', true);
   }

   getMedlisten();

});

   



   