


$(window).load(function () {
    // run code


    alert("hejhej")


});




$(document).ready(function () {
    alert("hej morten")
    $("#btn_lto_rap").click(function () {
        alert("MOrten")
    });

    $("#FM_rapporttype").click(function () {


        //alert("KUK?" + $("#FM_rapporttype").val())

        var raptype = $("#FM_rapporttype").val()
       
        //alert(raptype)

        if (raptype == "2") {

            $("#FM_abo_mtyp").attr('disabled', true);
            $("#FM_abo_progrp").attr('disabled', true);

        } else {

            $("#FM_abo_mtyp").removeAttr('disabled');
            $("#FM_abo_progrp").removeAttr('disabled');

        }


    });

   
   

});

