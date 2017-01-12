


$(window).load(function () {
    // run code





});











$(document).ready(function () {

    

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

