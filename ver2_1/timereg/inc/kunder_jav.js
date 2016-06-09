$(document).ready(function() {

        
    $("#visflkonti").mouseover(function () {
        
        $(this).css('cursor', 'hand');
    });


    $("#visflkonti").click(function () {

        //alert("hht")
        if ($(".tr_konti").css('display') == "none" ) {
            $(".tr_konti").css("display", "");
            $(".tr_konti").css("visibility", "visible");
            $(".tr_konti").show("fast");
            //$.scrollTo('1400px', 400);
            //$.scrollTo($('#div_indforstamgrp'), 4000);

        } else {

            $(".tr_konti").hide("fast");
            //$.scrollTo('1000px', 400);
            //$.scrollTo('-=100px', 1500);

        }
        
    });

   
    $(".tr_konti").hide("fast");

});