﻿$(document).ready(function () {


   $("#RFID_field").change(function () {

        var RFID_id = $("#RFID_field").val();
        //alert("hep")
        //alert(RFID_id)

        $("#scan").submit();
        //Alert("HER")
    });


   $('html, body').on('touchmove', function (e) {
        //prevent native touch activity like scrolling
        e.preventDefault();
    });

    $("#text_besked").delay(3000).fadeOut();

    $(".container").bind('click', function () {

        //alert("FOCUS")

        $("#RFID_field").focus();

    });



});


