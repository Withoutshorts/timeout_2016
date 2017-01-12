





$(document).ready(function() {


   

    $("#loadbar").hide(1000);



    $("#s_luk_igv").click(function () {

        $("#jobstatusigv").hide("fast")
        $("#eksport").hide("fast")
        
        
    });



    $('#s_luk_igv').mouseover(function () {

        $(this).css("cursor", "pointer");

    });


});

