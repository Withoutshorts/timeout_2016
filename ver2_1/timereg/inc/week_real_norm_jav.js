





$(document).ready(function() {



    globalWdt = $("#globalWdt").val()
    $("#medarbafstem_aar").css("width", globalWdt);
  

   
    $("#sp_filterava").mouseover(function () {

        $(this).css("cursor", "pointer");

    });

    $("#sp_filterava").click(function () {


        if ($("#div_filterava").css('display') == "none") {

            $("#div_filterava").css("display", "");
            $("#div_filterava").css("visibility", "visible");
            $("#div_filterava").show(1000);

            $("#div_filterava_show").val('1')
            $.cookie('filterava_show', '1');

        } else {

            $("#div_filterava").hide(1000);

            $("#div_filterava_show").val('0')
            $.cookie('filterava_show', '0');
        }
    });


    if ($.cookie('filterava_show') == '1') {

        $("#div_filterava").css("display", "");
        $("#div_filterava").css("visibility", "visible");
        $("#div_filterava").show(100);

    }


    $("#FM_eksporter_direkte").click(function () {

        
        if ($("#FM_eksporter_direkte").is(':checked') == true) {
            $("#periode").attr('target', '_blank');
        } else {
            $("#periode").attr('target', '');
            }

    });
   


});

$(window).load(function() {
    // run code
    $("#loadbar").hide(1000);
});