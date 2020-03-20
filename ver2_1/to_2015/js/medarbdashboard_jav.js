






$(window).load(function () {

    //visSubtotJobmed();

});


$(document).ready(function () {

    //alert("HEj")

   
    //$(".loader").hide("fast")
    //$(".loader").css("visibility", "hidden")
    //$(".loader").css("display", "none")

    $("#jq_medid").change(function () {

        //alert($("#jq_medid").val())
        if ($("#jq_medid").val() == 0) {
            $("#jq_medid > option").each(function () {
                $(this).prop('selected', true);
            });
        }

    });


    $("#jq_progrp").change(function () {
        $("#jq_medid > option").each(function () {
            $(this).prop('selected', false);
        });

        $("#dashboard_filter").submit();
    });


   

        //$(".panel-heading").mouseover(function () {
        //    alert("HER")
        //    $(this).css('cursor', 'pointer');
        //});

    
   

});
      
 




       
      