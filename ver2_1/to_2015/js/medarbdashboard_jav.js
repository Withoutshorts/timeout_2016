






$(window).load(function () {

    //visSubtotJobmed();

});


$(document).ready(function () {

    //alert("HEj")

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
      
});
      
 




       
      