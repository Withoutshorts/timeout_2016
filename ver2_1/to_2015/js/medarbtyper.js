$(document).ready(function () {

    $("#FM_normtimer_rul").change(function () {
        thissogval = $(this).val();

           if (thissogval == "0") {
               $(".div_normtimer_rul").hide();
          } else {
               $(".div_normtimer_rul").show();
          } 

     

    });

    thissogval = $("#FM_normtimer_rul").val();

    if (thissogval == "0") {
        $(".div_normtimer_rul").hide();
    } else {
        $(".div_normtimer_rul").show();
    } 
   
});