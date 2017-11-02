




$(document).ready(function () {

  

   




    $(".sletimperr_all").click(function () {

        //alert("HER")

        thisDay = $(this).val()
        isThisChk = $(this).is(':checked')

        //$(".sletimperr_" + thisDay).each(function () {

        //alert(isThisChk)

            if (isThisChk == true) {
                $(".sletimperr_" + thisDay).attr("checked", "checked")
                } else {
                $(".sletimperr_" + thisDay).attr("checked", "")
                }
          
        //});

       

    });


});



