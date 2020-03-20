
$(document).ready(function () {

   $("#systemeddelelser").hover(function () {
        
        var systemnews = $("#systemnewsids").val();
        $("#systemnewsids").val('');
        $("#antalnyemeddelelser").html('');
        
       if (systemnews != "") {

           var systemnewssplit = systemnews.split(",");
           for (var i = 0; i < systemnewssplit.length; i++) {
               document.cookie = "systemmeddelelse_" + systemnewssplit[i] + "_laest=1";
           }
        }

        
    });
    

});