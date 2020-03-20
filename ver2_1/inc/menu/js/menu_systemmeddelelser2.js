
$(document).ready(function () {
   
    $("#showNews").click(function () {
        $("#listofnews").css('display', 'block')

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

    $("#newsread").click(function () {
        
        $("#listofnews").hide();

    });
    

});