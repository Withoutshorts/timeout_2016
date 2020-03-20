
$(document).ready(function () {
   
    $("#showNews").click(function () {
        $("#listofnews").css('display', 'block')

       /* var systemnews = $("#systemnewsids").val();
        $("#systemnewsids").val('');
        $("#antalnyemeddelelser").html('');

        if (systemnews != "") {

            var systemnewssplit = systemnews.split(",");
            for (var i = 0; i < systemnewssplit.length; i++) {
                document.cookie = "systemmeddelelse_" + systemnewssplit[i] + "_laest=1";
            }



        } */

    });

    $("#newsread").click(function () {

        var systemnews = $("#systemnewsids").val();
       
        $("#listofnews").hide();

        if (systemnews != "") {

            var my_window = window.open("../to_2015/systemmeddelelserlaest.asp?sids=" + systemnews, "my_window", "height=1,width=1");
            
            setTimeout(function () { my_window.close(); }, 250);

            $("#systemnewsids").val('');
            $("#antalnyemeddelelser").html('');

        }

       // $("#settoread").submit();

        


    });
    

});