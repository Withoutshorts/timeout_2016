




$(window).load(function () {
    // run code





});











$(document).ready(function () {

 
    var lto = $("#lto").val();
    var ikkeloggetud = $("#ikkeloggetud").val();
    if ((lto == "miele" || lto == "esn") && ikkeloggetud == 0) {
        //alert("Skal checke ind")
    }

    if ((lto == "miele" || lto == "esn") && ikkeloggetud == 1) {
        //alert("logger ind automatisk")

        $("#container").submit();
    }


    $("#m_login").focus(function() {
    
        if ($("#m_login").val() == 'Brugernavn') {
            $("#m_login").val('')
        }
    });
   

    $("#m_pw").focus(function () {

        if ($("#m_pw").val() == 'Password') {
            $("#m_pw").val('')
        }
    });


  

});


