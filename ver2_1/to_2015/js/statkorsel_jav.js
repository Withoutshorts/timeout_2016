$(document).ready(function () {

    $("#afregnAlle").click(function () {

        if ($("#afregnAlle").prop("checked") == true) {
            ($(".afregnkorsel").prop("checked", true))
        } else {
            ($(".afregnkorsel").prop("checked", false))
        }
    });
   

});


