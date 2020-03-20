$(document).ready(function () {


    $("#FM_newpassword_repeat").keyup(function () {

        strnewpassword = $("#FM_newpassword").val();
        thisval = $(this).val();

        if (thisval != strnewpassword)
        {
            $('#statusbox').css('color', "darkred");
            $("#statusbox").removeClass("fa fa-check")
            $("#statusbox").text("Passwords dont match")
        } else
        {
            $('#statusbox').css('color', "green");
            $("#statusbox").addClass("fa fa-check")
            $("#statusbox").text("")
        }


    });

    

});


