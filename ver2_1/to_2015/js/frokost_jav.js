$(document).ready(function () {

  

    //alert("Check");

    var pauseBegyndt = 0  //Kun til test skal slettet igen

    $("#medarbId").focus();


    $("#medarbId").change(function () {

       // alert("hep")

        if (pauseBegyndt == 0)
        {
            $('#pauseCheck').css('color', 'green');
            $("#pauseCheck").delay(1000).fadeIn();
            $('#pauseCheck').css('visibility', 'inherit');
            $("#pauseCheck").delay(1000).fadeOut();
            pauseBegyndt = 1

            $("#medarbId").val('');
        }
        else
        {
            $('#pauseCheck').css('color', 'red');
            $("#pauseCheck").delay(1000).fadeIn();
           $('#pauseCheck').css('visibility', 'inherit');
            $("#pauseCheck").delay(1000).fadeOut();
            pauseBegyndt = 0

            $("#medarbId").val('');
        }
       

    });

    




});


