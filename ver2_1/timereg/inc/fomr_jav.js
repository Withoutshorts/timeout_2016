





$(document).ready(function() {

    $("#sp_grundlag").click(function () {

        if ($("#div_grundlag").css('display') == "none") {
            $("#div_grundlag").css("display", "");
            $("#div_grundlag").css("visibility", "visible");
            $("#div_grundlag").show("fast");
          

        } else {

            $("#div_grundlag").hide("fast");
           

        }

    });

    
    $("#sp_grundlag").mouseover(function () {
        $(this).css('cursor', 'pointer');
    });


});

$(window).load(function() {
    // run code
    $("#load").hide(1000);
});


function clearJobsog() {
    document.getElementById("FM_jobsog").value = ""
}