$(window).load(function () {
    // run code
    
});









    $(document).ready(function () {



        $("#loadbar").hide(1000);



        $(".glag").click(function () {
           //alert("her")
           var thisid = this.id
           var thisvallngt = thisid.length
           var thisvaltrim = thisid.slice(5, thisvallngt)
           //alert(thisvaltrim)

           showhideglag(thisvaltrim);
        });



        function showhideglag(thisid) {
            if ($("#divglag_" + thisid).css('display') == "none") {

                $("#divglag_" + thisid).css("display", "");
                $("#divglag_" + thisid).css("visibility", "visible");
                $("#divglag_" + thisid).show(4000);

            } else {

                $("#divglag_" + thisid).hide(1000);

            }
        }




});