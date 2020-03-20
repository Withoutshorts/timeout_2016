







$(document).ready(function () {


    $(".picmodal").click(function () {

        

        var modalid = this.id
        var idlngt = modalid.length
        var idtrim = modalid.slice(6, idlngt)

        //var modalidtxt = $("#myModal_" + idtrim);
        var modal = document.getElementById('myModal_' + idtrim);

        //alert("HER" + idtrim)

        modal.style.display = "block";

        window.onclick = function (event) {
            if (event.target == modal) {
                modal.style.display = "none";
            }
        }

    });



    $("#FM_visallemedarb").click(function () {


        var sesid = $("#FM_sesMid").val()
        $("#usemrn").val(sesid)

        $("#filterkri").submit();


    });

    $("#usemrn").change(function () {

        $("#filterkri").submit();
    });



    $("#udspec").click(function () {

        show_udspecdiv();

    });



    function show_udspecdiv() {


        if ($("#udspecdiv").css('display') == "none") {

            $("#udspecdiv").css("display", "");
            $("#udspecdiv").css("visibility", "visible");
            $("#udspecdiv").show(4000);

            $.cookie('c_udspecdiv', '1');

        } else {
            $("#udspecdiv").hide(1000);
            $.cookie('c_udspecdiv', '0');
        }

    }


    if ($.cookie('c_udspecdiv') == "1") {


        $("#udspecdiv").css("display", "");
        $("#udspecdiv").css("visibility", "visible");
        $("#udspecdiv").show(4000);
    
    }


});

$(window).load(function() {
    // run code
    $("#loadbar").hide(1000);
});


