


$(document).ready(function () {




    ////////// SMILEY pop Up //////////////////////
    function smileyshowhide() {

        //$("#s2").hide(1000);
        //$("#s2_td").css("background-color", "#EFF3FF");


        if ($("#s0").css('display') == "none") {

            $("#s0").css("display", "");
            $("#s0").css("visibility", "visible");
            $("#s0").show(4000);
            $("#s0_td").css("background-color", "#FFFF99");

            $.scrollTo("#s0", 300, { offset: -300 }); //

            //scrollTo($("#s0"), 800);

            //$.scrollTo('260px', 300, { offset: 100 });

        } else {

            $("#s0").hide(1000);
            $("#s0_td").css("background-color", "#EFf3FF");

            //$.scrollTo('100px', 300);

        }

    }

    $(".sA1_k").click(function () {
        //smileyshowhide()

        //alert("her")

        $("#s1").hide(1000);
        $("#s0").hide(1000);
    });


    $("#s1_k").click(function () {

       
        smileyshowhide()
    });


    




    $('#se_uegeafls_a').mouseover(function () {

        $(this).css("cursor", "pointer");

    });

    $('#se_uegeafls_a').click(function () {

        if ($("#dv_ugeafslutninger").css('display') == "none") {

            //$.cookie('showfilter', '0');
            $("#dv_ugeafslutninger").css("display", "");
            $("#dv_ugeafslutninger").css("visibility", "visible");
            $("#dv_ugeafslutninger").show(100);

        } else {

            $("#dv_ugeafslutninger").hide(100);
        }

    });

    ///////////////////////////////////////////////////////



});




function showlukalleuger() {
    if (document.getElementById("FM_afslutuge").checked == true) {
        document.getElementById("lukalleuger").style.visibility = "visible"
        document.getElementById("lukalleuger").style.display = ""
    } else {
        document.getElementById("lukalleuger").style.visibility = "hidden"
        document.getElementById("lukalleuger").style.display = "none"
        document.getElementById("FM_alleuger").checked = false
    }
}