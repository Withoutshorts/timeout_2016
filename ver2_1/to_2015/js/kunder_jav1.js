$(document).ready(function () {


  //  alert("HER")

    var budget_view = 0
    $("#editbudget").click(function () {

        if (budget_view == 0) {
            $(".budget_values").hide();
            $(".budget_input").show();
            budget_view = 1
        }
        else {
            $(".budget_input").hide();
            $(".budget_values").show();
            budget_view = 0
        }
        


        
    });




    $("#visflkonti").mouseover(function () {

        $(this).css('cursor', 'hand');
    });


    $("#visflkonti").click(function () {

        //alert("hht")
        if ($(".tr_konti").css('display') == "none") {
            $(".tr_konti").css("display", "");
            $(".tr_konti").css("visibility", "visible");
            $(".tr_konti").show("fast");
            //$.scrollTo('200px', 400);
            //$.scrollTo("#tr_konto_b", 4000);
            //$.scrollTop(300);

        } else {

            $(".tr_konti").hide("fast");
            //$.scrollTo('1000px', 400);
            //$.scrollTo('-=100px', 1500);
            //$.scrollTo('200px', 400);

        }

    });



    $(".tr_konti").hide("fast");






});