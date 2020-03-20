$(document).ready(function () {

  // alert("HERHER")

    $("#sog_kunde").keyup(function () {

        thisval = $("#sog_kunde").val();

        $.post("?sogtxt=" + thisval, { control: "FN_findkunder", AjaxUpdateField: "true" }, function (data) {

          //  alert("HER")

            $("#kunde_DD").show();
           
            $("#kunde_DD").html(data);

        });
    });

    $("#kunde_DD").change(function () {
        
        kundevalgt = $("#kunde_DD").val();
     //  alert(kundevalgt);

        $("#valgt_kunde").val(kundevalgt);
        $("#kunde_DD").hide();
        
       // $("#sog_kunde").val('');

        $.post("?kundeSEL=" + kundevalgt, { control: "FN_getkundenavn", AjaxUpdateField: "true" }, function (data) {

            $("#sog_kunde").val(data);

        });

    });

    $("#kunde_ADD").click(function () {
        
        valgt_kunde = $("#valgt_kunde").val();
        startaar = $("#startaar").val();
        slutaar = $("#slutaar").val();

       // alert(startaar + " - " + slutaar)

        findeskunde = $("#kundefindes_" + valgt_kunde).val();
        //alert(findeskunde)

        if (findeskunde != 1) {
            $.post("?kundeSEL=" + valgt_kunde + "&startaar=" + startaar + "&slutaar=" + slutaar, { control: "FN_getkundedata", AjaxUpdateField: "true" }, function (data) {

                $("#list_body").append(data);
                $("#sog_kunde").val('');
            });
        } else {
            $("#sog_kunde").val('');
            $("#errormessage").show();
            $("#errormessage").fadeOut(3000);
        }

    });


    var budgetview = 0
    $("#editbudget").click(function () {
        if (budgetview == 0) {
            $(".txtfield").hide();
            $(".inputfield").show();
            budgetview = 1
        }
        else {
            $(".inputfield").hide();
            $(".txtfield").show();           
            budgetview = 0
        }

    });
    

});