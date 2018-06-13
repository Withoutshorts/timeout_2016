



$(document).ready(function() {



    $("#sogjob_lukpas").click(function () {

        jobliste();

    });

    $("#sogjob").keyup(function () {

    
        jobliste();

    });
    

    function jobliste() {


        var sogVal = $("#sogjob").val();
        var sogjob_lukpas = 0;

        //alert($("#sogjob_lukpas").is(":checked"));

        //alert($("#sogjob_lukpas").prop('checked'));

        if ($("#sogjob_lukpas").is(":checked")) {
            sogjob_lukpas = 1;
        }



        //alert("her" + thisval + " m:" + usemrn + " sogjob_lukpas: " + sogjob_lukpas)


        $.post("?sogval=" + sogVal + "&sogjob_lukpas=" + sogjob_lukpas, { control: "FN_sogjobliste", AjaxUpdateField: "true" }, function (data) {

            $("#jobprint_joblisten").html(data);

        });



    };


    

  
});



