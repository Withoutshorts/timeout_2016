



$(document).ready(function() {



    $("#sogjob").keyup(function () {

    
        var sogVal = $("#sogjob").val();
        var sogjob_lukpas = 0;

        alert($get("sogjob_lukpas").checked)

        if ($get("sogjob_lukpas").checked == true) {
            sogjob_lukpas = 1;
        }
          

 
        //alert("her" + thisval + " m:" + usemrn)


        $.post("?sogval=" + sogVal + "&sogjob_lukpas=" + sogjob_lukpas, { control: "FN_sogjobliste", AjaxUpdateField: "true" }, function (data) {

            $("#jobprint_joblisten").html(data);

        });


    });
    
   
  
});



