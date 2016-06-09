



$(document).ready(function() {



    $("#sogjob").keyup(function () {

    
        var sogVal = $("#sogjob").val();

 
        //alert("her" + thisval + " m:" + usemrn)


        $.post("?sogval=" + sogVal, { control: "FN_sogjobliste", AjaxUpdateField: "true" }, function (data) {

            $("#jobprint_joblisten").html(data);

        });


    });
    
   
  
});



