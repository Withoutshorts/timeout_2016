













$(document).ready(function () {


   

    $("#kundeid").change(function () {

        joblist();

    });




    function joblist() {

   
        var kundeid = $("#kundeid").val()
        $.post("?kundeid=" + kundeid, { control: "FN_showjob", AjaxUpdateField: "true" }, function (data) {

            $("#FM_folderid").html(data);
            
        });
    }

    
        
   

});


