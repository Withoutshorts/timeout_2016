$(document).ready(function() {

    $("#allejob").click(function() {

        if ($("#allejob").is(':checked') == true) {
            $(".alljobCHK").attr('checked', true);
        } else {
            $(".alljobCHK").attr('checked', false);

        }

    });


    $("#FM_jobid").change(function() {
        aktiviteter_liste();
    });



    function aktiviteter_liste() {
    
        var thisC = $("#FM_jobid")
        //alert("her" + thisC)
        $.post("?func=opd_guide", { control: "FN_getCustDesc", AjaxUpdateField: "true", cust: thisC.val() }, function(data) {
            //$("#fajl").val(data);
            $("#jq_aktlist").html(data);

            //alert("her2")

        });

        //alert("her3")

    }


    aktiviteter_liste();



});