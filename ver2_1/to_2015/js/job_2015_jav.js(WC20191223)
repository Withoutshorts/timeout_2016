// JScript File
$(window).load(function () {
    // run code


    $("#loadbar").hide(1000);
});






$(document).ready(function () {


    $('.date').datepicker({

    });

   
    $("#FM_kunde").change(function () {
        getKpers();
    });

    // kontakpers //
    function getKpers() {
     
        var kundekpersopr_val = $("#kundekpersopr").val()
        var kid_val = $("#FM_kunde").val()

        //alert("her " + kid_val)

        $.post("?jq_kid=" + kid_val + "&jq_kundekpers=" + kundekpersopr_val, { control: "FN_kpers", AjaxUpdateField: "true", cust: 0 }, function (data) {
            //$("#FM_modtageradr").val(data);
            $("#FM_kpers").html(data);
            //$("#jobid").html(data);
            //alert(data)

        });


    }


});