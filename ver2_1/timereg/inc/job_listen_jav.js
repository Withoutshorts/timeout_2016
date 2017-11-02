// JScript File



$(window).load(function () {
    // run code
    $("#loadbar").hide(1000);
});



// Kundesøg //
function getKundelisten() {

    var visalle = 1 //bruges ikke
    var sog_val = $("#kunde_sog_step1")

    //alert("her: " + sog_val.val())
    if (sog_val.val() != "") {

        $.post("?jq_sogkunde=" + visalle, { control: "FN_getKundelisten_2013", AjaxUpdateField: "true", cust: sog_val.val() }, function (data) {
            //$("#FM_modtageradr").val(data);

            //var idlngt = sog_val.val()
            //alert(idlngt.lenght)
            //if (idlngt.lenght == 1) {
            //    alert("Du er ved at foretege en søgning..")
            //}
            //alert("Der søges...")
            //$("#test").val(data);
            $("#jq_kunde_sel").html(data);
          
            //$("#jq_kunde_sel").focus();

            //$("#Submit2").focus();


            //$("#jobid").html(data);

        });

    }


}




$(document).ready(function () {



    $("#st_cls_1").click(function () {

       

        $(".FM_listestatus_1").attr("checked", "checked");

        $("#st_cls_1").attr("checked", "checked");

      

    });

    $("#st_cls_2").click(function () {

        

        $(".FM_listestatus_2").attr("checked", "checked");
        $("#st_cls_2").attr("checked", "checked");
        

    });

    $("#st_cls_3").click(function () {

        

        $(".FM_listestatus_3").attr("checked", "checked");
        $("#st_cls_3").attr("checked", "checked");
        

    });

    $("#st_cls_4").click(function () {

        

        $(".FM_listestatus_4").attr("checked", "checked");
        $("#st_cls_4").attr("checked", "checked");
       
    });

    $("#st_cls_5").click(function () {



        $(".FM_listestatus_5").attr("checked", "checked");
        $("#st_cls_5").attr("checked", "checked");

    });


    $("#st_cls_0").click(function () {

       

        $(".FM_listestatus_0").attr("checked", "checked");
        $("#st_cls_0").attr("checked", "checked");
      

    });

   


    //alert("Så er siden klar..")

   


    function getKPerslisten() {

        var kundekpers = $("#FM_hd_kpers").val();
        var sog_val = $("#FM_kunde");



        //alert("her: " + sog_val.val() + " " + kundekpers)
        if (sog_val.val() != "") {

            $.post("?jq_kundekpers=" + kundekpers, { control: "FN_getKundeKperslisten", AjaxUpdateField: "true", cust2: sog_val.val() }, function (data) {
                //$("#FM_modtageradr").val(data);

                //var idlngt = sog_val.val()
                //alert(idlngt.lenght)
                //if (idlngt.lenght == 1) {
                //    alert("Du er ved at foretege en søgning..")
                //}
                //alert("Der søges...")
                $("#FM_kundekpers").html(data);

                //$("#jq_kunde_sel").focus();

                //$("#jobid").html(data);

            });

        }


    }

    $("#FM_kunde").change(function () {
        //alert("her")
        $("#FM_hd_kpers").val($("#FM_kundekpers").val());
        var kundekpers = $("#FM_hd_kpers").val();
        //alert("her: " + kundekpers);

        getKPerslisten();

    });

    getKPerslisten();

    $("#FM_kundekpers").change(function () {
        //alert($("#FM_kundekpers").val())
        $("#FM_hd_kpers").val($("#FM_kundekpers").val());
    });


    $("#kunde_sog_step1").keyup(function () {
        getKundelisten();
    });

    $("#kunde_sog_step1_but").click(function () {
        getKundelisten();
    });


});