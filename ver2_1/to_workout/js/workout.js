





$(document).ready(function () {


    //alert("godkendt");

    $('.date').datepicker({

    });


    $("#theme").change(function () {
       

            //alert(thisMid)
            var theme = $("#theme").val();
         

            $.post("?theme=" + theme, { control: "FN_hentakt", AjaxUpdateField: "true" }, function (data) {

                 //alert("OK")

                 //alert("godkendt: " + data);
                 //$("#test").val(data);
                 $("#FM_activity").html(data);

            });

    });


    $(".dvcon").mouseover(function () {

        $(this).css("cursor", "pointer");

    });

    $(".dvakt").mouseover(function () {

        $(this).css("cursor", "pointer");

    });

    $(".dvcon").click(function () {

      
        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(6, thisvallngt)
        thisval = thisvaltrim

        $("#dvakt_" + thisval).css("display", "");
        $("#dvakt_" + thisval).css("visibility", "visible")

    });

    $(".dvakt").click(function () {


        var thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(6, thisvallngt)
        thisval = thisvaltrim

        $("#dvakt_" + thisval).hide('fast')

        $("#dvakt_" + thisval).css("display", "none");
        $("#dvakt_" + thisval).css("visibility", "hidden")

    });

});






