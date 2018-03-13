




$(document).ready(function () {

   

    $(".faktor").change(function () {

        

        var thisid = this.id

        var idlngt = thisid.length
        var idtrim = thisid.slice(10, idlngt)

        

        timer_gk_opr = $("#timer_gk_opr_" + idtrim).val().replace(',', '.')
        
        faktor_gk = $("#faktor_gk_" + idtrim).val().replace(',', '.')

        //alert(idtrim + " timer: " + timer_gk_opr + " faktor_gk: " + faktor_gk)

        if (faktor_gk != 1) {
            var omregnet = timer_gk_opr * 1 + (timer_gk_opr * (faktor_gk / 100)) * 1
            omregnet = omregnet.toString().replace('.', ',')
        } else {
            omregnet = $("#timer_gk_opr_" + idtrim).val()
        }
        //alert(omregnet)
       
        $("#timer_gk_" + idtrim).val(omregnet);

    });



    $("#gkall").click(function () {

        //alert("HER: " + $('#gkall').prop('checked'))
        if ($('#gkall').prop('checked') == true) {
            $('.gk').prop('checked', true);
        } else {
            $('.gk').prop('checked', false);
        }
    });


});



