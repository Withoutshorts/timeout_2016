// JScript File




$(window).load(function () {
    // run code
    $("#loadbar").hide(1000);



});





$(document).ready(function () {

    //alert("Så er siden klar..")

  


    $(".rmenu").click(function () {
        var thisid = this.id
        var idlngt = thisid.length
        var idtrim = thisid.slice(2, idlngt)

        var lginTid = $("#FM_login_hh_" + idtrim).val()
        var lgudTid = $("#FM_login_mm_" + idtrim).val()
               
        //alert(idtrim)
        $("#FM_logud_hh_" + idtrim).val(lginTid)
        $("#FM_logud_mm_" + idtrim).val(lgudTid)
    });



});