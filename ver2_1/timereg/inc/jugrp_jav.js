// JScript File

$(window).load(function() {
    // run code



//getAktlisten()

});

$(document).ready(function () {
    //alert("her")



    $(".beregn_tf").keyup(function () {

        var thisid = this.id
        var thisval = this.value
        var idlngt = thisid.length
        var idtrim = thisid.slice(8, idlngt)


        //alert("her" + idtrim)

        beregnSpris(idtrim,0)
    });



    function beregnSpris(id, dothis) {


        var iStk = $("#ulevstk_" + id + "").val().replace(".", "")
        iStk = iStk.replace(",", ".")
        iStk = Math.round(iStk * 100) / 100

        var iStkPris = $("#ulevstp_" + id + "").val().replace(".", "")
        iStkPris = iStkPris.replace(",", ".")
        iStkPris = Math.round(iStkPris * 100) / 100


       // alert("her")

        var sPris = (iStk * iStkPris)
        sPris = String(Math.round(sPris * 100) / 100).replace(".", ",")
        $("#ulevpri_" + id + "").val(sPris)


        var iPris = $("#ulevpri_" + id + "").val().replace(".", "")
        iPris = iPris.replace(",", ".")
        iPris = Math.round(iPris * 100) / 100

        var iFaktor = $("#ulevfak_" + id + "").val().replace(".", "")
        iFaktor = iFaktor.replace(",", ".")
        iFaktor = Math.round(iFaktor * 100) / 100

        var sPris = (iPris * iFaktor)
        sPris = String(Math.round(sPris * 100) / 100).replace(".", ",")
        $("#ulevbel_" + id + "").val(sPris)
       

    }


    $(".beregn_to").keyup(function () {

        var thisid = this.id
        var thisval = this.value
        var idlngt = thisid.length
        var idtrim = thisid.slice(8, idlngt)


        beregnIpris(idtrim, 0)
    });


    function beregnIpris(id, dothis) {



        var sPris = $("#ulevbel_" + id + "").val().replace(".", "")
        sPris = sPris.replace(",", ".")
        sPris = Math.round(sPris * 100) / 100

        var iFaktor = $("#ulevfak_" + id + "").val().replace(".", "")
        iFaktor = iFaktor.replace(",", ".")
        iFaktor = Math.round(iFaktor * 100) / 100

        var iPris = (sPris / iFaktor)
        iPris = String(Math.round(iPris * 100) / 100).replace(".", ",")

        var iStk = $("#ulevstk_" + id + "").val().replace(".", "")
        iStk = iStk.replace(",", ".")
        iStk = Math.round(iStk * 100) / 100

        iStkPris = String(Math.round((iPris / iStk) * 100) / 100).replace(".", ",")

        $("#ulevstp_" + id + "").val(iStkPris)
        $("#ulevpri_" + id + "").val(iPris)


    }

   

});



	