$(document).ready(function () {

    //alert("awd")

    $(".submit").click(function () {

        var RFID_id = $("#RFID_field").val();
        //alert("hep")
        //alert(RFID_id)

        if (isNaN(RFID_id) == true)
        {

            decimalRfid = parseInt(RFID_id.toString(), 16)
            //alert(decimalRfid)

            RFID_idLength = RFID_id.length
            //alert(RFID_idLength)

            RfidZerozAdd = ""

            if (RFID_idLength < 10) {
                RFID_idZerosToAdd = 10 - RFID_idLength
                //alert(RFID_idZerosToAdd)
                RfidZeros = 0
                while (RfidZeros < (RFID_idZerosToAdd - 1)) {
                    RfidZeros += 1
                    RfidZerozAdd = RfidZerozAdd + "0"
                }
            }

            //alert(RfidZerozAdd)
            newDecimalRfid = RfidZerozAdd + decimalRfid

        }
        else
        {
            newDecimalRfid = RFID_id
        }
       
       //alert(newDecimalRfid)

       $("#RFID_field").val(newDecimalRfid);
       $("#scan").submit();



       // $("#scan").submit();
        //Alert("HER")
    });


   $('html, body').on('touchmove', function (e) {
        //prevent native touch activity like scrolling
        e.preventDefault();
    });

    $("#text_besked").delay(3000).fadeOut();

    $(".container").bind('click', function () {

        $("#RFID_field").focus();

    });

});


