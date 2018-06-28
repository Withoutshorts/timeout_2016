

$(document).ready(function() {



  /*  $(".korrigering_input").change(function () {

        thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(16, thisvallngt) 

       // alert(thisvaltrim)

       // alert("awd")
        var thisVal = $("#korrigering_opt_" + thisvaltrim).val();
        var currentOpt = $("#currentOpt_" + thisvaltrim).val();
       // alert(thisVal)

        currentOpt = currentOpt.replace(",", ".")

        var newOpt = parseFloat(currentOpt) + parseFloat(thisVal)
      //  alert(newOpt)

        document.getElementById('newOpt_' + thisvaltrim).setAttribute('value', newOpt);


    }); */

    $("#godkendalle").change(function () {

        if ($('#godkendalle').is(':checked') == true) {
            //alert("Godkend alle")

            $('.godkend_medarb').each(function () {
                $(this).prop('checked', true);
            });
        }
        if ($('#godkendalle').is(':checked') == false) {
            //alert("Godkend ingen")

            $('.godkend_medarb').each(function () {
                $(this).prop('checked', false);
            });
        }
        
    });

    $('input[type=radio][name=FM_medarb]').change(function () {

        if (this.value == '1') {
            $(".FM_medarbtype_SEL").prop("disabled", false);
            $(".FM_progrb_SEL").prop("disabled", true);
        }
        if (this.value == '2') {
            $(".FM_medarbtype_SEL").prop("disabled", true);
            $(".FM_progrb_SEL").prop("disabled", false);
        }

    });

    $('#feriekorri_multi_val').on('input', function (e) {

        //alert("click")
        multival = $("#feriekorri_multi_val").val();
       // alert(multival)

        $('.korrigering_input').each(function () {
            $(this).val(multival);
        });
    });

    $('#feriefriOpt_multi_val').on('input', function (e) {

       // alert("click")
        multival = $("#feriefriOpt_multi_val").val();
      //  alert(multival)

        $('.feriefriOpt_input').each(function () {
            $(this).val(multival);
        });
    });

    $('#scrollable').DataTable({
        "scrollY": "700px",
        "scrollCollapse": true,
        "paging": false
    });
    
});