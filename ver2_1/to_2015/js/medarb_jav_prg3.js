

$(document).ready(function () {

    $('#scrollable').DataTable({
        "scrollY": "250px",
        "scrollCollapse": true,
        "paging": false,
        "ordering": true
    });


$('.panel').on('shown.bs.collapse', function (e) {
    //alert("Her")
    var table = $('#scrollable').DataTable();
    table.columns.adjust().draw();
})

$(".prg_navn_orPrg").click(function () {
          
    //alert("organitorisk")
    thisprgid = $(this).attr('value')
    //alert(thisprgid)

    if ($(this).is(':checked') == true) {
        $(".prg_navn_orPrg").each(function () {

            if ($(this).attr('value') != thisprgid) {
                $(this).prop('checked', false);
                $(this).attr('disabled', true);
                //alert("1")
            }

        });
    }
    else {
        $(".prg_navn_orPrg").each(function () {

            if ($(this).attr('value') != thisprgid) {
                $(this).prop('checked', false);
                $(this).attr('disabled', false);
                //alert("2")
            }

        });
    }
    //alert("3")

});

$(".tilfoj_fjern_prg").click(function () {
 
    if ($("#tilfoj_fjern_prg").is(':checked') == true) {
        $("#tilfoj_fjern_prg").prop("checked", false)
    } else {
        $("#tilfoj_fjern_prg").prop("checked", true)
    }

    $(".prg_navn").each(function () {

        //alert("HER 2" + this.id)
        //$(this).attr('checked', true);

        if ($("#tilfoj_fjern_prg").is(':checked') == true) {

            $(this).prop("checked", true)
            } else {
            $(this).prop("checked", false)

        }

    });

});

});
