$(document).ready(function () {


    // Setup - add a text input to each header cell
    //$('#fckap thead th').not(":eq(0),:eq(1),:eq(2),:eq(3),:eq(4),:eq(5),:eq(6),:eq(7),:eq(8),:eq(9),:eq(10),:eq(11),:eq(12),:eq(13),:eq(14),:eq(15)").each(function () {

    //    var title = $('#fckap thead th').eq($(this).index()).text();
    //    var sogtekst = $("#sogtekst").val()
    //    $(this).html('<input type="text" placeholder="' + sogtekst + '" />');

    //});

    // DataTable



    var table = $('#fckap').DataTable({

        "aLengthMenu": [10, 25, 50, 75, 100, 500, 1000],
        "iDisplayLength": 100,
        "order": [1, "asc"]

    });

  


    // Apply the search
    //table.columns().every(function () {

    //    var coulmn = this;

    //    $('input', this.header()).on('keyup change', function () {
    //        if (coulmn.search() !== this.value) {
    //            coulmn
    //                .search(this.value)
    //                .draw();
    //        }
    //    });
    //});





});