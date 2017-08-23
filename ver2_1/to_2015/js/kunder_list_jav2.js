$(document).ready(function () {


    // Setup - add a text input to each header cell
    $('#kundeliste thead th').not(":eq(2),:eq(3),:eq(4),:eq(5),:eq(6),:eq(7),:eq(8),:eq(9)").each(function () {

        var title = $('#kundeliste thead th').eq($(this).index()).text();
        var sogtekst = $("#sogtekst").val()
        $(this).html('<input type="text" placeholder="' + sogtekst + '" />');
    });

    // DataTable



    var table = $('#kundeliste').DataTable({

        "aLengthMenu": [10, 25, 50, 75, 100, 500, 1000],
        "iDisplayLength": 1000,
        "order": [1, "asc"]

    });




    // Apply the search
    table.columns().every(function () {

        var coulmn = this;

        $('input', this.header()).on('keyup change', function () {
            if (coulmn.search() !== this.value) {
                coulmn
                    .search(this.value)
                    .draw();
            }
        });
    });





});