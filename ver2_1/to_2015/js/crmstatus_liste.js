$(document).ready(function () {

    // Setup - add a text input to each header cell
    $('#crm_liste thead th').not(":eq(2),:eq(3),:eq(5),:eq(6),:eq(7)").each(function () {

        var title = $('#crm_liste thead th').eq($(this).index()).text();
        $(this).html('<input type="text" placeholder="Search ' + title + '" />');
    });

    // DataTable
    var table = $('#crm_liste').DataTable({
        
        "aLengthMenu": [10, 25, 50, 75, 100, 500, 1000],
        "iDisplayLength": 100,
        "order": [0, "asc"]
    
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


