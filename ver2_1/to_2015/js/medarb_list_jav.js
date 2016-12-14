
jQuery.extend(jQuery.fn.dataTableExt.oSort, {
    "date-dmy-pre": function (a) {
        if (a == null || a == "") {
            return 0;
        }
        var date = a.split('-');
        return (date[2] + date[1] + date[0]) * 1;
    },

    "date-dmy-asc": function (a, b) {
        return ((a < b) ? -1 : ((a > b) ? 1 : 0));
    },

    "date-dmy-desc": function (a, b) {
        return ((a < b) ? 1 : ((a > b) ? -1 : 0));
    }
});


$(document).ready(function () {

   

    // Setup - add a text input to each header cell
    $('#medarb_list thead th').not(":eq(1),:eq(2),:eq(3),:eq(5),:eq(6),:eq(7),:eq(8)").each(function () {

        var title = $('#medarb_list thead th').eq($(this).index()).text();
        $(this).html('<input type="text" placeholder="Søg ' + title + '" />');
    });

    // DataTable
    var table = $('#medarb_list').DataTable({
        
        "aoColumnDefs": [
        { "sType": "date-dmy", "aTargets": [5,6] }
        ]
    
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


