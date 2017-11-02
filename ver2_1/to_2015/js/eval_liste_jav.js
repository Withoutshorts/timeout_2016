


$(document).ready(function ()
{

    $('#eval_liste').DataTable({

        "aoColumnDefs": [
        { "sType": "date-dmy", "aTargets": [5] }
        ],
        "aLengthMenu": [10, 25, 50, 75, 100, 500, 1000],
        "iDisplayLength": 100,
        "order": [0, "asc"]

    });

});