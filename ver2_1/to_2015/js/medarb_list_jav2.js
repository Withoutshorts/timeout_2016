$(document).ready(function () {

  
   

    // Setup - add a text input to each header cell
    $('#medarb_list thead th').not(":eq(1),:eq(2),:eq(3),:eq(5),:eq(6),:eq(7),:eq(8)").each(function () {

        var title = $('#medarb_list thead th').eq($(this).index()).text();
        $(this).html('<input type="text" placeholder="Søg ' + title + '" />');
    });

    // DataTable
    var table = $('#medarb_list').DataTable({
        
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


   var table = $('#medarb_list').dataTable({
        "aoColumns": [
            null,
            null,
            null,
            null,
            null,
           { "sType": "date-uk" },
           { "sType": "date-uk" },
           null,
           null
        ]
   });

   jQuery.extend(jQuery.fn.dataTableExt.oSort, {
       "date-uk-pre": function (a) {
           var ukDatea = a.split('/');
           return (ukDatea[2] + ukDatea[1] + ukDatea[0]) * 1;
       },

       "date-uk-asc": function (a, b) {
           return ((a < b) ? -1 : ((a > b) ? 1 : 0));
       },

       "date-uk-desc": function (a, b) {
           return ((a < b) ? 1 : ((a > b) ? -1 : 0));
       }
   });
    

   

});


