
$(document).ready(function () {


    $('.date').datepicker({

    });

    $("#kontoplan_table").DataTable({
        scrollY: "800px",
        scrollCollapse: true,
        paging: false,
        ordering: false
    });

   /* $("#kontoplan_table").DataTable({
        "language": {
            "decimal": ",",
            "thousands": "."
        }
    }); */

});
