$(document).ready(function() {


   $('.date').datepicker({

   });



    // Fuld - Delvis skift
    $(".delvisradio").click(function () {

        thisid = this.id

       // alert("value " + this.value);

        if (this.value == 1) {
            $(".fuld").each(function () {
                if (this.id == thisid) {
                    $(this).css("display", "none");
                }
            });
            $(".delvis").each(function () {
                if (this.id == thisid) {
                    $(this).css("display", "");
                }
            });

            $(".fuld_input").each(function () {
                if (this.id == thisid) {
                    $(this).val("");
                }
            });

        } else {
            $(".delvis").each(function () {
                if (this.id == thisid) {
                    $(this).css("display", "none");
                }
            });
            $(".fuld").each(function () {
                if (this.id == thisid) {
                    $(this).css("display", "");
                }
            });
        }

        


    });








    // Setup - add a text input to each header cell
    $('#kundeliste thead th').not(":eq(2),:eq(3),:eq(4),:eq(5),:eq(6),:eq(7),:eq(8),:eq(9)").each(function () {

        var title = $('#kundeliste thead th').eq($(this).index()).text();
        $(this).html('<input type="text" placeholder="S�g ' + title + '" />');
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

        
    $("#visflkonti").mouseover(function () {
        
        $(this).css('cursor', 'hand');
    });


    $("#visflkonti").click(function () {

        //alert("hht")
        if ($(".tr_konti").css('display') == "none" ) {
            $(".tr_konti").css("display", "");
            $(".tr_konti").css("visibility", "visible");
            $(".tr_konti").show("fast");
            //$.scrollTo('200px', 400);
            //$.scrollTo("#tr_konto_b", 4000);
            //$.scrollTop(300);

        } else {

            $(".tr_konti").hide("fast");
            //$.scrollTo('1000px', 400);
            //$.scrollTo('-=100px', 1500);
            //$.scrollTo('200px', 400);

        }
        
    });

   
    $(".tr_konti").hide("fast");

});