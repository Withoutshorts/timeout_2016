$(document).ready(function () {


    $("#FM_copy_follow_activejoblist1").click(function () {

        $("#FM_copy_follow_activejoblist2").prop('checked', false);

    });

    $("#FM_copy_follow_activejoblist2").click(function () {

        $("#FM_copy_follow_activejoblist1").prop('checked', false);

    });


    $("#FM_medtype").change(function () {


        strMedarbejdertype = $("#FM_medtype").val()
        strMedarbejderid = $("#medarbejderId").val()
        beforeempltype_id = $("#beforeempltype_id").val()
        //alert(strMedarbejderid)
        
        //$("#newempltype").attr('style', 'visibilty:inherit');
        $("#newempltype_name").html($("#FM_medtype option:selected").text());
        $("#newempltype_id").val(strMedarbejdertype);
        //alert(strMedarbejdertype)
        //alert(beforeempltype_id)
        if (strMedarbejdertype == beforeempltype_id) {  
            $("#newempltype").attr('style', 'display:none');
        } else {
            $("#newempltype").attr('style', 'display:normal');
        }


    });


    $("#FM_ansat2").click(function () {

        $("#deaktivnote").css("display", "");
        $("#deaktivnote").css("visibility", "visible");
        $("#deaktivnote").show("fast");

    });

    $("#FM_ansat1, #FM_ansat3").click(function () {

        
        
        $("#deaktivnote").hide("fast");

    });
   

    // Setup - add a text input to each header cell
    $('#medarb_list thead th').not(":eq(1),:eq(2),:eq(3),:eq(5),:eq(6),:eq(7)").each(function () {

        var title = $('#medarb_list thead th').eq($(this).index()).text();
        $(this).html('<input type="text" placeholder="Søg ' + title + '" />');
    });

    // DataTable
    var table = $('#medarb_list').DataTable({
        
    
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


