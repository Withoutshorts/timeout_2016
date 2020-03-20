

$(document).ready(function () {


    
    if ($.cookie('udvidet_sog') == '1') {

        $('#udvidet_sog').css('visibility', 'visible')
        $('#udvidet_sog').css('display', '')
        $('#udvidet_sog').show(200)
    }

     
    $('#ud_s_span').click(function () {

      

        //alert("HER")
        if ($('#udvidet_sog').css("visibility") == "hidden") {
            $('#udvidet_sog').css('visibility', 'visible')
            $('#udvidet_sog').css('display', '')
            $('#udvidet_sog').show(200)

            $.cookie('udvidet_sog', '1');

        } else {

            $('#udvidet_sog').hide(200)
            $('#udvidet_sog').css('display', 'none')
            $('#udvidet_sog').css('visibility', 'hidden')

            $.cookie('udvidet_sog', '0');
           
        }

    });


    


    $('#ud_s_span').mouseover(function () {

        $(this).css("cursor", "pointer");

    });


  
    

    $('.date').datepicker({

    });

    //alert("HepHep")
    $("#actionMenu").change(function () {

        //alert("Bulk Menu")
        //alert($("#actionMenu").val())

        actionMenuValue = $("#actionMenu").val()

        jobids = ""
        $("input:checkbox[name=job_bulkCHB]:checked").each(function () {

            jobids = jobids + this.id + ","

        });


        // Slet
        if (actionMenuValue == "5") {

            $("#FM_selectedJobids_delete").val(jobids);

            var modal = document.getElementById('actionMenu_delete');

            modal.style.display = "block";

            window.onclick = function (event) {
                if (event.target == modal) {
                    modal.style.display = "none";
                }
            }
        }


        // Easyreg
        if (actionMenuValue == "4") {

            $("#FM_selectedJobids_easyreg").val(jobids);

            var modal = document.getElementById('actionMenu_easyreg');

            modal.style.display = "block";

            window.onclick = function (event) {
                if (event.target == modal) {
                    modal.style.display = "none";
                }
            }
        }


        // Job Status
        if (actionMenuValue == "3") {

            $("#FM_selectedJobids_status").val(jobids);

            var modal = document.getElementById('actionMenu_jobstatus');

            modal.style.display = "block";

            window.onclick = function (event) {
                if (event.target == modal) {
                    modal.style.display = "none";
                }
            }
        }

        // Job slutdato
        if (actionMenuValue == "2") {

            $("#FM_selectedJobids_enddate").val(jobids);

            var modal = document.getElementById('actionMenu_enddate');

            modal.style.display = "block";

            window.onclick = function (event) {
                if (event.target == modal) {
                    modal.style.display = "none";
                }
            }
        }

        // Forretnings områder
        if (actionMenuValue == "1") {

            $("#FM_selectedJobids_forOm").val(jobids);

            var modal = document.getElementById('actionMenu_forOm');

            modal.style.display = "block";

            window.onclick = function (event) {
                if (event.target == modal) {
                    modal.style.display = "none";
                }
            }
        }

    });



    function getKPerslisten() {

        var kundekpers = $("#FM_hd_kpers").val();
        var sog_val = $("#FM_kunde");



        //alert("her: " + sog_val.val() + " " + kundekpers)
        if (sog_val.val() != "") {

            $.post("?jq_kundekpers=" + kundekpers, { control: "FN_getKundeKperslisten", AjaxUpdateField: "true", cust2: sog_val.val() }, function (data) {
                //$("#FM_modtageradr").val(data);

                //var idlngt = sog_val.val()
                //alert(idlngt.lenght)
                //if (idlngt.lenght == 1) {
                //    alert("Du er ved at foretege en søgning..")
                //}
                //alert("Der søges...")
                $("#FM_kundekpers").html(data);

                //$("#jq_kunde_sel").focus();

                //$("#jobid").html(data);

            });

        }


    }


    $("#FM_kunde").change(function () {
        //alert("her")
        $("#FM_hd_kpers").val($("#FM_kundekpers").val());
        var kundekpers = $("#FM_hd_kpers").val();
        //alert("her: " + kundekpers);

        getKPerslisten();

    });

    getKPerslisten();

    $("#FM_kundekpers").change(function () {
        //alert($("#FM_kundekpers").val())
        $("#FM_hd_kpers").val($("#FM_kundekpers").val());
    });



    $(".picmodal").click(function () {

        var modalid = this.id
        var idlngt = modalid.length
        var idtrim = modalid.slice(6, idlngt)

        //var modalidtxt = $("#myModal_" + idtrim);
        var modal = document.getElementById('myModal_' + idtrim);

        modal.style.display = "block";

        window.onclick = function (event) {
            if (event.target == modal) {
                modal.style.display = "none";
            }
        }

    });




    // Setup - add a text input to each header cell
    $('#jobliste thead th').not(":eq(1),:eq(2),:eq(3),:eq(4),:eq(5),:eq(6),:eq(7),:eq(8),:eq(9)").each(function () {

        var title = $('#jobliste thead th').eq($(this).index()).text();
        var sogtekst = $("#sogtekst").val()
        $(this).html('<input type="text" placeholder="' + sogtekst + '" />');
    });

    // DataTable



    var table = $('#jobliste').DataTable({

        "aLengthMenu": [10, 25, 50, 75, 100, 500, 1000],
        "iDisplayLength": 1000,
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