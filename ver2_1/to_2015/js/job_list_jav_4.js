

$(document).ready(function () {

    $.fn.dataTable.moment();

    $('#multiselect').click(function () {
        showCheckboxes();
    });



    //lastjobid
    lastjobid = $("#lastjobid").val()

    if (lastjobid != 0) {

        $([document.documentElement, document.body]).animate({
            scrollTop: $("#tr_jobid_" + lastjobid).offset().top - 200
        }, 1000);

    }

        
    

    StatusValgt();

    var expanded = false;

    function showCheckboxes() {

        var checkboxes = document.getElementById("checkboxes");
        if (!expanded) {
            checkboxes.style.display = "block";
            expanded = true;
            StatusfiltOpenCookie(1)
        } else {
            checkboxes.style.display = "none";
            expanded = false;
            StatusfiltOpenCookie(0)
        }
    }

    $('.statusfilt').change(function () {
        thisid = this.id

        strstatus = " (" + $('#status_navn_' + thisid).html() + ")"

       // alert(strstatus)

        strstatusall = $('#statustxt').html();

        if ($(this).prop('checked') == true) {

            if (strstatusall.indexOf(strstatus)) {
                $('#statustxt').append(strstatus)
            }
        }
        else {
            strstatusall = strstatusall.replace(strstatus, "");
            $('#statustxt').html(strstatusall)
        }

        StatusValgt();

    });

    function StatusfiltOpenCookie (val)
    {
        $.post("?open=" + val, { control: "statusfil_cookie", AjaxUpdateField: "true", cust: 0 }, function (data) {

          //  alert("Hep")

        });
    }

    function StatusValgt()
    {
        var valgte = 0
        $('.statusfilt').each(function () {
            if ($(this).prop('checked') == true)
            {
                valgte += 1;
            }
        });

        if (valgte == 0) {
            $('#statustxt').html("Vælg Status")
        } else {
            strnew = $('#statustxt').html();
            strnew = strnew.replace("Vælg Status", "")
            $('#statustxt').html(strnew)
        }
    }


    SkiftVisnign();


    $('#skiftvisning').click(function () {

        // alert("KLik")
        SkiftVisnign();

    });

    function SkiftVisnign() {
        simpleliste = $('#FM_simplelist').val();
        // alert(simpleliste)
        if (simpleliste == 0) {
            $('.list_advanced').hide()
            $('.list_simple').show()
            $('#FM_simplelist').val(1)
            $('#skiftvisning').html("<b>< Advanced ></b>")
        } else {
            
            $('.list_advanced').show()
            $('.list_simple').hide()
            $('#FM_simplelist').val(0)
            $('#skiftvisning').html("<b>< Simple ></b>")
        }

        // simpleliste = $('#FM_simplelist').val();
        $.post("?simpleliste=" + simpleliste, { control: "simpleliste_cookie", AjaxUpdateField: "true", cust: 0 }, function (data) {

            //  alert("Hep")

        });

        if ($.fn.dataTable.isDataTable('#xjobliste')) {
            table = $('#xjobliste').DataTable();
            table.destroy();
        }

        if (simpleliste == 0) {
            if (!$.fn.DataTable.isDataTable('#xjobliste')) {

                var table = $('#xjobliste').DataTable({
                    "ordering": true,
                    "paging": false,
                    "bInfo": false,
                    "bLengthChange": false
                });
            }
        }



    }


    if ($.cookie('udvidet_sog') == '1') {
        $('#udvidet_sog').css('visibility', 'visible')
        $('#udvidet_sog').css('display', '')
        $('#udvidet_sog').show(200)
    }


    $('#selectall').change(function () {
        if ($('#selectall').prop('checked') == true) {
            $('.job_bulkCHB').prop('checked', true)
        } else {
            $('.job_bulkCHB').prop('checked', false)
        }
    });

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