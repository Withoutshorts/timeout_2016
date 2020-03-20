$(document).ready(function () {


    if ($("#skiftjob").val() == 1) {
        LoadJobs(1);
    };


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

    $("#closebtn").click(function () {

        $("#myModal_picker").hide();

    });
   
    $(".loadjob").click(function () {
        LoadJobs(1);
    });

    $(".loadact").click(function () {

        FM_jobid_valgt = $("#FM_jobid_valgt").val();

        if (FM_jobid_valgt == "-1") {
            LoadJobs(1);
        } else {
            LoadActivities(1);
        }      

    });
    

    $(".scrollDown").mousedown(function () {
        
        $("#jobSection").scrollTop($("#jobSection").scrollTop() + 100)

    });

    $(".scrollUp").click(function () {
 
        $("#jobSection").scrollTop($("#jobSection").scrollTop() - 100)

    });

    navnOrderMode = 1
    $("#aktsort_navn").click(function () {

        if (navnOrderMode == 1) {
            LoadActivities(2);
            navnOrderMode = 2
        }
        else {
            LoadActivities(1);
            navnOrderMode = 1
        }

    });

    navnOrderMode = 1
    $("#jobsort_navn").click(function () {

        if (navnOrderMode == 1) {
            LoadJobs(2);
            navnOrderMode = 2
        }
        else {
            LoadJobs(1);
            navnOrderMode = 1
        }

    });

    $(".backbtn").click(function () {

        LoadJobs(1);

    });


    function LoadJobs(sorting)
    {
        $("#modal_overskrift").html("&nbsp; Velg Jobb &nbsp;")

        //$("#modal_overskrift").css("background-color", "yellow")
        $("#overskriftspan").css("background-color", "")

        $("#modal_underskrift").hide()
        
        if ($.fn.dataTable.isDataTable('#job_table')) {  
            table = $('#job_table').DataTable();
            table.destroy();
        }

        if ($.fn.dataTable.isDataTable('#akt_table')) {
            table = $('#akt_table').DataTable();
            table.destroy();
        }

        $("#akt_table").hide();
        $("#job_table").show();

        $("#aktsort_navn").hide();
        $("#jobsort_navn").show();

        $("#modal_job_SEL").html("");

        jq_newfilterval = $("#FM_job_0").val()
        jq_medid = $("#FM_medid").val()
        varTjDatoUS_man = $("#varTjDatoUS_man").val()
        jq_jobidc = $("#jq_jobidc").val()
        searchmode = 1
        sogSort = sorting


        // alert(jq_newfilterval)
        // alert(varTjDatoUS_man)
        // alert(jq_jobidc)

        $.post("?lto=" + lto + "&jq_newfilterval=" + jq_newfilterval + "&jq_medid=" + jq_medid + "&varTjDatoUS_man=" + varTjDatoUS_man + "&jq_jobidc=" + jq_jobidc + "&searchmode=" + searchmode + "&sogSort=" + sogSort, { control: "FN_sogjobogkunde", AjaxUpdateField: "true" }, function (data) {
           // alert("cc")

            $("#modal_job_SEL").html(data);
           

            var countTD = $("#job_table tbody tr:last > td").length;
           // alert("TD " + countTD)

            if (countTD < 5) {
                var i = countTD
                while (i < 5) {
                    $("#job_table tbody tr:last").append("<td style='display:none;'>Hej<td>")
                    i += 1
                }
                
            }

            $.fn.dataTable.ext.classes.sPageButton = 'btn btn-default btn-paging';
            $.fn.dataTable.ext.classes.sPageButtonActive = 'btn btn-default btn-paging-active';

            var table = $('#job_table').DataTable({
                "pageLength": 4,
                "lengthChange": false,
                "ordering": false,
                "pagingType": "simple_numbers",
                "bInfo": false,
                "fnDrawCallback": function (oSettings) {
                   // alert('DataTables has redrawn the table');
                   
                   $(".jobsel").click(function () {
                      // alert("Klik")
                       thisjobid = this.id
                       //alert(thisjobid)
                       //alert($("#jobnavn_sel_" + thisjobid).val())
                       $("#FM_jobnavn").html($("#jobnavn_sel_" + thisjobid).val())
                       $("#FM_jobid_0").val(thisjobid)
                       $("#FM_jobid_valgt").val(thisjobid);


                       $("#FM_aktnavn").html('')
                       $("#FM_aktid_0").val('-1')
                       $("#FM_aktivitetid").val('-1')


                       $.post("?jq_jobidc=" + thisjobid, { control: "FN_gemjobc", AjaxUpdateField: "true" }, function (data) {

                           //alert("HER OK")
                           //$("#test").val(data);

                       });

                       LoadActivities(1);
                       
                    });
                }
            });

            $(".backbtn").css("visibility", "hidden")
            /* dom: 'l<"toolbar">frtip',
                initComplete: function () {
                    $("div.toolbar").html('<span style="font-size:150%; background-color:#bcbcbc; margin-left:25px;" class="btn btn-default jobsort_navn sortbtn"><b>Navn</b> <span class="fa fa-sort"></span></span>');
                } */
           
        });
    }


    function LoadActivities(sorting)
    {
        $("#modal_overskrift").html("&nbsp;" + $("#FM_jobnavn").html() + "&nbsp;")

        $("#modal_overskrift").css("background-color", "")
        //$("#overskriftspan").css("background-color", "orange")

        $("#modal_underskrift").show()

        if ($.fn.dataTable.isDataTable('#job_table')) {
            table = $('#job_table').DataTable();
            table.destroy();
        }

        if ($.fn.dataTable.isDataTable('#akt_table')) {
            table = $('#akt_table').DataTable();
            table.destroy();
        }

        $("#job_table").hide();
        $("#akt_table").show();

        $("#jobsort_navn").hide();
        $("#aktsort_navn").show();

        $("#sort_btn").removeClass("jobsort_navn")
        $("#sort_btn").addClass("aktsort_navn")
        
        $("#modal_akt_SEL").html("");

        mobil_week_reg_akt_dd = $("#mobil_week_reg_akt_dd").val()
        mobil_week_reg_job_dd = $("#mobil_week_reg_job_dd").val()

        varTjDatoUS_man = $("#varTjDatoUS_man").val()

        jq_newfilterval = "-1"
        jq_jobid = $("#FM_jobid_valgt").val()
        
        jq_medid = $("#FM_medid").val()
        jq_aktid = $("#FM_akt_0").val()
        jq_pa = $("#FM_pa").val()

        jq_aktidc = $("#jq_aktidc").val()

        searchmode = 1
        sogSort = sorting

        $.post("?jq_newfilterval=" + jq_newfilterval + "&jq_jobid=" + jq_jobid + "&jq_medid=" + jq_medid + "&jq_aktid=" + jq_aktid + "&jq_pa=" + jq_pa + "&varTjDatoUS_man=" + varTjDatoUS_man + "&jq_aktidc=" + jq_aktidc + "&searchmode=" + searchmode + "&sogSort=" + sogSort, { control: "FN_sogakt", AjaxUpdateField: "true" }, function (data) {
           //alert("cc")
            //$("#dv_akt_test").html(data);
            $("#modal_akt_SEL").html(data);

            var countTD = $("#akt_table tbody tr:last > td").length;
           // alert("TD " + countTD)

            if (countTD < 4)
            {
                var i = countTD
                while (i < 4)
                {                   
                    $("#akt_table tbody tr:last").append("<td style='display:none;'>Hej<td>")
                    i += 1
                }                
            }

            $.fn.dataTable.ext.classes.sPageButton = 'btn btn-default btn-paging';
            $.fn.dataTable.ext.classes.sPageButtonActive = 'btn btn-default btn-paging-active';

            var table = $('#akt_table').DataTable({
                "pageLength": 4,
                "lengthChange": false,
                "ordering": false,
                "pagingType": "simple_numbers",
                "bInfo": false,
                "fnDrawCallback": function (oSettings) {
                    // alert('DataTables has redrawn the table');

                    $(".aktsel").click(function () {
                    //    alert("Klik")
                        thisaktid = this.id
                        //alert(thisjobid)
                      //  alert($("#aktnavn_sel_" + thisaktid).val())
                        $("#FM_aktnavn").html($("#aktnavn_sel_" + thisaktid).val())
                        $("#FM_aktid_0").val(thisaktid)
                        //   $("#FM_jobid_valgt").val(thisjobid);
                        $("#FM_aktivitetid").val(thisaktid)

                        $.post("?jq_aktidc=" + thisaktid, { control: "FN_gemaktc", AjaxUpdateField: "true" }, function (data) {

                        });

                        $("#myModal_picker").hide();
                    });

                    $(".backbtn").css("visibility", "")                  
                }
            });

        });

        
    }



    $("#RFID_field").change(function () {

        var RFID_id = $("#RFID_field").val();
        //alert("hep")
        //alert(RFID_id)

        $("#scan").submit();
        //Alert("HER")
    });


    $('html, body').on('touchmove', function (e) {
        //prevent native touch activity like scrolling
        e.preventDefault();
    });

    $("#text_besked").delay(3000).fadeOut();

    $(".container").bind('click', function () {

        //alert("FOCUS")

        $("#RFID_field").focus();

    });

});


