$(document).ready(function () {

    $(".guestsignField").bind('change', function () {
        
        CheckGuestSignFields();

        if (messege.length <= 0) {
            $("#guestsub").removeClass("btn-default");
            $("#guestsub").addClass("btn-success");
            $("#guestErrorMessege").html(""); 
        } else {
            $("#guestsub").removeClass("btn-success");
            $("#guestsub").addClass("btn-default");
        }
        
    });
    
    
    $("#guestsub").bind('click', function () {

        CheckGuestSignFields();

        $("#guestErrorMessege").html(messege); 

        if (messege.length <= 0) {

            lto = $("#lto").val();

            if (lto == "cflow") {
                //Sender sms at gæst er ankommet
                visitid = $("#visit").val();
                strTlf = ""
                
                $.post("?medid=" + visitid, { control: "GetVisitsTel", AjaxUpdateField: "true" }, function (data) {

                    strTlf = data
                    
                    if (strTlf.length > 8) {
                        
                        guest = $("#guestEmployee").val();
                        guest = guest.toString().replace("æ", "ae")
                        guest = guest.toString().replace("ø", "o")
                        guest = guest.toString().replace("å", "aa")
                        guest = guest.toString().replace("Æ", "Ae")
                        guest = guest.toString().replace("Ø", "O")
                        guest = guest.toString().replace("Å", "Aa")
                        
                        messege = guest + " er ankommet"
                        recipient = strTlf
                        
                      // var smswin = window.open("https://gatewayapi.com/rest/mtsms?token=gAws4-yJSP-ySXOjoDRNcgKvWgTzAbnB1nFEuOdPusswBtRWDQRqiHk35uE8rdYO&message=" + messege + "&recipients.0.msisdn=" + recipient)
                       // smswin.close();
                        
                        $.post("https://gatewayapi.com/rest/mtsms?token=gAws4-yJSP-ySXOjoDRNcgKvWgTzAbnB1nFEuOdPusswBtRWDQRqiHk35uE8rdYO&message=" + messege + "&recipients.0.msisdn=" + recipient, function (data) {
                            
                        });

                    }

                });

            }

            $("#guestlogin").submit();
        } 
        
    });


    var errorMessege1 = $("#errorMessege1").val();
    var errorMessege2 = $("#errorMessege2").val();
    var errorMessege3 = $("#errorMessege3").val();
    var errorMessege4 = $("#errorMessege4").val();
    var errorMessege5 = $("#errorMessege5").val();
    var errorMessege6 = $("#errorMessege6").val();

    function CheckGuestSignFields() {

        var lto = $("#lto").val();

        var virksomhed = $("#guestfirma").val();
        var navn = $("#guestEmployee").val();
        var tlf = $("#guestTlf").val();
        var visit = $("#besogInput").val();

        messege = ""

        messege = ($('#infoApr').is(":checked")) ? messege : errorMessege1.toString();
        messege = (visit.length <= 0) ? errorMessege2.toString() : messege;
        if (lto != "cool") {
            messege = (tlf.length <= 0) ? errorMessege3.toString() : messege;
        }
        messege = (navn.length <= 0) ? errorMessege4.toString() : messege;

        if (navn.indexOf(' ') == -1 && navn.length > 0) {
            messege = errorMessege5.toString()
        }

        messege = (virksomhed.length <= 0) ? errorMessege6.toString() : messege;
    }


    $("#besogInput").bind('keyup', function () {
        besogInput = $("#besogInput").val();

        if (besogInput != "" && besogInput.length > 0) {

            $.post("?besogInput=" + besogInput, { control: "FindEmployee", AjaxUpdateField: "true" }, function (data) {

                $("#besogslist").html(data)
                $("#besogslist").show();

                $("#besogslist").bind('click', function () {
                    var employeeSEL = $("#besogslist").val();
                    var employeeName = $("#besogslist").find(':selected').data('name');

                    if (employeeSEL != null && employeeSEL != "" && employeeSEL.length > 0) {
                        $("#besogInput").val(employeeName);
                        $("#visit").val(employeeSEL);

                        $("#besogslist").html("");
                        $("#besogslist").hide();
                    }

                });

            });

        } else {
            $("#besogslist").html("")
            $("#besogslist").hide();
        }

    });


    $("#guestfirma").bind('keyup', function () {

        guestval = $("#guestfirma").val();
        //alert(guestval)

        if (guestval != "" && guestval.length > 0) {
            $.post("?guestval=" + guestval, { control: "FindGuestCompany", AjaxUpdateField: "true" }, function (data) {

                if (data.length > 0) {
                    $("#guestfirmalist").html(data)
                    $("#guestfirmalist").show();
                } else {
                    $("#guestfirmalist").html('')
                    $("#guestfirmalist").hide();
                }

                $("#guestfirmalist").bind('click', function () {
                    var customerSEL = $("#guestfirmalist").val();

                    if (customerSEL != null && customerSEL != "" && customerSEL.length > 0) {
                        $("#guestfirma").val(customerSEL);

                        $("#guestfirmalist").html("")
                        $("#guestfirmalist").hide();
                    }


                });

            });
        } else {
            $("#guestfirmalist").html("")
            $("#guestfirmalist").hide();
        }


    });

    $("#guestEmployee").bind('keyup', function () {
        guestval = $("#guestEmployee").val();
        guestCompany = $("#guestfirma").val();

        if (guestval != "" && guestval.length > 0) {
            $.post("?guestval=" + guestval + "&guestCompany=" + guestCompany, { control: "FindGuestEmployee", AjaxUpdateField: "true" }, function (data) {

                if (data.length > 0) {
                    $("#guestEmployeelist").html(data)
                    $("#guestEmployeelist").show();
                } else {
                    $("#guestEmployeelist").html('')
                    $("#guestEmployeelist").hide();
                }

                $("#guestEmployeelist").bind('click', function () {
                    var guestSEL = $("#guestEmployeelist").val();
                    var guestTlf = $("#guestEmployeelist").find(':selected').data('tlf');

                    if (guestSEL != null && guestSEL != "" && guestSEL.length > 0) {
                        $("#guestEmployee").val(guestSEL);
                        $("#guestTlf").val(guestTlf);

                        $("#guestEmployeelist").html("")
                        $("#guestEmployeelist").hide();
                    }


                });

            });
        } else {
            $("#guestEmployeelist").html("")
            $("#guestEmployeelist").hide();
        }
    });


    $("#gustlist").bind('click', function () {
        $("#guestlistmodal").show();

        $("#closeguestlist").click(function () {
            $("#guestlistmodal").hide();
        });

    });
    

    vislogud = $("#vislogud").val();
    if (vislogud == "1") {
        $("#logudmodal").show();
    }

    var guestmodal = 0;
    $("#gustlogin").bind('click', function () {

        $("#guestmodal").show();
        guestmodal = 1;

        $("#closeguestmodal").click(function () {
            $("#guestmodal").hide();
            guestmodal = 0;
        });

    });


    $(".picmodal").click(function () {

        var modalid = this.id
        var idlngt = modalid.length
        var idtrim = modalid.slice(6, idlngt)


        //alert(idtrim)
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


    function LoadJobs(sorting) {
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

        lto = $("#lto").val()

        //alert(jq_newfilterval)
        //alert(varTjDatoUS_man)
        //alert(jq_jobidc)
        //alert(jq_medid + " ; " + jq_jobidc + " ; " + jq_newfilterval + " ; " + varTjDatoUS_man)

        $.post("?lto=" + lto + "&jq_newfilterval=" + jq_newfilterval + "&jq_medid=" + jq_medid + "&varTjDatoUS_man=" + varTjDatoUS_man + "&jq_jobidc=" + jq_jobidc + "&searchmode=" + searchmode + "&sogSort=" + sogSort, { control: "FN_sogjobogkunde", AjaxUpdateField: "true" }, function (data) {
            // alert("cc")

            $("#modal_job_SEL").html(data);


            var countTD = $("#job_table tbody tr:last > td").length;
            // alert("TD " + countTD)

            if (countTD < 5) {
                var i = countTD
                while (i < 5) {
                    $("#job_table tbody tr:last").append("<td style='display:none;'>&nbsp;<td>")
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

                    $(".jobsel").unbind('click').bind('click', function () {
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


    function LoadActivities(sorting) {
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

            if (countTD < 4) {
                var i = countTD
                while (i < 4) {
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

                    $(".aktsel").unbind('click').bind('click', function () {
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
        if (guestmodal != 1) {
            $("#RFID_field").focus();
        }

    });




    //alert($("#skiftjob").val())

    if ($("#skiftjob").val() == "1") {

        //alert($("#lto").val())

        var modal = document.getElementById('myModal_picker');

        modal.style.display = "block";

        window.onclick = function (event) {
            if (event.target == modal) {
                modal.style.display = "none";
            }
        }

        LoadJobs(1);

        //alert($("#jq_jobidc").val())


    };




});


