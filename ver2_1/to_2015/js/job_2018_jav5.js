// JScript File
$(window).load(function () {
    // run code


    $("#loadbar").hide(1000);
});






$(document).ready(function () {

    $('#opdaterafslut').click(function () {
        $('#afslut_val').val('1')
        $('#opretproj').submit();
    });


    /*
    $('#ud_s_span').click(function () {

        alert("HER")

        $('#udvidet_sog').css('visibility', 'visible')
        $('#udvidet_sog').css('display', '')

    });


    $('#ud_s_span').mouseover(function () {

        $(this).css("cursor", "pointer");

    });

    */

    $('.date').datepicker({

    });

    AktView1();

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


    getKpers();

    $("#fornmrbtn").click(function () {

        if ($("#forretningsomroderDropDown").is(":hidden")) {
            $("#forretningsomroderDropDown").show();
        } else {
            $("#forretningsomroderDropDown").hide();
        }

    });


    $("#FM_jobans_proc_1").keyup(function () {
        jobandelmaksval(1)
    });

    $("#FM_jobans_proc_2").keyup(function () {
        jobandelmaksval(2)
    });

    $("#FM_jobans_proc_3").keyup(function () {
        jobandelmaksval(3)
    });

    $("#FM_jobans_proc_4").keyup(function () {
        jobandelmaksval(4)
    });

    $("#FM_jobans_proc_5").keyup(function () {
        jobandelmaksval(5)
    });



    function jobandelmaksval(id) {

        var totalVal = 0;
        totalVal = $("#FM_jobans_proc_1").val().replace(",", ".") / 1 + $("#FM_jobans_proc_2").val().replace(",", ".") / 1 + $("#FM_jobans_proc_3").val().replace(",", ".") / 1 + $("#FM_jobans_proc_4").val().replace(",", ".") / 1 + $("#FM_jobans_proc_5").val().replace(",", ".") / 1

        //alert(totalVal / 1)
        if (totalVal / 1 > 100) {
            alert("Den samlede % må ikke overstige 100")
            $("#FM_jobans_proc_" + id).val('0')
        }

    }


    $("#FM_salgsans_proc_1").keyup(function () {


        salgsandelmaksval(1)

    });

    $("#FM_salgsans_proc_2").keyup(function () {

        salgsandelmaksval(2)

    });

    $("#FM_salgsans_proc_3").keyup(function () {

        salgsandelmaksval(3)

    });

    $("#FM_salgsans_proc_4").keyup(function () {

        salgsandelmaksval(4)

    });

    $("#FM_salgsans_proc_5").keyup(function () {

        salgsandelmaksval(5)

    });

    function salgsandelmaksval(id) {


        var totalVal = 0;
        totalVal = $("#FM_salgsans_proc_1").val().replace(",", ".") / 1 + $("#FM_salgsans_proc_2").val().replace(",", ".") / 1 + $("#FM_salgsans_proc_3").val().replace(",", ".") / 1 + $("#FM_salgsans_proc_4").val().replace(",", ".") / 1 + $("#FM_salgsans_proc_5").val().replace(",", ".") / 1

        //alert(totalVal / 1)
        if (totalVal / 1 >= 101) {
            alert("Den samlede % må ikke overstige 100")
            $("#FM_salgsans_proc_" + id).val('0')
        }

    }


    // Aktiviteter
    activitiesCreated = 0

    $("#newActivtyBtn").click(function () {
        //alert("Ny aktivitet")
        activitiesCreated = activitiesCreated + 1
        tdList = "<tr><td><input type='text' class='form-control input-small' name='FM_newactivityName_" + activitiesCreated + "' /></td>"
        tdList = tdList + "<td><select class='form-control input-small' name='FM_newactivityStatus_" + activitiesCreated + "'><option value='1'>Aktiv</option><option value='0'>Lukket</option><option value='2'>Passiv</option></select></td>"
        tdList = tdList + "<td id='typeHolder_" + activitiesCreated + "' style='width:75px;'></td>"
        tdList = tdList + "<td><input type='text' class='form-control input-small' name='FM_newactivitybudgetantal_" + activitiesCreated + "' value='0' /></td>"
        tdList = tdList + "<td><select class='form-control input-small' name='FM_newactivityBGR_" + activitiesCreated + "'><option value='0'>Ingen</option><option value='1'>Timer</option><option value='2'>Stk.</option></select></td>"
        tdList = tdList + "<td><div class='input-group date'><input type='text' style='width:90px;' class='form-control input-small' value='13-06-2018' name='FM_newactivitySTDate_" + activitiesCreated + "' /><span class='input-group-addon input-small'><span class='fa fa-calendar'></span></span></div></td>"
        tdList = tdList + "<td><div class='input-group date'><input type='text' style='width:90px;' class='form-control input-small' value='13-07-2018' name='FM_newactivityENDDate_" + activitiesCreated + "' /><span class='input-group-addon input-small'><span class='fa fa-calendar'></span></span></div></td>"
        tdList = tdList + "<td id='prgSelectHolder_" + activitiesCreated + "' style='width:135px;'></td>"
        tdList = tdList + "<td></td>"
        tdList = tdList + "<td><input type='hidden' value='" + activitiesCreated + "' name='FM_newActivities' /></td>"
        tdList = tdList + "</tr>"

        $("#activtyTable").append(tdList)

        // $("#prgSEL").prop('name', 'FM_newactivityPrgSEL_' + activitiesCreated);
        $("#prgSEL").clone().appendTo('#prgSelectHolder_' + activitiesCreated).prop('name', 'FM_newactivityPrgSEL_' + activitiesCreated);
        $("#akttypeSEL").clone().appendTo('#typeHolder_' + activitiesCreated).prop('name', 'FM_newactivityTypeSEL_' + activitiesCreated);

        $('.date').datepicker({

        });

        //$("#prgSEL").clone().appendTo('#prgSelectHolder_' + activitiesCreated);
        // $("#prgSEL").prop('name', 'EMPTY')

        // alert($("#prgSEL").prop('name').toString())


    });

    $(".edit_akt_btn").click(function () {

        thisid = this.id
        //alert("click")
        $("#FM_akt_navn_" + thisid).prop('disabled', false);
        $("#aktstatus_" + thisid).prop('disabled', false);
        $("#aktSTdate_" + thisid).addClass('date');
        $("#aktSLdate_" + thisid).addClass('date');
        $("#aktInputSTdate_" + thisid).prop('readonly', false);
        $("#aktInputSLdate_" + thisid).prop('readonly', false);
        $("#aktbudget_" + thisid).prop('readonly', false);
        $("#aktEnhed_" + thisid).prop('disabled', false);
        $("#FM_fakturerbart_" + thisid).prop('disabled', false);
        $("#FM_akt_prg_" + thisid).prop('disabled', false);
        $("#FM_akt_prisenhed_" + thisid).prop('disabled', false);


        $('.date').datepicker({

        });

        // for at vide denne er blevet redigeret       
        $("#updated_activities").append("<input type='hidden' value='" + thisid + "' name='FM_updated_activity' />")
    });

    $("#FM_status").change(function () {

        valgtStatus = $("#FM_status").val();

        if (valgtStatus == '3') {

            nxtTnr = $('#FM_nexttnr').val()

            $('#FM_tnr').val(nxtTnr)
            $('.tilbudsinfo').css('visibility', 'inherit');
            $('#FM_usetilbudsnr').prop('checked', true);
        }
        else {
            $('.tilbudsinfo').css('visibility', 'hidden');
            $('#FM_usetilbudsnr').prop('checked', false);
        }

    });

    /* $("#aktView1_btn").click(function () {
 
         $('.aktView2').css('display', 'none');
         $('.aktView3').css('display', 'none');
         $('#aktView3_btn').removeClass('active');
         $('#aktView2_btn').removeClass('active');
         $('#aktView1_btn').addClass('active');
         $('.aktView1').show();
 
     });
 
     $("#aktView2_btn").click(function () {
 
         $('.aktView1').css('display', 'none');
         $('.aktView3').css('display', 'none');
         $('#aktView1_btn').removeClass('active');
         $('#aktView3_btn').removeClass('active');
         $('#aktView2_btn').addClass('active');
         $('.aktView2').show();
 
     });
 
     $("#aktView3_btn").click(function () {
 
         $('.aktView1').css('display', 'none');
         $('.aktView2').css('display', 'none');
         $('#aktView1_btn').removeClass('active');
         $('#aktView2_btn').removeClass('active');
         $('#aktView3_btn').addClass('active');
         $('.aktView3').show();
 
     }); */

    function AktView1() {
        $('#aktView3_btn').removeClass('active');
        $('#aktView2_btn').removeClass('active');
        $('#aktView1_btn').addClass('active');

        $('.akttable').hide();

        $('.akttable_navn').show();
        $('.akttable_status').show();
        $('.akttable_type').show();
        $('.akttable_budget').show();
        $('.akttable_enhed').show();
        $('.akttable_stdato').show();
        $('.akttable_sldato').show();
        $('.akttable_medarb').show();
        $('.akttable_alloker').show();
        $('.akttable_actions').show();
        $('.newActivtyBtn_div').show();
    }

    $("#aktView1_btn").click(function () {

        AktView1();

    });

    $("#aktView2_btn").click(function () {

        $('#aktView1_btn').removeClass('active');
        $('#aktView3_btn').removeClass('active');
        $('#aktView2_btn').addClass('active');

        $('.akttable').hide();

        /* $('#aktheader_navn').show();
         $('#aktheader_status').show();
         $('#aktheader_type').show();
         $('#aktheader_budget').show();
         $('#aktheader_enhed').show();
         $('#aktheader_prisenh').show();
         $('#aktheader_ialt').show();
         $('#aktheader_fctid').show();
         $('#aktheader_fcdkk').show();
         $('#aktheader_realtid').show();
         $('#aktheader_realdkk').show(); */

        $('.akttable_navn').show();
        $('.akttable_status').show();
        $('.akttable_type').show();
        $('.akttable_budget').show();
        $('.akttable_enhed').show();
        $('.akttable_prisenh').show();
        $('.akttable_ialt').show();
        $('.akttable_fctid').show();
        $('.akttable_fcdkk').show();
        $('.akttable_realtid').show();
        $('.akttable_realdkk').show();
        $('.akttable_actions').show();


    });

    $("#aktView3_btn").click(function () {

        $('#aktView1_btn').removeClass('active');
        $('#aktView2_btn').removeClass('active');
        $('#aktView3_btn').addClass('active');

        $('.akttable').hide();

        /*  $('#aktheader_navn').show();
          $('#aktheader_status').show();
          $('#aktheader_type').show();
          $('#aktheader_budget').show();
          $('#aktheader_enhed').show();
          $('#aktheader_prisenh').show();
          $('#aktheader_fc').show();
          $('#aktheader_real').show();
          $('#aktheader_budgetfc').show();
          $('#aktheader_budgetreal').show();
          $('#aktheader_fcreal').show();
          $('#aktheader_faktid').show(); */

        $('.akttable_navn').show();
        // $('.akttable_status').show();
        // $('.akttable_type').show();
        $('.akttable_budget').show();
        $('.akttable_enhed').show();
        $('.akttable_prisenh').show();
        $('.akttable_fc').show();
        $('.akttable_real').show();
        $('.akttable_budgetfc').show();
        $('.akttable_budgetreal').show();
        $('.akttable_fcreal').show();
        $('.akttable_faktid').show();
        $('.akttable_actions').show();

    });


    $("#FM_kunde").change(function () {
        getKpers();
    });

    // kontakpers //
    
    function getKpers() {

        var kundekpersopr_val = $("#kundekpersopr").val()
        var kid_val = $("#FM_kunde").val()

        //alert("her " + kid_val)

        $.post("?jq_kid=" + kid_val + "&jq_kundekpers=" + kundekpersopr_val, { control: "FN_kpers", AjaxUpdateField: "true", cust: 0 }, function (data) {
            //$("#FM_modtageradr").val(data);
            // alert("change")
           // alert(data)
            $("#FM_kpers").html(data);

            //$("#jobid").html(data);
            //alert(data)

        });


    }

    $("#FM_beskrivelse").keyup(function () {
        thisVAL = $("#FM_beskrivelse").val();
        thisVALlng = thisVAL.length;
        $("#beskrivLNG").html("" + thisVALlng + "/1000")

        if (parseInt(thisVALlng) > 1000) {
            $("#beskrivLNG").css("color", "red");
        }

    });

    AktView1();

});