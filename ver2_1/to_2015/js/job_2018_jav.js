// JScript File
$(window).load(function () {
    // run code


    $("#loadbar").hide(1000);
});






$(document).ready(function () {


    $('.date').datepicker({

    });

    //alert("hepHep")

    activitiesCreated = 0

    $("#newActivtyBtn").click(function () {
        //alert("Ny aktivitet")
        activitiesCreated = activitiesCreated + 1
        tdList = "<tr><td><input type='text' class='form-control input-small' value='"+ activitiesCreated +"' name='FM_newactivityName_" + activitiesCreated +"' /></td>"
        tdList = tdList + "<td><select class='form-control input-small'><option>Aktiv</option><option>Passiv</option><option>Lukket</option></select></td>"
        tdList = tdList + "<td><input type='text' class='form-control input-small' /></td>"
        tdList = tdList + "<td><select class='form-control input-small'><option>Ingen</option><option>Timer</option><option>Stk.</option></select></td>"
        tdList = tdList + "<td><div class='input-group date'><input type='text' style='width:90px;' class='form-control input-small' id='' value='' /><span class='input-group-addon input-small'><span class='fa fa-calendar'></span></span></div></td>"
        tdList = tdList + "<td><div class='input-group date'><input type='text' style='width:90px;' class='form-control input-small' id='' value='' /><span class='input-group-addon input-small'><span class='fa fa-calendar'></span></span></div></td>"
        tdList = tdList + "<td><select class='form-control input-small'><option>Alle medarbejere</option><option>1</option><option>2</option></select></td>"
        tdList = tdList + "<td></td>"
        tdList = tdList + "<td><input type='text' value='" + activitiesCreated +"' name='FM_newActivities' /></td>"
        tdList = tdList + "</tr>"

        $("#activtyTable").append(tdList)

    });
    
    $(".edit_akt_btn").click(function () {

        thisid = this.id
        //alert("click")
        $("#aktstatus_" + thisid).prop('disabled', false);
        $("#aktSTdate_" + thisid).addClass('date'); 
        $("#aktSLdate_" + thisid).addClass('date'); 
        $("#aktInputSTdate_" + thisid).prop('readonly', false);
        $("#aktInputSLdate_" + thisid).prop('readonly', false);
        $("#aktbudget_" + thisid).prop('readonly', false);
        $("#aktEnhed_" + thisid).prop('disabled', false);
  
    });

    $("#FM_status").change(function () {

        valgtStatus = $("#FM_status").val();

        if (valgtStatus == '5') {

            nxtTnr = $('#FM_nexttnr').val()

            $('#FM_tnr').val(nxtTnr)
            $('.tilbudsinfo').css('visibility', 'inherit');
            $('#FM_usetilbudsnr').prop('checked', true);
        }
        else
        {
            $('.tilbudsinfo').css('visibility', 'hidden');
            $('#FM_usetilbudsnr').prop('checked', false);
        }

    });

    $("#aktView1_btn").click(function () {

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
            alert("change")
            $("#FM_kpers").html(data);
            
            //$("#jobid").html(data);
            //alert(data)

        });


    }


});