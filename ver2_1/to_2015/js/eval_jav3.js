




$(document).ready(function () {

    OmregnTimer();

    $('.date').datepicker({

    });
    
    function OmregnTimer() {

        eval_timerBrugt = $("#timerBrugt").val()
        eval_timePris = $("#timePris").val()

        var eval_timerBrugt_string = eval_timerBrugt.replace(',', '.')
        var eval_timePris_string = eval_timePris.replace(',', '.')

        var eval_omregnetOms_string = eval_timerBrugt_string * eval_timePris_string
        var evalomregnet = eval_omregnetOms_string.toString().replace('.', ',')


        $("#omregnetOms").val(evalomregnet);

        

        if ($("#CHECK_fastpris").val() == '0') {
            eval_original_price = $("#for_oms_result").val().replace(',', '.')
        }
        if ($("#CHECK_fastpris").val() == '1') {
            eval_original_price = $("#fastprisjob_oms").val().replace(',', '.')
        }

        
        //alert("HER: " + eval_original_price + " - " + eval_omregnetOms_string)

        var eval_lbn_forskel = (eval_original_price - eval_omregnetOms_string)

        
        var evalomregnet_string = eval_lbn_forskel.toString().replace('.', ',')
        $("#eval_lbn_forskel").val(evalomregnet_string);

    }

    $(".opdaterOms").keyup(function () {

        OmregnTimer();
        
    });


    $(".XXeval_sub_btn").click(function () {

            eval_vur = $("input[type='radio'][name='eval_vur']:checked").val()
            eval_jobid = $("#eval_jobid").val()
            eval_comment = $("#eval_comment").val()
            eval_suggested_hours = $("#timerBrugt").val()
            eval_suggested_hourly_rate = $("#timePris").val()
            eval_oms = $("#omregnetOms").val()

            if ($("#CHECK_fastpris").val() == '0') {
                eval_original_price = $("#for_oms_result").val().replace(',', '.')
            }
            if ($("#CHECK_fastpris").val() == '1') {
                eval_original_price = $("#fastprisjob_oms").val().replace(',', '.')
            }

            eval_lbn_forskel = eval_original_price - eval_oms.replace(',', '.')

            $.post("?eval_vur=" + eval_vur + "&eval_jobid=" + eval_jobid + "&eval_suggested=" + eval_oms + "&eval_comment=" + eval_comment + "&timerBrugt=" + eval_suggested_hours + "&timePris=" + eval_suggested_hourly_rate + "&omregnetOms=" + eval_oms + "&eval_lbn_forskel=" + eval_lbn_forskel + "&eval_original_price=" + eval_original_price, { control: "eval_submit_lbn_hours", AjaxUpdateField: "true" }, function (data) {

                alert("Opdateret")
                window.close();
            });


    });




});



