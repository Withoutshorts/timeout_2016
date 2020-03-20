




$(document).ready(function () {
    //alert("hej2")
    OmregnTimer();

    $('.date').datepicker({

    });

    function OmregnTimer() {

        eval_fak_svendetimer = $("#fak_svendetimer").val()
        eval_fak_svendetimePris = $("#fak_svendetimerPris").val()
        var eval_fak_svendetimer_string = eval_fak_svendetimer.replace(/\./g, '');
        eval_fak_svendetimer_string = eval_fak_svendetimer_string.replace(',', '.')

        var eval_fak_svendetimer_pris_string = eval_fak_svendetimePris.replace(/\./g, '');
        eval_fak_svendetimer_pris_string = eval_fak_svendetimer_pris_string.replace(',', '.')

        var svendeTimer_omregnet_string = eval_fak_svendetimer_string * eval_fak_svendetimer_pris_string
        var svendeTimer_omregnet = svendeTimer_omregnet_string.toString().replace('.', ',')
        var svendeTimer_omregnet_deci = parseFloat(svendeTimer_omregnet_string).toFixed(2).replace('.', ',')
        $("#omregnetOms_svendetimer").val(svendeTimer_omregnet_deci);


        ubemandet_maskine_timer = $("#ubemandet_maskine_timer").val()
        ubemandet_maskine_timePris = $("#ubemandet_maskine_timePris").val()

        var eval_ubemandet_maskine_timer_string = ubemandet_maskine_timer.replace(/\./g, '');
        eval_ubemandet_maskine_timer_string = eval_ubemandet_maskine_timer_string.replace(',', '.')

        var eval_ubemandet_maskine_timepris_string = ubemandet_maskine_timePris.replace(/\./g, '');
        eval_ubemandet_maskine_timepris_string = eval_ubemandet_maskine_timepris_string.replace(',', '.')

        var eval_omregnet_ubemandet_maskine_string = eval_ubemandet_maskine_timer_string * eval_ubemandet_maskine_timepris_string
        var eval_omregnet_ubemandet_maskine = eval_omregnet_ubemandet_maskine_string.toString().replace('.', ',')
        var eval_omregnet_ubemandet_maskine_deci = parseFloat(eval_omregnet_ubemandet_maskine_string).toFixed(2).replace('.', ',')
        $("#omregnetOms_ubemandet_maskinetimer").val(eval_omregnet_ubemandet_maskine_deci);


        lear_timer = $("#laer_timer").val()
        lear_timepris = $("#laer_timepris").val()

        var eval_lear_timer_string = lear_timer.replace(/\./g, '');
        eval_lear_timer_string = eval_lear_timer_string.replace(',', '.')

        var eval_lear_timepris_string = lear_timepris.replace(/\./g, '');
        eval_lear_timepris_string = eval_lear_timepris_string.replace(',', '.')
        var eval_lear_timer_omregnet_string = eval_lear_timer_string * eval_lear_timepris_string

        var eval_lear_timer_omrenget = eval_lear_timer_omregnet_string.toString().replace('.', ',')
        var eval_lear_timer_omrenget_deci = parseFloat(eval_lear_timer_omregnet_string).toFixed(2).replace('.', ',')
        $("#omregnet_lear_timer").val(eval_lear_timer_omrenget_deci);


        easy_reg_timer = $("#easy_reg_timer").val()
        easy_reg_timepris = $("#easy_reg_timepris").val()

        eval_easy_reg_timer_string = easy_reg_timer.replace(/\./g, '');
        var eval_easy_reg_timer_string = eval_easy_reg_timer_string.replace(',', '.')

        var eval_easy_reg_timepris_string = easy_reg_timepris.replace(/\./g, '');
        eval_easy_reg_timepris_string = eval_easy_reg_timepris_string.replace(',', '.')

        var omregnet_eval_easy_reg_timepris_string = eval_easy_reg_timer_string * eval_easy_reg_timepris_string
        var omregnet_eval_easy_reg_timepris = omregnet_eval_easy_reg_timepris_string.toString().replace('.', ',')

        var omregnet_eval_easy_reg_timepris_deci = parseFloat(omregnet_eval_easy_reg_timepris_string).toFixed(2).replace('.', ',')
        $("#omregnet_easy_reg_timer").val(omregnet_eval_easy_reg_timepris_deci);

        ikke_fakbar_tid_timer = $("#ikke_fakbar_tid_timer").val()
        ikke_fakbar_tid_timepris = $("#ikke_fakbar_tid_timepris").val()

        var eval_ikke_fakbar_tid_timer_string = ikke_fakbar_tid_timer.replace(/\./g, '');
        eval_ikke_fakbar_tid_timer_string = eval_ikke_fakbar_tid_timer_string.replace(',', '.')

        var eval_ikke_fakbar_tid_timepris_string = ikke_fakbar_tid_timepris.replace(/\./g, '');
        eval_ikke_fakbar_tid_timepris_string = eval_ikke_fakbar_tid_timepris_string.replace(',', '.')

        var omreget_eval_ikke_fakbar_timer_string = eval_ikke_fakbar_tid_timer_string * eval_ikke_fakbar_tid_timepris_string
        var omreget_eval_ikke_fakbar_timer = omreget_eval_ikke_fakbar_timer_string.toString().replace('.', ',')
        var omreget_eval_ikke_fakbar_timer_deci = parseFloat(omreget_eval_ikke_fakbar_timer_string).toFixed(2).replace('.', ',')
        $("#omregnet_ikke_fakbar").val(omreget_eval_ikke_fakbar_timer_deci);

        var total_forslaet_timer_string = parseFloat(eval_fak_svendetimer_string) + parseFloat(eval_ubemandet_maskine_timer_string) + parseFloat(eval_lear_timer_string) + parseFloat(eval_easy_reg_timer_string) + parseFloat(eval_ikke_fakbar_tid_timer_string)
        var total_forslaet_timer = total_forslaet_timer_string.toString().replace('.', ',')
        $("#forslaet_timer_total").val(total_forslaet_timer);

        //var total_forslaet_timepris_string = parseFloat(eval_fak_svendetimer_pris_string) + parseFloat(eval_eval_fak_svendetimePris_string) + parseFloat(eval_lear_timepris_string) + parseFloat(eval_easy_reg_timepris_string) + parseFloat(eval_ikke_fakbar_tid_timepris_string)
        //var total_forslaet_timepris = total_forslaet_timepris_string.toString().replace('.', ',')
        var total_forslaet_timepris = (eval_fak_svendetimer_string * eval_fak_svendetimer_pris_string) + (eval_ubemandet_maskine_timer_string * eval_ubemandet_maskine_timepris_string) + (eval_lear_timer_string * eval_lear_timepris_string) + (eval_easy_reg_timer_string * eval_easy_reg_timepris_string) + (eval_ikke_fakbar_tid_timer_string * eval_ikke_fakbar_tid_timepris_string)
        var total_F_timepris_gns = parseFloat(total_forslaet_timepris) / total_forslaet_timer_string
        var total_F_timepris_gns_deci = parseFloat(total_F_timepris_gns).toFixed(2).replace('.', ',')
        $("#forslaet_timepris_total").val(total_F_timepris_gns_deci);

        var forslaet_oms_string = parseFloat(svendeTimer_omregnet_string) + parseFloat(eval_omregnet_ubemandet_maskine_string) + parseFloat(eval_lear_timer_omregnet_string) + parseFloat(omregnet_eval_easy_reg_timepris_string) + parseFloat(omreget_eval_ikke_fakbar_timer_string)
        var forslaet_oms = forslaet_oms_string.toString().replace('.', ',')
        var forslaet_oms_deci = parseFloat(forslaet_oms_string).toFixed(2).replace('.', ',')
        $("#forslaet_oms").val(forslaet_oms_deci);

        if ($("#CHECK_fastpris").val() == '0') {
            var eval_original_price = $("#for_oms_result").val();            
        }
        if ($("#CHECK_fastpris").val() == '1') {
            var eval_original_price = $("#fastprisjob_oms").val()
        }

        eval_original_price = eval_original_price.replace(/\./g, '');
        eval_original_price = eval_original_price.replace(',', '.')

        var eval_forskel_oms = (eval_original_price - forslaet_oms_string)
        var eval_forskel_oms_string = eval_forskel_oms.toString().replace('.', ',')
        eval_forskel_oms_string = parseFloat(eval_forskel_oms_string).toFixed(2).replace('.', ',')
        $("#eval_lbn_forskel").val(eval_forskel_oms_string);
        //alert(eval_forskel_oms_string)
        //eval_timerBrugt = $("#timerBrugt").val()
        //eval_timePris = $("#timePris").val()
        //var eval_timerBrugt_string = eval_timerBrugt.replace(',', '.')
        //var eval_timePris_string = eval_timePris.replace(',', '.')

        //var eval_omregnetOms_string = eval_timerBrugt_string * eval_timePris_string
        //var evalomregnet = eval_omregnetOms_string.toString().replace('.', ',')


        //$("#omregnetOms").val(evalomregnet);






        //alert("HER: " + eval_original_price + " - " + eval_omregnetOms_string)

        //var eval_lbn_forskel = (eval_original_price - eval_omregnetOms_string)


        //var evalomregnet_string = eval_lbn_forskel.toString().replace('.', ',')
        //$("#eval_lbn_forskel").val(evalomregnet_string);

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



