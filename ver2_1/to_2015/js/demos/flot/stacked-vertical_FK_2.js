$(function () {

    //alert("awd")

    var fomrnavn1 = $("#fomrnavn_0").val();
    var fomrnavn2 = $("#fomrnavn_1").val();
    var fomrnavn3 = $("#fomrnavn_2").val();
    var fomrnavn4 = $("#fomrnavn_3").val();
    var fomrnavn5 = $("#fomrnavn_4").val();
    var fomrnavn6 = $("#fomrnavn_5").val();
    var fomrnavn7 = $("#fomrnavn_6").val();
    var fomrnavn8 = $("#fomrnavn_7").val();
    var fomrnavn9 = $("#fomrnavn_8").val();
    var fomrnavn10 = $("#fomrnavn_9").val();
    var fomrnavn11 = $("#fomrnavn_10").val();
    var fomrnavn12 = $("#fomrnavn_11").val();
    var fomrnavn13 = $("#fomrnavn_12").val();
    var fomrnavn14 = $("#fomrnavn_13").val();
    var fomrnavn15 = $("#fomrnavn_14").val();

    i = 1
    $(".chartContainer").each(function () {
        
        thisid = this.id
        //alert(thisid)
        var medarbnavn = $("#medarbnavn_" + thisid).val().toString();
        var forcastet_val_periode_1_fomr_1 = $("#periode_1_fomr_1_forecast_" + thisid).val(); forcastet_val_periode_1_fomr_1 = forcastet_val_periode_1_fomr_1.replace(",", "."); forcastet_val_periode_1_fomr_1 = parseFloat(forcastet_val_periode_1_fomr_1);
        var forcastet_val_periode_1_fomr_2 = $("#periode_1_fomr_2_forecast_" + thisid).val(); forcastet_val_periode_1_fomr_2 = forcastet_val_periode_1_fomr_2.replace(",", "."); forcastet_val_periode_1_fomr_2 = parseFloat(forcastet_val_periode_1_fomr_2);
        var forcastet_val_periode_1_fomr_3 = $("#periode_1_fomr_3_forecast_" + thisid).val(); forcastet_val_periode_1_fomr_3 = forcastet_val_periode_1_fomr_3.replace(",", "."); forcastet_val_periode_1_fomr_3 = parseFloat(forcastet_val_periode_1_fomr_3);
        var forcastet_val_periode_1_fomr_4 = $("#periode_1_fomr_4_forecast_" + thisid).val(); forcastet_val_periode_1_fomr_4 = forcastet_val_periode_1_fomr_4.replace(",", "."); forcastet_val_periode_1_fomr_4 = parseFloat(forcastet_val_periode_1_fomr_4);
        var forcastet_val_periode_1_fomr_5 = $("#periode_1_fomr_5_forecast_" + thisid).val(); forcastet_val_periode_1_fomr_5 = forcastet_val_periode_1_fomr_5.replace(",", "."); forcastet_val_periode_1_fomr_5 = parseFloat(forcastet_val_periode_1_fomr_5);
        var forcastet_val_periode_1_fomr_6 = $("#periode_1_fomr_6_forecast_" + thisid).val(); forcastet_val_periode_1_fomr_6 = forcastet_val_periode_1_fomr_6.replace(",", "."); forcastet_val_periode_1_fomr_6 = parseFloat(forcastet_val_periode_1_fomr_6);
        var forcastet_val_periode_1_fomr_7 = $("#periode_1_fomr_7_forecast_" + thisid).val(); forcastet_val_periode_1_fomr_7 = forcastet_val_periode_1_fomr_7.replace(",", "."); forcastet_val_periode_1_fomr_7 = parseFloat(forcastet_val_periode_1_fomr_7);
        var forcastet_val_periode_1_fomr_8 = $("#periode_1_fomr_8_forecast_" + thisid).val(); forcastet_val_periode_1_fomr_8 = forcastet_val_periode_1_fomr_8.replace(",", "."); forcastet_val_periode_1_fomr_8 = parseFloat(forcastet_val_periode_1_fomr_8);
        var forcastet_val_periode_1_fomr_9 = $("#periode_1_fomr_9_forecast_" + thisid).val(); forcastet_val_periode_1_fomr_9 = forcastet_val_periode_1_fomr_9.replace(",", "."); forcastet_val_periode_1_fomr_9 = parseFloat(forcastet_val_periode_1_fomr_9);
        var forcastet_val_periode_1_fomr_10 = $("#periode_1_fomr_10_forecast_" + thisid).val(); forcastet_val_periode_1_fomr_10 = forcastet_val_periode_1_fomr_10.replace(",", "."); forcastet_val_periode_1_fomr_10 = parseFloat(forcastet_val_periode_1_fomr_10);
        var forcastet_val_periode_1_fomr_11 = $("#periode_1_fomr_11_forecast_" + thisid).val(); forcastet_val_periode_1_fomr_11 = forcastet_val_periode_1_fomr_11.replace(",", "."); forcastet_val_periode_1_fomr_11 = parseFloat(forcastet_val_periode_1_fomr_11);
        var forcastet_val_periode_1_fomr_12 = $("#periode_1_fomr_12_forecast_" + thisid).val(); forcastet_val_periode_1_fomr_12 = forcastet_val_periode_1_fomr_12.replace(",", "."); forcastet_val_periode_1_fomr_12 = parseFloat(forcastet_val_periode_1_fomr_12);
        var forcastet_val_periode_1_fomr_13 = $("#periode_1_fomr_13_forecast_" + thisid).val(); forcastet_val_periode_1_fomr_13 = forcastet_val_periode_1_fomr_13.replace(",", "."); forcastet_val_periode_1_fomr_13 = parseFloat(forcastet_val_periode_1_fomr_13);
        var forcastet_val_periode_1_fomr_14 = $("#periode_1_fomr_14_forecast_" + thisid).val(); forcastet_val_periode_1_fomr_14 = forcastet_val_periode_1_fomr_14.replace(",", "."); forcastet_val_periode_1_fomr_14 = parseFloat(forcastet_val_periode_1_fomr_14);
        var forcastet_val_periode_1_fomr_15 = $("#periode_1_fomr_15_forecast_" + thisid).val(); forcastet_val_periode_1_fomr_15 = forcastet_val_periode_1_fomr_15.replace(",", "."); forcastet_val_periode_1_fomr_15 = parseFloat(forcastet_val_periode_1_fomr_15);
        
        var forcastet_val_periode_2_fomr_1 = $("#periode_2_fomr_1_forecast_" + thisid).val(); forcastet_val_periode_2_fomr_1 = forcastet_val_periode_2_fomr_1.replace(",", "."); forcastet_val_periode_2_fomr_1 = parseFloat(forcastet_val_periode_2_fomr_1);
        var forcastet_val_periode_2_fomr_2 = $("#periode_2_fomr_2_forecast_" + thisid).val(); forcastet_val_periode_2_fomr_2 = forcastet_val_periode_2_fomr_2.replace(",", "."); forcastet_val_periode_2_fomr_2 = parseFloat(forcastet_val_periode_2_fomr_2);
        var forcastet_val_periode_2_fomr_3 = $("#periode_2_fomr_3_forecast_" + thisid).val(); forcastet_val_periode_2_fomr_3 = forcastet_val_periode_2_fomr_3.replace(",", "."); forcastet_val_periode_2_fomr_3 = parseFloat(forcastet_val_periode_2_fomr_3);
        var forcastet_val_periode_2_fomr_4 = $("#periode_2_fomr_4_forecast_" + thisid).val(); forcastet_val_periode_2_fomr_4 = forcastet_val_periode_2_fomr_4.replace(",", "."); forcastet_val_periode_2_fomr_4 = parseFloat(forcastet_val_periode_2_fomr_4);
        var forcastet_val_periode_2_fomr_5 = $("#periode_2_fomr_5_forecast_" + thisid).val(); forcastet_val_periode_2_fomr_5 = forcastet_val_periode_2_fomr_5.replace(",", "."); forcastet_val_periode_2_fomr_5 = parseFloat(forcastet_val_periode_2_fomr_5);
        var forcastet_val_periode_2_fomr_6 = $("#periode_2_fomr_6_forecast_" + thisid).val(); forcastet_val_periode_2_fomr_6 = forcastet_val_periode_2_fomr_6.replace(",", "."); forcastet_val_periode_2_fomr_6 = parseFloat(forcastet_val_periode_2_fomr_6);
        var forcastet_val_periode_2_fomr_7 = $("#periode_2_fomr_7_forecast_" + thisid).val(); forcastet_val_periode_2_fomr_7 = forcastet_val_periode_2_fomr_7.replace(",", "."); forcastet_val_periode_2_fomr_7 = parseFloat(forcastet_val_periode_2_fomr_7);
        var forcastet_val_periode_2_fomr_8 = $("#periode_2_fomr_8_forecast_" + thisid).val(); forcastet_val_periode_2_fomr_8 = forcastet_val_periode_2_fomr_8.replace(",", "."); forcastet_val_periode_2_fomr_8 = parseFloat(forcastet_val_periode_2_fomr_8);
        var forcastet_val_periode_2_fomr_9 = $("#periode_2_fomr_9_forecast_" + thisid).val(); forcastet_val_periode_2_fomr_9 = forcastet_val_periode_2_fomr_9.replace(",", "."); forcastet_val_periode_2_fomr_9 = parseFloat(forcastet_val_periode_2_fomr_9);
        var forcastet_val_periode_2_fomr_10 = $("#periode_2_fomr_10_forecast_" + thisid).val(); forcastet_val_periode_2_fomr_10 = forcastet_val_periode_2_fomr_10.replace(",", "."); forcastet_val_periode_2_fomr_10 = parseFloat(forcastet_val_periode_2_fomr_10);
        var forcastet_val_periode_2_fomr_11 = $("#periode_2_fomr_11_forecast_" + thisid).val(); forcastet_val_periode_2_fomr_11 = forcastet_val_periode_2_fomr_11.replace(",", "."); forcastet_val_periode_2_fomr_11 = parseFloat(forcastet_val_periode_2_fomr_11);
        var forcastet_val_periode_2_fomr_12 = $("#periode_2_fomr_12_forecast_" + thisid).val(); forcastet_val_periode_2_fomr_12 = forcastet_val_periode_2_fomr_12.replace(",", "."); forcastet_val_periode_2_fomr_12 = parseFloat(forcastet_val_periode_2_fomr_12);
        var forcastet_val_periode_2_fomr_13 = $("#periode_2_fomr_13_forecast_" + thisid).val(); forcastet_val_periode_2_fomr_13 = forcastet_val_periode_2_fomr_13.replace(",", "."); forcastet_val_periode_2_fomr_13 = parseFloat(forcastet_val_periode_2_fomr_13);
        var forcastet_val_periode_2_fomr_14 = $("#periode_2_fomr_14_forecast_" + thisid).val(); forcastet_val_periode_2_fomr_14 = forcastet_val_periode_2_fomr_14.replace(",", "."); forcastet_val_periode_2_fomr_14 = parseFloat(forcastet_val_periode_2_fomr_14);
        var forcastet_val_periode_2_fomr_15 = $("#periode_2_fomr_15_forecast_" + thisid).val(); forcastet_val_periode_2_fomr_15 = forcastet_val_periode_2_fomr_15.replace(",", "."); forcastet_val_periode_2_fomr_15 = parseFloat(forcastet_val_periode_2_fomr_15);

        var forcastet_val_periode_3_fomr_1 = $("#periode_3_fomr_1_forecast_" + thisid).val(); forcastet_val_periode_3_fomr_1 = forcastet_val_periode_3_fomr_1.replace(",", "."); forcastet_val_periode_3_fomr_1 = parseFloat(forcastet_val_periode_3_fomr_1);
        var forcastet_val_periode_3_fomr_2 = $("#periode_3_fomr_2_forecast_" + thisid).val(); forcastet_val_periode_3_fomr_2 = forcastet_val_periode_3_fomr_2.replace(",", "."); forcastet_val_periode_3_fomr_2 = parseFloat(forcastet_val_periode_3_fomr_2);
        var forcastet_val_periode_3_fomr_3 = $("#periode_3_fomr_3_forecast_" + thisid).val(); forcastet_val_periode_3_fomr_3 = forcastet_val_periode_3_fomr_3.replace(",", "."); forcastet_val_periode_3_fomr_3 = parseFloat(forcastet_val_periode_3_fomr_3);
        var forcastet_val_periode_3_fomr_4 = $("#periode_3_fomr_4_forecast_" + thisid).val(); forcastet_val_periode_3_fomr_4 = forcastet_val_periode_3_fomr_4.replace(",", "."); forcastet_val_periode_3_fomr_4 = parseFloat(forcastet_val_periode_3_fomr_4);
        var forcastet_val_periode_3_fomr_5 = $("#periode_3_fomr_5_forecast_" + thisid).val(); forcastet_val_periode_3_fomr_5 = forcastet_val_periode_3_fomr_5.replace(",", "."); forcastet_val_periode_3_fomr_5 = parseFloat(forcastet_val_periode_3_fomr_5);
        var forcastet_val_periode_3_fomr_6 = $("#periode_3_fomr_6_forecast_" + thisid).val(); forcastet_val_periode_3_fomr_6 = forcastet_val_periode_3_fomr_6.replace(",", "."); forcastet_val_periode_3_fomr_6 = parseFloat(forcastet_val_periode_3_fomr_6);
        var forcastet_val_periode_3_fomr_7 = $("#periode_3_fomr_7_forecast_" + thisid).val(); forcastet_val_periode_3_fomr_7 = forcastet_val_periode_3_fomr_7.replace(",", "."); forcastet_val_periode_3_fomr_7 = parseFloat(forcastet_val_periode_3_fomr_7);
        var forcastet_val_periode_3_fomr_8 = $("#periode_3_fomr_8_forecast_" + thisid).val(); forcastet_val_periode_3_fomr_8 = forcastet_val_periode_3_fomr_8.replace(",", "."); forcastet_val_periode_3_fomr_8 = parseFloat(forcastet_val_periode_3_fomr_8);
        var forcastet_val_periode_3_fomr_9 = $("#periode_3_fomr_9_forecast_" + thisid).val(); forcastet_val_periode_3_fomr_9 = forcastet_val_periode_3_fomr_9.replace(",", "."); forcastet_val_periode_3_fomr_9 = parseFloat(forcastet_val_periode_3_fomr_9);
        var forcastet_val_periode_3_fomr_10 = $("#periode_3_fomr_10_forecast_" + thisid).val(); forcastet_val_periode_3_fomr_10 = forcastet_val_periode_3_fomr_10.replace(",", "."); forcastet_val_periode_3_fomr_10 = parseFloat(forcastet_val_periode_3_fomr_10);
        var forcastet_val_periode_3_fomr_11 = $("#periode_3_fomr_11_forecast_" + thisid).val(); forcastet_val_periode_3_fomr_11 = forcastet_val_periode_3_fomr_11.replace(",", "."); forcastet_val_periode_3_fomr_11 = parseFloat(forcastet_val_periode_3_fomr_11);
        var forcastet_val_periode_3_fomr_12 = $("#periode_3_fomr_12_forecast_" + thisid).val(); forcastet_val_periode_3_fomr_12 = forcastet_val_periode_3_fomr_12.replace(",", "."); forcastet_val_periode_3_fomr_12 = parseFloat(forcastet_val_periode_3_fomr_12);
        var forcastet_val_periode_3_fomr_13 = $("#periode_3_fomr_13_forecast_" + thisid).val(); forcastet_val_periode_3_fomr_13 = forcastet_val_periode_3_fomr_13.replace(",", "."); forcastet_val_periode_3_fomr_13 = parseFloat(forcastet_val_periode_3_fomr_13);
        var forcastet_val_periode_3_fomr_14 = $("#periode_3_fomr_14_forecast_" + thisid).val(); forcastet_val_periode_3_fomr_14 = forcastet_val_periode_3_fomr_14.replace(",", "."); forcastet_val_periode_3_fomr_14 = parseFloat(forcastet_val_periode_3_fomr_14);
        var forcastet_val_periode_3_fomr_15 = $("#periode_3_fomr_15_forecast_" + thisid).val(); forcastet_val_periode_3_fomr_15 = forcastet_val_periode_3_fomr_15.replace(",", "."); forcastet_val_periode_3_fomr_15 = parseFloat(forcastet_val_periode_3_fomr_15);
        
        var forcastet_val_periode_4_fomr_1 = $("#periode_4_fomr_1_forecast_" + thisid).val(); forcastet_val_periode_4_fomr_1 = forcastet_val_periode_4_fomr_1.replace(",", "."); forcastet_val_periode_4_fomr_1 = parseFloat(forcastet_val_periode_4_fomr_1);
        var forcastet_val_periode_4_fomr_2 = $("#periode_4_fomr_2_forecast_" + thisid).val(); forcastet_val_periode_4_fomr_2 = forcastet_val_periode_4_fomr_2.replace(",", "."); forcastet_val_periode_4_fomr_2 = parseFloat(forcastet_val_periode_4_fomr_2);
        var forcastet_val_periode_4_fomr_3 = $("#periode_4_fomr_3_forecast_" + thisid).val(); forcastet_val_periode_4_fomr_3 = forcastet_val_periode_4_fomr_3.replace(",", "."); forcastet_val_periode_4_fomr_3 = parseFloat(forcastet_val_periode_4_fomr_3);
        var forcastet_val_periode_4_fomr_4 = $("#periode_4_fomr_4_forecast_" + thisid).val(); forcastet_val_periode_4_fomr_4 = forcastet_val_periode_4_fomr_4.replace(",", "."); forcastet_val_periode_4_fomr_4 = parseFloat(forcastet_val_periode_4_fomr_4);
        var forcastet_val_periode_4_fomr_5 = $("#periode_4_fomr_5_forecast_" + thisid).val(); forcastet_val_periode_4_fomr_5 = forcastet_val_periode_4_fomr_5.replace(",", "."); forcastet_val_periode_4_fomr_5 = parseFloat(forcastet_val_periode_4_fomr_5);
        var forcastet_val_periode_4_fomr_6 = $("#periode_4_fomr_6_forecast_" + thisid).val(); forcastet_val_periode_4_fomr_6 = forcastet_val_periode_4_fomr_6.replace(",", "."); forcastet_val_periode_4_fomr_6 = parseFloat(forcastet_val_periode_4_fomr_6);
        var forcastet_val_periode_4_fomr_7 = $("#periode_4_fomr_7_forecast_" + thisid).val(); forcastet_val_periode_4_fomr_7 = forcastet_val_periode_4_fomr_7.replace(",", "."); forcastet_val_periode_4_fomr_7 = parseFloat(forcastet_val_periode_4_fomr_7);
        var forcastet_val_periode_4_fomr_8 = $("#periode_4_fomr_8_forecast_" + thisid).val(); forcastet_val_periode_4_fomr_8 = forcastet_val_periode_4_fomr_8.replace(",", "."); forcastet_val_periode_4_fomr_8 = parseFloat(forcastet_val_periode_4_fomr_8);
        var forcastet_val_periode_4_fomr_9 = $("#periode_4_fomr_9_forecast_" + thisid).val(); forcastet_val_periode_4_fomr_9 = forcastet_val_periode_4_fomr_9.replace(",", "."); forcastet_val_periode_4_fomr_9 = parseFloat(forcastet_val_periode_4_fomr_9);
        var forcastet_val_periode_4_fomr_10 = $("#periode_4_fomr_10_forecast_" + thisid).val(); forcastet_val_periode_4_fomr_10 = forcastet_val_periode_4_fomr_10.replace(",", "."); forcastet_val_periode_4_fomr_10 = parseFloat(forcastet_val_periode_4_fomr_10);
        var forcastet_val_periode_4_fomr_11 = $("#periode_4_fomr_11_forecast_" + thisid).val(); forcastet_val_periode_4_fomr_11 = forcastet_val_periode_4_fomr_11.replace(",", "."); forcastet_val_periode_4_fomr_11 = parseFloat(forcastet_val_periode_4_fomr_11);
        var forcastet_val_periode_4_fomr_12 = $("#periode_4_fomr_12_forecast_" + thisid).val(); forcastet_val_periode_4_fomr_12 = forcastet_val_periode_4_fomr_12.replace(",", "."); forcastet_val_periode_4_fomr_12 = parseFloat(forcastet_val_periode_4_fomr_12);
        var forcastet_val_periode_4_fomr_13 = $("#periode_4_fomr_13_forecast_" + thisid).val(); forcastet_val_periode_4_fomr_13 = forcastet_val_periode_4_fomr_13.replace(",", "."); forcastet_val_periode_4_fomr_13 = parseFloat(forcastet_val_periode_4_fomr_13);
        var forcastet_val_periode_4_fomr_14 = $("#periode_4_fomr_14_forecast_" + thisid).val(); forcastet_val_periode_4_fomr_14 = forcastet_val_periode_4_fomr_14.replace(",", "."); forcastet_val_periode_4_fomr_14 = parseFloat(forcastet_val_periode_4_fomr_14);
        var forcastet_val_periode_4_fomr_15 = $("#periode_4_fomr_15_forecast_" + thisid).val(); forcastet_val_periode_4_fomr_15 = forcastet_val_periode_4_fomr_15.replace(",", "."); forcastet_val_periode_4_fomr_15 = parseFloat(forcastet_val_periode_4_fomr_15);
        


        var registret_val_periode_1_fomr_1 = $("#periode_1_fomr_1_registreret_" + thisid).val(); registret_val_periode_1_fomr_1 = registret_val_periode_1_fomr_1.replace(",", "."); registret_val_periode_1_fomr_1 = parseFloat(registret_val_periode_1_fomr_1);
        var registret_val_periode_1_fomr_2 = $("#periode_1_fomr_2_registreret_" + thisid).val(); registret_val_periode_1_fomr_2 = registret_val_periode_1_fomr_2.replace(",", "."); registret_val_periode_1_fomr_2 = parseFloat(registret_val_periode_1_fomr_2);
        var registret_val_periode_1_fomr_3 = $("#periode_1_fomr_3_registreret_" + thisid).val(); registret_val_periode_1_fomr_3 = registret_val_periode_1_fomr_3.replace(",", "."); registret_val_periode_1_fomr_3 = parseFloat(registret_val_periode_1_fomr_3);
        var registret_val_periode_1_fomr_4 = $("#periode_1_fomr_4_registreret_" + thisid).val(); registret_val_periode_1_fomr_4 = registret_val_periode_1_fomr_4.replace(",", "."); registret_val_periode_1_fomr_4 = parseFloat(registret_val_periode_1_fomr_4);
        var registret_val_periode_1_fomr_5 = $("#periode_1_fomr_5_registreret_" + thisid).val(); registret_val_periode_1_fomr_5 = registret_val_periode_1_fomr_5.replace(",", "."); registret_val_periode_1_fomr_5 = parseFloat(registret_val_periode_1_fomr_5);
        var registret_val_periode_1_fomr_6 = $("#periode_1_fomr_6_registreret_" + thisid).val(); registret_val_periode_1_fomr_6 = registret_val_periode_1_fomr_6.replace(",", "."); registret_val_periode_1_fomr_6 = parseFloat(registret_val_periode_1_fomr_6);
        var registret_val_periode_1_fomr_7 = $("#periode_1_fomr_7_registreret_" + thisid).val(); registret_val_periode_1_fomr_7 = registret_val_periode_1_fomr_7.replace(",", "."); registret_val_periode_1_fomr_7 = parseFloat(registret_val_periode_1_fomr_7);
        var registret_val_periode_1_fomr_8 = $("#periode_1_fomr_8_registreret_" + thisid).val(); registret_val_periode_1_fomr_8 = registret_val_periode_1_fomr_8.replace(",", "."); registret_val_periode_1_fomr_8 = parseFloat(registret_val_periode_1_fomr_8);
        var registret_val_periode_1_fomr_9 = $("#periode_1_fomr_9_registreret_" + thisid).val(); registret_val_periode_1_fomr_9 = registret_val_periode_1_fomr_9.replace(",", "."); registret_val_periode_1_fomr_9 = parseFloat(registret_val_periode_1_fomr_9);
        var registret_val_periode_1_fomr_10 = $("#periode_1_fomr_10_registreret_" + thisid).val(); registret_val_periode_1_fomr_10 = registret_val_periode_1_fomr_10.replace(",", "."); registret_val_periode_1_fomr_10 = parseFloat(registret_val_periode_1_fomr_10);
        var registret_val_periode_1_fomr_11 = $("#periode_1_fomr_11_registreret_" + thisid).val(); registret_val_periode_1_fomr_11 = registret_val_periode_1_fomr_11.replace(",", "."); registret_val_periode_1_fomr_11 = parseFloat(registret_val_periode_1_fomr_11);
        var registret_val_periode_1_fomr_12 = $("#periode_1_fomr_12_registreret_" + thisid).val(); registret_val_periode_1_fomr_12 = registret_val_periode_1_fomr_12.replace(",", "."); registret_val_periode_1_fomr_12 = parseFloat(registret_val_periode_1_fomr_12);
        var registret_val_periode_1_fomr_13 = $("#periode_1_fomr_13_registreret_" + thisid).val(); registret_val_periode_1_fomr_13 = registret_val_periode_1_fomr_13.replace(",", "."); registret_val_periode_1_fomr_13 = parseFloat(registret_val_periode_1_fomr_13);
        var registret_val_periode_1_fomr_14 = $("#periode_1_fomr_14_registreret_" + thisid).val(); registret_val_periode_1_fomr_14 = registret_val_periode_1_fomr_14.replace(",", "."); registret_val_periode_1_fomr_14 = parseFloat(registret_val_periode_1_fomr_14);
        var registret_val_periode_1_fomr_15 = $("#periode_1_fomr_15_registreret_" + thisid).val(); registret_val_periode_1_fomr_15 = registret_val_periode_1_fomr_15.replace(",", "."); registret_val_periode_1_fomr_15 = parseFloat(registret_val_periode_1_fomr_15);

      /*  alert(registret_val_periode_1_fomr_1)
        alert(registret_val_periode_1_fomr_2)
        alert(registret_val_periode_1_fomr_3)
        alert(registret_val_periode_1_fomr_4)
        alert(registret_val_periode_1_fomr_5)
        alert(registret_val_periode_1_fomr_6)
        alert(registret_val_periode_1_fomr_7)
        alert(registret_val_periode_1_fomr_8)
        alert(registret_val_periode_1_fomr_9)
        alert(registret_val_periode_1_fomr_10)
        alert(registret_val_periode_1_fomr_11)
        alert(registret_val_periode_1_fomr_12)
        alert(registret_val_periode_1_fomr_13)
        alert(registret_val_periode_1_fomr_14)
        alert(registret_val_periode_1_fomr_15) */


        var registret_val_periode_2_fomr_1 = $("#periode_2_fomr_1_registreret_" + thisid).val(); registret_val_periode_2_fomr_1 = registret_val_periode_2_fomr_1.replace(",", "."); registret_val_periode_2_fomr_1 = parseFloat(registret_val_periode_2_fomr_1);
        var registret_val_periode_2_fomr_2 = $("#periode_2_fomr_2_registreret_" + thisid).val(); registret_val_periode_2_fomr_2 = registret_val_periode_2_fomr_2.replace(",", "."); registret_val_periode_2_fomr_2 = parseFloat(registret_val_periode_2_fomr_2);
        var registret_val_periode_2_fomr_3 = $("#periode_2_fomr_3_registreret_" + thisid).val(); registret_val_periode_2_fomr_3 = registret_val_periode_2_fomr_3.replace(",", "."); registret_val_periode_2_fomr_3 = parseFloat(registret_val_periode_2_fomr_3);
        var registret_val_periode_2_fomr_4 = $("#periode_2_fomr_4_registreret_" + thisid).val(); registret_val_periode_2_fomr_4 = registret_val_periode_2_fomr_4.replace(",", "."); registret_val_periode_2_fomr_4 = parseFloat(registret_val_periode_2_fomr_4);
        var registret_val_periode_2_fomr_5 = $("#periode_2_fomr_5_registreret_" + thisid).val(); registret_val_periode_2_fomr_5 = registret_val_periode_2_fomr_5.replace(",", "."); registret_val_periode_2_fomr_5 = parseFloat(registret_val_periode_2_fomr_5);
        var registret_val_periode_2_fomr_6 = $("#periode_2_fomr_6_registreret_" + thisid).val(); registret_val_periode_2_fomr_6 = registret_val_periode_2_fomr_6.replace(",", "."); registret_val_periode_2_fomr_6 = parseFloat(registret_val_periode_2_fomr_6);
        var registret_val_periode_2_fomr_7 = $("#periode_2_fomr_7_registreret_" + thisid).val(); registret_val_periode_2_fomr_7 = registret_val_periode_2_fomr_7.replace(",", "."); registret_val_periode_2_fomr_7 = parseFloat(registret_val_periode_2_fomr_7);
        var registret_val_periode_2_fomr_8 = $("#periode_2_fomr_8_registreret_" + thisid).val(); registret_val_periode_2_fomr_8 = registret_val_periode_2_fomr_8.replace(",", "."); registret_val_periode_2_fomr_8 = parseFloat(registret_val_periode_2_fomr_8);
        var registret_val_periode_2_fomr_9 = $("#periode_2_fomr_9_registreret_" + thisid).val(); registret_val_periode_2_fomr_9 = registret_val_periode_2_fomr_9.replace(",", "."); registret_val_periode_2_fomr_9 = parseFloat(registret_val_periode_2_fomr_9);
        var registret_val_periode_2_fomr_10 = $("#periode_2_fomr_10_registreret_" + thisid).val(); registret_val_periode_2_fomr_10 = registret_val_periode_2_fomr_10.replace(",", "."); registret_val_periode_2_fomr_10 = parseFloat(registret_val_periode_2_fomr_10);
        var registret_val_periode_2_fomr_11 = $("#periode_2_fomr_11_registreret_" + thisid).val(); registret_val_periode_2_fomr_11 = registret_val_periode_2_fomr_11.replace(",", "."); registret_val_periode_2_fomr_11 = parseFloat(registret_val_periode_2_fomr_11);
        var registret_val_periode_2_fomr_12 = $("#periode_2_fomr_12_registreret_" + thisid).val(); registret_val_periode_2_fomr_12 = registret_val_periode_2_fomr_12.replace(",", "."); registret_val_periode_2_fomr_12 = parseFloat(registret_val_periode_2_fomr_12);
        var registret_val_periode_2_fomr_13 = $("#periode_2_fomr_13_registreret_" + thisid).val(); registret_val_periode_2_fomr_13 = registret_val_periode_2_fomr_13.replace(",", "."); registret_val_periode_2_fomr_13 = parseFloat(registret_val_periode_2_fomr_13);
        var registret_val_periode_2_fomr_14 = $("#periode_2_fomr_14_registreret_" + thisid).val(); registret_val_periode_2_fomr_14 = registret_val_periode_2_fomr_14.replace(",", "."); registret_val_periode_2_fomr_14 = parseFloat(registret_val_periode_2_fomr_14);
        var registret_val_periode_2_fomr_15 = $("#periode_2_fomr_15_registreret_" + thisid).val(); registret_val_periode_2_fomr_15 = registret_val_periode_2_fomr_15.replace(",", "."); registret_val_periode_2_fomr_15 = parseFloat(registret_val_periode_2_fomr_15);

        var registret_val_periode_3_fomr_1 = $("#periode_3_fomr_1_registreret_" + thisid).val(); registret_val_periode_3_fomr_1 = registret_val_periode_3_fomr_1.replace(",", "."); registret_val_periode_3_fomr_1 = parseFloat(registret_val_periode_3_fomr_1);
        var registret_val_periode_3_fomr_2 = $("#periode_3_fomr_2_registreret_" + thisid).val(); registret_val_periode_3_fomr_2 = registret_val_periode_3_fomr_2.replace(",", "."); registret_val_periode_3_fomr_2 = parseFloat(registret_val_periode_3_fomr_2);
        var registret_val_periode_3_fomr_3 = $("#periode_3_fomr_3_registreret_" + thisid).val(); registret_val_periode_3_fomr_3 = registret_val_periode_3_fomr_3.replace(",", "."); registret_val_periode_3_fomr_3 = parseFloat(registret_val_periode_3_fomr_3);
        var registret_val_periode_3_fomr_4 = $("#periode_3_fomr_4_registreret_" + thisid).val(); registret_val_periode_3_fomr_4 = registret_val_periode_3_fomr_4.replace(",", "."); registret_val_periode_3_fomr_4 = parseFloat(registret_val_periode_3_fomr_4);
        var registret_val_periode_3_fomr_5 = $("#periode_3_fomr_5_registreret_" + thisid).val(); registret_val_periode_3_fomr_5 = registret_val_periode_3_fomr_5.replace(",", "."); registret_val_periode_3_fomr_5 = parseFloat(registret_val_periode_3_fomr_5);
        var registret_val_periode_3_fomr_6 = $("#periode_3_fomr_6_registreret_" + thisid).val(); registret_val_periode_3_fomr_6 = registret_val_periode_3_fomr_6.replace(",", "."); registret_val_periode_3_fomr_6 = parseFloat(registret_val_periode_3_fomr_6);
        var registret_val_periode_3_fomr_7 = $("#periode_3_fomr_7_registreret_" + thisid).val(); registret_val_periode_3_fomr_7 = registret_val_periode_3_fomr_7.replace(",", "."); registret_val_periode_3_fomr_7 = parseFloat(registret_val_periode_3_fomr_7);
        var registret_val_periode_3_fomr_8 = $("#periode_3_fomr_8_registreret_" + thisid).val(); registret_val_periode_3_fomr_8 = registret_val_periode_3_fomr_8.replace(",", "."); registret_val_periode_3_fomr_8 = parseFloat(registret_val_periode_3_fomr_8);
        var registret_val_periode_3_fomr_9 = $("#periode_3_fomr_9_registreret_" + thisid).val(); registret_val_periode_3_fomr_9 = registret_val_periode_3_fomr_9.replace(",", "."); registret_val_periode_3_fomr_9 = parseFloat(registret_val_periode_3_fomr_9);
        var registret_val_periode_3_fomr_10 = $("#periode_3_fomr_10_registreret_" + thisid).val(); registret_val_periode_3_fomr_10 = registret_val_periode_3_fomr_10.replace(",", "."); registret_val_periode_3_fomr_10 = parseFloat(registret_val_periode_3_fomr_10);
        var registret_val_periode_3_fomr_11 = $("#periode_3_fomr_11_registreret_" + thisid).val(); registret_val_periode_3_fomr_11 = registret_val_periode_3_fomr_11.replace(",", "."); registret_val_periode_3_fomr_11 = parseFloat(registret_val_periode_3_fomr_11);
        var registret_val_periode_3_fomr_12 = $("#periode_3_fomr_12_registreret_" + thisid).val(); registret_val_periode_3_fomr_12 = registret_val_periode_3_fomr_12.replace(",", "."); registret_val_periode_3_fomr_12 = parseFloat(registret_val_periode_3_fomr_12);
        var registret_val_periode_3_fomr_13 = $("#periode_3_fomr_13_registreret_" + thisid).val(); registret_val_periode_3_fomr_13 = registret_val_periode_3_fomr_13.replace(",", "."); registret_val_periode_3_fomr_13 = parseFloat(registret_val_periode_3_fomr_13);
        var registret_val_periode_3_fomr_14 = $("#periode_3_fomr_14_registreret_" + thisid).val(); registret_val_periode_3_fomr_14 = registret_val_periode_3_fomr_14.replace(",", "."); registret_val_periode_3_fomr_14 = parseFloat(registret_val_periode_3_fomr_14);
        var registret_val_periode_3_fomr_15 = $("#periode_3_fomr_15_registreret_" + thisid).val(); registret_val_periode_3_fomr_15 = registret_val_periode_3_fomr_15.replace(",", "."); registret_val_periode_3_fomr_15 = parseFloat(registret_val_periode_3_fomr_15);

        var registret_val_periode_4_fomr_1 = $("#periode_4_fomr_1_registreret_" + thisid).val(); registret_val_periode_4_fomr_1 = registret_val_periode_4_fomr_1.replace(",", "."); registret_val_periode_4_fomr_1 = parseFloat(registret_val_periode_4_fomr_1);
        var registret_val_periode_4_fomr_2 = $("#periode_4_fomr_2_registreret_" + thisid).val(); registret_val_periode_4_fomr_2 = registret_val_periode_4_fomr_2.replace(",", "."); registret_val_periode_4_fomr_2 = parseFloat(registret_val_periode_4_fomr_2);
        var registret_val_periode_4_fomr_3 = $("#periode_4_fomr_3_registreret_" + thisid).val(); registret_val_periode_4_fomr_3 = registret_val_periode_4_fomr_3.replace(",", "."); registret_val_periode_4_fomr_3 = parseFloat(registret_val_periode_4_fomr_3);
        var registret_val_periode_4_fomr_4 = $("#periode_4_fomr_4_registreret_" + thisid).val(); registret_val_periode_4_fomr_4 = registret_val_periode_4_fomr_4.replace(",", "."); registret_val_periode_4_fomr_4 = parseFloat(registret_val_periode_4_fomr_4);
        var registret_val_periode_4_fomr_5 = $("#periode_4_fomr_5_registreret_" + thisid).val(); registret_val_periode_4_fomr_5 = registret_val_periode_4_fomr_5.replace(",", "."); registret_val_periode_4_fomr_5 = parseFloat(registret_val_periode_4_fomr_5);
        var registret_val_periode_4_fomr_6 = $("#periode_4_fomr_6_registreret_" + thisid).val(); registret_val_periode_4_fomr_6 = registret_val_periode_4_fomr_6.replace(",", "."); registret_val_periode_4_fomr_6 = parseFloat(registret_val_periode_4_fomr_6);
        var registret_val_periode_4_fomr_7 = $("#periode_4_fomr_7_registreret_" + thisid).val(); registret_val_periode_4_fomr_7 = registret_val_periode_4_fomr_7.replace(",", "."); registret_val_periode_4_fomr_7 = parseFloat(registret_val_periode_4_fomr_7);
        var registret_val_periode_4_fomr_8 = $("#periode_4_fomr_8_registreret_" + thisid).val(); registret_val_periode_4_fomr_8 = registret_val_periode_4_fomr_8.replace(",", "."); registret_val_periode_4_fomr_8 = parseFloat(registret_val_periode_4_fomr_8);
        var registret_val_periode_4_fomr_9 = $("#periode_4_fomr_9_registreret_" + thisid).val(); registret_val_periode_4_fomr_9 = registret_val_periode_4_fomr_9.replace(",", "."); registret_val_periode_4_fomr_9 = parseFloat(registret_val_periode_4_fomr_9);
        var registret_val_periode_4_fomr_10 = $("#periode_4_fomr_10_registreret_" + thisid).val(); registret_val_periode_4_fomr_10 = registret_val_periode_4_fomr_10.replace(",", "."); registret_val_periode_4_fomr_10 = parseFloat(registret_val_periode_4_fomr_10);
        var registret_val_periode_4_fomr_11 = $("#periode_4_fomr_11_registreret_" + thisid).val(); registret_val_periode_4_fomr_11 = registret_val_periode_4_fomr_11.replace(",", "."); registret_val_periode_4_fomr_11 = parseFloat(registret_val_periode_4_fomr_11);
        var registret_val_periode_4_fomr_12 = $("#periode_4_fomr_12_registreret_" + thisid).val(); registret_val_periode_4_fomr_12 = registret_val_periode_4_fomr_12.replace(",", "."); registret_val_periode_4_fomr_12 = parseFloat(registret_val_periode_4_fomr_12);
        var registret_val_periode_4_fomr_13 = $("#periode_4_fomr_13_registreret_" + thisid).val(); registret_val_periode_4_fomr_13 = registret_val_periode_4_fomr_13.replace(",", "."); registret_val_periode_4_fomr_13 = parseFloat(registret_val_periode_4_fomr_13);
        var registret_val_periode_4_fomr_14 = $("#periode_4_fomr_14_registreret_" + thisid).val(); registret_val_periode_4_fomr_14 = registret_val_periode_4_fomr_14.replace(",", "."); registret_val_periode_4_fomr_14 = parseFloat(registret_val_periode_4_fomr_14);
        var registret_val_periode_4_fomr_15 = $("#periode_4_fomr_15_registreret_" + thisid).val(); registret_val_periode_4_fomr_15 = registret_val_periode_4_fomr_15.replace(",", "."); registret_val_periode_4_fomr_15 = parseFloat(registret_val_periode_4_fomr_15);

        var underskrift_periode_1 = $("#underskrift_periode_1").val();
        var underskrift_periode_2 = $("#underskrift_periode_2").val();
        var underskrift_periode_3 = $("#underskrift_periode_3").val();
        var underskrift_periode_4 = $("#underskrift_periode_4").val();
        
     /* alert(registret_val_periode_1_fomr_1)
        alert(registret_val_periode_1_fomr_2)
        alert(registret_val_periode_1_fomr_3)
        alert(registret_val_periode_1_fomr_4)
        alert(registret_val_periode_1_fomr_5)
        alert(registret_val_periode_1_fomr_6)
        alert(registret_val_periode_1_fomr_7)
        alert(registret_val_periode_1_fomr_8)
        alert(registret_val_periode_1_fomr_9)
        alert(registret_val_periode_1_fomr_10)
        
        alert(registret_val_periode_1_fomr_12) */

        //alert(medarbnavn)
        var chart = new CanvasJS.Chart(this, {
           /* title: {
                text: medarbnavn
             }, */
            axisY: {
                title: "Tid i procent %"
            },
            axisY: {
                maximum: 200,
                suffix: " %"
            },
            toolTip: {
                content: "{label} <br/>{name} : {y} % af norm"
            },
            data: [
                {
                    type: "stackedColumn",
                    color: $("#color_fomr_1").val(),
                    name: fomrnavn1,
                    showInLegend: false,
                    dataPoints: [
                        { label: "FC " + underskrift_periode_1, y: forcastet_val_periode_1_fomr_1 },
                        { label: "AC ", y: registret_val_periode_1_fomr_1 },  
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_2, y: forcastet_val_periode_2_fomr_1 },
                        { label: "AC ", y: registret_val_periode_2_fomr_1 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_3, y: forcastet_val_periode_3_fomr_1 },
                        { label: "AC ", y: registret_val_periode_3_fomr_1 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_4, y: forcastet_val_periode_4_fomr_1 },
                        { label: "AC ", y: registret_val_periode_4_fomr_1 }
                    ]
                },
                {
                    type: "stackedColumn",
                    color: $("#color_fomr_2").val(),
                    name: fomrnavn2,
                    showInLegend: false,
                    dataPoints: [
                        { label: "FC " + underskrift_periode_1, y: forcastet_val_periode_1_fomr_2 },
                        { label: "AC ", y: registret_val_periode_1_fomr_2 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_2, y: forcastet_val_periode_2_fomr_2 },
                        { label: "AC ", y: registret_val_periode_2_fomr_2 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_3, y: forcastet_val_periode_3_fomr_2 },
                        { label: "AC ", y: registret_val_periode_3_fomr_2 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_4, y: forcastet_val_periode_4_fomr_2 },
                        { label: "AC ", y: registret_val_periode_4_fomr_2 }
                        
                    ]
                },
                {
                    type: "stackedColumn",
                    color: $("#color_fomr_3").val(),
                    name: fomrnavn3,
                    showInLegend: false,
                    dataPoints: [
                        { label: "FC " + underskrift_periode_1, y: forcastet_val_periode_1_fomr_3 },
                        { label: "AC ", y: registret_val_periode_1_fomr_3 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_2, y: forcastet_val_periode_2_fomr_3 },
                        { label: "AC ", y: registret_val_periode_2_fomr_3 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_3, y: forcastet_val_periode_3_fomr_3 },
                        { label: "AC ", y: registret_val_periode_3_fomr_3 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_4, y: forcastet_val_periode_4_fomr_3 },
                        { label: "AC ", y: registret_val_periode_4_fomr_3 }
                        
                    ]
                },
                {
                    type: "stackedColumn",
                    color: $("#color_fomr_4").val(),
                    name: fomrnavn4,
                    showInLegend: false,
                    dataPoints: [
                        { label: "FC " + underskrift_periode_1, y: forcastet_val_periode_1_fomr_4 },
                        { label: "AC ", y: registret_val_periode_1_fomr_4 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_2, y: forcastet_val_periode_2_fomr_4 },
                        { label: "AC ", y: registret_val_periode_2_fomr_4 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_3, y: forcastet_val_periode_3_fomr_4 },
                        { label: "AC ", y: registret_val_periode_3_fomr_4 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_4, y: forcastet_val_periode_4_fomr_4 },
                        { label: "AC ", y: registret_val_periode_4_fomr_4 }

                        
                    ]
                },
                {
                    type: "stackedColumn",
                    color: $("#color_fomr_5").val(),
                    name: fomrnavn5,
                    showInLegend: false,
                    dataPoints: [
                        { label: "FC " + underskrift_periode_1, y: forcastet_val_periode_1_fomr_5 },
                        { label: "AC ", y: registret_val_periode_1_fomr_5 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_2, y: forcastet_val_periode_2_fomr_5 },
                        { label: "AC ", y: registret_val_periode_2_fomr_5 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_3, y: forcastet_val_periode_3_fomr_5 },
                        { label: "AC ", y: registret_val_periode_3_fomr_5 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_4, y: forcastet_val_periode_4_fomr_5 },
                        { label: "AC ", y: registret_val_periode_4_fomr_5 }


                    ]
                },
                {
                    type: "stackedColumn",
                    color: $("#color_fomr_6").val(),
                    name: fomrnavn6,
                    showInLegend: false,
                    dataPoints: [
                        { label: "FC " + underskrift_periode_1, y: forcastet_val_periode_1_fomr_6 },
                        { label: "AC ", y: registret_val_periode_1_fomr_6 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_2, y: forcastet_val_periode_2_fomr_6 },
                        { label: "AC ", y: registret_val_periode_2_fomr_6 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_3, y: forcastet_val_periode_3_fomr_6 },
                        { label: "AC ", y: registret_val_periode_3_fomr_6 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_4, y: forcastet_val_periode_4_fomr_6 },
                        { label: "AC ", y: registret_val_periode_4_fomr_6 }


                    ]
                },
                {
                    type: "stackedColumn",
                    color: $("#color_fomr_7").val(),
                    name: fomrnavn7,
                    showInLegend: false,
                    dataPoints: [
                        { label: "FC " + underskrift_periode_1, y: forcastet_val_periode_1_fomr_7 },
                        { label: "AC ", y: registret_val_periode_1_fomr_7 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_2, y: forcastet_val_periode_2_fomr_7 },
                        { label: "AC ", y: registret_val_periode_2_fomr_7 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_3, y: forcastet_val_periode_3_fomr_7 },
                        { label: "AC ", y: registret_val_periode_3_fomr_7 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_4, y: forcastet_val_periode_4_fomr_7 },
                        { label: "AC ", y: registret_val_periode_4_fomr_7 }


                    ]
                },
                {
                    type: "stackedColumn",
                    color: $("#color_fomr_8").val(),
                    name: fomrnavn8,
                    showInLegend: false,
                    dataPoints: [
                        { label: "FC " + underskrift_periode_1, y: forcastet_val_periode_1_fomr_8 },
                        { label: "AC ", y: registret_val_periode_1_fomr_8 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_2, y: forcastet_val_periode_2_fomr_8 },
                        { label: "AC ", y: registret_val_periode_2_fomr_8 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_3, y: forcastet_val_periode_3_fomr_8 },
                        { label: "AC ", y: registret_val_periode_3_fomr_8 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_4, y: forcastet_val_periode_4_fomr_8 },
                        { label: "AC ", y: registret_val_periode_4_fomr_8 }


                    ]
                },
                {
                    type: "stackedColumn",
                    color: $("#color_fomr_9").val(),
                    name: fomrnavn9,
                    showInLegend: false,
                    dataPoints: [
                        { label: "FC " + underskrift_periode_1, y: forcastet_val_periode_1_fomr_9 },
                        { label: "AC ", y: registret_val_periode_1_fomr_9 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_2, y: forcastet_val_periode_2_fomr_9 },
                        { label: "AC ", y: registret_val_periode_2_fomr_9 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_3, y: forcastet_val_periode_3_fomr_9 },
                        { label: "AC ", y: registret_val_periode_3_fomr_9 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_4, y: forcastet_val_periode_4_fomr_9 },
                        { label: "AC ", y: registret_val_periode_4_fomr_9 }


                    ]
                },
                {
                    type: "stackedColumn",
                    color: $("#color_fomr_10").val(),
                    name: fomrnavn10,
                    showInLegend: false,
                    dataPoints: [
                        { label: "FC " + underskrift_periode_1, y: forcastet_val_periode_1_fomr_10 },
                        { label: "AC ", y: registret_val_periode_1_fomr_10 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_2, y: forcastet_val_periode_2_fomr_10 },
                        { label: "AC ", y: registret_val_periode_2_fomr_10 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_3, y: forcastet_val_periode_3_fomr_10 },
                        { label: "AC ", y: registret_val_periode_3_fomr_10 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_4, y: forcastet_val_periode_4_fomr_10 },
                        { label: "AC ", y: registret_val_periode_4_fomr_10 }


                    ]
                },
                {
                    type: "stackedColumn",
                    color: $("#color_fomr_11").val(),
                    name: fomrnavn11,
                    showInLegend: false,
                    dataPoints: [
                        { label: "FC " + underskrift_periode_1, y: forcastet_val_periode_1_fomr_11 },
                        { label: "AC ", y: registret_val_periode_1_fomr_11 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_2, y: forcastet_val_periode_2_fomr_11 },
                        { label: "AC ", y: registret_val_periode_2_fomr_11 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_3, y: forcastet_val_periode_3_fomr_11 },
                        { label: "AC ", y: registret_val_periode_3_fomr_11 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_4, y: forcastet_val_periode_4_fomr_11 },
                        { label: "AC ", y: registret_val_periode_4_fomr_11 }


                    ]
                },
                {
                    type: "stackedColumn",
                    color: $("#color_fomr_12").val(),
                    name: fomrnavn12,
                    showInLegend: false,
                    dataPoints: [
                        { label: "FC " + underskrift_periode_1, y: forcastet_val_periode_1_fomr_12 },
                        { label: "AC ", y: registret_val_periode_1_fomr_12 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_2, y: forcastet_val_periode_2_fomr_12 },
                        { label: "AC ", y: registret_val_periode_2_fomr_12 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_3, y: forcastet_val_periode_3_fomr_12 },
                        { label: "AC ", y: registret_val_periode_3_fomr_12 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_4, y: forcastet_val_periode_4_fomr_12 },
                        { label: "AC ", y: registret_val_periode_4_fomr_12 }


                    ]
                },
                {
                    type: "stackedColumn",
                    color: $("#color_fomr_13").val(),
                    name: fomrnavn13,
                    showInLegend: false,
                    dataPoints: [
                        { label: "FC " + underskrift_periode_1, y: forcastet_val_periode_1_fomr_13 },
                        { label: "AC ", y: registret_val_periode_1_fomr_13 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_2, y: forcastet_val_periode_2_fomr_13 },
                        { label: "AC ", y: registret_val_periode_2_fomr_13 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_3, y: forcastet_val_periode_3_fomr_13 },
                        { label: "AC ", y: registret_val_periode_3_fomr_13 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_4, y: forcastet_val_periode_4_fomr_13 },
                        { label: "AC ", y: registret_val_periode_4_fomr_13 }


                    ]
                },
                {
                    type: "stackedColumn",
                    color: $("#color_fomr_14").val(),
                    name: fomrnavn14,
                    showInLegend: false,
                    dataPoints: [
                        { label: "FC " + underskrift_periode_1, y: forcastet_val_periode_1_fomr_14 },
                        { label: "AC ", y: registret_val_periode_1_fomr_14 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_2, y: forcastet_val_periode_2_fomr_14 },
                        { label: "AC ", y: registret_val_periode_2_fomr_14 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_3, y: forcastet_val_periode_3_fomr_14 },
                        { label: "AC ", y: registret_val_periode_3_fomr_14 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_4, y: forcastet_val_periode_4_fomr_14 },
                        { label: "AC ", y: registret_val_periode_4_fomr_14 }


                    ]
                },
                {
                    type: "stackedColumn",
                    color: $("#color_fomr_15").val(),
                    name: fomrnavn15,
                    showInLegend: false,
                    dataPoints: [
                        { label: "FC " + underskrift_periode_1, y: forcastet_val_periode_1_fomr_15 },
                        { label: "AC ", y: registret_val_periode_1_fomr_15 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_2, y: forcastet_val_periode_2_fomr_15 },
                        { label: "AC ", y: registret_val_periode_2_fomr_15 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_3, y: forcastet_val_periode_3_fomr_15 },
                        { label: "AC ", y: registret_val_periode_3_fomr_15 },
                        { label: " ", y: 0 },
                        { label: "FC " + underskrift_periode_4, y: forcastet_val_periode_4_fomr_15 },
                        { label: "AC ", y: registret_val_periode_4_fomr_15 }


                    ]
                }
            ]
        });
        chart.render();

    });

});


$(document).ready(function () {

    var underskrift_periode_1 = $("#underskrift_periode_1").val();
    var underskrift_periode_2 = $("#underskrift_periode_2").val();
    var underskrift_periode_3 = $("#underskrift_periode_3").val();
    var underskrift_periode_4 = $("#underskrift_periode_4").val();

    var totReg_1_1 = $("#totreg_0_0").val(); totReg_1_1 = totReg_1_1.replace(',', '.'); totReg_1_1 = parseFloat(totReg_1_1)
    var totReg_1_2 = $("#totreg_0_1").val(); totReg_1_2 = totReg_1_2.replace(',', '.'); totReg_1_2 = parseFloat(totReg_1_2)
    var totReg_1_3 = $("#totreg_0_2").val(); totReg_1_3 = totReg_1_3.replace(',', '.'); totReg_1_3 = parseFloat(totReg_1_3)
    var totReg_1_4 = $("#totreg_0_3").val(); totReg_1_4 = totReg_1_4.replace(',', '.'); totReg_1_4 = parseFloat(totReg_1_4)
    
    var totReg_2_1 = $("#totreg_1_0").val(); totReg_2_1 = totReg_2_1.replace(',', '.'); totReg_2_1 = parseFloat(totReg_2_1)
    var totReg_2_2 = $("#totreg_1_1").val(); totReg_2_2 = totReg_2_2.replace(',', '.'); totReg_2_2 = parseFloat(totReg_2_2)
    var totReg_2_3 = $("#totreg_1_2").val(); totReg_2_3 = totReg_2_3.replace(',', '.'); totReg_2_3 = parseFloat(totReg_2_3)
    var totReg_2_4 = $("#totreg_1_3").val(); totReg_2_4 = totReg_2_4.replace(',', '.'); totReg_2_4 = parseFloat(totReg_2_4)
   
    var totReg_3_1 = $("#totreg_2_0").val(); totReg_3_1 = totReg_3_1.replace(',', '.'); totReg_3_1 = parseFloat(totReg_3_1)
    var totReg_3_2 = $("#totreg_2_1").val(); totReg_3_2 = totReg_3_2.replace(',', '.'); totReg_3_2 = parseFloat(totReg_3_2)
    var totReg_3_3 = $("#totreg_2_2").val(); totReg_3_3 = totReg_3_3.replace(',', '.'); totReg_3_3 = parseFloat(totReg_3_3)
    var totReg_3_4 = $("#totreg_2_3").val(); totReg_3_4 = totReg_3_4.replace(',', '.'); totReg_3_4 = parseFloat(totReg_3_4)
    
    var totReg_4_1 = $("#totreg_3_0").val(); totReg_4_1 = totReg_4_1.replace(',', '.'); totReg_4_1 = parseFloat(totReg_4_1)
    var totReg_4_2 = $("#totreg_3_1").val(); totReg_4_2 = totReg_4_2.replace(',', '.'); totReg_4_2 = parseFloat(totReg_4_2)
    var totReg_4_3 = $("#totreg_3_2").val(); totReg_4_3 = totReg_4_3.replace(',', '.'); totReg_4_3 = parseFloat(totReg_4_3)
    var totReg_4_4 = $("#totreg_3_3").val(); totReg_4_4 = totReg_4_4.replace(',', '.'); totReg_4_4 = parseFloat(totReg_4_4)

    var totReg_5_1 = $("#totreg_4_0").val(); totReg_5_1 = totReg_5_1.replace(',', '.'); totReg_5_1 = parseFloat(totReg_5_1)
    var totReg_5_2 = $("#totreg_4_1").val(); totReg_5_2 = totReg_5_2.replace(',', '.'); totReg_5_2 = parseFloat(totReg_5_2)
    var totReg_5_3 = $("#totreg_4_2").val(); totReg_5_3 = totReg_5_3.replace(',', '.'); totReg_5_3 = parseFloat(totReg_5_3)
    var totReg_5_4 = $("#totreg_4_3").val(); totReg_5_4 = totReg_5_4.replace(',', '.'); totReg_5_4 = parseFloat(totReg_5_4)

    var totReg_6_1 = $("#totreg_5_0").val(); totReg_6_1 = totReg_6_1.replace(',', '.'); totReg_6_1 = parseFloat(totReg_6_1)
    var totReg_6_2 = $("#totreg_5_1").val(); totReg_6_2 = totReg_6_2.replace(',', '.'); totReg_6_2 = parseFloat(totReg_6_2)
    var totReg_6_3 = $("#totreg_5_2").val(); totReg_6_3 = totReg_6_3.replace(',', '.'); totReg_6_3 = parseFloat(totReg_6_3)
    var totReg_6_4 = $("#totreg_5_3").val(); totReg_6_4 = totReg_6_4.replace(',', '.'); totReg_6_4 = parseFloat(totReg_6_4)

    var totReg_7_1 = $("#totreg_6_0").val(); totReg_7_1 = totReg_7_1.replace(',', '.'); totReg_7_1 = parseFloat(totReg_7_1)
    var totReg_7_2 = $("#totreg_6_1").val(); totReg_7_2 = totReg_7_2.replace(',', '.'); totReg_7_2 = parseFloat(totReg_7_2)
    var totReg_7_3 = $("#totreg_6_2").val(); totReg_7_3 = totReg_7_3.replace(',', '.'); totReg_7_3 = parseFloat(totReg_7_3)
    var totReg_7_4 = $("#totreg_6_3").val(); totReg_7_4 = totReg_7_4.replace(',', '.'); totReg_7_4 = parseFloat(totReg_7_4)

    var totReg_8_1 = $("#totreg_7_0").val(); totReg_8_1 = totReg_8_1.replace(',', '.'); totReg_8_1 = parseFloat(totReg_8_1)
    var totReg_8_2 = $("#totreg_7_1").val(); totReg_8_2 = totReg_8_2.replace(',', '.'); totReg_8_2 = parseFloat(totReg_8_2)
    var totReg_8_3 = $("#totreg_7_2").val(); totReg_8_3 = totReg_8_3.replace(',', '.'); totReg_8_3 = parseFloat(totReg_8_3)
    var totReg_8_4 = $("#totreg_7_3").val(); totReg_8_4 = totReg_8_4.replace(',', '.'); totReg_8_4 = parseFloat(totReg_8_4)

    var totReg_9_1 = $("#totreg_8_0").val(); totReg_9_1 = totReg_9_1.replace(',', '.'); totReg_9_1 = parseFloat(totReg_9_1)
    var totReg_9_2 = $("#totreg_8_1").val(); totReg_9_2 = totReg_9_2.replace(',', '.'); totReg_9_2 = parseFloat(totReg_9_2)
    var totReg_9_3 = $("#totreg_8_2").val(); totReg_9_3 = totReg_9_3.replace(',', '.'); totReg_9_3 = parseFloat(totReg_9_3)
    var totReg_9_4 = $("#totreg_8_3").val(); totReg_9_4 = totReg_9_4.replace(',', '.'); totReg_9_4 = parseFloat(totReg_9_4)

    var totReg_10_1 = $("#totreg_9_0").val(); totReg_10_1 = totReg_10_1.replace(',', '.'); totReg_10_1 = parseFloat(totReg_10_1)
    var totReg_10_2 = $("#totreg_9_1").val(); totReg_10_2 = totReg_10_2.replace(',', '.'); totReg_10_2 = parseFloat(totReg_10_2)
    var totReg_10_3 = $("#totreg_9_2").val(); totReg_10_3 = totReg_10_3.replace(',', '.'); totReg_10_3 = parseFloat(totReg_10_3)
    var totReg_10_4 = $("#totreg_9_3").val(); totReg_10_4 = totReg_10_4.replace(',', '.'); totReg_10_4 = parseFloat(totReg_10_4)
   
    var totReg_11_1 = $("#totreg_10_0").val(); totReg_11_1 = totReg_11_1.replace(',', '.'); totReg_11_1 = parseFloat(totReg_11_1)
    var totReg_11_2 = $("#totreg_10_1").val(); totReg_11_2 = totReg_11_2.replace(',', '.'); totReg_11_2 = parseFloat(totReg_11_2)
    var totReg_11_3 = $("#totreg_10_2").val(); totReg_11_3 = totReg_11_3.replace(',', '.'); totReg_11_3 = parseFloat(totReg_11_3)
    var totReg_11_4 = $("#totreg_10_3").val(); totReg_11_4 = totReg_11_4.replace(',', '.'); totReg_11_4 = parseFloat(totReg_11_4)
   
    var totReg_12_1 = $("#totreg_11_0").val(); totReg_12_1 = totReg_12_1.replace(',', '.'); totReg_12_1 = parseFloat(totReg_12_1)
    var totReg_12_2 = $("#totreg_11_1").val(); totReg_12_2 = totReg_12_2.replace(',', '.'); totReg_12_2 = parseFloat(totReg_12_2)
    var totReg_12_3 = $("#totreg_11_2").val(); totReg_12_3 = totReg_12_3.replace(',', '.'); totReg_12_3 = parseFloat(totReg_12_3)
    var totReg_12_4 = $("#totreg_11_3").val(); totReg_12_4 = totReg_12_4.replace(',', '.'); totReg_12_4 = parseFloat(totReg_12_4)

    var totReg_13_1 = $("#totreg_12_0").val(); totReg_13_1 = totReg_13_1.replace(',', '.'); totReg_13_1 = parseFloat(totReg_13_1)
    var totReg_13_2 = $("#totreg_12_1").val(); totReg_13_2 = totReg_13_2.replace(',', '.'); totReg_13_2 = parseFloat(totReg_13_2)
    var totReg_13_3 = $("#totreg_12_2").val(); totReg_13_3 = totReg_13_3.replace(',', '.'); totReg_13_3 = parseFloat(totReg_13_3)
    var totReg_13_4 = $("#totreg_12_3").val(); totReg_13_4 = totReg_13_4.replace(',', '.'); totReg_13_4 = parseFloat(totReg_13_4)

    var totReg_14_1 = $("#totreg_13_0").val(); totReg_14_1 = totReg_14_1.replace(',', '.'); totReg_14_1 = parseFloat(totReg_14_1)
    var totReg_14_2 = $("#totreg_13_1").val(); totReg_14_2 = totReg_14_2.replace(',', '.'); totReg_14_2 = parseFloat(totReg_14_2)
    var totReg_14_3 = $("#totreg_13_2").val(); totReg_14_3 = totReg_14_3.replace(',', '.'); totReg_14_3 = parseFloat(totReg_14_3)
    var totReg_14_4 = $("#totreg_13_3").val(); totReg_14_4 = totReg_14_4.replace(',', '.'); totReg_14_4 = parseFloat(totReg_14_4)

    var totReg_15_1 = $("#totreg_14_0").val(); totReg_15_1 = totReg_15_1.replace(',', '.'); totReg_15_1 = parseFloat(totReg_15_1)
    var totReg_15_2 = $("#totreg_14_1").val(); totReg_15_2 = totReg_15_2.replace(',', '.'); totReg_15_2 = parseFloat(totReg_15_2)
    var totReg_15_3 = $("#totreg_14_2").val(); totReg_15_3 = totReg_15_3.replace(',', '.'); totReg_15_3 = parseFloat(totReg_15_3)
    var totReg_15_4 = $("#totreg_14_3").val(); totReg_15_4 = totReg_15_4.replace(',', '.'); totReg_15_4 = parseFloat(totReg_15_4)

    // FC
    var totFC_1_1 = $("#tofc_0_0").val(); totFC_1_1 = totFC_1_1.replace(',', '.'); totFC_1_1 = parseFloat(totFC_1_1)
    var totFC_1_2 = $("#tofc_0_1").val(); totFC_1_2 = totFC_1_2.replace(',', '.'); totFC_1_2 = parseFloat(totFC_1_2)
    var totFC_1_3 = $("#tofc_0_2").val(); totFC_1_3 = totFC_1_3.replace(',', '.'); totFC_1_3 = parseFloat(totFC_1_3)
    var totFC_1_4 = $("#tofc_0_3").val(); totFC_1_4 = totFC_1_4.replace(',', '.'); totFC_1_4 = parseFloat(totFC_1_4)

    var totFC_2_1 = $("#tofc_1_0").val(); totFC_2_1 = totFC_2_1.replace(',', '.'); totFC_2_1 = parseFloat(totFC_2_1)
    var totFC_2_2 = $("#tofc_1_1").val(); totFC_2_2 = totFC_2_2.replace(',', '.'); totFC_2_2 = parseFloat(totFC_2_2)
    var totFC_2_3 = $("#tofc_1_2").val(); totFC_2_3 = totFC_2_3.replace(',', '.'); totFC_2_3 = parseFloat(totFC_2_3)
    var totFC_2_4 = $("#tofc_1_3").val(); totFC_2_4 = totFC_2_4.replace(',', '.'); totFC_2_4 = parseFloat(totFC_2_4)

    var totFC_3_1 = $("#tofc_2_0").val(); totFC_3_1 = totFC_3_1.replace(',', '.'); totFC_3_1 = parseFloat(totFC_3_1)
    var totFC_3_2 = $("#tofc_2_1").val(); totFC_3_2 = totFC_3_2.replace(',', '.'); totFC_3_2 = parseFloat(totFC_3_2)
    var totFC_3_3 = $("#tofc_2_2").val(); totFC_3_3 = totFC_3_3.replace(',', '.'); totFC_3_3 = parseFloat(totFC_3_3)
    var totFC_3_4 = $("#tofc_2_3").val(); totFC_3_4 = totFC_3_4.replace(',', '.'); totFC_3_4 = parseFloat(totFC_3_4)

    var totFC_4_1 = $("#tofc_3_0").val(); totFC_4_1 = totFC_4_1.replace(',', '.'); totFC_4_1 = parseFloat(totFC_4_1)
    var totFC_4_2 = $("#tofc_3_1").val(); totFC_4_2 = totFC_4_2.replace(',', '.'); totFC_4_2 = parseFloat(totFC_4_2)
    var totFC_4_3 = $("#tofc_3_2").val(); totFC_4_3 = totFC_4_3.replace(',', '.'); totFC_4_3 = parseFloat(totFC_4_3)
    var totFC_4_4 = $("#tofc_3_3").val(); totFC_4_4 = totFC_4_4.replace(',', '.'); totFC_4_4 = parseFloat(totFC_4_4)

    var totFC_5_1 = $("#tofc_4_0").val(); totFC_5_1 = totFC_5_1.replace(',', '.'); totFC_5_1 = parseFloat(totFC_5_1)
    var totFC_5_2 = $("#tofc_4_1").val(); totFC_5_2 = totFC_5_2.replace(',', '.'); totFC_5_2 = parseFloat(totFC_5_2)
    var totFC_5_3 = $("#tofc_4_2").val(); totFC_5_3 = totFC_5_3.replace(',', '.'); totFC_5_3 = parseFloat(totFC_5_3)
    var totFC_5_4 = $("#tofc_4_3").val(); totFC_5_4 = totFC_5_4.replace(',', '.'); totFC_5_4 = parseFloat(totFC_5_4)

    var totFC_6_1 = $("#tofc_5_0").val(); totFC_6_1 = totFC_6_1.replace(',', '.'); totFC_6_1 = parseFloat(totFC_6_1)
    var totFC_6_2 = $("#tofc_5_1").val(); totFC_6_2 = totFC_6_2.replace(',', '.'); totFC_6_2 = parseFloat(totFC_6_2)
    var totFC_6_3 = $("#tofc_5_2").val(); totFC_6_3 = totFC_6_3.replace(',', '.'); totFC_6_3 = parseFloat(totFC_6_3)
    var totFC_6_4 = $("#tofc_5_3").val(); totFC_6_4 = totFC_6_4.replace(',', '.'); totFC_6_4 = parseFloat(totFC_6_4)

    var totFC_7_1 = $("#tofc_6_0").val(); totFC_7_1 = totFC_7_1.replace(',', '.'); totFC_7_1 = parseFloat(totFC_7_1)
    var totFC_7_2 = $("#tofc_6_1").val(); totFC_7_2 = totFC_7_2.replace(',', '.'); totFC_7_2 = parseFloat(totFC_7_2)
    var totFC_7_3 = $("#tofc_6_2").val(); totFC_7_3 = totFC_7_3.replace(',', '.'); totFC_7_3 = parseFloat(totFC_7_3)
    var totFC_7_4 = $("#tofc_6_3").val(); totFC_7_4 = totFC_7_4.replace(',', '.'); totFC_7_4 = parseFloat(totFC_7_4)

    var totFC_8_1 = $("#tofc_7_0").val(); totFC_8_1 = totFC_8_1.replace(',', '.'); totFC_8_1 = parseFloat(totFC_8_1)
    var totFC_8_2 = $("#tofc_7_1").val(); totFC_8_2 = totFC_8_2.replace(',', '.'); totFC_8_2 = parseFloat(totFC_8_2)
    var totFC_8_3 = $("#tofc_7_2").val(); totFC_8_3 = totFC_8_3.replace(',', '.'); totFC_8_3 = parseFloat(totFC_8_3)
    var totFC_8_4 = $("#tofc_7_3").val(); totFC_8_4 = totFC_8_4.replace(',', '.'); totFC_8_4 = parseFloat(totFC_8_4)

    var totFC_9_1 = $("#tofc_8_0").val(); totFC_9_1 = totFC_9_1.replace(',', '.'); totFC_9_1 = parseFloat(totFC_9_1)
    var totFC_9_2 = $("#tofc_8_1").val(); totFC_9_2 = totFC_9_2.replace(',', '.'); totFC_9_2 = parseFloat(totFC_9_2)
    var totFC_9_3 = $("#tofc_8_2").val(); totFC_9_3 = totFC_9_3.replace(',', '.'); totFC_9_3 = parseFloat(totFC_9_3)
    var totFC_9_4 = $("#tofc_8_3").val(); totFC_9_4 = totFC_9_4.replace(',', '.'); totFC_9_4 = parseFloat(totFC_9_4)

    var totFC_10_1 = $("#tofc_9_0").val(); totFC_10_1 = totFC_10_1.replace(',', '.'); totFC_10_1 = parseFloat(totFC_10_1)
    var totFC_10_2 = $("#tofc_9_1").val(); totFC_10_2 = totFC_10_2.replace(',', '.'); totFC_10_2 = parseFloat(totFC_10_2)
    var totFC_10_3 = $("#tofc_9_2").val(); totFC_10_3 = totFC_10_3.replace(',', '.'); totFC_10_3 = parseFloat(totFC_10_3)
    var totFC_10_4 = $("#tofc_9_3").val(); totFC_10_4 = totFC_10_4.replace(',', '.'); totFC_10_4 = parseFloat(totFC_10_4)

    var totFC_11_1 = $("#tofc_10_0").val(); totFC_11_1 = totFC_11_1.replace(',', '.'); totFC_11_1 = parseFloat(totFC_11_1)
    var totFC_11_2 = $("#tofc_10_1").val(); totFC_11_2 = totFC_11_2.replace(',', '.'); totFC_11_2 = parseFloat(totFC_11_2)
    var totFC_11_3 = $("#tofc_10_2").val(); totFC_11_3 = totFC_11_3.replace(',', '.'); totFC_11_3 = parseFloat(totFC_11_3)
    var totFC_11_4 = $("#tofc_10_3").val(); totFC_11_4 = totFC_11_4.replace(',', '.'); totFC_11_4 = parseFloat(totFC_11_4)

    var totFC_12_1 = $("#tofc_11_0").val(); totFC_12_1 = totFC_12_1.replace(',', '.'); totFC_12_1 = parseFloat(totFC_12_1)
    var totFC_12_2 = $("#tofc_11_1").val(); totFC_12_2 = totFC_12_2.replace(',', '.'); totFC_12_2 = parseFloat(totFC_12_2)
    var totFC_12_3 = $("#tofc_11_2").val(); totFC_12_3 = totFC_12_3.replace(',', '.'); totFC_12_3 = parseFloat(totFC_12_3)
    var totFC_12_4 = $("#tofc_11_3").val(); totFC_12_4 = totFC_12_4.replace(',', '.'); totFC_12_4 = parseFloat(totFC_12_4)

    var totFC_13_1 = $("#tofc_12_0").val(); totFC_13_1 = totFC_13_1.replace(',', '.'); totFC_13_1 = parseFloat(totFC_13_1)
    var totFC_13_2 = $("#tofc_12_1").val(); totFC_13_2 = totFC_13_2.replace(',', '.'); totFC_13_2 = parseFloat(totFC_13_2)
    var totFC_13_3 = $("#tofc_12_2").val(); totFC_13_3 = totFC_13_3.replace(',', '.'); totFC_13_3 = parseFloat(totFC_13_3)
    var totFC_13_4 = $("#tofc_12_3").val(); totFC_13_4 = totFC_13_4.replace(',', '.'); totFC_13_4 = parseFloat(totFC_13_4)

    var totFC_14_1 = $("#tofc_13_0").val(); totFC_14_1 = totFC_14_1.replace(',', '.'); totFC_14_1 = parseFloat(totFC_14_1)
    var totFC_14_2 = $("#tofc_13_1").val(); totFC_14_2 = totFC_14_2.replace(',', '.'); totFC_14_2 = parseFloat(totFC_14_2)
    var totFC_14_3 = $("#tofc_13_2").val(); totFC_14_3 = totFC_14_3.replace(',', '.'); totFC_14_3 = parseFloat(totFC_14_3)
    var totFC_14_4 = $("#tofc_13_3").val(); totFC_14_4 = totFC_14_4.replace(',', '.'); totFC_14_4 = parseFloat(totFC_14_4)

    var totFC_15_1 = $("#tofc_14_0").val(); totFC_15_1 = totFC_15_1.replace(',', '.'); totFC_15_1 = parseFloat(totFC_15_1)
    var totFC_15_2 = $("#tofc_14_1").val(); totFC_15_2 = totFC_15_2.replace(',', '.'); totFC_15_2 = parseFloat(totFC_15_2)
    var totFC_15_3 = $("#tofc_14_2").val(); totFC_15_3 = totFC_15_3.replace(',', '.'); totFC_15_3 = parseFloat(totFC_15_3)
    var totFC_15_4 = $("#tofc_14_3").val(); totFC_15_4 = totFC_15_4.replace(',', '.'); totFC_15_4 = parseFloat(totFC_15_4)

    var fomrnavn1 = $("#fomrnavn_0").val();
    var fomrnavn2 = $("#fomrnavn_1").val();
    var fomrnavn3 = $("#fomrnavn_2").val();
    var fomrnavn4 = $("#fomrnavn_3").val();
    var fomrnavn5 = $("#fomrnavn_4").val();
    var fomrnavn6 = $("#fomrnavn_5").val();
    var fomrnavn7 = $("#fomrnavn_6").val();
    var fomrnavn8 = $("#fomrnavn_7").val();
    var fomrnavn9 = $("#fomrnavn_8").val();
    var fomrnavn10 = $("#fomrnavn_9").val();
    var fomrnavn11 = $("#fomrnavn_10").val();
    var fomrnavn12 = $("#fomrnavn_11").val();
    var fomrnavn13 = $("#fomrnavn_12").val();
    var fomrnavn14 = $("#fomrnavn_13").val();
    var fomrnavn15 = $("#fomrnavn_14").val();

    var chart = new CanvasJS.Chart("totalChart", {
        /* title: {
             text: medarbnavn
          }, */
        axisY: {
            title: "Tid i procent %"
        },
        axisY: {
            maximum: 200,
            suffix: " %"
        },
        toolTip: {
            content: "{label} <br/>{name} : {y} % af norm"
        },
        data: [
            {
                type: "stackedColumn",
                color: $("#color_fomr_1").val(),
                name: fomrnavn1,
                showInLegend: false,
                dataPoints: [
                    { label: "FC " + underskrift_periode_1, y: totFC_1_1 },
                    { label: "AC ", y: totReg_1_1 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_2, y: totFC_1_2 },
                    { label: "AC ", y: totReg_1_2 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_3, y: totFC_1_3 },
                    { label: "AC ", y: totReg_1_3 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_4, y: totFC_1_4 },
                    { label: "AC ", y: totReg_1_4 }
                ]
            },
            {
                type: "stackedColumn",
                color: $("#color_fomr_2").val(),
                name: fomrnavn2,
                showInLegend: false,
                dataPoints: [
                    { label: "FC " + underskrift_periode_1, y: totFC_2_1 },
                    { label: "AC ", y: totReg_2_1 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_2, y: totFC_2_2 },
                    { label: "AC ", y: totReg_2_2 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_3, y: totFC_2_3 },
                    { label: "AC ", y: totReg_2_3 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_4, y: totFC_2_4 },
                    { label: "AC ", y: totReg_2_4 }

                ]
            },
            {
                type: "stackedColumn",
                color: $("#color_fomr_3").val(),
                name: fomrnavn3,
                showInLegend: false,
                dataPoints: [
                    { label: "FC " + underskrift_periode_1, y: totFC_3_1 },
                    { label: "AC ", y: totReg_3_1 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_2, y: totFC_3_2 },
                    { label: "AC ", y: totReg_3_2 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_3, y: totFC_3_3 },
                    { label: "AC ", y: totReg_3_3 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_4, y: totFC_3_4 },
                    { label: "AC ", y: totReg_3_4 }

                ]
            },
            {
                type: "stackedColumn",
                color: $("#color_fomr_4").val(),
                name: fomrnavn4,
                showInLegend: false,
                dataPoints: [
                    { label: "FC " + underskrift_periode_1, y: totFC_4_1 },
                    { label: "AC ", y: totReg_4_1 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_2, y: totFC_4_2 },
                    { label: "AC ", y: totReg_4_2 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_3, y: totFC_4_3 },
                    { label: "AC ", y: totReg_4_3 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_4, y: totFC_4_4 },
                    { label: "AC ", y: totReg_4_4 }


                ]
            },
            {
                type: "stackedColumn",
                color: $("#color_fomr_5").val(),
                name: fomrnavn5,
                showInLegend: false,
                dataPoints: [
                    { label: "FC " + underskrift_periode_1, y: totFC_5_1 },
                    { label: "AC ", y: totReg_5_1 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_2, y: totFC_5_2 },
                    { label: "AC ", y: totReg_5_2 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_3, y: totFC_5_3 },
                    { label: "AC ", y: totReg_5_3 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_4, y: totFC_5_4 },
                    { label: "AC ", y: totReg_5_4 }


                ]
            },
            {
                type: "stackedColumn",
                color: $("#color_fomr_6").val(),
                name: fomrnavn6,
                showInLegend: false,
                dataPoints: [
                    { label: "FC " + underskrift_periode_1, y: totFC_6_1 },
                    { label: "AC ", y: totReg_6_1 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_2, y: totFC_6_2 },
                    { label: "AC ", y: totReg_6_2 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_3, y: totFC_6_3 },
                    { label: "AC ", y: totReg_6_3 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_4, y: totFC_6_4 },
                    { label: "AC ", y: totReg_6_4 }


                ]
            },
            {
                type: "stackedColumn",
                color: $("#color_fomr_7").val(),
                name: fomrnavn7,
                showInLegend: false,
                dataPoints: [
                    { label: "FC " + underskrift_periode_1, y: totFC_7_1 },
                    { label: "AC ", y: totReg_7_1 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_2, y: totFC_7_2 },
                    { label: "AC ", y: totReg_7_2 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_3, y: totFC_7_3 },
                    { label: "AC ", y: totReg_7_3 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_4, y: totFC_7_4 },
                    { label: "AC ", y: totReg_7_4 }


                ]
            },
            {
                type: "stackedColumn",
                color: $("#color_fomr_8").val(),
                name: fomrnavn8,
                showInLegend: false,
                dataPoints: [
                    { label: "FC " + underskrift_periode_1, y: totFC_8_1 },
                    { label: "AC ", y: totReg_8_1 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_2, y: totFC_8_2 },
                    { label: "AC ", y: totReg_8_2 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_3, y: totFC_8_3 },
                    { label: "AC ", y: totReg_8_3 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_4, y: totFC_8_4 },
                    { label: "AC ", y: totReg_8_4 }


                ]
            },
            {
                type: "stackedColumn",
                color: $("#color_fomr_9").val(),
                name: fomrnavn9,
                showInLegend: false,
                dataPoints: [
                    { label: "FC " + underskrift_periode_1, y: totFC_9_1 },
                    { label: "AC ", y: totReg_9_1 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_2, y: totFC_9_2 },
                    { label: "AC ", y: totReg_9_2 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_3, y: totFC_9_3 },
                    { label: "AC ", y: totReg_9_3 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_4, y: totFC_9_4 },
                    { label: "AC ", y: totReg_9_4 }


                ]
            },
            {
                type: "stackedColumn",
                color: $("#color_fomr_10").val(),
                name: fomrnavn10,
                showInLegend: false,
                dataPoints: [
                    { label: "FC " + underskrift_periode_1, y: totFC_10_1 },
                    { label: "AC ", y: totReg_10_1 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_2, y: totFC_10_2 },
                    { label: "AC ", y: totReg_10_2 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_3, y: totFC_10_3 },
                    { label: "AC ", y: totReg_10_3 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_4, y: totFC_10_4 },
                    { label: "AC ", y: totReg_10_4 }


                ]
            },
            {
                type: "stackedColumn",
                color: $("#color_fomr_11").val(),
                name: fomrnavn11,
                showInLegend: false,
                dataPoints: [
                    { label: "FC " + underskrift_periode_1, y: totFC_11_1 },
                    { label: "AC ", y: totReg_11_1 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_2, y: totFC_11_2 },
                    { label: "AC ", y: totReg_11_2 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_3, y: totFC_11_3 },
                    { label: "AC ", y: totReg_11_3 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_4, y: totFC_11_4 },
                    { label: "AC ", y: totReg_11_4 }


                ]
            },
            {
                type: "stackedColumn",
                color: $("#color_fomr_12").val(),
                name: fomrnavn12,
                showInLegend: false,
                dataPoints: [
                    { label: "FC " + underskrift_periode_1, y: totFC_12_1 },
                    { label: "AC ", y: totReg_12_1 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_2, y: totFC_12_2 },
                    { label: "AC ", y: totReg_12_2 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_3, y: totFC_12_3 },
                    { label: "AC ", y: totReg_12_3 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_4, y: totFC_12_4 },
                    { label: "AC ", y: totReg_12_4 }


                ]
            },
            {
                type: "stackedColumn",
                color: $("#color_fomr_13").val(),
                name: fomrnavn13,
                showInLegend: false,
                dataPoints: [
                    { label: "FC " + underskrift_periode_1, y: totFC_13_1 },
                    { label: "AC ", y: totReg_13_1 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_2, y: totFC_13_2 },
                    { label: "AC ", y: totReg_13_2 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_3, y: totFC_13_3 },
                    { label: "AC ", y: totReg_13_3 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_4, y: totFC_13_4 },
                    { label: "AC ", y: totReg_13_4 }


                ]
            },
            {
                type: "stackedColumn",
                color: $("#color_fomr_14").val(),
                name: fomrnavn14,
                showInLegend: false,
                dataPoints: [
                    { label: "FC " + underskrift_periode_1, y: totFC_14_1 },
                    { label: "AC ", y: totReg_14_1 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_2, y: totFC_14_2 },
                    { label: "AC ", y: totReg_14_2 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_3, y: totFC_14_3 },
                    { label: "AC ", y: totReg_14_3 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_4, y: totFC_14_4 },
                    { label: "AC ", y: totReg_14_4 }


                ]
            },
            {
                type: "stackedColumn",
                color: $("#color_fomr_15").val(),
                name: fomrnavn15,
                showInLegend: false,
                dataPoints: [
                    { label: "FC " + underskrift_periode_1, y: totFC_15_1 },
                    { label: "AC ", y: totReg_15_1 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_2, y: totFC_15_2 },
                    { label: "AC ", y: totReg_15_2 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_3, y: totFC_15_3 },
                    { label: "AC ", y: totReg_15_3 },
                    { label: " ", y: 0 },
                    { label: "FC " + underskrift_periode_4, y: totFC_15_4 },
                    { label: "AC ", y: totReg_15_4 }


                ]
            }
        ]
    });
    chart.render();


});