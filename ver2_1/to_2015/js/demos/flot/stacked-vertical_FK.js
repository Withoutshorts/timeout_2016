$(function () {

    //alert("awd")

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
            },
            toolTip: {
                content: "{label} <br/>{name} : {y} % af norm"
            },
            data: [
                {
                    type: "stackedColumn",
                    color: $("#color_fomr_1").val(),
                    name: "Projekt1",
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
                    name: "Projekt2",
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
                    name: "Projekt3",
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
                    name: "Projekt4",
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
                    name: "Projekt5",
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
                    name: "Projekt6",
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
                    name: "Projekt7",
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
                    name: "Projekt8",
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
                    name: "Projekt9",
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
                    name: "Projekt10",
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
                    name: "Projekt11",
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
                    name: "Projekt12",
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
                    name: "Projekt13",
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
                    name: "Projekt14",
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
                    name: "Projekt15",
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