

$(document).ready(function () {

    var expenceTotal = 1;
    var lto = $("#lto").val();
    var matid = 0 //$("#matid_0").val();
    matid = parseInt(matid);

    var totalExpenesToMake = 1;

    $("#removeExpence").hide();

    var nameTxt = $("#nameTxt").val();
    var priceTxt = $("#priceTxt").val();
    var notbillableTxt = $("#notbillableTxt").val();
    var billableTxt = $("#billableTxt").val();
    var companypaidTxt = $("#companypaidTxt").val();
    var personalTxt = $("#personalTxt").val();
    var selectimageTxt = $("#selectimageTxt").val();

  //  alert("Hep")
    
    $("#addExpence").click(function () {

        $("#removeExpence").hide();
        $("#addExpence").hide();

        totalExpenesToMake++;

        // $("#expenceTable").clone().appendTo("#expencestacker");

        strExpenceSection = "<table style='width:100%; border:solid 1px #c9c9c9; margin-top:10px;' id='expenceTable_" + expenceTotal +"'>"

        strExpenceSection += "<tr>"
        strExpenceSection += "<td colspan='2'><input type='text' name='FM_udlaeg_navn' id='FM_udlaeg_navn_" + expenceTotal + "' placeholder='" + nameTxt + "' class='form-control' /></td>"
        strExpenceSection += "</tr>"

        //Valuta
        strExpenceSection += "<tr>"
        strExpenceSection += "<td>"
        strExpenceSection += "<input type='number' name='FM_udlaeg_belob' id='FM_udlaeg_belob_" + expenceTotal + "' placeholder='" + priceTxt +"' class='form-control' value='' />"
        strExpenceSection += "</td>"
        strExpenceSection += "<td>"
        strExpenceSection += "<select class='form-control' id='FM_udlaeg_valuta_" + expenceTotal + "'>"

        strExpenceSection += "</select>"
        strExpenceSection += "</td>"
        strExpenceSection += "</tr>"

        // Type
        strExpenceSection += "<tr>"
        strExpenceSection += "<td colspan='2'>"
        strExpenceSection += "<select class='form-control' id='FM_udlaeg_gruppe_" + expenceTotal +"'>"
        
        strExpenceSection += "</select>"
        strExpenceSection += "</td>"
        strExpenceSection += "</tr>"

        // Fakt bar
        strExpenceSection += "<tr>"
        strExpenceSection += "<td colspan='2'>"
        strExpenceSection += "<select class='form-control' id='FM_udlaeg_faktbar_" + expenceTotal +"'>"

        if (lto == "nt") {
            strExpenceSection += "<option value='1'>" + notbillableTxt +"</option>"
            strExpenceSection += "<option value='2'>" + billableTxt +" 100%</option>"
            strExpenceSection += "<option value='5'>" + billableTxt +" 50%</option>"
        } else {
            strExpenceSection += "<option value='1'>" + notbillableTxt +"</option>"
            strExpenceSection += "<option value='2'>" + billableTxt +"</option>"
        }
        
        strExpenceSection += "</select>"
        strExpenceSection += "</td>"
        strExpenceSection += "</tr>"


        // Personlig elelr firma
        strExpenceSection += "<tr>"
        strExpenceSection += "<td colspan='2'>"
        strExpenceSection += "<select class='form-control' id='FM_udlaeg_form_" + expenceTotal + "'>"
        if (lto == "nt") {
            strExpenceSection += "<option value='0'>" + companypaidTxt +"</option>"
            strExpenceSection += "<option value='1'>" + personalTxt +"</option>"
        }
        else {
            strExpenceSection += "<option value='1'>" + personalTxt +"</option>"
            strExpenceSection += "<option value='0'>" + companypaidTxt +"</option>"
        }
        
        strExpenceSection += "</select>"
        strExpenceSection += "</td>"
        strExpenceSection += "</tr>"

        strExpenceSection += "<tr>"
        strExpenceSection += "<td colspan='2' style='text-align:center;'>"
        strExpenceSection += "<form ENCTYPE='multipart/form-data' method='post' id='image_upload_" + expenceTotal +"'>"
        strExpenceSection += "<label class='btn btn-default btn-lg'>"
        strExpenceSection += "<INPUT id='" + expenceTotal + "' NAME='fileupload1' TYPE='file' style='width:400px; display:none;' onchange='readURL(this);'>"
        strExpenceSection += "<b>" + selectimageTxt +"</b>"
        strExpenceSection += "</label"
        strExpenceSection += "</form>"
        strExpenceSection += "</td>"
        strExpenceSection += "</tr>"

        strExpenceSection += "<tr>"
        strExpenceSection += "<td colspan='2' style='text-align:center;'>"
        strExpenceSection += "<img id='imageholder_"+ expenceTotal +"' alt='' border='0' style='width:50px;'>"
        strExpenceSection += "</td>"
        strExpenceSection += "</tr>"


        strExpenceSection += "</table>"


        $("#expencestacker").append(strExpenceSection)
        
        id = 1
        $.post("?id=" + id, { control: "GetCurrencies", AjaxUpdateField: "true" }, function (data) {

            $("#FM_udlaeg_valuta_" + expenceTotal).append(data)

            $.post("?id=" + id, { control: "GetExpenecTypes", AjaxUpdateField: "true" }, function (data) {
               
                $("#FM_udlaeg_gruppe_" + expenceTotal).append(data)

                $("#removeExpence").show();
                $("#addExpence").show();

                expenceTotal += 1
            });

            

        });

        


    });

    $("#removeExpence").click(function () {
        var expenceToRemvoe = (parseInt(expenceTotal) - 1)
        $("#expenceTable_" + expenceToRemvoe).remove();

        expenceTotal--;
        totalExpenesToMake--;

      //  alert(expenceTotal + " " + totalExpenesToMake)

        if (expenceTotal == 1)
        {
            $("#removeExpence").hide();
        }

    });



    $("#formclick").click(function () {

        alert("HEJ")
        $("#image_upload").submit();
        

    });

    var ARRAY_udlaeg_navn = [];
    var ARRAY_udlaeg_belob = [];
    var ARRAY_udlaeg_valuta = [];
    var ARRAY_udlaeg_gruppe = [];
    var ARRAY_udlaeg_faktbar = [];
    var ARRAY_udlaeg_form = [];
    var ARRAY_udlaeg_valutakode = [];


    $("#sbmExpence").click(function () {

        if (expenceTotal == 1 && $("#FM_udlaeg_navn_0").val() == "")
        {
            $("#dietsform").submit();
            return;
        }

        // Get exepence id
        $.post("?id=1", { control: "GetExpenceId", AjaxUpdateField: "true" }, function (data) {
            matid = data
            matid = parseInt(matid)
       });

        var error = 0

        $("#errorMessage").text("")

        endpoint = 'convert'
        access_key = 'de1b777d2882c4fe895b0ade03dbb001';

        FM_datoer = $("#FM_datoer").val();

        if (mobil_week_reg_job_dd != 1) {
            jq_jobid = $("#FM_jobid").val();
        } else {
            jq_jobid = $("#dv_job").val();
        }

        if (mobil_week_reg_akt_dd != 1) {
            aktid = $("#FM_aktid").val();
        } else {
            aktid = $("#dv_akt").val();
        }

        if (jq_jobid == 0 || jq_jobid == -1) {
            error = 1
            $("#errorMessage").text("Der mangler at blive valgt et job")
        }

        // alert("HEPHEP")

        for (var i = 0; i < expenceTotal; i++) {
            // Tjekker alle felter for fejl

            FM_udlaeg_navn = $("#FM_udlaeg_navn_" + i).val();
           
            FM_udlaeg_belob = $("#FM_udlaeg_belob_" + i).val();
            FM_udlaeg_belob = FM_udlaeg_belob.replace(",", ".")

            FM_udlaeg_valuta = $("#FM_udlaeg_valuta_" + i).val();
            FM_udlaeg_gruppe = $("#FM_udlaeg_gruppe_" + i).val();
            FM_udlaeg_faktbar = $("#FM_udlaeg_faktbar_" + i).val();
            FM_udlaeg_form = $("#FM_udlaeg_form_" + i).val();

            $("#FM_udlaeg_valuta_" + i + " > option").each(function () {
                if ($(this).is(':selected')) {
                    FM_udlaeg_valutakode = $(this).data('valutakode');
                }
            });

            if (FM_udlaeg_navn == "") {
                error = 1
                $("#errorMessage").text("Der mangler at blive angivet et navn")
                window.scrollTo(0,0);
            }

            if (FM_udlaeg_belob == "") {
                error = 1
                $("#errorMessage").text("Der mangler at blive angivet et beløb")
                window.scrollTo(0, 0);
            }
            
            if (i == 0) {
                ARRAY_udlaeg_navn = FM_udlaeg_navn
                ARRAY_udlaeg_belob = FM_udlaeg_belob
                ARRAY_udlaeg_valuta = FM_udlaeg_valuta
                ARRAY_udlaeg_gruppe = FM_udlaeg_gruppe
                ARRAY_udlaeg_faktbar = FM_udlaeg_faktbar
                ARRAY_udlaeg_form = FM_udlaeg_form
                ARRAY_udlaeg_valutakode = FM_udlaeg_valutakode
            }
            else {
                ARRAY_udlaeg_navn += "," + FM_udlaeg_navn
                ARRAY_udlaeg_belob += "," + FM_udlaeg_belob
                ARRAY_udlaeg_valuta += "," + FM_udlaeg_valuta
                ARRAY_udlaeg_gruppe += "," + FM_udlaeg_gruppe
                ARRAY_udlaeg_faktbar += "," + FM_udlaeg_faktbar
                ARRAY_udlaeg_form += "," + FM_udlaeg_form
                ARRAY_udlaeg_valutakode += "," + FM_udlaeg_valutakode
            }
          //  alert("Val koder " + ARRAY_udlaeg_valutakode)
        }

        ARRAY_udlaeg_navn = ARRAY_udlaeg_navn.split(',');
        ARRAY_udlaeg_belob = ARRAY_udlaeg_belob.split(',');
        ARRAY_udlaeg_valuta = ARRAY_udlaeg_valuta.split(',');
        ARRAY_udlaeg_gruppe = ARRAY_udlaeg_gruppe.split(',');
        ARRAY_udlaeg_faktbar = ARRAY_udlaeg_faktbar.split(',');
        ARRAY_udlaeg_form = ARRAY_udlaeg_form.split(',');
        ARRAY_udlaeg_valutakode = ARRAY_udlaeg_valutakode.split(',');

      //  alert("i " + i + " L " + ARRAY_udlaeg_navn.length)

        var totalCreated = 0;

        if (error == 0)
        {

            $("#uploadbar").show();

           // var udlaeg_navnne = ['Hans', 'Jens', 'Morten'];
           // alert(udlaeg_navnne[0])

            for (var i = 0; i < expenceTotal; i++) {

                
              //  alert(udlaeg_navn)
                

                FM_udlaeg_navn = $("#FM_udlaeg_navn_" + i).val();

                FM_udlaeg_belob = $("#FM_udlaeg_belob_" + i).val();
                FM_udlaeg_belob = FM_udlaeg_belob.replace(",", ".")

                FM_udlaeg_valuta = $("#FM_udlaeg_valuta_" + i).val();
                FM_udlaeg_gruppe = $("#FM_udlaeg_gruppe_" + i).val();
                FM_udlaeg_faktbar = $("#FM_udlaeg_faktbar_" + i).val();
                FM_udlaeg_form = $("#FM_udlaeg_form_" + i).val();

                var from = "NA"
                
                $("#FM_udlaeg_valuta_" + i + " > option").each(function () {
                    if ($(this).is(':selected')) {
                        from = $(this).data('valutakode');
                    }
                });

                to = 'DKK'; // Basic valuta
                amount = $("#FM_udlaeg_belob_" + i).val();
                amount = amount.replace(/\,/g, '.')

              //  alert("For AJ " + FM_udlaeg_navn + " " + FM_udlaeg_belob + " " + FM_udlaeg_valuta + " " + FM_udlaeg_gruppe + " " + FM_udlaeg_faktbar + " " + FM_udlaeg_form + " " + FM_datoer + " " + jq_jobid + " " + aktid)

                if (from != "NA") {
                    $.ajax({
                        url: 'https://apilayer.net/api/' + endpoint + '?access_key=' + access_key + '&from=' + from + '&to=' + to + '&amount=' + amount + "&format=1",
                        dataType: 'jsonp',
                        success: function (json) {
                            var kurs = parseFloat(json.result) / parseFloat(amount)
                            var udlaeg_navn = udlaeg_navnne[totalCreated]
                            //  alert("Result " + json.result + " KURS " + kurs)
                          //  alert("Navn " + udlaeg_navn)
                          //  alert("Aj " + to + " " + kurs + " from " + from + " " + json.result + " " + FM_udlaeg_navn + " " + FM_udlaeg_belob + " " + FM_udlaeg_valuta + " " + FM_udlaeg_gruppe + " " + FM_udlaeg_faktbar + " " + FM_udlaeg_form + " " + FM_datoer + " " + jq_jobid + " " + aktid)

                         //   IndlaesExpence(to, kurs, json.result, FM_udlaeg_navn, FM_udlaeg_belob, FM_udlaeg_valuta, FM_udlaeg_gruppe, FM_udlaeg_faktbar, FM_udlaeg_form, FM_datoer, jq_jobid, aktid);

                            totalCreated += 1;
                        },
                        error: function (json) {

                            alert("error")

                        }
                        
                    });

                }

                
                
            }

            CallAjax()

        }


        /*

        endpoint = 'convert'
        access_key = 'de1b777d2882c4fe895b0ade03dbb001';

        var from = "NA"

        $("#FM_udlaeg_valuta > option").each(function () {
            if ($(this).is(':selected')) {
                from = $(this).data('valutakode');
            }          
        });

       // alert(from)

       // from = 'USD';
        to = 'DKK'; // Basic valuta
        amount = $("#FM_udlaeg_belob").val();
        amount = amount.replace(/\,/g, '.')

        //amount = amount.replace(".", ",")

        if (from != "NA") {
            $.ajax({
                url: 'https://apilayer.net/api/' + endpoint + '?access_key=' + access_key + '&from=' + from + '&to=' + to + '&amount=' + amount + "&format=1",
                dataType: 'jsonp',
                success: function (json) {
                    // access the conversion result in json.result
                   // alert(json.result);
                    var kurs = parseFloat(json.result) / parseFloat(amount)
                   // alert(kurs)
                   // alert(kurs)
                    IndlaesExpence(to, kurs, json.result);
                },
                error: function (json) {

                    alert("error")

                }
            });
        } */

    });

    var i = 0
    
  //  alert(" HEJ " + udlaeg_navnne[0])
   // alert("lengt " + ARRAY_udlaeg_navn.length)
    function CallAjax()
    {

      //  alert(ARRAY_udlaeg_navn.length)

     // alert("Ajax her")
        
        to = "DKK"
        amount = 5; 
       // alert(i)

        var udlaeg_navn = ARRAY_udlaeg_navn[i]
        var udlaeg_belob = ARRAY_udlaeg_belob[i]
        var udlaeg_valuta = ARRAY_udlaeg_valuta[i]
        var udlaeg_gruppe = ARRAY_udlaeg_gruppe[i]
        var udlaeg_faktbar = ARRAY_udlaeg_faktbar[i]
        var udlaeg_form = ARRAY_udlaeg_form[i]
        var udlaeg_fromvalutakode = ARRAY_udlaeg_valutakode[i]

        amount = udlaeg_belob
        from = udlaeg_fromvalutakode

        $.ajax({
            url: 'https://apilayer.net/api/' + endpoint + '?access_key=' + access_key + '&from=' + from + '&to=' + to + '&amount=' + amount + "&format=1",
            dataType: 'jsonp',
            success: function (json) {
                var kurs = parseFloat(json.result) / parseFloat(amount)
                
                //  alert("Result " + json.result + " KURS " + kurs)
                //  alert("Navn " + udlaeg_navn)
                //  alert("Aj " + to + " " + kurs + " from " + from + " " + json.result + " " + FM_udlaeg_navn + " " + FM_udlaeg_belob + " " + FM_udlaeg_valuta + " " + FM_udlaeg_gruppe + " " + FM_udlaeg_faktbar + " " + FM_udlaeg_form + " " + FM_datoer + " " + jq_jobid + " " + aktid)

              //  alert("Navn " + udlaeg_navn + " belob " + udlaeg_belob + " valuta " + udlaeg_valuta + " udlaeg_gruppe " + udlaeg_gruppe + " udlaeg_faktbar " + udlaeg_faktbar + " udlaeg_form " + udlaeg_form)

                IndlaesExpence(to, kurs, json.result, udlaeg_navn, udlaeg_belob, udlaeg_valuta, udlaeg_gruppe, udlaeg_faktbar, udlaeg_form, FM_datoer, jq_jobid, aktid, i);

                if (i < (ARRAY_udlaeg_navn.length - 1))
                {
                    i += 1;
                    CallAjax(i)
                }
                
            },
            error: function (json) {

               // alert("error")

            }

        });
    }


    var expenesCreated = 0
    function IndlaesExpence(basic_valuta, basic_kurs, basic_belob, FM_udlaeg_navn, FM_udlaeg_belob, FM_udlaeg_valuta, FM_udlaeg_gruppe, FM_udlaeg_faktbar, FM_udlaeg_form, FM_datoer, jq_jobid, aktid, expenceNumber)
    {
       // alert("Func " + basic_valuta + " " + basic_kurs + " " + basic_belob + " " + FM_udlaeg_navn + " " + FM_udlaeg_belob + " " + FM_udlaeg_valuta + " " + FM_udlaeg_gruppe + " " + FM_udlaeg_faktbar + " " + FM_udlaeg_form + " " + FM_datoer + " " + jq_jobid + " " + aktid)
       // alert(basic_valuta + " " + basic_kurs + " " + basic_belob)
        /*
        mobil_week_reg_job_dd = $("#mobil_week_reg_job_dd").val();
        mobil_week_reg_akt_dd = $("#mobil_week_reg_akt_dd").val();

        if (mobil_week_reg_job_dd != 1) {
            jq_jobid = $("#FM_jobid").val();
        } else {
            jq_jobid = $("#dv_job").val();
        }

        if (mobil_week_reg_akt_dd != 1) {
            aktid = $("#FM_aktid").val();
        } else {
            aktid = $("#dv_akt").val();
        }


        FM_udlaeg_navn = $("#FM_udlaeg_navn").val();

        FM_udlaeg_belob = $("#FM_udlaeg_belob").val();
        FM_udlaeg_belob = FM_udlaeg_belob.replace(",", ".")

        FM_udlaeg_valuta = $("#FM_udlaeg_valuta").val();
        FM_udlaeg_gruppe = $("#FM_udlaeg_gruppe").val();
        FM_udlaeg_faktbar = $("#FM_udlaeg_faktbar").val();
        FM_udlaeg_form = $("#FM_udlaeg_form").val();

        FM_datoer = $("#FM_datoer").val();

        error = 0

        if (jq_jobid == 0 || jq_jobid == -1) {
            error = 1
            $("#errorMessage").text("Der mangler at blive valgt et job")
        }

        if (FM_udlaeg_navn == "") {
            error = 1
            $("#errorMessage").text("Der mangler at blive angivet et navn")
        }

        if (FM_udlaeg_belob == "") {
            error = 1
            $("#errorMessage").text("Der mangler at blive angivet et beløb")
        } */

        //$("#jobidholderMatreg").val(jq_jobid);

      //  alert("jobid " + jq_jobid + " aktid " + aktid + " navn " + FM_udlaeg_navn + " belob " + FM_udlaeg_belob + " valuta " + FM_udlaeg_valuta + " gruppe " + FM_udlaeg_gruppe + " faktbar " + FM_udlaeg_faktbar + " form " + FM_udlaeg_form)


       // if (error == 0) {

           // alert("7894578678786")
        $.post("?jq_jobid=" + jq_jobid + "&aktid=" + aktid + "&FM_udlaeg_valuta=" + FM_udlaeg_valuta + "&FM_udlaeg_form=" + FM_udlaeg_form + "&FM_udlaeg_belob=" + FM_udlaeg_belob + "&FM_udlaeg_navn=" + FM_udlaeg_navn + "&FM_udlaeg_gruppe=" + FM_udlaeg_gruppe + "&FM_datoer=" + FM_datoer + "&basic_valuta=" + basic_valuta + "&basic_kurs=" + basic_kurs + "&basic_belob=" + basic_belob + "&FM_udlaeg_faktbar=" + FM_udlaeg_faktbar, { control: "FN_createOutlay", AjaxUpdateField: "true" }, function (data) {

              // alert("data " + data)
                

                id = 1
              //  $.post("?id=" + id, { control: "GetExpenceId", AjaxUpdateField: "true" }, function (data) {
                  //  alert("Mat id ") + data
                    //matid = data
                    
                    matid += 1

               // alert(" MAT POST " + matid)

                   // alert("MAT " + matid)
                    var fd = new FormData($("#image_upload_" + expenceNumber)[0])
                   // alert(matid)
                    $.ajax({
                        type: "POST",
                        url: "../timereg/upload_bin.asp?matUpload=1&matId=" + matid + "&thisfile=timetag_mobile",
                        data: fd,
                        processData: false,
                        contentType: false,
                        success: function (data, status) {
                            //this will execute when form is submited without errors
                            // alert("Suc")
                            expenesCreated++;
                            UpdateProgressbar(expenesCreated);
                        },
                        error: function (data, status) {
                            //this will execute when get any error
                            alert(status)
                        },
                    });
                   
               // });

              /*  var fd = new FormData($("#image_upload_" + expenceNumber)[0])

                $.ajax({
                    type: "POST",
                    url: "../timereg/upload_bin.asp?matUpload=1&matId=" + matid + "&thisfile=timetag_mobile",
                    data: fd,
                    processData: false,
                    contentType: false,
                    success: function (data, status) {
                        //this will execute when form is submited without errors
                        alert("Suc")
                    },
                    error: function (data, status) {
                        //this will execute when get any error
                        alert(status)
                    },
                }); */


               // $("#image_upload").submit();
               // alert("done")
               // i = 1;
                //  $("#formtest").submit();
               // alert("Done")

             /*   $("form").each(function () {
                   
                    var fd = new FormData($(this)[0]);
                    $.ajax({
                        type: "POST",
                        url: "../timereg/upload_bin.asp?matUpload=1&matId=" + matid + "&thisfile=timetag_mobile",
                        data: fd,
                        processData: false,
                        contentType: false,
                        success: function (data, status) {
                            //this will execute when form is submited without errors
                            alert("Suc")
                        },
                        error: function (data, status) {
                            //this will execute when get any error
                            alert(status)
                        },
                    });
                }); */


            });

       // }
    }

    function UpdateProgressbar(doneIndex) {

        var percentDone = doneIndex / totalExpenesToMake
        percentDone *= 100

        $("#progressbar").css("width", percentDone+"%")
        

      //  alert("Done " + doneIndex + " Ud af " + totalExpenesToMake + " Percent " + percentDone)

        if (percentDone >= 100) {
            $("#dietsform").submit();
           // location.reload();
        }

    }


    // Diets

    var dagsfradrag = $("#dagsfradrag").val();
    var morgenmadBrugt = 0
    var frokostBrugt = 0
    var aftensmadBrugt = 0

    $("#antalmorgenmad").keyup(function () {
        var morgenmadPrice = $("#morgenmadprice").val();
        thisval = $(this).val();
        morgenmadBrugt = morgenmadPrice * thisval;
        morgenmadBrugt = parseFloat(morgenmadBrugt)
        UdregnTotalDiet();
    });

    $("#antalfrokost").keyup(function () {
        var forkostPrice = $("#frokostprice").val();
        thisval = $(this).val();
        frokostBrugt = forkostPrice * thisval;
        frokostBrugt = parseFloat(frokostBrugt)
        UdregnTotalDiet();
    });

    $("#antalaftensmad").keyup(function () {
        var aftensmadPrice = $("#aftensmadprice").val();
        thisval = $(this).val();
        aftensmadBrugt = aftensmadPrice * thisval;
        aftensmadBrugt = parseFloat(aftensmadBrugt)
        UdregnTotalDiet();
    });

    function UdregnTotalDiet() {
        var total = dagsfradrag - (morgenmadBrugt + frokostBrugt + aftensmadBrugt)

        $("#dietTotal").html(total)
    }



});

