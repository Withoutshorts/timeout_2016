

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
        if (lto != "lm") {
            strExpenceSection += "<tr>"
            strExpenceSection += "<td colspan='2'>"
            strExpenceSection += "<select class='form-control' id='FM_udlaeg_faktbar_" + expenceTotal + "'>"

            if (lto == "nt") {
                strExpenceSection += "<option value='1'>" + notbillableTxt + "</option>"
                strExpenceSection += "<option value='2'>" + billableTxt + " 100%</option>"
                strExpenceSection += "<option value='5'>" + billableTxt + " 50%</option>"
            } else {
                strExpenceSection += "<option value='1'>" + notbillableTxt + "</option>"
                strExpenceSection += "<option value='2'>" + billableTxt + "</option>"
            }

            strExpenceSection += "</select>"
            strExpenceSection += "</td>"
            strExpenceSection += "</tr>"
        } else {
            strExpenceSection += "<input type='hidden' id='FM_udlaeg_faktbar_" + expenceTotal + "' value='1' />"
        }
        


        // Personlig elelr firma
        strExpenceSection += "<tr>"
        strExpenceSection += "<td colspan='2'>"
        strExpenceSection += "<select class='form-control' id='FM_udlaeg_form_" + expenceTotal + "'>"
        if (lto == "nt") {
            strExpenceSection += "<option value='0'>" + companypaidTxt +"</option>"
            strExpenceSection += "<option value='1'>" + personalTxt +"</option>"
        }
        else {
            strExpenceSection += "<option value='1'>" + personalTxt + "</option>"

            if (lto != "lm") {
                strExpenceSection += "<option value='0'>" + companypaidTxt + "</option>"
            }
        }

        if (lto == "lm") {
            strExpenceSection += "<option value='20'>M.Card</option>"
            strExpenceSection += "<option value='21'>Workplus</option>"
            strExpenceSection += "<option value='22'>BroBizz</option>"
            strExpenceSection += "<option value='23'>Rejsekort</option>"
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



    var ARRAY_udlaeg_navn = [];
    var ARRAY_udlaeg_belob = [];
    var ARRAY_udlaeg_valuta = [];
    var ARRAY_udlaeg_gruppe = [];
    var ARRAY_udlaeg_faktbar = [];
    var ARRAY_udlaeg_form = [];
    var ARRAY_udlaeg_valutakode = [];


    $("#sbmExpence").click(function () {
        
        strStdatoDiet = $("#strStdatoDiet").val();
        strSldatoDiet = $("#strSldatoDiet").val();

        

        
        if (strStdatoDiet != "" && strSldatoDiet != "") {

            FM_diet_sttime = $("#FM_diet_sttime").val();
            FM_diet_sltime = $("#FM_diet_sltime").val();

            if (FM_diet_sttime == "" || FM_diet_sltime == "") {
                $("#errorMessage").text("Indtast venligt til og fratidspunkt på diæter")
                window.scrollTo(0, 0);
                return;
            }
        }

        kmantal = $("#antalkm").val();
        kmsted = $("#kmregsted").val();
        kmfra = $("#fradistination").val();
        kmtil = $("#tildistination").val();
        kmretur = $("#returdistination").val();
        if (kmantal != "" && kmantal != "0") {
            if (kmsted == "" || kmfra == "" || kmtil == "" || kmretur == "") {
                $("#errorMessage").text("Udfyld venligst alle felter ved KM registrering")
                window.scrollTo(0, 0);
                return;
            }
        }

        SubmitTheRest();

            // Nedenstående er hvis den skal tjekke om der er for mange timer på dagen til at man kan oprette diæter]
           /* $.post("?strStdatoDiet=" + strStdatoDiet + "&strSldatoDiet=" + strSldatoDiet, { control: "CheckHoursInDietPeriod", AjaxUpdateField: "true" }, function (data) {
                
                if (data == 1 || data == "1") {
                    $("#errorMessage").text("Der kan ikke indtastes diæter på dage der er registreret mere end 10 timer")
                    window.scrollTo(0, 0);
                }
                else {
                    SubmitTheRest();
                }

            });
        }
        else {
            SubmitTheRest();
        } */
    });

    function SubmitTheRest() {

        if (expenceTotal == 1 && $("#FM_udlaeg_navn_0").val() == "") {
            UploadOvernatninger();
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
                window.scrollTo(0, 0);
            }

            if (lto != "lm") {

                if (FM_udlaeg_belob == "") {
                    error = 1
                    $("#errorMessage").text("Der mangler at blive angivet et beløb")
                    window.scrollTo(0, 0);
                }

            } else {

                if (FM_udlaeg_form == "1" && FM_udlaeg_belob == "") {
                    error = 1
                    $("#errorMessage").text("Der mangler at blive angivet et beløb")
                    window.scrollTo(0, 0);
                }

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

        if (error == 0) {

            $("#uploadbar").show();

            // var udlaeg_navnne = ['Hans', 'Jens', 'Morten'];
            // alert(udlaeg_navnne[0])

            /* for (var i = 0; i < expenceTotal; i++) {
 
                 
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
 
                 
                 
             } */

            CallAjax()

        }

    }

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

        if (udlaeg_belob == "" || udlaeg_belob == "0") {
            amount = 1
        }
        else {
            amount = udlaeg_belob
        }
        
        from = udlaeg_fromvalutakode

        $.ajax({
            url: 'https://apilayer.net/api/' + endpoint + '?access_key=' + access_key + '&from=' + from + '&to=' + to + '&amount=' + amount + "&format=1",
            dataType: 'jsonp',
            success: function (json) {
                
                var kurs = parseFloat(json.result) / parseFloat(amount)

                if (udlaeg_belob == "" || udlaeg_belob == "0") {
                    jsonresult = 0
                    udlaeg_belob = 0
                }
                else {
                    jsonresult = json.result
                }

                //  alert("Result " + json.result + " KURS " + kurs)
                //  alert("Navn " + udlaeg_navn)
                //  alert("Aj " + to + " " + kurs + " from " + from + " " + json.result + " " + FM_udlaeg_navn + " " + FM_udlaeg_belob + " " + FM_udlaeg_valuta + " " + FM_udlaeg_gruppe + " " + FM_udlaeg_faktbar + " " + FM_udlaeg_form + " " + FM_datoer + " " + jq_jobid + " " + aktid)

              //  alert("Navn " + udlaeg_navn + " belob " + udlaeg_belob + " valuta " + udlaeg_valuta + " udlaeg_gruppe " + udlaeg_gruppe + " udlaeg_faktbar " + udlaeg_faktbar + " udlaeg_form " + udlaeg_form)

                IndlaesExpence(to, kurs, jsonresult, udlaeg_navn, udlaeg_belob, udlaeg_valuta, udlaeg_gruppe, udlaeg_faktbar, udlaeg_form, FM_datoer, jq_jobid, aktid, i);

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
                  
            });

       // }
    }

    function UpdateProgressbar(doneIndex) {

        var percentDone = doneIndex / totalExpenesToMake
        percentDone *= 100

        $("#progressbar").css("width", percentDone+"%")
        

      //  alert("Done " + doneIndex + " Ud af " + totalExpenesToMake + " Percent " + percentDone)

        if (percentDone >= 100) {

            UploadOvernatninger();

           // location.reload();
        }

    }


    // KM
    $("#antalkm").keyup(function () {
        thisval = thisval = $(this).val();
        thisval = thisval.replace(",", ".")
        prisprkm = $("#prisprkm").val();
        prisprkm = prisprkm.replace(",", ".")

        prisprkm = parseFloat(prisprkm)
        thisval = parseFloat(thisval)

        var kmpristot = 0

        if (prisprkm > 0 && thisval > 0) {
            kmpristot = prisprkm * thisval
        }

        $("#kmtotalpris").html(kmpristot + " Kr.");

    });

    // Diets
    var dagsfradrag = $("#dagsfradrag").val();
    var morgenmadBrugt = 0
    var frokostBrugt = 0
    var aftensmadBrugt = 0

    $("#antalmorgenmad").keyup(function () {
        var morgenmadPrice = $("#morgenmadprice").val();
        morgenmadPrice = morgenmadPrice.replace(",", ".")
        thisval = $(this).val();
        thisval = thisval.replace(",", ".")
        morgenmadBrugt = morgenmadPrice * thisval;
        morgenmadBrugt = parseFloat(morgenmadBrugt)
        UdregnTotalDiet();
    });

    $("#antalfrokost").keyup(function () {
        var forkostPrice = $("#frokostprice").val();
        forkostPrice = forkostPrice.replace(",", ".")
        thisval = $(this).val();
        thisval = thisval.replace(",", ".")
        frokostBrugt = forkostPrice * thisval;
        frokostBrugt = parseFloat(frokostBrugt)
        UdregnTotalDiet();
    });

    $("#antalaftensmad").keyup(function () {
        var aftensmadPrice = $("#aftensmadprice").val();
        aftensmadPrice = aftensmadPrice.replace(",", ".")
        thisval = $(this).val();
        thisval = thisval.replace(",", ".")
        aftensmadBrugt = aftensmadPrice * thisval;
        aftensmadBrugt = parseFloat(aftensmadBrugt)
        UdregnTotalDiet();
    });


    $("#strStdatoDiet").change(function () {
        UdregnTotalDiet();
    });

    $("#strSldatoDiet").change(function () {
        UdregnTotalDiet();
    });

    $("#FM_diet_sttime").change(function () {
        UdregnTotalDiet();
    });

    $("#FM_diet_sltime").change(function () {
        UdregnTotalDiet();
    });

    function UdregnTotalDiet() {
     
        strStdatoDiet = $("#strStdatoDiet").val();
        strSldatoDiet = $("#strSldatoDiet").val();

        if (strStdatoDiet != "" && strSldatoDiet != "") {
            
            std = strStdatoDiet.split("-")
            stdArr = std[2]
            stdMonth = (std[1] - 1)
            stdDay = std[0]

            sld = strSldatoDiet.split("-")
            sldArr = sld[2]
            sldMonth = (sld[1] - 1)
            sldDay = sld[0]

            FM_diet_sttime = $("#FM_diet_sttime").val();
            if (FM_diet_sttime != "") {
                startime = FM_diet_sttime.split(":")
                startimeHour = startime[0];
                startimeMinute = startime[1];
            } else {
                startimeHour = 0
                startimeMinute = 0
            }

            FM_diet_sltime = $("#FM_diet_sltime").val();
            if (FM_diet_sltime != "") {
                sluttime = FM_diet_sltime.split(":")
                sluttimehour = sluttime[0];
                sluttimeminute = sluttime[1];
            } else {
                sluttimehour = 0
                sluttimeminute = 0
            }


            var startDate = new Date();
            startDate.setFullYear(stdArr)
            startDate.setMonth(stdMonth)
            startDate.setDate(stdDay)
            startDate.setHours(startimeHour)
            startDate.setMinutes(startimeMinute)

            var endDate = new Date();
            endDate.setFullYear(sldArr)
            endDate.setMonth(sldMonth)
            endDate.setDate(sldDay)
            endDate.setHours(sluttimehour)
            endDate.setMinutes(sluttimeminute)

            var seconds = (endDate.getTime() - startDate.getTime()) / 1000;
            days = ((seconds / 60) / 60) / 24
            
            if (days < 1) {
                $('#antalmorgenmad').prop('readonly', true);
                $('#antalfrokost').prop('readonly', true);
                $('#antalaftensmad').prop('readonly', true);

                $('#antalmorgenmad').val(0);
                $('#antalfrokost').val(0);
                $('#antalaftensmad').val(0);
            }
            else {
                $('#antalmorgenmad').prop('readonly', false);
                $('#antalfrokost').prop('readonly', false);
                $('#antalaftensmad').prop('readonly', false);
            }


        } else {
            days = 1
        }


        totaldagsfradrag = 0
        totaldagsfradrag = days * dagsfradrag

       // alert(days + " " + totaldagsfradrag)
       // alert(morgenmadBrugt + " " + frokostBrugt + " " + aftensmadBrugt)
        var total = totaldagsfradrag - (morgenmadBrugt + frokostBrugt + aftensmadBrugt)
        total = total.toFixed(2);
        
        $("#dietTotal").html(total)
    }


    // Overnatningner
    $("#antalovernatninger").keyup(function () {
        overnatninger = $("#antalovernatninger").val();
        prisprovernatning = $("#prisprovernatning").val();

        overnatninger.replace(",", ".")
        prisprovernatning.replace(",", ".")

        overnatninger = parseFloat(overnatninger)
        prisprovernatning = parseFloat(prisprovernatning)

        totalPrisForOvernatninger = overnatninger * prisprovernatning

        $("#overnatningerTotalPrice").html(totalPrisForOvernatninger + " kr.");

       // alert(overnatninger + " " + prisprovernatning)
        
    });

    function UploadOvernatninger() {

        prisprovernatning = $("#prisprovernatning").val();
        overnatninger = $("#antalovernatninger").val();

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

        FM_udlaeg_navn = "Overnatning"
        FM_udlaeg_valuta = 1
        FM_udlaeg_form = 0
        FM_datoer = $("#FM_datoer").val();
        FM_udlaeg_gruppe = 0
        basic_valuta = "DKK"
        basic_kurs = 1
        basic_belob = prisprovernatning
        FM_udlaeg_faktbar = 1
        antal = parseInt(overnatninger)
    

        if (antal > 0) {

            if (jq_jobid == 0 || jq_jobid == -1) {
                error = 1
                $("#errorMessage").text("Der mangler at blive valgt et job")
                window.scrollTo(0, 0);
                return;
            }

            $.post("?jq_jobid=" + jq_jobid + "&aktid=" + aktid + "&FM_udlaeg_valuta=" + FM_udlaeg_valuta + "&FM_udlaeg_form=" + FM_udlaeg_form + "&FM_udlaeg_belob=" + prisprovernatning + "&FM_udlaeg_navn=" + FM_udlaeg_navn + "&FM_udlaeg_gruppe=" + FM_udlaeg_gruppe + "&FM_datoer=" + FM_datoer + "&basic_valuta=" + basic_valuta + "&basic_kurs=" + basic_kurs + "&basic_belob=" + basic_belob + "&FM_udlaeg_faktbar=" + FM_udlaeg_faktbar + "&antal=" + antal, { control: "FN_createOutlay", AjaxUpdateField: "true" }, function (data) {
                UploadKmReg();
            });

        } else {
            UploadKmReg();
        }
        

    }


    $("#fradistination").click(function () {

        $("#dv_fradistination").show();
        
    });

    $("#tildistination").click(function () {

        $("#dv_tildistination").show();

    });

    $("#returdistination").click(function () {

        $("#dv_returdistination").show();

    });

    var bopalfra = 0;
    var bopaltil = 0;
    var bopalretur = 0;

    $(".fradistination_forvalgt").click(function () {

        thisid = this.id
        if (thisid == "bopal") {
            bopalfra = 1
        } else {
            bopalfra = 0
        }

        thisadress = $(this).html();
        $("#fradistination").val(thisadress);
        $("#dv_fradistination").hide();

    });

    $(".tildistination_forvalgt").click(function () {

        thisid = this.id
        if (thisid == "bopal") {
            bopaltil = 1
        } else {
            bopaltil = 0
        }

        thisadress = $(this).html();
        $("#tildistination").val(thisadress);
        $("#dv_tildistination").hide();

    });

    $(".returdistination_forvalgt").click(function () {

        thisid = this.id
        if (thisid == "bopal") {
            bopalretur = 1
        } else {
            bopalretur = 0
        }

        thisadress = $(this).html();
        $("#returdistination").val(thisadress);
        $("#dv_returdistination").hide();

    });


    $("#fradistination").change(function () {

        $("#dv_fradistination").hide();
        
    });

    $("#tildistination").change(function () {

        $("#dv_tildistination").hide();

    });

    $("#returdistination").change(function () {

        $("#dv_returdistination").hide();

    });

    function UploadKmReg() {

        antalkm = $("#antalkm").val();       
        sted = $("#kmregsted").val();
        fraDistination = $("#fradistination").val();
        tilDistination = $("#tildistination").val();
        returDistination = $("#returdistination").val();
        usebopal = 0
        if (bopalfra == 1 || bopaltil == 1 || bopalretur == 1) {
            usebopal = 1
        }

        if (sted != "") {
            sted = "Formål og sted: " + sted
        }
       
        if (fraDistination != "") {
            fraDistination = "Fra: " + fraDistination
        }

        if (tilDistination != "") {
            tilDistination = "Til: " + tilDistination
        }

        if (returDistination) {
            returDistination = "Retur: " + returDistination
        }

        distination = sted + " " + fraDistination + " " + tilDistination + " " + returDistination

        koregnr = $("#koregnr").val();
        kmdato = $("#FM_datoer").val();

        if (mobil_week_reg_job_dd != 1) {
            jq_jobid = $("#FM_jobid").val();
        } else {
            jq_jobid = $("#dv_job").val();
        }


        if (antalkm != "") {
            $.post("?antalkm=" + antalkm + "&distination=" + distination + "&koregnr=" + koregnr + "&kmdato=" + kmdato + "&jobid=" + jq_jobid + "&usebopal=" + usebopal, { control: "KmReg", AjaxUpdateField: "true" }, function (data) {
                $("#dietsform").submit();
            });
        }
        else {
            $("#dietsform").submit();
        }
    }



});

