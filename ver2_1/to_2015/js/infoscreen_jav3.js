

$(document).ready(function () {


    startTime();

    
    if ($(document).width() > 2500) {
       // alert("3 r;kker 3 3")
        Cookies.set('infoscreenSize', '1');
    }
    else {
      //  alert("2 r;kker")
        Cookies.set('infoscreenSize', '0');
    }

    $("#alarm").bind('click', function () {
       // alert("awd")
        lto = $("#lto").val();

        if (lto == "cflow") {
            $.post("?medid=1", { control: "GetEmployeeTel", AjaxUpdateField: "true" }, function (data) {

                recipient = data.split(",")
               // alert(recipient.length)

                recipients = ""
                for (var i = 0; i < recipient.length; i++) {
                    if (i == 0) {
                        recipients += "recipients.0.msisdn=" + recipient[0]
                    } else {
                        recipients += "&recipients." + i + ".msisdn=" + recipient[i]
                    }
                }
              //  alert(recipients)

                messege = "Test sms fra timeout - undskyld forstyrelsen"
               // var smswin = window.open("https://gatewayapi.com/rest/mtsms?token=gAws4-yJSP-ySXOjoDRNcgKvWgTzAbnB1nFEuOdPusswBtRWDQRqiHk35uE8rdYO&message=" + messege + "&" + recipients)
               // smswin.close();

                $.post("https://gatewayapi.com/rest/mtsms?token=gAws4-yJSP-ySXOjoDRNcgKvWgTzAbnB1nFEuOdPusswBtRWDQRqiHk35uE8rdYO&message=" + messege + "&" + recipients + messege + "&recipients.0.msisdn=" + recipient, function (data) {

                });

            });
        }

        if (lto == "cool") {
           $.post("?medid=1", { control: "Mail", AjaxUpdateField: "true" }, function (data) {
                alert(data)
            });
        }

    });


    $('.date').datepicker({

    });




    $(".container").bind('click', function () {

        $("#keyfield").focus();

    });

    $("#keyfield").change(function () {

        var key = $("#keyfield").val();

        $("#scan").submit();

    });


    var idToEdit = 0;
    $(".context-menu-one").bind('mousedown', function () {
        thisid = this.id;
        idToEdit = thisid;

        // alert(idToEdit);
    });

    $(".context-menu-two").bind('mousedown', function () {
        thisid = this.id;
        idToEdit = thisid;

        // alert(idToEdit);
    });

    var forrejsetxt = $("#rightclick_txt_forrejse").val()
    var arbejdhjemmetxt = $("#rightclick_txt_arbejdhjemme").val()
    var ferietxt = $("#rightclick_txt_ferie").val()
    var sygTxt = $("#rightclick_txt_syg").val()
    var guestFrokostTxt = $("#rightclick_txt_guestfrokost").val()
    // alert(ferietxt + sygTxt + guestFrokostTxt)
    $.contextMenu({
        selector: '.context-menu-one',
        callback: function (key, options) {
            var m = "clicked: " + key;

          //  window.console && console.log(m) || alert(m + " " + idToEdit);
            UploadHours(key, idToEdit);
        },
        items: {
            "forrejse": { name: forrejsetxt, icon: "edit" },
            "arbejdhjemme": { name: arbejdhjemmetxt, icon: "edit" },
            "ferie": { name: ferietxt, icon: "edit" },
            "syg": { name: sygTxt, icon: "edit" }
        }
    });

    $('.context-menu-one').on('click', function (e) {
        console.log('clicked', this);
    }) 

    $.contextMenu({
        selector: '.context-menu-two',
        callback: function (key, options) {
            var m = "clicked: " + key;

          //  window.console && console.log(m) || alert(m + " " + idToEdit);
            AddGuestToLunch(idToEdit);
        },
        items: {
            "frokost": { name: guestFrokostTxt.toString(), icon: "edit" }
        }
    });

    $('.context-menu-two').on('click', function (e) {
        console.log('clicked', this);
    }) 

    function UploadHours(key, medid) {

     //   alert(key + " medid " + medid);
        aktid = 0 // aktid er hardcoded ind til videre. her under bestemes hvilket aktivitet timerne skal ud på.
        switch (key)
        {
            case "ferie":
               // aktid = '13';
                DatoPopUp(medid, key);
                break;

            case "syg":
                aktid = '30'; 
                break;

            case "arbejdhjemme":
                aktid = '32';
                break;

            case "forrejse":
                // aktid = '33'
                DatoPopUp(medid, key);
                break;
        }

        if (key == "syg" || key == "arbejdhjemme") {
        
            $.post("?aktid=" + aktid + "&medid=" + medid, { control: "UploadHoursToAkt", AjaxUpdateField: "true" }, function (data) {

                //   alert("cc");
                location.reload();
            });
        }

    }

    function DatoPopUp(medid, key)
    {
        $('#datoPopMedid').val(medid);
        $('#datoPopKey').val(key);
        $('#datopop').show();
    }

    $("#datoPopUpGodkend").bind('click', function () {

        startdato = $("#datoPopSt").val();
        slutdato = $("#datoPopSl").val();
        medid = $('#datoPopMedid').val();
        key = $('#datoPopKey').val();

        //  alert("medid " + medid + " key " + key + " std " + startdato + " stl " + slutdato)

        if (startdato == "" || slutdato == "") {
            alert("Angiv start og slutdato")
        }
        else {

            $.post("?medid=" + medid + "&regkey=" + key + "&startdato=" + startdato + "&slutdato=" + slutdato, { control: "DatoPopReg", AjaxUpdateField: "true" }, function (data) {
                location.reload();
            });

        }   

    });

    $("#datoPopUpAnnulller").bind('click', function () {
        $('#datopop').hide();
    });

    function AddGuestToLunch(medid) {
       // alert("Hep")
        $.post("?medid=" + medid, { control: "AddGuestToLunch", AjaxUpdateField: "true" }, function (data) {

          //  alert("cc");
            location.reload();

        });
    }









    $("#sort_name").click(function () {

        sortTable(0);

    });

    $("#sort_color").click(function () {

        sortTable(1);

    });


    function sortTable(sortTd) {
        var table, rows, switching, i, x, y, shouldSwitch;
        table = document.getElementById("myTable");
        switching = true;
        /*Make a loop that will continue until
        no switching has been done:*/
        while (switching) {
            //start by saying: no switching is done:
            switching = false;
            rows = table.getElementsByTagName("TR");
            /*Loop through all table rows (except the
            first, which contains table headers):*/
            for (i = 1; i < (rows.length - 1); i++) {
                //start by saying there should be no switching:
                shouldSwitch = false;
                /*Get the two elements you want to compare,
                one from current row and one from the next:*/
                x = rows[i].getElementsByTagName("TD")[sortTd];
                y = rows[i + 1].getElementsByTagName("TD")[sortTd];
                //check if the two rows should switch place:
                if (x.innerHTML.toLowerCase() > y.innerHTML.toLowerCase()) {
                    //if so, mark as a switch and break the loop:
                    shouldSwitch = true;
                    break;
                }
            }
            if (shouldSwitch) {
                /*If a switch has been marked, make the switch
                and mark that a switch has been done:*/
                rows[i].parentNode.insertBefore(rows[i + 1], rows[i]);
                switching = true;
            }
        }
    }


    function startTime() {
        var today = new Date();
        var h = today.getHours();
        var m = today.getMinutes();
        var s = today.getSeconds();
        m = checkTime(m);
        s = checkTime(s);
        document.getElementById('live_clock').innerHTML =
            h + ":" + m + ":" + s;
        var t = setTimeout(startTime, 500);
    }

    function checkTime(i) {
        if (i < 10) { i = "0" + i };  // add zero in front of numbers < 10
        return i;
    }




});


