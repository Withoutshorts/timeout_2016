




$(document).ready(function () {

    $('.date').datepicker({

    });

    if ($("#bruguge").is(':checked') == true) {
        document.getElementById('bruguge_selector').style.display = "";
        document.getElementById('bruguge_selector_of').style.display = "none";
    } else {
        document.getElementById('bruguge_selector').style.display = "none";
        document.getElementById('bruguge_selector_of').style.display = "";
    }

    $(".dateweeknum").click(function () {

        if ($("#bruguge").is(':checked') == true) {
            document.getElementById('bruguge_selector').style.display = "";
            document.getElementById('bruguge_selector_of').style.display = "none";
        } else {
            document.getElementById('bruguge_selector').style.display = "none";
            document.getElementById('bruguge_selector_of').style.display = "";
        }

    });


    $(".medarbliste").click(function () {

        thisid = this.id
        var medarbTR = document.getElementById('tr_medarbliste_' + thisid)
        var medarbheader = document.getElementById('aktvisfont_' + thisid);

        if (medarbTR.style.display !== 'none') {
            medarbTR.style.display = 'none'
            medarbheader.style.color = ""
        } else {
            medarbTR.style.display = ''
            medarbheader.style.color = "black"
        }

    });





    $("#godkendalle").click(function () {

        //alert("alle"+ this + $(".godkendalle").attr("checked"))

        $(".godkendbox").each(function () {

            //alert("HER 2" + this.id)
            //$(this).attr('checked', true);

            if ($(this).prop('disabled') != true) {

                if ($("#godkendalle").is(':checked') == true) {

                    $(this).prop("checked", true)
                } else {
                    $(this).prop("checked", false)
                }
            }

        });

        if ($("#godkendalle").is(':checked') == true) {
            $("#afvisalle").prop("checked", false)
            $(".afvismedarb").prop("checked", false)
        }


        //afvisgodkendAll();

    });


    $(".afvisbox").click(function () {

        $(".sendemail").hide('fast')
        
        $(".noemail").css("display","");
        $(".noemail").css("visibility","visible")

        

    });

    $(".godkendbox").click(function () {

        $(".sendemail").hide('fast')

        $(".noemail").css("display", "");
        $(".noemail").css("visibility", "visible")



    });



    $("#afvisalle").click(function () {

        //alert("alle"+ this + $(".godkendalle").attr("checked"))

        $(".afvisbox").each(function () {

            //alert("HER 2" + this.id)
            //$(this).attr('checked', true);

            if ($(this).prop('disabled') != true) {

                if ($("#afvisalle").is(':checked') == true) {

                    $(this).prop("checked", true)
                    $("#godkendalle").prop("checked", false)
                } else {
                    $(this).prop("checked", false)
                }

            }

        });


        if ($("#afvisalle").is(':checked') == true) {
            $("#godkendalle").prop("checked", false)
            $(".godkendmedarb").prop("checked", false)

        }


        //afvisgodkendAll();

    });




    $(".godkendmedarb").click(function () {



        thisid = this.id
        var arr = thisid.split('_');

        var medid = arr[1];


        $(".godkendbox_" + medid).each(function () {

            if ($(this).prop('disabled') != true) {

                if ($("#" + thisid).is(':checked') == true) {

                    $(this).prop("checked", true)
                } else {
                    $(this).prop("checked", false)
                }

            }

        });

        if ($("#" + thisid).is(':checked') == true) {
            $("#medarbafvis_" + medid).prop("checked", false)
        }


        //afvisgodkendAll();
    });



    $(".afvismedarb").click(function () {



        thisid = this.id
        var arr = thisid.split('_');

        var medid = arr[1];


        $(".afvisbox_" + medid).each(function () {


            if ($(this).prop('disabled') != true) {

                if ($("#" + thisid).is(':checked') == true) {

                    $(this).prop("checked", true)
                } else {
                    $(this).prop("checked", false)
                }

            }
        });

        if ($("#" + thisid).is(':checked') == true) {
            $("#medarbgodkend_" + medid).prop("checked", false)
        }

        //afvisgodkendAll();
    });


    

    $("#godkendknap").click(function () {

        afvisgodkendAll();

        $("#godkendform").submit();

    });

    /// Submitter til DB

    //$("#XXgodkendknap").click(function () {


    //$(".godkendmedarb, .afvismedarb, .godkendalle, .afvisalle").click(function () {
    function afvisgodkendAll() {


        tids = "";
        $(".godkendbox").each(function () {



            var thisid = this.id
            var arr = thisid.split('_');


            var medid = arr[1];
            var aktid = arr[2];
            var tid = arr[3];



            if ($("#" + thisid).is(':checked') == true) {


                if (tid > 0) {
                    tids += ", " + tid
                }

            }

        });

        //alert("TIDS: " + tids)


     
        $.post("?multi=1&tid=" + tids, { control: "godkenduge", AjaxUpdateField: "true" }, function (data) {

            //$("#div_jobid").html(data);


        });



        globalMid = 0;
        globalTids = "";
        lastMid = 0;
        tids = "";


        $(".afvisbox").each(function () {

            //alert("HEj")

            var thisid = this.id;

            var arr = thisid.split('_');

            var medid = arr[1];
            var aktid = arr[2];
            var tid = arr[3];


            if ($("#" + thisid).is(':checked') == true) {


                if (tid > 0) {
                    tids += ", " + tid
                }

            }

            //alert(tids)

        });



        //alert("afvist");

        $.post("?multi=1&tid=" + tids, { control: "afvisuge", AjaxUpdateField: "true" }, function (data) {

            //alert("HEJ: " + data)
            //$("#div_jobid").html(data);

            //$("#div_jobid").html(data);

        });

        /// SEND IKKE MAIL VIA DENNE MERE 20171003
        t = 0
        if (t != 0) {

            $(".decl_tids").each(function () {

                //alert("HEr 9")

                var thisMid = $(this).val();

                //alert(thisMid)
                var ThisTids = $("#decl_tids_" + thisMid).val();
                var kommentar = $("#decline_comment_" + thisMid).val()
                $("#decl_tids_" + thisMid).val('');

                if (ThisTids != "") {
                    // alert("HER 8 decl_tids: " + ThisTids + " mid: " + thisMid)

                    //$("#decl_tids_mid_" + globalMid).val('');
                    //alert("KØR DDD")

                    $.post("?medid=" + thisMid + "&tids=" + ThisTids + "&kommentar=" + kommentar, { control: "emailnoti", AjaxUpdateField: "true" }, function (data) {

                        //alert("godkendt");

                        //$("#div_jobid").html(data);

                    });

                }

            });
        };



        // $("#godkendform").submit();

    }
    //});



    $('.sendemail').hover(function () {
        $(this).css('cursor', 'pointer');
    });



    $(".sendemail").click(function () {



        var thisid = this.id;
        var arr = thisid.split('_');
        var thisMid = arr[1];
        var thisAkt = arr[2];

        //alert(thisMid + "_" + thisAkt)

        var ThisTids = $("#decl_tids_" + thisMid + "_" + thisAkt).val();

        //alert("ThisTids: " +ThisTids)


        var kommentar = $("#decline_comment_" + thisMid + "_" + thisAkt).val()
        //$("#decl_tids_" + thisMid + "_" + thisAkt).val('');

        if (ThisTids != "") {
            // alert("HER 8 decl_tids: " + ThisTids + " mid: " + thisMid)

            //$("#decl_tids_mid_" + globalMid).val('');
            //alert("KØR DDD")

            $.post("?medid=" + thisMid + "&tids=" + ThisTids + "&kommentar=" + kommentar, { control: "emailnoti", AjaxUpdateField: "true" }, function (data) {

                //alert("godkendt");

                //$("#div_jobid").html(data);

            });

        } else {
            alert("No work was done - You need to submit declined info before sending email");
        };


        $("#decline_comment_" + thisMid + "_" + thisAkt).val('')
        $("#sendemail_" + thisMid + "_" + thisAkt).html('Mail send!');

        
       
        var time = 1;

        var interval = setInterval(function() { 
            if (time <= 1) { 
                //alert(time);
                $("#sendemail_" + thisMid + "_" + thisAkt).html('Send email with declined info >>');
                time++;
            }
            else { 
                clearInterval(interval);
            }
        }, 5000);



    });



    //$(".godkendbox").each(function () {
    $(".Xgodkendbox").click(function () {

        //alert("HER GK DISABLE")

        var thisid = this.id
        var arr = thisid.split('_');


        var medid = arr[1];
        var aktid = arr[2];
        var tid = arr[3];



        if ($("#" + thisid).is(':checked') == true) {

            startdato = $("#startdato").val()
            slutdato = $("#slutdato").val()
            godkendjobid = $("#godkendjobid").val()



            //alert(startdato + slutdato + "medid:" + medid + "aktid: " + aktid + "godkendjobid:" + godkendjobid + " tid: " + tid)
            //return true

            $.post("?tid=" + tid + "&medid=" + medid + "&startdato=" + startdato + "&slutdato=" + slutdato + "&aktid=" + aktid + "&godkendjobid=" + godkendjobid, { control: "godkenduge", AjaxUpdateField: "true" }, function (data) {

                //alert("godkendt");

                //$("#div_jobid").html(data);

            });

        }

    });



    $(".Xafvisbox").click(function () {


        var thisid = this.id;

        var arr = thisid.split('_');

        var medid = arr[1];
        var aktid = arr[2];
        var tid = arr[3];

        if ($("#" + thisid).is(':checked') == true) {

            startdato = $("#startdato").val()
            slutdato = $("#slutdato").val()
            godkendjobid = $("#godkendjobid").val()


            //alert(startdato + slutdato + "tid: " + tid + " medid:" + medid + "aktid: " + aktid + "godkendjobid:" + godkendjobid)
            //return true

            decl_tids = $("#decl_tids_" + medid).val();
            //alert(decl_tids)
            $("#decl_tids_" + medid).val(decl_tids + " OR tid = " + tid);
            //$("#decl_tids_mid_" + medid).val(medid);
            globalMid = medid;
            globalTids = $("#decl_tids_" + globalMid).val();

            $.post("?tid=" + tid + "&medid=" + medid + "&startdato=" + startdato + "&slutdato=" + slutdato + "&aktid=" + aktid + "&godkendjobid=" + godkendjobid, { control: "afvisuge", AjaxUpdateField: "true" }, function (data) {

                //alert("afvist");

                //$("#div_jobid").html(data);


            });



            //lastMid = medid

        } // CHECKED = TRUE


    });




  /*  $(".godkendmedarb_akt").click(function () {

        var thisid = this.id
        var arr = thisid.split('_');


        var medid = arr[1];
        var aktid = arr[2];
        var tid = arr[3];

        alert(medid)
        alert(aktid)
        alert(tid)

        startdato = $("#startdato").val()
        slutdato = $("#slutdato").val()
        godkendjobid = $("#godkendjobid").val()

        //alert(startdato)
        //alert(slutdato)
        //alert(godkendjobid)

        $.post("?tid=" + tid + "&medid=" + medid + "&startdato=" + startdato + "&slutdato=" + slutdato + "&aktid=" + aktid + "&godkendjobid=" + godkendjobid, { control: "godkenduge", AjaxUpdateField: "true" }, function (data) {

            alert("godkendt2");

            //$("#div_jobid").html(data);

        });


    }); */



    $(".aktivilist").click(function () {

        //alert("hej")


        var modalid = this.id
        //var idlngt = modalid.length
        //var idtrim = modalid.slice(6, idlngt)

        //var modalidtxt = $("#myModal_" + idtrim);

        var modal = document.getElementById('aktivilist_' + modalid);
        var medarbheader = document.getElementById('aktvisfont_' + modalid);
        //var modal = document.getElementsByClassName("aktivilist_" + modalid)


        //alert("awd");

        if (modal.style.display !== 'none') {
            modal.style.display = 'none';
            medarbheader.style.color = ""
            //alert("normal")
        }
        else {
            modal.style.display = '';
            medarbheader.style.color = "black"
            //alert("none")
        }

    });

    /* if (modal.style.display !== 'none') {
         modal.style.display = 'none';
     }
     else {
         modal.style.display = 'block';
     } */



    //alert("her")

});



