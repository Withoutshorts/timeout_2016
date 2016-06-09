

/////////// Søg job og kunder hoved søge filter ///////////////



//alert("herher")

$(window).load(function () {
    // run code

    //load_cdown


   




});






$(document).ready(function () {



    $(".a_showkom").click(function () {


        //alert("her")

        var a = $(this);
        var pos = a.position()
        var left = pos.left
        var top = pos.top
        //alert(top)
        $("#kom").css("top", top + 100);
        $("#kom").css("left", "100");

        $.scrollTo("#kom", 300, { offset: -200 });
        $(".FM_kommentar").blur();
        $(".FM_kommentar").focus();

    });



    $(".afase").click(function () {
        idval = this.id
        //alert(idval)
        vzb = $(".td_" + idval + "").css("visibility")
        if (vzb == "hidden") {
            $(".td_" + idval + "").css("display", "");
            $(".td_" + idval + "").css("visibility", "visible");

            $("#faseshow_" + idval + "").css("visibility", "visible");

            document.getElementById("faseshow_" + idval + "").value = "visible"

            $.scrollTo(this, 300, { offset: -200 });

        } else {
            $(".td_" + idval + "").css("display", "none");
            $(".td_" + idval + "").css("visibility", "hidden");
            $("#faseshow_" + idval + "").css("visibility", "hidden");
            document.getElementById("faseshow_" + idval + "").value = "hidden"

            $.scrollTo(this, 300, { offset: -300 }); //
            //$.scrollTo(this, 800, { offset: 200 });
            //$.scrollTo('400px', 800);
        }


    });





    $("#ea_kn_1").click(function () {
        fordelaesyregtimer(1);
    });

    $("#ea_kn_2").click(function () {
        fordelaesyregtimer(2);
    });

    $("#ea_kn_3").click(function () {
        fordelaesyregtimer(3);
    });

    $("#ea_kn_4").click(function () {
        fordelaesyregtimer(4);
    });

    $("#ea_kn_5").click(function () {
        fordelaesyregtimer(5);
    });

    $("#ea_kn_6").click(function () {
        fordelaesyregtimer(6);
    });

    $("#ea_kn_7").click(function () {
        fordelaesyregtimer(7);
    });



    function fordelaesyregtimer(id) {

        var antallinier = $("#antalaktlinier").val();
        //alert(antallinier)
        var antaltimer = $("#ea_" + id + "").val();

        if (antaltimer.length > 0) {

            antaltimer = antaltimer.replace(",", ".")
            if (antallinier == 0) {
                antallinier = 1;
            }


            var res = 0;

            res = Math.round(antaltimer / antallinier * 100) / 100
            alert("Der bliver tildelt: " + res + " time til hver Easyreg. akt.\nFindes der timer i forvejen bliver de nye timer lagt til.")

            var preval = 0;
            var newval = 0;
            var oprVerdi = 0;
            var i = 1; //iRowloop

            for (i = 1; i < 500; i++) { //bør svare til maks antal EA linier


               

                if (res == 0) {

                    //oprVerdi = $("#FM_timer_opr_" + i + "_" + id).val()
                    //alert("her: " + oprVerdi)
                    //(document.getElementById("FM_" + dagtype + "_opr_" + nummer + "").value / 1);
                    //$("#FM_timer_" + i + "_" + id).val(oprVerdi)
                    $("#FM_timer_" + i + "_" + id).val(0)

                } else {

                    //if ($("#FM_timer_" + i + "_" + id).exists()) {
                    //var thisid_len = $("#FM_timer_" + i + "_" + id).val()

                    if ($("#FM_timer_" + i + "_" + id).length > 0) {

                        if (!validZip($("#FM_timer_" + i + "_" + id).val())) {
                            //alert("Der er angivet et ugyldigt tegn.")
                            $(".dcls_" + id).val(0)

                            return false

                        } else {

                            preval = $("#FM_timer_" + i + "_" + id).val()
                            preval = jQuery.trim(preval)


                            if (preval != 0 && preval.length > 0) {
                                preval = preval.replace(",", ".")
                                preval = preval
                                newval = String(Math.round((preval / 1 + res / 1) * 100) / 100)
                                newval = newval.replace(".", ",")
                            } else {
                                //alert(res.length + "  res: " + res)
                                if (res != 0) {
                                    preval = 0
                                    newval = String(Math.round((res / 1) * 100) / 100)
                                } else {
                                    newval = 0
                                }
                            }



                            $("#FM_timer_" + i + "_" + id).val(newval)
                            //$("#FM_timer_opr_" + i + "_" + id).val(newval)


                        }

                    } else {
                        //alert(res+" i: " + i)
                        res = String(res)
                        res = res.replace(".", ",")
                        $("#FM_timer_" + i + "_" + id).val(res)
                        //$("#FM_timer_opr_" + i + "_" + id).val(res)
                    }


                } // nulstil


            } // for

        } //length ea
    }
    // end func //


    $(".fjeasyregakt").click(function () {

        var thisid = this.id
        var thisC = thisid
       
        $.post("?fjerneasyakt=1", { control: "FN_fjeasyregakt", AjaxUpdateField: "true", cust: thisid }, function (data) {

            //$.scrollTo('+=250px', 200);
             $("#"+thisid).hide(200);

             //alert("ok")
         });    
      

    });



    $(".jobid").click(function () {

        var thisid = this.id

        //alert(thisid)

        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(6, thisvallngt)
        thisval = thisvaltrim

        //alert(thisval)


        str = $("#seljobid").val()
        str = String(str).replace("," + thisval + ",", ",")
        $("#seljobid").val(str)





    });



    $("#tildel_sel_medarb_bt").click(function () {


        tildelaktivejobliste_load()

    });

    function tildelaktivejobliste_load() {

        //alert("gfh")

        var vlgtPrgrp = $("#tildel_sel_pgrp")
        vlgtmedarb = $("#tildel_sel_medarb").val()


        //var medCookie = $.cookie("tildel_sel_medarb_c")
        $.cookie("tildel_sel_medarb_c", vlgtmedarb)

        //alert(vlgtmedarb)

        //$("#BtnCustDescUpdate").data("cust", thisC.val());
        $.post("?vlgtmedarb=" + vlgtmedarb, { control: "FN_getTildelJob", AjaxUpdateField: "true", prgrp: vlgtPrgrp.val() }, function (data) {
            //$("#fajl").val(data);

            $("#div_tildeljoblisten").html(data);

        });





    }


  




    $(".dcls_1").keyup(function () {

        alertDiv(this.id)
        //dagstotal(1)
        tjektimer(this.id)

    });

    $(".dcls_2").keyup(function () {

        alertDiv(this.id)
        //dagstotal(2)
        tjektimer(this.id)

    });

    $(".dcls_3").keyup(function () {

        alertDiv(this.id)
        //dagstotal(3)
        tjektimer(this.id)

    });

    $(".dcls_4").keyup(function () {

        alertDiv(this.id)
        //dagstotal(4)
        tjektimer(this.id)

    });

    $(".dcls_5").keyup(function () {

        alertDiv(this.id)
        //dagstotal(5)
        tjektimer(this.id)

    });

    $(".dcls_6").keyup(function () {

        alertDiv(this.id)
        //dagstotal(6)
        tjektimer(this.id)

    });

    $(".dcls_7").keyup(function () {

        alertDiv(this.id)
        //dagstotal(7)
        tjektimer(this.id)

    });


    function alertDiv(rloopThis) {
        var thisid = rloopThis
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(9, thisvallngt - 2)
        thisval = thisvaltrim


        //alert(thisval)
        //alert($("#FM_timer_" + this.id + "_1").val())
        jobid = $("#FM_jobid_" + thisval).val()
        $("#FM_jobid_timerOn_" + jobid).val('1')

    }

    function dagstotal(dag) {

        var timerval = 0;
        var tfltVal = 0;
        var loops = $("#loops").val();


        for (i = 1; i <= loops; i++) {
            tfltVal = jQuery.trim($("#FM_timer_" + i + "_" + dag + "").val())
            //$("#FM_timer_" + i + "_" + dag + "").val(tfltVal.replace(".", ","))
            if (tfltVal.length > 0) {
                //alert($("#FM_timer_" + i + "_1").val() + " # " + tfltVal.length)
                if (tfltVal / 1 > 0) {
                    tfltVal = tfltVal.replace(",", ".")
                    timerval = timerval + (tfltVal / 1)
                }
            }

        }

        timerval = String(Math.round(timerval * 100) / 100).replace(".", ",")
        $("#timer_" + dag).val(timerval)

        //var p = $("div:dagstotaler");
        //var position = p.position();
        //alert(position.left + " # " + position.top +" ## "+  position.bottom)
        //var top = screen.height;
        //top = top - 300
        //$("div:dagstotaler").css("top", top);

        //$("#sm1").focus()


    }




  


    $(".a_showhideaktbesk").click(function () {
        thisid = this.id
        var idlngt = thisid.length
        var idtrim = thisid.slice(8, idlngt)
        aktid = idtrim

        //alert(aktid)

        hgt = $("#div_aktbesk_" + aktid + "").css("height")

        if (hgt == "100px") {
            $("#div_aktbesk_" + aktid + "").css("height", "14px");
            $("#div_aktbesk_" + aktid + "").css("overflow", "hidden");
            //document.getElementById(aktid).style.height = "14px";
            //document.getElementById(aktid).style.overflow = "hidden";

            $.scrollTo(this, 300, { offset: -300 }); //

        } else {
            //document.getElementById(aktid).style.height = "100px";
            //document.getElementById(aktid).style.overflow = "auto";
            $("#div_aktbesk_" + aktid + "").css("height", "100px");
            $("#div_aktbesk_" + aktid + "").css("overflow", "auto");

            $.scrollTo(this, 300, { offset: -300 }); //
        }


    });




    function validZip(inZip) {
        if (inZip == "") {
            return true
        }
        if (isNum_treg(inZip)) {
            return true
        }
        return false
    }

    function tjektimer(thisid) {



        if (!validZip(document.getElementById(thisid).value)) {
            alert("Der er angivet et ugyldigt tegn.")
            document.getElementById(thisid).value = '';
            document.getElementById(thisid).focus()
            document.getElementById(thisid).select()
            return false
        }
        return true
    }

    function tjekkm(dagtype, nummer) {

        if (!validZip(document.getElementById(thisid).value)) {
            alert("Der er angivet et ugyldigt tegn.")
            document.getElementById(thisid).value = '';
            document.getElementById(thisid).focus()
            document.getElementById(thisid).select()
            return false
        }
        return true
    }



    function isNum_treg(passedVal) {
        invalidChars = " /:;<>abcdefghijklmnopqrstuvwxyzæøå"

        //alert("her")

        if (passedVal == "") {
            return false
        }

        for (i = 0; i < invalidChars.length; i++) {
            badChar = invalidChars.charAt(i)
            if (passedVal.indexOf(badChar, 0) != -1) {
                return false
            }
        }

        for (i = 0; i < passedVal.length; i++) {
            if (passedVal.charAt(i) == "." || passedVal.charAt(i) == "-") {
                return true
            }
            else {
                if (passedVal.charAt(i) < "0") {
                    return false
                }
                if (passedVal.charAt(i) > "9") {
                    return false
                }
            }
            return true
        }

    }






});







