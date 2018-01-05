






$(window).load(function () {

    //visSubtotJobmed();

});


$(document).ready(function () {


   



    $(".sp_jid").mouseover(function () {
        $(this).css('cursor', 'pointer');
    });


    $(".sp_p").mouseover(function () {
        $(this).css('cursor', 'pointer');
    });

    $("#a_afdm").mouseover(function () {
        $(this).css('cursor', 'pointer');
    });

    $("#a_jobakt").mouseover(function () {
        $(this).css('cursor', 'pointer');
    });


    $("#load").hide(1000);



    $(".sp_p").click(function () {

        var thisid = this.id
        var idlngt = thisid.length
        var idtrim = thisid.slice(5, idlngt)




        if ($(".afd_p_" + idtrim).css('display') == "none") {


            //$(".afd_p").css("display", "none");
            //$(".afd_p").css("visibility", "hidden");
            $(".afd_p").hide(100);
            $(".afd_p").css("background-color", "");

            $(".afd_p_" + idtrim).css("display", "");
            $(".afd_p_" + idtrim).css("visibility", "visible");
            $(".afd_p_" + idtrim).show(100);




        } else {

            $(".afd_p_" + idtrim).css("display", "none");
            $(".afd_p_" + idtrim).css("visibility", "hidden");
            $(".afd_p_" + idtrim).hide(100);



        }


    });


    $("#a_afdm").click(function () {

        if ($("#dv_afdm").css('display') == "none") {



            $("#dv_afdm").css("display", "");
            $("#dv_afdm").css("visibility", "visible");
            $("#dv_afdm").show(100);

            if ($("#d_jobakt").css('display') == "none") {

                $("#dv_afdm").css("left", "0px");
                $("#a_afdm").css("left", "150px");

            } else {

                $("#dv_afdm").css("left", "1250px");
                $("#a_afdm").css("left", "1250px");

            }


        } else {

            $("#dv_afdm").css("display", "none");
            $("#dv_afdm").css("visibility", "hidden");
            $("#dv_afdm").hide(100);

            $("#dv_afdm").css("left", "0px");
            $("#a_afdm").css("left", "150px");


        }

    });


    $("#a_jobakt").click(function () {

        if ($("#d_jobakt").css('display') == "none") {



            $("#d_jobakt").css("display", "");
            $("#d_jobakt").css("visibility", "visible");
            $("#d_jobakt").show(100);

            if ($("#dv_afdm").css('display') == "none") {


                $("#a_afdm").css("left", "150px");

            } else {

                $("#dv_afdm").css("left", "1250px");
                $("#a_afdm").css("left", "1250px");
            }


        } else {

            $("#d_jobakt").css("display", "none");
            $("#d_jobakt").css("visibility", "hidden");
            $("#d_jobakt").hide(100);

            $("#dv_afdm").css("left", "0px");
            $("#a_afdm").css("left", "150px");

        }

    });



    $(".sp_jid").click(function () {

        var thisid = this.id
        var idlngt = thisid.length
        var idtrim = thisid.slice(3, idlngt)



        if ($(".tr_" + idtrim).css('display') == "none") {

            $(".tr_" + idtrim).css("display", "");
            $(".tr_" + idtrim).css("visibility", "visible");
            $(".tr_" + idtrim).show(2000);

            str = idtrim
            var n = str.indexOf("_");
            if (n == -1) {

                $("#sp_" + idtrim).html('<b>[-]&nbsp;&nbsp;</b>');
            }



        } else {

            $(".tr_" + idtrim).css("display", "none");
            $(".tr_" + idtrim).css("visibility", "hidden");
            $(".tr_" + idtrim).hide(300);

            str = idtrim
            var n = str.indexOf("_");

            if (n == -1) {
                $("#sp_" + idtrim).html('<b>[+]&nbsp;&nbsp;</b>');
            }

        }


    });


    // farveskift ved tildel budgettimer //
    $(".jobakt_budgettimer_FY, .jobakt_budgettimer_job").keyup(function () {

        var thisid = this.id
        var idlngt = thisid.length
        var idtrim = thisid.slice(19, idlngt)


        budgettimer = $("#FM_timerbudget_FY0_" + idtrim).val().replace(",", ".")
        realtimer = $("#realtimer_jobakt_" + idtrim).val().replace(",", ".")

        thisval = Math.round((budgettimer / 1 - realtimer / 1) * 100) / 100

        if (thisval < 0) {
            $("#h2t_jobakt_" + idtrim).css("background-color", "red")
        } else {
            $("#h2t_jobakt_" + idtrim).css("background-color", "")
        }

        thisval = String(thisval).replace(".", ",")

        $("#h2t_jobakt_" + idtrim).val(thisval)



        budgettimer1 = $("#FM_timerbudget_FY1_" + idtrim).val().replace(",", ".")
        budgettimer2 = $("#FM_timerbudget_FY2_" + idtrim).val().replace(",", ".")

        budgettimerTotFY = (budgettimer / 1 + budgettimer1 / 1 + budgettimer2 / 1)
        jobbudgettimer = $("#budgettimer_jobakt_" + idtrim).val().replace(",", ".")


        if ((jobbudgettimer / 1 - budgettimerTotFY / 1) < 0) {
            $("#budgettimer_jobakt_" + idtrim).css("background-color", "red")
        } else {
            $("#budgettimer_jobakt_" + idtrim).css("background-color", "")
        }





        //alert(budgettimer)

        //budgettimerColor(h1totjobakt, h2totjobakt, budgettimer, idtrim);

    });



    ////////////////////////////////////////////////////////////////////////////////////
    // Beregn H1H2 MEDARBEJDER FORECAST///
    $(".mh1, .mh2").keyup(function () {


        //medarbejder total på job / akt h1 + h2
        var thisid = this.id
        var idlngt = thisid.length
        var idtrim = thisid.slice(15, idlngt)


        //////////////////////////////////////////////////////////////////////////////////
        //// AKT --> Joblinje


        // H1 //
        h1jobTot = 0;

        if (aktid != 0) { //Akt sumer til job



            var arr = idtrim.split('_');


            $(".mh1h_akt_" + jobid + "_" + arr[2]).each(function () {



                h1job = $(this).val().replace(".", "")
                h1job = h1job.replace(",", ".")

                if (h1job == "NaN" || (h1job == "-Infinity") || h1job == "") {
                    h1job = 0
                }

                h1jobTot = (h1jobTot / 1 + h1job / 1)


            });





            //alert("#mh1h_jobaktmid_" + jobid + "_0_" + "_" + arr[2])

            h1jobTot = Math.round(h1jobTot)
            h1jobTot = h1jobTot
            h1jobTot = String(h1jobTot).replace(".", ",")
            $("#mh1h_jobaktmid_" + jobid + "_0_" + arr[2]).val(h1jobTot)

            //$("#h1t_jobakt_" + jobid + "_0").val(h1jobTot) //h1t_jobakt


        }


        // H2 //
        h2jobTot = 0;

        if (aktid != 0) { //Akt sumer til job



            var arr = idtrim.split('_');


            $(".mh2h_akt_" + jobid + "_" + arr[2]).each(function () {



                h2job = $(this).val().replace(".", "")
                h2job = h2job.replace(",", ".")

                if (h2job == "NaN" || (h2job == "-Infinity") || h2job == "") {
                    h2job = 0
                }

                h2jobTot = (h2jobTot / 1 + h2job / 1)


            });


            //alert("#mh2h_jobaktmid_" + jobid + "_0_" + "_" + arr[2])

            h2jobTot = Math.round(h2jobTot)
            h2jobTot = h2jobTot
            h2jobTot = String(h2jobTot).replace(".", ",")
            $("#mh2h_jobaktmid_" + jobid + "_0_" + arr[2]).val(h2jobTot)
            //$("#h2t_jobakt_" + jobid + "_0").val(h2jobTot) //h1t_jobakt




        }





        // SLUT sammentælling akt->job
        //////////////////////////////////////////////////////////////////////////////////////



        beregnh1h2(idtrim);



    });







    /////////////////////////////////////////////////////////////////////////////////////////



    function beregnh1h2(idtrim) {


        //alert(idtrim)

        h1 = $("#mh1h_jobaktmid_" + idtrim).val()
        h2 = $("#mh2h_jobaktmid_" + idtrim).val()

        //alert(h2)

        h1 = h1.replace(",", ".")
        h2 = h2.replace(",", ".")


        if (h1 == "NaN" || (h1 == "-Infinity") || h1 == "") {
            h1 = 0
        }

        if (h2 == "NaN" || (h2 == "-Infinity") || h2 == "") {
            h2 = 0
        }


        h1h2 = Math.round(h1 / 1 + h2 / 1)
        h1h2 = String(h1h2).replace(".", ",")


        $("#mh12h_jobaktmid_" + idtrim).val(h1h2)


        jobid = $("#h1_jobid_" + idtrim).val()
        aktid = $("#h1_aktid_" + idtrim).val()

        h1totjobakt = 0
        h2totjobakt = 0

        h1jobakt = 0
        h2jobakt = 0

        //for (i = 0; i <= 10; i++) {


        // TOTAL PR. JOB/  AKT LINJE
        $(".mh1h_jobaktmid_" + jobid + "_" + aktid).each(function () {



            h1jobakt = $(this).val().replace(".", "")
            h1jobakt = h1jobakt.replace(",", ".")
            if (h1jobakt == "NaN" || (h1jobakt == "-Infinity") || h1jobakt == "") {
                h1jobakt = 0
            }



            h1totjobakt = h1totjobakt + (h1jobakt / 1)


        });



        $(".mh2h_jobaktmid_" + jobid + "_" + aktid).each(function () {



            h2jobakt = $(this).val().replace(".", "")
            h2jobakt = h2jobakt.replace(",", ".")
            if (h2jobakt == "NaN" || (h2jobakt == "-Infinity") || h2jobakt == "") {
                h2jobakt = 0
            }



            h2totjobakt = h2totjobakt + (h2jobakt / 1)

        });


        //alert(h1totjobakt)

        $("#h1h_jobaktmid_" + jobid + "_" + aktid).val(h1totjobakt)
        $("#h2h_jobaktmid_" + jobid + "_" + aktid).val(h2totjobakt)





        //budgettimer_jobakt_
        budgettimer = 0;
        budgettimer = $("#FM_timerbudget_FY0_" + jobid + "_" + aktid).val().replace(",", ".")


        budgettimerColor(h1totjobakt, h2totjobakt, budgettimer, jobid + "_" + aktid);



        //


        // GT på job / AKT 


        h1totjobaktGT = 0;
        $(".mh1h_jobaktmid_" + jobid + "_0").each(function () {



            h1jobakt = $(this).val().replace(".", "")



            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(14, idlngt)

            var arr = idtrim.split('_');

            //alert(this.id + " aar1: " + arr[1])

            if (arr[1] != 0) {
                //h1h_jobaktmid_<%=jobid%>_<%=aktid %>

                h1jobakt = h1jobakt.replace(",", ".")
                if (h1jobakt == "NaN" || (h1jobakt == "-Infinity") || h1jobakt == "") {
                    h1jobakt = 0
                }


                //alert(thisid +" "+ h1jobakt)

                h1totjobaktGT = h1totjobaktGT + (h1jobakt / 1)

            }


        });


        h1totjobaktGT = Math.round(h1totjobaktGT)
        h1totjobaktGT = h1totjobaktGT
        h1totjobaktGT = String(h1totjobaktGT).replace(".", ",")


        h2totjobaktGT = 0;
        $(".mh2h_jobaktmid_" + jobid + "_0").each(function () {

            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(14, idlngt)

            var arr = idtrim.split('_');

            if (arr[1] != 0) {

                h2jobakt = $(this).val().replace(".", "")
                h2jobakt = h2jobakt.replace(",", ".")
                if (h2jobakt == "NaN" || (h2jobakt == "-Infinity") || h2jobakt == "") {
                    h2jobakt = 0
                }



                h2totjobaktGT = h2totjobaktGT + (h2jobakt / 1)

            }

        });

        h2totjobaktGT = Math.round(h2totjobaktGT)
        h2totjobaktGT = h2totjobaktGT
        h2totjobaktGT = String(h2totjobaktGT).replace(".", ",")

        $("#h1h_jobaktmid_" + jobid + "_0").val(h1totjobaktGT)
        $("#h2h_jobaktmid_" + jobid + "_0").val(h2totjobaktGT)


        // JOB og AKT Sammentæller totaler ///


        // AFDELINGER
        phigh = $("#phigh").val()

        afdelingTot(aktid, jobid, phigh);




        totalalleh1h2(1);


        //alert("ok")

    }


    //////////////////////////////////////////////////////////////////////////////////////







    function afdelingTot(aktid, jobid, phigh) {


        for (p = 0; p <= phigh; p++) {

            fs_totTot = 0
            $(".fs_" + jobid + "_" + aktid + "_" + p).each(function () {



                fs_tot = $(this).val().replace(".", "")
                fs_tot = fs_tot.replace(",", ".")


                if (fs_tot == "NaN" || (fs_tot == "-Infinity") || fs_tot == "") {
                    fs_tot = 0
                }

                fs_totTot = (fs_totTot / 1 + fs_tot / 1)

                //alert(fs_totTot)


            });

            fs_totTot = String(fs_totTot).replace(".", ",")

            $("#afd_jobakt_" + jobid + "_" + aktid + "_" + p).val(fs_totTot)

        }


    }













    //////////////////////////////////////////////////////////////////////////////////////
    // SAMMENTÆLLING AKT --> JOB
    //////////////////////////////////////////////////////////////////////////////////////



    //////////////////////////////////////////////////////////////////////////////////
    //// Timepriser JOB
    $(".jobakt_budgettimep_job").keyup(function () {

        //alert("her")

        var thisid = this.id
        var idlngt = thisid.length
        var idtrim = thisid.slice(19, idlngt)

        var arr = idtrim.split('_');

        //alert(idtrim)

        ///////////////////////////////////////////////////////////////////////////////////////////////
        // Sammentælling Akt --> Job 

        // TPJOB //
        jobbudgettp = 0;
        if (arr[1] != 0) { //Akt sumer til job
            i = 0
            $(".jobakt_budgettimep_" + arr[0]).each(function () {



                tottp1 = $(this).val().replace(".", "")
                tottp1 = tottp1.replace(",", ".")



                if (tottp1 == "NaN" || (tottp1 == "-Infinity") || tottp1 == "") {
                    tottp1 = 0
                }

                jobbudgettp = (jobbudgettp / 1 + tottp1 / 1)

                if (tottp1 != 0) {
                    i = i + 1
                }



            });

            if (i == 0) {
                i = 1
            }


            jobbudgettp = jobbudgettp / i
            jobbudgettp = Math.round(jobbudgettp)
            jobbudgettp = String(jobbudgettp).replace(".", ",")

            //alert("jer" + jobbudgettp)

            $("#budgettimep_jobakt_" + arr[0] + "_0").val(jobbudgettp)


        }



        // SLUT sammentælling akt->job
        //////////////////////////////////////////////////////////////////////////////////////

        totalalleh1h2(1)

    });


















    //////////////////////////////////////////////////////////////////////////////////
    //// JOBTIMER
    $(".jobakt_budgettimer_job").keyup(function () {



        var thisid = this.id
        var idlngt = thisid.length
        var idtrim = thisid.slice(19, idlngt)

        var arr = idtrim.split('_');



        ///////////////////////////////////////////////////////////////////////////////////////////////
        // Sammentælling Akt --> Job 

        // FY0 //
        jobbudgetTot = 0;
        if (arr[1] != 0) { //Akt sumer til job

            $(".jobakt_budgettimer_" + arr[0]).each(function () {



                jobbudget = $(this).val().replace(".", "")
                jobbudget = jobbudget.replace(",", ".")

                //alert(jobbudget)

                if (jobbudget == "NaN" || (jobbudget == "-Infinity") || jobbudget == "") {
                    jobbudget = 0
                }

                jobbudgetTot = (jobbudgetTot / 1 + jobbudget / 1)


            });



            //jobbudgetTot = Math.round((jobbudgetTot) / 100) * 100
            jobbudgetTot = jobbudgetTot
            jobbudgetTot = String(jobbudgetTot).replace(".", ",")
            $("#budgettimer_jobakt_" + arr[0] + "_0").val(jobbudgetTot)


        }



        // SLUT sammentælling akt->job
        //////////////////////////////////////////////////////////////////////////////////////

        totalalleh1h2(1)

    });


   

      



        //////////////////////////////////////////////////////////////////////////////////
        //// FY 0 JOB
        $(".jobakt_budgettimer_FY").keyup(function () {



            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(19, idlngt)

            var arr = idtrim.split('_');



            ///////////////////////////////////////////////////////////////////////////////////////////////
            // Sammentælling Akt --> Job 

            // FY0 //
            totFy0Tot = 0;
            if (arr[1] != 0) { //Akt sumer til job
               
                $(".jobakt_fy0_" + arr[0]).each(function () {



                    totFy0 = $(this).val().replace(".", "")
                    totFy0 = totFy0.replace(",", ".")

                    //alert(totFy0)

                    if (totFy0 == "NaN" || (totFy0 == "-Infinity") || totFy0 == "") {
                        totFy0 = 0
                    }

                    totFy0Tot = (totFy0Tot / 1 + totFy0 / 1)

                 
                });

              

                //totFy0Tot = Math.round((totFy0Tot) / 100) * 100
                totFy0Tot = totFy0Tot 
                totFy0Tot = String(totFy0Tot).replace(".", ",")
                $("#FM_timerbudget_FY0_" + arr[0] + "_0").val(totFy0Tot)


            }



            // SLUT sammentælling akt->job
            //////////////////////////////////////////////////////////////////////////////////////

            totalalleh1h2(1)

        });


        //////////////////////////////////////////////////////////////////////////////////
        //// FY 1 JOB
        $(".jobakt_budgettimer_FY1").keyup(function () {



            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(19, idlngt)

            var arr = idtrim.split('_');



            ///////////////////////////////////////////////////////////////////////////////////////////////
            // Sammentælling Akt --> Job 

            // FY1 //
            totFy1Tot = 0;
            if (arr[1] != 0) { //Akt sumer til job
               
                $(".jobakt_fy1_" + arr[0]).each(function () {



                    totFy1 = $(this).val().replace(".", "")
                    totFy1 = totFy1.replace(",", ".")

                    //alert(totFy1)

                    if (totFy1 == "NaN" || (totFy1 == "-Infinity") || totFy1 == "") {
                        totFy1 = 0
                    }

                    totFy1Tot = (totFy1Tot / 1 + totFy1 / 1)

                   
                });

              

                //totFy1Tot = Math.round((totFy1Tot) / 100) * 100
                totFy1Tot = totFy1Tot
                totFy1Tot = String(totFy1Tot).replace(".", ",")
                $("#FM_timerbudget_FY1_" + arr[0] + "_0").val(totFy1Tot)


            }



            // SLUT sammentælling akt->job
            //////////////////////////////////////////////////////////////////////////////////////

            totalalleh1h2(1)

        });




        //////////////////////////////////////////////////////////////////////////////////
        //// FY 2 JOB
        $(".jobakt_budgettimer_FY2").keyup(function () {



            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(19, idlngt)

            var arr = idtrim.split('_');



            ///////////////////////////////////////////////////////////////////////////////////////////////
            // Sammentælling Akt --> Job 

            // FY1 //
            totFy2Tot = 0;
            if (arr[1] != 0) { //Akt sumer til job
                $(".jobakt_fy2_" + arr[0]).each(function () {



                    totFy2 = $(this).val().replace(".", "")
                    totFy2 = totFy2.replace(",", ".")

                    //alert(totFy2)

                    if (totFy2 == "NaN" || (totFy2 == "-Infinity") || totFy2 == "") {
                        totFy2 = 0
                    }

                    totFy2Tot = (totFy2Tot / 1 + totFy2 / 1)

                   
                });

               

                //totFy2Tot = Math.round((totFy2Tot) / 100) * 100
                totFy2Tot = totFy2Tot 
                totFy2Tot = String(totFy2Tot).replace(".", ",")
                $("#FM_timerbudget_FY2_" + arr[0] + "_0").val(totFy2Tot)


            }



            // SLUT sammentælling akt->job
            //////////////////////////////////////////////////////////////////////////////////////

            totalalleh1h2(1)

        });





        function beregntpH1(idtrim) {


            

            var arr = idtrim.split('_');



            ///////////////////////////////////////////////////////////////////////////////////////////////
            // Sammentælling Akt --> Job 

            // TP1 //
            tottp1Tot = 0;
            if (arr[1] != 0) { //Akt sumer til job
                i = 0
                $(".jobakt_tph1_" + arr[0]).each(function () {



                    tottp1 = $(this).val().replace(".", "")
                    tottp1 = tottp1.replace(",", ".")

                    //alert(tottp1)

                    if (tottp1 == "NaN" || (tottp1 == "-Infinity") || tottp1 == "") {
                        tottp1 = 0
                    }

                    tottp1Tot = (tottp1Tot / 1 + tottp1 / 1)

                    if (tottp1 != 0) {
                        i = i + 1
                    }



                });

                if (i == 0) {
                    i = 1
                }


                tottp1Tot = tottp1Tot / i
                tottp1Tot = Math.round(tottp1Tot)
                tottp1Tot = String(tottp1Tot).replace(".", ",")
                $("#fctimepris_h1_jobakt_" + arr[0] + "_0").val(tottp1Tot)


            }


        }






        //////////////////////////////////////////////////////////////////////////////////
        //// Timepriser H1 & H2 JOB
        $(".jobakt_tph1").keyup (function () {

            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(21, idlngt)
           

            beregntpH1(idtrim);


          
            // SLUT sammentælling akt->job
            //////////////////////////////////////////////////////////////////////////////////////

               totalalleh1h2(1)

        });



        //////////////////////////////////////////////////////////////////////////////////
        //// Timepriser H1 & H2 JOB
        $(".jobakt_tph2").keyup(function () {



            var thisid = this.id
            var idlngt = thisid.length
            var idtrim = thisid.slice(21, idlngt)

            var arr = idtrim.split('_');

            // jobakt_tph2_<%=jobid %>



            ///////////////////////////////////////////////////////////////////////////////////////////////
            // Sammentælling Akt --> Job 

            // TP2 //
            tottp2Tot = 0;
            if (arr[1] != 0) { //Akt sumer til job
                i = 0
                $(".jobakt_tph2_" + arr[0]).each(function () {



                    tottp2 = $(this).val().replace(".", "")
                    tottp2 = tottp2.replace(",", ".")

                    //alert(tottp2)

                    if (tottp2 == "NaN" || (tottp2 == "-Infinity") || tottp2 == "") {
                        tottp2 = 0
                    }

                    tottp2Tot = (tottp2Tot / 1 + tottp2 / 1)

                    if (tottp2 != 0) {
                        i = i + 1
                    }



                });

                if (i == 0) {
                    i = 1
                }

                
                tottp2Tot = tottp2Tot / i
                tottp2Tot = Math.round(tottp2Tot)
                tottp2Tot = String(tottp2Tot).replace(".", ",")
                $("#fctimepris_h2_jobakt_" + arr[0] + "_0").val(tottp2Tot)

            }
            // SLUT sammentælling akt->job
            //////////////////////////////////////////////////////////////////////////////////////



            totalalleh1h2(1)

        });


     



        // SLUT sammentælling akt->job
        //////////////////////////////////////////////////////////////////////////////////////
        //////////////////////////////////////////////////////////////////////////////////////
        //////////////////////////////////////////////////////////////////////////////////////











        function beregnbudget(jobid, aktid){
        
          ////////////////////////////////////////////////////////////////////////////////
            // Budget //
            h1timer = 0;
            h1timer = $("#h1t_jobakt_" + jobid + "_" + aktid).val().replace(",", ".")
            fctimepris = 0;
            fctimepris = $("#fctimepris_h1_jobakt_" + jobid + "_" + aktid).val().replace(",", ".")

            h2timer = 0;
            h2timer = $("#h2t_jobakt_" + jobid + "_" + aktid).val().replace(",", ".")
            fctimeprish2 = 0;
            fctimeprish2 = $("#fctimepris_h2_jobakt_" + jobid + "_" + aktid).val().replace(",", ".")

            //alert(budgettimer)


            realtimer = 0;
            realtimer = $("#realtimer_jobakt_" + jobid + "_" + aktid).val().replace(",", ".")

            realtimep = 0;
            realtimep = $("#realtimep_jobakt_" + jobid + "_" + aktid).val().replace(",", ".")

            //alert(realtimer + "*" + realtimep + " = " + (realtimer / 1 * realtimep / 1) + " + " + h2timer / 1 + "*" + fctimeprish2 / 1 + "=" + (h2timer / 1 * fctimeprish2 / 1))

            budgeth1 = Math.round((h1timer / 1 * fctimepris / 1) / 100) * 100
            budgeth2 = Math.round(((realtimer / 1 * realtimep / 1) + (h2timer / 1 * fctimeprish2 / 1)) / 100) * 100

            //alert(budgeth2 / 1)

            budgeth1h2 = (budgeth1 / 1 + budgeth2 / 1)

            budgeth1 = String(budgeth1).replace(".", ",")
            budgeth2 = String(budgeth2).replace(".", ",")
            budgeth1h2 = String(budgeth1h2).replace(".", ",")
            $("#budget_h1_jobakt_" + jobid + "_" + aktid).val(budgeth1)
            $("#budget_h2_jobakt_" + jobid + "_" + aktid).val(budgeth2)

            if (realtimer == 0) {
                $("#budget_jobakt_" + jobid + "_" + aktid).val(budgeth1)
            } else {

                $("#budget_jobakt_" + jobid + "_" + aktid).val(budgeth2)
            }

            //$("#budgetgt").css('background-color','#999999')
            $("#budgetgt").val('Opdater!')

            //alert("her")
            //////////////////////////////////////////////////////////////////////////////////////
        }

       


        function totalalleh1h2(loadberegn) {


            mhigh = $("#mhigh").val()
            xhigh = $("#xhigh").val()
            phigh = $("#phigh").val()
            //alert(xhigh)
            //xhigh = 4

            // JOB og AKT Sammentæller totaler ///
            for (i = 0; i <= xhigh; i++) {

                

                jobid = $("#FM_jobid_" + i).val()
                aktid = $("#FM_aktid_" + i).val()

                //alert("her: " + jobid + "aktid: " + aktid)
                budgettimer = 0;
                //alert(jobid + "_" + aktid + " budgettimer: " + budgettimer)
               

                h1 = $("#h1h_jobaktmid_" + jobid + "_" + aktid).val()
                $("#h1t_jobakt_" + jobid + "_" + aktid).val(h1)

                h1 = h1.replace(".", "")
                h1 = h1.replace(",", ".")

              

                h2 = $("#h2h_jobaktmid_" + jobid + "_" + aktid).val()
                $("#h2t_jobakt_" + jobid + "_" + aktid).val(h2)
                h2 = h2.replace(".", "")
                h2 = h2.replace(",", ".")

                //h2 = 0

                realtimer = $("#realtimer_jobakt_" + jobid + "_" + aktid).val().replace(".", "")
                realtimer = realtimer.replace(",",".")
                //realtimer = 0

                // budgettimer_jobakt_
                // FY0 budgetTimer
                budgettimer = $("#FM_timerbudget_FY0_" + jobid + "_" + aktid).val().replace(",", ".")
                budgettimer = budgettimer.replace(",",".")
                
                //budgettimer = 0
           


                if (h1 == "NaN" || h1 == "-Infinity" || h1 == "") {
                    h1 = 0
                }

                if (h2 == "NaN" || h2 == "-Infinity" || h2 == "") {
                    h2 = 0
                }

                if (realtimer == "NaN" || realtimer == "-Infinity" || realtimer == "") {
                    realtimer = 0
                }

                if (budgettimer == "NaN" || budgettimer == "-Infinity" || budgettimer == "") {
                    budgettimer = 0
                }

                //alert(h1 + "# " + h2 + "# " + budgettimer + "# " + realtimer)

                h1h2total = Math.round((h1/1 + h2/1) * 100) / 100
                h1h2total = String(h1h2total).replace(".", ",")
                $("#h12t_jobakt_" + jobid + "_" + aktid).val(h1h2total)


               
                
                budgettimerColor(h1, h2, budgettimer, jobid+"_"+aktid);


                if (loadberegn == 1) {
                
                    beregnbudget(jobid, aktid);



                        }


           
           


                            // MEDARBEJDERE / NORM
                            for (l = 0; l <= mhigh; l++) {

                                //alert(l)

                                mh1 = $("#totmedarbh1_" + l).val().replace(".", "")
                                mh1 = mh1.replace(",", ".")
                                mh2 = $("#totmedarbh2_" + l).val().replace(".", "")
                                mh2 = mh2.replace(",", ".")

                                mn1 = $("#totmedarbn1_" + l).val().replace(".", "")
                                mn1 = mn1.replace(",", ".")
                                mn2 = $("#totmedarbn2_" + l).val().replace(".", "")
                                mn2 = mn2.replace(",", ".")
                                //alert()
                

                                if (mh1 == "NaN" || (mh1 == "-Infinity") || mh1 == "") {
                                    mh1 = 0
                                }

                                if (mh2 == "NaN" || (mh2 == "-Infinity") || mh2 == "") {
                                    mh2 = 0
                                }

                                if (mn1 == "NaN" || (mn1 == "-Infinity") || mn1 == "") {
                                    mn1 = 0
                                }

                                if (mn2 == "NaN" || (mn2 == "-Infinity") || mn2 == "") {
                                    mn2 = 0
                                }


                                if ((mh1/1) > mn1) {
                                    $("#totmedarbh1_" + l).css("background-color", "red")
                                } else {

                                    if (mh1 = "0,00") {
                                        $("#totmedarbh1_" + l).css("background-color", "")
                                    } else {
                                        $("#totmedarbh1_" + l).css("background-color", "greenyellow")
                                    }
                                }

                                if ((mh2/1) > mn2) {
                                    $("#totmedarbh2_" + l).css("background-color", "red")
                                } else {
                                    if (mh1 = "0,00") {
                                        $("#totmedarbh2_" + l).css("background-color", "")
                                    } else {
                                        $("#totmedarbh2_" + l).css("background-color", "greenyellow")
                                    }
                                }



                            }


            } //loadberegn
           
            

        
        
        } /// All H1H2 totaler func




        function budgettimerColor(h1, h2, budgettimer, jobidaktid){

         

            if ((h1 / 1 + h2 / 1) > budgettimer && budgettimer != 0) {
                $("#FM_timerbudget_FY0_" + jobidaktid).css("background-color", "red")
            } else {

                if (budgettimer == 0) {
                    $("#FM_timerbudget_FY0_" + jobidaktid).css("background-color", "")
                } else {
                    $("#FM_timerbudget_FY0_" + jobidaktid).css("background-color", "")
                }

            }


        }









        //ved load 
        //totalalleh1h2(0); //== REGNES I ASP

        //samler på faser
        // AFDELINGER
        // JOB og AKT Sammentæller totaler ///
        xhigh = $("#xhigh").val()
        phigh = $("#phigh").val()
        //alert("xhigh:" + xhigh)
        for (i = 0; i <= xhigh; i++) {

            jobid = $("#FM_jobid_" + i).val()
            aktid = $("#FM_aktid_" + i).val()

            //alert("her")

            afdelingTot(aktid, jobid, phigh);
        }
      

    });
      
 




       
      