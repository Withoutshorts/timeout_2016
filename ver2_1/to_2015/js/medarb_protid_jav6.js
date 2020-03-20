$(document).ready(function () {

    var sortering = $("#FM_sortering").val();
    var version = $("#version").val();

    var alletxt = $("#alletxt").val();


    $("#jogsog").keyup(function () {
        thissogval = $(this).val();
       
      /*  if (thissogval == "") {
            $("#FM_job").hide();
        } else {
            $("#FM_job").show();
        } */
        
        $.post("?FM_sogval=" + thissogval, { control: "SogJobFilter", AjaxUpdateField2: "true" }, function (data) {
            $("#FM_job").html("<option value='-1' data-jobnavn='" + alletxt + "'>" + alletxt +"</option>")
            $("#FM_job").append(data)
        });

    });




    $("#FM_job").change(function () {
        $('#FM_job option').each(function () {
            if ($(this).is(':selected'))
            {
                thisjobnr = $(this).val();
                thisjobnavn = $(this).data('jobnavn');
                thisjobid = $(this).data('jobid');
                TilfojJobTilValgte(thisjobnavn, thisjobnr);
            }
            
        });
    });

    function TilfojJobTilValgte(jobnavn, jobnr) {
      /*  var jobvalgt = $('#jobvalgt').val();
        var jobvalgtarr = jobvalgt.split(",");

        var findes = 0;
        for (i = 0; i < jobvalgtarr.length; i++)
        {
            if (jobnr == jobvalgtarr[i])
            {
                findes = 1
            }
        } */

        if (jobnr == -1) {
            $('.jobboxElement').remove();
        }
        else
        {
            FjernJobFraValgte(-1)
        }

        var findes = 0;

        $('.FM_job').each(function () {
            thisjobnr = this.id;

            if (thisjobnr == jobnr) {
                findes = 1;
            }

        });

        if (findes == 0)
        {
            // Ramme på div - background-color:#fafafa; border:2px solid #171717;

            strjobnavn = jobnavn + " ("+ jobnr +")"
            if (jobnr == -1)
            {
                strjobnavn = jobnavn
            }


          //  $('#jobvalgt').val(jobvalgt + "," + jobnr)
            jobbox = "<div class='jobboxElement' id='jobbox_" + jobnr + "' style='height:25px; display:inline-block; margin:2px; margin-right:7px;'>" + strjobnavn + " &nbsp <span style='cursor:pointer;' class='fa fa-times fjernjob' id='" + jobnr + "' ></span> &nbsp "
            jobbox += "<input type='hidden' value='" + jobnr + "' name='FM_job' class='FM_job' id='" + jobnr + "' />"
            jobbox += "</div>"
            $('#jobbox').append(jobbox)
        }
    }

  /*  $(".fjernjob").live('click', function () {
        thisjobnr = this.id;
        alert(thisjobnr)
    }); */

    $(document).on("click", ".fjernjob", function () {
        thisjobnr = this.id;
        FjernJobFraValgte(thisjobnr)
    });

    function FjernJobFraValgte(jobnr)
    {
        $("#jobbox_" + jobnr).remove();        
    }

    if (version == 0 || version == "0") {
        if (sortering == 0 || sortering == "0") {
            $(".expandJob").click(function () {
                thisid = this.id

                if ($(".akt_" + thisid).is(":not(':hidden')")) {
                    $(".akt_" + thisid).hide();
                    $(".medarb_" + thisid).hide();

                    $("#icon_" + thisid).removeClass('fa-angle-up')
                    $("#icon_" + thisid).addClass('fa-angle-down')
                } else {
                    $(".akt_" + thisid).show();

                    $("#icon_" + thisid).removeClass('fa-angle-down')
                    $("#icon_" + thisid).addClass('fa-angle-up')
                }

            });

            $(".expandAkt").click(function () {
                thisid = this.id

                if ($(".medarb_" + thisid).is(":not(':hidden')")) {
                    $(".medarb_" + thisid).hide();

                    $("#icon_" + thisid).removeClass('fa-angle-up')
                    $("#icon_" + thisid).addClass('fa-angle-down')

                } else {
                    $(".medarb_" + thisid).show();

                    $("#icon_" + thisid).removeClass('fa-angle-down')
                    $("#icon_" + thisid).addClass('fa-angle-up')
                }
            });
        } else {
            $(".expandJob").click(function () {
                thisid = this.id

                if ($(".medarb_" + thisid).is(":not(':hidden')")) {
                    $(".medarb_" + thisid).hide();
                    $(".akt_" + thisid).hide();
                    $("#icon_" + thisid).removeClass('fa-angle-up')
                    $("#icon_" + thisid).addClass('fa-angle-down')
                } else {
                    $(".medarb_" + thisid).show();
                    $("#icon_" + thisid).removeClass('fa-angle-down')
                    $("#icon_" + thisid).addClass('fa-angle-up')
                }

            });

            $(".expandMedarb").click(function () {
                thisid = this.id

                if ($(".akt_" + thisid).is(":not(':hidden')")) {
                    $(".akt_" + thisid).hide();
                    $("#icon_" + thisid).removeClass('fa-angle-up')
                    $("#icon_" + thisid).addClass('fa-angle-down')
                } else {
                    $(".akt_" + thisid).show();
                    $("#icon_" + thisid).removeClass('fa-angle-down')
                    $("#icon_" + thisid).addClass('fa-angle-up')
                }
            });
        }
    } else {
        $(".expandMedarb").click(function () {
            thisid = this.id

            if ($(".job_" + thisid).is(":not(':hidden')")) {
                $(".job_" + thisid).hide();
                $(".akt_" + thisid).hide();
                $("#icon_" + thisid).removeClass('fa-angle-up')
                $("#icon_" + thisid).addClass('fa-angle-down')
            } else {
                $(".job_" + thisid).show();
                $("#icon_" + thisid).removeClass('fa-angle-down')
                $("#icon_" + thisid).addClass('fa-angle-up')
            }
        });
        
        $(".expandJob").click(function () {
            thisid = this.id
            if ($(".akt_" + thisid).is(":not(':hidden')")) {
                $(".akt_" + thisid).hide();
                $("#icon_" + thisid).removeClass('fa-angle-up')
                $("#icon_" + thisid).addClass('fa-angle-down')
            } else {
                $(".akt_" + thisid).show();
                $("#icon_" + thisid).removeClass('fa-angle-down')
                $("#icon_" + thisid).addClass('fa-angle-up')
            }
        });
    }


    window.onafterprint = function (e) {
        $(".container").removeClass('zoom')
        $(".luft").show();
        $(".printhide").show();
        $(".fixed-navbar-vert").show();
        $(".fixed-navbar-hoz").show();
    };

    $("#nyprint").click(function () {
        $("#back-to-top").hide();
        $(".container").addClass('zoom')
        $(".luft").hide();
        $(".printhide").hide();
        $(".fixed-navbar-vert").hide();
        $(".fixed-navbar-hoz").hide();

        window.print();
    });



});
