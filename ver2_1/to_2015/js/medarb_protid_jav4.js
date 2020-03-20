$(document).ready(function () {

    var sortering = $("#FM_sortering").val();
    var version = $("#version").val();
    
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
