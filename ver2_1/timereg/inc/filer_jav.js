


$(window).load(function () {
    // run code



    $("#loadbar").hide(1000);



    for (i = 0; i < 2000; i++) { //bør svare til maks antal foldere linier

        //alert($.cookie('showfolder_' + i))
        if ($.cookie('showfolder_' + i) == "1") {

            //alert("her" + i)
            $(".foid_" + i).css("display", "");
            $(".foid_" + i).css("visibility", "visible");
            $(".foid_" + i).show(100);

            $("#foid_" + i).html("<img src='../ill/folder_document.png'  alt='' border='0'/>")

        }

    }

});











$(document).ready(function () {


   

    $("#kundeid").change(function () {

        joblist();

    });




    function joblist() {

   
        var kundeid = $("#kundeid").val()
        //alert(kundeid)
        $.post("?kundeid=" + kundeid, { control: "FN_showjob", AjaxUpdateField: "true" }, function (data) {

            $("#jobid").html(data);
            
        });
    }

    
        



 

 
        $(".showfolder").mouseover(function () {
            //alert("her")
            $(this).css("cursor", "pointer");
        });


        $(".showfolder").click(function () {

            var thisid = this.id
            var thisvallngt = thisid.length
            var thisvaltrim = thisid.slice(5, thisvallngt)
            thisid = thisvaltrim

            $(".trfo").css("background-color", "#FFFFFF");

            //alert($(".foid_" + thisid).css('display'))
            if ($(".foid_" + thisid).css('display') == "none") {

                $(".foid_" + thisid).css("display", "");
                $(".foid_" + thisid).css("visibility", "visible");
                $(".foid_" + thisid).show(100);

                $("#trid_" + thisid).css("background-color", "#FFFF99");

                $.scrollTo(this, 200, { offset: -100 });


                $("#foid_" + thisid).html("<img src='../ill/folder_document.png'  alt='' border='0'/>")

                $.cookie('showfolder_' + thisid, '1');

            } else {

                $(".foid_" + thisid).hide(100);
                $.cookie('showfolder_' + thisid, '0');

                $("#trid_" + thisid).css("background-color", "#FFFFE1");
                $("#foid_" + thisid).html("<img src='../ill/folder.png'  alt='' border='0'/>")


                $.scrollTo(this, 200, { offset: -200 });

            }



        });




    });

   




