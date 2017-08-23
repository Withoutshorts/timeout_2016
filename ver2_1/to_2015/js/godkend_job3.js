




$(document).ready(function () {

   $("#godkendalle").click(function () {

       //alert("alle"+ this + $(".godkendalle").attr("checked"))

            $(".godkendbox").each(function () {

           //alert("HER 2" + this.id)
           //$(this).attr('checked', true);

           if ($("#godkendalle").is(':checked') == true) {

               $(this).prop("checked", true)
           } else {
               $(this).prop("checked", false)
           }

            });

            if ($("#godkendalle").is(':checked') == true) {
                $("#afvisalle").prop("checked", false)
                $(".afvismedarb").prop("checked", false)
            }

 
   });


   $("#afvisalle").click(function () {

       //alert("alle"+ this + $(".godkendalle").attr("checked"))

       $(".afvisbox").each(function () {

           //alert("HER 2" + this.id)
           //$(this).attr('checked', true);

           if ($("#afvisalle").is(':checked') == true) {

               $(this).prop("checked", true)
               $("#godkendalle").prop("checked", false)
           } else {
               $(this).prop("checked", false)
           }

       });


       if ($("#afvisalle").is(':checked') == true) {
           $("#godkendalle").prop("checked", false)
           $(".godkendmedarb").prop("checked", false)
           
       } 
       

   });




   $(".godkendmedarb").click(function () {

      

       thisid = this.id
       var arr = thisid.split('_');

       var medid = arr[1];
       
      
       $(".godkendbox_" + medid).each(function () {

           

           if ($("#"+thisid).is(':checked') == true) {

               $(this).prop("checked", true)
           } else {
               $(this).prop("checked", false)
           }

       });

       if ($("#" + thisid).is(':checked') == true) {
           $("#medarbafvis_" + medid).prop("checked", false)
       }
     
   });



   $(".afvismedarb").click(function () {



       thisid = this.id
       var arr = thisid.split('_');

       var medid = arr[1];


       $(".afvisbox_" + medid).each(function () {



           if ($("#" + thisid).is(':checked') == true) {

               $(this).prop("checked", true)
           } else {
               $(this).prop("checked", false)
           }

       });

       if ($("#" + thisid).is(':checked') == true) {
           $("#medarbgodkend_" + medid).prop("checked", false)
       }
    

   });

     


    /// Submitter til DB

    $(".godkendknap").click(function () {

       

        //var $godkendboxes = $('input[class=godkendbox]:checked');
       

            //$godkendboxes.each(function () {
        $(".godkendbox").each(function () {

               
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




            //var $afvisboxes = $('input[class=afvisbox]:checked');

            //$afvisboxes.each(function () {
        globalMid = 0;
        globalTids = "";

        lastMid = 0;

      

         $(".afvisbox").each(function () {

             var thisid = this.id;

                var arr = thisid.split('_');

                var medid = arr[1];
                var aktid = arr[2];
                var tid = arr[3];

                if ($("#" + thisid).is(':checked') == true) {

                startdato = $("#startdato").val()
                slutdato = $("#slutdato").val()
                godkendjobid = $("#godkendjobid").val()


                //alert(startdato + slutdato + "medid:" + medid + "aktid: " + aktid + "godkendjobid:" + godkendjobid)
                //return true

                decl_tids = $("#decl_tids_" + medid).val();
                $("#decl_tids_" + medid).val(decl_tids + " OR tid = " + tid);
                //$("#decl_tids_mid_" + medid).val(medid);
                globalMid = medid;
                globalTids = $("#decl_tids_" + globalMid).val();

                $.post("?tid=" + tid + "&medid=" + medid + "&startdato=" + startdato + "&slutdato=" + slutdato + "&aktid=" + aktid + "&godkendjobid=" + godkendjobid, { control: "afvisuge", AjaxUpdateField: "true" }, function (data) {

                    //alert("godkendt");

                    //$("#div_jobid").html(data);
                    

                });


              
                lastMid = medid

                } // CHECKED = TRUE




            });
           

      

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
        

          $("#godkendform").submit();
    });


    
    function sendmail() {

        

      
    };
   


   


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



