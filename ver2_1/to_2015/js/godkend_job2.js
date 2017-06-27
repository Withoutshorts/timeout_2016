




$(document).ready(function () {

   $(".godkendalle").click(function () {

        //alert("alle")



        var godkendalleboxses = $('input[type=checkbox]');
        
        if ($(".godkendalle").attr("checked"))
        {
            $(".godkendalle").attr("checked",false)
            godkendalleboxses.each(function () {
                godkendalleboxses.attr("checked", false)            
            });
            //alert("0")
        }
        else 
        {
            $(".godkendalle").attr("checked",true)
            godkendalleboxses.each(function () {
            godkendalleboxses.prop("checked", true)
            });
            //alert("1")
        }

       
        });




      /*  var $godkendalleakt = $('input[class=godkendmedarb]');

        if ($(".godkendalle").attr("checked")) 
        {
            $godkendmedarb.each(function () {
                //alert("hej")
                $godkendmedarb.attr("checked", false);
            });
        }
        else
        {
            $godkendmedarb.each(function () {
                //alert("farvel")
                $godkendmedarb.prop("checked", true);
            });
        } */
                    
    
   

    $(".godkendmedarb").click(function () {


        thisid = this.id
        var thisvallngt = thisid.length
        var thisvaltrim = thisid.slice(14, thisvallngt)

        var $medarbgodkedboxes = $('input[id=godkendstatus_' + thisvaltrim + ']');

       if ($("#medarbgodkend_" + thisvaltrim).attr("checked")) {
           //alert("1")           
           $medarbgodkedboxes.each(function () {
               $medarbgodkedboxes.attr("checked", false);
           });
           $("#medarbgodkend_" + thisvaltrim).attr("checked", false)
       }

       else  {
           //alert("0")          
           $medarbgodkedboxes.each(function () {
               $medarbgodkedboxes.prop("checked", true);
           });
           $("#medarbgodkend_" + thisvaltrim).attr("checked", true)
        }        
  

     }); 

     


    $(".godkendknap").click(function () {

        var $godkedboxes = $('input[name=godkendbox]:checked');
        var godkedboxeslng = $('input[name=godkendbox]:checked').length;
        var antalgodkendt = 0

            $godkedboxes.each(function () {

            antalgodkendt = antalgodkendt + 1

            var thisid = this.id
            var thisvallngt = thisid.length
            var thisvaltrim = thisid.slice(14, thisvallngt)
            
            var thisaktid = this.className
            var thisaktidlengt = thisaktid.length
            var thisaktidtrim = thisaktid.slice(11, thisaktidlengt)

            //alert(thisaktidtrim)

            startdato = $("#startdato").val()
            slutdato = $("#slutdato").val()
            godkendjobid = $("#godkendjobid").val()

            medid = thisvaltrim
            
           // alert(startdato + slutdato)


            $.post("?medid=" + medid + "&startdato=" + startdato + "&slutdato=" + slutdato + "&thisaktidtrim=" + thisaktidtrim + "&godkendjobid=" + godkendjobid, { control: "godkenduge", AjaxUpdateField: "true" }, function (data) {

               //alert("godkendt");
                
            });        


        });

            //alert(antalgodkendt)
            //alert(godkedboxeslng)
        //if (antalgodkendt == godkedboxeslng)
        //  {
        //      $("#godkendform").submit();
        //  }
    });


   


   

    //alert("her")

});



