
$(window).load(function () {

    
  
});



$(document).ready(function () {


    
    $("#FM_soster").change(function () {


        var thisval = $("#FM_soster option:selected").val()
        
        if (thisval > 0) {
            

            //$("#FM_gruppe option:selected").val(-1)
            $("#FM_gruppe option[value='-1']").attr('selected', 'selected');
        }

    });




   


});