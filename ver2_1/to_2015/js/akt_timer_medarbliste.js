




$(document).ready(function () {

    $(".medarbliste").click(function () {
        
        thisid = this.id
        var medarbTR = document.getElementById('tr_medarbliste_' + thisid)

        if (medarbTR.style.display !== 'none')
        {
            medarbTR.style.display = 'none'
        } else {
            medarbTR.style.display = ''
        }

});



});