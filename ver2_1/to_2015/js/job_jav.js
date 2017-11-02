




$(document).ready(function () {


    

    $(".advanced_filter").click(function () {


        var advancedFilter = document.getElementById('udvidet_sog');
        var ud_s_span = document.getElementById('ud_s_span');

        if (advancedFilter.style.display == "none")
        {
            advancedFilter.style.display = "";
            ud_s_span.innerHTML = "-";
        } else
        {
            advancedFilter.style.display = "none";
            ud_s_span.innerHTML = "+";
        }


    });


});



