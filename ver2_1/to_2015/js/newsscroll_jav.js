

$(document).ready(function () {


    setInterval(myMethod, 15);

    newssection = document.getElementById("newssection");
   // alert(newssectio + "  " + newssection.scrollHeight)
    setTimeout(function () {
        scollDirection = 1
    }, 2000);

    function myMethod() {
        //this will repeat every 5 seconds
        //you can reset counter here

        if (scollDirection == 1) {
            $("#newssection").scrollTop($("#newssection").scrollTop() + 1)
        }
        else if (scollDirection = -1)
        {
            $("#newssection").scrollTop($("#newssection").scrollTop() - 1)
        }

        if (newssection.scrollTop == (newssection.scrollHeight - newssection.offsetHeight))
        {
            setTimeout(function () {
                scollDirection = -1
            }, 2000);
        }

        if (newssection.scrollTop <= 0) {
            setTimeout(function () {
                scollDirection = 1
            }, 2000);

        }

        //$("#newssection").scrollTop($("#newssection").scrollTop() + 1);
       

    }




});


