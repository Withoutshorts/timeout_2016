

$(document).ready(function () {

    //alert("HEJe23")

    //scrollAbleTable();
    scrollAbleTableOnLoad();
    document.getElementById("screen_overflow").style.overflowX = "inherit";


  /*  var mainTable = document.getElementById("main-table");
    var tableHeight = mainTable.offsetHeight;
    if (tableHeight > 300) {
        var fauxTable = document.getElementById("faux-table");
        document.getElementById("table-wrap").className += ' ' + 'fixedON';
        var clonedElement = mainTable.cloneNode(true);
        clonedElement.id = "";
        fauxTable.appendChild(clonedElement);        
    } else if (tableHeight < 300) {
        fauxTable.parentNode.removeChild(fauxTable);
    } */
    $(".fp_jid").click(function () {

        //scrollAbleTable();
        scrollAbleTableJobClick();

    });

    $("#fitScreen_chb").click(function () {

        if ($('#fitScreen_chb').is(':checked') == false) {
            //alert("Fit screen")
            document.getElementById("screen_overflow").style.overflowX = "inherit";
            //document.getElementById("faux-table").style.width = "400px !important"

        } else
        {
           // alert("Dont Fit screen")
            document.getElementById("screen_overflow").style.overflowX = "scroll";
        }

    });

});


function scrollAbleTableOnLoad() {

    var mainTable = document.getElementById("main-table");
    var tableHeight = mainTable.offsetHeight;
    
    //alert("her1")

    //if (tableHeight > 700) {  
        var fauxTable = document.getElementById("faux-table");
        fauxTable.style.display = "";
        document.getElementById("table-wrap").className += ' ' + 'fixedON';
        var clonedElement = mainTable.cloneNode(true);
        clonedElement.id = "clonedTable";
        fauxTable.appendChild(clonedElement);
        //alert("LAvet")
    //} else
    //{

        if (tableHeight < 700)
        {
            var fauxTable = document.getElementById("faux-table");
            //alert("slette")
            fauxTable.style.display = "none";
            mainTable.style.height = "auto";
            //alert("Removed")
        }
    //}


}

function scrollAbleTableJobClick() {

    var mainTable = document.getElementById("main-table");
    var tableHeight = mainTable.offsetHeight;

    var fauxTable = document.getElementById("faux-table");

    //alert("her1")

    if (tableHeight > 700)
    {
        fauxTable.style.display = "";
        
    } else
    {
        fauxTable.style.display = "none";
        mainTable.style.height = "auto";
    }

}


