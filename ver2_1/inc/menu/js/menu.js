









$(document).ready(function () {


    $("#luk_menuslider").mouseover(function () {

        $(this).css("cursor", "pointer");

    });

    

    $("#luk_menuslider").click(function () {

        $(".menu-slider").css("left", "-240px");
        
    });
    

    $(".showLeft").click(function () {
       
        $(".menupkt_n2").hide("fast");
     
        
        var thisid = this.id
        var idlngt = thisid.length
        var idtrim = thisid.slice(5, idlngt)
        var menuid = idtrim

        if (menuid > 0) {

        $(".menu-slider").css("left", "60px");
       
        
        $("#ul_menu-slider-" + menuid).css("display", "");
        $("#ul_menu-slider-" + menuid).css("visibility", "visible");
        $("#ul_menu-slider-" + menuid).show(100);

        }

    });


    $(".menu-slider").click(function () {

        $(".menupkt_n2").hide("fast");
        $(".menu-slider").css("left", "-240px");



    });

    


});





/*
var menuLeft = document.getElementById('menu-slider-1'),
body = document.body;


showLeft.onclick = function () {



    classie.toggle(this, 'active');
    classie.toggle(menuLeft, 'menu-slider-open');
    disableOther('showLeft');
};

function disableOther(button) {
    if (button !== 'showLeft') {
        classie.toggle(showLeft, 'disabled');
    }
}

container.onclick = function () {
    if (showLeft.className == "active") {
        classie.toggle(showLeft, 'active');
        classie.toggle(menuLeft, 'menu-slider-open');
        disableOther('showLeft');
    }
}
*/



