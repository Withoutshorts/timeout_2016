


function popUp(URL, width, height, left, top) {
    window.open(URL, 'navn', 'left=' + left + ',top=' + top + ',toolbar=0,scrollbars=1,location=0,statusbar=1,menubar=0,resizable=1,width=' + width + ',height=' + height + '');
}


$(document).ready(function () {

    

    $("#FM_sogkpers").keyup(function () {
        //alert("her")
        hentkpers()
    });


    $("#bt_sogkpers").click(function () {
        //alert("her")
        hentkpers()
     
    });

    function hentkpers() {

        sogVal = $("#FM_sogkpers").val()

        $.post("?jqsogval=" + sogVal, { control: "FN_sogktak", AjaxUpdateField: "true" }, function (data) {

            //alert(data)

            $("#kpers_div").html(data);

        });


    }



});