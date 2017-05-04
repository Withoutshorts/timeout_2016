









$(document).ready(function() {


    

   
    //alert("HER")

    ////////////////////// ONLOAD FUNCTIONS  ////////////////////////// 

    jq_lto = $("#jq_lto").val()

    jq_stempeluron = $("#jq_stempeluron").val()

    if (jq_stempeluron == 1) {


        getLocation()

      
    }

   



    // geo gps getLocation
    var x = document.getElementById("demo");
    function getLocation() {


        if (navigator.geolocation) {
            navigator.geolocation.getCurrentPosition(showPosition);
        } else {
            x.innerHTML = "Geolocation is not supported by this browser.";
        }
    }
    function showPosition(position) {

        //x.innerHTML = "Latitude: " + position.coords.latitude +
        //"<br>Longitude: " + position.coords.longitude;

        //alert(position.coords.longitude)

        $("#jq_longitude").val(position.coords.longitude)
        $("#jq_latitude").val(position.coords.latitude)


        jq_longitude = $("#jq_longitude").val()
        jq_latitude = $("#jq_latitude").val()

        jq_lto = $("#jq_lto").val()
        jq_medid = $("#jq_medid").val()

        //alert("HER: " + jq_latitude)

        if ($("#jq_longitude").val() != 0) {


            $.post("?jq_longitude=" + jq_longitude + "&jq_latitude=" + jq_latitude + "&jq_medid=" + jq_medid + "&jq_lto=" + jq_lto, { control: "FN_geolocation", AjaxUpdateField: "true" }, function (data) {

                //$("#FM_job").val(data)

                //alert("Du er her: " + jq_latitude)

            });



        }

    }


});

$(window).load(function() {
    // run code
    
});