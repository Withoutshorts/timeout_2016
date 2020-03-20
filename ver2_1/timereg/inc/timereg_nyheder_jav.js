$(document).ready(function () {
   

    $("#closenews").click(function () {

        // alert("hep")

        $.post("?jq_kid=1", { control: "FN_nyhederlaest", AjaxUpdateField: "true", cust: 0 }, function (data) {

            // alert("Kalst")
        });


        document.getElementById('newsmodal').style.display = "none";

    });

    $(".nyhedsbillede").click(function () {

        var modalid = this.id
        var idlngt = modalid.length
        var idtrim = modalid.slice(6, idlngt)

        //var modalidtxt = $("#myModal_" + idtrim);
        var modal = document.getElementById('nyhedsbillede_' + idtrim);

        modal.style.display = "block";

        window.onclick = function (event) {
            if (event.target == modal) {
                modal.style.display = "none";
            }
        }

    });

});