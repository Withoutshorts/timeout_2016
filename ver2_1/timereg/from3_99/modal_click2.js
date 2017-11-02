




$(document).ready(function () {

   // alert("included")

    $(".modal-click").click(function () {

        var modalid = this.id
        var idlngt = modalid.length
        var idtrim = modalid.slice(6, idlngt)
        //alert(idtrim)
        var modal = document.getElementById('modalOpen_' + idtrim);

        modal.style.display = "block";

        window.onclick = function (event) {
            if (event.target == modal) {
                modal.style.display = "none";
            }
        }
    });



});