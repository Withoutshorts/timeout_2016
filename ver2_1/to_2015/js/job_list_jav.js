

$(document).ready(function () {

    $('.date').datepicker({

    });

    //alert("HepHep")
    $("#actionMenu").change(function () {

        //alert("Bulk Menu")
        //alert($("#actionMenu").val())

        actionMenuValue = $("#actionMenu").val()

        jobids = ""
        $("input:checkbox[name=job_bulkCHB]:checked").each(function () {

            jobids = jobids + this.id + ","

        });

        // Job Status
        if (actionMenuValue == "3") {

            $("#FM_selectedJobids_status").val(jobids);

            var modal = document.getElementById('actionMenu_jobstatus');

            modal.style.display = "block";

            window.onclick = function (event) {
                if (event.target == modal) {
                    modal.style.display = "none";
                }
            }
        }

        // Job slutdato
        if (actionMenuValue == "2") {

            $("#FM_selectedJobids_enddate").val(jobids);

            var modal = document.getElementById('actionMenu_enddate');

            modal.style.display = "block";

            window.onclick = function (event) {
                if (event.target == modal) {
                    modal.style.display = "none";
                }
            }
        }

        // Forretnings områder
        if (actionMenuValue == "1") {

            $("#FM_selectedJobids_forOm").val(jobids);

            var modal = document.getElementById('actionMenu_forOm');

            modal.style.display = "block";

            window.onclick = function (event) {
                if (event.target == modal) {
                    modal.style.display = "none";
                }
            }
        }

    });



});