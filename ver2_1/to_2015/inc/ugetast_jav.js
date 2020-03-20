$(document).ready(function () {

    var feltnr = $("#lastfeltnr").val();
    feltnr = parseInt(feltnr)

    $(".addRow").click(function () {
        feltnr = feltnr + 1
        thisdate = $(this).attr('data-thisdate')
        thisid = this.id
        var newRow = "<tr>"
        newRow += "<td>"
        newRow += "<input type='hidden' name='FM_feltnr' value='" + feltnr + "' />"
        newRow += "<input type='hidden' name='FM_dato_" + feltnr + "' value='" + thisdate + "' />"
        newRow += "<input type='hidden' name='FM_timeid_" + feltnr +"' value='-1' />"

        newRow += "<input type='time' name='FM_sttid_" + feltnr + "' class='form-control input-small' /></td>"
        newRow += "<td><input type='time' name='FM_sltid_" + feltnr + "' class='form-control input-small' /></td>"

        newRow += "<td><select name='FM_jobnr_" + feltnr + "' id='" + feltnr + "' class='form-control input-small jobSEL'>" + $("#defaultjobopr").val() + "</select></td>"
        newRow += "<td><select name='FM_aktid_" + feltnr + "' class='form-control input-small' id='aktSEL_" + feltnr + "'><option value='-1'>-</option></select></td>"
        newRow += "</tr>"
       
        $(".tr_" + thisid).each(function (index) {
            lastIndex = index
        });
       // alert(lastIndex)
        $(".tr_" + thisid).each(function (index) {
            if (index == lastIndex) {
                $(this).after(newRow)
            }
        });


       // $(".tr_" + thisid).after(newRow)

        var currentRowspan = $("#title_" + thisid).attr('rowspan')
        var newRowspan = parseInt(currentRowspan) + 1
        $("#title_" + thisid).attr('rowspan', newRowspan)
    });

    $(document).on('change', '.jobSEL', function () {
        thisid = this.id;

        jobnr = $(this).val();

        $.post("?jobnr=" + jobnr, { control: "findaktiviteter", AjaxUpdateField: "true" }, function (data) {
            $("#aktSEL_" + thisid).html(data)
        });

        
    });

});


