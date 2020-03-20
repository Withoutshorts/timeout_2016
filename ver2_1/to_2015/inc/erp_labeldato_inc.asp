


<script type="text/javascript">
      function IsValidDate(input) {
            var dayfield = input.split(".")[0];
            var monthfield = input.split(".")[1];
            var yearfield = input.split(".")[2];
            var returnval;
                        var dayobj = new Date(yearfield, monthfield - 1, dayfield)
                        if ((dayobj.getMonth() + 1 != monthfield) || (dayobj.getDate() != dayfield) || (dayobj.getFullYear() != yearfield)) {
                              alert("Forkert datoformat, skriv venligst dd.mm.yyyy eller d.m.yyyy");
                              returnval = false;
                        }
                        else {
                              returnval = true;
                        }
                        return returnval;
                  }
    $(document).ready(function () {

                    
                        //Make datepickers
                        $.datepicker.setDefaults($.extend({ showOn: 'button', constrainInput: true, showOtherMonths: true, showWeeks: true, minDate: new Date(2002, 1, 1), firstDay: 1, changeFirstDay: false, buttonImage: '../ill/popupcal.gif', start: 6, buttonImageOnly: true, dateFormat: 'd.m.yy', changeMonth: true, changeYear: true }));

                       
                            $("#FM_labeldato").datepicker();
                       

                        $("#FM_labeldato").change(function() { if (IsValidDate($("#FM_labeldato").val()) == true) { var datestring = $(this).val(); var datesplit = datestring.split('.'); $("[name=FM_labeldato_dag]").val(datesplit[0]); $("[name=FM_labeldato_mrd]").val(datesplit[1]); $("[name=FM_labeldato_aar]").val(datesplit[2]); } else { alert("Dato ikke gemt"); } });

                        //hide unused fields and show relevant
                        $(".popupcalImg").hide();
                        //$("[name=FM_labeldato_dag]").hide(); $("[name=FM_labeldato_mrd]").hide(); $("[name=FM_labeldato_aar]").hide();
                        $("#FM_labeldato").css("display", "inline");
                  });
</script>
<input type="text" id="FM_labeldato" value="<%=day(labelDato)%>.<%=month(labelDato)%>.<%=year(labelDato)%>" style="display:none; margin-right:5px; width:80px;" />

<!--
<input type="hidden" id="FM_labeldato" value="<%=day(istSlutDato)%>.<%=month(istSlutDato)%>.<%=year(istSlutDato)%>" style="display:none; margin-right:5px; width:70px;" />
-->
<input name="FM_labeldato_dag" id="FM_labeldato_dag" type="hidden" value="<%=day(labelDato)%>" />
<input name="FM_labeldato_mrd" id="FM_labeldato_mrd" type="hidden" value="<%=month(labelDato)%>" />
<input name="FM_labeldato_aar" id="FM_labeldato_aar" type="hidden" value="<%=year(labelDato)%>" />

		
				
<!-- Weekselecter SLUT -->
