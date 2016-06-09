


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
                  $(document).ready(function() {
                        //Make datepickers
                        $.datepicker.setDefaults($.extend({ showOn: 'button', constrainInput: true, showOtherMonths: true, showWeeks: true, minDate: new Date(2002, 1, 1), firstDay: 1, changeFirstDay: false, buttonImage: '../ill/popupcal.gif', start: 6, buttonImageOnly: true, dateFormat: 'd.m.yy', changeMonth: true, changeYear: true }));

                        $("#FM_istdato2").datepicker();
                        $("#FM_istdato2").change(function() { if (IsValidDate($("#FM_istdato2").val()) == true) { var datestring = $(this).val(); var datesplit = datestring.split('.'); $("[name=FM_istdato2_dag]").val(datesplit[0]); $("[name=FM_istdato2_mrd]").val(datesplit[1]); $("[name=FM_istdato2_aar]").val(datesplit[2]); } else { alert("Dato ikke gemt"); } });

                        //hide unused fields and show relevant
                        $(".popupcalImg").hide();
                        //$("[name=FM_istdato2_dag]").hide(); $("[name=FM_istdato2_mrd]").hide(); $("[name=FM_istdato2_aar]").hide();
                        $("#FM_istdato2").css("display", "inline");
                  });
</script>
<input type="text" id="FM_istdato2" value="<%=day(istSlutDato)%>.<%=month(istSlutDato)%>.<%=year(istSlutDato)%>" style="display:none; margin-right:5px; width:80px;" />

<!--
<input type="hidden" id="FM_istdato2" value="<%=day(istSlutDato)%>.<%=month(istSlutDato)%>.<%=year(istSlutDato)%>" style="display:none; margin-right:5px; width:70px;" />
-->
<input name="FM_istdato2_dag" id="FM_istdato2_dag" type="hidden" value="<%=day(istSlutDato)%>" />
<input name="FM_istdato2_mrd" id="FM_istdato2_mrd" type="hidden" value="<%=month(istSlutDato)%>" />
<input name="FM_istdato2_aar" id="FM_istdato2_aar" type="hidden" value="<%=year(istSlutDato)%>" />

		
				
<!-- Weekselecter SLUT -->
