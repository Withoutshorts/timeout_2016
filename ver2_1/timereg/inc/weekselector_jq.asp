



<!--
Bruges bl.a af:




DENNE FILK SKAL IKKE BRUGE ERP DATO cookie
BRUGES KUN AF webblik_joblisten_21 (Gantt)

-->


<% 

select case thisfile
case "erp_oprfak_fs", "fak_serviceaft_osigt.asp", "erp_tilfakturering.asp", "erp_opr_faktura.asp", "erp_serviceaft_saldo.asp", "erp_fak_rykker", "erp_opr_faktura", "erp_opr_faktura_kontojob"
useErpCookie = 1
case else
useErpCookie = 0
end select

'Response.Write "strDag: " & request("FM_start_dag")
'Response.write "thisfile: " & thisfile & "<br>"

if thisfile = "fak_serviceaft_osigt.asp" OR thisfile = "joblog" OR _
thisfile = "serviceaft_osigt.asp" OR thisfile = "materiale_stat.asp" OR _
thisfile = "materialer_ordrer.asp" OR thisfile = "stempelur" OR thisfile = "fomr" OR _
thisfile = "webblik_joblisten.asp" OR thisfile = "webblik_joblisten21.asp" OR thisfile = "webblik_milepale.asp" OR _
thisfile = "erp_tilfakturering.asp" OR thisfile = "webblik_tilfakturering.asp" OR _
thisfile = "webblik.asp" OR thisfile = "joblog_k" OR thisfile = "erp_opr_faktura.asp" OR _
thisfile = "erp_serviceaft_saldo.asp" OR thisfile = "crmhistorik" OR thisfile = "smileystatus.asp" OR thisfile = "stopur_2008.asp" then
      if len(request("FM_usedatokri")) <> 0 then
      useDatofilter = 1
      else
      useDatofilter = 0
      end if
else
      if thisfile = "joblog_timberebal" OR thisfile = "erp_fak_rykker" OR thisfile = "erp_opr_faktura" then
      useDatofilter = 1
      else
      useDatofilter = cint(request("FM_usedatointerval"))
      end if
end if

'Response.write thisfile & "<br>"
'Response.write "useDatofilter" & useDatofilter & "bruglistedato" &   bruglistedato & "    request(FM_slut_mrd):" & request("FM_slut_mrd")

'Response.flush

if useDatofilter = 1 then

if cint(bruglistedato) = 1 then

     strSQLgdt = "SELECT datomidtpkt FROM gantts WHERE id ="& gantt_liste

     'Response.write strSQLgdt
     'Response.flush
     oRec6.open strSQLgdt, oConn, 3
     if not oRec6.EOF then


        strMrd = month(oRec6("datomidtpkt"))
        strDag = day(oRec6("datomidtpkt"))
        strAar = year(oRec6("datomidtpkt")) 
        
        strDag_slut = strDag
        strMrd_slut = strMrd
        strAar_slut = strAar
     
     end if
     oRec6.close
      

   



else

strMrd = request("FM_start_mrd")
strDag = request("FM_start_dag")
strAar = right(request("FM_start_aar"),4) 
strDag_slut = request("FM_slut_dag")
strMrd_slut = request("FM_slut_mrd")
strAar_slut = right(request("FM_slut_aar"),4)


end if


else
    
    'Response.Write "her"
      '*Brug cookie eller dagsdato?
      if (len(Request.Cookies("datoer")("st_md")) <> 0 AND useErpCookie = 0) OR (len(Request.Cookies("erp_datoer")("st_md")) <> 0 AND useErpCookie = 1) then
      'if Request.Cookies("st_dag").haskeys then      
            
            if useErpCookie = 1 then
            strMrd = Request.Cookies("erp_datoer")("st_md")
            strDag = Request.Cookies("erp_datoer")("st_dag")
            strAar = Request.Cookies("erp_datoer")("st_aar") 
                
                    'if thisfile = "erp_tilfakturering.asp" then
                  '   strDag_slut = day(now()) 
                    '    strMrd_slut = month(now()) 
                    '    strAar_slut = right(year(now()),2) 
                    'else
                        strDag_slut = Request.Cookies("erp_datoer")("sl_dag")
                        strMrd_slut = Request.Cookies("erp_datoer")("sl_md")
                        strAar_slut = Request.Cookies("erp_datoer")("sl_aar")
                    'end if
                    
            else


            
            if Request.Cookies("datoer")("st_md") <> "" then
            
            strMrd = Request.Cookies("datoer")("st_md")
            strDag = Request.Cookies("datoer")("st_dag")
            strAar = Request.Cookies("datoer")("st_aar") 
                    
            strDag_slut = Request.Cookies("datoer")("sl_dag")
            strMrd_slut = Request.Cookies("datoer")("sl_md")
            strAar_slut = Request.Cookies("datoer")("sl_aar")

            else

            strMrd = month(now)
            strDag = day(now)
            strAar = datepart("yyyy", now)
                    
            strDag_slut = strDag 'Request.Cookies("datoer")("sl_dag")
            strMrd_slut = strMrd 'Request.Cookies("datoer")("sl_md")
            strAar_slut = strAar 'Request.Cookies("datoer")("sl_aar")

            end if

            end if
            
            'StrTdato = date-31
            'StrUdato = date 
            
      else
      
            StrTdato = date-31
            StrUdato = date 
            
            '* Bruges i weekselector *
            if month(now()) = 1 then
            strMrd = 12
            else
            strMrd = month(now()) - 1
            end if
            
            strDag = day(now())
            
            if month(now()) = 1 then
            strAar = right(year(now()) - 1, 2)
            else
            strAar = right(year(now()), 2) 
            end if
            
            strMrd_slut = month(now())
            strAar_slut = right(year(now()), 2) 
            
            if strDag > "28" then
            strDag_slut = "1"
            strMrd_slut = strMrd_slut + 1
            else
            strDag_slut = day(now())
            end if
      
      end if
      
end if


                  Select case strMrd
                  case 4, 6, 9, 11
                        if strDag = 31 then
                        strDag = 30
                        else
                        strDag = strDag
                        end if
                  case 2
                        if strDag > 28 then
                            select case strAar
                            case "04","08", "12", "16", "20", "24", "28", "32", "36", "40", "44", "48", "52", "56", "60"
                            strDag = 29
                            case else
                            strDag = 28
                            end select
                        else
                        strDag = strDag
                        end if
                  case else
                  strDag = strDag
                  end select
                  
                  
                  Select case strMrd_slut
                  case 4, 6, 9, 11
                        if strDag_slut = 31 then
                        strDag_slut = 30
                        else
                        strDag_slut = strDag_slut
                        end if
                  case 2
                        if strDag_slut > 28 then
                            
                            select case strAar_slut
                            case "04","08", "12", "16", "20", "24", "28", "32", "36", "40", "44", "48", "52", "56", "60"
                            strDag_slut = 29
                            case else
                            strDag_slut = 28
                            end select
                            
                        
                        else
                        strDag_slut = strDag_slut
                        end if
                  case else
                  strDag_slut = strDag_slut
                  end select

'** Indsætter cookie **



'if useErpCookie = 1 then
'Response.Cookies("erp_datoer")("st_dag") = strDag
'Response.Cookies("erp_datoer")("st_md") = strMrd
'Response.Cookies("erp_datoer")("st_aar") = strAar
'Response.Cookies("erp_datoer")("sl_dag") = strDag_slut
'Response.Cookies("erp_datoer")("sl_md") = strMrd_slut
'Response.Cookies("erp_datoer")("sl_aar") = strAar_slut
'Response.Cookies("erp_datoer").Expires = date + 10    
'else
Response.Cookies("datoer")("st_dag") = strDag
Response.Cookies("datoer")("st_md") = strMrd
Response.Cookies("datoer")("st_aar") = strAar
Response.Cookies("datoer")("sl_dag") = strDag_slut
Response.Cookies("datoer")("sl_md") = strMrd_slut
Response.Cookies("datoer")("sl_aar") = strAar_slut
Response.Cookies("datoer").Expires = date + 10        
'end if
%>                
<!--#include file="dato2_b.asp"-->
<%
if dontshowDD <> "1" then
%>
<script src="../inc/jquery/ui.datepicker.js" type="text/javascript"></script>


<script type="text/javascript">
      function IsValidDate(input) {
            var dayfield = input.split("/")[0];
            var monthfield = input.split("/")[1];
            var yearfield = input.split("/")[2];

            //alert(yearfield)

            //if (yearfield.length == 2) {
            //    yearfield = "20" + yearfield

            //}

            var returnval;
                        var dayobj = new Date(yearfield, monthfield - 1, dayfield)
                        if ((dayobj.getMonth() + 1 != monthfield) || (dayobj.getDate() != dayfield) || (dayobj.getFullYear() != yearfield)) {
                              alert("Forkert datoformat, skriv venligst dd/mm/yyyy eller d/m/yyyy");
                              returnval = false;
                        }
                        else {
                              returnval = true;
                        }
                        return returnval;
                  }
                  $(document).ready(function() {
                        //Make datepickers
                        $.datepicker.setDefaults($.extend({ showOn: 'button', constrainInput: true, showOtherMonths: true, showWeeks: true, minDate: new Date(2002, 1, 1), firstDay: 1, changeFirstDay: false, buttonImage: '../ill/popupcal.gif', start: 6, buttonImageOnly: true, dateFormat: 'd/m/yy', changeMonth: true, changeYear: true }));

                        $("#FM_startdate").datepicker();
                        $("#FM_startdate").change(function() { if (IsValidDate($("#FM_startdate").val()) == true) { var datestring = $(this).val(); var datesplit = datestring.split('/'); $("#FM_start_dag").val(datesplit[0]); $("#FM_start_mrd").val(datesplit[1]); $("#FM_start_aar").val(datesplit[2]); } else { alert("dato ikke gemt"); } });

                        $("#FM_enddate").datepicker();
                        $("#FM_enddate").change(function() { if (IsValidDate($("#FM_enddate").val()) == true) { var datestring = $(this).val(); var datesplit = datestring.split('/'); $("#FM_slut_dag").val(datesplit[0]); $("#FM_slut_mrd").val(datesplit[1]); $("#FM_slut_aar").val(datesplit[2]); } });

                        //hide unused fields and show relevant
                        $(".popupcalImg").hide();
                        $("#FM_start_dag").hide(); $("#FM_start_mrd").hide(); $("#FM_start_aar").hide();
                        $("#FM_startdate").css("display", "inline");
                        $("#FM_slut_dag").hide(); $("#FM_slut_mrd").hide(); $("#FM_slut_aar").hide();
                        $("#FM_enddate").css("display", "inline");

                  });
</script>

<%if len(strAar) < 4 then 
strAar = "20" & strAar 
end if%>

<%if len(strAar_slut) < 4 then 
strAar_slut = "20" & strAar_slut 
end if%>

<span id="fraLabel">Fra:&nbsp;<input type="text" id="FM_startdate" value="<%=strDag%>/<%=strMrd%>/<%=strAar%>" style="display:none; background-color : #ffffff; border : thin black; font : 10px verdana;" />
<select name="FM_start_dag" id="FM_start_dag" style="background-color : #ffffff; border : thin black; font : 10px verdana;">
            <option value="<%=strDag%>"><%=strDag%></option> 
            <option value="1">1</option>
            <option value="2">2</option>
            <option value="3">3</option>
            <option value="4">4</option>
            <option value="5">5</option>
            <option value="6">6</option>
            <option value="7">7</option>
            <option value="8">8</option>
            <option value="9">9</option>
            <option value="10">10</option>
            <option value="11">11</option>
            <option value="12">12</option>
            <option value="13">13</option>
            <option value="14">14</option>
            <option value="15">15</option>
            <option value="16">16</option>
            <option value="17">17</option>
            <option value="18">18</option>
            <option value="19">19</option>
            <option value="20">20</option>
            <option value="21">21</option>
            <option value="22">22</option>
            <option value="23">23</option>
            <option value="24">24</option>
            <option value="25">25</option>
            <option value="26">26</option>
            <option value="27">27</option>
            <option value="28">28</option>
            <option value="29">29</option>
            <option value="30">30</option>
            <option value="31">31</option></select>&nbsp;&nbsp;
            
            <select name="FM_start_mrd" id="FM_start_mrd" style="background-color : #ffffff; border : thin black; font : 10px verdana;">
            <option value="<%=strMrd%>"><%=strMrdNavn%></option>
            <option value="1">jan</option>
            <option value="2">feb</option>
            <option value="3">mar</option>
            <option value="4">apr</option>
            <option value="5">maj</option>
            <option value="6">jun</option>
            <option value="7">jul</option>
            <option value="8">aug</option>
            <option value="9">sep</option>
            <option value="10">okt</option>
            <option value="11">nov</option>
            <option value="12">dec</option></select>&nbsp;&nbsp;
            <select name="FM_start_aar" id="FM_start_aar" style="background-color : #ffffff; border : thin black; font : 10px verdana;">
            <option value="<%=strAar%>"><%=strAar%></option>
            
            <%for x = -5 to 10 
            useY = datepart("yyyy", dateadd("yyyy", x, date()))%>
            <option value="<%=useY%>"><%=right(useY, 2)%></option>
            <%next %>
            </select>&nbsp;<a href="javascript:NewWin_popupcal('../inc/regular/popupcalender_inc.asp?use=1')"><img src="../ill/popupcal.gif" alt="" border="0" width="16" height="15" class="popupcalImg"></a></span>
            
            
            
            
            <span id="tilLabel"><br>Til:&nbsp;<input type="text" value="<%=strDag_slut%>/<%=strMrd_slut%>/<%=strAar_slut%>" id="FM_enddate" style="display:none; background-color : #ffffff; border : thin black; font : 10px verdana; margin-left:6px;"/><img src="../ill/blank.gif" width="6" height="1" alt="" border="0"> 
            <select name="FM_slut_dag" id="FM_slut_dag" style="background-color : #ffffff; border : thin black; font : 10px verdana;">
            <option value="<%=strDag_slut%>"><%=strDag_slut%></option> 
            <option value="1">1</option>
            <option value="2">2</option>
            <option value="3">3</option>
            <option value="4">4</option>
            <option value="5">5</option>
            <option value="6">6</option>
            <option value="7">7</option>
            <option value="8">8</option>
            <option value="9">9</option>
            <option value="10">10</option>
            <option value="11">11</option>
            <option value="12">12</option>
            <option value="13">13</option>
            <option value="14">14</option>
            <option value="15">15</option>
            <option value="16">16</option>
            <option value="17">17</option>
            <option value="18">18</option>
            <option value="19">19</option>
            <option value="20">20</option>
            <option value="21">21</option>
            <option value="22">22</option>
            <option value="23">23</option>
            <option value="24">24</option>
            <option value="25">25</option>
            <option value="26">26</option>
            <option value="27">27</option>
            <option value="28">28</option>
            <option value="29">29</option>
            <option value="30">30</option>
            <option value="31">31</option></select>&nbsp;&nbsp;
            <select name="FM_slut_mrd" id="FM_slut_mrd" style="background-color : #ffffff; border : thin black; font : 10px verdana;">
            <option value="<%=strMrd_slut%>"><%=strMrdNavn_slut%></option>
            <option value="1">jan</option>
            <option value="2">feb</option>
            <option value="3">mar</option>
            <option value="4">apr</option>
            <option value="5">maj</option>
            <option value="6">jun</option>
            <option value="7">jul</option>
            <option value="8">aug</option>
            <option value="9">sep</option>
            <option value="10">okt</option>
            <option value="11">nov</option>
            <option value="12">dec</option></select>&nbsp;&nbsp;
            <select name="FM_slut_aar" id="FM_slut_aar" style="background-color : #ffffff; border : thin black; font : 10px verdana;">
            <option value="<%=strAar_slut%>"><%=strAar_slut%></option>
            <%for x = -5 to 10
            useY = datepart("yyyy", dateadd("yyyy", x, date()))%>
            <option value="<%=useY%>"><%=right(useY, 2)%></option>
            <%next %>
            </select>&nbsp;<a href="javascript:NewWin_popupcal('../inc/regular/popupcalender_inc.asp?use=4')"><img src="../ill/popupcal.gif" alt="" border="0" width="16" height="15" class="popupcalImg"></a></span>
                        
<!-- Weekselecter SLUT -->
<%end if %>


 
 
 

 
 

