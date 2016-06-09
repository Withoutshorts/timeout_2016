	<!--#include file="../inc/connection/conn_db_inc.asp"-->
	<!--#include file="../inc/errors/error_inc.asp"-->
	<%
	if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="../timereg/inc/dato.asp"-->
	<!--#include file="inc/regular/pda_topmenu_inc.asp"-->
	<!--#include file="../timereg/inc/convertDate.asp"-->

<%
function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
end function



%>
	<div id="sindhold" style="position:absolute; left:10; top:48; visibility:visible;">
	<table cellspacing="0" cellpadding="0" border="0" width="166">
	<tr>
    <td valign="top">
			 <!-------------------------------Sideindhold------------------------------------->
			<br><img src="../ill/header_reg.gif" alt="" border="0"><hr align="left" width="166" size="1" color="#000000" noshade>
		</td>
</tr>
</table>
<br>
<%
'************************** Finder hvilke projektgrupper brugeren er medlem af **********
		 
		f = 0
		Redim brugergruppeId(f)
		
		strSQL = "SELECT Mnavn, Mid, type, ProjektgruppeId, MedarbejderId FROM medarbejdere, medarbejdertyper, progrupperelationer WHERE medarbejdere.Mid="& session("Mid")&" AND medarbejdertyper.id = medarbejdere.Medarbejdertype AND MedarbejderId = Mid" 
			oRec.Open strSQL, oConn, 0, 1 
			
			While Not oRec.EOF
			Redim preserve brugergruppeId(f) 
				strMnavn = oRec("Mnavn")
				strMnr = oRec("Mid")
				strMType = oRec("type")
				brugergruppeId(f) = oRec("ProjektgruppeId")
				f = f + 1
				oRec.MoveNext
			Wend
		
		oRec.Close

%>

<% 
 '** Sætter jobnr på det sidst udskrevne job. ***
  areJobPrinted = 0
  strLastjobnr = 0
  jobPrinted = "0#"
  aktPrinted = "0#"
  func = request("func")
  id = request("id")
  
  select case func
  case "step2"
  '**************************** Rettigheds tjeck aktiviteter **************************
  	for intcounter = 0 to f - 1  
  
	strSQLkri3 = strSQLkri3 &" aktiviteter.job = "& id &" AND aktiviteter.projektgruppe1 = "& brugergruppeId(intcounter) &""_
	&" OR aktiviteter.job = "& id &" AND aktiviteter.projektgruppe2 = "& brugergruppeId(intcounter) &""_
	&" OR aktiviteter.job = "& id &" AND aktiviteter.projektgruppe3 = "& brugergruppeId(intcounter) &""_
	&" OR aktiviteter.job = "& id &" AND aktiviteter.projektgruppe4 = "& brugergruppeId(intcounter) &""_
	&" OR aktiviteter.job = "& id &" AND aktiviteter.projektgruppe5 = "& brugergruppeId(intcounter) &" OR "
	
	next
  	
	'** Trimmer sql states ***
	strSQLkri3_len = len(strSQLkri3)
	strSQLkri3_left = strSQLkri3_len - 3
	strSQLkri3_use = left(strSQLkri3, strSQLkri3_left) 
  	strSQLkri3 = strSQLkri3_use	
  
  %>
  <table cellspacing="0" cellpadding="0" border="0" width="166">
  <tr>
	<td>Søn</td>
	<td>Man</td>
	<td>Tir</td>
	<td>Ons</td>
	<td>Tor</td>
	<td>Fre</td>
	<td>Lør</td>
</tr>
  <%
  strSQL = "SELECT job.id AS id, jobnr, jobnavn, jobknr, kkundenavn, aktiviteter.id AS akt, aktiviteter.navn AS navn, aktiviteter.job FROM job, kunder, aktiviteter WHERE " & strSQLkri3 & " ORDER BY navn" 

	oRec.cachesize = 100
	'Response.write (strSQL)
	oRec.Open strSQL, oConn, 0, 1 
	
	lastAktid = 0
	While Not oRec.EOF
	
					if lastAktid <> oRec("akt") then
					%>
					<tr><td colspan="7" bgcolor="#5582D2" height="1"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td></tr>	
					<tr>
						<td valign=top colspan="7">&nbsp;&nbsp;<b><%=left(oRec("navn"), 12)%></b>&nbsp;&nbsp;</td></tr>
					<tr>
						<td><input type="text" name="FM_timer_val" size="2" maxlength="4"></td>
						<td><input type="text" name="FM_timer_val" size="2" maxlength="4"></td>
						<td><input type="text" name="FM_timer_val" size="2" maxlength="4"></td>
						<td><input type="text" name="FM_timer_val" size="2" maxlength="4"></td>
						<td><input type="text" name="FM_timer_val" size="2" maxlength="4"></td>
						<td><input type="text" name="FM_timer_val" size="2" maxlength="4"></td>
						<td><input type="text" name="FM_timer_val" size="2" maxlength="4"></td>
					</tr>
					<%
					end if
					
	lastAktid = oRec("akt")
	oRec.movenext
	wend
	oRec.close
  
  case else
  '**************************** Rettigheds tjeck **************************
  	for intcounter = 0 to f - 1  
  
  	strSQLkri = strSQLkri &" jobstatus = 1 AND fakturerbart = 1 AND kunder.Kid = job.jobknr AND job.projektgruppe1 = "&brugergruppeId(intcounter)&""_
	&" OR "_
	&" jobstatus = 1 AND fakturerbart = 1 AND kunder.Kid = job.jobknr AND job.projektgruppe2 = "&brugergruppeId(intcounter)&""_
	&" OR "_
	&" jobstatus = 1 AND fakturerbart = 1 AND kunder.Kid = job.jobknr AND job.projektgruppe3 = "&brugergruppeId(intcounter)&""_
	&" OR "_
	&" jobstatus = 1 AND fakturerbart = 1 AND kunder.Kid = job.jobknr AND job.projektgruppe4 = "&brugergruppeId(intcounter)&""_
	&" OR "_
  	&" jobstatus = 1 AND fakturerbart = 1 AND kunder.Kid = job.jobknr AND job.projektgruppe5 = "&brugergruppeId(intcounter)&""_
	&" OR "
	
	next
  	
	'** Trimmer de 2 sql states ***
	strSQLkri_len = len(strSQLkri)
	strSQLkri_left = strSQLkri_len - 3
	strSQLkri_use = left(strSQLkri, strSQLkri_left) 
  	strSQLkri = strSQLkri_use
  	
	%>
	<table cellspacing="0" cellpadding="0" border="0" width="166">
	<tr bgcolor="#5582D2">
		<td width="3"><img src="../ill/venstre_hjorne.gif" alt="" border="0"></td>
		<td style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 0px; border-color : #003399; border-style : solid;"><font class="stor-hvid">&nbsp;Eksterne job</font></td>
		<td  width="3" align="right"><img src="../ill/hojre_hjorne.gif" alt="" border="0"></td>
	</tr>
	</table>

	<table cellspacing="0" cellpadding="0" border="0" width="166">
	<%
	strSQL = "SELECT job.id AS id, jobnr, jobnavn, jobknr, kkundenavn FROM job, kunder WHERE " & strSQLkri & " ORDER BY jobnavn" 
	
	oRec.cachesize = 100
	'Response.write (strSQL)
	oRec.Open strSQL, oConn, 0, 1 
	
	lastJobid = 0
	While Not oRec.EOF
	
					if lastJobid <> oRec("id") then
					%>
					<tr><td colspan="3" bgcolor="#5582D2" height="1"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td></tr>	
					<tr>
						<td valign=top>&nbsp;&nbsp;<%=oRec("Jobnr")%>&nbsp;&nbsp;</td>
						<td valign=top>&nbsp;&nbsp;<b><%=left(oRec("Jobnavn"), 12)%></b>&nbsp;&nbsp;<font size="1">(<%=left(oRec("kkundenavn"), 7)%>)</font></td><td><a href='pda_timereg.asp?func=step2&vmenu=tsa&id=<%=oRec("id")%>'><img src='ill/pillillexp.gif' alt='' width='16' height='14' border='0'></a></td>
					</tr>
					<%
					end if
					
	lastJobid = oRec("id")
	oRec.movenext
	wend
	oRec.close
	
	end select
	%>
	</table>

</div>
<%
end if
%>
<!--#include file="../inc/regular/footer_inc.asp"-->
 




