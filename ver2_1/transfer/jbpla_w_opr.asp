<%Response.buffer = true%>
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
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<%
	
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	
	select case func
	case "oprdb", "reddb"
	
	sttid = request("FM_sttid")
	sltid = request("FM_sltid")
	
	jobid = request("FM_jobid")
	medid = request("FM_mid")
	dato = request("FM_dato")
	aktid = request("FM_aktid")
	
	strSQLA = "SELECT id, aktstartdato, aktslutdato FROM aktiviteter WHERE id = "& aktid
	'Response.write strSQLA
	'Response.flush
	oRec.open strSQLA, oConn, 3 
	if not oRec.EOF then
	aktstkri = oRec("aktstartdato")
	aktslkri = oRec("aktslutdato")
	end if
	oRec.close
	
	if cdate(formatdatetime(aktstkri, 2)) > cdate(formatdatetime(dato, 2)) OR cdate(formatdatetime(aktslkri, 2)) < cdate(formatdatetime(dato, 2)) then
	%>
	<div name="tb" id="tb" style="position:absolute; display:; visibility:visible; left:20; top:20; background-color:#ffffff;">
	Den valgte dato ligger udenfor den valgte aktivitets start og slut dato!<br>
	Den valgte dato er: <%=formatdatetime(dato, 2)%> <br><br>
	Den aktive periode for aktiviteten er:  <%=formatdatetime(aktstkri, 2) %> til <%=formatdatetime(aktslkri, 2)%><br><br><br>
	<a href="Javascript:history.back()"><a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a><br>&nbsp;
	</div>
	<%
	else
	
	'Response.write dato&" "&sttid &", "& dato&" "&sltid
	
	timerthis_temp = datediff("n", sttid, sltid)
	timerthis_temp = (timerthis_temp/60)
	
	timerthis = replace(timerthis_temp, ",",".")
	
	sttid = formatdatetime(sttid, 3)
	sltid = formatdatetime(sltid, 3)
	
	if func = "reddb" then
	strSQL = "UPDATE ressourcer SET jobid = "& jobid &", aktid = "& aktid&", timer = "& timerthis &", mid = "& medid &", starttp = '"& sttid &"', sluttp = '"& sltid &"', dato = '"& dato &"' WHERE id ="&id
	else
	strSQL = "INSERT INTO ressourcer SET jobid = "& jobid &", aktid = "& aktid&", timer = "& timerthis &", mid = "& medid &", starttp = '"& sttid &"', sluttp = '"& sltid &"', dato = '"& dato &"'"
	end if
	
	oConn.execute(strSQL)
	Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
	Response.Write("<script language=""JavaScript"">window.close();</script>")
   
   	end if
	
	case "step2"
	
	if len(request("aid")) <> 0 then
	aid = request("aid")
	else
	aid = 0
	end if
	
	jobid = request("FM_opr_jobid")
	medid = request("FM_opr_mid")
	dato = request("FM_start_aar")&"/"&request("FM_start_mrd")&"/"&request("FM_start_dag")
	
	sttid = request("klok_timer")&":"&request("klok_min")&":00"
	sltid = request("klok_timer_slut")&":"&request("klok_min_slut")&":00"
	
	if formatdatetime(sltid, 3) <= formatdatetime(sttid, 3) then
	sltid = dateadd("n", 15, dato&" "&sttid)
	else
	sltid = sltid
	end if
	
	if id <> 0 then
	func = "reddb"
	funcpil = "opdaterpil"
	else
	func = "oprdb"
	funcpil = "opretpil"
	end if
	%>
	<div name="step2" id="step2" style="position:absolute; display:; visibility:visible; left:20; top:20; background-color:#eff3ff; width:220; z-index:10000; border:1px #5582d2 solid; padding:5px;">
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF" width="300">
	<form action="jbpla_w_opr.asp?func=<%=func%>" method="post" name="opr" id="opr">
	<input type="hidden" name="id" id="id" value="<%=id%>">
	<input type="hidden" name="FM_dato" id="FM_dato" value="<%=dato%>">
	<input type="hidden" name="FM_sttid" id="FM_sttid" value="<%=sttid%>">
	<input type="hidden" name="FM_sltid" id="FM_sltid" value="<%=sltid%>">
	<input type="hidden" name="FM_mid" id="FM_mid" value="<%=medid%>">
	<input type="hidden" name="FM_jobid" id="FM_jobid" value="<%=jobid%>">
	<tr>
		<td align=right style="padding-right:3px;">Aktivitet:</td>
		<td><select name="FM_aktid" id="FM_aktid" style="width:250;">
			<%strSQLJ = "SELECT id, navn FROM aktiviteter WHERE job = "& jobid &" AND aktstatus = 1 AND fakturerbar <> 2 ORDER BY navn"
				oRec.open strSQLJ, oConn, 3 
				while not oRec.EOF 
				if cint(aid) = oRec("id") then
				aidsel = "SELECTED"
				else
				aidsel = ""
				end if%>
				<option value="<%=oRec("id")%>" <%=aidsel%>><%=left(oRec("navn"), 30)%></option>
				<%
				oRec.movenext
				wend
				oRec.close 
				%>
			</select></td>
	</tr>
	<tr><td colspan=2 align=center><br><input type="image" src="../ill/<%=funcpil%>.gif"></td></tr>
	</form>
	</table>
	
	</div>
	<div name="tb" id="tb" style="position:absolute; display:; visibility:visible; left:20; top:120; background-color:#ffffff;">
	<a href="Javascript:history.back()"><a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a><br>&nbsp;
	</div>
	
	<%
	
	
	case else
	%> 
	<script LANGUAGE="javascript">
		function NewWin_popupcal(url)    {
		window.open(url, 'Calpick', 'width=250,height=250,scrollbars=no,toolbar=no');    
		}
	</script>
	<%
	if len(request("medid")) <> 0 then
	medid = request("medid")
	else
	medid = 0
	end if
	
	if len(request("jobid")) <> 0 then
	jobid = request("jobid")
	else
	jobid = 0
	end if
	
	if len(request("aid")) <> 0 then
	aid = request("aid")
	else
	aid = 0
	end if
	
	dato = request("dato")
	strDag = day(dato)
	strMrd = month(dato)
	strAar = year(dato)
	strKlokkeslet = request("sttid")
	strKlokkeslet_slut = dateadd("n", 30, dato&" "&strKlokkeslet)
	
	formatstrKlokkeslet = FormatDateTime(strKlokkeslet,4)
	klok_timer = left(formatstrKlokkeslet,2)
	klok_min = right(formatstrKlokkeslet,2)
	
	formatstrKlokkeslet_slut = FormatDateTime(strKlokkeslet_slut,4)
	klok_timer_slut = left(formatstrKlokkeslet_slut,2)
	klok_min_slut = right(formatstrKlokkeslet_slut,2)
	
	jobstKri = request("datostkri") 
	jobslKri = request("datoslkri")

%>
<!--#include file="inc/dato2_b.asp"-->
<div name="opretny" id="opretny" style="position:absolute; display:; visibility:visible; left:20; top:20; background-color:#eff3ff; width:220; z-index:10000; border:1px #5582d2 solid; padding:5px;">
	
	
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF" width="300">
	<form action="jbpla_w_opr.asp?func=step2" method="post" name="opr" id="opr">
	<input type="hidden" name="id" id="id" value="<%=id%>">
	<input type="hidden" name="aid" id="aid" value="<%=aid%>">
	<tr>
		<td align=right style="padding-right:3px;">Dato:</td>
		<td><select name="FM_start_dag">
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
		<option value="31">31</option></select>
		
		<select name="FM_start_mrd">
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
	   	<option value="12">dec</option></select>
		
		
		<select name="FM_start_aar">
		<option value="<%=strAar%>"><%=strAar%></option>
		<option value="2002">2002</option>
		<option value="2003">2003</option>
	   	<option value="2004">2004</option>
	   	<option value="2005">2005</option>
		<option value="2006">2006</option>
		<option value="2007">2007</option></select>&nbsp;&nbsp;<a href="javascript:NewWin_popupcal('../inc/regular/popupcalender_inc.asp?use=1')"><img src="../ill/popupcal.gif" alt="" border="0" width="16" height="15"></a>
		</td>
	</tr>	
			
	
	<tr><td align=right style="padding-right:3px;">Start kl: </td>
		<td><select name="klok_timer">
		<option value="<%=klok_timer%>" SELECTED><%=klok_timer%></option>
		<option value="08">08</option>
	   	<option value="09">09</option>
	   	<option value="10">10</option>
	   	<option value="11">11</option>
	   	<option value="12">12</option>
	   	<option value="13">13</option>
	   	<option value="14">14</option>
	   	<option value="15">15</option>
	   	<option value="16">16</option>
	   	<!--<option value="17">17</option>
	   	<option value="18">18</option>
	   	<option value="19">19</option>
	   	<option value="20">20</option>
	   	<option value="21">21</option>-->
	  	</select>&nbsp;<b>:</b>
		
		<select name="klok_min">
		<%if func <> "opret" then%>
		<option value="<%=klok_min%>" selected><%=klok_min%></option>
		<%end if%>
		<option value="00">00</option>
		<option value="15">15</option>
	   	<option value="30">30</option>
	   	<option value="45">45</option>
		</select>
		</td>
	</tr>
	
	
	<tr><td align=right style="padding-right:3px;">Slut kl:</td>
		<td>
		<select name="klok_timer_slut">
		<option value="<%=klok_timer_slut%>" SELECTED><%=klok_timer_slut%></option>
		<option value="08">08</option>
	   	<option value="09">09</option>
	   	<option value="10">10</option>
	   	<option value="11">11</option>
	   	<option value="12">12</option>
	   	<option value="13">13</option>
	   	<option value="14">14</option>
	   	<option value="15">15</option>
	   	<option value="16">16</option>
	   	<!--<option value="17">17</option>
	   	<option value="18">18</option>
	   	<option value="19">19</option>
	   	<option value="20">20</option>
	   	<option value="21">21</option>-->
	  	</select>&nbsp;<b>:</b>
		
		<select name="klok_min_slut">
		<%if func <> "opret" then%>
		<option value="<%=klok_min_slut%>" selected><%=klok_min_slut%></option>
		<%else%>
		<option value="30" SELECTED>30</option>
		<%end if%>
		<option value="00">00</option>
		<option value="15">15</option>
	   	<option value="30">30</option>
	   	<option value="45">45</option>
		</select>
		</td>
	</tr>
	<tr><td align=right style="padding-right:3px;">Medarbejder:</td><td><select name="FM_opr_mid" id="FM_opr_mid">
	<%		
			strSQLM = "SELECT mnavn, mid, mnr FROM medarbejdere WHERE mansat <> 2 ORDER BY mnavn"
			oRec.open strSQLM, oConn, 3 
			while not oRec.EOF 
			if cint(medid) = oRec("mid") then
			ts = "SELECTED"
			else
			ts = ""
			end if
			%>
			<option value="<%=oRec("mid")%>" <%=ts%>><%=oRec("mnavn")%></option>
			<%
			oRec.movenext
			wend
			oRec.close 
	%>
	</select></td></tr>
	<tr><td align=right style="padding-right:3px;">Job:</td><td><select name="FM_opr_jobid" id="FM_opr_jobid" style="width:250;">
	
	<%
	strSQLJ = "SELECT jobnavn, id, jobnr FROM job WHERE jobstatus = 1 AND fakturerbart = 1 AND (jobstartdato <= '"& jobstKri &"' AND jobslutdato >= '"& jobslKri &"') ORDER BY jobnavn"
	oRec.open strSQLJ, oConn, 3 
	while not oRec.EOF 
	if cint(jobid) = cint(oRec("id")) then
	ts2 = "SELECTED"
	else
	ts2 = ""
	end if
	%>
	<option value="<%=oRec("id")%>" <%=ts2%>><%=left(oRec("jobnavn"), 30)%> (<%=oRec("jobnr")%>)</option>
	<%
	oRec.movenext
	wend
	oRec.close 
	%>
	</select></td></tr>
	
	<tr><td colspan=2 align=center><br><input type="image" src="../ill/step2pil.gif"></td></tr>
	</form>
	</table>

	</div>	
	<%
	end select
	end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
