<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/job_func.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function
		
	
	select case func
	case "slet"
	'*** Her spørges om det er ok at der slettes en medarbejder ***
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:210; top:100; background-color:#ffff99; visibility:visible; border:2px #000000 dashed; padding:15px;"">
	<table cellspacing="2" cellpadding="2" border="0" bgcolor="#ffff99" width=400>
	<tr>
	    <td><img src="../ill/alert_lille.gif" width="22" height="19" alt="" border="0">&nbsp;<b>Slet gruppe?</b><br>
			Du er ved at <b>slette</b> en Stam-aktivitets gruppe. Er dette korrekt?</td>
	</tr>
	<tr>
	   <td><br><a href="akt_gruppe.asp?menu=job&func=sletok&id=<%=id%>" class=vmenu>Ja, slet gruppe.</a><br><br>
	   <a href="javascript:history.back()" class=red>Nej, tilbage.</a></td>
	</tr>
	</table>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%
	case "sletok"
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM akt_gruppe WHERE id = "& id &"")
	Response.redirect "akt_gruppe.asp?menu=job&shokselector=1"
	
	case "dbopr", "dbred"
	'*** Her indsættes en ny type i db ****
		if len(request("FM_navn")) = 0 then
		%>
		<!--#include file="../inc/regular/header_inc.asp"-->
		<!--#include file="../inc/regular/topmenu_inc.asp"-->
		<%
		errortype = 8
		call showError(errortype)
		
		else
		
		strNavn = SQLBless(request("FM_navn"))
		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
		
		if func = "dbopr" then
		oConn.execute("INSERT INTO akt_gruppe (navn, editor, dato) VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"')")
		else
		oConn.execute("UPDATE akt_gruppe SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"' WHERE id = "&id&"")
		end if
		
		Response.redirect "akt_gruppe.asp?menu=job&shokselector=1"
		end if
	
	
	case "kopier"
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%call tsamainmenu(3)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	call jobtopmenu()
	%>
	</div>
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20; top:122; visibility:visible;">
	<h3><img src="../ill/aktstam_48.png" alt="" border="0">&nbsp;Kopier Stam-aktivitetsgruppe</h3>
	
	<%
	strSQL = "SELECT navn FROM akt_gruppe WHERE id =" & id
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
	
	grpNavn = oRec("navn")
	
	oRec.movenext
	wend
	
	oRec.close
	%>
	<table cellspacing="1" cellpadding="2" border="0" width="500">
	<form method=post action="akt_gruppe.asp?func=kopierja&id=<%=id%>">
	<tr><td colspan=2><h4>Stam-aktivitetsgruppe:</h4></td></tr>
	<tr>
		<td align=right><b>Gruppenavn:</b></td><td><input type="text" name="FM_grpnavn" id="FM_grpnavn" size=40 value="Kopi af: <%=grpNavn%>"></td>
	</tr>
	<tr><td colspan=2><br><h4>Aktiviteterne:</h4></td></tr>
	<tr>
		<td align=right valign=top style="padding-top:6px;"><b>Aktivitetsnavn:</b></td>
		<td style="padding-top:2px;"><input type="checkbox" name="FM_udskift_navn" id="FM_udskift_navn" value="1">Ja, omdøb navn på aktiviteter.<br>
		Udskift denne del af aktivitetsnavnet:<br> 
		<input type="text" name="FM_aktnavn_for" id="FM_aktnavn_for" value=""><br>
		med dette: <br>
		<input type="text" name="FM_aktnavn_efter" id="FM_aktnavn_efter" value=""></td>
	</tr>
	<tr>
		<td align=right valign=top style="padding-top:20px;"><b>Faktor:</b></td>
		<td style="padding-top:16px;"><input type="checkbox" name="FM_udskift_faktor" id="FM_udskift_faktor" value="1">Ja, udskift faktor.<br>
		<input type="text" name="FM_faktor" id="FM_faktor" value="" size=2></td>
	</tr>
	<tr>
		<td align=right valign=top style="padding-top:20px;"><b>Forretningsområde:</b></td>
			<td style="padding-top:16px;"><input type="checkbox" name="FM_udskift_fomr" id="FM_udskift_fomr" value="1">Ja, udskift forretningsområde.<br>
			<select name="FM_fomr">
		<option value="0">Ingen</option>
		<%
		strSQL = "SELECT id, navn FROM fomr ORDER BY navn"
		oRec.open strSQL, oConn, 3 
		while not oRec.EOF 
		%>
		<option value="<%=oRec("id")%>"><%=oRec("navn")%></option>
		<%
		oRec.movenext
		wend
		oRec.close 
		%>
		</select>
		</td>
	</tr>
	<tr>
		<td align=right valign=top style="padding-top:20px;"><b>Tidslås:</b></td>
		<td style="padding-top:16px;"><input type="checkbox" name="FM_udskift_tidslaas" id="FM_udskift_tidslaas" value="1">Ja, udskift tidslås.<br>
		<input type="checkbox" name="FM_tidslaas" id="FM_tidslaas" value="1">Tidslås aktiv. Der skal <b>kun</b> kunne registreres timer mellem:<br> 
		kl. <input type="text" name="FM_tidslaas_start" id="FM_tidslaas_start" size="3" value="07:00"> og kl.
		<input type="text" name="FM_tidslaas_slut" id="FM_tidslaas_slut" size="3" value="23:30"> (Format: tt:mm)</td>
	</tr>
	<tr>
		<td colspan=2 align=center><br><br><input type="submit" value="Kopier gruppe"></td>
	</tr>
	</form>
	</table>
	
	</div>
	<%
	case "kopierja"
	
	'Response.write "så er vi her"
	
			
			'** Faktor **
			if len(request("FM_udskift_faktor")) <> 0 then
			udskiftFaktor = 1
			else
			udskiftFaktor = 0
			end if
			aktFaktor = replace(request("FM_faktor"), ",", ".")
				
			%>
			<!--#include file="inc/isint_func.asp"-->
			<%
			call erDetInt(dblFaktor)
			if isInt > 0 then
			%>
			<!--#include file="../inc/regular/header_inc.asp"-->
			<!--#include file="../inc/regular/topmenu_inc.asp"-->
			<%
			errortype = 51
			call showError(errortype)
			isInt = 0
			response.end
			end if 
				
					
				
				'*** tidslaas validering ***
				'** Tidslaas **
				udskiftTidslaas = request("FM_udskift_tidslaas")
				
				if len(request("FM_tidslaas")) <> 0 then
				tidslaas = 1
				else
				tidslaas = 0
				end if
				
				tidslaas_st = request("FM_tidslaas_start") & ":00"
				tidslaas_sl = request("FM_tidslaas_slut") & ":00"
					
				function SQLBless3(s3)
				dim tmp3
				tmp3 = s3
				tmp3 = replace(tmp3, ":", "")
				SQLBless3 = tmp3
				end function
				
				
				idagErrTjek = day(now)&"/"&month(now)&"/"&year(now)
				
				for t = 1 to 2
				
				select case t
				case 1
				tdsl = tidslaas_st
				case 2
				tdsl = tidslaas_sl
				end select
				
				'Response.write SQLBless3(trim(tSttid(y))) & "<br>"
				'Response.flush
				
				call erDetInt(SQLBless3(trim(tdsl)))
				if isInt > 0 then
					antalErr = 1
					errortype = 63
					useleftdiv = "t"
					%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
					call showError(errortype)
					response.end
				end if	
				isInt = 0
				
				'*** Punktum  i angivelse ved registrering af klokkeslet
				if instr(trim(tdsl), ".") <> 0 OR instr(trim(tdsl), ",") <> 0 then
					antalErr = 1
					errortype = 66
					useleftdiv = "t"
					%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
					call showError(errortype)
					response.end
				end if	
				
				if len(trim(tdsl)) <> 0 then
				
				'Response.write idagErrTjek &" "& tSttid(y)&":00" &" - "& isdate(idagErrTjek &" "& tSttid(y)&":00") &"<br>"
					if isdate(idagErrTjek &" "& tdsl) = false then
						antalErr = 1
						errortype = 64
						useleftdiv = "t"
						%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
						call showError(errortype)
						response.end
					end if
				end if
				
			
			next
				
				
			'*** Validering ok! *** 		
			
			
			'** Gruppe **
			intAktfavgp = id
			grpNavn = SQLBless(request("FM_grpnavn"))
			strEditor = session("user")
			strDato = day(now) &"-" & month(now) & "-" & year(now)
			
			'** Opretter gruppe ***
			strSQLopr = "INSERT INTO akt_gruppe (navn, editor, dato) VALUES ('"& grpNavn &"', '"& strEditor &"', '"& strDato &"')"
			oConn.execute strSQLopr
			
			
			'*** Henter gruppe ID ***
			strSQL = "SELECT id FROM akt_gruppe WHERE id <> 0 ORDER BY id DESC"
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
				
				newaktgrpId = oRec("id")
			
			end if
			oRec.close   
			
			'** Navn **
			if len(request("FM_udskift_navn")) <> 0 then
			udskiftNavn = 1
			else
			udskiftNavn = 0
			end if
			aktNavn_for = SQLBless(request("FM_aktnavn_for")) 
			aktNavn_efter = SQLBless(request("FM_aktnavn_efter"))
			
			
			function SQLBless4(s, glnavn, nytnavn)
			tmpthis = s
			tmpthis = replace(lcase(tmpthis), ""&lcase(glnavn)&"", ""&nytnavn&"")
			SQLBless4 = tmpthis
			end function
			
			
			'** Forretningsområde ***
			if len(request("FM_udskift_fomr")) <> 0 then
			udskiftFomr = 1
			else
			udskiftFomr = 0
			end if
			aktFomr = request("FM_fomr") 
			
			
			
				
				
							'**** Henter aktiviteter fra oprindelig gruppe ***
							strSQL2 = "select id, navn, fakturerbar, budgettimer, fomr, faktor, aktbudget, aktstatus, tidslaas, tidslaas_st, tidslaas_sl FROM aktiviteter WHERE aktFavorit = "& intAktfavgp 
								oRec2.open strSQL2, oConn, 3
									while not oRec2.EOF
									
									aktNavn = replace(oRec2("navn"), "'", "''")
									
									if cint(udskiftNavn) = 1 then
									aktNavn = SQLBless4(trim(aktNavn),aktNavn_for,trim(aktNavn_efter))
									end if
									
									aktFakbar = oRec2("fakturerbar")
									
									if cint(udskiftFomr) = 0 then
									aktFomr = oRec2("fomr")
									else
									aktFomr = aktFomr
									end if
									
									if cint(udskiftFaktor) = 0 then
									aktFaktor = replace(oRec2("faktor"), ",", ".")
									else
									aktFaktor = aktFaktor
									end if
									
									if len(aktFaktor) <> 0 then
									aktFaktor = aktFaktor
									else
									aktFaktor = 0
									end if
									
									aktBudget = replace(oRec2("aktbudget"), ",", ".")
									aktstatus = oRec2("aktstatus")
									
									if cint(tidslaas) = 0 then 
									
										tidslaas = oRec2("tidslaas")
										tidslaas_st = oRec2("tidslaas_st")
										tidslaas_sl = oRec2("tidslaas_sl")
										
										
										if len(tidslaas) <> 0 then
										tidslaas = tidslaas
										else
										tidslaas = 0
										end if
										
										
										if len(trim(tidslaas_st)) <> 0 AND len(trim(tidslaas_sl)) <> 0 then
										tidslaas_st = left(formatdatetime(tidslaas_st, 3), 7)
										tidslaas_sl = left(formatdatetime(tidslaas_sl, 3), 7)
										else
										tidslaas_st = "07:00:00"
										tidslaas_sl = "23:30:00"
										end if
									
									end if
									
									
									aktBudgettimer = replace(oRec2("budgettimer"), ",", ".")
									
									
									strSQLins = "INSERT INTO aktiviteter "_
									&" (navn, dato, editor, job, fakturerbar, "_
									&" projektgruppe1, projektgruppe2, projektgruppe3, "_
									&" projektgruppe4, projektgruppe5, projektgruppe6, "_
									&" projektgruppe7, projektgruppe8, projektgruppe9, "_
									&" projektgruppe10, aktstartdato, aktslutdato, "_
									&" budgettimer, fomr, faktor, aktBudget, aktstatus, tidslaas, tidslaas_st, tidslaas_sl, aktfavorit) VALUES "_
									&"('"& aktNavn &"', "_
									&"'"& strDato &"', "_ 
									&"'"& strEditor &"', "_
									&"0, "_ 
									&""& aktFakbar &", "_
									&"10,1,1,1,1,1,1,1,1,1, "_ 
									&"'2001/1/1', "_ 
									&"'2044/1/1', "_
									&""& aktBudgettimer & ", "& aktFomr &", "_
									&""& aktFaktor &", "& aktBudget &", "& aktstatus &", "_
									&""&tidslaas&", '"&tidslaas_st&"', '"&tidslaas_sl&"', "& newaktgrpId &")"
									
									
									'Response.write strSQLins
									'Response.flush
									oConn.execute(strSQLins)
									
									
									'*** Henter det netop oprettede akt-id ***
									strSQLid = "SELECT id FROM aktiviteter ORDER BY id DESC"
									oRec3.open strSQLid, oConn, 3
									if not oRec3.EOF then
									useAktid = oRec3("id")
									end if
									oRec3.close
									
									if len(useAktid) <> 0 then
									useAktid = useAktid
									else
									useAktid = 0
									end if
									
									
									'*** Timepriser ***
									strSQLtp1 = "SELECT jobid, aktid, medarbid, timeprisalt, 6timepris FROM timepriser WHERE aktid = " & oRec2("id")
									oRec3.open strSQLtp1, oConn, 3
									while not oRec3.EOF
									 	
										strjnr = oRec3("jobid")
										medarbid = oRec3("medarbid")
										timeprisalt = replace(oRec3("timeprisalt"), ",",".")
										this6timepris = replace(oRec3("6timepris"), ",",".")
										
										strSQLtp2 = "INSERT INTO timepriser (jobid, aktid, medarbid, timeprisalt, 6timepris) VALUES ("& strjnr &", "& useAktid &", "& medarbid &", "&timeprisalt&", "& this6timepris &")"
										oConn.execute (strSQLtp2)
										
										
									oRec3.movenext
									wend
									oRec3.close
									
									
									
							
							oRec2.movenext
							wend
							oRec2.close
					
						
		
	Response.redirect "akt_gruppe.asp"
	
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	
	else
	strSQL = "SELECT navn, editor, dato FROM akt_gruppe WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "opdaterpil" 
	end if
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%call tsamainmenu(3)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	call jobtopmenu()
	%>
	</div>
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20; top:122; visibility:visible;">
	<h3><img src="../ill/aktstam_48.png" alt="" border="0">&nbsp;Stam-aktivitetsgrupper</h3>
	<table cellspacing="0" cellpadding="0" border="0" width="400">
	<tr><form action="akt_gruppe.asp?menu=crm&func=<%=dbfunc%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
    	<td valign="top" colspan="2">&nbsp;</td>
	</tr>
	<%if dbfunc = "dbred" then%>
	<tr>
		<td colspan="2" valign="bottom" style="height:30;">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></td>
	</tr>
	<%end if%>
	<tr>
		<td><b>Stam-aktivitetsgruppe:</b> (Navn)</td>
		<td><input type="text" name="FM_navn" value="<%=strNavn%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td colspan="2"><br><br><img src="ill/blank.gif" width="100" height="1" alt="" border="0">
		<input type="image" src="../ill/<%=varSubVal%>.gif"></td>
	</tr>
	</form>
	</table>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%case else%>
	<script language="javascript">
	<!--
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	//-->
	</script>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%call tsamainmenu(3)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	call jobtopmenu()
	%>
	</div>
	
	<!--#include file="inc/convertDate.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20; top:122; visibility:visible;">
	<h3><img src="../ill/aktstam_48.png" alt="" border="0">&nbsp;Stam-aktivitetsgrupper</h3>
	<a href="akt_gruppe.asp?menu=job&func=opret">Opret ny aktivitetsgruppe&nbsp;&nbsp;<img src="../ill/pil_groen_lille.gif" width="17" height="17" alt="" border="0"></a><br>
	<br>
	<table cellspacing="0" cellpadding="0" border="0" width="600" bgcolor="#EFF3FF">
 	<tr bgcolor="5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td colspan=5 valign="top"><img src="../ill/tabel_top.gif" width="687" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td class='alt'><b>Id</b></td>
		<td class='alt'><b>Stam-aktivitets gruppe</b></td>
		<td class='alt'><b> Se gruppe </b>(antal akt.)</td>
		<td class='alt'><b>Slet gruppe?</b></td>
		<td class='alt'><b>Kopier gruppe?</b></td>
	</tr>
	<%
	strSQL = "SELECT id, navn FROM akt_gruppe WHERE id <> 2 ORDER BY navn" 'gl. kørsels gruppe
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	%>
	<tr>
		<td bgcolor="#5582D2" colspan="7"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr onmouseover="mOvr('gift',this,'#B4C7EF');" onmouseout="mOut(this,'');">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td><%=oRec("id")%></td>
					
					<%
					'** Henter aktiviteter i den aktuelle gruppe ****
					strSQL2 = "SELECT count(id) AS antal FROM aktiviteter WHERE aktfavorit = "&oRec("id")&""
					oRec2.open strSQL2, oConn, 3
					if not oRec2.EOF then
					intAntal = oRec2("antal")
					end if
					oRec2.close 
					%>
		
		<%if oRec("id") = 2 then%>
		<td height="20"><%=oRec("navn")%></td>
		<%else%>
		<td height="20"><a href="akt_gruppe.asp?menu=job&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%> </a></td>
		<%end if%>
		<td><a href='aktiv.asp?menu=job&func=favorit&id=<%=oRec("id")%>&stamakgrpnavn=<%=oRec("navn")%>' class='vmenuglobal'>Se / Rediger Stam-aktiviteter i grp.&nbsp;</a>(<%=intAntal%>)</td>
		<%if oRec("id") <= 2 then%>
		<td>&nbsp;</td>
		<%else%>
			<%if intAntal = 0 then%>
			<td><a href="akt_gruppe.asp?menu=job&func=slet&id=<%=oRec("id")%>"><img src="../ill/slet_16.gif" alt="Slet Stamaktivitetes-gruppe" border="0"></a></td>
			<%else%>
			<td>&nbsp;</td>
			<%end if%>
			
		<%end if%>
		<td><a href="akt_gruppe.asp?func=kopier&id=<%=oRec("id")%>" class=rmenu><img src="../ill/aktstam_kopier_24.gif" alt="Kopier gruppe" border="0"></a></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<%
	x = 0
	oRec.movenext
	wend
	%>	
	<tr bgcolor="#5582D2">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan="5" valign="bottom"><img src="../ill/tabel_top.gif" width="687" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>
	</table>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	
	<div style="position:relative; background-color:#ffffe1; top:0px; left:0px; visibility:visible; border:1px red dashed; padding:15px; width:400px;"><img src="../ill/ac0005-24.gif" width="24" height="24" alt="" border="0">&nbsp;<b>Side note(r):<br></b>
	<b>Default</b> gruppen er systemgruppe og kan ikke slettes, men den kan godt omdøbes.
	<br>&nbsp;
	</div>
	
	
	</div>
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
