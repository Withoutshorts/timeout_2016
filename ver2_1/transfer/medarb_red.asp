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
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	select case func
	case "slet"
	'*** Her spørges om det er ok at der slettes en medarbejder ***
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:190; top:120;">
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF">
	<tr>
	    <td style="!border: 1px; border-color: #8CAAE6; border-style: solid; padding-right:5; padding-left:5; padding-top:5; padding-bottom:5;"><img src="../ill/alert.gif" width="44" height="45" alt="" border="0"><br>
		Du er ved at <b>slette</b> en medarbejder. <br>
		Du vil samtidig slette alle timeregistreringer for denne medarbejder.<br>
		Timeregistreringerne vil <b>ikke kunne genskabes</b>. <br>
		<br>Er du sikker på at du vil slette denne medarbejder?<br><br>
		<a href="medarb_red.asp?menu=medarb&func=sletok&id=<%=id%>">Ja</a>&nbsp;&nbsp;&nbsp;<a href="javascript:history.back()">Nej</a></td>
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
	oConn.execute("DELETE FROM medarbejdere WHERE Mid = "& id &"")
	oConn.execute("DELETE FROM progrupperelationer WHERE MedarbejderId = "& id  &"") 'projektgruppeId = 10 AND
	oConn.execute("DELETE FROM timer WHERE tmnr = "& id &"")
		
	Response.redirect "medarb.asp?menu=medarb"
	
	case "dbred", "dbopr" 
	'*** Her tjekkes om alle required felter er udfyldt. ***
	if len(Request("FM_login")) = 0 OR len(Request("FM_pw")) = 0 OR len(Request("FM_mnr")) = 0 then
	
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<%
	errortype = 9
	call showError(errortype)
	
	else
	 
			%>
			<!--#include file="inc/isint_func.asp"-->
			<%
			call erDetInt(Request("FM_mnr"))
			if isInt > 0 then
			%>
			<!--#include file="../inc/regular/header_inc.asp"-->
			<!--#include file="../inc/regular/topmenu_inc.asp"-->
			<%
			errortype = 22
			call showError(errortype)
			
			isInt = 0
			else
		'*** Hvis alle nødvendige er udfyldt ***
	
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function
		
			strNavn = Request("FM_navn")
			strAnsat = Request("FM_ansat")
			
			if strAnsat = "1" then
			strAnsat = strAnsat
			else
			strAnsat = "2"
			end if
			
			strLogin = Request("FM_login")
			strPw = Request("FM_pw")
			strMnr = Request("FM_mnr")
			strEditor = session("user")
			strDato = session("dato")
			strMedinfo = SQLBless(Request("FM_medinfo"))
			strBrugergruppe = Request("FM_bgruppe")
			strMedarbejdertype = Request("FM_medtype")
			strEmail = Request("FM_email")
			if len(request("FM_tsacrm")) <> 0 then 
			inttsacrm = request("FM_tsacrm")
			else
			inttsacrm = 0
			end if
			
			
			'*** Her tjekkes for dubletter i db ***
			strSQL = "SELECT Mnr FROM medarbejdere WHERE Mnr ="& strMnr &" AND Mid <> "& id
			oRec.open strSQL, oConn, 3	
			if oRec.EOF then
			strMnrOK = "y"
			end if
			oRec.close
			
			strSQL = "SELECT login FROM medarbejdere WHERE login ='"& strLogin &"' AND Mid <> "& id
			oRec.open strSQL, oConn, 3	
			if oRec.EOF then
			strLoginOK = "y"
			end if
			oRec.close
			
			'Checker for dubletter i DB før record opdateres eller indsættes
			if strMnrOK = "y" AND strLoginOK = "y" then
			
					if func = "dbred" then
					oConn.execute("UPDATE medarbejdere SET"_
					&" Mnavn = '"& strNavn &"',"_
					&" Mnr = "& strMnr &","_
					&" Mansat = '"& strAnsat &"',"_
					&" login = '"& strLogin &"',"_
					&" pw = '"& strPw &"',"_
					&" Brugergruppe = "& strBrugergruppe &","_
					&" Medarbejdertype = "& strMedarbejdertype &","_
					&" editor = '"& strEditor &"',"_
					&" dato = '"& strDato &"',"_
					&" Medarbejderinfo = '"& strMedinfo &"', "_
					&" Email = '"& strEmail &"',"_
					&" tsacrm = "& inttsacrm &""_
					&" WHERE Mid = "& id &"")
					  
					Response.Redirect "medarb.asp?menu=medarb"
					end if
					
					
					if func = "dbopr" then
					
					
					Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
					
					' Sætter Charsettet til ISO-8859-1 
					Mailer.CharSet = 2
					' Afsenderens navn 
					Mailer.FromName = "timeout2.1"
					' Afsenderens e-mail 
					Mailer.FromAddress = "support@outzource.dk"
					Mailer.RemoteHost = "webmail.abusiness.dk"
					' Modtagerens navn og e-mail
					Mailer.AddRecipient strNavn, strEmail
					Mailer.AddBCC "Support", "support@outzource.dk" 
					' Mailens emne
					Mailer.Subject = "Velkommen til TimeOut 2.1"
					
					' Selve teksten
					Mailer.BodyText = "" & "Hej "& strNavn & vbCrLf _ 
					& "Du er blevet oprettet som bruger i TimeOut 2.1" & vbCrLf _ 
					& "Dit brugernavn er: " & strLogin & " og dit password er: " & strPw & vbCrLf & vbCrLf _ 
					& "Gem disse oplysninger, til du skal logge ind i TimeOut 2.1."  & vbCrLf _ 
					& "Du kan altid selv ændre dem når du er logget på systemet." & vbCrLf & vbCrLf _ 
					& "Adressen til TimeOut 2.1 er: https://outzource.dk/"&lto&""& vbCrLf & vbCrLf _ 
					& "Med venlig hilsen"& vbCrLf & vbCrLf & strEditor & vbCrLf 
					
					
						oConn.execute("INSERT INTO medarbejdere"_
						&" (Mnavn, Mnr, Mansat, login, pw, Brugergruppe, Medarbejdertype, editor, dato, Medarbejderinfo, Email, tsacrm) VALUES ("_
						&" '"& strNavn &"',"_
						&" "& strMnr &","_
						&" '"& strAnsat &"',"_
						&" '"& strLogin &"',"_
						&" '"& strPw &"',"_
						&" "& strBrugergruppe &","_
						&" "& strMedarbejdertype &","_
						&" '"& strEditor &"',"_
						&" '"& strDato &"',"_
						&" '"& strMedinfo &"',"_
						&" '"& strEmail &"',"_
						&" "& inttsacrm &")")
						
						strSQL = "SELECT Mid FROM medarbejdere WHERE Mnr=" & strMnr &""
						oRec.open strSQL, oConn, 3
						if not oRec.EOF then
						intMid = oRec("Mid")
						end if
						oRec.close
						
						oConn.execute("INSERT INTO progrupperelationer (projektgruppeId, MedarbejderId) VALUES (10, "& intMid  &")")
					
					
					%>
					<!--#include file="../inc/regular/header_inc.asp"-->
					<!--#include file="../inc/regular/topmenu_inc.asp"-->
					<div id="sindhold" style="position:absolute; left:190; top:50;">
					<br><img src="../ill/header_medarb.gif" alt="" border="0" width="279" height="34"><hr align="left" width="630" size="1" color="#000000" noshade><br><br>
					<table cellspacing="0" cellpadding="0" border="0">
					<tr><td>
					<%
					If Mailer.SendMail Then
    				%>
					<br><img src="../ill/info.gif" width="42" height="38" alt="" border="0"><br><br> 
					Du har netop oprettet <b><%=strNavn%></b> i timeout2.1.<br><br>
					Der er afsendt en email til medarbejderen med login og password oplysninger,<br>
					som <b><%=strNavn%></b> bør modtage indenfor et par minutter.<br>
					<br><br>
					<table cellspacing="0" cellpadding="0" border="0">
					<tr>
						<td width="500">
						<br><br><a href="medarb.asp?menu=medarb">OK! og videre...&nbsp;<img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a>
						</td>
					</tr>
					</table><%
					Else
    				Response.Write "Fejl...<br>" & Mailer.Response
  					End if
					%>
					</td></tr></table>
				<%			
				end if
			else
			'*** Hvis dublet fandtes ***
			%>
			<!--#include file="../inc/regular/header_inc.asp"-->
			<!--#include file="../inc/regular/topmenu_inc.asp"-->
			<%
			if strMnrOK <> "y" then
			errortype = 10
			call showError(errortype)
			else
			
			if strLoginOK <> "y" then
			errortype = 11
			call showError(errortype)
						end if
					end if
				end if
			end if
		end if
	case else
	'**************************** Medarbejder Data ************************************
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<script language="javascript">
	<!--
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	//-->
	
	function showdisable() {
	//if ( == ) {
	//document.getElementBy("FM_ansat2").status = "checked"
	alert("Når en medarbejder bliver deaktiveret kan denne ikke længere logge sig ind i TimeOut. Medarbejderen vil ikke længere indgå som bruger når antallet af aktive brugere bliver opgjort. (TimeOut Service Aftale). Alle medarbejderens data vil forblive som de er, og vil stadigvæk være tilgængelige.")
	//}
	}
	
	</script>
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!--#include file="../inc/regular/rmenu.asp"-->
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:190; top:50;">
	<br><img src="../ill/header_medarb.gif" alt="" border="0" width="279" height="34"><hr align="left" width="630" size="1" color="#000000" noshade><br><br>
	<table cellspacing="0" cellpadding="0" border="0" width="80%">
	<%
	if func = "red" then
		strSQL = "SELECT mid, mnavn, mnr, mansat, login, pw, lastlogin, brugergruppe, medarbejdertype, type, navn, medarbejdere.editor, medarbejdere.dato, Medarbejderinfo, email, tsacrm FROM medarbejdere, brugergrupper, medarbejdertyper WHERE mid = "& id &" AND brugergrupper.id = brugergruppe AND medarbejdertyper.id = medarbejdertype"
		
		oRec.open strSQL, oConn, 3
		
		if not oRec.EOF then
		strNavn = oRec("Mnavn")
		strAnsat = oRec("Mansat")
		strLogin = oRec("login")
		strPw = oRec("pw")
		strMnr = oRec("Mnr")
		strEditor = oRec("editor")
		strDato = oRec("dato")
		strMedinfo = oRec("Medarbejderinfo")
		dbfunc = "dbred"
		varSubVal = "Opdater!"
		strEmail = oRec("email")
		intCRM = oRec("tsacrm")
		end if
		oRec.close
		
		'if strAnsat = "Yes" then
		'strAnsat = "checked"
		'else
		'strAnsat = ""
		'end if
		
		if intCRM = 1 then
		strCRMchecked = "checked"
		else
		strCRMchecked = ""
		end if
	%>
	<tr>
		<td colspan="2" valign="bottom" style="height:30;">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></td>
	</tr>
	<%
	else
		strSQL = "SELECT Mnr FROM medarbejdere ORDER BY Mnr"
		oRec.open strSQL, oConn, 3
		
		oRec.movelast
		
		strNavn = ""
		strAnsat = ""
		strLogin = ""
		strPw = ""
		strMnr = (oRec("Mnr") + 1)
		strEditor = ""
		strDato = ""
		strMedinfo = ""
		dbfunc = "dbopr"
		varSubVal = "Opret!" 
		strEmail = ""
		
		oRec.close
	end if
	%>
	<tr>
		<td>&nbsp;</td>
	</tr>	
	<tr>
	<form action="medarb_red.asp?menu=medarb&func=<%=dbfunc%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
		<td>Navn:</td>
		<td><input type="text" name="FM_navn" value="<%=strNavn%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<%if level < 3 then%>
	<tr>
		<td>Medarbejder nr.</td>
		<td><input type="text" name="FM_Mnr" value="<%=strMnr%>" size="10" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<%
	else%>
	<input type="hidden" name="FM_Mnr" value="<%=strMnr%>">
	<%end if
	if level < 3 then%>
	<tr>
		<td>Medarbejderkonto:</td>
		<%
		if strAnsat <> "2" then
		chk1 = "CHECKED"
		chk2 = ""
		else
		chk1 = ""
		chk2 = "CHECKED"
		end if%>
		<td>Aktiv:<input type="radio" name="FM_ansat" id=FM_ansat1 value="1" <%=chk1%>>
		Deaktiveret:<input type="radio" name="FM_ansat" id=FM_ansat2 onClick="showdisable()"; value="2" <%=chk2%>></td>
	</tr>
	<%
	else
		%>
		<input type="hidden" name="FM_ansat" value="<%=strAnsat%>">
		<%
	end if%>
	<tr>
		<td>Login:</td>
		<td><input type="text" name="FM_login" value="<%=strLogin%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td>Password:</td>
		<td><input type="password" name="FM_pw" value="<%=strPw%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td>Email:</td>
		<td><input type="text" name="FM_email" value="<%=strEmail%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	</tr>
	<%if level <= 2 OR level = 6 then%>
	<tr>
		<td>Medarbejdertype:</td>
		<td><select name="FM_medtype" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		<%
		if func = "red" then
			strSQL = "SELECT medarbejdertype, type, id FROM medarbejdere, medarbejdertyper WHERE mid = "& id &" AND medarbejdertyper.id = medarbejdertype"
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
			%>
			<option value="<%=oRec("id")%>" SELECTED><%=oRec("type")%></option>
			<%
			end if
			oRec.close
		end if
		
		strSQL = "SELECT id, type FROM medarbejdertyper ORDER BY type"
		oRec.open strSQL, oConn, 3
		While not oRec.EOF%>
		<option value="<%=oRec("id")%>"><%=oRec("type")%></option>
		<%
		oRec.movenext
		Wend
		oRec.close
		%>
		</select></td>
	</tr>
	<%
	else
		strSQL = "SELECT medarbejdertype, type, id FROM medarbejdere, medarbejdertyper WHERE Mid = "& id &" AND medarbejdertyper.id = medarbejdertype"
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then%>
		<input type="hidden" name="FM_medtype" value="<%=oRec("id")%>">
		<%
		end if
		oRec.close
	end if%>
	<tr>
		<td valign="top">Medarbejder info:</td>
		<td><textarea cols="40" rows="5" name="FM_medinfo" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"><%=strMedinfo%></textarea></td>
	</tr>
	<%if level <= 2 OR level = 6 then%>
	<tr>
		<td>Brugergruppe:</td>
		<td><select name="FM_bgruppe" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		<%
		'** Finder den allerede valgte brugergruppe **
		if func = "red" then
			strSQL = "SELECT brugergruppe, navn, id FROM medarbejdere, brugergrupper WHERE Mid = "& id &" AND brugergrupper.id = brugergruppe"
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
			%>
			<option value="<%=oRec("id")%>" SELECTED><%=oRec("navn")%></option>
			<%
			end if
			oRec.close
		end if
		
		'** Hvis der oprettes ny medarbejder sættes brugergruppe 2 () til default ***
		if licensType = "CRM" then
		strWHEREKlaus = " "
		else
		strWHEREKlaus = " WHERE rettigheder <= 2 OR rettigheder = 6 "
		end if
		
		strSQL = "SELECT id, navn FROM brugergrupper "& strWHEREKlaus &" Order by id"
		oRec.open strSQL, oConn, 3
		While not oRec.EOF
		if oRec("id") = 7 AND func <> "red" then%>
		<option value="<%=oRec("id")%>" SELECTED><%=oRec("navn")%></option>
		<%else%>
		<option value="<%=oRec("id")%>"><%=oRec("navn")%></option>
		<%end if
		oRec.movenext
		Wend
		oRec.close
		%>
		</select></td>
	</tr>
	<%
	'*** Hvis bruger ikke har admin rettigheder ****
	else
		strSQL = "SELECT brugergruppe, navn, id FROM medarbejdere, brugergrupper WHERE Mid = "& id &" AND brugergrupper.id = brugergruppe"
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then%>
		<input type="hidden" name="FM_bgruppe" value="<%=oRec("id")%>">
		<%
		end if
		oRec.close
	end if%>
	<%if licensType = "CRM" then%>
	<tr>
		<td colspan="2"><br>Skal startsiden være i TSA (default) eller i
		 CRM <input type="checkbox" name="FM_tsacrm" value="1" <%=strCRMchecked%> style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">&nbsp;systemet.</td>
	</tr>
	<%end if%>
	<tr>
		<td colspan="2"><br><br><img src="ill/blank.gif" width="200" height="1" alt="" border="0"><input type="submit" value="<%=varsubVal%>" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid; cursor:hand; width:116;"></td>
	</tr>
	</form>
	</table>
<br><br>
<br>
<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
<br>
</div>
<%end select%>
<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
