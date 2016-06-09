<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<%

sub crmaktstatheader
%>
<table border=0 cellpadding=0 cellspacing=0 width="450">
	<tr>
	<td valign="top" width="163"><img src="../ill/logo_bg.gif" width="163" height="53" alt="" border="0"></td>
	<td valign="bottom"><b>Timeout Kontrolpanel - Incident status</b><br>
	Tilføj, fjern eller rediger Incident status.</td>
	</tr>
	</table><br>
<%
end sub


if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
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
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20; top:10; visibility:visible;">
	<%call crmaktstatheader%>
	<div id="slet" style="position:relative; left:0; top:10; visibility:visible; border:1px red dashed; background-color:#ffff99; width:300px; padding:10px;">
	<table cellspacing="2" cellpadding="2" border="0">
	<tr>
	    <td>Du er ved at <b>slette</b> en status. Er dette korrekt?<br><br>
		Du skal være opmærksom på at der kan ligge response tider på incident der er gemt med denne status.</td>
	</tr>
	<tr>
	   <td><a href="sdsk_status.asp?menu=tok&func=sletok&id=<%=id%>">Ja</a>&nbsp;&nbsp;&nbsp;<a href="javascript:history.back()">Nej</a></td>
	</tr>
	</table>
	</div>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%
	case "sletok"
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM sdsk_status WHERE id = "& id &"")
	Response.redirect "sdsk_status.asp?menu=tok&shokselector=1"
	
	case "dbopr", "dbred"
	'*** Her indsættes en ny type i db ****
		if len(request("FM_navn")) = 0 then
		%>
		<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
		<%
		errortype = 8
		call showError(errortype)
		
		else
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function
		
		strNavn = SQLBless(request("FM_navn"))
		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
		strFarve = request("FM_farve")
		
		if len(request("FM_vispahovedliste")) <> 0 then
		vispahovedliste = request("FM_vispahovedliste")
		else
		vispahovedliste = 0
		end if
		
		if func = "dbopr" then
		oConn.execute("INSERT INTO sdsk_status (navn, editor, dato, farve, vispahovedliste) VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', '"& strFarve &"', "& vispahovedliste &")")
		else
		oConn.execute("UPDATE sdsk_status SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"', farve = '"& strFarve &"', vispahovedliste = "& vispahovedliste &" WHERE id = "&id&"")
		end if
		
		Response.redirect "sdsk_status.asp?menu=tok&shokselector=1"
		end if
	
	case "rsptid"
	
	rsptid_a = request("FM_rsptid_a")
	rsptid_b = request("FM_rsptid_b")
	
	
	strSQL = "UPDATE sdsk_status SET rsptid_a = 0 WHERE id <> 0" 
	oConn.execute(strSQL)
	strSQL = "UPDATE sdsk_status SET rsptid_a = 1 WHERE id = " & rsptid_a
	oConn.execute(strSQL)
	strSQL = "UPDATE sdsk_status SET rsptid_b = 0 WHERE id <> 0" 
	oConn.execute(strSQL)
	strSQL = "UPDATE sdsk_status SET rsptid_b = 1 WHERE id = " & rsptid_b
	oConn.execute(strSQL)
	
	Response.redirect "sdsk_status.asp"
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	
	else
	strSQL = "SELECT navn, editor, dato, farve, vispahovedliste FROM sdsk_status WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	strFarve = oRec("farve")
	vispahovedliste = oRec("vispahovedliste")
	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "opdaterpil" 
	end if
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20; top:10; visibility:visible;">
	<%call crmaktstatheader%>
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr><form action="sdsk_status.asp?menu=tok&func=<%=dbfunc%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
    	<td valign="top" colspan="2"><font class="pageheader">Kontrolpanel | Incident status | <%=varbroedkrumme%></font></td>
	</tr>
	<%if dbfunc = "dbred" then%>
	<tr>
		<td colspan="2" valign="bottom" style="height:30;">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></td>
	</tr>
	<%end if%>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td width=50><b>Navn:</b></td>
		<td><input type="text" name="FM_navn" value="<%=strNavn%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td colspan=2><br>
		<%if cint(vispahovedliste) = 1 then
		vispahovedlisteCHK = "CHECKED"
		else
		vispahovedlisteCHK = ""
		end if%>
		<input type="checkbox" name="FM_vispahovedliste" id="FM_vispahovedliste" value="1" <%=vispahovedlisteCHK%>> Inkluder denne status som forvalgt kriterie<br>
		 på Incident hovedlisten. </td>
	</tr>
	<tr>
		<td><br><b>Farve:</b></td>
		<td>
		<%select case strFarve
		case "orange"
		sel_1 = ""
		sel_2 = "SELECTED"
		sel_3 = ""
		sel_4 = ""
		sel_5 = ""
		sel_6 = ""
		case "darkRed"
		sel_1 = ""
		sel_2 = ""
		sel_3 = "SELECTED"
		sel_4 = ""
		sel_5 = ""
		sel_6 = ""
		case "#8caae6"
		sel_1 = ""
		sel_2 = ""
		sel_3 = ""
		sel_4 = "SELECTED"
		sel_5 = ""
		sel_6 = ""
		case "#003399"
		sel_1 = ""
		sel_2 = ""
		sel_3 = ""
		sel_4 = ""
		sel_5 = ""
		sel_6 = "SELECTED"
		case "#cccccc"
		sel_1 = ""
		sel_2 = ""
		sel_3 = ""
		sel_4 = ""
		sel_5 = "SELECTED"
		sel_6 = ""
		case else
		sel_1 = "SELECTED"
		sel_2 = ""
		sel_3 = ""
		sel_4 = ""
		sel_5 = ""
		sel_6 = ""
		end select%>
		
		<br><select name="FM_farve" id="FM_farve">
	<option value="#ffff99" <%=sel_1%>>Gul</option>
	<option value="orange" <%=sel_2%>>Orange</option>
	<option value="darkRed" <%=sel_3%>>Rød</option>
	<option value="#8caae6" <%=sel_4%>>Blå</option>
	<option value="#003399" <%=sel_6%>>Mørk Blå</option>
	<option value="#cccccc" <%=sel_5%>>Grå</option>
	</select></td>
	</tr>
	<tr>
		<td colspan="2"><br><br><img src="ill/blank.gif" width="100" height="1" alt="" border="0"><input type="image" src="../ill/<%=varSubVal%>.gif"></td>
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
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<script language="javascript">
	<!--
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	//-->
	</script>
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20; top:10; visibility:visible;">
	<%call crmaktstatheader%>
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr>
    <td valign="top"><font class="pageheader">Kontrolpanel | Incident status</font>
	<br>
	<br>
	<!--Sortér efter <a href="sdsk_status.asp?menu=tok&sort=navn">Navn</a> eller <a href="sdsk_status.asp?menu=tok&sort=nr">Id nr.</a>-->
	<img src="../ill/blank.gif" width="405" height="1" alt="" border="0">
	<a href="sdsk_status.asp?menu=tok&func=opret">Opret ny status mulighed <img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a><br>
	<br>
	</td></tr>
	</table>
	
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<form action="sdsk_status.asp?func=rsptid" method="post">
	<tr bgcolor="#5582D2">
		<td width="8" valign=top rowspan=2><img src="../ill/tabel_top_left.gif" width="8" height="32" alt="" border="0"></td>
		<td colspan=5 valign="top" style="border-top:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right valign=top rowspan=2><img src="../ill/tabel_top_right.gif" width="8" height="32" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td width="40" class=alt><b>Id</b></td>
		<td class=alt><b>Status navn</b></td>
		<td class=alt><b>Slet?</b></td>
		<td class=alt><b>Responstid A</b></td>
		<td class=alt><b>Responstid B</b></td>
		
	</tr>
	<%
	'sort = Request("sort")
	'if sort = "navn" then
	'strSQL = "SELECT id, navn, rsptid_a, rsptid_b FROM sdsk_status ORDER BY navn"
	'else
	strSQL = "SELECT id, navn, rsptid_a, rsptid_b FROM sdsk_status ORDER BY rsptid_a DESC, rsptid_b DESC"
	'end if
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	%>
	<tr>
		<td bgcolor="#003399" colspan="7"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="#ffffff" onmouseover="mOvr('gift',this,'#B4C7EF');" onmouseout="mOut(this,'#ffffff');">
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td><%=oRec("id")%></td>
		<td height="20"><a href="sdsk_status.asp?menu=tok&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%> </a></td>
		<td><a href="sdsk_status.asp?menu=tok&func=slet&id=<%=oRec("id")%>"><img src="../ill/slet.gif" width="20" height="20" alt="" border="0"></a></td>
		<td align=center>
		
		<%if cint(oRec("rsptid_a"))  = 1 then
		rsptid_a_CHK = "CHECKED"
		else
		rsptid_a_CHK = ""
		end if%>
		
		<input type="radio" name="FM_rsptid_a" value="<%=oRec("id")%>" <%=rsptid_a_CHK%>></td>
		<td align=center>
		
		<%if cint(oRec("rsptid_b"))  = 1 then
		rsptid_b_CHK = "CHECKED"
		else
		rsptid_b_CHK = ""
		end if%>
		
		<input type="radio" name="FM_rsptid_b" value="<%=oRec("id")%>" <%=rsptid_b_CHK%>></td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<%
	x = 0
	oRec.movenext
	wend
	%>
	<tr bgcolor="#ffffff">
		<td width="8" valign=top height=50 style="border-bottom:1px #003399 solid; border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td colspan=5 valign="middle" style="border-bottom:1px #003399 solid;" align=right><input type="submit" value="Opdater Rsp. tider"></td>
		<td align=right valign=top style="border-bottom:1px #003399 solid; border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td width="8" valign=top height=20 style="border-bottom:1px #003399 solid; border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td colspan=5 valign="top" style="border-bottom:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right valign=top style="border-bottom:1px #003399 solid; border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	</form>	
	</table>
	<br><br><br>
	<a href="Javascript:window.close()" class=vmenu>Luk vindue, og vend tilbage til kontrolpanel</a><br><br>
	
			<div id="jobinfo" style="position:relative; left:0; top:20; width:400; visibility:visible; padding:10px; border:1px red dashed; background-color:#ffffe1;">
			<table>
			<tr><td>
			<img src="../ill/ac0005-24.gif" width="24" height="24" alt="" border="0">&nbsp;<b>Side note(r):<br></b>
			Respons tider bruges til at holde øje med hvor lang der går fra en incident er oprettet til den skifter til den valgte status.
			<br><br>Når en responstid er gemt på en incident kan denne ikke overskrives, 
			men hvis oprettelses tiden på det pågældende incident fornyes bliver respons tiderne A og B nulstillet.
			<br>&nbsp;</td></tr>
			</table>
			</div>
	</div>
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
