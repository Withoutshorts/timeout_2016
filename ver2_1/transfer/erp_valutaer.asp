<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/sdsk_func.asp"-->
<%





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
	
	thisfile = "sdsk_kontrolpanel"
	kview = "j"
	
	if len(request("lastedit")) <> 0 then
	lastedit = request("lastedit")
	else
	lastedit = 0
	end if
	
	select case func
	case "slet"
	%>
		<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
		<!--#include file="../inc/regular/topmenu_inc.asp"-->
		<%
	
	slttxtalt = ""
	slturlalt = ""
	
	slttxt = "Du er ved at <b>slette</b> en valuta.<br>"_
    &"Du vil samtidig nulstille alle fakturaer der er oprettet med denne valuta til at blive vist med grundvalutaen.<br>"
	
    slturl = "erp_valutaer.asp?menu=tok&func=sletok&id="&id
	
	call sltque(slturl,slttxt,slturlalt,slttxtalt,210,100)
	
	case "sletok"
	'*** Her slettes en valuta ***
	oConn.execute("DELETE FROM valutaer WHERE id = "& id &"")
	Response.redirect "erp_valutaer.asp?menu=tok"
	
	case "dbopr", "dbred"
	'*** Her indsættes en ny type i db ****
		
		
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function
		
		function SQLBless2(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless2 = tmp
		end function
		
		err = 0
		
		
		
		if len(trim(request("FM_navn"))) <> 0 then
		strNavn = SQLBless(request("FM_navn"))
		else
		err = 110
		end if
		
		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
		
		
		if len(trim(request("FM_kode"))) <> 0 then
		strKode = SQLBless(request("FM_kode"))
		else
		err = 111
		end if
		
		    %>
			<!--#include file="inc/isint_func.asp"-->
			<%
			call erDetInt(request("FM_kurs"))
			
					if isInt > 0 then
					err = 112
					else
		            intKurs = SQLBless2(request("FM_kurs"))
		            end if
		    
		            isInt = 0
		
		if err = 0 then
		
		
		        if func = "dbopr" then
		        strSQLopr = "INSERT INTO valutaer (valuta, editor, dato, valutakode, kurs) "_
		        &" VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', "_
		        &" '"& strKode &"', "& intKurs &")"
        		
		        'Response.Write strSQLopr
		        'Response.flush
		        oConn.execute(strSQLopr)
        				
        				
        			
        				
		        else
		        oConn.execute("UPDATE valutaer SET valuta ='"& strNavn &"', editor = '" &strEditor &"', "_
		        &" dato = '" & strDato &"', valutakode = '"& strKode &"', kurs = "& intKurs &""_
		        &" WHERE id = "&id&"")
        		
		        lastedit = id
        		
		        end if
		
		
		
		Response.redirect "erp_valutaer.asp?menu=tok&shokselector=1&lastedit="&lastedit
		
		
		else
		
		%>
		<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
		<%
		errortype = err
		call showError(errortype)
		
		
		end if
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	
	strNavn = ""
	strTimepris = ""
	varSubVal = "opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	rsptid = 0
	intkunweekend = 1
	
	else
	strSQL = "SELECT valuta, editor, dato, valutakode, kurs FROM valutaer WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	
	strNavn = oRec("valuta")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	strKode = oRec("valutakode")
	intKurs = oRec("kurs")
	
	
	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "opdaterpil" 
	end if
	
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20; top:62; visibility:visible;">
	
	<br><br><br>
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr><form action="erp_valutaer.asp?menu=tok&func=<%=dbfunc%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
    	<td valign="top" colspan="2"><h3><img src="../ill/ac0001-24.gif" />  Valutaer - <%=varbroedkrumme%></h3></td>
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
		<td style="padding-top:10px;"><b>Valuta:</b></td>
		<td style="padding-top:10px;"><input type="text" name="FM_navn" value="<%=strNavn%>" size="40" style="border: 1px #86B5E4 solid;"></td>
	</tr>
	<tr>
		<td style="padding-top:10px;"><b>Valuta ISO kode:</b> </td>
		<td style="padding-top:10px;"><input type="text" name="FM_kode" value="<%=strKode%>" size="5" style="border: 1px #86B5E4 solid;"> (DKR, SEK, EUR, USD etc.)</td>
	</tr>
	<tr>
		<td style="padding-top:10px;"><b>Kurs:</b><br />
		 Afrundes til 2 decimaler</td>
		<td style="padding-top:10px;"><input type="text" name="FM_kurs" value="<%=intKurs %>" size=10 style="border: 1px #86B5E4 solid;"> (89,34)</td>
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
	<%case else
	
	
	%>
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
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20px; top:62px; visibility:visible;">
	
	<h3>
        <img src="../ill/ac0001-24.gif" /> Valutaer</h3>
	<a href="erp_valutaer.asp?menu=tok&func=opret&prio_grp=<%=prio_grp%>">Opret ny valuta <img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a><br>
	
	
	
	
	
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr bgcolor="#5582D2">
		<td width="8" valign=top rowspan=2><img src="../ill/tabel_top_left.gif" width="8" height="32" alt="" border="0"></td>
		<td colspan=5 valign="top" style="border-top:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td width="8" align=right valign=top rowspan=2><img src="../ill/tabel_top_right.gif" width="8" height="32" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td width="50" class=alt><b>Id</b></td>
		<td width="300" class=alt><b>Valuta</b> (grundvaluta)</td>
		<td class=alt><b>Valuta ISO kode</b></td>
		<td class=alt><b>Kurs</b></td>
		<td width="50" class=alt>
            &nbsp;</td>
	</tr>
	<%
	strSQL = "SELECT id, valuta, valutakode, kurs, grundvaluta FROM valutaer ORDER BY valuta"
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	
	if cint(oRec("id")) = cint(lastedit) then
	bgCol = "#ffff99"
	else
	bgCol = "#ffffff"
	end if%>
	<tr>
		<td bgcolor="#003399" colspan="7"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="<%=bgCol%>" onmouseover="mOvr('gift',this,'#B4C7EF');" onmouseout="mOut(this,'<%=bgCol%>');">
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td><%=oRec("id")%></td>
		<td height="20"><a href="erp_valutaer.asp?menu=tok&func=red&id=<%=oRec("id")%>"><%=oRec("valuta")%> </a>
		&nbsp;&nbsp;(<%=oRec("grundvaluta")%>)</td>
		<td><%=oRec("valutakode")%></td>
		<td><%=oRec("kurs")%></td>
		<td style="padding-top:2px;">
		<a href="erp_valutaer.asp?menu=tok&func=slet&id=<%=oRec("id")%>"><img src="../ill/slet_16.gif" alt="" border="0"></a>
		&nbsp;</td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<%
	x = 0
	oRec.movenext
	wend
	%>
	<tr bgcolor="#5582D2">
		<td width="8" valign=top height=20 style="border-bottom:1px #003399 solid; border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td colspan=5 valign="top" style="border-bottom:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right valign=top style="border-bottom:1px #003399 solid; border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>	
	</table>
	<br><br><br>
	<a href="Javascript:window.close()">[Luk vindue]</a><br><br>
        &nbsp;
	</div>
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
