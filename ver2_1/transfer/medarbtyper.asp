<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/medarb_func.asp"-->

<%
'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
Session.LCID = 1030

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
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	
	
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%call tsamainmenu(6)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	call medarbtopmenu()
	%>
	</div>
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20; top:132;">
	<h3><img src="../ill/ac0019-24.gif" width="24" height="24" alt="" border="0">&nbsp;Medarbejdertyper - Slet?</h3>
	<div style="border:1px red dashed; padding:10px; background-color:#ffff99;">
	<table cellspacing="2" cellpadding="2" border="0">
	<tr>
	    <td><img src="../ill/alert.gif" width="44" height="45" alt="" border="0">Du er ved at <b>slette</b> en medarbejdertype. Er dette korrekt?</td>
	</tr>
	<tr>
	   <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	   <a href="medarbtyper.asp?menu=medarb&func=sletok&id=<%=id%>">Ja</a>&nbsp;&nbsp;&nbsp;<a href="javascript:history.back()">Nej</a></td>
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
	oConn.execute("DELETE FROM medarbejdertyper WHERE id = "& id &"")
	Response.redirect "medarbtyper.asp?menu=medarb"
	
	case "med"
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
	
	
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%call tsamainmenu(6)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	call medarbtopmenu()
	%>
	</div>
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20; top:132;">
	<h3><img src="../ill/ac0019-24.gif" width="24" height="24" alt="" border="0">&nbsp;Medarbejdertyper.</h3><%
	strSQL = "SELECT id, type FROM medarbejdertyper WHERE id=" & id
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
	gruppeNavn = oRec("type")
	end if
	oRec.close
	%>
	<br><br>
	Medarbejdere af typen <b><%=gruppeNavn%></b>:<br>
	<br>
	<table cellspacing="0" cellpadding="0" border="0" width="600" bgcolor="#EFF3FF">
	<tr bgcolor="5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="31" alt="" border="0"></td>
		<td colspan=3 valign="top"><img src="../ill/tabel_top.gif" width="587" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="31" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td width="40" class=alt><b>Navn</b></td>
		<td colspan=2>&nbsp;</td>
	</tr>
	<%
	strSQL = "SELECT Mnavn, Mid FROM medarbejdere WHERE medarbejdertype = "&id&" ORDER BY Mnavn"
	oRec.open strSQL, oConn, 3
	
	while not oRec.EOF 
	%>
	<tr>
		<td bgcolor="#5582D2" colspan="5"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr onmouseover="mOvr('gift',this,'#B4C7EF');" onmouseout="mOut(this,'');">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td height="20"><a href="medarb_red.asp?menu=medarb&func=red&id=<%=oRec("Mid")%>"><%=oRec("Mnavn")%> </a></td>
		<td colspan=2>&nbsp;</td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<%
	oRec.movenext
	wend
	%>
	<tr bgcolor="#5582D2">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan="3" valign="bottom"><img src="../ill/tabel_top.gif" width="587" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>	
	</table>
	<br><br>
<br>
<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
<br>
<br>
	<%
	
	case "dbopr", "dbred"
	%>
	<!--#include file="inc/isint_func.asp"-->
	<%
	'*** Her indsættes en ny type i db ****
		if len(request("FM_navn")) = 0 then
		%>
		<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
		<!--#include file="../inc/regular/topmenu_inc.asp"-->
		<%
		errortype = 8
		call showError(errortype)
		
		else
		
				function SQLBlessDOT(s)
				dim tmpdot
				tmpdot = s
				tmpdot = replace(tmpdot, ".", ",")
				SQLBlessDOT = tmpdot
				end function
				
				
				call erDetInt(SQLBlessDOT(request("FM_Timepris")))
				if isInt > 0 then
				%>
				<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
				<!--#include file="../inc/regular/topmenu_inc.asp"-->
				<%
				errortype = 39
				call showError(errortype)
				isInt = 0
						
				else
					
				call erDetInt(SQLBlessDOT(request("FM_Kostpris")))
				if isInt > 0 then
				%>
				<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
				<!--#include file="../inc/regular/topmenu_inc.asp"-->
				<%
				errortype = 39
				call showError(errortype)
				isInt = 0
						
				else
					
				
					if len(request("FM_Kostpris")) = 0 OR len(request("FM_timepris")) = 0 then
					%>
					<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
					<!--#include file="../inc/regular/topmenu_inc.asp"-->
					<%
					errortype = 39
					call showError(errortype)
					
					else
						
						isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_son")))
						int1 = isInt 
						
						isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_man")))
						int2 = isInt
						
						isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_tir")))
						int3 = isInt
						
						isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_ons")))
						int4 = isInt
						
						isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_tor")))
						int5 = isInt
						
						isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_fre")))
						int6 = isInt
						
						isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_lor")))
						int7 = isInt
						
						isInt = 0
						call erDetInt(request("FM_Kostpris")) 
						int8 = isInt
						
						isInt = 0
						call erDetInt(request("FM_timepris"))
						int9 = isInt
						
						'isInt = 0
						'call erDetInt(request("FM_timepris1"))
						'int10 = isInt
						
						isInt = 0
						call erDetInt(request("FM_timepris2"))
						int11 = isInt
						
						isInt = 0
						call erDetInt(request("FM_timepris3"))
						int12 = isInt
						
						isInt = 0
						call erDetInt(request("FM_timepris4"))
						int13 = isInt
						
						isInt = 0
						call erDetInt(request("FM_timepris5"))
						int14 = isInt
						
						if int1 > 0 OR int2 > 0 OR int3 > 0 OR int4 > 0 OR int5 > 0 OR int6 > 0 OR int7 > 0 OR int8 > 0 OR int9 > 0 OR int11 > 0 OR int12 > 0 OR int13 > 0 OR int14 > 0  then
						%>
						<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
						<!--#include file="../inc/regular/topmenu_inc.asp"-->
						<%
						errortype = 39
						call showError(errortype)
						isInt = 0
					
						else
							
							
							
							function SQLBless(s)
							dim tmp
							tmp = s
							tmp = replace(tmp, "'", "''")
							SQLBless = tmp
							end function
							
							function SQLBlessDOT2(s)
							dim tmpdot
							tmpdot = s
							tmpdot = replace(tmpdot, ",", ".")
							SQLBlessDOT2 = tmpdot
							end function
							
							strNavn = SQLBless(request("FM_navn"))
							strTimepris = SQLBlessDOT2(request("FM_Timepris"))
							strTimepris1 = strTimepris 'SQLBlessDOT2(request("FM_timepris1"))
							strTimepris2 = SQLBlessDOT2(request("FM_timepris2"))
							strTimepris3 = SQLBlessDOT2(request("FM_timepris3"))
							strTimepris4 = SQLBlessDOT2(request("FM_timepris4"))
							strTimepris5 = SQLBlessDOT2(request("FM_timepris5"))
							
							
							if len(strTimepris) <> 0 then
							strTimepris = strTimepris
							else
							strTimepris = 0
							end if
							
							if len(strTimepris1) <> 0 then
							strTimepris1 = strTimepris1
							else
							strTimepris1 = 0
							end if
							
							if len(strTimepris2) <> 0 then
							strTimepris2 = strTimepris2
							else
							strTimepris2 = 0
							end if
							
							if len(strTimepris3) <> 0 then
							strTimepris3 = strTimepris3
							else
							strTimepris3 = 0
							end if
							
							if len(strTimepris4) <> 0 then
							strTimepris4 = strTimepris4
							else
							strTimepris4 = 0
							end if
							
							if len(strTimepris5) <> 0 then
							strTimepris5 = strTimepris5
							else
							strTimepris5 = 0
							end if
							
							
							dubKostpris = request("FM_kostpris")
							normtimer_son = SQLBlessDOT2(request("FM_norm_son"))
							normtimer_man = SQLBlessDOT2(request("FM_norm_man"))
							normtimer_tir =	SQLBlessDOT2(request("FM_norm_tir"))
							normtimer_ons = SQLBlessDOT2(request("FM_norm_ons"))
							normtimer_tor = SQLBlessDOT2(request("FM_norm_tor"))
							normtimer_fre = SQLBlessDOT2(request("FM_norm_fre"))
							normtimer_lor = SQLBlessDOT2(request("FM_norm_lor"))
							
							
							if len(dubKostpris) <> 0 then
							dubKostpris = dubKostpris
							else
							dubKostpris = 0
							end if
							
							
							
							if len(normtimer_man) <> 0 then
							normtimer_man = normtimer_man
							else
							normtimer_man = 0
							end if
							
							if len(normtimer_tir) <> 0 then
							normtimer_tir = normtimer_tir
							else
							normtimer_tir = 0
							end if
							
							if len(normtimer_ons) <> 0 then
							normtimer_ons = normtimer_ons
							else
							normtimer_ons = 0
							end if
							
							if len(normtimer_tor) <> 0 then
							normtimer_tor = normtimer_tor
							else
							normtimer_tor = 0
							end if
							
							if len(normtimer_fre) <> 0 then
							normtimer_fre = normtimer_fre
							else
							normtimer_fre = 0
							end if
							
							if len(normtimer_lor) <> 0 then
							normtimer_lor = normtimer_lor
							else
							normtimer_lor = 0
							end if
							
							if len(normtimer_son) <> 0 then
							normtimer_son = normtimer_son
							else
							normtimer_son = 0
							end if
							
							tp0_valuta = request("FM_valuta_0")
							tp1_valuta = tp0_valuta
							tp2_valuta = request("FM_valuta_2")
							tp3_valuta = request("FM_valuta_3")
							tp4_valuta = request("FM_valuta_4")
							tp5_valuta = request("FM_valuta_5")
							
							
							
							if request("FM_Kostpris") < 0 OR strTimepris < 0 OR strTimepris1 < 0 OR strTimepris2 < 0 OR strTimepris3 < 0 OR strTimepris4 < 0 OR strTimepris5 < 0 then
							%>
							<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
							<!--#include file="../inc/regular/topmenu_inc.asp"-->
							<%
							errortype = 39
							call showError(errortype)
							
							else
		
				
				strEditor = session("user")
				strDato = session("dato")
				
				if func = "dbopr" then
				oConn.execute("INSERT INTO medarbejdertyper (type, timepris, editor, dato, kostpris, normtimer_son, normtimer_man, normtimer_tir, normtimer_ons, normtimer_tor, normtimer_fre, normtimer_lor, "_
				&" timepris_a1, timepris_a2, timepris_a3, timepris_a4, timepris_a5, "_
				&" tp0_valuta, tp1_valuta, tp2_valuta, tp3_valuta, tp4_valuta, tp5_valuta) VALUES"_
				&" ('"& strNavn &"', "& strTimepris &", '"& strEditor &"', '"& strDato &"', "& dubKostpris &", "_
				&" "& normtimer_son &", "& normtimer_man &", "& normtimer_tir &", "& normtimer_ons &", "_
				&" "& normtimer_tor &", "& normtimer_fre &", "& normtimer_lor &", "& strTimepris1 &", "_
				&" "& strTimepris2 &", "& strTimepris3 &", "& strTimepris4 &", "& strTimepris5 &""_
				&" "& tp0_valuta &","& tp1_valuta &","& tp2_valuta &","& tp3_valuta &","& tp4_valuta &","& tp5_valuta &""_
				&" )")
				else
				oConn.execute("UPDATE medarbejdertyper SET type ='"& strNavn &"', timepris = "& strTimepris &", editor = '" &strEditor &"', dato = '" & strDato &"', kostpris = "& dubKostpris &", "_
				&" normtimer_son = "& normtimer_son &", normtimer_man="& normtimer_man &", normtimer_tir="& normtimer_tir &", "_
				&" normtimer_ons="& normtimer_ons &", normtimer_tor="& normtimer_tor &", normtimer_fre="& normtimer_fre &", normtimer_lor="& normtimer_lor &", "_
				&" timepris_a1 = "& strTimepris1 &", timepris_a2 = "& strTimepris2 &", timepris_a3 = "& strTimepris3 &", "_
				&" timepris_a4 = "& strTimepris4 &", timepris_a5 = "& strTimepris5 &", "_
				&" tp0_valuta = "& tp0_valuta &", tp1_valuta = "& tp1_valuta &", "_
				&" tp2_valuta = "& tp2_valuta &", tp3_valuta = "& tp3_valuta &", "_
				&" tp4_valuta = "& tp4_valuta &", tp5_valuta = "& tp5_valuta &""_
				&" WHERE id = "&id&"")
				end if
		
				Response.redirect "medarbtyper.asp?menu=medarb"
				end if 'validering
				end if 'validering
				end if 'validering
				end if 'validering
			end if 'validering
		end if 'validering
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	
	strTimepris = 0
	dubKostpris = 0
	strTimepris1 = 0
	strTimepris2 = 0
	strTimepris3 = 0
	strTimepris4 = 0
	strTimepris5 = 0
	
	normtimer_son = 0
	normtimer_man = 0
	normtimer_tir = 0
	normtimer_ons = 0
	normtimer_tor = 0
	normtimer_fre = 0
	normtimer_lor = 0
	
	else
	strSQL = "SELECT type, editor, dato, timepris, kostpris, normtimer_son, normtimer_man, "_
	&" normtimer_tir, normtimer_ons, normtimer_tor, normtimer_fre, normtimer_lor, timepris_a1, "_
	&" timepris_a2, timepris_a3, timepris_a4, timepris_a5, "_
	&" tp0_valuta, tp1_valuta, tp2_valuta, tp3_valuta, tp4_valuta, tp5_valuta "_
	&" FROM medarbejdertyper WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	
	strNavn = oRec("type")
	strTimepris = oRec("timepris")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	dubKostpris = oRec("kostpris")
	normtimer_son = oRec("normtimer_son")
	normtimer_man = oRec("normtimer_man")
	normtimer_tir = oRec("normtimer_tir")
	normtimer_ons = oRec("normtimer_ons")
	normtimer_tor = oRec("normtimer_tor")
	normtimer_fre = oRec("normtimer_fre")
	normtimer_lor  = oRec("normtimer_lor")
	strTimepris1 = oRec("timepris_a1")
	strTimepris2 = oRec("timepris_a2")
	strTimepris3 = oRec("timepris_a3")
	strTimepris4 = oRec("timepris_a4")
	strTimepris5 = oRec("timepris_a5")
	
	tp0_valuta = oRec("tp0_valuta")
	tp1_valuta = oRec("tp1_valuta")
	tp2_valuta = oRec("tp2_valuta")
	tp3_valuta = oRec("tp3_valuta")
	tp4_valuta = oRec("tp4_valuta")
	tp5_valuta = oRec("tp5_valuta")
	
	end if
	oRec.close
	
	
	if len(trim(dubKostpris)) <> 0 then
	dubKostpris = dubKostpris
	else
	dubKostpris = 0
	end if
	
	
	ugetotal = formatnumber(normtimer_son + normtimer_man + normtimer_tir + normtimer_ons + normtimer_tor + normtimer_fre + normtimer_lor, 2)
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "opdaterpil" 
	end if
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	
	
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%call tsamainmenu(6)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	call medarbtopmenu()
	%>
	</div>
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20; top:132;">
	<h3><img src="../ill/ac0019-24.gif" width="24" height="24" alt="" border="0">&nbsp;Medarbejdertyper.</h3>
	
	
	<table cellspacing="0" cellpadding="0" border="0" width="500" bgcolor="#EFF3FF">
	<form action="medarbtyper.asp?menu=medarb&func=<%=dbfunc%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
	<tr bgcolor="5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="31" alt="" border="0"></td>
		<td colspan=2 valign="top" style="border-top:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="31" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
    	<td colspan="2" class=alt><b><%=varbroedkrumme%> medarbejdertype</b><br />
		<%if dbfunc = "dbred" then%>
		Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b>
		<%end if%>
		</td>
	</tr>
	<tr>
		<td valign="top" ><img src="../ill/tabel_top.gif" width="1" height="40" alt="" border="0"></td>
		<td width=150><br><b>Navn:</b></td>
		<td><br><input type="text" name="FM_navn" value="<%=strNavn%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="40" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="70" alt="" border="0"></td>
		<td valign="top" style="padding-top:3;"><b>Normeret tid: </b></td>
		<td valign="top">
		<table cellspacing=3 cellpadding=0 border=0><tr>
		<td>
		Mandag</td><td>
		Tirsdag</td><td>
		Onsdag</td><td>
		Torsdag</td><td>
		Fredag</td><td>
		Lørdag</td>
		<td>Søndag</td></tr>
		<tr>
		<td><input type="text" name="FM_norm_man" value="<%=normtimer_man%>" size="3" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		</td><td><input type="text" name="FM_norm_tir" value="<%=normtimer_tir%>" size="3" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		</td><td><input type="text" name="FM_norm_ons" value="<%=normtimer_ons%>" size="3" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		</td><td><input type="text" name="FM_norm_tor" value="<%=normtimer_tor%>" size="3" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		</td><td><input type="text" name="FM_norm_fre" value="<%=normtimer_fre%>" size="3" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		</td><td><input type="text" name="FM_norm_lor" value="<%=normtimer_lor%>" size="3" style="!border: 1px; background-color: #FFFFFF; border-color: darkred; border-style: solid;">
		</td>
		<td><input type="text" name="FM_norm_son" value="<%=normtimer_son%>" size="3" style="!border: 1px; background-color: #FFFFFF; border-color: darkred; border-style: solid;">
		</td>
		</tr>
		<tr><td colspan=7>Timer&nbsp;&nbsp;<b>Total:</b>&nbsp;<u><%=ugetotal%></u></td></tr>
		</table></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="70" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td><b>Timepris 1:</b><br />
		(generel timepris)</td>
		<td><input type="text" name="FM_timepris" value="<%=strTimepris%>" size="10" style="border:1px #86B5E4 solid; background-color:#FFFFFF;">&nbsp;
		<%call valutaKoder(0, tp0_valuta) %>
		
		
		
		</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	
	<!--
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="40" alt="" border="0"></td>
		<td valign="top" style="padding-top:5px;"><b>Timepris 1:</b></td>
		<td><input type="hidden" name="FM_timepris1" value="<%=strTimepris1%>" size="10" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">&nbsp;
		<%call valutaKoder(1, valuta) %><br />
		 (sæt alt. timepris 1 = generel tp. pga. stamakt) &nbsp;&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="40" alt="" border="0"></td>
	</tr>
	-->
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td><b>Timepris 2:</b></td>
		<td><input type="text" name="FM_timepris2" value="<%=strTimepris2%>" size="10" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">&nbsp;
		<%call valutaKoder(2, tp2_valuta) %>&nbsp;&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td><b>Timepris 3:</b></td>
		<td><input type="text" name="FM_timepris3" value="<%=strTimepris3%>" size="10" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">&nbsp;
		<%call valutaKoder(3, tp3_valuta) %>&nbsp;&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td><b>Timepris 4:</b></td>
		<td><input type="text" name="FM_timepris4" value="<%=strTimepris4%>" size="10" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">&nbsp;
		<%call valutaKoder(4, tp4_valuta) %>
		&nbsp;&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td><b>Timepris 5:</b></td>
		<td><input type="text" name="FM_timepris5" value="<%=strTimepris5%>" size="10" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">&nbsp;
		<%call valutaKoder(5, tp5_valuta) %>
		&nbsp;&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="120" alt="" border="0"></td>
		<td valign="top"><br><b>Intern kostpris:</b></td>
		<td valign="top"><br><input type="text" name="FM_kostpris" value="<%=dubKostpris%>" size="10" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">&nbsp;DKK
		<br><br><br>
		<input type="image" src="../ill/<%=varSubVal%>.gif"><br><br>&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="120" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan="2" valign="bottom" style="border-bottom:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>	
	</form>
	</table>
	
	<div style="position:absolute; width:200px; left:530px; top:50px; background-color:#FFFFe1; border:1px red dashed; padding:10px;">
	<b>
    <img src="../ill/about.gif" /> Info:</b><br />
		<b>Generel timepris</b> <br />
		Den generelle timepris er den timepris der automatisk bliver tildelt på et job 
		ved joboprettelse. De andre timepriser kan vælges ved redigering af job eller 
		de kan forvælges på stam-aktiviteter så den bliver forvalgt på den enkelte stam-aktivitet.
		
       
	</div>
	
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
	<%call tsamainmenu(6)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	call medarbtopmenu()
	%>
	</div>
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20; top:132;">
	<h3><img src="../ill/ac0019-24.gif" width="24" height="24" alt="" border="0">&nbsp;Medarbejdertyper.</h3>
	<div id="opretny" style="position:absolute; left:490; top:10; visibility:visible;">
	<a href="medarbtyper.asp?menu=medarb&func=opret">Opret ny type&nbsp;<img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a>
	</div>
	<table cellspacing="0" cellpadding="0" border="0" width="600" bgcolor="#EFF3FF">
	<tr bgcolor="5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="31" alt="" border="0"></td>
		<td colspan="7" valign="top"><img src="../ill/tabel_top.gif" width="587" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="31" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td height="30" class=alt><a href="medarbtyper.asp?menu=medarb&sort=nr" class=alt><b>Id</b></a></td>
		<td width="60" class=alt><a href="medarbtyper.asp?menu=medarb&sort=navn" class=alt><b>Navn</b></a></td>
		<td align="left" width="70" class=alt>Medlemmer</td>
		<td class=alt align=right>General timepris</td>
		<td class=alt align=right>Intern Kostpris</td>
		<td class=alt align=right>Normeret tid</td>
		<td>&nbsp;&nbsp;</td>
	</tr>
	<%
	sort = Request("sort")
	if sort = "navn" then
	odrBy = " ORDER BY type"
	else
	odrBy = " ORDER BY id"
	end if
	
	strSQL = "SELECT mt.id, type, timepris, kostpris, normtimer_son, "_
	&" normtimer_man, normtimer_tir, normtimer_ons, normtimer_tor, "_
	&" normtimer_fre, normtimer_lor, tp0_valuta, v.valutakode FROM medarbejdertyper mt "_
	&" LEFT JOIN valutaer v ON (v.id = tp0_valuta) "_
	&""& odrBy
	
	
	oRec.open strSQL, oConn, 3
	
	x = 0
	c = 0
	ugetotal = 0
	while not oRec.EOF
	ugetotal = formatnumber(oRec("normtimer_son") + oRec("normtimer_man") + oRec("normtimer_tir") + oRec("normtimer_ons") + oRec("normtimer_tor") + oRec("normtimer_fre") + oRec("normtimer_lor"), 2)
	
	 
	strSQL2 = "SELECT Mid FROM medarbejdere WHERE Medarbejdertype = "& oRec("id")
	oRec2.open strSQL2, oConn, 3
	while not oRec2.EOF 
	x = x + 1
	oRec2.movenext
	wend
	oRec2.close
	Antal = x
	
	
	select case right(c,1)
	case 0,2,4,6,8
	trbg = "#EFF3FF"
	case else
	trbg = "#FFFFFF"
	end select%>
	<tr>
		<td bgcolor="#5582D2" colspan="9"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="<%=trbg%>" onmouseover="mOvr('gift',this,'#B4C7EF');" onmouseout="mOut(this,'<%=trbg%>');">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td><%=oRec("id")%></td>
		<td height="20">&nbsp;&nbsp;<a href="medarbtyper.asp?menu=medarb&func=red&id=<%=oRec("id")%>"><%=oRec("type")%> <img src="../ill/edit.gif" alt="Se medarbejdertype" border="0"> </a></td>
		<td><a href="medarbtyper.asp?menu=medarb&func=med&id=<%=oRec("id")%>" class=vmenuglobal>Se medlemmer (<%=antal%>)</a></td>
		<td align=right><%=oRec("timepris")%>
            &nbsp;<%=oRec("valutakode") %></td>
		<td align=right><%=oRec("kostpris")%> DKK</td>
		<td align=right><%=formatnumber(ugetotal)%> timer</td>
		<%if x > 0 then%>
		<td>&nbsp;</td>
		<%else%>
		<td style="padding-left:10px;"><a href="medarbtyper.asp?menu=medarb&func=slet&id=<%=oRec("id")%>"><img src="../ill/slet_16.gif" width="16" height="16" alt="" border="0"></a></td>
		<%end if%>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<%
	c = c + 1
	x = 0
	oRec.movenext
	wend
	%>
	<tr bgcolor="#5582D2">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan="7" valign="bottom"><img src="../ill/tabel_top.gif" width="587" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>		
	</table>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
