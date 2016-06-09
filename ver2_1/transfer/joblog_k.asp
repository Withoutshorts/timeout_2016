<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<SCRIPT language=javascript src="inc/serviceaft_osigt_func.js"></script>
<%
if session("user") = "" then
%>
<!--#include file="../inc/regular/header_inc.asp"-->
<%
	errortype = 5
	call showError(errortype)
	else
	
	thisfile = "joblog_k"
	'**** Job ****
	
	
	if request("uselogin") = "n" then
		uselogin = "n"
		kundeid = request("usekid")
	else
		uselogin = "y"
		kundeid = session("Mid")
	end if
	
	
	
	'*** Finder det valgte job ***
	if len(request("FM_seljob")) <> 0 then
	jobnr = request("FM_seljob")
	else
	jobnr = 0
	end if
	
	if len(request("func")) <> 0 then
	func = request("func")
	else
	func = "tim"
	end if
	
	if len(request("FM_orderby_medarb")) <> 0 AND request("FM_orderby_medarb") <> 0 then
	fordelpamedarb = 1
	orderByKri = "Tjobnavn, Tjobnr, Tmnavn, Tdato DESC, Tmnavn"
	else
	fordelpamedarb = 0
	orderByKri = "Tjobnavn, Tjobnr, Tdato DESC"
	end if
	
	
	if len(request("FM_visgodkendte")) <> 0 then
	visgodkendte = request("FM_visgodkendte")
		select case visgodkendte
		case 1
		visGodkendtfilterKri = " AND godkendtstatus = 1"
		case 2
		visGodkendtfilterKri = " AND godkendtstatus = 0"
		case else 
		visGodkendtfilterKri = " AND godkendtstatus <> 99"
		end select
	else
	visgodkendte = 0
	visGodkendtfilterKri = " AND godkendtstatus <> 99"
	end if
	
	
	if len(request("FM_ignorerperiode")) <> 0 then
	visheleperiode = request("FM_ignorerperiode")
	vishelperchk = "CHECKED"
	else
	visheleperiode = 0
	vishelperchk = ""
	end if
	
	session.lcid = 1030
	
	'*** Indsætter kundekommentar ***
	if func = "kom" then
	
	timerid = request("FM_timerid")
	dagsdato = year(now) &"/"& month(now) &"/"& day(now)
	str_dagsdato = formatdatetime(dagsdato, 1)
	
	'if len(request("FM_godkendTimer")) <> 0 then
	'godkendtimer = 1
	'else
	godkendtimer = 0
	'end if
	
	komm = "<br><font color=#5582d2><u>"& str_dagsdato &", "& session("user") &":</u><br></font>" & replace(request("FM_kom"), "'", "''")
	strSQLkomm = "INSERT INTO incidentnoter (editor, dato, note, timerid, todoid, status) "_
	&" VALUES ('"& session("user") &"', '"& dagsdato &"', '"& komm &"', "& timerid &", 0, "& godkendtimer &")"
	
	'Response.write strSQLkomm
	oConn.execute(strSQLkomm)
	Response.redirect "joblog_k.asp?func=tim&FM_seljob="&jobnr&"&uselogin="&uselogin&"&usekid="&kundeid&"&FM_ignorerperiode="&visheleperiode&"&FM_orderby_medarb="&fordelpamedarb&"&FM_visgodkendte="&visgodkendte
	end if
	
	
	'**** kunde/overordnet godkender *** 
	if func = "godkendt" then
	
	intGodkendtetimer = Split(request("FM_godkend"), ", ")
			For b = 0 to Ubound(intGodkendtetimer)
				strSQL = "UPDATE timer SET godkendtstatus = 1, godkendtstatusaf = '"& session("user") &"' WHERE tid = " & intGodkendtetimer(b)
				'Response.write strSQL &"<br>"
				oConn.execute(strSQL)
			next
	
	Response.redirect "joblog_k.asp?func=tim&FM_seljob="&jobnr&"&uselogin="&uselogin&"&usekid="&kundeid&"&FM_ignorerperiode="&visheleperiode&"&FM_orderby_medarb="&fordelpamedarb&"&FM_visgodkendte="&visgodkendte
	
	end if
	
	
	
	if request("print") <> "j" then%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<%
	else%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->	
	<%end if%>
	
	<!--#include file="inc/convertDate.asp"-->
	<script language="javascript">
	<!--
	
	function hideprint(){
	document.getElementById("printknap").style.visibility = "hidden"
	document.getElementById("printknap").style.display = "none"
	}
	
	function showkommentar(tid){
	document.getElementById("FM_timerid").value = tid
	document.getElementById("kommentar").style.visibility = "visible"
	document.getElementById("kommentar").style.display = ""
	document.getElementById("FM_kom").focus()
	document.getElementById("FM_kom").select()
	}
	
	function hidekommentar(){
	document.getElementById("FM_timerid").value = 0
	document.getElementById("kommentar").style.visibility = "hidden"
	document.getElementById("kommentar").style.display = "none"
	}
	
	
	
	function checkAll(field)
	{
	field.checked = true;
	for (i = 0; i < field.length; i++)
		field[i].checked = true ;
	}
	
	function uncheckAll(field){
	field.checked = false;
	for (i = 0; i < field.length; i++)
		field[i].checked = false ;
	}
	
	
	
	//-->
	</script>
	
	
	<!---- Logo ------->
	<%
	
	
		if request("print") <> "j" then
		ptopher = 82
		plefther = 640
		else
		ptopher = 90
		plefther = 690
		end if
		%>
		<div id="ill" style="position:absolute; left:<%=plefther%>; top:<%=ptopher%>; z-index=10;">
			<%
			strSQL = "SELECT useasfak, logo, id, filnavn FROM kunder, filer WHERE kid = "& kundeid &" AND filer.id = kunder.logo"
			
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
			logonavn = "<img src='../inc/upload/"&lto&"/"&oRec("filnavn")&"' alt='' border='0'>"
			nologo = "y"
			else
			logonavn = "(Firma logo kan uploades her.)"
			nologo = "n"
			end if
			oRec.close
			'Response.write "<font class=megetlillesilver>Denne service er leveret af:<br>"& logonavn
			Response.write logonavn
			
			%>
		</div>
	<%'end if%>
	<!---- LogoSlut------->
		
	<%
	'*** Konfigurerer menupkt navne efter lto. ***
	select case lto 
	case "kringit"
	pkt1 = "Seviceordre"
	pkt2 = "Aftaler"
	pkt3 = "Filarkiv"
	pkt4 = "Infobase"
	pkt1oskrift = "Seviceordre"
	case else
	pkt1 = "Timeregistreringer"
	pkt2 = "Aftaler"
	pkt3 = "Filarkiv"
	pkt4 = "Infobase"
	pkt1oskrift = "Job"
	end select
	
			'*** Hvis der ikke vises print side udskrives menu her **************
			if request("print") <> "j" then
	
					
					'*************** Henter kunde oplysninger ************************************
					strSQL = "SELECT Kid, kkundenavn, kkundenr, adresse, postnr, city, land, telefon, cvr, logo FROM kunder WHERE Kid =" & kundeid		
					oRec.open strSQL, oConn, 3
					if not oRec.EOF then 
						intKnr = oRec("kkundenr")
						strKnavn = oRec("kkundenavn")
						strKadr = oRec("adresse")
						strKpostnr = oRec("postnr")
						strBy = oRec("city")
						strLand = oRec("land")
						intKid = oRec("Kid")
						intCVR = oRec("cvr")
						intTlf = oRec("telefon")
						logo = oRec("logo")
					end if
					
					oRec.close
					%>
					<br><br>
					<%
					topknavn = 60
					%>
					<table border="0" cellpadding="0" cellspacing="0" style="position:absolute; left:10; top:<%=topknavn%>; z-index:0;">
					<tr>
						<td colspan="3"><br><br><img src="../ill/ac0009-24.gif" width="24" height="24" alt="" border="0">&nbsp;<b>Kontakt:</b></td>
					</tr>
					<tr>
						<td valign="top" colspan="3"><%=strKnavn%><br>
						<%=strKadr%><br><%=strKpostnr%>&nbsp;<%=strBy%>
						<%if len(trim(intTlf)) <> 0 then%>
						<br>Tlf:&nbsp;<%=intTlf%><br>
						<%end if%>
						&nbsp;</td>
					</tr>
					<tr>
						<td colspan="3" width="150"><br><b>Oversigter:</b><br>
						<img src='../ill/pile_selected.gif' alt='' border='0'>&nbsp;<a href="joblog_k.asp?func=tim&usekid=<%=intKid%>&uselogin=<%=uselogin%>&FM_seljob=<%=jobnr%>" class='vmenu'><%=pkt1%></a><br>
						<img src='../ill/pile_selected.gif' alt='' border='0'>&nbsp;<a href="joblog_k.asp?func=aft&usekid=<%=intKid%>&uselogin=<%=uselogin%>&FM_seljob=<%=jobnr%>" class='vmenu'><%=pkt2%></a><br>
						<img src='../ill/pile_selected.gif' alt='' border='0'>&nbsp;<a href="javascript:popUp('filer.asp?kundeid=<%=intKid%>&jobid=0&kundelogin=1', '600', '600','200', '50')" target="_self" class='vmenu'><%=pkt3%></a><br>
						<img src='../ill/pile_selected.gif' alt='' border='0'>&nbsp;<a href="infobase.asp?menu=kund&usekview=j&id=<%=intKid%>&uselogin=<%=uselogin%>&FM_seljob=<%=jobnr%>" class='vmenu'><%=pkt4%></a>
						</td>
					</tr>
					
				</table>
				<!-- slut ventre menu -->
				<%end if%>
	
	
	
	<%
	if request("print") = "j" then
	'*** Sideheader ****
	%>
	<div id="sheader" style="position:absolute; left:0; top:0; visibility:visible; z-index:100;">
	<%
		select case lto 
		case "kringit"
		topbarprint = "kringit/kring_value_print"
		case else
		topbarprint = "logo_topbar_print"
		end select 
		%>
		<table cellspacing="0" cellpadding="0" border="0">
		<tr>
			<td bgcolor="#ffffff" width="650"><img src="../ill/<%=topbarprint%>.gif" alt="" border="0"></td>
		</tr>
		<tr>
			<td bgcolor="#FFFFFF" style="padding-right:2; padding-top:3;">
			<p id="printknap" name="printknap" style="position:relative; visibility:visible; display:;">
			<!--<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0">&nbsp;Tilbage</a>
			<img src="../ill/blank.gif" width="10" height="1" alt="" border="0">-->
			<a href="javascript:window.close()"><img src="../ill/luk_xp.gif" width="30" height="28" alt="" border="0">&nbsp;Luk</a>
			<img src="../ill/blank.gif" width="10" height="1" alt="" border="0">
			<a href="javascript:window.print()" onClick="hideprint();"><img src="../ill/print_xp.gif" width="28" height="30" alt="" border="0">&nbsp;Print</a>
			</p>
			</td>
		</tr>
		</table>
		</div>
	<%end if%>
	
	
	
	
	<%'**** Sideindhold ****
	
	if request("print") <> "j" then
	sideDivTop = "55px"
	sideDivLeft = "190px"
	else
	sideDivTop = "80px"
	sideDivLeft = "20px"
	end if
	
	%>
	<div id="sindhold" style="position:absolute; left:<%=sideDivLeft%>; top:<%=sideDivTop%>; visibility:visible; z-index:100;">
	<%
	'***** Filter ****
	
	
		if len(request("filter_per")) <> 0 then
		filterKri_kundelogin = request("filter_per")
		else
			if len(request.cookies("so_filterKri")) <> 0 then
			filterKri_kundelogin = request.cookies("so_filterKri")
			else
 			filterKri_kundelogin = 0
			end if 
		end if
		
		select case filterKri_kundelogin
		case 0
		chkfilt0 = "CHECKED"
		chkfilt1 = ""
		chkfilt2 = ""
		useFilterKri = ""
		case 1
		chkfilt0 = ""
		chkfilt1 = ""
		chkfilt2 = "CHECKED"
		useFilterKri = " AND advitype = 1 "
		case 2
		chkfilt0 = ""
		chkfilt1 = "CHECKED"
		chkfilt2 = ""
		useFilterKri = " AND advitype = 2 "
		end select
		
		
			filterStatus = request("status")
			select case filterStatus
			case "vislukkede"
			chkfilt3 = ""
			chkfilt4 = ""
			chkfilt5 = "CHECKED"
			useFilterKri = useFilterKri & " AND status = 0 "
			case "visalle"
			chkfilt3 = "CHECKED"
			chkfilt4 = ""
			chkfilt5 = ""
			useFilterKri = useFilterKri & ""
			case else
			chkfilt3 = ""
			chkfilt4 = "CHECKED"
			chkfilt5 = ""
			useFilterKri = useFilterKri & " AND status = 1 "
			end select
			
		
	
	if request("print") <> "j" then%>
	<br>
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF">
	<form action="joblog_k.asp?menu=<%=menu%>&func=<%=func%>&FM_usedatokri=1" method="post">
	<input type="hidden" name="usekid" value="<%=intKid%>">
	<input type="hidden" name="uselogin" value="<%=uselogin%>">
	<tr bgcolor="#5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td colspan=2 width=384 valign="top" style="border-top:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td colspan=2 valign="top" class="alt">&nbsp;<b>Vælg filter kriterier:</b></td>
	</tr>
	
			
	
	<%select case func%>
	
	<%case "tim"%>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="19" alt="" border="0"></td>
		<td colspan=2 style="padding-top:5px;">&nbsp;<b><%=pkt1oskrift%>:</b></td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="19" alt="" border="0"></td>
	</tr>	
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="45" alt="" border="0"></td>
		<td colspan=2 style="padding-left:15;">
		<%
		strSQL = "SELECT jobnavn, jobnr, id, budgettimer, jobslutdato, jobTpris FROM job WHERE jobknr =" & kundeid &" AND (fakturerbart = 1 AND kundeok = 1) OR (fakturerbart = 1 AND kundeok = 2 AND jobstatus = 0) ORDER BY jobnavn"		
		'Response.write strSQL
		strJobnrs = 0
		%>
		<select name="FM_seljob" id="FM_seljob" style="width:300; font-size:9;">
		<option value="0">Alle <%=pkt1oskrift%></option>
		<%
		oRec.open strSQL, oConn, 3 
		while not oRec.EOF 
		if cint(jobnr) = oRec("jobnr") then
		seljid = "SELECTED"
		else
		seljid = ""
		end if 
		
		strJobnrs = strJobnrs & "," & oRec("jobnr")
		
		%>
		<option value="<%=oRec("jobnr")%>" <%=seljid%>><%=oRec("jobnavn")%> &nbsp;(<%=oRec("jobnr")%>)</option>
		<%
		oRec.movenext
		wend
		oRec.close 
		%>
		</select>
		<br>
		
		<%if request("FM_viskorsel") = "1" then
		kchk = "CHECKED"
		else
		kchk = ""
		end if%>
		<input type="checkbox" name="FM_viskorsel" id="FM_viskorsel" value="1" <%=kchk%>>Vis kørsel.
		
		</td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="45" alt="" border="0"></td>
	</tr>
	
	<%end select%>
	
	
	<%select case func 
	case "aft", "tim"
	%>
	<tr>
		<td valign="top" style="border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td colspan=2 style="padding-top:3px;"><b>Periode:</b>
		<%if func = "aft" then%>
		<br>Vis kun aftaler med <b>startdato</b> i det valgte interval:
		<%end if%>
		</td>
		<td valign="top" align=right style="border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>	
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="45" alt="" border="0"></td>
		<td>&nbsp;</td>
		<td style="padding-top:3px; padding-left:15px;"><!--#include file="inc/weekselector_s.asp"--></td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="45" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td colspan=2 valign=top style="padding-left:15; padding-top:2;">
		<input type="checkbox" name="FM_ignorerperiode" id="FM_ignorerperiode" value="1" <%=vishelperchk%>>Ignorer periode.<br></td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	
	<%if func = "tim" then%>
	
	<tr>
		<td valign="top" style="border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td colspan=2 style="padding-top:10px;"><b>Øvrige filter kriterier:</b></td>
		<td valign="top" align=right style="border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>	
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="50" alt="" border="0"></td>
		<td colspan=2 valign=top style="padding-left:15; padding-top:2;">
	
		<%if fordelpamedarb = 1 then
		chkFordelmedarb = "CHECKED"
		else
		chkFordelmedarb = ""
		end if%>
		
		<input type="checkbox" name="FM_orderby_medarb" id="FM_orderby_medarb" value="1" <%=chkFordelmedarb%>>Fordel på medarbejdere.<br>
		
		<%
		
		Select case visgodkendte
		case 1
		chkVisgodkendte = ""
		chkVisgodkendte1 = "CHECKED"
		chkVisgodkendte2 = ""
		case 2
		chkVisgodkendte = ""
		chkVisgodkendte1 = ""
		chkVisgodkendte2 = "CHECKED"
		case else
		chkVisgodkendte = "CHECKED"
		chkVisgodkendte1 = ""
		chkVisgodkendte2 = ""
		end select%>
		<input type="radio" name="FM_visgodkendte" id="FM_visgodkendte" value="0" <%=chkVisgodkendte%>>Vis Alle
		<input type="radio" name="FM_visgodkendte" id="FM_visgodkendte" value="1" <%=chkVisgodkendte1%>>Vis kun godkendte
		<input type="radio" name="FM_visgodkendte" id="FM_visgodkendte" value="2" <%=chkVisgodkendte2%>>Vis kun ikke godkendte
		</td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="50" alt="" border="0"></td>
	</tr>
	<%end if%>
	
	
	<%end select%>
	
	<%if func = "aft" then%>
	<input type="hidden" name="FM_seljob" value="<%=jobnr%>">
	<tr>
		<td valign="top" style="border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td colspan=2><br>&nbsp;&nbsp;<b>Type:</b><br>
			<input type="radio" name="filter_per" id="filter_per" value="0" <%=chkfilt0%>>Vis Alle&nbsp;&nbsp;
			<input type="radio" name="filter_per" id="filter_per" value="1" <%=chkfilt2%>><b>Enheds/klip</b> aftaler.&nbsp;&nbsp;
			<input type="radio" name="filter_per" id="filter_per" value="2" <%=chkfilt1%>><b>Periode</b> aftaler.
		</td>
		<td valign="top" align=right style="border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top" style="border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td colspan=2><br>&nbsp;&nbsp;<b>Status:</b><br>
			<input type="radio" name="status" id="status" value="visalle" <%=chkfilt3%>>Vis Alle&nbsp;&nbsp;
			<input type="radio" name="status" id="status" value="visaktive" <%=chkfilt4%>>Aktive.&nbsp;&nbsp;
			<input type="radio" name="status" id="status" value="vislukkede" <%=chkfilt5%>>Lukkede.&nbsp;&nbsp;
		</td>
		<td valign="top" align=right style="border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	
	<%end if%>
		
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="32" alt="" border="0"></td>
		<td colspan=2 style="padding-right:35px;" align=right><input type="image" src="../ill/statpil.gif"></td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="32" alt="" border="0"></td>
	</tr>	
	<tr bgcolor="#EFF3FF">
			<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
			<td colspan=2 valign="bottom" style="border-bottom:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
			<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr></form>
	<tr><td colspan=4 bgcolor="#ffffff"><br>
	<a href="javascript:NewWin_large('joblog_k.asp?FM_orderby_medarb=<%=fordelpamedarb%>&func=<%=func%>&FM_seljob=<%=jobnr%>&print=j&uselogin=<%=uselogin%>&usekid=<%=kundeid%>&FM_ignorerperiode=<%=visheleperiode%>&FM_visgodkendte=<%=visgodkendte%>&FM_start_mrd=<%=request("FM_start_mrd")%>&FM_start_dag=<%=request("FM_start_dag")%>&FM_start_aar=<%=request("FM_start_aar")%>&FM_slut_dag=<%=request("FM_slut_dag")%>&FM_slut_mrd=<%=request("FM_slut_mrd")%>&FM_slut_aar=<%=request("FM_slut_aar")%>&strJobnrs=<%=strJobnrs%>&status=<%=request("status")%>&filter_per=<%=request("filter_per")%>');" target="_self" class="vmenu">&nbsp;Print venlig version&nbsp;<img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a>
	<br></td></tr>
	</table>
	<%end if '** Print
	
	
	if request("print") = "j" then
	
	strMrd = Request.Cookies("datoer")("st_md")
	strDag = Request.Cookies("datoer")("st_dag")
	strAar = Request.Cookies("datoer")("st_aar") 
	strDag_slut = Request.Cookies("datoer")("sl_dag")
	strMrd_slut = Request.Cookies("datoer")("sl_md")
	strAar_slut = Request.Cookies("datoer")("sl_aar")
	
	strStartDato = strAar&"/"&strMrd&"/"&strDag
	strSlutDato = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut
	else
	strStartDato = strAar&"/"&strMrd&"/"&strDag
	strSlutDato = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut
	end if
	
	
	if func <> "fil" then
	
		if request("print") = "j" then
				if visheleperiode = "1" then
				Response.write "<br>Periode afgrænsning: Ingen"
				else
				Response.write "<br>Periode afgrænsning: " & formatdatetime(strDag&"/"&strMrd&"/"&strAar, 2) & " til " & formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 2)
				end if
		end if
	
	end if
	
	
	'*** Sideindhold ***
	select case func
	case "aft"
	'****************************** Aftaler. ***********************************'
	%>
	<!--#include file="inc/serviceaft_osigt_inc.asp"-->
	<%
	case "tim"
	
	showtimerkorsel = 1
	call visjoblog(showtimerkorsel) 'timer
	
	if request("FM_viskorsel") = "1" then
	showtimerkorsel = 2
	call visjoblog(showtimerkorsel) 'km
	end if
	
	%>
		<div id=kommentar name=kommentar style="position:absolute; left:90; top:180; visibility:hidden; display:none; width:400; background-color:#ffffe1; border:1px #003399 solid;">
		<table border=0 cellpadding=0 cellspacing=0 width=400>
		<form action="joblog_k.asp?func=kom" method="post">
		<input type="hidden" name="FM_timerid" id="FM_timerid" value="0">
		<input type="hidden" name="usekid" value="<%=intKid%>">
		<input type="hidden" name="uselogin" value="<%=uselogin%>">
		<input type="hidden" name="FM_seljob" value="<%=jobnr%>">
		<input type="hidden" name="FM_ignorerperiode" value="<%=visheleperiode%>">
		<input type="hidden" name="FM_orderby_medarb" value="<%=fordelpamedarb%>">
		<input type="hidden" name="FM_visgodkendte" value="<%=visgodkendte%>">
		
		<tr><td style="padding:5px;">
		<b>Kommentar:</b><br>
		<textarea cols="45" rows="5" name="FM_kom" id="FM_kom"></textarea><br>
		<!--<input type="checkbox" name="FM_godkendTimer" id="FM_godkendTimer" value="1">Godkend timer/enheder.-->
		</td></tr>	
		<tr><td align=center>
		<input type="submit" value="Opdater">
		</td></tr>
		<tr><td style="padding:5px;"><a href="#" onClick='hidekommentar()' class=vmenu>Luk vindue</a></td></tr>
		</form>
		</table>	
		</div>	
	<%
	end select

	
	
	
	
	'****** Subs og Funktioner ****
	public lastmedarbnavn, medarbtimer, medarbEnheder
	sub fordelpamedarbejdere%>
					
					<tr>
						<td bgcolor="#5582D2" colspan="8"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					</tr>
					<tr>
						<td bgcolor="<%=medbgcol%>" align=right colspan="8" height=20 style="border-left:1 #003399 solid; border-right:1 #003399 solid; padding-right:10;"><%=lastmedarbnavn%>:&nbsp; <b><%=formatnumber(medarbtimer, 2)%></b> <%=enhed%>. = <b><%=medarbEnheder%></b> Enheder.</td>
					</tr>
	
	<%end sub
	
	
	sub subtotaler
	%>
	<tr>
		<td bgcolor="#5582D2" colspan="8"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr>
	<%
	
	'if showtimerkorsel = 1 then
	
	'**** Timer indtastet total på job ***
	strSQL2 = "SELECT sum(timer) AS sumtimer FROM timer WHERE tfaktim <> 5 AND tjobnr = "& jobnr
	oRec2.open strSQL2, oConn, 3 
	if not oRec2.EOF then
	timerTotpajob = oRec2("sumtimer")
	end if
	oRec2.close  
	
	'else
	
	''**** KM indtastet på job ***
	'strSQL = "SELECT sum(timer) AS sumtimer FROM timer WHERE tfaktim = 5 AND tjobnr = "& jobnr
	'oRec.open strSQL, oConn, 3 
	'if not oRec.EOF then
	'timerTotpajob = oRec("sumtimer")
	'end if
	'oRec.close  
	
	'end if
	
	if len(medarbEnhederPer) <> 0 then
	medarbEnhederPer = medarbEnhederPer
	else
	medarbEnhederPer = 0
	end if
	
	if len(timerTotpajob) <> 0 then
	timerTotpajob = timerTotpajob
	else
	timerTotpajob = 0
	end if
	
	if len(timerTotpajobPer) <> 0 then
	timerTotpajobPer = timerTotpajobPer
	else
	timerTotpajobPer = 0
	end if
	
	if len(jobForkalkulerettimer) <> 0 then
	jobForkalkulerettimer = jobForkalkulerettimer
	else
	jobForkalkulerettimer = 1
	end if
	
	if len(jobForkalfakbare) <> 0 then
	jobForkalfakbare = jobForkalfakbare
	else
	jobForkalfakbare = 0
	end if
	
	if len(fakbareTimerTotpajob) <> 0 then
	fakbareTimerTotpajob = fakbareTimerTotpajob
	else
	fakbareTimerTotpajob = 0
	end if
	%>
	
		<td bgcolor="#d6dff5" align=right colspan="8" valign="top" height=20 style="border-left:1 #003399 solid; border-right:1 #003399 solid; padding-right:10; padding-top:3;">
		Ialt på <b><%=lastjobnavn%></b> i den <u>valgte periode</u>: <b><%=formatnumber(timerTotpajobPer, 2)%></b> <%=enhed%>.<br>
		Svarende til <b><%=formatnumber(medarbEnhederPer, 2)%></b> Enheder.<br>
		<%if showtimerkorsel = 1 then%>
		<!--Resterende timer total:  <b><=formatnumber(jobForkalfakbare - timerTotpajob, 2)%></b> <=enhed%>. <br>
		Procent af job afsluttet: <b><
		if jobForkalkulerettimer > 0 then
		Response.write formatpercent(timerTotpajob/jobForkalkulerettimer, 0)
		else
		Response.write "0 %"
		end if%></b><br> -->
		Kontakt den jobansvarlige: <a href="mailto:<%=emailjobans1%>">her</a>.
		<%end if%><br>&nbsp;
		</td>
	</tr>
	<%
	medarbEnhederPer = 0
	timerTotpajobPer = 0
	timerTotpajob = 0
	jobForkalkulerettimer = 1
	jobForkalfakbare = 0
	fakbareTimerTotpajob = 0 
	end sub
	
	
	
	
	


public medbgcol, timerTotpajob, timerTotpajobPer, medarbEnhederPer, jobForkalkulerettimer, jobForkalfakbare, fakbareTimerTotpajob, showtimerkorsel, enhed, lastjobnavn, emailjobans1, strJobnrs 
function visjoblog(showtimerkorsel)

if showtimerkorsel = 1 then
enhed = "Timer"
tft = " tfaktim <> 5 "
medbgcol = "#ffff99"
dnavn = "tim"
else
enhed = "Km"
tft = " tfaktim = 5 "
medbgcol = "#cccccc"
dnavn = "km"
end if

%>

<br><br><br>	
<h3><%=enhed%>:</h3>



<table border="0" width=80% cellpadding="0" cellspacing="0" bgcolor="#ffffff" style="page-break-after: always;">
<form action="joblog_k.asp?menu=<%=menu%>&func=godkendt" method="post" id="godkendtimer" name="godkendtimer">
<input type="hidden" name="usekid" value="<%=intKid%>">
<input type="hidden" name="uselogin" value="<%=uselogin%>">
<input type="hidden" name="FM_seljob" value="<%=jobnr%>">
<input type="hidden" name="FM_ignorerperiode" value="<%=visheleperiode%>">
<input type="hidden" name="FM_orderby_medarb" value="<%=fordelpamedarb%>">
<input type="hidden" name="FM_visgodkendte" value="<%=visgodkendte%>"> 

<%

	
	lastjobnavn = ""
	lastmedarb = 0
	lastSid = 0
	
	if visheleperiode = "1" then
	perintKri = ""
	else
	perintKri = " AND (Tdato >= '"& strStartDato &"' AND Tdato <= '"& strSlutDato &"') "
	end if
	
	if jobnr <> 0 then
	tjobnrKri = " tjobnr = "& jobnr 
	else
	
	if request("print") <> "j" then
	strJobnrs = strJobnrs
	else
	strJobnrs = request("strJobnrs")
	end if 
		
		if len(strJobnrs) <> 0 then
			valgtejobnr = split(strJobnrs, ",")
			
			for t = 0 to UBOUND(valgtejobnr)
				if t = 0 then
				tjobnrKri = "(tjobnr = " & valgtejobnr(t) 
				else
				tjobnrKri = tjobnrKri & " OR tjobnr = " & valgtejobnr(t)
				end if
			next 
			
			tjobnrKri = tjobnrKri & ")"
		else
		tjobnrKri = " tjobnr = 0" 
		end if
	
	end if
	
	strSQL = "SELECT Tdato, TasteDato, Tjobnr, Tjobnavn, TAktivitetNavn AS Anavn,"_
	&" TAktivitetId, Tknavn, Tknr, Tmnr, Tmnavn, Timer, Tid, "_
	&" Tfaktim, TimePris, Timerkom, offentlig, s.navn AS aftnavn, "_
	&" sttid, sltid, s.id AS sid, a.faktor, godkendtstatus, godkendtstatusaf FROM timer"_
	&" LEFT JOIN job j ON (j.jobnr = "& jobnr &") "_
	&" LEFT JOIN serviceaft s ON (s.id = j.serviceaft) "_
	&" LEFT JOIN aktiviteter a ON (a.id = TAktivitetId) "_
	&" WHERE "& tjobnrKri &" "& perintKri &" AND "& tft & visGodkendtfilterKri &""_
	&" ORDER BY "& orderByKri
	oRec.open strSQL, oConn, 0, 1
	v = 0
	while not oRec.EOF
		strWeekNum = datepart("ww", oRec("Tdato"))
		id = 1
				
				if (lastmedarb <> oRec("Tmnr") OR lastjobnr <> oRec("Tjobnr")) AND fordelpamedarb = 1 then
					if v <> 0 then
					
					call fordelpamedarbejdere()%>
					
					<%
					medarbtimer = 0
					medarbEnheder = 0
					end if
				end if
				
				if lastjobnr <> oRec("Tjobnr") then
				%>
				<%if v <> 0 then
					call subtotaler%>
				<%
				timerTotpajob = 0
				end if
				
				if lastmedarb <> oRec("Tmnr") OR lastjobnr <> oRec("Tjobnr") then
					if v <> 0 then%>
					<tr>
						<td height=100 style="border-top:1px #003399 solid;" bgcolor="#ffffff" colspan="8"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					</tr>
					<%
					end if
				end if
				%>
				
				
				<tr>
					<td bgcolor="#ffffff" colspan="5" height=60>
					<img src="../ill/ac0063-24.gif" width="24" height="24" alt="" border="0">&nbsp;<%=pkt1oskrift%>: <b><%=oRec("Tjobnavn")%> (<%=oRec("Tjobnr")%>)</b><br>
					Kontakt: <b><%=oRec("Tknavn")%>&nbsp;(<%=oRec("Tknr")%>)</b><br>
					<%
					
						'*** Jobansvarlige ***
						
						strSQL2 = "SELECT mnavn, mnr, mid, job.id, jobnavn, jobnr, jobans1, jobstatus, jobstartdato, jobslutdato, budgettimer, ikkebudgettimer, email FROM job LEFT JOIN medarbejdere ON (mid = jobans1) WHERE job.jobnr = "& oRec("Tjobnr") &""
						oRec2.open strSQL2, oConn, 3 
						if not oRec2.EOF then
						jobans1txt = oRec2("mnavn") & " ("& oRec2("mnr") &")"
						jobstatus = oRec2("jobstatus")
						jobstartdato = oRec2("jobstartdato")
						jobslutdato = oRec2("jobslutdato")
						jobForkalkulerettimer = (oRec2("budgettimer") + oRec2("ikkebudgettimer"))
						jobForkalfakbare = oRec2("budgettimer")
						emailjobans1 = oRec2("email")
						strJobnavn = oRec2("jobnavn")
						strJobnr = oRec2("jobnr")
						end if
						oRec2.close
						
						
						strSQL2 = "SELECT mnavn, mnr, mid, jobans2, email FROM job, medarbejdere WHERE job.jobnr = "& oRec("Tjobnr") &" AND mid = jobans2"
						oRec2.open strSQL2, oConn, 3 
						if not oRec2.EOF then
						jobans2txt = oRec2("mnavn") & " ("& oRec2("mnr") &")"
						emailjobans2 = oRec2("email")
						end if
						oRec2.close
						
						
						if showtimerkorsel = 1 then%>
						Jobansvarlig 1: <b><a href="mailto:<%=emailjobans1%>" class=vmenu><%=jobans1txt%></a></b>
						<%if len(trim(jobans2txt)) <> 0 then%>
						<br>
	   					Jobansvarlig 2: <b><a href="mailto:<%=emailjobans2%>" class=vmenu><%=jobans2txt%></a></b>
						<%end if
						jobans1txt = ""
						jobans2txt = ""
						%><br>
						<%select case jobstatus
						case 1
						thisjobstatus = "Aktiv"
						case 2
						thisjobstatus = "Passiv"
						case 0
						thisjobstatus = "Lukket"
						end select%>
						Status: <b><%=thisjobstatus%></b><br>
						Periode: <b><%=formatdatetime(jobstartdato, 1)%> - <%=formatdatetime(jobslutdato, 1)%></b>
						<br>
						Forkalkuleret timer: <b><%=formatnumber(jobForkalkulerettimer, 2)%></b>
						<br>Aftale: <b><%=oRec("aftnavn")%></b>
						<br>&nbsp;
						<%end if%>
					</td>
					<td bgcolor="<%=jobbgcol%>" colspan="3" align=right height=50 valign="bottom">
					<%if print <> "j" then%>Godkend timeregistreringer: <br><a href="#" name="CheckAll" onClick="checkAll(document.godkendtimer.FM_godkend)"><img src="../ill/alle.gif" border="0"></a>
					&nbsp;&nbsp;<a href="#" name="UnCheckAll" onClick="uncheckAll(document.godkendtimer.FM_godkend)"><img src="../ill/ingen.gif" border="0"></a>
					<%end if%>
					</td>
				</tr>
				
				
				<!-- New header Sep. 2005 -->
				<tr bgcolor="#5582D2">
					<td rowspan=2 height=30 style="border-top:1px #003399 solid; border-left:1px #003399 solid;">&nbsp;</td>
					<td colspan=6 class='alt' valign=middle style="border-top:1px #003399 solid;">&nbsp;</td>
					<td rowspan=2 style="border-top:1px #003399 solid; border-right:1px #003399 solid;">&nbsp;</td>
				</tr>
				<tr bgcolor="#5582D2">
					<td class='alt' width="120"><b>&nbsp;&nbsp;<!--Uge-->&nbsp;&nbsp;&nbsp;Dato</b></td>
					<td class='alt'>&nbsp;&nbsp;<b>Aktivitet</b></td>
					<td class='alt' width=80 align=right style="padding-right:3px;"><b><%=enhed%></b><br>Tid</td>
					<td class='alt' width=80 align=right style="padding-right:3px;"><b>Enh.</b></td>
					<td class='alt' align=right style="padding-right:3px;"><b>Medarb.</b></td>
					<td class='alt' align=right>&nbsp;<b>Godkend</b></td>
				</tr>
				
				<%end if%>
				<tr>
					<td bgcolor="#5582D2" colspan="8"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
				</tr>
				<tr>
				<td valign="top" style="border-left:1 #003399 solid;" height="20">&nbsp;</td>
				<td valign="top" style="padding-top:3px;">
				<%=formatdatetime(oRec("Tdato"), 2)%>&nbsp;</td>
				
				<td valign="top" style="padding-top:3px;"><b>
				<%if len(oRec("Anavn")) > 35 then
				Response.write left(oRec("Anavn"), 35)&".."
				else
				Response.write oRec("Anavn")
				end if
				%></b>
				
				
					<%'**** Kundekommentarer ****%>
				
					<%
					strSQLin = "SELECT note FROM incidentnoter WHERE timerid = "& oRec("tid") &" ORDER BY id"
					oRec3.open strSQLin, oConn, 3 
					i = 0
					Timergodkendt = 0
					while not oRec3.EOF 
					
					if i = 0 then
					%>
					<font class=megetlillesort>
					<%
					end if
					
					Response.write "<br>"& oRec3("note") 
					
					'if oRec3("status") = 1 then
					'Timergodkendt = 1
					'Response.write "<br><font class=roed><b>Registrering godkendt!</b></font>"
					'end if
					
					i =  i + 1
					oRec3.movenext
					wend
					oRec3.close  
					
					%>
					</font>
					<%
				
				if oRec("godkendtstatus") = 0 then%>
				&nbsp;&nbsp;<a href="#" onClick=showkommentar('<%=oRec("tid")%>')><img src="../ill/blyant.gif" width="12" height="11" alt="Tilføj kommentar" border="0"></a>
				<%end if%>
				</td>
				
				<td align="right" style="padding-top:3px; padding-right:3px;" valign="top"><b><%=formatnumber(oRec("Timer"), 2)%></b>
				<%
				if len(oRec("sttid")) <> 0 then
					if left(formatdatetime(oRec("sttid"), 3), 5) <> "00:00" then
					Response.write "<font class=megetlillesort><br>"& left(formatdatetime(oRec("sttid"), 3), 5) & " - " & left(formatdatetime(oRec("sltid"), 3), 5)
					end if
				end if%>
				</td>
				<%
				enheder = 0
				enheder = oRec("faktor") * oRec("timer")
				%>
				<td align="right" style="padding-top:3px; padding-right:3px;" valign="top"><b><%=formatnumber(enheder, 2)%></b></td>
				<%
				medarbtimer = medarbtimer + oRec("Timer")
				medarbEnheder = medarbEnheder + enheder
				timerTotpajobPer = timerTotpajobPer + oRec("Timer")
				medarbEnhederPer = medarbEnhederPer + enheder
				
				if oRec("tfaktim") = 1 then
				fakbareTimerTotpajob = fakbareTimerTotpajob + oRec("Timer")
				end if
				
				%>
				<td align=right style="padding-top:3px; padding-right:3px;" valign="top"><%=oRec("Tmnavn")%></td>
				<td align=right style="padding-top:3px;" valign="top" class=lille>
				<%if oRec("godkendtstatus") = 0 then%>
				<input type="checkbox" name="FM_godkend" id="FM_godkend" value="<%=oRec("tid")%>">
				<%else%>
				<font color="limegreen"><b><i>V</i></b></font> &nbsp;(<%=oRec("godkendtstatusaf") %>)
				<%end if%></td>
				<td valign="top" style="border-right:1 #003399 solid;" height="20">&nbsp;</td></tr>
				
				<%if len(oRec("timerkom")) <> 0 AND oRec("offentlig") <> 1 then 'tilgængelig for kunde%>
				<tr>
					<td valign="top" style="border-left:1 #003399 solid;" height="20">&nbsp;</td>
					<td>&nbsp;</td>
					<td valign="top" width=315><br><u>Medarbejder kommentar:</u><br>
					<%=oRec("timerkom")%><br>&nbsp;</td>
					<td colspan="4">&nbsp;</td>
					<td valign="top" style="border-right:1 #003399 solid;" height="20">&nbsp;</td></tr>
				</tr>
				<%end if%>
				
			<%v = v + 1
			lastmedarbnavn = oRec("Tmnavn")
			lastmedarb = oRec("Tmnr")
			lastjobnr = oRec("Tjobnr")
			lastjobnavn = oRec("Tjobnavn")
			lastSid = oRec("sid")
	oRec.movenext
	wend
	oRec.Close
	'Set oRec = Nothing

	if v = 0 then
	%>
	<tr>
		<td bgcolor="#5582D2" colspan="8"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr>
		<td bgcolor="#ffffff" colspan="8" height=60 style="border-left:1 #003399 solid; border-right:1 #003399 solid; padding-left:10;"><font class=roed><b>Der er ikke registreret <%=enhed%> i den valgte periode!</b></font></td>
	</tr>
	<%
	end if
	
	if v > 0  AND fordelpamedarb = 1 then
	
	call fordelpamedarbejdere()
	
	end if
	
	
call subtotaler
%>
<tr bgcolor="#5582D2">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan="6" valign="bottom"><img src="../ill/tabel_top.gif" width="609" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>
	<%if request("print") <> "j" then%>
	<tr bgcolor="#ffffff">
		<td colspan="8" align=right height=50 valign="bottom"><input type="submit" value="Godkend time-registreringer"></td>
	</tr>
	</form>
	<%end if%>
</table>


<br><br>&nbsp;

<%end function
	

Set oRec = Nothing
end if %>
<!--#include file="../inc/regular/footer_inc.asp"-->