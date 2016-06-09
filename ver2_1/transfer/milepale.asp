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
	
	strInternt = request("int") 
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	rdir = request("rdir")
	
	jid = request("jid")
	if len(jid) <> 0 then
	jid = jid 
	else
	jid = 0
	end if
	
	thisfile = "milepale"
	headergif = "../ill/job_milepale_header.gif"
	
	strKnr = request("kundeid")
	
	
	
	
	'************************** funktioner *********************************''
	function topfaneblade()
	%>
	
				<span id=visakt name=visakt style="position:absolute; left:600; top:148; visibility:visible; z-index:150;">
					<table cellspacing="0" cellpadding="0" border="0" bgcolor="#5582D2">
						<tr>
							<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
							<td valign="top"><img src="../ill/tabel_top.gif" width="74" height="1" alt="" border="0"></td>
							<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
						</tr>
						<tr bgcolor="#5582D2">
							<td><a href="aktiv.asp?menu=job&jobid=<%=jid%>&jobnavn=<%=jobnavn%>&rdir=<%=rdir%>" class='alt'>Aktiviteter</a></td>
						</tr>
						</table>
				</span>
				
				<span id=visfak name=visfak style="position:absolute; left:330; top:150; visibility:visible; z-index:150;">
					<table cellspacing="0" cellpadding="0" border="0" bgcolor="#5582D2">
						<tr>
							<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
							<td valign="top"><img src="../ill/tabel_top.gif" width="74" height="1" alt="" border="0"></td>
							<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
						</tr>
						<tr bgcolor="#5582D2">
							<td><a href="milepale.asp?menu=job&jid=<%=jid%>&rdir=<%=rdir%>&kundeid=<%=strKnr%>" class='alt'>Milepæle</a></td><!--Fakturering-->
						</tr>
						</table>
				</span>
				
				
				<span id=visgannt name=visgannt style="position:absolute; left:210; top:148; visibility:visible; z-index:100;">
					<table cellspacing="0" cellpadding="0" border="0" bgcolor="#5582D2">
						<tr>
							<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
							<td valign="top"><img src="../ill/tabel_top.gif" width="104" height="1" alt="" border="0"></td>
							<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
						</tr>
						<tr bgcolor="#5582D2">
							<td><a href="javascript:NewWin_large('jobplanner.asp?menu=job&id=<%=jid%>')" class='alt'>Gannt</a></td>
						</tr>
						</table>
				</span>
				
				
				
				<span id=visres name=visres style="position:absolute; left:120; top:148; visibility:visible; z-index:150;">
					<table cellspacing="0" cellpadding="0" border="0" bgcolor="#5582D2">
						<tr>
							<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
							<td valign="top"><img src="../ill/tabel_top.gif" width="74" height="1" alt="" border="0"></td>
							<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
						</tr>
						<tr bgcolor="#5582D2">
							<td><a href="javascript:NewWin_large('ressource_belaeg_jbpla.asp?id=<%=jid%>&showonejob=1')" class='alt'>Ressourcer</a></td>
						</tr>
						</table>
				</span>
				
				<span id=visstamdata name=visstamdata style="position:absolute; left:30; top:148; visibility:visible; z-index:250;">
					<table cellspacing="0" cellpadding="0" border="0" bgcolor="#5582D2">
						<tr>
							<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
							<td valign="top"><img src="../ill/tabel_top.gif" width="74" height="1" alt="" border="0"></td>
							<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
						</tr>
						<tr bgcolor="#5582D2">
							<td><a href="jobs.asp?menu=job&func=red&id=<%=jid%>&int=1&rdir=<%=rdir%>" class='alt'>Stamdata</a></td>
						</tr>
						</table>
				</span>
				
				<span id=vistpris name=vistpris style="position:absolute; left:510; top:148; visibility:visible; z-index:250;">
				<table cellspacing="0" cellpadding="0" border="0" bgcolor="#5582D2">
					<tr>
						<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
						<td valign="top"><img src="../ill/tabel_top.gif" width="74" height="1" alt="" border="0"></td>
						<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
					</tr>
					<tr bgcolor="#5582D2">
						<td><a href="jobs.asp?menu=job&func=red&id=<%=jid%>&int=1&rdir=<%=rdir%>&showdiv=tpriser" class='alt'>Timepriser</a></td>
					</tr>
					</table>
			</span>
			
			<span id=filearkiv name=filearkiv style="position:absolute; left:650; top:128; visibility:visible; z-index:0;">
			<table cellspacing="0" cellpadding="0" border="0" bgcolor="#5582D2">
				<tr>
					<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
					<td valign="top"><img src="../ill/tabel_top.gif" width="74" height="1" alt="" border="0"></td>
					<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
				</tr>
				<tr bgcolor="#5582D2">
					<td><a href="javascript:popUp('filer.asp?kundeid=<%=strKnr%>&jobid=<%=jid%>', '600', '600','200', '50')" target="_self" class='alt'>Filarkiv</a></td>
				</tr>
				</table>
		</span>
				
	<%
	end function
	
	
	function SQLBless2(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless2 = tmp
	end function
	
	'************************************************************************************************
	
	
	
	select case func
	case "slet"
	'*** Her spørges om det er ok at der slettes en medarbejder ***
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:210; top:100; background-color:#ffff99; visibility:visible; border:2px #000000 dashed; padding:15px;"">
	<table cellspacing="2" cellpadding="2" border="0" bgcolor="#ffff99" width=400>
	<tr>
	    <td><img src="../ill/alert_lille.gif" width="22" height="19" alt="" border="0">&nbsp;<b>Slet milepæl?</b><br><br>Du er ved at <b>slette</b> en milepæl. Er dette korrekt?</td>
	</tr>
	<tr>
	   <td><a href="milepale.asp?menu=job&func=sletok&id=<%=id%>&rdir=<%=rdir%>&jid=<%=jid%>">Ja</a>&nbsp;&nbsp;&nbsp;<a href="javascript:history.back()">Nej</a></td>
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
	oConn.execute("DELETE FROM milepale WHERE id = "& id &"")
	Response.redirect "milepale.asp?menu=job&shokselector=1&rdir="&rdir&"&jid="&jid
	
	case "dbopr", "dbred"
			
		if len(request("FM_navn")) = 0 then
		%>
		<!--#include file="../inc/regular/header_inc.asp"-->
		<%
		errortype = 61
		call showError(errortype)
		
		else
			
			'*tjekker om startdag eksisterer ** 
				if Request("FM_start_dag") > 28 then 
				select case Request("FM_start_mrd")
				case "2"
				strStartDay = 28
				case "4", "6", "9", "11"
				strStartDay = 30
				case else
				strStartDay = Request("FM_start_dag")
				end select
				else
				strStartDay = Request("FM_start_dag")
				end if
			
			strNavn = request("FM_navn")
			intType = request("FM_type")
			milepal_dato = Request("FM_start_aar") &"/" & Request("FM_start_mrd") & "/" & strStartDay 
			strBesk = SQLBless2(request("FM_besk"))
			jid = request("jid")
			timestamp = year(now) & "/" & month(now) &"/" & day(now)
			
			if func = "dbopr" then
			strSQL = "INSERT INTO milepale (navn, type, milepal_dato, beskrivelse, editor, dato, jid) VALUES ('"& strNavn &"', "& intType &", '"& milepal_dato &"', '"& strBesk &"', '"& session("user") &"', '"& timestamp &"', "& jid &")"
			oConn.execute(strSQL)
			
			else
			strSQL = "UPDATE milepale SET navn = '"& strNavn &"', type = "& intType &", milepal_dato = '"& milepal_dato &"', beskrivelse = '"& strBesk &"', editor = '"& session("user") &"', dato = '"& timestamp &"' WHERE id = "& id
			oConn.execute(strSQL)
			
			end if
			
			Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
			Response.Write("<script language=""JavaScript"">window.close();</script>")
			
		end if 'validering
	
	
	
	
	case "opr", "red" 
		
		
		if func = "opr" then
		jid = request("jid")
		func = "dbopr"
		subm = "opretpil"
		else
		
		strSQL = "SELECT milepale.id AS id, milepale.navn AS navn, milepal_dato, type, ikon, beskrivelse FROM milepale LEFT JOIN milepale_typer ON (milepale_typer.id = milepale.type) WHERE milepale.id = "& id 
		oRec.open strSQL, oConn, 3 
		if not oRec.EOF then
			strNavn = oRec("navn")
			strTdato = oRec("milepal_dato")
			strBesk = oRec("beskrivelse")
			intType = oRec("type")
		end if
		oRec.close
		
		jid = request("jid")
		func = "dbred"
		subm = "opdaterpil"
		end if
	%>
	
	<script>
	function NewWin_popupcal(url)    {
		window.open(url, 'Calpick', 'width=250,height=250,scrollbars=no,toolbar=no');    
		}
	</script>
	
	
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<%
	dtop = 20
	dleft = 10
	%>
	<!--#include file="inc/dato2.asp"-->
	
	<!-------------------------------Sideindhold------------------------------------->
	

	<div id="sindhold" style="position:absolute; left:<%=dleft%>; top:<%=dtop%>; visibility:visible; z-index:50;">
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF" width="600">
	<tr><form action="milepale.asp?menu=job&func=<%=func%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
	<input type="hidden" name="rdir" value="<%=rdir%>">
	<tr bgcolor="#5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td colspan=2 valign="top"><img src="../ill/tabel_top.gif" width="584" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td colspan=2 valign="top" class="alt" style="padding-top:4;">&nbsp;</td>
	</tr>
	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="40" alt="" border="0"></td>
		<td colspan=2 valign="top" style="padding-top:4;">
		</td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="40" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td valign="top" width=150 align=right style="padding-top:7; padding-right:4;"><b>Milepæl:</b></td>
		<td valign="top" style="padding-top:4;"><input type="text" name="FM_navn" id="FM_navn" value="<%=strNavn%>">&nbsp;(Overskrift)</td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	
	
	<tr bgcolor="#EFF3FF">
		<td valign="top" style="border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td>&nbsp;</td>
		<td valign="top" style="padding-top:4;"><b>Kunde og job:</b> (Kun aktive job, som du er tilknyttet via dine projektgrupper)</td>
		<td valign="top" style="border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#EFF3FF">
		<td valign="top" style="border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td>&nbsp;</td>
		<td valign="top" style="padding-top:4;">
			<select name="jid" id="jid" style="width:450px;">
			<%
			call hentbgrppamedarb(session("mid"))
			
			strSQL = "SELECT j.jobnavn, j.jobnr, j.id, j.jobknr, kkundenavn, kkundenr, kid FROM job j "_
			&" LEFT JOIN kunder ON (kid = j.jobknr) WHERE j.jobstatus = 1 "& strPgrpSQLkri &" ORDER BY kkundenavn, j.jobnavn"
			
			oRec.open strSQL, oConn, 3 
			while not oRec.EOF 
			if cint(jid) = oRec("id") then
			jsel = "SELECTED"
			else
			jsel = ""
			end if
			%>
			<option value="<%=oRec("id")%>" <%=jsel%>><%=oRec("kkundenavn")%> | <%=oRec("jobnavn")%> (<%=oRec("jobnr")%>) </option>
			<%
			oRec.movenext
			wend
			oRec.close %>
			</select>&nbsp;</td>
		<td valign="top" style="border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	
	
	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td align=right style="padding-top:4; padding-right:4;"><b>Dato:</b></td>
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
		<option value="31">31</option></select>&nbsp;
		
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
		<option value="<%=strAar%>">
		<%if id <> 0 then%>
		20<%=strAar%>
		<%else%>
		<%=strAar%>
		<%end if%></option>
		<option value="2002">2002</option>
		<option value="2003">2003</option>
	   	<option value="2004">2004</option>
	   	<option value="2005">2005</option>
		<option value="2006">2006</option>
		<option value="2007">2007</option></select>&nbsp;&nbsp;<a href="javascript:NewWin_popupcal('../inc/regular/popupcalender_inc.asp?use=1')"><img src="../ill/popupcal.gif" alt="" border="0" width="16" height="15"></a></td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		</tr>
	</tr>
	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td valign="top" align=right style="padding-top:4; padding-right:4;"><b>Type:</b></td>
		<td valign="top" style="padding-top:4;">
		<select name="FM_type">
		<%
		strSQL = "SELECT id, navn FROM milepale_typer ORDER BY navn"
		oRec.open strSQL, oConn, 3
		while not oRec.EOF
		if intType = oRec("id") then
		selthis = "SELECTED"
		else
		selthis = ""
		end if
		%>
		<option value="<%=oRec("id")%>" <%=selthis%>><%=oRec("navn")%></option>
		<%
		oRec.movenext
		wend
		oRec.close%>
		</select>&nbsp;(Sættes i kontrolpanelet)
		</td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="80" alt="" border="0"></td>
		<td valign="top" align=right style="padding-top:4; padding-right:4;"><b>Beskrivelse:</b></td>
		<td valign="top" style="padding-top:4;"><textarea cols="60" rows="4" name="FM_besk" id="FM_besk" id="FM_besk"><%=strBesk%></textarea></td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="80" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="50" alt="" border="0"></td>
		<td colspan=2 valign="top" align=center style="padding-top:4;">
		<input type="image" src="../ill/<%=subm%>.gif">
		</td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="50" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan=2 valign="bottom"><img src="../ill/tabel_top.gif" width="584" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>
	</form>
</table>
<br><br>
<!--
<a href="Javascript:history.back()" class=vmenu><img src="../ill/soeg-knap_tilbage.gif" width="16" height="16" alt="" border="0">&nbsp;Tilbage</a>
-->
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
	
	
	
	<%
	
	'**** henter job info ***************
		strSQL = "SELECT jobstartdato, jobslutdato, jobnavn FROM job WHERE id = "& jid
		oRec.open strSQL, oConn, 3 
		if not oRec.EOF then
		 	jstdato = oRec("jobstartdato")
			jenddato = oRec("jobslutdato")
			jobnavn = oRec("jobnavn")
		end if
		oRec.close
	
	call faneblade("mile", "")
	%>	
	<div id="sindhold" style="position:absolute; left:20; top:120; visibility:visible; z-index:50;">
<h3><img src="../ill/ac0063-24.gif" width="24" height="24" alt="" border="0">&nbsp;Jobdata</h3>
<br>
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF" width="600">
	<tr>
	<tr bgcolor="#5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="30" alt="" border="0"></td>
		<td colspan=4 valign="top"><img src="../ill/tabel_top.gif" width="586" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="30" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td class=alt width=200 style="padding-top:4; padding-bottom:2; padding-right:20;"><b><%=jobnavn%></b></td><td colspan=3 valign="top" class="alt" align=right style="padding-top:4; padding-bottom:2; padding-right:20;">
		<a href="javascript:popUp('milepale.asp?menu=job&func=opr&jid=<%=jid%>&rdir=<%=rdir%>','650','500','250','120');" target="_self" class=alt>Opret ny milepæl&nbsp;<img src="../ill/pillillexp.gif" width="16" height="18" alt="" border="0"></a></td>
	</tr>
	<tr bgcolor="#EFF3FF">
		<td valign="top" height=10 style="border-left:1px #003399 solid;">&nbsp;</td>
		<td colspan=4 valign="top" style="padding-top:4; padding-bottom:4;"><b>Milepæle</b><br>Dette kan være en deadline, en del-aflevering eller måske en del-fakturering.<br>Opret nye milepæle typer i kontrolpanelet.<br>&nbsp;</td>
		<td valign="top" align=right style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	
		<tr>
			<td colspan=6 bgcolor=#5582d2 height=1 style="border-right:1px #003399 solid; border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		</tr>
		<tr onmouseover="mOvr('gift',this,'#B4C7EF');" onmouseout="mOut(this,'');">
			<td valign="top" height=25 style="border-left:1px #003399 solid;">&nbsp;</td>
			<td valign="top" style="padding-right:4; padding-top:2;"><img src="../ill/mp_gron.gif" width="10" height="10" alt="" border="0"><img src="../ill/blank.gif" width="10" height="1" alt="" border="0"><font class=megetlillesilver><%=formatdatetime(jstdato, 1)%></font></td>
			<td valign="top" style="padding-top:2;"><b>Job startdato</b></td>
			<td valign="top" colspan=2>&nbsp;</td>
			<td valign="top" align=right style="border-right:1px #003399 solid;">&nbsp;</td>
		</tr>
		<tr>
			<td colspan=6 bgcolor="#5582d2" height=1 style="border-right:1px #003399 solid; border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		</tr>
		<%
		strSQL = "SELECT milepale.id AS id, milepale.navn AS navn, milepal_dato, milepale_typer.navn AS type, ikon, beskrivelse FROM milepale LEFT JOIN milepale_typer ON (milepale_typer.id = milepale.type) WHERE jid = "& jid &" ORDER BY milepal_dato"
		
		oRec.open strSQL, oConn, 3
		while not oRec.EOF 
		%>
		<tr onmouseover="mOvr('gift',this,'#B4C7EF');" onmouseout="mOut(this,'');">
			<td valign="top" style="border-left:1px #003399 solid;">&nbsp;</td>
			<td valign="top" width=120 style="padding-right:4; padding-top:2;"><img src="../ill/mp_<%=oRec("ikon")%>.gif" width="10" height="10" alt="" border="0"><img src="../ill/blank.gif" width="10" height="1" alt="" border="0"><font class=megetlillesilver><%=oRec("milepal_dato")%></font></td>
			<td valign="top" width=100 style="padding-top:2;">
			<%=oRec("type")%></td>
			<td valign="top" width=300 style="padding-left:4; padding-top:2; padding-bottom:3;">
			<a href="javascript:popUp('milepale.asp?func=red&id=<%=oRec("id")%>&jid=<%=jid%>&rdir=webblik','650','500','250','120');" target="_self" class=vmenu><%=oRec("navn")%></a>
			<%if len(oRec("beskrivelse")) <> 0 then%>
			<br>
			<font class=lillegray><%=oRec("beskrivelse")%></font>
			<%end if%></td>
			<td valign="top" width=100 style="padding-left:4; padding-top:2;"><a href="milepale.asp?menu=job&id=<%=oRec("id")%>&jid=<%=jid%>&func=slet&rdir=<%=rdir%>"><img src="../ill/slet_eks.gif" width="20" height="20" alt="" border="0"></a></td>
			<td valign="top" align=right style="border-right:1px #003399 solid;">&nbsp;</td>
		</tr>
		<tr>
			<td colspan=6 bgcolor=#5582d2 height=1 style="border-right:1px #003399 solid; border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		</tr>
		<%
		oRec.movenext
		wend
		oRec.close
		%>
		
		<tr onmouseover="mOvr('gift',this,'#B4C7EF');" onmouseout="mOut(this,'');">
			<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
			<td valign="top" height=25 style="padding-right:4; padding-top:2;"><img src="../ill/mp_end.gif" width="10" height="10" alt="" border="0"><img src="../ill/blank.gif" width="10" height="1" alt="" border="0"><font class=megetlillesilver><%=formatdatetime(jenddato, 1)%></font></td>
			<td valign="top" style="padding-top:2;"><b>Job slutdato</b></td>
			<td valign="top" colspan=2>&nbsp;</td>
			<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
		</tr>
		<tr>
			<td colspan=6 bgcolor=#5582d2 height=1 style="border-right:1px #003399 solid; border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		</tr>
	<tr bgcolor="#5582d2">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan=4 valign="bottom"><img src="../ill/tabel_top.gif" width="586" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>
	</table>
	<br><br>
	<a href="timereg.asp?menu=timereg" class=vmenu><img src="../ill/soeg-knap_tilbage.gif" width="16" height="16" alt="" border="0">&nbsp;Til timregistrering</a>
	<br><br>&nbsp;
	</div>
<%end select%>
<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
