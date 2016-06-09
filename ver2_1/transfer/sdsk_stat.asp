<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/helligdage_func.asp"-->
<!--#include file="inc/isint_func.asp"-->
<!--#include file="inc/sdsk_inc.asp"-->

<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else
	%>
	
	<%
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	func = request("func")
	
	thisfile = "sdsk"
	
	if func = "dbopr" then
	id = 0
	else
		if len(request("id")) <> 0 then
		id = request("id")
		else
		id = 0
		end if
	end if
	
	dato = day(now)&"/"&month(now)&"/"&year(now)
	tid = time
	
	
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%call sdskmainmenu(4)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	'call sdsktopmenu()
	%>
	</div>
	
	
	<div id="sideindhold" style="position:absolute; left:20; top:132; visibility:visible;">
	<h3><img src='../ill/ac0033-24.gif' width='24' height='24' alt='' border='0'> Servicedesk - Incident Statistik</h3>
	
			
	<%
	'* Kunde
	if len(request("FM_kontakt")) <> 0 AND request("FM_kontakt") <> 0 then
	kontaktKri = " kundeid = "& request("FM_kontakt")
	selKid = request("FM_kontakt")
	else
	selKid = 0
	kontaktKri = " kundeid <> 0"
	end if
	
	'* Ansavrlig
	if len(request("FM_ansv")) <> 0 then
		if request("FM_ansv") <> 0 then
		ansvKri = " AND s.ansvarlig = "& request("FM_ansv")
		ansSel = request("FM_ansv")
		else
		ansSel = 0
		ansvKri = ""
		end if
	else
	ansSel = 0
	ansvKri = " AND s.ansvarlig <> "& ansSel
	end if
	
	'* Dato
	
	call datocookie
	stDatoSQL = strAar &"/"& strMrd &"/"& strDag
	slDatoSQL = strAar_slut &"/"& strMrd_slut &"/"& strDag_slut
	
	if request("usedatokri") = "j" then
	datoKriJa = "CHECKED"
	datoKriNej = ""
	datoKri = " AND s.dato BETWEEN '"& stDatoSQL &"' AND '"& slDatoSQL &"'"
	else
	datoKri = ""
	datoKriJa = ""
	datoKriNej = "CHECKED"
	end if
	%>	
	<div style="position:relative; padding:10px; border-left:1px #8caae6 solid; border-top:1px #8caae6 solid; border-right:2px #5582d2 solid; border-bottom:2px #5582d2 solid; width:600; background-color:#ffffff;">
	<table cellspacing=0 cellpadding=2 border=0 width=600>
	<form action="sdsk_stat.asp" method="post">
	<tr bgcolor="#ffffff">
		<td><b>Vælg kontakt:</b></td><td colspan=2>
		<select name="FM_kontakt" style="font-family: arial; font-size: 10px; width:240px; height:12px;">
			<option value="0">Alle</option>
			<%
			
			ketypeKri = " ketype <> 'e'"
			
					strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE "& ketypeKri &" ORDER BY Kkundenavn"
					'Response.write strSQL
					'Response.flush
					oRec.open strSQL, oConn, 3
					while not oRec.EOF
					
					if cint(selKid) = cint(oRec("Kid")) then
					isSelected = "SELECTED"
					else
					isSelected = ""
					end if
					%>
					<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%>&nbsp;(<%=oRec("Kkundenr")%>)</option>
					<%
					oRec.movenext
					wend
					oRec.close
					%>
			</select>
		</td>
		</tr>
		<tr>
		<td><b>Ansvarlig (Intern):</b></td><td colspan=2><select name="FM_ansv" style="font-family: arial; font-size: 10px; width:240px; height:12px;">
							<option value="0">Alle</option>
							
							<%
							strSQL = "SELECT mid, mnavn FROM medarbejdere WHERE mid <> 0 AND mansat <> 2 ORDER BY mnavn"
							oRec.open strSQL, oConn, 3
							while not oRec.EOF
							if cint(ansSel) = cint(oRec("mid")) then
							isSelected = "SELECTED"
							else
							isSelected = ""
							end if%>
							<option value="<%=oRec("mid")%>" <%=isSelected%>><%=oRec("mnavn")%></option>
							<%
							oRec.movenext
							wend
							
							oRec.close
							%>
							</select>
	</td>
	</tr>
	<tr bgcolor="#ffffff">
		<td><b>Benyt datointerval:</b><br>
		<input type="radio" name="usedatokri" value="j" <%=datoKriJa%>>&nbsp;Ja&nbsp;&nbsp;
		<input type="radio" name="usedatokri" value="n" <%=datoKriNej%>>&nbsp;Nej</td>
	
		<!--#include file="inc/weekselector_b.asp"-->
	</tr>
	<tr bgcolor="#ffffff"><td colspan=3 align=right style="padding-right:50px;"><input type="submit" value="  Søg  "></td></tr>
	</form>
	</table>
	</div>
			
			
	
	
	<div style="position:relative; border-left:1px #8caae6 solid; border-top:1px #8caae6 solid; border-right:2px #5582d2 solid; border-bottom:2px #5582d2 solid; width:875px; top:20px; padding:10px; background-color:#ffffe1;">
	<table cellspacing=0 cellpadding=2 border=0 width=100% bgcolor="#ffffe1">
	<tr>
		<td valign=top><b>Antal fordelt på status:</b><br>
	<%
	
	intTotalall = 1
	
	antalTot = 0
	strSQL = "SELECT count(s.id) AS antal, st.navn AS navn "_
	&" FROM sdsk s"_
	&" LEFT JOIN sdsk_status st ON (st.id = s.status)"_
	&" WHERE s.id <> 0 AND "& kontaktKri &" "& ansvKri &" "& datoKri &" GROUP BY s.status ORDER BY antal DESC"
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
	
	Response.write oRec("navn") & ": <b>" & oRec("antal") &"</b><br>"
	
	antalTot = antalTot + cint(oRec("antal"))
	oRec.movenext
	wend
	
	oRec.close
	%>
	</td>
	
	<td valign=top><b>Antal fordelt på type:</b><br> (Top 10 med flest incidents)<br>
	<%
	strSQL = "SELECT count(s.id) AS antal, tp.navn AS navn "_
	&" FROM sdsk s"_
	&" LEFT JOIN sdsk_typer tp ON (tp.id = s.type)"_
	&" WHERE s.id <> 0 AND "& kontaktKri &" "& ansvKri &" "& datoKri &" GROUP BY s.type ORDER BY antal DESC"
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
	
	Response.write oRec("navn") & ": <b>" & oRec("antal") &"</b><br>"
	
	oRec.movenext
	wend
	
	oRec.close
	%></td>
	
	
	<td valign=top><b>Antal fordelt på prioiteter:</b><br>
	<%
	strSQL = "SELECT count(s.id) AS antal, p.navn AS navn "_
	&" FROM sdsk s"_
	&" LEFT JOIN sdsk_prioitet p ON (p.id = s.prioitet)"_
	&" WHERE s.id <> 0 AND "& kontaktKri &" "& ansvKri &" "& datoKri &" GROUP BY s.prioitet ORDER BY antal DESC"
	
	'Response.write strSQL 
	'Response.flush
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
	
	Response.write oRec("navn") & ": <b>" & oRec("antal") &"</b><br>"
	
	oRec.movenext
	wend
	
	oRec.close
	%></td>
	
	<td valign=top><b>Antal uden tildelt ansvarlig:</b><br>
	<%
	strSQL = "SELECT count(s.id) AS antal "_
	&" FROM sdsk s"_
	&" WHERE s.ansvarlig = 0 AND "& kontaktKri &" "& datoKri 
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
	
	Response.write "Der er ialt: <b>" & oRec("antal") &"</b> incidents u. ansvarlige.<br>"
	
	oRec.movenext
	wend
	
	oRec.close
	%></td>
	</tr>
	</table>
	</div>
	
	
	<%
	blnIE4 = False
strUA = Request.ServerVariables("HTTP_USER_AGENT")
intMSIE = Instr(strUA, "MSIE ") + 5
If intMSIE > 5 Then
  If CInt(Mid(strUA, intMSIE, Instr(intMSIE, strUA, ".") - intMSIE)) >= 4 Then 
    blnIE4 = True
  End If 
End If


If blnIE4 Then




'*** create the pie chart ***
intTotalall = antalTot
iLineCnt = 3
iDegPos = 0
iRectTop = -110



%>
							<%
							'************************************************************************
							'** Påbegyndt
							'************************************************************************
							%>


				
							<DIV STYLE="font-family:Tahoma,Verdana,Arial,sans-serif; position:absolute; top:346; left:450; width:400; background-color:#ffffff;  border-top:1px #8caae6 solid; border-left:1px #8caae6 solid; border-bottom:2px #5582d2 solid; border-right:2px #5582d2 solid; padding:10px;">
							<b>Antal påbegyndt til tiden / Antal påbegyndt forsent.</b>
							<table cellspacing=0 cellpadding=2 border=0 width=100%>
							<tr>
							<td valign=top><br><u>Antal påbegyndt forsent:</u><br>
							<%
							antalOvertiden = 0
							strSQL = "SELECT count(s.id) AS antal, p.responstid, p.navn AS navn "_
							&" FROM sdsk s"_
							&" LEFT JOIN sdsk_prioitet p ON (p.id = s.prioitet)"_
							&" WHERE s.id <> 0 AND "& kontaktKri &" "& ansvKri &" "& datoKri &" AND (p.responstid < s.rsptid_a_tid AND s.rsptid_a_id <> 0) GROUP BY s.prioitet ORDER BY antal DESC"
							
							'Response.write strSQL 
							'Response.flush
							oRec.open strSQL, oConn, 3
							while not oRec.EOF
							
							Response.write oRec("navn") & ": <b>" & oRec("antal") &"</b><br>"
							antalOvertiden = antalOvertiden + cint(oRec("antal"))
							oRec.movenext
							wend
							
							oRec.close
							%></td>
							
							<td valign=top><br><u>Antal påbegyndt. til tiden:</u><br>
							<%
							antalTiltiden = 0
							strSQL = "SELECT count(s.id) AS antal, p.responstid, p.navn AS navn "_
							&" FROM sdsk s"_
							&" LEFT JOIN sdsk_prioitet p ON (p.id = s.prioitet)"_
							&" WHERE s.id <> 0 AND "& kontaktKri &" "& ansvKri &" "& datoKri &" AND (p.responstid >= s.rsptid_a_tid AND s.rsptid_a_id <> 0) GROUP BY s.prioitet ORDER BY antal DESC"
							
							'Response.write strSQL 
							'Response.flush
							oRec.open strSQL, oConn, 3
							while not oRec.EOF
							
							Response.write oRec("navn") & ": <b>" & oRec("antal") &"</b><br>"
							antalTiltiden = antalTiltiden + cint(oRec("antal")) 
							oRec.movenext
							wend
							
							oRec.close
							%>
							
							</td>
							</tr>
							</table>
							
				
				
				<%
				if antalTiltiden <> 0 AND intTotalall <> 0 then
				tiltidenPieStr = CStr(Round(antalTiltiden / intTotalall * 360))
				tiltidenPro = FormatPercent(antalTiltiden / intTotalall, 1) 
				else
				tiltidenPieStr = 0
				tiltidenPro = FormatPercent(0, 1) 
				end if
				
				if antalOvertiden <> 0 AND intTotalall <> 0 then
				overtidenPieStr = CStr(Round(antalOvertiden / intTotalall * 360))
				overtidenPro = FormatPercent(antalOvertiden / intTotalall, 1) 
				else
				overtidenPieStr = 0
				overtidenPro = FormatPercent(0, 1) 
				end if
				%>

				<OBJECT ID="sgPie" STYLE="position:relative; height:230; width:300; top:20; left:0" CLASSID="CLSID:369303C2-D7AC-11D0-89D5-00A0C90833E6">
				
				<PARAM NAME="Line0001" VALUE="SetLineColor(0, 0, 0)"> 
				<PARAM NAME="Line0002" VALUE="SetLineStyle(0)">
				<PARAM NAME="Line0003" VALUE="SetFillColor(192, 192, 192">
				<PARAM NAME="Line0004" VALUE="Oval(-135, -95, 75, 75, 0)">
				
				
				<!--<param name='Line0005' value='Rect(100, <= CStr(iRectTop) + 5 %>, 20, 15, 0)'>-->
				<param name='Line0005' value='SetLineStyle(1)'>
				<PARAM NAME='Line0006' VALUE='SetFillColor(253, 153, 0)'>
				<PARAM NAME='Line0007' VALUE='Pie(-140, -100, 75, 75, <%= CStr(iDegPos) %>, <%= tiltidenPieStr %> , 0)'>
				<PARAM NAME='Line0008' VALUE='Rect(15, <%= CStr(iRectTop) %>, 20, 15, 0)'>
				
				<%
				iDegPos = iDegPos + tiltidenPieStr
				%>
				
				<!--<param name='Line0010' value='Rect(100, <= CStr(iRectTop) + 25 %>, 20, 15, 0)'>-->
				<param name='Line0009' value='SetLineStyle(1)'>
				<PARAM NAME='Line0010' VALUE='SetFillColor(255, 55, 0)'>
				<PARAM NAME='Line0011' VALUE='Pie(-140, -100, 75, 75, <%= CStr(iDegPos) %>, <%= overtidenPieStr %> , 0)'>
				<PARAM NAME='Line0012' VALUE='Rect(15, <%= CStr(iRectTop) + 20 %>, 20, 15, 0)'>
				
				
				</OBJECT>	
				
			
				
				<DIV STYLE="font-family:Tahoma,Verdana,Arial,sans-serif; position:relative; top:-204; left:190; width:400;">
				<B>Påbegyndt til tiden.</B>  &nbsp; Count:  &nbsp;(<B><%= tiltidenPro %></B>) 
				</DIV>
				
				<DIV STYLE="font-family:Tahoma,Verdana,Arial,sans-serif; position:relative; top:-199; left:190; width:400;">
				<B>Påbegyndt forsent.</B>  &nbsp; Count:  &nbsp;(<B><%= overtidenPro %></B>)
				</DIV>
				
				
				</div>
				
				
				
				<%
				'************************************************************************
				'** Afsluttet 
				'************************************************************************
				%>
				
				
				
				
					<DIV STYLE="font-family:Tahoma,Verdana,Arial,sans-serif; position:absolute; top:346; left:0; width:400; background-color:#ffffff; border-top:1px #8caae6 solid; border-left:1px #8caae6 solid; border-bottom:2px #5582d2 solid; border-right:2px #5582d2 solid; padding:10px;">
							<b>Antal incidents afsluttet forsent / til tiden.</b>
							<table cellspacing=0 cellpadding=2 border=0 width=100%>
							<tr>
								<td valign=top><br><u>Antal afsl. forsent:</u><br>
								<%
								afslForsent = 0
								strSQL = "SELECT count(s.id) AS antal, p.responstid, p.navn AS navn "_
								&" FROM sdsk s"_
								&" LEFT JOIN sdsk_prioitet p ON (p.id = s.prioitet)"_
								&" WHERE s.id <> 0 AND "& kontaktKri &" "& ansvKri &" "& datoKri &" AND (p.responstid < s.rsptid_b_tid AND s.rsptid_b_id <> 0) GROUP BY s.prioitet ORDER BY antal DESC"
								
								'Response.write strSQL 
								'Response.flush
								oRec.open strSQL, oConn, 3
								while not oRec.EOF
								
								Response.write oRec("navn") & ": <b>" & oRec("antal") &"</b><br>"
								afslForsent = afslForsent + cint(oRec("antal"))
								oRec.movenext
								wend
								
								oRec.close
								%></td>
								
								<td valign=top><br><u>Antal afsl. til tiden:</u><br>
								<%
								afslTiltiden = 0
								strSQL = "SELECT count(s.id) AS antal, p.responstid, p.navn AS navn "_
								&" FROM sdsk s"_
								&" LEFT JOIN sdsk_prioitet p ON (p.id = s.prioitet)"_
								&" WHERE s.id <> 0 AND "& kontaktKri &" "& ansvKri &" "& datoKri &" AND (p.responstid >= s.rsptid_b_tid AND s.rsptid_b_id <> 0) GROUP BY s.prioitet ORDER BY antal DESC"
								
								'Response.write strSQL 
								'Response.flush
								oRec.open strSQL, oConn, 3
								while not oRec.EOF
								
								Response.write oRec("navn") & ": <b>" & oRec("antal") &"</b><br>"
								afslTiltiden = afslTiltiden + cint(oRec("antal"))
								oRec.movenext
								wend
								
								oRec.close
								%></td>
								
								</tr>
							</table>
							
				
				
				<%
				if afslTiltiden <> 0 AND intTotalall <> 0 then
				afslTiltidenPieStr = CStr(Round(afslTiltiden / intTotalall * 360))
				afslTiltidenPro = FormatPercent(afslTiltiden / intTotalall, 1) 
				else
				afslTiltidenPieStr = 0
				afslTiltidenPro = FormatPercent(0, 1) 
				end if
				
				if afslForsent <> 0 AND intTotalall <> 0 then
				afslForsentPieStr = CStr(Round(afslForsent / intTotalall * 360))
				afslForsentPro = FormatPercent(afslForsent / intTotalall, 1) 
				else
				afslForsentPieStr = 0
				afslForsentPro = FormatPercent(0, 1) 
				end if
				%>

				<OBJECT ID="sgPie" STYLE="position:relative; height:230; width:300; top:20; left:0" CLASSID="CLSID:369303C2-D7AC-11D0-89D5-00A0C90833E6">
				
				<PARAM NAME="Line0001" VALUE="SetLineColor(0, 0, 0)"> 
				<PARAM NAME="Line0002" VALUE="SetLineStyle(0)">
				<PARAM NAME="Line0003" VALUE="SetFillColor(192, 192, 192">
				<PARAM NAME="Line0004" VALUE="Oval(-135, -95, 75, 75, 0)">
				
				
				<!--<param name='Line0005' value='Rect(100, <= CStr(iRectTop) + 5 %>, 20, 15, 0)'>-->
				<param name='Line0005' value='SetLineStyle(1)'>
				<PARAM NAME='Line0006' VALUE='SetFillColor(85, 233, 66)'>
				<PARAM NAME='Line0007' VALUE='Pie(-140, -100, 75, 75, <%= CStr(iDegPos) %>, <%= afslTiltidenPieStr %> , 0)'>
				<PARAM NAME='Line0008' VALUE='Rect(15, <%= CStr(iRectTop) %>, 20, 15, 0)'>
				
				<%
				iDegPos = iDegPos + afslTiltidenPieStr
				%>
				
				<!--<param name='Line0010' value='Rect(100, <= CStr(iRectTop) + 25 %>, 20, 15, 0)'>-->
				<param name='Line0009' value='SetLineStyle(1)'>
				<PARAM NAME='Line0010' VALUE='SetFillColor(233, 66, 93)'>
				<PARAM NAME='Line0011' VALUE='Pie(-140, -100, 75, 75, <%= CStr(iDegPos) %>, <%= overtidenPieStr %> , 0)'>
				<PARAM NAME='Line0012' VALUE='Rect(15, <%= CStr(iRectTop) + 20 %>, 20, 15, 0)'>
				
				
				</OBJECT>	
				
			
				
				<DIV STYLE="font-family:Tahoma,Verdana,Arial,sans-serif; position:relative; top:-204; left:190; width:400;">
				<B>Afsluttet til tiden.</B>  &nbsp; Count:  &nbsp;(<B><%= afslTiltidenPro %></B>) 
				</DIV>
				
				<DIV STYLE="font-family:Tahoma,Verdana,Arial,sans-serif; position:relative; top:-199; left:190; width:400;">
				<B>Afsluttet forsent.</B>  &nbsp; Count:  &nbsp;(<B><%= afslForsentPro %></B>)
				</DIV>
				
				
				</div>
				
				
<%end if%>


	
	
</div>
<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
	
	
