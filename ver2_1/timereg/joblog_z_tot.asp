<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<%
if session("user") = "" then
%>
<!--#include file="../inc/regular/header_inc.asp"-->
<%
	errortype = 5
	call showError(errortype)
	else
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->	
	<%if request("print") <> "j" then%>
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!--#include file="../inc/regular/rmenu.asp"-->
	<%end if%>
	<!--#include file="inc/convertDate.asp"-->
	<script Language="JavaScript">
    function NewWin(url)    {
    window.open(url, 'print', 'width=620,height=680,scrollbars=yes,toolbar=no');    
	}
	</script>
	<script language="javascript">
	function showweek(weeknum,antalweeks) {
	for (i=0;i<antalweeks;i++){
	if(i == weeknum){
	document.all["weekdiv_"+weeknum+""].style.visibility = "visible";
	document.all["knap_"+i+""].style.borderColor = "darkRed";
	}
	else{
	document.all["weekdiv_"+i+""].style.visibility = "hidden";
	document.all["knap_"+i+""].style.borderColor = "lightBlue";
	}
	}
	}
	</script> 
	<script language="javascript">
	<!--
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	//-->
	</script>
	<%
	Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"
	Response.write time
	
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
		end function
	
	
	
	function opbygning()
	'*** Opbygning af oversigt medarbejdere *****
			areMaP = instr(strMedarbPrinted, ", " &oRec("Tmnr")&"#")
			
			if areMaP = "0" then
			Redim preserve medarbejder(t)
			Redim Preserve medarbID(t)
			medarbID(t) = oRec("Tmnr")  
			medarbejder(t) = oRec("Tmnavn")
			Redim preserve medarbTtot(t) 
			Redim preserve medarbOmstot(t)
			t = t + 1
			end if
			
			strMedarbPrinted = strMedarbPrinted & ", " & oRec("Tmnr") &"#" 
			
			'*** Opbygning af oversigt Aktiviteter *****
			areAktP = instr(strAktPrinted, ", " &oRec("TAktivitetId")&"#")
			
			if areAktP = "0" then
			Redim preserve aktivitetID(a)
			Redim preserve aktiviteter(a)
			Redim preserve jobnavn(a)
			Redim preserve jobnr(a)
			aktivitetID(a) = oRec("TAktivitetId")
			aktiviteter(a) = oRec("Anavn")
			jobnavn(a) = oRec("Tjobnavn")
			jobnr(a) = oRec("Tjobnr")
			Redim preserve timerAktTper(a)
			Redim preserve timerAktOmsper(a)
			a = a + 1
			end if
			
			strAktPrinted = strAktPrinted & ", " & oRec("TAktivitetId") &"#" 
			
			'*** Opbygning af weeknumbers *****
			areWeekP = instr(strWeekNumPrinted, ", " &strWeekNum&"#")
			if areWeekP = "0" then
			Redim preserve weeknumbers(w)
			weeknumbers(w) = strWeekNum
			w = w + 1
			end if
			
			
			'Response.write strYearNum &"-" & strWeekNum &"&nbsp;&nbsp;:"
			'*** Opbygning af yearnumbers *****
			areYearP = instr(strYearNumPrinted, ", " & strYearNum &"#")
			if areYearP = "0" then
			y = y + 1
			strWeekNumPrinted = "#0"
			yearweeks(y,week) = strWeekNum
			week = week + 1
			else
				if areWeekP = "0" then
				yearweeks(y,week) = strWeekNum
				week = week + 1
				end if
			end if
			
			
			strWeekNumPrinted = strWeekNumPrinted & ", " & strWeekNum &"#" 
			strYearNumPrinted = strYearNumPrinted & ", " & strYearNum &"#" 
			'**************  opbygning slut ********************
	end function
	
	
	function AktprMedTot()
	strSQL = "SELECT timer AS Ztimer, tdato, TAktivitetId, TimePris, Tfaktim FROM timer WHERE TAktivitetId = "& aktivitetID(a) &" AND Tmnr = "& medarbID(t) &""
				oRec2.open strSQL, oConn, 3
				while not oRec2.EOF
					if datepart("ww", oRec2("tdato"),2,2) => weeknumbers(w-1) AND datepart("ww", oRec2("tdato"),2,2) <= weeknumbers(0) then 
						intZtimer = intZtimer + csng(oRec2("Ztimer"))
						if oRec2("Tfaktim") = "1" then
						omsAktMed = omsAktMed + (csng(oRec2("Ztimer")) * ccur(oRec2("TimePris")))
						else
						omsAktMed = 0
						end if
					end if
				oRec2.movenext
				wend
				oRec2.close
			Response.write SQLBless(intZtimer) 
			intZtimer = 0
	end function
	
	'***************          Start sideopbygning og variable          ************
	'** S�tter Periode v�rdier **
	FM_medarb = request("FM_medarb")
	FM_job = request("FM_job")
	FM_Aar = request("FM_Aar")
	eks = request("eks")
	func = request("func")
	weekselector = request("weekselector")
	
	if weekselector = "j" then 
	StrTdato = request("FM_start_mrd") &"/" & request("FM_start_dag") & "/" & request("FM_start_aar")
	StrUdato = request("FM_slut_mrd") &"/" & request("FM_slut_dag") & "/" & request("FM_slut_aar")
	strWSelStartDato = cdate(StrTdato)
	strWSelEndDato = cdate(StrUdato) 
	id = "1"
		'if request("FM_start_aar") <> request("FM_slut_aar") then
		strReqAar_slut = request("FM_slut_aar")
		'end if
	else
	StrTdato = (month(now()) - 1) &"/1/" & year(now())
	StrUdato = month(now()) &"/" & day(now()) & "/" & year(now())
	id = "1"
	strReqAar_slut = 2025
	end if
	'**** 
	
	linket = request("linket")
	
	selmedarb = FM_medarb 'request("selmedarb")
	selaktid = request("selaktid")
	intJobnr = FM_job 'request("jobnr")
	
	if intJobnr = "0" then
	intJobnrKri = ""
	else
		Dim intJobnrKriValues 
		Dim i
		intJobnrKriValues = Split(intJobnr, ", ")
		For i = 0 to Ubound(intJobnrKriValues)
		'Response.write intJobnrKriValues(i)
		intJobnrKri = intJobnrKri & "Tjobnr = '" & intJobnrKriValues(i) &"' OR "
		next
		intJobnrKriCount = len(intJobnrKri)
		trimintJobnrKri = left(intJobnrKri, (intJobnrKriCount-3))
		intJobnrKri = trimintJobnrKri & " AND "
	end if
	
	if len(Request("mrd")) <> "0" then
		strReqMrd = Request("mrd")
	else
	strReqMrd = month(now)
	end if
	
	'**** finder ud af om der er valgt et �r ***
	if weekselector <> "j" then 
		if len(Request("year")) <> "0" then
			if Request("year") = "-1" then
			strReqAar = "0"
			else
			strReqAar = Request("year")
			end if	
		else
		strReqAar = right(year(now), 2)
		end if
		strReqAar_slut = strReqAar
	else
	strReqAar = request("FM_start_aar")
	end if	
	
	lastFakdag = Request("lastFakdag")
	
	'*** Bruges som mark�r af mrd knapper **
	if weekselector = "j" then
	framrd = datepart("m", StrTdato)
	tilmrd = datepart("m", StrUdato)
	end if


if request("print") <> "j" then%>
	<!--#include file="inc/joblog_z_mrdk.asp"-->
	<%
end if 'print

'*** Finder de medarbejdere der er valgt ***
if len(selmedarb) <> "0" then
	if selmedarb > "0" then
	'selmedarbKri = "Tmnr = " & selmedarb &" AND "
		Dim intMedArbVal 
		Dim b
		intMedArbVal = Split(selmedarb, ", ")
		For b = 0 to Ubound(intMedArbVal)
		selmedarbKri = selmedarbKri & "Tmnr = " & intMedArbVal(b) &" OR "
		next
		intMedArbKriCount = len(selmedarbKri)
		trimintMedArbVal = left(selmedarbKri, (intMedArbKriCount-3))
		selmedarbKri = trimintMedArbVal & " AND "
	else
	selmedarbKri = "Tmnr <> 0 AND "
	end if
else
selmedarbKri = "Tmnr <> 0 AND "
end if

Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"
Response.write selmedarbKri

'*** Finder de aktiviteter der er valgt ***
if len(request("selaktid")) <> "0" then
selaktidKri = "AND TAktivitetId = " & selaktid &""
else
selaktidKri = ""
end if

if request("print") <> "j" then
visWeekSel = "visible"
else
visWeekSel = "hidden"
end if%>	
<!--#include file="inc/weekselector.asp"-->
<%
StrTdato = "1/1/2001"
StrUdato = "12/12/2007"
	
	'** Fakturerbart / Ikke fakturerbart job. **
	strSQL = "SELECT Tdato, TasteDato, Tjobnr, Tjobnavn, TAktivitetNavn AS Anavn, TAktivitetId, Tknavn, Tmnr, Tmnavn, Timer, Tid, Tfaktim, TimePris, Timerkom FROM timer WHERE "& selmedarbKri & intJobnrKri &" Tid > 0 "& selaktidKri &" ORDER BY Tdato DESC, Tmnavn"
	oRec.open strSQL, oConn, 0, 1
	
	strYearNumPrinted = "0#"
	strWeekNumPrinted = "0#"
	strMedarbPrinted = "0#"
	strAktPrinted = "0#"
	a = 0
	t = 0
	w = 0
	v = 0
	y = -1
	week = 0
	
	Redim medarbID(t)
	Redim medarbejder(t)
	Redim medarbOmstot(t)
	Redim medarbTtot(t)
	Redim aktivitetID(a)
	Redim aktiviteter(a)
	Redim jobnavn(a)
	Redim jobnr(a)
	Redim weeknumbers(w)
	'Redim yearnumbers(y)
	Dim yearweeks(2,52)
	Redim timerAktTper(a)
	Redim timerAktOmsper(a)
	
	while not oRec.EOF
		StrTdato = oRec("Tdato")
		strWeekNum = datepart("ww", StrTdato,2,2)
		strYearNum = datepart("yyyy", StrTdato)
		id = 1%>
		<!--#include file="inc/dato2.asp"-->
		<%
		if cint(strReqAar) = cint(strReqAar_slut) then
			if cint(strAar) = cint(strReqAar) OR cint(strReqAar) = "0" then
			if cint(strReqMrd) = cint(strMrd) OR cint(strReqMrd) = "0" then
			if datepart("ww", oRec("tdato"),2,2) => frauge AND datepart("ww", oRec("tdato"),2,2) <= tiluge then 
			
			call opbygning()
			
			end if
			end if
			end if
		else
			if cDate(StrTdato) => strWSelStartDato AND cDate(StrTdato) <= strWSelEndDato then
				call opbygning()
			end if
		end if
	
	
	oRec.movenext
	wend	  
		
	oRec.Close
	Set oRec = Nothing
	
'*************************       Z        *****************************************************
%>
<!-- venstre side div -->
<%if request("print") <> "j" then
leftloading = "280"
else
leftloading = "140"
end if
%>
<div id="loading" name="loading" style="position:absolute; top:300; left:<%=leftloading%>; width:300; height:50; z-index:1000; visibility:visible; background-color:#ffffff; !border: 1px; border-color: DarkRed; border-style: solid; padding-left:4px; padding-top:5px;">
<b>Henter information...</b><br>
Denne handling kan tage op til 20 sek. alt efter din forbindelse.
</div>
<%if request("print") <> "j" then%>		
<div id="z" style="position:absolute; left:180; top:55; width:300; height:600; visibility:visible; z-index:100;">
<div id="logknap" style="position:absolute; top:35; left:304; width:32; height:30px; z-index:100; !border: 1px; background-color:#ffffff; border-color: lightBlue; border-style: solid; padding-left:4px; padding-top:3px;">
<a href="word.asp?eks=<%=request("eks")%>&jobnr=<%=intJobnr%>&mrd=<%=strReqMrd%>&year=<%=strReqAar%>"><img src="../ill/word.gif" width="23" height="22" alt="" border="0"></a>
</div>
<div style="position:absolute; top:35; left:338; width:90; height:30px; z-index:100; !border: 1px; background-color:#ffffff; border-color: lightBlue; border-style: solid; padding-left:4px; padding-top:5px;">
<a href="stat.asp?menu=stat&func=sel&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>"><img src="../ill/pil_lysblaa_left.gif" width="11" height="6" alt="" border="0">&nbsp;Statistik</a>
</div>
<div style="position:absolute; top:35; left:430; width:33; height:30px; z-index:100; !border: 1px; background-color:#ffffff; border-color: lightBlue; border-style: solid; padding-left:4px; padding-top:3px;">
	<a href="joblog.asp?menu=stat&mrd=<%=strReqMrd%>&jobnr=<%=intJobnr%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>"><img src="../ill/joblog.gif" alt="Joblog" border="0" width="23" height="22"></a>
</div>
<div style="position:absolute; top:35; left:465; width:33; height:30px; z-index:100; !border: 1px; background-color:#ffffff; border-color: lightBlue; border-style: solid; padding-left:4px; padding-top:3px;">
	<a href="joblog_z.asp?menu=stat&mrd=<%=strReqMrd%>&jobnr=<%=intJobnr%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>"><img src="../ill/sum.gif" alt="Joblog" border="0" width="23" height="22"></a>
</div>
<div style="position:absolute; top:35; left:500; width:32; height:30px; z-index:100; !border: 1px; background-color:#ffffff; border-color: lightBlue; border-style: solid; padding-left:4px; padding-top:3px;">
<a href="javascript:NewWin('joblog_z.asp?menu=stat&print=j&mrd=<%=strReqMrd%>&jobnr=<%=intJobnr%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>')" target="_self"><img src="../ill/small_print.gif" width="23" height="22" alt="" border="0"></a>
</div>
<%else%>
<div id="z" style="position:absolute; left:20; top:15; width:300; height:600; visibility:visible; z-index:100;">
<div style="position:absolute; top:35; left:466; width:32; height:30px; z-index:100; !border: 1px; background-color:#ffffff; border-color: lightBlue; border-style: solid; padding-left:4px; padding-top:3px;">
<a href="javascript:window.close()"><img src="../ill/luk.gif" width="23" height="22" alt="" border="0"></a>
</div>
<div style="position:absolute; top:35; left:500; width:32; height:30px; z-index:100; !border: 1px; background-color:#ffffff; border-color: lightBlue; border-style: solid; padding-left:4px; padding-top:3px;">
<a href="javascript:window.print()"><img src="../ill/small_print.gif" width="23" height="22" alt="" border="0"></a>
</div>
<%end if%>
<table cellspacing="0" cellpadding="0" border="0" width="600">
<tr><td><br><br><font class="pageheader">Timefordeling</font><br>
<%if weekselector = "j" then
%>
<div id="periodeShow" style="position:absolute; left:155; top:106; visibility:visible; z-index:100;">
<table width="200" celpadding="2" cellspacing="1" border="0">
<tr><td height="13" align="left" valign="top" style="!border: 1 px; border-color: DarkRed; border-style: solid; padding-left : 4px; padding-right : 4px;">
<font size="1"><%
Response.write showStrTdato &"&nbsp;(uge&nbsp;"& frauge &")&nbsp;" &" - " & showStrUdato &"&nbsp;(uge&nbsp;"& tiluge &")&nbsp;"
%>
</td></tr></table></div>
<%
end if

if request("print") = "j" then
Response.write "20" & strReqAar 
end if%>
<br></td></tr></table>


<!-------------------------------Sideindhold------------------------------------->
<%
'** Hvis der findes uger med registreringer i den valgte periode.**
if week > 0 then
'*** Periode total ****
%>
<!-- H�jre side div -->
<%if request("print") <> "j" then%>	
<div id="akttotal" style="position:absolute; left:450; top:188; visibility:visible; z-index:200;">
<%else%>
<div id="akttotal" style="position:absolute; left:300; top:120; visibility:visible; z-index:200;">
<%end if%>
<table cellspacing="1" cellpadding="0" border="0" width="261">
<tr>
	<td colspan="2"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	<td align="right" style="!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px;  padding-right : 4px;"><font size="1">timer</td>
	<td><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	<td align="right" style="!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px;  padding-right : 4px;"><font size="1">oms�tning</td>
</tr>
<tr>
	<td width="115" style="!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px; padding-right : 1px;"><font size="1">uge: <b><%=weeknumbers(w-1)&" - "& weeknumbers(0)%></b></td>
	<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	<td width="59" align="right" style="!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px;  padding-right : 4px;">
	<b><%
	for y = 0 To y - 1
	for week = 0 To week - 1
	for a = 0 To a - 1
	for t = 0 To t - 1
				strSQL = "SELECT timer AS Ztimer, tdato, Tfaktim, TAktivitetId, TimePris FROM timer WHERE Tmnr = "&medarbID(t)&" AND TAktivitetId = "& aktivitetID(a) &""
				oRec2.open strSQL, oConn, 0, 1  'adOpenForwardOnly, adLockReadOnly
				
				while not oRec2.EOF
				'Response.write datepart("ww", oRec2("tdato"),2,2) &"="  & yearweeks(y,week) &"<br>"
				if datepart("ww", oRec2("tdato"),2,2) = yearweeks(y,week) then 'weeknumbers(w) then  ''AND datepart("yyyy", oRec2("tdato")) = yearnumbers(0) then 
					
					'** S�tter total timeforbrug pr aktivitet i den valgte periode ***
					timerAktTper(a) = timerAktTper(a) + csng(oRec2("Ztimer"))
					medarbTtot(t) = medarbTtot(t) + csng(oRec2("Ztimer"))
					
					if oRec2("Tfaktim") = "1" then
					'** S�tter total timeforbrug p� en aktivitet i den valgte periode ***
					timerAktOmsper(a) = timerAktOmsper(a) + (csng(oRec2("Ztimer")) * ccur(oRec2("TimePris")))
					medarbOmstot(t) = medarbOmstot(t) + (csng(oRec2("Ztimer")) * ccur(oRec2("TimePris")))
					else
					timerAktOmsper(a) = 0
					medarbOmstot(t) = medarbOmstot(t) 
					end if 
				end if
				
				oRec2.movenext
				wend
				oRec2.close
	next
	  timerPerTot = timerPerTot + timerAktTper(a)
	  omsTotPer = omsTotPer + timerAktOmsper(a)
	next
	Response.write SQLBless(timerPerTot)
	timerPerTot = 0
	next
	next
	
	%></b></td>
	<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	<td align="right" width="85" style="!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px;  padding-right : 2px;"><b><%=SQLBless(ccur(omsTotPer))%></b></td>
	</tr>
<tr>
	<td height=20 colspan="5"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td></tr>
</table>
<%omsTotPer = 0

'*** Aktiviteter total ****
%>
<table cellspacing="0" cellpadding="0" border="0" width="261">	
<tr>
	<td colspan="2" style="!border: 1 px; background-color: LightBlue; border-color: silver; border-style: solid; padding-left : 4px;">Aktiviteter:</td>
	<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	<td align="right" style="!border: 1 px; background-color: LightBlue; border-color: silver; border-style: solid; padding-left : 4px;"><font size=1>timer&nbsp;</td>
	<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	<td align="right" style="!border: 1 px; background-color: LightBlue; border-color: silver; border-style: solid; padding-left : 4px;"><font size=1>oms�tning&nbsp;</td></tr>
<%
For a = 0 to a - 1
	%>
	<tr><td height=2 colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td></tr>
			<tr><td colspan="2" width="115" style="!border: 1 px; background-color: #ffffff; border-color: LightBlue; border-style: solid; padding-left : 4px;"><font size="1"><%=jobnr(a) &"&nbsp;" & jobnavn(a) &"</font><b><br>"& left(aktiviteter(a), 16) %></b></td>
			<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
			<td align="right" valign="bottom" style="!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px; padding-right : 1px;"><%
			'call aktiviteterTot()
			Response.write SQLBless(timerAktTper(a)) 
			%></td>
			<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
		<td height=20 valign="bottom" align="right" style="!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px;  padding-right : 1px;"><%=SQLBless(ccur(timerAktOmsper(a)))%></td>
	</tr>
	<%'omsAkt = 0 
next%>
	<tr><td colspan="6" height=20><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td></tr>
	<tr>
		<td colspan="2" style="!border: 1 px; background-color: #d2691e; border-color: silver; border-style: solid; padding-left : 4px;">Medarbejdere:</td>
		<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
		<td align="right" style="!border: 1 px; background-color: #d2691e; border-color: silver; border-style: solid; padding-left : 4px;"><font size=1>timer&nbsp;</td>
		<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
		<td align="right" style="!border: 1 px; background-color: #d2691e; border-color: silver; border-style: solid; padding-left : 4px;"><font size=1>oms�tning&nbsp;</td></tr>
	</tr>
	<%
	'** Medarbejder total **
	For t = 0 to t - 1
	%>
	
	<tr><td height=2 colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td></tr>
			<tr>
			<td colspan=2 style="!border: 1 px; background-color: #FFFFFF; border-color: #d2691e; border-style: solid; padding-left : 4px;"><%=medarbejder(t) %></td>
			<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
			<td align=right style="!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px; padding-right : 4px;">
			<%
			'call medarbejdereTot()
			Response.write SQLBless(medarbTtot(t)) 
			%></td>
			<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
		<td valign="bottom" align="right" style="!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px;  padding-right : 1px;"><%=SQLBless(ccur(medarbOmstot(t)))%></td>
	</tr>
	<%
	next
	%>
	</table><br>
	<br>
	
<%
'*** aktiviteter pr. medarbejder ****
%>	
<table cellspacing="0" cellpadding="0" border="0" width="261">	
<tr><td colspan="2" height=20 style="padding-left : 4px;">Aktiviteter fordelt pr. medarbejder:</td>
<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
		<td valign="bottom" align="right"><font size=1>timer&nbsp;&nbsp;&nbsp;</td>
		<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
		<td valign="bottom" align="right"><font size=1>oms�tning&nbsp;</td></tr>
</tr>
<%
For t = 0 to t - 1
if t > 0 then
varHeight = 20
else 
varHeight = 2
end if%>
<tr>
	<td height=<%=varHeight%> colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td></tr>
<tr>
	<td colspan=6 style="!border: 1 px; background-color: #FFFFFF; border-color: #d2691e; border-style: solid; padding-left : 4px;"><b><%=medarbejder(t) %></b></td>
</tr>
<%

For a = 0 to a - 1

	%>
	<tr>
		<td height=2 colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td></tr>
	<tr>
			<td colspan=2 width="125" style="!border: 1 px; background-color: #FFFFFF; border-color: LightBlue; border-style: solid; padding-left : 4px;"><font size="1"><%=jobnr(a) &"&nbsp;" & jobnavn(a) &"</font><b><br>" & left(aktiviteter(a), 16) %></b></td>
			<td width=1><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
			<td width="50" align=right valign="bottom" style="!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px; padding-right : 1px;"><%
			call aktprMedTot()
			%></td>
			<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
			<td width="60" valign="bottom" align="right" style="!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px;  padding-right : 1px;"><%=SQLBless(ccur(omsAktMed))%></td>
			</tr>
<%omsAktMed = 0
	next
next%>
</table><br><br><br>
<!-- slut h�jreside div -->
</div>
<%
Set oRec2 = Nothing
else
%>
<div id="ingenregi" style="position:absolute; left:35; top:140; width:60%; height:600; visibility:visible; z-index:100;">
<table width="480" celpadding="2" cellspacing="1" border="0">
<tr><td height="50" align="center" valign="middle" style="!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px; padding-right : 4px;">Ingen registreringer i den valgte periode.</td></tr></table>
</div>
<%
end if
end if 
Response.write time
%>
<!--#include file="../inc/regular/footer_inc.asp"-->
<script>
document.all["loading"].style.visibility = "hidden";
</script>