<table cellspacing="0" cellpadding="2" border="0" width=540>
	<tr><td colspan="7">Følgende timer er projekteret på din profil i denne uge:<img src="../ill/blank.gif" width="200" height="1" alt="" border="0"><a href="#" onClick=hiderestimer();>Luk vindue</a></td></tr>
	<tr><td colspan=2 valign="top">
	
	<%
	'*** tildelte timer ***
	strSQL3 = "SELECT DISTINCT(ressourcer.id), ressourcer.dato AS rdato, timer AS sumtimer, job.jobnavn, jobid, job.jobnr AS jnr, job.id AS jid, aktiviteter.navn AS anavn FROM ressourcer LEFT JOIN job ON (job.id = ressourcer.jobid) LEFT JOIN aktiviteter ON (aktiviteter.id = aktid) WHERE (ressourcer.dato >= '"&varTjDatoUS_son&"' AND ressourcer.dato <= '"&varTjDatoUS_lor&"') AND mid=" & usemrn &" GROUP BY ressourcer.id ORDER BY ressourcer.dato, jobnavn"
	oRec3.open strSQL3, oConn, 3 
	x = 0
	lastresdato = 0
	while not oRec3.EOF 
		
		if x > 0 AND (lastresdato <> oRec3("rdato"))then%>
		</table></td><td valign="top">
			<%if lastresdato <> oRec3("rdato") then%>
			<table cellspacing="0" cellpadding="0" border="0" width=120><tr><td colspan=2><b><%=formatdatetime(oRec3("rdato"),2)%></b></td></tr>
			<%end if
		else
			if lastresdato <> oRec3("rdato") then%>
			<table cellspacing="0" cellpadding="0" border="0" width=120><tr><td colspan=2><b><%=formatdatetime(oRec3("rdato"),2)%></b></td></tr>
			<%end if
		end if
		
		if lastjobid <> oRec3("jid") OR (lastresdato <> oRec3("rdato")) then%>
		<tr><td colspan=2 class=lille><u>(<%=oRec3("jnr")%>) <%=oRec3("jobnavn")%></u></td></tr>
		<%end if%>
		<tr><td class=lille><%=oRec3("anavn")%></td><td class=lille><%=formatnumber(oRec3("sumtimer"),2)%></td></tr>
	<%
	timertottildelt = timertottildelt + oRec3("sumtimer")
	lastjobid = oRec3("jid")
	lastresdato = oRec3("rdato")
	x = x + 1
	oRec3.movenext
	wend
	oRec3.close
	%>
	</table></td></tr>
	<%if x = 0 then%>
	<tr><td colspan=2><img src="../ill/blank.gif" width="500" height="1" alt="" border="0"><br>Der er ikke tildelt timer på din profil i denne uge.<br>Der kan tildeles timer ved at redigere et job.<br><br>&nbsp;</td></tr>
	<%end if%>
	</table>
	</div>

<!--Viser timer pr. dag i toppen af siden -->
<div style="position:absolute; left:537; top:47;">
	<table cellspacing="0" cellpadding="0" border="0" width="402">
	<tr bgcolor="#5582d2">
		<td style="border-left: 1px solid #003399; border-top: 1px solid #003399;">&nbsp;</td>
		<td class=alt style="padding-left: 2px; border-top: 1px solid #003399;"><font size=1>S <%=day(tjekdag(1))&"-"&month(tjekdag(1))%></td>
		<td class=alt style="padding-left: 2px; border-top: 1px solid #003399;"><font size=1>M <%=day(tjekdag(2))&"-"&month(tjekdag(2))%></td>
		<td class=alt style="padding-left: 2px; border-top: 1px solid #003399;"><font size=1>T <%=day(tjekdag(3))&"-"&month(tjekdag(3))%></td>
		<td class=alt style="padding-left: 2px; border-top: 1px solid #003399;"><font size=1>O <%=day(tjekdag(4))&"-"&month(tjekdag(4))%></td>
		<td class=alt style="padding-left: 2px; border-top: 1px solid #003399;"><font size=1>T <%=day(tjekdag(5))&"-"&month(tjekdag(5))%></td>
		<td class=alt style="padding-left: 2px; border-top: 1px solid #003399;"><font size=1>F <%=day(tjekdag(6))&"-"&month(tjekdag(6))%></td>
		<td class=alt style="border-right: 1px solid #003399; padding-left: 2px; border-top: 1px solid #003399;"><font size=1>L <%=day(tjekdag(7))&"-"&month(tjekdag(7))%></td>
	</tr>
	<tr bgcolor="#ffffff">
	<%
	for intcounter = 1 to 7
	
	select case intcounter
	case 1
	useSQLd = varTjDatoUS_son
	case 2
	useSQLd = varTjDatoUS_man
	case 3
	useSQLd = varTjDatoUS_tir
	case 4
	useSQLd = varTjDatoUS_ons
	case 5
	useSQLd = varTjDatoUS_tor
	case 6
	useSQLd = varTjDatoUS_fre
	case 7
	useSQLd = varTjDatoUS_lor
	end select 
	
	strSQL = "SELECT sum(timer) AS timer_indtastet FROM timer WHERE Tmnr = "& intMnr &" AND Tdato='"& useSQLd &"' AND tfaktim <> 5"
		
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
		select case intcounter
		case 1
		sonTimerTot = oRec("timer_indtastet")
		case 2
		manTimerTot = oRec("timer_indtastet")
		case 3
		tirTimerTot = oRec("timer_indtastet")
		case 4
		onsTimerTot = oRec("timer_indtastet")
		case 5
		torTimerTot = oRec("timer_indtastet")
		case 6
		freTimerTot = oRec("timer_indtastet")
		case 7
		lorTimerTot = oRec("timer_indtastet")
		end select 
	end if
	oRec.close
	
	next
	
	if len(sonTimerTot) <> 0 then
	sonTimerTot = sonTimerTot
	else
	sonTimerTot = 0
	end if
	
	if len(manTimerTot) <> 0 then
	manTimerTot = manTimerTot
	else
	manTimerTot = 0
	end if
	
	if len(tirTimerTot) <> 0 then
	tirTimerTot = tirTimerTot
	else
	tirTimerTot = 0
	end if
	
	if len(onsTimerTot) <> 0 then
	onsTimerTot = onsTimerTot
	else
	onsTimerTot = 0
	end if
	
	if len(torTimerTot) <> 0 then
	torTimerTot = torTimerTot
	else
	torTimerTot = 0
	end if
	
	if len(freTimerTot) <> 0 then
	freTimerTot = freTimerTot
	else
	freTimerTot = 0
	end if
	
	if len(lorTimerTot) <> 0 then
	lorTimerTot = lorTimerTot
	else
	lorTimerTot = 0
	end if
	
	TimerTot = sonTimerTot + manTimerTot + tirTimerTot + onsTimerTot + torTimerTot + freTimerTot + lorTimerTot 
	
	'**** normeret tid ***
	strSQL = "SELECT mid, medarbejdertype, normtimer_son, id, normtimer_man, normtimer_tir, normtimer_ons, normtimer_tor, normtimer_fre, normtimer_lor FROM medarbejdertyper, medarbejdere WHERE medarbejdere.mid = "& session("mid") &" AND id = medarbejdertype"  
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
							
	sontimer = oRec("normtimer_son")
	mantimer = oRec("normtimer_man")
	tirtimer = oRec("normtimer_tir")
	onstimer = oRec("normtimer_ons")
	tortimer = oRec("normtimer_tor")
	fretimer = oRec("normtimer_fre")
	lortimer = oRec("normtimer_lor")
	
	end if
	oRec.close
	
	thisnormeret = (sontimer + mantimer + tirtimer + onstimer +tortimer + fretimer + lortimer)
	%>
	<td style="border-top : 1px; border-bottom : 0px; border-left : 1px; border-right : 1px; border-color : #003399; border-style : solid; padding-left:4;">Timer:</td>
	<td  style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 1px; border-color : #003399; border-style : solid;" align="center"><input type="text" name="FM_son_total" id="FM_son_total" value="<%=SQLBless2(sonTimerTot)%>" style='!border: 1px; background-color: #ffffff; border-color: #ffffff; border-style: solid; padding-left : 1px;padding-right : 2px; width:35px;'>
	<%if sonTimerTot > 0 then%><a href="#" onClick="showtimedetail('timedetailson');"><img src="../ill/lilleplus.gif" width="8" height="7" alt="Se allerede registrerede timer på denne dato" border="0"></a>&nbsp;
	<%else%><img src="../ill/blank.gif" width="8" height="7" alt="" border="0"><%end if%></td>
	<td  style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 1px; border-color : #003399; border-style : solid;" align="center"><input type="text" name="FM_man_total" id="FM_man_total" value="<%=SQLBless2(manTimerTot)%>" style='!border: 1px; background-color: #ffffff; border-color: #ffffff; border-style: solid; padding-left : 1px;padding-right : 2px; width:35px;'>
	<%if manTimerTot > 0 then%><a href="#" onClick="showtimedetail('timedetailman');"><img src="../ill/lilleplus.gif" width="8" height="7" alt="Se allerede registrerede timer på denne dato" border="0"></a>&nbsp;
	<%else%><img src="../ill/blank.gif" width="8" height="7" alt="" border="0"><%end if%></td>
	<td  style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 1px; border-color : #003399; border-style : solid;" align="center"><input type="text" name="FM_tir_total" id="FM_tir_total" value="<%=SQLBless2(tirTimerTot)%>" style='!border: 1px; background-color: #ffffff; border-color: #ffffff; border-style: solid; padding-left : 1px;padding-right : 2px; width:35px;'>
	<%if tirTimerTot > 0 then%><a href="#" onClick="showtimedetail('timedetailtir');"><img src="../ill/lilleplus.gif" width="8" height="7" alt="Se allerede registrerede timer på denne dato" border="0"></a>&nbsp;
	<%else%><img src="../ill/blank.gif" width="8" height="7" alt="" border="0"><%end if%></td>
	<td  style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 1px; border-color : #003399; border-style : solid;" align="center"><input type="text" name="FM_ons_total" id="FM_ons_total" value="<%=SQLBless2(onsTimerTot)%>" style='!border: 1px; background-color: #ffffff; border-color: #ffffff; border-style: solid; padding-left : 1px;padding-right : 2px; width:35px;'>
	<%if onsTimerTot > 0 then%><a href="#" onClick="showtimedetail('timedetailons');"><img src="../ill/lilleplus.gif" width="8" height="7" alt="Se allerede registrerede timer på denne dato" border="0"></a>&nbsp;
	<%else%><img src="../ill/blank.gif" width="8" height="7" alt="" border="0"><%end if%></td>
	<td  style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 1px; border-color : #003399; border-style : solid;" align="center"><input type="text" name="FM_tor_total" id="FM_tor_total" value="<%=SQLBless2(torTimerTot)%>" style='!border: 1px; background-color: #ffffff; border-color: #ffffff; border-style: solid; padding-left : 1px;padding-right : 2px; width:35px;'>
	<%if torTimerTot > 0 then%><a href="#" onClick="showtimedetail('timedetailtor');"><img src="../ill/lilleplus.gif" width="8" height="7" alt="Se allerede registrerede timer på denne dato" border="0"></a>&nbsp;
	<%else%><img src="../ill/blank.gif" width="8" height="7" alt="" border="0"><%end if%></td>
	<td  style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 1px; border-color : #003399; border-style : solid;" align="center"><input type="text" name="FM_fre_total" id="FM_fre_total" value="<%=SQLBless2(freTimerTot)%>" style='!border: 1px; background-color: #ffffff; border-color: #ffffff; border-style: solid; padding-left : 1px;padding-right : 2px; width:35px;'>
	<%if freTimerTot > 0 then%><a href="#" onClick="showtimedetail('timedetailfre');"><img src="../ill/lilleplus.gif" width="8" height="7" alt="Se allerede registrerede timer på denne dato" border="0"></a>&nbsp;
	<%else%><img src="../ill/blank.gif" width="8" height="7" alt="" border="0"><%end if%></td>
	<td  style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 1px; border-color : #003399; border-style : solid;" align="center"><input type="text" name="FM_lor_total" id="FM_lor_total" value="<%=SQLBless2(lorTimerTot)%>" style='!border: 1px; background-color: #ffffff; border-color: #ffffff; border-style: solid; padding-left : 1px;padding-right : 2px; width:35px;'>
	<%if lorTimerTot > 0 then%><a href="#" onClick="showtimedetail('timedetaillor');"><img src="../ill/lilleplus.gif" width="8" height="7" alt="Se allerede registrerede timer på denne dato" border="0"></a>&nbsp;
	<%else%><img src="../ill/blank.gif" width="8" height="7" alt="" border="0"><%end if%></td>
	</tr>
	<%
	if len(timertottildelt) <> 0 then
	timertottildelt = timertottildelt
	else
	timertottildelt = 0
	end if
	%>
	<tr bgcolor="#ffffff">
	<td colspan="6" style="border-top : 1px; border-bottom : 1px; border-left : 1px; border-right : 1px; border-color : #003399; border-style : solid;  padding-left:4;">Normeret tid: <b><%=formatnumber(thisnormeret, 2)%></b>&nbsp;/&nbsp;Din uge: <a href="#" class=vmenu onClick="showrestimer()"><%=formatnumber(timertottildelt, 2)%></a><img src="../ill/blank.gif" width="50" height="1" alt="" border="0">Indtastet uge <b><%=strWeek%></b>:</td>
	<td colspan="2" align="center" style="border-top : 1px; border-bottom : 1px; border-left : 0px; border-right : 1px; border-color : #003399; border-style : solid;  padding-left:4;"><input type="text" name="FM_week_total" id="FM_week_total" value="<%=formatnumber(SQLBless2(TimerTot), 2)%>" style='!border: 1px; background-color: #ffffff; border-color: #ffffff; border-style: solid; padding-left : 1px;padding-right : 2px; width:35px;' maxlenght=4> timer</td>
	</tr>
	</table>