

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
	<td  style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 1px; border-color : #003399; border-style : solid;" align="center"><%=SQLBless(sonTimerTot)%>&nbsp;
	<%if sonTimerTot > 24 then%><img src="../ill/alert_lille.gif" width="22" height="19" alt="Der er indtastet mere end 24 timer." border="0"><%end if%>
	<%if sonTimerTot > 0 then%><a href="#" onClick="showtimedetail('timedetailson');"><img src="../ill/lilleplus.gif" width="8" height="7" alt="Se allerede registrerede timer på denne dato" border="0"></a>&nbsp;
	<%else%><img src="../ill/blank.gif" width="8" height="7" alt="" border="0"><%end if%></td>
	<td  style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 1px; border-color : #003399; border-style : solid;" align="center"><%=SQLBless(manTimerTot)%>&nbsp;
	<%if manTimerTot > 24 then%><img src="../ill/alert_lille.gif" width="22" height="19" alt="Der er indtastet mere end 24 timer." border="0"><%end if%>
	<%if manTimerTot > 0 then%><a href="#" onClick="showtimedetail('timedetailman');"><img src="../ill/lilleplus.gif" width="8" height="7" alt="Se allerede registrerede timer på denne dato" border="0"></a>&nbsp;
	<%else%><img src="../ill/blank.gif" width="8" height="7" alt="" border="0"><%end if%></td>
	<td  style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 1px; border-color : #003399; border-style : solid;" align="center"><%=SQLBless(tirTimerTot)%>&nbsp;
	<%if tirTimerTot > 24 then%><img src="../ill/alert_lille.gif" width="22" height="19" alt="Der er indtastet mere end 24 timer." border="0"><%end if%>
	<%if tirTimerTot > 0 then%><a href="#" onClick="showtimedetail('timedetailtir');"><img src="../ill/lilleplus.gif" width="8" height="7" alt="Se allerede registrerede timer på denne dato" border="0"></a>&nbsp;
	<%else%><img src="../ill/blank.gif" width="8" height="7" alt="" border="0"><%end if%></td>
	<td  style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 1px; border-color : #003399; border-style : solid;" align="center"><%=SQLBless(onsTimerTot)%>&nbsp;
	<%if onsTimerTot > 24 then%><img src="../ill/alert_lille.gif" width="22" height="19" alt="Der er indtastet mere end 24 timer." border="0"><%end if%>
	<%if onsTimerTot > 0 then%><a href="#" onClick="showtimedetail('timedetailons');"><img src="../ill/lilleplus.gif" width="8" height="7" alt="Se allerede registrerede timer på denne dato" border="0"></a>&nbsp;
	<%else%><img src="../ill/blank.gif" width="8" height="7" alt="" border="0"><%end if%></td>
	<td  style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 1px; border-color : #003399; border-style : solid;" align="center"><%=SQLBless(torTimerTot)%>&nbsp;
	<%if torTimerTot > 24 then%><img src="../ill/alert_lille.gif" width="22" height="19" alt="Der er indtastet mere end 24 timer." border="0"><%end if%>
	<%if torTimerTot > 0 then%><a href="#" onClick="showtimedetail('timedetailtor');"><img src="../ill/lilleplus.gif" width="8" height="7" alt="Se allerede registrerede timer på denne dato" border="0"></a>&nbsp;
	<%else%><img src="../ill/blank.gif" width="8" height="7" alt="" border="0"><%end if%></td>
	<td  style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 1px; border-color : #003399; border-style : solid;" align="center"><%=SQLBless(freTimerTot)%>&nbsp;
	<%if freTimerTot > 24 then%><img src="../ill/alert_lille.gif" width="22" height="19" alt="Der er indtastet mere end 24 timer." border="0"><%end if%>
	<%if freTimerTot > 0 then%><a href="#" onClick="showtimedetail('timedetailfre');"><img src="../ill/lilleplus.gif" width="8" height="7" alt="Se allerede registrerede timer på denne dato" border="0"></a>&nbsp;
	<%else%><img src="../ill/blank.gif" width="8" height="7" alt="" border="0"><%end if%></td>
	<td  style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 1px; border-color : #003399; border-style : solid;" align="center"><%=SQLBless(lorTimerTot)%>&nbsp;
	<%if lorTimerTot > 24 then%><img src="../ill/alert_lille.gif" width="22" height="19" alt="Der er indtastet mere end 24 timer." border="0"><%end if%>
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
	<td colspan="6" style="border-top : 1px; border-bottom : 1px; border-left : 1px; border-right : 1px; border-color : #003399; border-style : solid;  padding-left:4;">Normeret tid: <b><%=formatnumber(thisnormeret, 2)%></b>&nbsp;/&nbsp;<a href="#" class=vmenu onClick="showrestimer()"><%=formatnumber(timertottildelt, 2)%></a><img src="../ill/blank.gif" width="50" height="1" alt="" border="0">Indtastet denne uge:</td>
	<td colspan="2" align="center" style="border-top : 1px; border-bottom : 1px; border-left : 0px; border-right : 1px; border-color : #003399; border-style : solid;  padding-left:4;"><%=SQLBless(TimerTot)%> timer</td>
	</tr>
	</table>

