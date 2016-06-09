

<table cellspacing="0" cellpadding="0" border="0" height=100>
	<tr><td colspan="7" valign=top>Følgende timer er projekteret på din profil i denne uge:<img src="../ill/blank.gif" width="500" height="1" alt="" border="0"><a href="#" onClick=hiderestimer(); class=red>[Luk vindue]</a>&nbsp;&nbsp;</td></tr>
	<tr><td colspan=2 valign="top">
	
	<%
	'*** tildelte timer ***
	strSQL3 = "SELECT ressourcer.id, ressourcer.dato AS rdato, timer AS sumtimer, job.jobnavn, jobid, job.jobnr AS jnr, job.id AS jid FROM ressourcer "_
	&" LEFT JOIN job ON (job.id = ressourcer.jobid) WHERE (ressourcer.dato >= '"&varTjDatoUS_man&"' AND ressourcer.dato <= '"&varTjDatoUS_son&"') AND mid=" & usemrn &""_
	&" GROUP BY ressourcer.dato, jobid, jobid ORDER BY ressourcer.dato, jobnavn"
	
	'Response.write strSQL3
	
	oRec3.open strSQL3, oConn, 3 
	x = 0
	lastresdato = 0
	while not oRec3.EOF 
		
		if x > 0 AND (lastresdato <> oRec3("rdato"))then%>
		</table></td><td valign="top">
			<%if lastresdato <> oRec3("rdato") then%>
			<table cellspacing="0" cellpadding="0" border="0"><tr><td colspan=2 style="padding-left:5px;"><b><%=formatdatetime(oRec3("rdato"),2)%></b></td></tr>
			<%end if
		else
			if lastresdato <> oRec3("rdato") then%>
			<table cellspacing="0" cellpadding="0" border="0"><tr><td colspan=2 style="padding-left:5px;"><b><%=formatdatetime(oRec3("rdato"),2)%></b></td></tr>
			<%end if
		end if
		%>
		<tr><td class=lille width=200 style="padding-left:5px;"><%=left(oRec3("jobnavn"), 15)%> (<%=oRec3("jnr")%>)</td>
		<td width=40 class=lille bgcolor="#cccccc" style="padding-left:5px; border-left:1px #999999 dashed; border-right:1px #999999 dashed;"><%=formatnumber(oRec3("sumtimer"),2)%></td></tr>
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
	<tr><td colspan=2><img src="../ill/blank.gif" width="600" height="1" alt="" border="0">
	<br>
	<font color=red><b>!</b></font> <b>Der er ikke tildelt ressourcetimer på din profil i denne uge.</b><br>
	Der kan tildeles timer ved at redigere et job.<br><br>&nbsp;</td></tr>
	<%end if%>
	</table>
	</div>

<!--Viser timer pr. dag i toppen af siden -->
<div style="position:absolute; left:537; top:47;">
	<table cellspacing="1" cellpadding="0" border="0" width="402" bgcolor="#003399">
	<tr bgcolor="#5582d2">
		<td>&nbsp;</td>
		<td class=alt style="padding-left: 2px;"><font size=1>M <%=day(tjekdag(1))&"-"&month(tjekdag(1))%></td>
		<td class=alt style="padding-left: 2px;"><font size=1>T <%=day(tjekdag(2))&"-"&month(tjekdag(2))%></td>
		<td class=alt style="padding-left: 2px;"><font size=1>O <%=day(tjekdag(3))&"-"&month(tjekdag(3))%></td>
		<td class=alt style="padding-left: 2px;"><font size=1>T <%=day(tjekdag(4))&"-"&month(tjekdag(4))%></td>
		<td class=alt style="padding-left: 2px;"><font size=1>F <%=day(tjekdag(5))&"-"&month(tjekdag(5))%></td>
		<td class=alt style="padding-left: 2px;"><font size=1>L <%=day(tjekdag(6))&"-"&month(tjekdag(6))%></td>
		<td class=alt><font size=1>S <%=day(tjekdag(7))&"-"&month(tjekdag(7))%></td>
	</tr>
	<tr bgcolor="#ffffff">
	<%
	for intcounter = 1 to 7
	
	select case intcounter
	case 1
	useSQLd = varTjDatoUS_man
	case 2
	useSQLd = varTjDatoUS_tir
	case 3
	useSQLd = varTjDatoUS_ons
	case 4
	useSQLd = varTjDatoUS_tor
	case 5
	useSQLd = varTjDatoUS_fre
	case 6
	useSQLd = varTjDatoUS_lor
	case 7
	useSQLd = varTjDatoUS_son
	end select 
	
	strSQL = "SELECT sum(timer) AS timer_indtastet FROM timer WHERE Tmnr = "& intMnr &" AND Tdato='"& useSQLd &"' AND tfaktim <> 5"
	'Response.write intcounter & "<br>"& strSQL & "<br>"
		
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
		select case intcounter
		case 7
		sonTimerTot = oRec("timer_indtastet")
		case 1
		manTimerTot = oRec("timer_indtastet")
		case 2
		tirTimerTot = oRec("timer_indtastet")
		case 3
		onsTimerTot = oRec("timer_indtastet")
		case 4
		torTimerTot = oRec("timer_indtastet")
		case 5
		freTimerTot = oRec("timer_indtastet")
		case 6
		lorTimerTot = oRec("timer_indtastet")
		end select 
	end if
	oRec.close
	
	
	'**** login historik ****
	
	strSQL = "SELECT l.id AS lid, l.mid AS lmid, l.login, l.logud, "_
	&" s.navn AS stempelurnavn, s.faktor, s.minimum FROM login_historik l"_
	&" LEFT JOIN stempelur s ON (s.id = l.stempelurindstilling) WHERE l.dato = '"& useSQLd &"' AND l.mid = " & intMnr &""_
	&" ORDER BY l.login" 
	
	'Response.write strSQL
	'Response.flush
	
	x = 0
	oRec.open strSQL, oConn, 3 
	while not oRec.EOF 
	
		timerThis = 0
		timerThisDIFF = 0
		
		if len(oRec("login")) <> 0 AND len(oRec("logud")) <> 0 then
		
		loginTidAfr = left(formatdatetime(oRec("login"), 3), 5)
		logudTidAfr = left(formatdatetime(oRec("logud"), 3), 5)
		
		timerThisDIFF = datediff("s", loginTidAfr, logudTidAfr)/60
		
		if timerThisDIFF < oRec("minimum") then
			timerThisDIFF = oRec("minimum")
		end if
		
			
			if oRec("faktor") > 0 then
			useFaktor = oRec("faktor")
			else
			useFaktor = 0
			end if
		
		
		timerThis = (timerThisDIFF * useFaktor)
		'totaltimer = totaltimer + timerThis
		'Response.write oRec("lid") & ": " &  timerThis &" - "
		end if
		
		
		timTemp = formatnumber(timerThis/60, 3)
		timTemp_komma = split(timTemp, ",")
		
		for x = 0 to UBOUND(timTemp_komma)
			
			if x = 0 then
			thours = timTemp_komma(x)
			end if
			
			if x = 1 then
			tmin = timerThis - (thours * 60)
			end if
			
		next
		
		'totalhours = totalhours + cint(thours)
		'totalmin = totalmin + tmin
	
		select case intcounter
		case 1
		manMin = manMin + timerThis 
		case 2
		tirMin = tirMin + timerThis 
		case 3
		onsMin = onsMin + timerThis 
		case 4
		torMin = tormin + timerThis 
		case 5
		freMin = freMin + timerThis 
		case 6
		lorMin = lorMin + timerThis 
		case 7
		sonMin = sonMin + timerThis 
		end select 
		
		'thours = 0
		'tmin = 0
		
	oRec.movenext
	wend
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
	<td style="padding-left:4px;" class=lille>Reg. job timer:</td>
	
	<td align="right"><input type="text" name="FM_man_total" id="FM_man_total" value="<%=SQLBless2(manTimerTot)%>" style='!border: 1px; background-color: #ffffff; border-color: #ffffff; border-style: solid; padding-left : 1px;padding-right : 2px; width:25px; font-size:9px; font-family:arial;'>
	<%if manTimerTot > 0 then%><a href="#" onClick="showtimedetail('timedetailman');"><img src="../ill/lilleplus.gif" width="8" height="7" alt="Se allerede registrerede timer på denne dato" border="0"></a>
	<%else%><img src="../ill/blank.gif" width="8" height="7" alt="" border="0">
	<%end if%></td>
	
	<td align="center"><input type="text" name="FM_tir_total" id="FM_tir_total" value="<%=SQLBless2(tirTimerTot)%>" style='!border: 1px; background-color: #ffffff; border-color: #ffffff; border-style: solid; padding-left : 1px;padding-right : 2px; width:25px; font-size:9px; font-family:arial;'> 
	<%if tirTimerTot > 0 then%><a href="#" onClick="showtimedetail('timedetailtir');"><img src="../ill/lilleplus.gif" width="8" height="7" alt="Se allerede registrerede timer på denne dato" border="0"></a>
	<%else%><img src="../ill/blank.gif" width="8" height="7" alt="" border="0">
	<%end if%></td>
	
	<td align="center"><input type="text" name="FM_ons_total" id="FM_ons_total" value="<%=SQLBless2(onsTimerTot)%>" style='!border: 1px; background-color: #ffffff; border-color: #ffffff; border-style: solid; padding-left : 1px;padding-right : 2px; width:25px; font-size:9px; font-family:arial;'>
	<%if onsTimerTot > 0 then%><a href="#" onClick="showtimedetail('timedetailons');"><img src="../ill/lilleplus.gif" width="8" height="7" alt="Se allerede registrerede timer på denne dato" border="0"></a>
	<%else%>
	<img src="../ill/blank.gif" width="8" height="7" alt="" border="0">
	<%end if%></td>
	
	<td align="center"><input type="text" name="FM_tor_total" id="FM_tor_total" value="<%=SQLBless2(torTimerTot)%>" style='!border: 1px; background-color: #ffffff; border-color: #ffffff; border-style: solid; padding-left : 1px;padding-right : 2px; width:25px; font-size:9px; font-family:arial;'>
	<%if torTimerTot > 0 then%><a href="#" onClick="showtimedetail('timedetailtor');"><img src="../ill/lilleplus.gif" width="8" height="7" alt="Se allerede registrerede timer på denne dato" border="0"></a>
	<%else%>
	<img src="../ill/blank.gif" width="8" height="7" alt="" border="0">
	<%end if%></td>
	
	<td align="center"><input type="text" name="FM_fre_total" id="FM_fre_total" value="<%=SQLBless2(freTimerTot)%>" style='!border: 1px; background-color: #ffffff; border-color: #ffffff; border-style: solid; padding-left : 1px;padding-right : 2px; width:25px; font-size:9px; font-family:arial;'>
	<%if freTimerTot > 0 then%><a href="#" onClick="showtimedetail('timedetailfre');"><img src="../ill/lilleplus.gif" width="8" height="7" alt="Se allerede registrerede timer på denne dato" border="0"></a>
	<%else%>
	<img src="../ill/blank.gif" width="8" height="7" alt="" border="0">
	<%end if%></td>
	
	<td align="center"><input type="text" name="FM_lor_total" id="FM_lor_total" value="<%=SQLBless2(lorTimerTot)%>" style='!border: 1px; background-color: #ffffff; border-color: #ffffff; border-style: solid; padding-left : 1px;padding-right : 2px; width:25px; font-size:9px; font-family:arial;'> 
	<%if lorTimerTot > 0 then%><a href="#" onClick="showtimedetail('timedetaillor');"><img src="../ill/lilleplus.gif" width="8" height="7" alt="Se allerede registrerede timer på denne dato" border="0"></a>
	<%else%>
	<img src="../ill/blank.gif" width="8" height="7" alt="" border="0">
	<%end if%></td>
	
	<td align="center"><input type="text" name="FM_son_total" id="FM_son_total" value="<%=SQLBless2(sonTimerTot)%>" style='!border: 1px; background-color: #ffffff; border-color: #ffffff; border-style: solid; padding-left : 1px;padding-right : 2px; width:25px; font-size:9px; font-family:arial;'> 
	<%if sonTimerTot > 0 then%><a href="#" onClick="showtimedetail('timedetailson');"><img src="../ill/lilleplus.gif" width="8" height="7" alt="Se allerede registrerede timer på denne dato" border="0"></a>
	<%else%>
	<img src="../ill/blank.gif" width="8" height="7" alt="" border="0">
	<%end if%></td>
	</tr>
	
	
	
	
	<%
	if len(timertottildelt) <> 0 then
	timertottildelt = timertottildelt
	else
	timertottildelt = 0
	end if
	%>
	<tr bgcolor="#ffffff">
	<td colspan="6" style="padding-left:4px;">Normeret tid: <b><%=formatnumber(thisnormeret, 2)%></b>&nbsp;/&nbsp;Din uge: <a href="#" class=vmenu onClick="showrestimer()"><%=formatnumber(timertottildelt, 2)%></a><!--<img src="../ill/blank.gif" width="30" height="1" alt="" border="0">Indtastet uge <b><=strWeek%></b>:--></td>
	<td colspan="2" align="center"><input type="text" name="FM_week_total" id="FM_week_total" value="<%=formatnumber(SQLBless2(TimerTot), 2)%>" style='!border: 1px; background-color: #ffffff; border-color: #ffffff; border-style: solid; padding-left : 1px;padding-right : 2px; width:25px; font-size:9px;' maxlenght=4> timer</td>
	</tr>
	
	<%if session("stempelur") <> 0 then%>
	<tr bgcolor="#ffff99">
		<td style="padding-left:4px;"><font class=megetlillesort>Login timer:</font>&nbsp;<a href="#" onClick="visstempelur()" class=vmenu><b>[Se]</b></a></td>
		<td align=center><%call timerogminutberegning(manMin)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%></td>
		<td align=center><%call timerogminutberegning(tirMin)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%></td>
		<td align=center><%call timerogminutberegning(onsMin)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%></td>
		<td align=center><%call timerogminutberegning(torMin)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%></td>
		<td align=center><%call timerogminutberegning(freMin)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%></td>
		<td align=center><%call timerogminutberegning(lorMin)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%></td>
		<td align=center><%call timerogminutberegning(sonMin)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%></td>
	</tr>
	<%end if%>
	</table>