
<%

'********************** Dato linie sub **********************
public statdatoM3, statdatoM2, statdatoM1
sub datoeroverskrift
	'*** Datoer ***
		
		%>
		<tr>
		<%
		
		y = 0
		for y = -2 TO numoffdaysorweeksinperiode 
				
				
				call antaliperiode(periodeSel, y, startdato, monthUse)
				
				if y = -2 then 'lastMonth <> newMonth OR%>
				<td align=right bgcolor="#eff3ff" style="padding-left:5px; padding-right:5px;">&nbsp;</td>
				<%
				'csvTxtTop = csvTxtTop &  ";"
				lastMonth = newMonth
				end if
				
				if y = -1 then
				
				%>
				<td valign=top bgcolor="#eff3ff" align=right class=lille style="padding-top:1px; padding-right:5px; border-left:2px #8caae6 solid;">&nbsp;</td>
				<td valign=top bgcolor="#eff3ff" align=center style="padding:2px;">
				<%
				dato_3 = left(monthname(datepart("m", dateadd("m", -3, startdato))), 3) & " " & right(newYear, 2)%>
				<b><%=dato_3%></b>
				<%statdatoM3 = dateadd("m", -3, startdato)%>
				</td>
				<td valign=top bgcolor="#eff3ff" align=center style="padding:2px;">
				<%dato_2 = left(monthname(datepart("m", dateadd("m", -2, startdato))), 3) & " " & right(newYear, 2)%>
				<b><%=dato_2%></b>
				<%statdatoM2 = dateadd("m", -2, startdato)%>
				</td>
				<td valign=top bgcolor="#eff3ff" align=center style="padding:2px; border-right:2px #8caae6 solid;">
				<%dato_1 = left(monthname(datepart("m", dateadd("m", -1, startdato))), 3) & " " & right(newYear, 2)%>
				<b><%=dato_1%></b>
				<%statdatoM1 = dateadd("m", -1, startdato)%>
				</td>
				
				<%
					if skrivCsvFil = 1 then
					csvDatoer(35) = dato_3 
					csvDatoer(34) = dato_2 
					csvDatoer(33) = dato_1 
					end if
				
				
				end if
				
				
				if y > -1 then
					
					if periodeSel = 1 then
						
						select case weekday(nday, 2)
						case 6
						dbg = "Gainsboro"
						case 2
						dbg = "#eff3ff"
						case 3
						dbg = "#eff3ff"
						case 4
						dbg = "#eff3ff"
						case 5
						dbg = "#eff3ff"
						case 1
						dbg = "#eff3ff"
						case 7
						dbg = "Gainsboro"
						end select
						
					call helligdage(nday)
					if erHellig = 1 then
					dbg = "Gainsboro"
					end if
				
					if cdate(date) = cdate(nday) then
					dbg = "#FFff99"
					end if
				
				else
					dbg = "#eff3ff"
				end if
				
				
				%>
				
				<td valign=top bgcolor="<%=dbg%>" align=center style="padding:1px;"><b>
				<%
				
				select case periodeSel 
				case 1
				tdvalThis = left(formatdatetime(nday, 2), 2)
					
					if skrivCsvFil = 1 then
					csvDatoer(y) = left(monthname(datepart("m", nday)), 3) & "-" & right(newYear, 2) & " " & tdvalThis
					end if
					
				case 12
				tdvalThis =  left(monthname(datepart("m", nday)), 3) & " " & right(newYear, 2)
					
					if skrivCsvFil = 1 then
					csvDatoer(y) = tdvalThis
					end if
					
				end select
				
				Response.write tdvalThis
				
				
				%></b>
				</td>
				<%
				
				
				
				end if
				
			next
			
	
	end sub
	
	
	
	
	sub medarbtotal
	
				'** Skal total linie vises? ***
				timersum = 0
				timerTxt = ""
				for y = 0 TO numoffdaysorweeksinperiode
				
				timersum = timersum + medarbTotalTimer(y)
				timerTxt = timerTxt & "<td bgcolor=#ffdfdf align=center><b>"&formatnumber(medarbTotalTimer(y), 1)&"</b></td>"
				
				medarbTotalTimer(y) = 0
				next
				
				
		if timersum <> 0 then%>
		<tr>	
			<td height=20 bgcolor="#ffdfdf" align=right style="padding-right:15px;"><b>Total:</b></td>
			<td valign=top bgcolor="#ffdfdf" align=right class=lille style="padding-top:1px; padding-right:5px; border-left:2px #8caae6 solid;">
		Forecast:<br>
		Realiseret:<br>
		Afvigelse:</td>	<%
		
		'csvTxt = csvTxt & ";Total;Forecast;Realiseret;Afvigelse"& 
				
				for s = 0 to 2
				
					select case s
					case 0
					bdleft = 0
					bdright = 0
					case 1
					bdleft = 0
					bdright = 0
					case 2
					bdleft = 0
					bdright = 2
					end select
					
					%>
					<td valign=top bgcolor="#ffdfdf" align=right class=lille style="padding-top:1px; padding-right:5px;  border-right:<%=bdright%> #8caae6 solid; border-left:<%=bdleft%> #8caae6 solid;">
					<%
					
					afvTot = 0
					
					select case s
					case 0
					Response.write formatnumber(totTildtimPer(1), 1) &"<br>"& formatnumber(totFaktisk3, 1) 
					
					if totFaktisk3 <> 0 then
							
						if totTildtimPer(1) = 0 then
						afvTot = totFaktisk3
						else
						afvTot = 100 -((totTildtimPer(1) / totFaktisk3) * 100)
						end if
					
							
							'if afvTot >= 100 then
							'afvTot = afvTot - 100 
							'else
							'avtTot = 100 - afvTot
							'end if
							
					else
							if totTildtimPer(1) = 0 then
							afvTot = 0
							else
							afvTot = totTildtimPer(1) '100
							end if
					end if
					
					totTildtimPer(1) = 0
					totFaktisk3 = 0
					
					case 1
					Response.write formatnumber(totTildtimPer(2), 1) &"<br>"& formatnumber(totFaktisk2, 1) &"" 
					
					if totFaktisk2 <> 0 then
						
						
						if totTildtimPer(2) = 0 then
						afvTot = totFaktisk3
						else
						afvTot = 100 -((totTildtimPer(2) / totFaktisk2) * 100)
						end if
							
						'if afvTot >= 100 then
						'afvTot = afvTot - 100 
						'else
						'avtTot = 100 - afvTot
						'end if
							
					else
							if totTildtimPer(2) = 0 then
							afvTot = 0
							else
							afvTot = totTildtimPer(2) '100
							end if
					end if
					
					totFaktisk2 = 0
					totTildtimPer(2) = 0
					
					case 2
					Response.write formatnumber(totTildtimPer(3), 1) &"<br>"& formatnumber(totFaktisk1, 1)
					
					if totFaktisk1 <> 0 then
						
						
						if totTildtimPer(3) = 0 then
						afvTot = totFaktisk3
						else
						afvTot = 100 -((totTildtimPer(3) / totFaktisk1) * 100)
						end if
						
						'if afvTot >= 100 then
						'afvTot = afvTot - 100 
						'else
						'avtTot = 100 - afvTot
						'end if
						
						
					else
							if totTildtimPer(3) = 0 then
							afvTot = 0
							else
							afvTot = totTildtimPer(3) '100
							end if
					end if
					
					totFaktisk1 = 0
					totTildtimPer(3) = 0
					
					end select
					
					
					
					'*** Afvigelse ***
					Response.write "<br>"& formatnumber(afvTot, 0) &"%"
					afvTot = 0
					%>
					</td>
				<%next
				
				
				
				
				Response.write timerTxt
				%>
				</tr>
				<%
				end if
				%>
	
	<%
	end sub
	
	
	
	
	
		'***************** Opdater db sub ***********************
		function opdaterdb(timertd)
			
			'Response.write timertd & "!<br>"
			
			if len(trim(timertd)) <> 0 then
				
				'**** Sletter evt. gamle registreringer ***
				strSQLdel = "DELETE FROM ressourcer WHERE jobid = "& tjobid(y) &" AND mid = "& tmedid(y) &" AND dato = '"& tDato &"'"
				'Response.write strSQLdel & "<br>"
				oConn.execute(strSQLdel)
				
				'*** Indsætter nye registreringer ***
				if timertd <> 0 then
				
				timertdkomma = replace(timertd, ".", ",")
				sttid = "09:00:00"
				sltid = formatdatetime(dateadd("h", timertdkomma, datoer(y)&" "&sttid ), 3)
				
				strSQLupd = "INSERT INTO ressourcer (jobid, mid, dato, timer, starttp, sluttp) VALUES ("& tjobid(y) &", "& tmedid(y) &", '"& tDato &"', "& timertd &", '"& sttid &"', '"& sltid &"')"
				'Response.write strSQLupd & "<hr>"
				'Response.flush
				oConn.execute(strSQLupd)
				
				end if
			
			end if
		end function
		
		
		
		
		
		
		
		
		
		
	'*** Function periodeantal ***
	public newMonth, nday, newYear
	function antaliperiode(peri, ycount, sdato, mduse)
	
				select case peri
				case 1
					if ycount = 0 OR ycount = -2 then
					nday = dateadd("d", 0, sdato)
					else
					nday = dateadd("d", 1, nday)
					end if
					
					newMonth = mduse 
					newYear = yearthis
					
				
					
				case 12
					
					if ycount = 0 OR ycount = -2  then
					nday = dateadd("m", 0, sdato)
					else
					nday = dateadd("m", 1, nday)
					end if
					
					newMonth = datepart("m", nday)
					newYear = datePart("yyyy", nDay)
					
					
				end select
	
	end function
	
	
	
	
	
	
	'***** Function Normtimer + opdater dage ****
	'*** Henter normeret timer ved 1/1 dag. ***
	public tildeltimer_timer
	function normtimer(tildeltimer_timer, dato, medid, jobid)
	
	sqlDato = year(dato)&"/"&month(dato)&"/"&day(dato)
	
	
		if tildeltimer_timer = 7 then
		
				select case weekday(dato)
				case 1
				normdag = "normtimer_son"
				case 2
				normdag = "normtimer_man"
				case 3
				normdag = "normtimer_tir"
				case 4
				normdag = "normtimer_ons"
				case 5
				normdag = "normtimer_tor"
				case 6
				normdag = "normtimer_fre"
				case 7
				normdag = "normtimer_lor"
				case else
				normdag = "normtimer_man"
				end select
				
				
			strSQL = "SELECT medarbejdertype, "& normdag &" AS timer FROM medarbejdere m"_
			&" LEFT JOIN medarbejdertyper t ON (t.id = m. medarbejdertype)"_
			&" WHERE m.mid = " & medid
			
			'Response.write strSQL & "<hr>"
			'Response.flush
			
			oRec.open strSQL, oConn, 3 
			while not oRec.EOF 
				timerThis = oRec("timer")
			oRec.movenext
			wend
			oRec.close 
			
			
			if len(timerThis) <> 0 then
			timerThis = replace(timerThis, ",",".")
			else
			timerThis = 0
			end if
		
		else
		timerThis = tildeltimer_timer
		end if
		
		sttid = "09:00:00"
		sltid = formatdatetime(dateadd("h", timerThis, dato &" "&sttid ), 3)
		
		'*** Sletter gamle ressource registreringer ****
		strSQL_del = "DELETE FROM ressourcer WHERE dato = '"& sqlDato &"' "_
		&" AND jobid = "& jobid &" AND mid = " & medid
		oConn.execute(strSQL_del)
		
		if timerThis <> 0 then
		'**** Oprettter nye ressource registreringer ****
		strSQL = "INSERT INTO ressourcer (dato, mid, jobid, aktid, starttp, sluttp, timer) VALUES "_
		&" ('"& sqlDato &"', "& medid &", "& jobid &", 0, '"& sttid &"', '"& sltid &"', "& timerThis &")"
		'Response.write strSQL & "<br><br>"
		'Response.flush
		oConn.execute(strSQL)
		end if
		
	end function
	'***************************************************
	
	
	
	
	
	
	
	public ntimPer 
	function normtimerPer(medid)
			
			
			ntimPer = 0
			
			for n = 1 to 7 
			
			
					
						
						select case n 
						case 7
						normdag = "normtimer_son"
						nd1 = 0
						case 1
						normdag = "normtimer_man"
						nd2 = 0
						case 2
						normdag = "normtimer_tir"
						nd3 = 0
						case 3
						normdag = "normtimer_ons"
						nd4 = 0
						case 4
						normdag = "normtimer_tor"
						nd5 = 0
						case 5
						normdag = "normtimer_fre"
						nd6 = 0
						case 6
						normdag = "normtimer_lor"
						nd7 = 0
						end select
						
						
					strSQL = "SELECT medarbejdertype, "& normdag &" AS timer FROM medarbejdere m"_
					&" LEFT JOIN medarbejdertyper t ON (t.id = m. medarbejdertype)"_
					&" WHERE m.mid = " & medid
					
					'Response.write strSQL & "<hr>"
					'Response.flush
					
					oRec.open strSQL, oConn, 3 
					if not oRec.EOF then
						ntimPerThis = oRec("timer")
						
					end if
					oRec.close 
					
					
					ntimPer = ntimPer + ntimPerThis	
							
					
					
			next
			
			if len(ntimPer) <> 0 then
			ntimPer = ntimPer
			else
			ntimPer = 0
			end if
			
			
	end function
	
	
	dim totTildtimPer
	redim totTildtimPer(3)
	
	public tildtimPer 
	function tildelttimerPer(medid, jobid, datothis, s)
			
			
			tildtimPer = 0
			
					strSQL2 = "SELECT sum(r.timer) AS sumtimer FROM ressourcer r WHERE r.mid = "& medid &" AND r.jobid = "& jobid &" AND (YEAR(r.dato) = "& year(datothis) &" AND MONTH(r.dato) = "& month(datothis) &")" 'AND r.dato BETWEEN '"& stdato &"' AND '"& sldato &"'"
					
					'Response.write strSQL
					'Response.flush
					
					oRec2.open strSQL2, oConn, 3 
					if not oRec2.EOF then
						tildtimPer = oRec2("sumtimer")
					'oRec.movenext
					end if
					oRec2.close 
					
					if len(tildtimPer) <> 0 then
					tildtimPer = tildtimPer
					else
					tildtimPer = 0
					end if
					
					
					select case s
					case 0
					totTildtimPer(1) = totTildtimPer(1) + tildtimPer
					case 1
					totTildtimPer(2) = totTildtimPer(2) + tildtimPer
					case 2
					totTildtimPer(3) = totTildtimPer(3) + tildtimPer
					end select
					
					'Response.write "#"& totTildtimPer &"<br>"
					
	end function
	
	
	
	function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
	end function
	
	
	
	
	
	'**** Opdater weekender ********
	
	public tTimertildelt_tp
	function opd_weekends_yes_no(timertd)
	
	'tTimertildelt_tp = 0
	
	if len(trim(timertd)) <> 0 then
		tTimertildelt_tp = timertd
		tTimertildelt_tp = formatnumber(tTimertildelt_tp, 1)
		tTimertildelt_tp = replace(tTimertildelt_tp, ",", ".")
	else
		timertd = 0
		tTimertildelt_tp = timertd '0
	end if
	
	
	'Response.write "tTimertildelt_tp" & tTimertildelt_tp & "<br>"
	
		if cint(weekend_2_On) = 0 then
		
				if weekday(tDato, 2) < 6 then
					
					call helligdage(tDato)
					
					if erHellig <> 1 then
						call opdaterdb(tTimertildelt_tp)
						tildeltsofar = tildeltsofar + timertd
						tildeltialt = tildeltialt + timertd
					end if
					
				end if
			else
					
					call opdaterdb(tTimertildelt_tp)
					tildeltsofar = tildeltsofar + timertd
					tildeltialt = tildeltialt + timertd
			end if

	end function
	
	%>
	
	
		
		