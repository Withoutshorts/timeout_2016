




<%

'********************** Dato linie sub **********************
public statdatoM3, statdatoM2, statdatoM1, datoMD_3, datoMD_2, datoMD_1, datoAR_3, datoAR_2, datoAR_1, csvTxtTopMd, timerProcExt, dividerThisMth, dividerAll, antalJobAktlinierGrand, antalJobAktMedlinier
public firstdayOfFirstWeekThis, firstdayOfLastWeekThis, firstThursdayOfFirstWeekThis, dividerThisMthArr, firstOfmonth
redim dividerThisMth(14)



function datoogDivider(nday) ', firstThursdayOfFirstWeekThis


                    'response.write "nday: "& nday & "<br>"

                    thisMth = month(nday)
                    
                    if y = 0 then
                    thisMthFirst = "1-"& month(nday) & "-"& year(nday)
                    else
                    thisMthFirst = dateAdd("m", 1, firstThursdayOfFirstWeekThis)
                    end if   
                    
                    
                    dayOfWeekThis = datePart("w",thisMthFirst,2,2)
        
                    '1 = (dayOfWeek - d) 
                    dThis = (dayOfWeekThis - 1)
                    firstdayOfFirstWeekThis = dateAdd("d", -dThis, thisMthFirst)
                    firstThursdayOfFirstWeekThis = firstdayOfFirstWeekThis 'dateAdd("d", +3, firstdayOfFirstWeekThis)

                    if month(firstThursdayOfFirstWeekThis) = month(nday) then
                    firstThursdayOfFirstWeekThis = firstThursdayOfFirstWeekThis
                    else
                    firstThursdayOfFirstWeekThis = dateAdd("d", 7, firstThursdayOfFirstWeekThis)
                    end if 
            
                    nextMonth = dateAdd("m", 1, thisMthFirst)
                    nextMonthFirst = "1-"& month(nextMonth) & "-"& year(nextMonth)   
                    lastDInMonth = dateAdd("d", -1, nextMonthFirst) 'for at sikere vi kommer tilbage midt i ugen f�r. Der er den sidste uge i m�neden f�r.

                    dayOfLastWeekThis = datePart("w",lastDInMonth,2,2) - 1
                    firstdayOfLastWeekThis = dateAdd("d", -dayOfLastWeekThis, lastDInMonth)
                 

                    dividerThisMth(thisMth) = datediff("w", firstThursdayOfFirstWeekThis, firstdayOfLastWeekThis,2,2) + 1
                    firstOfmonth = thisMthFirst   
                    

                    dividerAll = dividerAll + 1 


end function


sub datoeroverskrift(val)
	'*** Datoer ***

                    '*** Timer eller proc extension
                    if cint(showasproc) = 1 then
                    timerProcExt = " %"
                    else
                    timerProcExt = " t."
                    end if
		
		%>
		<tr>
		<%
		
		y = -1
		for y = 0 TO numoffdaysorweeksinperiode 
				
				
				
				call antaliperiode(periodeSel, y, startdato, monthUse)
				
				
				
										if y = 0 then %>
                                        <td bgcolor="#ffffff" style="padding:3px 3px 3px 3px;">Medarbejder</td>
										<td bgcolor="#ffffff" style="padding:3px 3px 3px 3px;">Kunde<br /><b>Projekt / Job</b>										
										<%if val = 2 then %>
										<br /><%=medarbKundeoplysX(x, 4)%> (<%=medarbKundeoplysX(x, 7)%>)
										<%end if %>
										<img src="../ill/blank.gif" width="200" height="1" border="0" />
										</td>
                                        
                                        <%if cint(visRamme) = 1 then %>
                                        <td bgcolor="#ffffff" style="width:60px; padding:3px 3px 3px 3px;">�rsramme <br /><span style="color:#999999; font-size:9px;">Timer pr. �r</span></td>
                                        <%end if %>
                                        
                                        <%if cint(visAkt) = 1 then %>
                                        <td bgcolor="#ffffff" style="width:100px; padding:3px 3px 3px 3px;">Aktivitet</td>
										<%end if

										lastMonth = newMonth
										end if
										
										
										
										
				                        if y = 0 then
										
										statdatoM3 = dateadd("m", -3, startdato)
										statdatoM2 = dateadd("m", -2, startdato)
										statdatoM1 = dateadd("m", -1, startdato)
										datoMD_3 = datepart("m", dateadd("m", -3, startdato), 2,2) 
										datoMD_2 = datepart("m", dateadd("m", -2, startdato), 2,2) 
										datoMD_1 = datepart("m", dateadd("m", -1, startdato), 2,2) 
										datoAR_3 = datepart("yyyy",statdatoM3, 2,2)
										datoAR_2 = datepart("yyyy",statdatoM2, 2,2)
										datoAR_1 = datepart("yyyy",statdatoM1, 2,2)
                                        
										
										
												
												
												if skrivCsvFil = 1 then
												csvDatoerMD(35) = datoMD_3 
												csvDatoerMD(34) = datoMD_2 
												csvDatoerMD(33) = datoMD_1 
												
												csvDatoerAR(35) = datoAR_3 
												csvDatoerAR(34) = datoAR_2 
												csvDatoerAR(33) = datoAR_1 
												end if
												
												
										
										end if
				
				
					
				dbg = "#ffffff"
				%>
				<td bgcolor="<%=dbg%>" align=center style="padding:3px; width:80px;">
				<%
				

               
                 call datoogDivider(nday)


				select case periodeSel
				'case 1
				'tdvalThis = left(monthname(datepart("d", nday)), 3) & " " & right(newYear, 2)
				case 3
				tdvalThis = "<span style=""font-size:9px;"">"& left(monthname(datepart("m", nday,2,2)), 3) & " "& newYear &"</span><br>Uge: " & datepart("ww", nday, 2,2) 
				case else '6,12
				tdvalThis = left(monthname(datepart("m", nday,2,2)), 3) & " " & right(newYear, 2) '& "<br> Fm "& firstdayOfFirstWeekThis & "<br>Lm"& firstdayOfLastWeekThis & "<br>=="& dividerThisMth(month(nday))  'nday & " D:" & dividerThisMth(month(nday))'
                end select
				    
                
              
                 
                        
                
    
                     

                  
					'dividerThisMth(thisMth) + 1 

					if skrivCsvFil = 1 then
					csvDatoerUge(y) = datepart("ww", nday,2,2)
					csvDatoerMD(y) = datepart("m", nday,2,2)
					csvDatoerAR(y) = datepart("yyyy", nday,2,2)
                    end if
					
				     Response.write tdvalThis & "<br><span style=""color:#999999; font-size:9px;"">"& timerProcExt 
                    'Response.write "Y: "& y &" wdif: "& dividerThisMth(thisMth) &" firstThursdayOfFirstWeekThis: "& firstThursdayOfFirstWeekThis & " nday: "& nday 
                    '&"thisMth: "& thisMth &" antalD: "& lastDInMonth &"  1. 1-" & datepart("m", nday,2,2)  &"-"& datepart("yyyy", nday,2,2) &"</span>"
				
				
				%></span></b>
                </td>
				<%
				
				
			next
			
			
			if val <> 2 AND cint(visStatus) = 1 then
			%>
            <td bgcolor="#ffffff" style="padding:3px 3px 3px 3px; white-space:nowrap;">Budget job<br /><span style="font-size:10px; color:#999999;">timer, bel�b og rest.</span></td>
			<td bgcolor="#ffffff" style="padding:3px 3px 3px 3px; white-space:nowrap;">Forecast<br /><span style="font-size:10px; color:#999999;">ialt (uanset valgt periode)</span></td>
				<td bgcolor="#ffffff" style="padding:3px 3px 3px 3px; white-space:nowrap;">Realiseret<br /><span style="font-size:10px; color:#999999;">ialt (uanset valgt periode)</span></td>
                <td bgcolor="#ffffff" style="padding:3px 3px 3px 3px; white-space:nowrap;">Saldo forecast/real.<br /><span style="font-size:10px; color:#999999;">saldo  (uanset valgt periode)</span></td>
			<%end if
	
	end sub
	


	
	

	sub datoeroverskriftCSV(val)
	'*** Datoer ***
		
		
		
		y = -1
		for y = 0 TO numoffdaysorweeksinperiode 
				
				
				
				call antaliperiode(periodeSel, y, startdato, monthUse)
				
				
				
										
										
										
				                        if y = 0 then
										
										statdatoM3 = dateadd("m", -3, startdato)
										statdatoM2 = dateadd("m", -2, startdato)
										statdatoM1 = dateadd("m", -1, startdato)
										datoMD_3 = datepart("m", dateadd("m", -3, startdato), 2,2) 
										datoMD_2 = datepart("m", dateadd("m", -2, startdato), 2,2) 
										datoMD_1 = datepart("m", dateadd("m", -1, startdato), 2,2) 
										datoAR_3 = datepart("yyyy",statdatoM3, 2,2)
										datoAR_2 = datepart("yyyy",statdatoM2, 2,2)
										datoAR_1 = datepart("yyyy",statdatoM1, 2,2)
										
										
												
												
												if skrivCsvFil = 1 then
												csvDatoerMD(35) = datoMD_3 
												csvDatoerMD(34) = datoMD_2 
												csvDatoerMD(33) = datoMD_1 
												
												csvDatoerAR(35) = datoAR_3 
												csvDatoerAR(34) = datoAR_2 
												csvDatoerAR(33) = datoAR_1 
												end if
												
												
										
										end if
				
				
					
				
				
					
					
					csvDatoerUge(y) = datepart("ww", nday,2,2)
					csvDatoerMD(y) = datepart("m", nday,2,2)
					csvDatoerAR(y) = datepart("yyyy", nday,2,2)
					
                    '*** Timer eller proc extension
                    if cint(showasproc) = 1 then
                    timerProcExt = " %"
                    else
                    timerProcExt = " tim."
                    end if

                     if periodeSel = 3 then
                    datoOskrift = "Uge "& csvDatoerUge(y) &" - "& left(monthname(csvDatoerMD(y)), 3) &" "& csvDatoerAR(y)  

                    csvTxtTopMd = csvTxtTopMd & datoOskrift &" Forecast " & timerProcExt &";"
                    else
                    datoOskrift = left(monthname(csvDatoerMD(y)), 3) &" "& csvDatoerAR(y) 

                    csvTxtTopMd = csvTxtTopMd & datoOskrift &" Forecast "& timerProcExt &";"
                    end if
					

                       if cint(expvisreal) = 1 then
                        csvTxtTopMd = csvTxtTopMd & "Real. timer "& datoOskrift &";"
                       end if

                 call datoogDivider(nday)
				
				
			next
			
			
			
	end sub
	
	
                public jobStartKriThisM, datoIntervalThisM
                function stKriInterval612(y,jobStartKri)

                jobStartKri = "1-"& month(jobStartKri) &"-"& year(jobStartKri)
                jobStartKriThisM = dateAdd("m", y, jobStartKri)
                
                select case month(jobStartKriThisM)
                case "1","3","5","7","8","10","12"
                datoIntervalThisM = 30
                case 2
                    select case year(jobStartKri)
                    case "2000", "2004", "2008", "2012", "2016", "2020", "2024", "2028", "2032", "2036", "2040", "2044"
                    datoIntervalThisM = 28
                    case else
                    datoIntervalThisM = 27
                    end select
                case else
                datoIntervalThisM = 29
                end select

                end function


                function stKriInterval3(y, jobStartKri)


                if y = 0 then
                jobStartKriNormT = day(jobStartKri) &"-"& month(jobStartKri) &"-"& year(jobStartKri)
                end if
                jobStartKri = dateAdd("d", y*7, jobStartKriNormT)

                end function
	
	
	'*******************************
	'*** Sumline pr. medarbejder ***
    '*******************************
    public gtmedarblinier
    gtmedarblinier = 0
	sub medarbtotal

                if antalJobAktMedlinier = 0 OR len(trim(antalJobAktMedlinier)) = 0 then
                antalJobAktMedlinier = 1
                end if
	
				'** Skal total linie vises? ***
				timersum = 0
				timerTxt = ""
				for y = 0 TO numoffdaysorweeksinperiode
				
				        
				         '*** Afvigelse %
				        afvThis = 0 
				        if cdbl(medarbTotalTimer(y)) <= cdbl(medarbTimerReal(y)) then
						    if medarbTimerReal(y) <> 0 then
                            afvThis = 100 - (medarbTotalTimer(y)/medarbTimerReal(y)) * 100
                            else
                            afvThis = 0
                            end if
                        else
                            if medarbTotalTimer(y) <> 0 then
                            afvThis = 100 - (medarbTimerReal(y)/medarbTotalTimer(y)) * 100
                            else
                            afvThis = 0
                            end if
                        end if
				        
				
				timersum = timersum + medarbTotalTimer(y)


                if periodesel = "6" OR periodesel = "12" then
                
                call stKriInterval612(y,jobStartKri)
                
                'response.write "<hr>" & mTotthisMid & "<br><br>"
                'response.flush

                call normtimerPer(mTotthisMid, jobStartKriThisM, datoIntervalThisM, 1)
                normtimerperThisM = formatnumber(ntimPer, 0) 
                else
        

                    'response.write "y: "& y & "<br>" 
                
                'call stKriInterval3(y,jobStartKri)
                if y = 0 then
                jobStartKriNormT = day(jobStartKri) &"-"& month(jobStartKri) &"-"& year(jobStartKri)
                end if
                jobStartKri = dateAdd("d", y*7, jobStartKriNormT)


                call normtimerPer(mTotthisMid, jobStartKri, 6, 1) 
                normtimerperThisM = formatnumber(ntimPer, 1)
          
                end if 

                belagProcM = 0
                if normtimerperThisM <> 0 AND medarbTotalTimer(y) <> 0 then
                    if medarbTotalTimer(y) < normtimerperThisM then
                    belagProcM = (medarbTotalTimer(y) / normtimerperThisM ) * 100
                    else
                    belagProcM = (normtimerperThisM / medarbTotalTimer(y)) * 100
                    end if
                else
                belagProcM = 0
                end if


				timerTxt = timerTxt & "<td valign=bottom style='padding:2px 2px 2px 2px; color:#5582d2; font-size:10px; line-height:12px;'><span style='color:#999999;'>n: "& normtimerperThisM &" t.</span>"

                resttimerThisM = 0
                resttimerThisM = (normtimerperThisM - medarbTotalTimer(y))
                
                
                if medarbTotalTimer(y) <> 0 then 
               
                if cint(vis_simpel) = 2 then

                thishgt = formatnumber(medarbTotalProc(y)/ 2, 0) 

                if cdbl(thishgt) < 10 then
                thisBgc = "#DCF5BD" 
                end if

                if cdbl(thishgt) >= 10 AND cdbl(thishgt) < 25 then
                thisBgc = "yellowgreen" 
                end if

                 if cdbl(thishgt) >= 25 AND cdbl(thishgt) < 50 then
                thisBgc = "#6CAE1C" 
                end if

                 if cdbl(thishgt) >= 50 then
                thisBgc = "red" 
                end if

                if cdbl(thishgt) < 10 then
                 thishgt = 10
                end if

                timerTxt = timerTxt &"<div style='height:"& thishgt &"px; background-color:"& thisBgc &"; padding:2px; color:#000000; line-height:10px;'>" 
                timerTxt = timerTxt &"f: <b>"& formatnumber(medarbTotalTimer(y), 0)&"</b> t. <span style=""font-size:8px; color:#000000""> ("& formatnumber(belagProcM, 0) &"%) </span>" 
                if medarbTotalProc(y) > 50 then 'vis saldo n�r der er brugt mere end 50%
                timerTxt = timerTxt &"<br><u>s: "& formatnumber(resttimerThisM, 0)&" t.</u>"
                end if
                timerTxt = timerTxt &"</div>"


                else

                timerTxt = timerTxt &"<br>f: "& formatnumber(medarbTotalTimer(y), 0)&" t. "
                
                if belagProcM <> 0 then
                timerTxt = timerTxt &"<span style=""font-size:9px; color:#999999;""> ("& formatnumber(belagProcM, 0) &"%) </span>"
                end if

                timerTxt = timerTxt &"<br><u>s: "& formatnumber(resttimerThisM, 0)&" t. </u>"

                end if

               

                end if

                afvThiMtot = 0
                if medarbTimerReal(y) <> 0 then
                afvThiMtot = (medarbTotalTimer(y) / medarbTimerReal(y)) * 100
                end if
                
                if medarbTimerReal(y) <> 0 AND media <> "print" then

                 
                 if cint(vis_simpel) <> 2 AND periodesel <> 3 then
                 timerTxt = timerTxt &"</br>"
                 timerTxt = timerTxt &" r: "&formatnumber(medarbTimerReal(y), 0)&" t. "
                
                 if formatnumber(afvThiMtot,0) <> 0 AND formatnumber(afvThiMtot,0) <> 100 then
                 'timerTxt = timerTxt &" afv.: "& formatnumber(afvThiMtot,0) &"%"
                 end if

                 end if
                
                else
                timerTxtJob = timerTxtJob & "&nbsp;"
                end if
				
                timerTxt = timerTxt & "</td>"

                
				
                medarbTotalProc(y) = 0
				medarbTotalTimer(y) = 0 
				medarbTimerReal(y) = 0
				next
				
			 if timersum <> 0 then
                
                     
                       saldoTimerJobtotGrand = (forecastTimerJobtotGrand-realTimerJobtotGrand)

                       saldoTimerTotJobMedignPerGrand = (forcastTimerTotJobMedignPerGrand - realTimerTotJobMedignPerGrand)


                
                %><!-- medarb total -->

               <%
               if vis_simpel <> 2 then    
               bgmTot = "#d6dff5"
               bgmfnt = "#5582d2"
               else
               bgmTot = "#ffffff"
               bgmfnt = "#000000"
                end if  %>
					<tr class="xmedtotal" bgcolor="<%=bgmTot %>">	
                       
						<td height=20 align=right valign="top" style="padding:2px 5px 2px 2px; color:<%=bgmfnt%>;">Total <%=mNavnTot %><br /><span style="font-size:9px;">Norm, Real.</font></td>
                         <td style="padding:2px 5px 2px 5px;" valign="top">

                             <%if cint(vis_simpel) = 2 then %>
                             <span id="sp_medarbjoblist_<%=lastxmid%>" class="sp_medarbjoblist" style="color:<%=bgmfnt%>;"><b>[+] Vis alle</b></span> <%=thisMedarbJoblist %>
                             <%else %>
                             &nbsp;
                             <%end if %>

                         </td>

                        <%if cint(visRamme) = 1 then %>
                        <td valign=top style="padding:2px 5px 2px 2px; color:<%=bgmfnt%>"><%=formatnumber(thisMTotRamme,0) %><br />&nbsp;</td>
                        <%end if %>

                          <%if cint(visAkt) = 1 then %>
                        <td>&nbsp;</td>
					    <%end if

						Response.write timerTxt


                        if cint(visStatus) = 1 then
						%>
                        <td>&nbsp;</td>
						<td valign=top align=right style="padding:2px 5px 2px 2px; color:<%=bgmfnt%>"><b><%=formatnumber(forecastTimerJobtotGrand,0) %></b> <!-- (<%=formatnumber(forecastProcJobtotGrand/(dividerAll*antalJobAktMedlinier),0) %>%) --> <!--<%=forecastProcJobtotGrand&"/"& dividerAll &"**"& antalJobAktMedlinier %>--><br />
                        <span style="font-size:9px; color:#999999;"><%=formatnumber(forcastTimerTotJobMedignPerGrand, 0) %></span></td>
					    <td valign=top align=right style="padding:2px 5px 2px 2px; color:#5582d2;"><%=formatnumber(realTimerJobtotGrand,2) %><br />
                         <span style="font-size:9px; color:#999999;"><%=formatnumber(realTimerTotJobMedignPerGrand, 2) %></span></td>

                         <td valign=top align=right style="padding:2px 5px 2px 2px; color:#5582d2;"><%=formatnumber(saldoTimerJobtotGrand,2) %><br />
                         <span style="font-size:9px; color:#999999;"><%=formatnumber(saldoTimerTotJobMedignPerGrand, 2) %></span></td>

                        <%end if %>

						</tr>
						<%
						'Response.flush
                        forecastProcJobtotGrand = 0
						forecastTimerJobtotGrand = 0
						realTimerJobtotGrand = 0
                        thisMTotRamme = 0

                        forcastTimerTotJobMedignPerGrand = 0
                        realTimerTotJobMedignPerGrand = 0

                        '*** skifter medarbejder og derfor skal de nulstilles ***'
                        forecastTimerMJobtotGrand = 0
		                realTimerMJobtotGrand = 0 
                        
                        saldoTimerJobtotGrand = 0
                        saldoTimerTotJobMedignPerGrand = 0 

                        antalJobAktlinierGrand = antalJobAktlinierGrand + antalJobAktMedlinier
                        antalJobAktMedlinier = 0 

                        thisMedarbJoblist = ""

                        gtmedarblinier = gtmedarblinier + 1 
                        

				end if 'vis timersum (sub totaler)
				%>
	
	<%
	end sub

    
    
    '*************************************************
    '*** Sumline pr. job indenfor hver medarbejder ***
    '*************************************************
    public forcastTimerTotJobMedignPerGrand, realTimerTotJobMedignPerGrand
    public realignPerGrandGrand, realGrandGrand, fcignPerGrandGrand, fcGrandGrand, fpGrandGrand
    public saldoTimerMJobtotGrand, saldoTimerTotJobMedignPer, realTimerTotJobMedignPer, jobBgtGrandGrand, jobTimerGrandGrand

    
	sub jobtotalprmedarb
	
				'** Skal total linie vises? ***
				timersumJob = 0
				timerTxtJob = ""
				for y = 0 TO numoffdaysorweeksinperiode
				
				        
				         '*** Afvigelse %
				        afvThisJob = 0 
				        if (cdbl(timerMJobTotal(y)) <= cdbl(realMJobTotal(y))) AND realMJobTotal(y) <> 0 then
						    if timerMJobTotal(y) <> 0 then
                            afvThisJob = 100 - (timerMJobTotal(y)/realMJobTotal(y)) * 100
                            else
                            afvThisJob = 0
                            end if
                        else
                            if timerMJobTotal(y) <> 0 then
                            afvThisJob = 100 - (realMJobTotal(y)/timerMJobTotal(y)) * 100
                            else
                            afvThisJob = 0
                            end if
                        end if
				        
				
				timersumJob = timersumJob + timerMJobTotal(y)
				timerTxtJob = timerTxtJob & "<td bgcolor=snow style='padding:2px 2px 2px 2px; font-size:10px; color:#999999;'>"
                
                if timerMJobTotal(y) <> 0 then
                timerTxtJob = timerTxtJob & "f: " & formatnumber(timerMJobTotal(y), 0) & " t. "' ("& formatnumber(procMJobTotal(y), 0) & "%)" ' antanlJoblinierprM 'antalJobAktMedlinier 
                end if
                
                if realMJobTotal(y) <> 0 AND media <> "print" then
                timerTxtJob = timerTxtJob &"<br><span style=""font-size:10px;"">r: "&formatnumber(realMJobTotal(y), 0)&" t. " 'afv.: "& formatnumber(afvThisJob,0) &"% </span>"
                else
                timerTxtJob = timerTxtJob & "<br>&nbsp;"
                end if
				
                timerTxtJob = timerTxtJob & "</td>"
            
                fpGrandtotal(y) = fpGrandtotal(y) + procMJobTotal(y)
                fcGrandtotal(y) = fcGrandtotal(y) + timerMJobTotal(y)
                realGrandtotal(y) = realGrandtotal(y) + realMJobTotal(y)
                

				timerMJobTotal(y) = 0
				realMJobTotal(y) = 0
                procMJobTotal(y) = 0
				next

                if cint(vis_simpel) <> 1 AND cint(vis_simpel) <> 2 then
				
				if timersumJob <> 0 then%>
					<tr class="xjobtotal">
                      
						<td height=20 bgcolor="snow" align=right style="padding:2px 5px 2px 2px; color:#999999;">Total forecast job: <span style="font-size:9px;"><br /> Real.</span></td>
                            <td bgcolor="snow">&nbsp;</td>

                          <%if cint(visRamme) = 1 then %>
                        <td bgcolor="snow">&nbsp;</td>
                        <%end if %>

                          <%if cint(visAkt) = 1 then %>
                        <td bgcolor="snow">&nbsp;</td>
                        <%end if %>

                    
					    <%
						Response.write timerTxtJob
						 
                      

                        saldoTimerMJobtotGrand = (forecastTimerMJobtotGrand - realTimerMJobtotGrand)
                        saldoTimerTotJobMedignPer = (forcastTimerTotJobMedignPer - realTimerTotJobMedignPer)   
                        
                        if cint(visStatus) = 1 then%>
                        <td bgcolor="snow" align=right class=lille style="padding:2px 2px 2px 2px;">&nbsp;</td>


						<td bgcolor="snow" align=right class=lille style="padding:2px 2px 2px 2px; color:#999999;"><%=formatnumber(forecastTimerMJobtotGrand,0) %> <!-- (<%=formatnumber(forecastProcMJobtotGrand/(dividerAll*antalJobAktlinier),0) %>%) <%=forecastProcMJobtotGrand &"/"& dividerAll &"*"& antalJobAktlinier %>--><br />
                        <span style="font-size:9px; color:#999999;"><%=formatnumber(forcastTimerTotJobMedignPer, 0) %></span></td>
					    <td bgcolor="snow" class=lille align=right style="padding:2px 2px 2px 2px; color:#999999;"><%=formatnumber(realTimerMJobtotGrand,2) %><br />
                        <span style="font-size:9px; color:#999999;"><%=formatnumber(realTimerTotJobMedignPer, 2) %></span></td>

                         <td bgcolor="snow" class=lille align=right style="padding:2px 2px 2px 2px; color:#999999;"><%=formatnumber(saldoTimerMJobtotGrand,2) %><br />
                        <span style="font-size:9px; color:#999999;"><%=formatnumber(saldoTimerTotJobMedignPer, 2) %></span></td>

                        <%end if %>


						</tr>
						<%
						

                        'fcGrandGrand = fcGrandGrand + forecastTimerMJobtotGrand
                        'fcignPerGrandGrand = fcignPerGrandGrand + forcastTimerTotJobMedignPer
                        'realGrandGrand = realGrandGrand + realTimerMJobtotGrand 
                        'realignPerGrandGrand = realignPerGrandGrand + realTimerTotJobMedignPer

               

                        forecastProcMJobtotGrand = 0
						forecastTimerMJobtotGrand = 0
						realTimerMJobtotGrand = 0
                        forcastTimerTotJobMedignPer = 0
                        realTimerTotJobMedignPer = 0

                        saldoTimerMJobtotGrand = 0
                        saldoTimerTotJobMedignPer = 0
               
                  end if
				end if            
                                    
                        antalJobAktMedlinier = antalJobAktMedlinier + antalJobAktlinier
                        antalJobAktlinier = 0
               
				%>
	
	<%
	end sub
	
	

    '*** Grand-Grandtotal ***
	sub grandtotal
	
     

				'** Skal total linie vises? ***
				for y = 0 TO numoffdaysorweeksinperiode
				
				if gtmedarblinier <> 0 then
                gtmedarblinier = gtmedarblinier
                else 
                gtmedarblinier = 1
                end if   
				        
				
				
				timerTxt = timerTxt & "<td bgcolor=#F7F7F7 valign=top class=lille style='padding:2px 2px 2px 2px'>"
                
                if fcGrandtotal(y) <> 0 then
                timerTxt = timerTxt & "<b>f: "&formatnumber(fcGrandtotal(y), 0)&"</b> t. "' (" &formatnumber(fpGrandtotal(y)/gtmedarblinier, 0)&"%)"
                end if
                
                if realGrandtotal(y) <> 0 AND media <> "print" then
                timerTxtJob = timerTxtJob &"<br><span style=""font-size:9px;"">r: "&formatnumber(realGrandtotal(y), 0)&" t.</span>"
                else
                timerTxtJob = timerTxtJob & "<br>&nbsp;"
                end if

				
                timerTxtJob = timerTxtJob & "</td>"

                timerTxt = timerTxt & timerTxtJob

                timerTxtJob = ""
            	next
				
			    
                saldoGrandGrand = (fcGrandGrand - realGrandGrand)
                saldoignPerGrandGrand = (fcignPerGrandGrand - realignPerGrandGrand)
                
                %>
					<tr>	
                      
						<td height=20 bgcolor="#F7F7F7" align=right style="padding:2px 5px 2px 2px;"><b>Grandtotal:</b><font class=megetlillesort><br /> Real.</font></td>
                          <td bgcolor="#F7F7F7">&nbsp;</td>

                        <%if cint(visRamme) = 1 then %>
                        <td bgcolor="#F7F7F7">&nbsp;</td>
                        <%end if %>

                        <%if cint(visAkt) = 1 then %>
                        <td bgcolor="#F7F7F7">&nbsp;</td>
                        <%end if %>
					    
                        <%
						Response.write timerTxt
						%>


                        <%if cint(visStatus) = 1 then
                            
                           if antalJobAktlinierGrand = 0 then
                            antalJobAktlinierGrand = 1 
                            end if %>
                        <td bgcolor="#F7F7F7" align=right style="padding:2px 2px 2px 2px;"><%=formatnumber(jobTimerGrandGrand,0) %><br />
                        <span style="font-size:9px; color:#999999;"><%=formatnumber(jobBgtGrandGrand, 0) %> DKK</span></td>

						<td bgcolor="#F7F7F7" valign=top align=right style="padding:2px 5px 2px 2px;"><b><%=formatnumber(fcGrandGrand,0) %></b> <!--  (<%=formatnumber(fpGrandGrand/(dividerAll*antalJobAktlinierGrand),0) %>%) --><br />
                        <span style="font-size:9px; color:#999999;"><%=formatnumber(fcignPerGrandGrand, 0) %></span></td>
					    <td bgcolor="#F7F7F7" valign=top align=right style="padding:2px 5px 2px 2px;"><%=formatnumber(realGrandGrand,2) %><br />
                         <span style="font-size:9px; color:#999999;"><%=formatnumber(realignPerGrandGrand, 2) %> *</span></td>

                            <td bgcolor="#F7F7F7" valign=top align=right style="padding:2px 5px 2px 2px;"><%=formatnumber(saldoGrandGrand,2) %><br />
                         <span style="font-size:9px; color:#999999;"><%=formatnumber(saldoignPerGrandGrand, 2) %></span></td>
                        <%end if %>
				   </tr>
						<%
						
				%>
	
	<%
	end sub
	
	

	
	
	'*** Function periodeantal ***
	public newMonth, nday, newYear, newWeek
	function antaliperiode(peri, ycount, sdato, mduse)
	            
	            
	            
	                if ycount = 0 OR ycount = -2  then
					    nDay = dateadd("d", 0, sdato)
					else
					    select case periodeSel
				        case 1
				        nDay = dateAdd("d",ycount, sdato)
				        case 3
				        nDay = dateAdd("ww",ycount, sdato)
				        case 6,12
				        nDay = dateAdd("m",ycount, sdato)
				        end select
					end if
					
			    
					'Response.Write "nDay" & nDay
				
					
					newWeek = datepart("ww", nDay, 2,2)
					newMonth = datepart("m", nday, 2,2)
					newYear = datePart("yyyy", nDay, 2,2)
					
					
				
	
	end function
	
	
	
	
	
	
	
	
	
	
	
	
	
	function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
	end function
	
	
	function tilfojNylinie(medarbSel, visRamme, visAktiv)
	
        
        %>
	
		<!--<br>
				<table cellspacing=0 cellpadding=0 border=0 width=<%=leftwdt %>>
				<tr><td align=right><input type="submit" value="Indl�s forecast!"></td></tr>
				</table>
				-->
				
				
				
				
				<%
				skrivCsvFil = 0
				'call datoeroverskrift
				
                if medarbSel <> "" AND medarbSel <> 0 then
				usemrn = medarbSel
                else
                usemrn = session("mid")
                end if
				
				if medarbSel <> 0 then
					call hentbgrppamedarb(medarbSel)
				else
					strPgrpSQLkri = ""
				end if
				
				
				
				'******************************************************
				'**** Henter job fra Guiden Dine aktive job ***********
				'******************************************************
				
                varUseJob = "("

                

				strSQL4 = "SELECT tu.id, tu.medarb, tu.jobid FROM timereg_usejob AS tu "_
                &" LEFT JOIN job AS j ON (j.id = tu.jobid)"_
                &" WHERE "& varAktivejob &" tu.medarb = "& usemrn &" AND tu.forvalgt = 1 AND j.risiko > -1 GROUP BY tu.jobid" 'AND ("& jobSQLkri &")

			

                'if lto = "essens" then
			    'Response.Write strjIdWrt &"<br>"& strSQL4 &" .."& jobSQLkri &"<br><br>"
				'Response.flush
                'end if
				
                tuJobisWrt = ""
				oRec3.open strSQL4, oConn, 3
				
				while not oRec3.EOF 
                    
                 '** strjIdWrt = Job med forecast p� i periode / eller Real timer **'
                   'if (instr(strjIdWrt, "#"&oRec3("jobid")&"#") = 0) then

                   if instr(tuJobisWrt, "#"&oRec3("jobid")&"#") = 0 then

                   '** Vis kun de job i guiden der ikke allerede findes ressource timer p� I VALGTE PERIODE ***'
                   strSQL5 = "SELECT jobid FROM ressourcer_md AS rmd WHERE rmd.medid = "& usemrn & " AND rmd.jobid = "& oRec3("jobid") &" AND "_
                   &" ((rmd.md >= "& startdatoMD &" AND rmd.aar = "& startdatoYY &") "& orandval &" (rmd.md <= "& slutdatoMD &" AND rmd.aar = "& slutdatoYY &"))"_
                   &" GROUP BY jobid"
				  
                  'Response.write strSQL5 & "<br><br>"
                  'Response.flush


                         
               

                   oRec5.open strSQL5, oConn, 3
                   
                   if not oRec5.EOF then
                   varUseJob = varUseJob
                   else
                   varUseJob = varUseJob & " j.id = "& oRec3("jobid") & " OR "
                   tuJobisWrt = tuJobisWrt & "#"& oRec3("jobid") &"#"
                   end if
				   
                   oRec5.close
                   '******* 
                   'end if


                  end if 'tuJobisWrt
                oRec3.movenext
				wend 
				
				oRec3.close 

                   

              
				
				if varUseJob = "(" then
				varUseJob = " j.id = 0 AND "
				else
				varUseJob_len = len(varUseJob)
				varUseJob_left = varUseJob_len - 3
				varUseJob_use = left(varUseJob, varUseJob_left) 
				varUseJob = varUseJob_use & ") AND "
				end if



                select case lto 
                case "commutemedia", "intranet - local"
                jobstSQLkri = " (j.jobstatus = 1 OR j.jobstatus = 3) "
                case else
                jobstSQLkri = " j.jobstatus = 1 "
                end select

                jobstSQLkri = jobstSQLkri &" AND j.risiko > -1"
				
                
				
				'************** Main SQL Call (Ikke angivne ressource timer) *********************'
				'**Kun hvis der ikke er s�gt p� et specifikt jobnr

               
                strSQL = "SELECT j.id AS id, jobnr, jobnavn, jobknr, kkundenavn, kkundenr, count(a.id) AS antalakt,"_
				&" j.budgettimer, j.ikkebudgettimer, "_
				&" kid, jobans1, jobans2, j.jobstatus FROM job j"_
                &" LEFT JOIN kunder k ON (k.Kid = j.jobknr)"_
				&" LEFT JOIN aktiviteter a ON (a.job = j.id)"
				
				strSQL = strSQL &" WHERE ("& varUseJob &" "& jobstSQLkri 

                'strSQL = strSQL &" AND j.risiko > -1"
				
				'if jobidsel <> 0 then
				'strSQL = strSQL &" AND j.id = "& jobidsel &""
				'end if
				
				strSQL = strSQL &" AND k.Kid = j.jobknr "& jobNavnKri &") "& strPgrpSQLkri & " GROUP BY j.id, a.job "_
				&" ORDER BY k.kkundenavn, j.jobnavn, j.jobnr" 
				

                  
              

                %>

              
				
				
				
				
				<tr id="newline_<%=usemrn%>" style="left:5px; top:5px; visibility:hidden; display:none; background-color:#FFFFe1; padding:3px; border:1px #999999 solid;">
                   
				<td colspan="2" style="padding:2px 5px 4px 5px;"><b>Tilf�j nyt forecast:</b>
				
                 
				<br /><select id="selFM_jobid_<%=usemrn%>" name="selFM_jobid" style="width:365px;" onChange="setjobidval('<%=usemrn%>')" class="tilfoj_ny_job">
                        <option value=0>V�lg job .. (fra pers. aktivliste, uden forecast)</option>
                 <%
                 
                
                               
								
								 
								j = 0
								firstJsel = 0
                                lastKnavn = ""
								oRec.open strSQL, oConn, 3
                                while not oRec.EOF 
								
						
								
								'if instr(strjIdWrt, "#"&oRec("id")&"#") = 0 then
								
								if j = 0 then
								firstJsel = oRec("id")
								end if
								
								'if cint(jobidSel) = cint(oRec("id")) then
								'jobselected = "SELECTED"
								'else
								jobselected = ""
								'end if
								
								
								'delit = "...................................."
                                
                                jobnavnid = replace(left(oRec("jobnavn"), 25), "'", "") & " ("& oRec("jobnr") &")"
                                
                                'len_jobnavnid = len(jobnavnid)
                                'if len(len_jobnavnid) < 25 then
                                'lenst = 25
                                'else
                                'lenst = len_jobnavnid + 2
                                'end if 
                                
                                'if len(lenst) <> 0 AND len(len_jobnavnid) <> 0 AND (lenst-len_jobnavnid) > 0 then 
                                'delit_antal = left(delit, (lenst-len_jobnavnid)) 
                                'else
                                'delit_antal = "........"
                                'end if

                                if oRec("jobstatus") = 3 then
                                jobnavnid = jobnavnid & " (tilbud) "
                                end if
								
								if (cint(session("mid")) = cint(usemrn) OR level = 1 OR cint(session("mid")) = cint(oRec("Jobans1")) OR cint(session("mid")) = cint(oRec("Jobans2"))) then


                                if lastKnavn <> oRec("kkundenavn") then
                                if j <> 0 then%>
                                <option value="<%=oRec("id") %>" disabled></option>
                                <%end if %>

                                <option value="<%=oRec("id") %>" disabled><%=oRec("kkundenavn") &" ("& oRec("kkundenavn")&")"%></option>
                                <%
                                end if
		                        %>
								<option value="<%=oRec("id") %>" <%=jobselected %>><%=jobnavnid%></option>
				                <!--<option><%=(cint(session("mid")) &"="& cint(usemrn) &" OR level = 1 OR "& cint(session("mid")) &"="& cint(oRec("Jobans1")) &"OR"& cint(session("mid")) &"="& cint(oRec("Jobans2"))) %> </option>-->
                                <%
                                    lastKnavn = oRec("kkundenavn")

                          
				                j = j + 1
				                
				                end if
								
								'end if 'instr
								
								
								oRec.movenext
								wend
								oRec.close


                                  

                            
								
								if j = 0 then%>
								<option value=0>Ingen aktive job, hvor du har adgang til</option>
								<option value=0>at angive forecast for den valgte medarbejder.</option>
                                <option value=0>Eller forecast er allerede angivet.</option>
                                
                                

                                <%end if %>
								
								</select>

                            
                           
                        

                                 &nbsp;&nbsp;<a href="#" id="hnl_<%=usemrn%>" class="red">[X]</a>
                                <%if (cint(session("mid")) = cint(usemrn) OR level = 1) then%>
				            <br /><a href="javascript:popUp('guiden_2006.asp?mid=<%=usemrn%>&lc=res','700','500','150','120');" target="_self"; class=vmenu>Personlig aktiv jobliste >> </a>
				            <%end if %>
								</td>

                             

                                <%if cint(visRamme) = 1 then %>
                                <!-- Ramme -->
                                <td valign=top style="padding:2px 2px 2px 2px;">&nbsp;</td>
                                <%end if %>


                                <%if cint(visAktiv) = 1 then %>
                                <!-- Aktiviteter --->
                                <td valign=top style="padding:2px 2px 2px 2px;">
                                    <!--<textarea id="tdtestFM_jobid_<%=usemrn%>"></textarea>tdtestFM_jobid_<%=usemrn%>-->
                                <select id="tdFM_jobid_<%=usemrn%>" class="nlFM_jobid" name="sFM_aktid" style="width:100px;"><option>..</option></select>
                                &nbsp;</td>
                                <%end if %>
								
								
								            <%
											for y = 0 TO numoffdaysorweeksinperiode
													
													call antaliperiode(periodeSel, y, startdato, monthUse)
													lastMonth = newMonth
											%>	
													
											<td valign=top style="padding:2px 2px 2px 2px;">
											<input type="hidden" name="FM_jobid" id="FM_jobid_<%=usemrn%>_<%=y%>" value="<%=firstJsel%>">
                                            <input class="hdFM_jobid_<%=usemrn%>" type="hidden" name="FM_aktid" id="Hidden1" value="0">
                                            <input type="hidden" name="FM_aktid_old" id="Hidden13" value="0">
											
											<input type="hidden" name="FM_medarbid" id="Hidden2" value="<%=usemrn%>">
											<input type="hidden" name="FM_dato" id="Hidden3" value="<%=nday%>">
											
											
											
											<%if (month(cdate(dagsdato)) <= month(cdate(nday)) AND year(cdate(dagsdato)) = year(cdate(nday))) OR (year(cdate(dagsdato)) < year(cdate(nday))) OR level = 1 then %>
											<input type="text" name="FM_timer" id="FM_timer_<%=usemrn%>_0_<%=y %>" style="width:50px; font-size:10px;" value="">
											<!--<input id="Button1" type="button" value=">>" style="font-size:8px;" onclick="copyTimer('<%=usemrn%>','0', '<%=y %>')" />-->
                                                <span class="btn_timer_kopy" style="padding:1px; background-color:#CCCCCC; font-size:8px;" onclick="copyTimer('<%=usemrn%>','0', '<%=y %>')">>></span>

											<%else %>
											<input type="hidden" name="FM_timer" id="Text1" value="">
											<%end if %>
											
											<input type="hidden" name="FM_timer" id="Hidden4" value="#">
											</td>
											<%next%>
											
                                            <%if cint(visStatus) = 1 then%>
                                            <td style="background-color:#FFFFFF;">&nbsp;</td>
											<td style="background-color:#FFFFFF;">&nbsp;</td>
                                            <td style="background-color:#FFFFFF;">&nbsp;</td>
											<td style="background-color:#FFFFFF;">&nbsp;</td>
                                            <%end if%>
	
	</tr>
	
	<!--
	</table>
	<table cellspacing=0 cellpadding=0 border=0 width=100%>
		<tr><td align=right style="padding:10px 0px 5px 0px;">
		<input type="submit" value="Indl�s forecast >> "></td></tr>
</table>
	</div>
	-->
	<%
	
           


	end function


    function tilfojNyliniePaaJob(medidNL, jobid)
	
	
		usemrn = medidNL
				
				
				
				
				
				
				
				'******************************************************
				'**** Henter job  ***********
				'******************************************************
				
				
				strSQLnl = "SELECT id, jobnavn, jobnr FROM job WHERE id = "& jobid
				'Response.Write strSQLnl
                'Response.flush
                jobidNL = 0
               
				oRec4.open strSQLnl, oConn, 3 
				if not oRec4.EOF then
				jobnavnid = replace(left(oRec4("jobnavn"), 25), "'", "") & " ("& oRec4("jobnr") &")"
                jobidNL = oRec4("id") 
                end if
				oRec4.close
				
				%>
				
				
				<tr id="nylinjepaajob_<%=jobid %>_<%=medidNL %>" style="left:5px; top:5px; visibility:hidden; display:none; background-color:#FFFFE1; padding:3px; border:1px #999999 solid;">
                <td bgcolor="#FFFFFF" >&nbsp;</td>

                    
				<td style="padding:2px 5px 4px 5px; font-size:11px;">Tilf�j nyt forecast p�<br />
                <b><%=jobnavnid %>:</b>&nbsp;<a href="#" id="jnl_<%=jobid %>_<%=medidNL%>" class="red_j" style="font-size:11px; color:red;">[X]</a>
				<br />
                <input id="Hidden11" name="selFM_jobid" value="<%=jobidNL%>" type="hidden" />
                 <%
                 
                
                               
								
								
								
								
								
								%>
								</td>
                                <!-- Ramme -->
                                <%if cint(visRamme) = 1 then %>
                                <td valign=top style="padding:2px 2px 2px 2px;">&nbsp;</td>
                                <%end if %>
                    
                 


                                <!-- Aktiviteter --->
                     <%if cint(visAkt) = 1 then %>          
                     <td valign=top style="padding:2px 2px 2px 2px;">
                    <%
                     
                       

                  
                     call hentaktiviteter(positiv_aktivering_akt_val, medarbKundeoplysX(x, 0), medarbKundeoplysX(x, 3), aty_sql_realhoursAkt)
                 
                  

                    aSel = ""
                    aktText = ""
                   %>
                    <select class="aaNLFM_jobid" id="aaNLFM_jobid_<%=jobid %>_<%=medidNL %>" name="sFM_aktid" style="width:100px;">
                    <option value="0">(uspecificeret)</option>
                    <option value="0">V�lg aktivitet..?</option>
                    <%
                  
                    oRec4.open strSQLa, oConn, 3
                    While not oRec4.EOF 
        
                      
                    
                                       

                     if ISNULL(oRec4("fase")) <> true AND len(trim(oRec4("fase"))) <> 0 then
                     fsNavn = " | fase: "& oRec4("fase")
                     else
                     fsNavn = ""
                     end if

                     if cdbl(medarbKundeoplysX(x, 17)) = cdbl(oRec4("aid")) then
                     aSel = "SELECTED"
                     aktText = oRec4("aktnavn") & fsNavn
                     else
                     aSel = ""
                     end if

                     
                     
                      if media <> "print" AND media <> "eksport" then%>
                    <option value="<%=oRec4("aid")%>" <%=aSel %>><%=oRec4("aktnavn") &" "& fsNavn%></option>
                    <%end if

                      
                       
                    
                    oRec4.movenext
                    wend
                    oRec4.close 
                    
                 

                    %>
        
                    </select>
                    
                    </td>
					<%end if
								
											
											
											
											for y = 0 TO numoffdaysorweeksinperiode
													
											call antaliperiode(periodeSel, y, startdato, monthUse)
											lastMonth = newMonth
											%>	
													
											<td valign=top style="padding:2px 2px 2px 2px;">
											<input type="hidden" name="FM_jobid" id="Hidden5" value="<%=jobidNL%>">
                                            <input class="ahNLFM_jobid_<%=jobid %>_<%=medidNL %>" type="hidden" name="FM_aktid" id="Hidden6" value="0">
                                            <input type="hidden" name="FM_aktid_old" id="Hidden12" value="0">
											<input type="hidden" name="FM_medarbid" id="Hidden7" value="<%=medidNL%>">
											<input type="hidden" name="FM_dato" id="Hidden8" value="<%=nday%>">
											
											
											
											<%if (month(cdate(dagsdato)) <= month(cdate(nday)) AND year(cdate(dagsdato)) = year(cdate(nday))) OR (year(cdate(dagsdato)) < year(cdate(nday))) OR level = 1 then %>
											<input type="text" name="FM_timer" id="FM_timer_<%=medidNL %>_<%=jobid %>_<%=y %>" style="width:50px;" value="">
											<span class="btn_timer_kopy" style="padding:1px; background-color:#CCCCCC; font-size:8px;" onclick="copyTimer('<%=medidNL%>','<%=jobid %>', '<%=y %>')">>></span>

											<%else %>
											<input type="hidden" name="FM_timer" id="Hidden9" value="">
											<%end if %>
											
											<input type="hidden" name="FM_timer" id="Hidden10" value="#">
											</td>
											<%next%>

                                            <%if cint(visStatus) = 1 then %>
											<td bgcolor="#FFFFFF" >&nbsp;</td>
											<td bgcolor="#FFFFFF" >&nbsp;</td>
	                                        <td bgcolor="#FFFFFF" >&nbsp;</td>
                                            <td bgcolor="#FFFFFF" >&nbsp;</td>
                                            <%end if %>
	</tr>
	
	
	<%
	end function
	
    
    
    public arrsNormThis, arrsIdThis
    sub arsRammeSub

     
               arrsNormThis = 0
               arrsIdThis = 0
               strSQLram = "SELECT timer, id FROM ressourcer_ramme WHERE jobid = "& medarbKundeoplysX(x, 0) &" AND medid = "& medarbKundeoplysX(x, 3) &" AND aar = " & aarsRamme
               'Response.Write strSQLram
               'Response.flush
               oRec6.open strSQLram, oConn, 3
               if not oRec6.EOF then
               arrsNormThis = formatnumber(oRec6("timer"), 2)
               arrsIdThis = oRec6("id")
               end if
               oRec6.close 
               
                
               

    end sub

   
    public mNavnTot 
    function medarbejderLinje(mid,jobStartKri,datoInterval,rdimnb, vis)
    
    
    call meStamdata(mid)
        
       mNavnTot = meNavn

       if cint(vis) = 0 then%>
       <tr bgcolor="#F7F7F7" class="tr_medarb" style="visibility:visible; display:;"><!-- 5C75AA-->
       <%end if %>

        
		<td style="padding:2px 3px 3px 5px; white-space:nowrap; background-color:#F7F7F7;">
            <!--
		    <if vis_simpel = 2 then %>
            <span class="sp_medarbjoblist" id="sp_medarbjoblist_<=mid%>" style="color:#5582d2;"><b>[-]</b></span>
            <end if %>
            -->
            
            <span style="font-size:14px; line-height:18px;"><b><%=meTxt%></b></span>
		
		
		
		
		<!--<a href="#" onClick="showtildeltimer('<=xmid(x)%>','<=strDageChkboxOne%>','<=id%>','<=periodeSel%>')" class=vmenuglobal>Ny</a>-->
		
		
        
        <span style="font-size:9px; color:#999999; line-height:10px;">
        <br>
		Sidst opd: <%=meforecaststamp %></span>
		
            <!-- 
        <%if periodesel = "6" OR periodesel = "12" then %>
		<br />Norm timer pr. md ~ 
        <%'call normtimerPer(mid, jobStartKri, datoInterval, 0) %>
        <%="<b>"& formatnumber(ntimPer/periodesel, 1) & "</b>  timer"%>

         <%
          
         else %>
        
            <br />Norm timer pr. uge ~ 
        <%'call normtimerPer(mid, jobStartKri, 6, 0) %>
        <%="<b>"& formatnumber(ntimPer, 1) & "</b>  timer"%>

        <%
          
        end if %>
                -->
           
                

        <%if media <> "print" then 
        call positiv_aktivering_akt_fn()
        if cint(pa_aktlist) = 1 then%>
               <br /><a href="#" id="anl_<%=mid%>" class="rodstor">Tilf�j forecast p� job +</a> 
        <%end if %>

        <%else %>
        &nbsp;
		<%end if%>
		
		</td>
		
		
		
		<%
		cspan = rdimnb  '12
		
            
        if cint(vis) = 0 then%>
		<td align=right style="padding-right:5px; background-color:#FFFFFF;" colspan=<%=cspan+6%>>&nbsp;</td>
	</tr>

    <%end if


    end function

	%>
	
	
		
		