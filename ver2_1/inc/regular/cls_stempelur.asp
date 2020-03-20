 
 <%
	'******************************************************************************
     public aabentLogudfindes, lastDay
     function logindhistorik_week_60_100(usemrn, visning, varTjDatoUS_manSQL, varTjDatoUS_sonSQL)
     lastDay = 0
     aabentLogudfindes = 0
     %>

                                           <%
                                           select case visning
                                            case 0,1
                                            tblWdt = 100
                                            case 2, 3
                                            tblWdt = 100
                                            end select
                                               
                                               
                                               
                                           if cint(visning) = 1 then %>
                                           <table width="<%=tblWdt %>%" border="0" class="table table-striped">
                                              <thead>
                                           <%else %>
                                            <table width="<%=tblWdt %>%" border="0">
                                            <%end if %>
                                              <tr>

                                                   <%if cint(visning) = 1 then %>
                                                   <td>Medarbejder</td><td>Init</td>
                                                   <%end if %>
                                                  
                                                  <%if cint(visning) = 0 OR cint(visning) = 1 OR cint(visning) = 3 then %>
                                                  <td colspan="2">Dato</td>
                                                  <%end if %>
                                                  
                                                  <td colspan="2">Klokkeslet</td><td align=right>hh:mm</td><td align=right>100 tals</td>

                                                  <%if cint(visning) = 1 then %>
                                                   <td align=right>Timer til løn</td><td align=right>Pause</td><td align=right>Projekttimer</td><td>&nbsp;</td>
                                                   <%end if %>

                                                  <%if visning = 1 then %>
                                                  </thead>
                                                  <%end if %>
                                              </tr>

                                         <%
                                              'visning 0 godkend request timer, viser valgt medarb. pr. uge
                                              'visning 1 godkend request timer, vis alle medarb. pr. uge
                                              'visning 2 ugeseddel pr. uge

                                              if cint(visning) = 0 OR cint(visning) = 2 OR cint(visning) = 3 then

                                                logintjkDatoSql = " AND DATE(login) BETWEEN '"& varTjDatoUS_manSQL & "' AND '"& varTjDatoUS_sonSQL &"'"
                                                'logintjkDatoSql = year(oRec5("tdato")) & "/" & month(oRec5("tdato")) &"/"& day(oRec5("tdato")) 

                                                '******* Tjekker logindtid ********'
                                                timerMin100Tot = 0

                                                if cint(visning) = 2 OR cint(visning) = 3 then 'viser også nuværende dvs logud IS NULL 
                                                strSQL = "SELECT id, logud, login, minutter, stempelurindstilling FROM login_historik WHERE mid = "& usemrn &""_
                                                &" AND stempelurindstilling <> -10 "& logintjkDatoSql &" ORDER BY login LIMIT 50"
                                                else
                                                 strSQL = "SELECT id, logud, login, minutter, stempelurindstilling FROM login_historik WHERE mid = "& usemrn &""_
                                                &" AND stempelurindstilling <> -10 "& logintjkDatoSql &" AND logud IS NOT NULL ORDER BY login LIMIT 50"
                                                end if
                                                

                                                else 'viser alle


                                                 if usemrn <> 0 then 
                                                 midSQL = "l.mid = "& usemrn
                                                 else
                                                 midSQL = "l.mid <> 0 "
                                                 end if

                                                logintjkDatoSql = " AND DATE(l.login) BETWEEN '"& varTjDatoUS_manSQL & "' AND '"& varTjDatoUS_sonSQL &"'"
                                                timerMin100Tot = 0
                                                strSQL = "SELECT id, logud, l.login, minutter, mnavn, init, m.mid, stempelurindstilling FROM login_historik l "_
                                                &" LEFT JOIN medarbejdere m ON (m.mid = l.mid) WHERE "& midSQL &" "_
                                                &" AND stempelurindstilling > 0 "& logintjkDatoSql &" AND logud IS NOT NULL ORDER BY m.mid, l.login LIMIT 300"
                                            
                                                

                                                end if



                                            'if session("mid") = 1 then
                                            'Response.write strSQL
                                            'Response.flush
                                            'end if
                                            
                                            lcount = 0
                                            thisDayC = 0
                                            oRec.open strSQL, oConn, 3
                                            while not oRec.EOF 
                                            
                                          

                                            login = oRec("login")
                                            thisDay = datepart("w", login, 2,2)

                                            timerThis = 0
                                            'ER DER LOGGET UD
                                            if isNull(oRec("logud")) <> true then 'visning 0, visning 1
                                            logud = oRec("logud")

                                            timerThis = formatnumber(oRec("minutter")/60, 2)
                                            minutterThisDB = oRec("minutter")

                                            else
                                            aabentLogudfindes = 1
                                            logud = now
                                            timerThis = formatnumber(datediff("n", login, logud, 2, 2)/60, 2)
                                            minutterThisDB = formatnumber(datediff("n", login, logud, 2, 2), 0)
                                            end if
                                            
                                           
                                            

                                            
                                            
                                            if cdbl(timerThis) < 1 then
                                            timerThis = 0
                                            end if

                                            'response.Write "timerThis_komma: " & timerThis_komma & ""& oRec("minutter") &" timerThis: " & timerThis & "<br>"
                                            'response.flush
                                            timerThis_komma = instr(timerThis, ",")
                                            if timerThis_komma <> 0 then
                                            timerThis = left(timerThis, timerThis_komma - 1)
                                            end if

                                            timerThis_komma = instr(timerThis, ".")
                                            if timerThis_komma <> 0 then
                                            timerThis = left(timerThis, timerThis_komma - 1)
                                            end if

                                            minutterThis = 0
                                            minutterThis = formatnumber(minutterThisDB/60, 2)
                                            minutterThis = right(minutterThis, 2)
                                            minutterThis = formatnumber((minutterThis*60)/100, 0)
                                          

                                            'minutterThis = formatnumber(minutterThis * 60/100, 0)
                                            minutter100 = formatnumber(minutterThis/60*100, 0) 

                                            if minutterThis < 10 then
                                            minutterThis = "0"& minutterThis
                                            end if

                                            if minutter100 < 10 then
                                            minutter100 = "0"& minutter100
                                            end if

                                            

                                            if cint(visning) = 0 OR cint(visning) = 2 OR cint(visning) = 3 then

                                            'Response.write "<tr><td>"& left(weekdayname(weekday(login)), 3) &".</td><td>"& formatdatetime(login, 2) &"</td><td>"& left(formatdatetime(login, 3), 5) & " - " & left(formatdatetime(logud, 3), 5) & "</td><td style=""width:100px;"">&nbsp;</td><td align=right>"& timerThis &":"& minutterThis &"</td><td align=right>"& timerThis &","& minutter100 &"</td></tr>" 


                                            
                                                'if session("mid") = 1 then
                                                'Response.Write "HER: "& thisfile &" lastDay: " & lastDay &"<> thisday: "& thisday & "AND lcount > 0: "& lcount &" AND thisDayC > 1: " & thisDayC &"<br>"
                                                'end if
                                                
                                                

                                                if cint(lastDay) <> cint(thisday) AND lcount > 0 then 'AND thisDayC > 1 then
                                                    call dagsTotal(timerThisDayTot, visning)
                                                end if

                                                if cint(lastDay) <> cint(thisday) AND lcount > 0 then
                                                timerThisDayTot = 0
                                                thisDayC = 0
                                                end if
                                             
                                             if datepart("w", login, 2,2) > 5 then 'lørdag & søndag
                                             trLinje = "<tr style=""background-color:silver;"">"
                                             else
                                             trLinje = "<tr>"
                                             end if

                                             if cint(visning) = 0 OR cint(visning) = 1 OR cint(visning) = 3 then
                                             trLinje = trLinje &"<td>"& left(weekdayname(weekday(login)), 3) &".</td>"
                                             trLinje = trLinje &"<td>"& formatdatetime(login, 2) &"</td>"
                                             end if

                                           

                                             if oRec("stempelurindstilling") <> -1 then
                                             sign = ""
                                                if isNull(oRec("logud")) = true AND visning = 2 then 'BLINK
                                                trLinje = trLinje & "<td>"& left(formatdatetime(login, 3), 5) & " - <span class=blink style=""color:red;"">" & left(formatdatetime(logud, 3), 5) & "</span></td>"
                                                else
                                                trLinje = trLinje & "<td>"& left(formatdatetime(login, 3), 5) & " - " & left(formatdatetime(logud, 3), 5) & "</td>"
                                                end if
                                             else
                                             trLinje = trLinje & "<td>Pause</td>"
                                             sign = "-"
                                             end if
                                             
                                             if visning = 0 then
                                             trLinje = trLinje & "<td style=""width:100px;"">&nbsp;</td>"
                                             end if

                                             if visning = 2 OR visning = 3 then
                                             trLinje = trLinje & "<td style=""width:20px;"">&nbsp;</td>"
                                             end if


                                             if isNull(oRec("logud")) = true AND (visning = 2 OR visning = 3) then 'BLINK
                                             trLinje = trLinje & "<td align=right class=blink style=""color:red;"">"& timerThis &":"& minutterThis &"</td>"_
                                             &"<td align=right class=blink style=""color:red;"">"& timerThis &","& minutter100 &"</td>"_
                                             &"</tr>" 
                                             else
                                             trLinje = trLinje & "<td align=right>"&sign&""& timerThis &":"& minutterThis &"</td>"_
                                             &"<td align=right>"&sign&""& timerThis &","& minutter100 &"</td>"_
                                             &"</tr>" 
                                             end if

                                             response.write trLinje

                                            else 'visning = 1

                                                     Response.write "<tr>"_
                                                     &"<td>"& oRec("mnavn")&"</td><td>"& oRec("init")&"</td>"_
                                                     &"<td>"& left(weekdayname(weekday(login)), 3) &".</td>"_
                                                     &"<td>"& formatdatetime(login, 2) &"</td>"_
                                                     &"<td>"& left(formatdatetime(login, 3), 5) & " - " & left(formatdatetime(logud, 3), 5) & "</td>"_
                                                     &"<td style=""width:100px;"">&nbsp;</td><td align=right>"& timerThis &":"& minutterThis &"</td>"_
                                                     &"<td align=right>"& timerThis &","& minutter100 &"</td>"


                                                     datoSQL = year(login) &"/"& month(login) &"/"& day(login)

                                                     lonTimerTreg = 0
                                                     strSQLtimerLon = "SELECT SUM(timer) AS sumtimer FROM timer WHERE tdato = '"& datoSQL &"' AND tmnr = "& oRec("mid") &" AND tfaktim > 5 AND tfaktim <> 10 GROUP BY tmnr"

                                                     'Response.write strSQLtimerLon
                                                     'Response.flush
                                                     oRec4.open strSQLtimerLon, oConn, 3
                                                     if not oRec4.EOF then

                                                     lonTimerTreg = oRec4("sumtimer")

                                                     end if

                                                     oRec4.close

                                                     Response.write "<td align=right>"
                                             
                                                     if visning = 1 then'godkend request_timer
                                                     Response.write "<a href=""godkend_request_timer.asp?FM_medarb="& oRec("mid") &"&visning=0&varTjDatoUS_man="& varTjDatoUS_man &""" target=""_blank"">"& formatnumber(lonTimerTreg, 2) &"</a></td>"
                                                     else
                                                     Response.write formatnumber(lonTimerTreg, 2)
                                                     end if

                                                     pauseTimerTreg = 0
                                                     strSQLtimerPause = "SELECT SUM(timer) AS sumtimer FROM timer WHERE tdato = '"& datoSQL &"' AND tmnr = "& oRec("mid") &" AND tfaktim = 10 GROUP BY tmnr"

                                                     'Response.write strSQLtimerLon
                                                     'Response.flush
                                                     oRec4.open strSQLtimerPause, oConn, 3
                                                     if not oRec4.EOF then

                                                     pauseTimerTreg = oRec4("sumtimer")

                                                     end if

                                                     oRec4.close

                                                     Response.write "<td align=right>"& formatnumber(pauseTimerTreg, 2) &"</td>"

                                                     projektTimerTreg = 0
                                                     strSQLtimerprojekt = "SELECT SUM(timer) AS sumtimer FROM timer WHERE tdato = '"& datoSQL &"' AND tmnr = "& oRec("mid") &" AND (tfaktim = 1 OR tfaktim = 2) GROUP BY tmnr"
                                                     oRec4.open strSQLtimerprojekt, oConn, 3
                                                     if not oRec4.EOF then

                                                     projektTimerTreg = oRec4("sumtimer")

                                                     end if
                                                     oRec4.close

                                                     Response.write "<td align=right>"& formatnumber(projektTimerTreg, 2) &"</td>"

                                                     'ALERT HVIS løntimer afviger? 
                                                     alert = 0
                                                     lonTimerstempel100 = timerThis &","& minutter100
                                                     if ((formatnumber(lonTimerstempel100, 2) - formatnumber(lonTimerTreg+pauseTimerTreg, 2)) > "0,25") then
                                                     alert = 1
                                                     end if


                                                     'SKAL DER VISES ALERT HVIS projekttimer afviger? 
                                                     if (cdbl(lonTimerstempel100) - formatnumber(projektTimerTreg, 2)) > 1 then
                                                     'alert = 1
                                                     end if

                                                     if cint(alert) = 1 then
                                                     Response.write "<td align=right>&nbsp;&nbsp;&nbsp;<span style=color:red>!</span></td>"
                                                     else
                                                     Response.write "<td align=right>&nbsp;&nbsp;&nbsp;</td>"
                                                     end if


                                                     Response.write "</tr>" 


                                            end if

                                             
                                             if oRec("stempelurindstilling") <> -1 then
                                             timerMin100this = (timerThis &","& minutter100) * 1
                                             else
                                             timerMin100this = (timerThis &","& minutter100) * -1
                                             end if

                                             timerMin100Tot = timerMin100Tot*1 + (timerMin100this*1)

                                             lastDay = datepart("w", login, 2,2)
                                             
                                             if oRec("stempelurindstilling") <> -1 then
                                             timerThisDayTot = (timerThisDayTot*1)+(minutterThisDB*1)
                                             else
                                             timerThisDayTot = (timerThisDayTot*1)+(minutterThisDB*-1)
                                             end if

                                             thisDayC = thisDayC + 1

                                            lcount = lcount + 1
                                            oRec.movenext
                                            wend
                                            oRec.close
                                        
                                        if lcount = 0 then
                                        %>
                                        <tr><td>Ingen loginds i valgte periode</td></tr>
                                            <%else 
                                                
                                                      if visning <> 1 then
                                                
                                                        'if thisDayC > 1 then

                                                            call dagsTotal(timerThisDayTot, visning)
                                                   
                                                        'end if
                                                

                                                              if visning = 0 or visning = 1 then 
                                                              %>
                                                              <tr>
                                                                  <td><b>Total:</b></td>
                                                                  <td>&nbsp;</td>
                                                                  <td>&nbsp;</td>
                                                                  <td>&nbsp;</td>
                                                                  <td>&nbsp;</td>
                                                                  <td align=right><b><%=formatnumber(timerMin100Tot, 2) %></b></td>
                                                  
                                                              </tr>
                                                               <%    
                                                             end if

                                                       end if
                                            
                                         end if %>
                                        </table>




     <%end function



        public timerThisDayTot, thisDayC, timerThisDayLontimerTotTxt
            function dagsTotal(timerThisDayTot, visning)

                        timerThisDayLontimerTotTxt = 0

                        'if session("mid") = 1 then
                        'Response.Write "HER: " & timerThisDayTot
                        'end if

                        'timerThisTot = 0
                        timerThisTot = formatnumber(timerThisDayTot/60, 2)
                        if timerThisTot < 1 then
                        timerThisTot = 0
                        end if

                        timerThisTot_komma = instr(timerThisTot, ",")
                        if cint(timerThisTot_komma) <> 0 then
                        timerThisTot = left(timerThisTot,  timerThisTot_komma - 1)
                        end if

                        timerThisTot_komma = instr(timerThisTot, ".")
                        if cint(timerThisTot_komma) <> 0 then
                        timerThisTot = left(timerThisTot,  timerThisTot_komma - 1)
                        end if

                        minutterThisTot = 0
                        minutterThisTot = formatnumber(timerThisDayTot/60, 2)
                        minutterThisTot = right(minutterThisTot, 2)
                        minutterThisTot = formatnumber((minutterThisTot*60)/100, 0)
                                          

                        'minutterThis = formatnumber(minutterThis * 60/100, 0)
                        minutter100Tot = formatnumber(minutterThisTot/60*100, 0) 

                        if minutterThisTot < 10 then
                        minutterThisTot = "0"& minutterThisTot
                        end if

                        if minutter100Tot < 10 then
                        minutter100Tot = "0"& minutter100Tot
                        end if


               
                'style=""width:100px;""
                trLinjeTot = "<tr>"

                if visning = 0 Or visning = 1 Or visning = 3 then
                trLinjeTot = trLinjeTot &"<td>&nbsp;</td>"
                trLinjeTot = trLinjeTot &"<td>&nbsp;</td>"
                end if
                
                trLinjeTot = trLinjeTot &"<td>&nbsp;</td>"
                trLinjeTot = trLinjeTot &"<td>&nbsp;</td><td align=right><b>"& timerThisTot &":"& minutterThisTot &"</b></td>"
                trLinjeTot = trLinjeTot &"<td align=right><b>"& timerThisTot &","& minutter100Tot &"</b></td>"
                trLinjeTot = trLinjeTot &"</tr>" 


                response.write trLinjeTot

                timerThisDayLontimerTotTxt = timerThisTot &","& minutter100Tot 
                timerThisDayTot = 0
                thisDayC = 0

                

            end function

'** END 


'***********************************************************
'**** MAIN LOGUD FUNKTION. MONOTOR, LOGUD FUNTION fra HOVEDMENU.


'*** Fordel stempleurs timer ud på aktiviteter
'*** H&L timer CFLOW
'*** Projekttimer
'***********************************************************
function fordelStempelurstimer(uid, lto, dothis, logudDone, LogudDateTime, thisid) 


logudTid = LogudDateTime 'year(loginDato) &"/"& month(loginDato)&"/"& day(loginDato) & " " & logudHH &":"& logudMM & ":00"
logudTidafrundet = 0

overtid50typ = 54
overtid100typ = 51

select case lto 
case "cflow"

        call medariprogrpFn(uid)
        if (instr(medariprogrpTxt, "#4#") <> 0 OR instr(medariprogrpTxt, "#2#") <> 0) AND (instr(medariprogrpTxt, "#14#") = 0 AND instr(medariprogrpTxt, "#16#") = 0 AND instr(medariprogrpTxt, "#19#") = 0) then 'Produktion ONLY
        'if session("mid") = 1 then


          overtid50typ = 54
          overtid100typ = 51

             '* afrund mellen 15:30 og 15:45 til 15.30
             if datepart("h", logudTid, 2,2) = 15 then

                select case datepart("n", logudTid, 2,2)
                case 31,32,33,34,35,36,37,38,39,40,41,42,43,44,45
                
                        logudTid = year(LogudDateTime) &"/"& month(LogudDateTime)&"/"& day(LogudDateTime) & " 15:30:00"
                        logudTidafrundet = 1
                        logudTidopr = LogudDateTime
                end select
            

             end if

         else

        
             '*** Andre projektgrupper end produktion **'
             if instr(medariprogrpTxt, "#16#") <> 0 then 'Enginnering (fastlønn)
             overtid50typ = 91
             overtid100typ = 92
             else
             overtid50typ = 54
             overtid100typ = 51
             end if


         end if

         '**TEST
         'if session("mid") = 1 then
              '* afrund mellen 15:30 og 15:45 til 15.30
         '    if datepart("h", logudTid, 2,2) = 11 then

         '       select case datepart("n", logudTid, 2,2)
         '       case 15,16,17,18,19,20,21,22,23,24,25
                
         '               logudTid = year(LogudDateTime) &"/"& month(LogudDateTime)&"/"& day(LogudDateTime) & " 15:30:00"
         '               logudTidafrundet = 1
         '               logudTidopr = LogudDateTime
         '       end select
            

         '    end if
         'end if

end select

datoThis = day(now)&"/"& month(now)&"/"&year(now) 'Sættes længere nede = Altid logindato

if dothis = 2 then 'fra komme/gå side. Opdater alle
strSQL = "SELECT logud, id, login, stempelurindstilling FROM login_historik WHERE logud IS NOT NULL AND id = " & thisid
else
strSQL = "SELECT id, login, stempelurindstilling FROM login_historik WHERE mid = " & uid &" AND stempelurindstilling <> -1 AND logud IS NULL ORDER BY id DESC LIMIT 1"
end if

'if session("mid") = 1 then
'response.write "strSQL: " & strSQL
'response.flush
'end if

oRec.open strSQL, oConn, 3 
if not oRec.EOF then


    if cint(dothis) = 2 then '** Opdater hele ugen fra komme/gå siden

         logudHHthis = datepart("h", oRec("logud"), 2,2)
         logudMMthis = datepart("n", oRec("logud"), 2,2)

         if logudHHthis < 10 then
         logudHHthis = "0"&logudHHthis
         end if

         if logudMMthis < 10 then
         logudMMthis = "0"&logudMMthis
         end if

    logudTid = year(oRec("logud")) &"/"& month(oRec("logud"))&"/"& day(oRec("logud")) & " " & logudHHthis &":"& logudMMthis & ":00"
    else
    logudTid = logudTid
    end if


    loginTid = oRec("login") 'year(loginDato) &"/"& month(loginDato)&"/"& day(loginDato) & " " & loginHH &":"& loginMM & ":00"
	datoThis = day(loginTid)&"/"& month(loginTid)&"/"&year(loginTid)


	if len(loginTid) <> 0 AND len(logudTid) <> 0 then

	loginTidAfr = left(formatdatetime(loginTid, 3), 5)
	logudTidAfr = left(formatdatetime(logudTid, 3), 5)

       

        if dateDiff("d", cdate(loginTid), cdate(logudTid), 2,2) > 0 then 'Der er logget ud efter midnat
            minThisDIFF = datediff("s", loginTidAfr, "23:59")/60
            minThisDIFF = minThisDIFF + 1 + datediff("s", "00:00", logudTidAfr)/60
        else
            minThisDIFF = datediff("s", loginTidAfr, logudTidAfr)/60
        end if
	
    else

    minThisDIFF = 0

	end if


        'if uid = 1 AND lto = "cflow" then
        '    Response.write loginTidAfr & " : " & logudTidAfr & " = " & minThisDIFF
        '    Response.end

        'end if

    'if session("mid") = 1 AND lto = "cflow" then

    '    LogudDateTime = year(now)&"/"& month(now)&"/"&day(now)&" 02:00:00"
    '    logudTid = LogudDateTime
    '    loginTid_opr = oRec("login") 'year(now)&"/"& month(now)&"/"&day(now)&" 07:30:00"
    'else

        loginTid_opr = oRec("login")

    'end if
    ignoreLogout = 0
    select case lto  
    case "cflow" 
            if cdbl(minThisDIFF) < 0 AND oRec("stempelurindstilling") = 1 then '0 = Det er ok med NUL minutter, ellers kan folk ikke logge hurtigt ind og ud
            ignoreLogout = 1
            end if
    case else
            ignoreLogout = 0
    end select

	
    if cint(ignoreLogout) = 0 then

    if dothis = 2 then 'FRA KOMME GÅ SIDEN behold opr.
    strSQL2 = "UPDATE login_historik SET logud = '"& logudTid &"', minutter = "& minThisDIFF &" WHERE id = " & oRec("id")
    else

        if cint(logudTidafrundet) = 1 then
        logudTidopr = logudTidopr
        manuelt_afsluttetSQL = ", manuelt_afsluttet = 3" 'stempelur tilrettet pga. åbningstider
        else
        logudTidopr = logudTid
        manuelt_afsluttetSQL = ""
        end if

	    strSQL2 = "UPDATE login_historik SET logud = '"& logudTid &"', logud_first = '"& logudTidopr &"', minutter = "& minThisDIFF &" "& manuelt_afsluttetSQL &" WHERE id = " & oRec("id")
	end if
        'if session("mid") = 1 then
        'Response.Write strSQL2
        'Response.end	
        'end if
    oConn.execute(strSQL2)






    select case lto
    case "intranet - local", "cflow"
    

        '*** Overtid ***'
        'if len(trim(request("FM_timer_overtid"))) <> 0 AND len(trim(request("fm_overtid_ok"))) <> 0 then
        'timertilindlasning = request("FM_timer_overtid")
        'else
        'timertilindlasning = 0
        'end if
           



        if len(trim(request("indlaspaajob"))) <> 0 then
        indlaspaajob = request("indlaspaajob")
        else
        indlaspaajob = 0
        end if

        
        timerthis_mtx_tot = 0

        'if timertilindlasning <> 0 then
            thisDayWeekday = datepart("w", datoThis, 2,2)

            select case thisDayWeekday
            case 6,7
            ftaktim = overtid100typ 'overtid50typ 
            'Overtid 50% = Husk SKAL VÆRE spiciel for Engineering og Automasjon
            case else
            ftaktim = 30 'Ordinær
            end select

           

            loginTid = formatdatetime(oRec("login"), 3)
            logudTidmatrix = formatdatetime(logudTid, 3)
            loginDate = formatdatetime(oRec("login"), 2)
            call matrixtimespan_2018(uid, datoThis, 0, loginTid, logudTidmatrix, loginDate, lto)

            call overtidsTillaeg(oRec("stempelurindstilling"), lto, loginTid_opr, logudTid, uid, timerthis_mtx, ftaktim)

            timerthis_mtx_komma = replace(timerthis_mtx, ".", ",")
            timerthis_mtx_tot = (timerthis_mtx_tot * 1) + (timerthis_mtx_komma * 1)

          

            '*** Frokost / Pause
            if cdbl(timerthis_mtx) > 3 AND cint(indlaspaajob) = 0 then 'Kun når der stemples ud

                pausefindes = 0
                strSQLpausefindes = "SELECT tid FROM timer WHERE tdato = '"& year(datoThis) & "/"& month(datoThis) & "/" & day(datoThis) &"' AND tmnr = "& uid &" AND tfaktim = 10" 

                'If session("mid") = 1 then
                '    Response.Write strSQLpausefindes & "<br>"
                '    Response.flush

                'end if


                oRec4.open strSQLpausefindes, oConn, 2
                if not oRec4.EOF then
                pausefindes = 1
                end if
                oRec4.close

                '*** Kun 1 pause pr. dag.
                if cint(pausefindes) = 0 then

                    call overtidsTillaeg(oRec("stempelurindstilling"), lto, loginTid_opr, logudTid, uid, "-0,5", ftaktim)
                    call overtidsTillaeg(oRec("stempelurindstilling"), lto, loginTid_opr, logudTid, uid, "0,5", 10)

                    haltimespause = (1/2) 

                    timerthis_mtx_tot = (timerthis_mtx_tot * 1) - (haltimespause * 1)

                end if

            end if

            
            'if session("mid") = 1 then
            'Response.write timerthis_mtx_tot & "<br>"
            'Response.flush
            'end if

            'timertilindlasning = 0
        'end if



       
         '****** Overarbejde  / Overtid ************************************************
         '*** Engineering ALDRIG overtid altid føres explicit ************** 20190308 **
         '******************************************************************************
         if instr(medariprogrpTxt, "#16#") = 0 then


                 select case thisDayWeekday
                 case 6,7
                 '** SKAL IKKE BEREGNE overtod 50 og 100% i weekenden da all timer skal på overtid 100% type 51

                 case else
                 '*** Hverdage *** Beregne overtid 50 og 100%

                    ftaktim = overtid50typ  'Overtid 50% type. Beregn om det skal være 50%
                    loginTid = formatdatetime(oRec("login"), 3)
                    logudTidmatrix = formatdatetime(logudTid, 3)
                    loginDate = formatdatetime(oRec("login"), 2)

        
                    call matrixtimespan_2018(uid, datoThis, 1, loginTid, logudTidmatrix, loginDate, lto)

                    'Alrdirg overtid når der er arbejdet mindre end 8 timer 29101015
                    arbejdsdagTimerIalt = dateDiff("h", loginDate &" "& loginTid, logudTid, 2,2)

                    'if session("mid") = 1 then
                    'Response.Write "arbejdsdagTimerIalt: "& arbejdsdagTimerIalt & "loginTid: "& loginTid & "logudTid: " & logudTid
                    'Response.end
                    'end if

                    if cdbl(arbejdsdagTimerIalt) < 8 then
                    'ftaktim sættes til ORDINÆR 
                    ftaktim = 30 ' Ordinær 
                    end if

                    call overtidsTillaeg(oRec("stempelurindstilling"), lto, loginTid_opr, logudTid, uid, timerthis_mtx, ftaktim)

                    timerthis_mtx_komma = replace(timerthis_mtx, ".", ",")
                    timerthis_mtx_tot = (timerthis_mtx_tot * 1) + (timerthis_mtx_komma * 1)
        

                    'if session("mid") = 1 then
                    'Response.write "medariprogrpTxt: "& medariprogrpTxt &"<hr> loginTid: "& loginTid &" logudTidmatrix: "& logudTidmatrix &" logudTid: "& logudTid &" timerthis_mtx: "& timerthis_mtx &" timerthis_mtx: "& timerthis_mtx_komma & "<br>"& timerthis_mtx_tot & "<br>"
                    'Response.flush
                    'end if


                    ftaktim = overtid100typ 'Overtid 100% type. Beregn om det skal være 100%
                    loginTid = formatdatetime(oRec("login"), 3)
                    logudTidmatrix = formatdatetime(logudTid, 3)
                    loginDate = formatdatetime(oRec("login"), 2)
                    call matrixtimespan_2018(uid, datoThis, 2, loginTid, logudTidmatrix, loginDate, lto)

                    'Alrdirg overtid når der er arbejdet mindre end 8 timer 29101015
                    arbejdsdagTimerIalt = dateDiff("h", loginDate &" "& loginTid, logudTid, 2,2)
                    if cdbl(arbejdsdagTimerIalt) < 8 then
                    'ftaktim sættes til ORDINÆR 
                    ftaktim = 30 ' Ordinær 
                    end if

                    call overtidsTillaeg(oRec("stempelurindstilling"), lto, loginTid_opr, logudTid, uid, timerthis_mtx, ftaktim)
            
                    timerthis_mtx_komma = replace(timerthis_mtx, ".", ",")
                    timerthis_mtx_tot = (timerthis_mtx_tot * 1) + (timerthis_mtx_komma * 1)


      

                        '*** Timer efter midnat **'
                        ftaktim = 90 'Overtid 100% type. Beregn om det skal være 100%
                        loginTid = formatdatetime(oRec("login"), 3)
                        logudTidmatrix = formatdatetime(logudTid, 3)
                        loginDate = formatdatetime(oRec("login"), 2)
                        call matrixtimespan_2018(uid, datoThis, 8, loginTid, logudTidmatrix, loginDate, lto)

                        call overtidsTillaeg(oRec("stempelurindstilling"), lto, loginTid_opr, logudTid, uid, timerthis_mtx, ftaktim)
                        timerthis_mtx_komma = replace(timerthis_mtx, ".", ",") 
                        timerthis_mtx_tot = (timerthis_mtx_tot * 1) + (timerthis_mtx_komma * 1)

      

                 end select ' select case thisDayWeekday


            
         end if ' if instr(medariprogrpTxt, "#16#") <> 0 then = Engineering SKAL ikke have overtid, skal altid føres explixcit






         timertilindlasning = 0

        '*** Rejsetid
        if len(trim(request("FM_rejsetid"))) <> 0 AND request("FM_rejsetid") <> 0 then
        timertilindlasning = request("FM_rejsetid")
        else
        timertilindlasning = 0
        end if

        if timertilindlasning <> 0 then
        ftaktim = 52 'Standard 50% type. Beregn om det skal være 100%
        call overtidsTillaeg(oRec("stempelurindstilling"), lto, loginTid_opr, logudTid, uid, timertilindlasning, ftaktim)
        
        timerthis_mtx_komma = replace(timertilindlasning, ".", ",")
        timerthis_mtx_tot = (timerthis_mtx_tot * 1) - (timerthis_mtx_komma * 1)

        timertilindlasning = 0
        
        

        'if session("mid") = 1 then
        '    Response.write timerthis_mtx_tot & "<br>"
        '    Response.flush
        '    end if
    
        end if





        '*** FM_arbute_no 
        if len(trim(request("FM_arbute_no"))) <> 0 then
        timertilindlasning = 1
        else
        timertilindlasning = 0
        end if

        if timertilindlasning <> 0 then
        ftaktim = 50 
        call overtidsTillaeg(oRec("stempelurindstilling"), lto, loginTid_opr, logudTid, uid, timerthis_mtx_tot, ftaktim)
        timertilindlasning = 0
        end if



          '*** FM_arbute_world
        if len(trim(request("FM_arbute_world"))) <> 0 then
        timertilindlasning = 1
        else
        timertilindlasning = 0
        end if

        if timertilindlasning <> 0 then
        ftaktim = 60 
        call overtidsTillaeg(oRec("stempelurindstilling"), lto, loginTid_opr, logudTid, uid, timerthis_mtx_tot, ftaktim)
        timertilindlasning = 0
        end if

           '*** FM_arbute_teamleder
        if len(trim(request("FM_arbute_teamleder"))) <> 0 then
        timertilindlasning = 1
        else
        timertilindlasning = 0
        end if

        if timertilindlasning <> 0 then
        ftaktim = 61 
        call overtidsTillaeg(oRec("stempelurindstilling"), lto, loginTid_opr, logudTid, uid, timerthis_mtx_tot, ftaktim)
        timertilindlasning = 0
        end if


        '*** Evt timer ønskes udbetalt
        if len(trim(request("fm_overtidonskesudbetalt"))) <> 0 then
        timertilindlasning = request("fm_overtidonskesudbetalt")
        else
        timertilindlasning = 0
        end if
        
        if timertilindlasning <> 0 then
        ftaktim = 33 
        call overtidsTillaeg(oRec("stempelurindstilling"), lto, loginTid_opr, logudTid, uid, timerthis_mtx_tot, ftaktim)
        timertilindlasning = 0
        end if


        '*** INDLÆS på projekter **'
        if cint(fromweblogud) <> 1 then

             'AND instr(medariprogrpTxt, "#3#") = 0
             if (instr(medariprogrpTxt, "#14#") = 0 AND instr(medariprogrpTxt, "#16#") = 0 AND instr(medariprogrpTxt, "#19#") = 0) OR (instr(medariprogrpTxt, "#4#") <> 0 OR instr(medariprogrpTxt, "#2#") <> 0) then

            'if aktivitetid <> 0 AND cdbl(timerthis_mtx_tot) <> 0 then
            insertUpdate = 2 '2: FROM sesaba
            insUpdDate = loginTid
            extsysid = 0
            timerkom = ""
            koregnr = "" 
            destination = ""
            usebopal = 0
            call indlasTimerTfaktimAktid(lto, uid, timerthis_mtx_tot, 0, aktivitetid, insertUpdate, insUpdDate, extsysid, timerkom, koregnr, destination, usebopal)

            'end if


            end if

        end if
        

    end select



	

    end if     'cint(ignoreLogout) = 0
	
end if
oRec.close 


end function
'************************************ END FUNCTION FORDEL *******************************************************







'****************************************************************************************************************
'*** FIDNES DER TIMER
'****************************************************************************************************************
  sub findesDerTimer(io, dt, medarbSel)

     
     ddTreg = year(dt) & "/"& month(dt) & "/" & day(dt) 
    

     %>
           <span style="color:#000000; font-size:9px;">

        <%'FINDES der fraværs registreringer på dagen? 
            
           tcnt = 0
            strSQLfindesfravertimer = "SELECT taktivitetnavn, timer FROM timer WHERE tdato = '"& ddTreg &"' AND tmnr = "& medarbSel &" AND tfaktim > 6 LIMIT 3"

            'if session("mid") = 1 then
            'response.write strSQLfindesfravertimer
            'response.flush
            'end if

            oRec6.open strSQLfindesfravertimer, oConn, 3
            
            while not oRec6.EOF

            if tcnt > 0 then
            %><br /><%
            end if
             
            %>
            <%=oRec6("taktivitetnavn") & ": "& formatnumber(oRec6("timer"), 2) %> t.
            <%

            tcnt = tcnt + 1
            oRec6.movenext
            wend
            oRec6.close

             %>

        </span>

        


     <%
     end sub



 public forvalgt_stempelur

 function fv_stempelur()
 forvalgt_stempelur = 1
    
    strSQLfv = "SELECT id FROM stempelur WHERE forvalgt = 1"
    'Response.write strSQLfv
    'Response.flush

    oRec5.open strSQLfv, oConn, 3
    if not oRec5.EOF then

    forvalgt_stempelur = oRec5("id")

    end if
    oRec5.close


end function



public use_ig_sltid, loginTid, manuelt_afsluttet, stpause, stpause2
function stpauser_ignorePer(lgi_dato,lgi_sttid_hh,lgi_sttid_mm, medid)


    strSperretid_grp = ""
    strSQL = "SELECT ignorertid_st, ignorertid_sl, stpause, stpause2, sperretid_grp FROM licens WHERE id = 1"
	oRec.open strSQL, oConn, 3
    if not oRec.EOF then
    '*** sperretid_grp2 og ignorertid_st 2, ignorertid_sl 2 SKAL aktiveres her. De er lavet i DB og kontrolapanel. '20180823 

    ig_sttid = oRec("ignorertid_st")
    ig_sltid = oRec("ignorertid_sl")
    
    stpause = oRec("stpause")
    stpause2 = oRec("stpause2")

    strSperretid_grp = oRec("sperretid_grp")
   
    end if
    oRec.close
    
    if len(ig_sttid) <> 0 then
    ig_sttid = formatdatetime(ig_sttid, 3)
    end if
    
    if len(ig_sltid) <> 0 then
    ig_sltid = formatdatetime(ig_sltid, 3)
    end if
    
    
                    '*******************************************
					'*** Test for ignorer periode v. login *****
					testloginTid = year(lgi_dato) &"/"& month(lgi_dato)&"/"& day(lgi_dato) & " " & lgi_sttid_hh &":"& lgi_sttid_mm & ":00"
					testig_sttid = year(lgi_dato) &"/"& month(lgi_dato)&"/"& day(lgi_dato) & " " & ig_sttid
					testig_sltid = year(lgi_dato) &"/"& month(lgi_dato)&"/"& day(lgi_dato) & " " & ig_sltid
					
					'Response.Write cdate(testloginTid) &" > "& cdate(testig_sttid) &" AND "& cdate(testloginTid) &" < "& cdate(testig_sltid) &" then <br>"
					
					use_ig_sltid = 0
					
                    '**** TEST er med i gruppe *** 'Cflow 20180823
                    sperretid_grpMember = 1
                    if len(trim(strSperretid_grp)) <> 0 then

                        call medariprogrpFn(medid)

                        if instr(strSperretid_grp, "-") = 0 then 'gælder kun for medlem af grupper
                            sperretid_grpMember = 0
                            strSperretid_grpArr = split(strSperretid_grp, ",")
                            for pi = 0 to ubound(strSperretid_grpArr)
                            
                                    if instr(medariprogrpTxt, "#"& trim(strSperretid_grpArr(pi)) &"#") <> 0 then
                                    sperretid_grpMember = 1
                                    end if
                            next
                        else  'gælder kun for IKKE medlem af disse grupper (hvis minus i kontrolpanel)
                    
                            sperretid_grpMember = 1 
                            strSperretid_grp = replace(strSperretid_grp, "-", "")
                            strSperretid_grpArr = split(strSperretid_grp, ",")
                            for pi = 0 to ubound(strSperretid_grpArr)
                            
                                    if instr(medariprogrpTxt, "#"& trim(strSperretid_grpArr(pi)) &"#") <> 0 then
                                    sperretid_grpMember = 0
                                    end if
                            next

                        end if


                    end if           


					if (cdate(testloginTid) > cdate(testig_sttid) AND cdate(testloginTid) < cdate(testig_sltid) AND sperretid_grpMember = 1) then


                        if lto = "cflow" AND instr(medariprogrpTxt, "#22#") <> 0 then 'ignorer Alligevel sperretid for medarbejdere i denne projektgruppe 20191015
                        use_ig_sltid = 0
					    manuelt_afsluttet = 0
					    loginTid = year(lgi_dato) &"/"& month(lgi_dato)&"/"& day(lgi_dato) & " " & lgi_sttid_hh &":"& lgi_sttid_mm & ":00"
                        else
		    			use_ig_sltid = 1
			    		loginTid = year(lgi_dato) &"/"& month(lgi_dato)&"/"& day(lgi_dato) & " " & ig_sltid
				    	manuelt_afsluttet = 3
					    end if
         
                    else
					use_ig_sltid = 0
					manuelt_afsluttet = 0
					loginTid = year(lgi_dato) &"/"& month(lgi_dato)&"/"& day(lgi_dato) & " " & lgi_sttid_hh &":"& lgi_sttid_mm & ":00"
					end if



					'Response.Write loginTid & "<hr>"


end function








'****************************************************************************
'**** FINDES DER ET login der allerede dækker over denne periode? 
'****************************************************************************

public errKonflikt, strLogindkonflikt
function stempelur_tidskonflikt(thisId, thisMid, kloginTid, klogudTid, kthisDato, io)
                
                

                errKonflikt = 0
                if io <> 1 then 
                '1 logind fra logind siden
                '2 logind fra terminal
                '3 Manuelt logind / eller logud via stempelur siden
                
                logIntidnn = datepart("n", kloginTid, 2, 2)
                if len(trim(logIntidnn)) = 1 then
                logIntidnn = "0"&logIntidnn
                end if


                logUdtidnn = datepart("n", klogudTid, 2, 2)
                if len(trim(logUdtidnn)) = 1 then
                logUdtidnn = "0"&logUdtidnn
                end if

                kloginTid = datepart("yyyy", kloginTid, 2, 2) & "-" & datepart("m", kloginTid, 2, 2) & "-" & datepart("d", kloginTid, 2, 2) & " " & datepart("h", kloginTid, 2, 2) & ":"& logIntidnn & ":00" 
                klogudTid = datepart("yyyy", klogudTid, 2, 2) & "-" & datepart("m", klogudTid, 2, 2) & "-" & datepart("d", klogudTid, 2, 2) & " " & datepart("h", klogudTid, 2, 2) & ":"& logUdtidnn & ":00" 
                
                kthisDato = datepart("yyyy", kthisDato, 2, 2) & "-" & datepart("m", kthisDato, 2, 2) & "-" & datepart("d", kthisDato, 2, 2)
                
                strSQL = "SELECT dato, login, logud FROM login_historik WHERE mid = "& thisMid &" AND dato = '"& kthisDato &"'"_
		        &" AND stempelurindstilling <> -1 AND id <> "& thisId &" AND minutter <> 0 AND ((login <= '"& kloginTid &"' AND logud > '"& kloginTid &"') "_
		        &" OR (login < '"& klogudTid &"' AND logud > '"& klogudTid &"') "_
		        &" OR (login >= '"& kloginTid &"' AND logud < '"& klogudTid &"')) " 'AND logud <= '"& klogudTid &"' 20180426
		        
                'if session("mid") = 1 then
		        'Response.Write "<br>KOnflikt SQL: "& strSQL & "<br>"
		        'Response.flush
                'end if
		        
		        
		        
		        oRec4.open strSQL, oConn, 3
		        if not oRec4.EOF then
		        
		            if io = 3 then '1 fra TO ellers fra Terminal
		                strLogindkonflikt = oRec4("login") & " til " & oRec4("logud") 
		                errortype = 134
		                call showError(errortype)
		                Response.end
		            else
		                errKonflikt = 1
		            end if
		        

		        end if
		        oRec4.close
		        
		        end if
		        
                'if session("mid") = 1 then
		        'Response.Write "<br>errKn:" & errKonflikt & "<br>"
                'Response.write "<br>klogudTid: "& klogudTid &"<br>logudTid B:" & logudTid
                'end if   


end function








'*************************************************************************
'*** Tjekker logind status ved logind i TimeOut eller fra terminal
'*************************************************************************

public fo_logud, fo_oprettetnytlogin, fo_autoafsluttet, fo_afsluttetlogin, hoursDiff, lastfoid, fo_reentry
function logindStatus(strUsrId, intStempelur, io, tid)

'if session("mid") = 1 then
'Response.write "fo_oprettetnytlogin<br>"
'end if

fo_afsluttetlogin = 0
fo_autoafsluttet = 0
fo_oprettetnytlogin = 0
errIndlaesTerminal = 0
'** io = indlæs / overfør 
'io = 1 Fra logind i Timeout via logindsiden eller intern Monitor Terminal CFLOW
'io = 2 Indlæses fra externe terminal. CST


              
if datepart("n", tid) < 10 then
nTid = "0"&datepart("n", tid)              
else
nTid = datepart("n", tid)
end if
                    
LoginDateTime = year(tid)&"/"& month(tid)&"/"&day(tid)&" "& datepart("h", tid) &":"& nTid &":00" 
LoginDato = year(tid)&"/"& month(tid)&"/"&day(tid)
			
              

                    '*** Henter seneste login / logud ***
				    strSQLlog = "SELECT id, dato, login, logud, mid, stempelurindstilling FROM login_historik WHERE "_
				    &" mid = "& strUsrId &" AND stempelurindstilling > -1 AND logud IS NULL ORDER BY login DESC limit 1" 
				    'id limit 0, 1"
                    '0, 10
    				
				    fo_logud = 0
				    
                    'if session("mid") = 1 then
				    'Response.write strSQLlog & "<br>"
				    'Response.end
                    'end if 
                    hoursDiff = 0
                    lastfoid = 0
				    mailissentonce = 0
				    oRec.open strSQLlog, oConn, 3
                    if not oRec.EOF then
                                
                                
                                lastfoid = oRec("id")        
                        
                       
                                '*** Autologud = Firma lukketid på dagen  ***'
                                '*** Lukker gammel logind **'
                                hoursDiff = 0
                                hoursDiff = datediff("h", oRec("login"), LoginDateTime, 2,2)
                                
                                'if session("mid") = 1 then
                                'Response.Write "<br>id: "& oRec("id") &" logind: "& oRec("login") &" og logud: "& LoginDateTime &",  hoursDiff:" & hoursDiff & "<br>"
                                'Response.end
                                'end if


                                'if cdate(LoginDato) <> cdate(oRec("dato")) then
                                '** Hvis der er mere end 20 timer mellem 20 logind, må det betragtes som at  **'
                                '** man har glemt at logge ind.                                              **'
                                '** TimeOut afslutter seneste logind med lukketid - tilføjer pauser og opretter et nyt logind. **'
                                '** AUTO LOGUD SKAL ALTID VÆRE SAMME DAG som LOGIN

                                if cdbl(hoursDiff) > 20 then
                                'logudDag = weekday(oRec("dato"), 2)
                                logudDag = datepart("w", oRec("dato"), 2, 2)
                                
                                'if session("mid") = 1 then
                                'Response.Write " ================================= "& logudDag &" dato:"& oRec("dato") &": wdn:"& weekdayname(datepart("w", oRec("dato"), 2, 2), 2,2) &"<br>"
                                'Response.end
                                'end if
                                
                                select case logudDag
                                case 1
                                dagkri = "normtid_sl_man"
                                case 2
                                dagkri = "normtid_sl_tir"
                                case 3
                                dagkri = "normtid_sl_ons"
                                case 4
                                dagkri = "normtid_sl_tor"
                                case 5
                                dagkri = "normtid_sl_fre"
                                case 6
                                dagkri = "normtid_sl_lor"
                                case 7
                                dagkri = "normtid_sl_son"
                                end select
                                
                                '*** Henter firmalukketid ****
                                lukketid = "23:59:00"
                                strSQL = "SELECT "& dagkri &" AS lukketid FROM licens WHERE id = 1 "
                                
                                'Response.Write strSQL
                                'Response.flush
                                oRec2.open strSQL, oConn, 3
                                if not oRec2.EOF then
                                    if len(trim(oRec2("lukketid"))) <> 0 then
                                    lukketid = oRec2("lukketid")
                                    else
                                    lukketid = "23:59:00"
                                    end if
                                End if
                                oRec2.close
                                
                                'Response.Write "<br>Lukketid: "& formatdatetime(lukketid, 3)
                                'Response.end
                                
                                LogudDateTime = year(oRec("login"))&"/"& month(oRec("login"))&"/"&day(oRec("login"))&" "& formatdatetime(lukketid, 3)
                        
                               


                               
                                '** Er der logget ind efter lukketid og glemt at blive loggetud? 
                                '** Logud tid = Logind tid 
                                '** Minutter = 0
                                'if cDate(loginTidAfr) > cDate(logudTidAfr) then
                                hoursDiffTjkLukketid = 0
                                hoursDiffTjkLukketid = datediff("n", LoginDateTime, oRec("login"), 2,2)
                                if hoursDiffTjkLukketid < 0 AND session("mid") = 1 then
                                LogudDateTime = year(oRec("login"))&"/"& month(oRec("login"))&"/"&day(oRec("login"))&" "& formatdatetime(oRec("login"), 3)
                                end if


                                '**** Minutter beregning ***
                                loginTidAfr = formatdatetime(oRec("login"), 3)
                                logudTidAfr = formatdatetime(LogudDateTime, 3)
                               

                               

                                minThisDIFF = datediff("s", loginTidAfr, logudTidAfr)/60
                                minThisDIFF = replace(formatnumber(minThisDIFF, 0), ".", "")
	                            minThisDIFF = replace(formatnumber(minThisDIFF, 0), ",", ".")
                                minThisDIFF = abs(minThisDIFF)
	                            
	                            '*** Er der tidkonflikt med manulet oprettede loginds fra TO ???
                                errKonflikt = 0
                                call stempelur_tidskonflikt(oRec("id"), strUsrId, oRec("login"), LogudDateTime, oRec("dato"), io)

                        
                                
                                'if session("mid") = 1 then 
                                'Response.Write "<br>Tidspunkter: "& cDate(loginTidAfr) &">"& cDate(logudTidAfr) & " LogudDateTime:  " & LogudDateTime & " hoursDiffTjkLukketid: " & hoursDiffTjkLukketid
                                'Response.end
                                'end if
         
                                'if session("mid") = 1 then
                                'Response.Write "<br> ================================= errKonflikt: "& errKonflikt &" logudDag dato:"& oRec("dato") &": wdn:"& weekdayname(datepart("w", oRec("dato"), 2, 2), 2,2) &"<br>"
                                'Response.end
                                'end if



                                '*** Der er ikke tidskonfikt og logud oprettes *****'
                                if cint(errKonflikt) <> 1 then 
	                            
	                                strSQLupd = "UPDATE login_historik SET "_
	                                &" logud = '"& LogudDateTime &"', minutter = "& minThisDIFF &", "_
	                                &" manuelt_afsluttet = 2, logud_first = '"& LogudDateTime &"' WHERE id = "& oRec("id")
    				                
                                    'if session("mid") = 1 then
				                    'Response.Write "<br>AUTO: "& strSQLupd & "<br>"
				                    'Response.end
                                    'end if
    				                
				                    oConn.execute(strSQLupd)
    				                fo_afsluttetlogin = 1
				                    fo_logud = 2
				                    

                                    '***** Opretter AUTO pause ****'
				                    'Response.Write "her"

                                    loginDTp = year(LogudDateTime) &"/"& month(LogudDateTime)&"/"& day(LogudDateTime)
                                    loginTidp = loginDTp & " 00:00:00"
                                    logudTidp = loginDTp & " 00:00:00"
                                    medarbSel = session("mid")

                                    '** Standard pauser fra Licens ****
                                    psDt = loginDTp & " 00:00:00" 'SKAL VÆRE DEN DATO DET OPRINDELIGE LOGIN ER FOERETAGET
                                    call stPauserFralicens(psDt)


                                    '*** Adgang for specielle projektgrupper ****'
                                    call stPauserProgrp(medarbSel, p1_grp, p2_grp, p3_grp, p4_grp)
				                    
                                    
                                    '***********************************************************               
	                                '*** Tilføj Pauser *****
                                    '***********************************************************

                                    if cint(p1_prg_on) = 1 OR cint(p2_prg_on) = 1 OR cint(p3_prg_on) = 1 OR cint(p4_prg_on) = 1 then
                                    '*** Tømmer pauser så der er altid kun MAKS er indlæst 2-4 pauser pr. dag pr. medarb.
                                    

                                    if len(trim(LogudDateTime)) <> 0 AND isNull(LogudDateTime) <> true AND isDate(LogudDateTime) = true then
                                    LoginDatoDelpau = year(LogudDateTime) &"/"& month(LogudDateTime) &"/"& day(LogudDateTime) '20150115 year(LoginDato) &"/"& month(LoginDato) &"/"& day(LoginDato)
	                                else
                                    LoginDatoDelpau = "2001-01-01"
                                    end if    

                                    'if len(trim(LoginDato)) <> 0 AND isNull(LoginDato) <> true AND isDate(LoginDato) = true then
                                    'LoginDatoDelpau = year(LoginDato) &"/"& month(LoginDato) &"/"& day(LoginDato)
	                                'else
                                    'LoginDatoDelpau = "2001-01-01"
                                    'end if                               
     
                                    
                                    if len(trim(medarbSel)) <> 0 then
                                    medarbSel = medarbSel
                                    else
                                    medarbSel = 0
                                    end if

                                    strSQLpDel = "DELETE FROM login_historik WHERE stempelurindstilling = -1 AND dato = '"& LoginDatoDelpau &"' AND mid = "& medarbSel
	                                'if io = "2" then ' (stempelut
                                    'Response.write strSQLpDel
                                    'Response.flush
                                    'end if
                                    oConn.execute(strSQLpDel)
                                    end if
	                                
                                    'Response.Write strSQLpDel & "<br>"
                                    p1 = stPauseLic_1
	                                if cint(p1_prg_on) = 1 AND cint(p1on) = 1 then 'if p1on <> 0 then
	                            
	                                    call tilfojPauser(0,medarbSel,loginDTp,loginTidp,logudTidp,p1,p1_komm)
					        
					                end if
					        
					                'Response.Write " p2on " & p2on
					                
                                    p2 = stPauseLic_2
					                  if cint(p2_prg_on) = 1 AND cint(p2on) = 1 then 'if p2on <> 0 then
					            
                                         call tilfojPauser(0,medarbSel,loginDTp,loginTidp,logudTidp,p2,p2_komm)
	                        
	                                end if

                                        
                                     p3 = stPauseLic_3
					                  if cint(p3_prg_on) = 1 AND cint(p3on) = 1 then 'if p2on <> 0 then
					            
                                         call tilfojPauser(0,medarbSel,loginDTp,loginTidp,logudTidp,p3,p3_komm)
	                        
	                                end if


                                     p4 = stPauseLic_4
					                  if cint(p4_prg_on) = 1 AND cint(p4on) = 1 then 'if p2on <> 0 then
					            
                                         call tilfojPauser(0,medarbSel,loginDTp,loginTidp,logudTidp,p4,p4_komm)
	                        
	                                end if
                
                                   


                                    'Response.end
                                                
				                    '** Send ikke mail
                                    ltotval = "xcflow"
                                    select case ltotval
                                    case "cflow"


                                    case else 

				                    '***** Oprettter Mail object ***
			                        if cint(mailissentonce) = 0 AND (request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\login.asp") _
			                        AND (request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\stempelur_terminal_ns.asp") then

 
			                        Set myMail=CreateObject("CDO.Message")
                                    myMail.Subject="TimeOut - Medarbejder har glemt at logge ud"
                                    myMail.From = "timeout_no_reply@outzource.dk"
				      
                                    
			                        'Modtagerens navn og e-mail
			                        select case lto
			                        case "dencker" 

                                    'myMail.To= "Anders Dencker <dv@dencker.net>"
                                    myMail.To= "Lon <lon@dencker.net>"
                                    myMail.Cc= "Timeout Support <support@outzource.dk>"

                                    'if len(trim(medarbEmail(x))) <> 0 then
                                    'myMail.To= ""& medarbNavn(x) &"<"& medarbEmail(x) &">"
                                    'end if       

			                        'Mailer.AddRecipient "Anders Dencker", "dv@dencker.net"
                                    'Mailer.AddCC "TimeOut Support", "sk@outzource.dk"

                                    call meStamdata(session("mid"))
                                    if len(trim(meEmail)) <> 0 then
                                    'Mailer.AddCC ""& meNavn &"", ""& meEmail &""
                                    myMail.Cc= ""& meNavn &"<"& meEmail &">"
                                    end if
                                    
			                        case "outz"
			                        'Mailer.AddRecipient "OutZourCE", "sk@outzource.dk"
                                    myMail.To= "Support <support@outzource.dk>"
                                    case "jttek"

                                    'Mailer.AddRecipient "JT-Teknik", "jt@jtteknik.dk"
                                    myMail.To= "JT-Teknik <jt@jtteknik.dk>"
                                    
                                    call meStamdata(session("mid"))
                                    if len(trim(meEmail)) <> 0 then
                                    'Mailer.AddCC ""& meNavn &"", ""& meEmail &""
                                    myMail.Cc= ""& meNavn &"<"& meEmail &">"
                                    end if

                                    'Mailer.AddCC "TimeOut Support", "sk@outzource.dk"

			                        case else
			                        'Mailer.AddRecipient "OutZourCE", "timeout_no_reply@outzource.dk"
                                    myMail.To= "Support <support@outzource.dk>"
			                        end select
                        			
			                       
                        			
			                                ' Selve teksten
					                        txtEmailHtml = "<br>Medarbejder "& session("user") &" har glemt at logge ud "& weekdayname(weekday(oRec("dato"), 1)) &" d. "& oRec("dato") &"<br><br>"_ 
					                        & "Der er oprettet en auto-logud tid der er sat til jeres firmas normale arbejdstid. (Sættes i kontrolpanelet)<br><br>"_ 
					                        & "Med venlig hilsen<br>TimeOut Stempelur Service<br>" 
					                        
					                      
                                               myMail.htmlBody= "<html><head><title></title>"_
			                                   &"<LINK rel=""stylesheet"" type=""text/css"" href=""http://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/style/timeout_style_print.css""></head>"_
			                                   &"<body>"_ 
			                                   & txtEmailHtml & "</body></html>"

                                           'myMail.AddAttachment "d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_14\inc\log\data\"& file

                                            myMail.Configuration.Fields.Item _
                                            ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
                                            'Name or IP of remote SMTP server
                                   
                                            if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
                                                smtpServer = "webout.smtp.nu"
                                            else
                                                smtpServer = "formrelay.rackhosting.com" 
                                            end if
                    
                                            myMail.Configuration.Fields.Item _
                                            ("http://schemas.microsoft.com/cdo/configuration/smtpserver")= smtpServer

                                            'Server port
                                            myMail.Configuration.Fields.Item _
                                            ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25
                                            myMail.Configuration.Fields.Update
                    
                                            if len(trim(meEmail)) <> 0 then
                                            myMail.Send
                                            end if
                                            set myMail=nothing


                        				
                        			mailissentonce = 1

                               
			                        end if ''** Mail
                                    end select 'lto
				                
				                
				                else
				                
				                '** Ignorer logind pga tidskonflikt //Dvs. at man har oprettet/afsluttet et manuelt logind i TimeOut // browser har været gået ned
                                fo_logud = 3
				                end if
				                                
				                                
				                
				                else ' cint(hoursDiff) <= 20 then
				                
                                '** Der findes et igangværende uafsluttet logind ***'
				                '** CST Ekstern terminal, så skal alle logind overføres uanset interval --> login_historik_terminal
                                '** CFLOW Ekstern terminal, så skal alle logind overføres uanset interval.
				                '** Kun fra TimeOut, hvis f.eks browser er genstartet, eller CPU gået ned.
				                '** Så skal der ikke oprettes et nyt logind. **'
				                    
                                    '** IO settings ****
                                    '** 1 fra TO logind siden / Monitor (Cflow) 
				                    '** 2 fra Ekstern Terminal fra DB tabel Login_historik_terminal (overførsel)
				                    '** 3 fra Stempelurhistorik / komme / gå tider siden manuelt eller logiud (bruges ikke nov. 2012) Disse indlæases via stempelur.asp -> indlæs
                                    if io = 1 then 'fra login siden i TO 
				                    '** Indlæser INGEN TING da man logger ind via TO, og er er et igangværende logind
				                    
                                              
                                    fo_reentry = 1
                                    fo_logud = 3

                                   
				                    else 'Fra LoginSide + Terminal
                                    '*** Opretter nyt eller afslutter eksisterende
				                    
				                            fo_autoafsluttet = 1
				                            
                                            '*** ÆNDRET 20180517 CFLOW ***
                                            '*** Der kan være logget ud efter 24. Derfor behøver logud dato ikke være den samme som logind dato. Autologud > 20 timer. SKAL ALTID VÆRE SAMEM DATO
                                            '*** = year(LogudDateTime) &"/"& month(LogudDateTime)&"/"& day(LogudDateTime)     

                
                                            'if cdbl(strUsrId) = 1 then
                                            'LogudDateTimeDDMMYYY_TIME = "8-"& month(LoginDateTime)&"-"&year(LoginDateTime)&" 00:32:00" 
                                            'LogudDateTime = "8-"& month(LoginDateTime)&"-"&year(LoginDateTime)&" 00:32:00"  'LoginDateTime    
                                            'else
				                            LogudDateTimeDDMMYYY_TIME = day(LoginDateTime)&"-"& month(LoginDateTime)&"-"&year(LoginDateTime)&" "& formatdatetime(LoginDateTime, 3)
                                            LogudDateTime = LoginDateTime                      
                                            'end if

                                            '**** Minutter beregning ***
                                            loginTidAfr = oRec("login") 'formatdatetime(oRec("login"), 3)
                                            logudTidAfr = LogudDateTimeDDMMYYY_TIME 'formatdatetime(LogudDateTimeDDMMYYY_TIME, 3)
                               
                                            'Response.Write "<br>Min: "& oRec("login") & " - "& LogudDateTimeDDMMYYY_TIME
                               
                                            minThisDIFF = datediff("s", loginTidAfr, logudTidAfr)/60
                                            minThisDIFF = replace(formatnumber(minThisDIFF, 0), ".", "")
	                                        minThisDIFF = replace(formatnumber(minThisDIFF, 0), ",", ".")
                                            minThisDIFF = abs(minThisDIFF)
	                            
	                            
	                                        '** Ignorer Terminal hvis der er oprettet et manuelt fra TO i samme peridode **'
	                                        if cDate(LogudDateTimeDDMMYYY_TIME) >= cDate(oRec("login")) then
	                            
	                                            '*** Er der tidkonflikt med manulet oprettede loginds fra TO ???
                                                errKonflikt = 0
                                                call stempelur_tidskonflikt(oRec("id"), strUsrId, oRec("login"), LogudDateTime, oRec("dato"), io)
                                    
                                                if cint(errKonflikt) <> 1 then 
    	                            

                                                '*** Afslutter eksisterende logind ****' 
	                                            strSQLupd = "UPDATE login_historik SET "_
	                                            &" logud = '"& LogudDateTime &"', minutter = "& minThisDIFF &",  "_
	                                            &" logud_first = '"& LogudDateTime &"' WHERE id = "& oRec("id")
    				                
				                                'Response.Write "<br>Alm. logud: "& strSQLupd & "<br>"
				                                'Response.flush
    				                
				                                oConn.execute(strSQLupd)
                                                fo_afsluttetlogin = 1
                                                


                                                                '** PAUSER ***
                                                                psloginTidp = LoginDato
                                                                pslogudTidp = psloginTidp



                                                                call stPauserFralicens(LoginDato)

                                                                '** tjekker om der skal tilføjes pause til projektgruppe ***' 
                                                                call stPauserProgrp(strUsrId, p1_grp, p2_grp, p3_grp, p4_grp)

                                                    
                                                                if cint(p1_prg_on) = 1 OR cint(p2_prg_on) = 1 OR cint(p3_prg_on) = 1 OR cint(p4_prg_on) = 1 then
                                                                '**** Tømmer pauser så der er altid kun er indlæst 2 pauser pr. dag pr. medarb. ****
                                                                LoginDatoDelpau = year(LoginDato) &"/"& month(LoginDato) &"/"& day(LoginDato)
                                                                strSQLpDel = "DELETE FROM login_historik WHERE stempelurindstilling = -1  AND dato = '"& LoginDatoDelpau &"' AND mid = "& strUsrId
	                                                            'if io = 2 then
                                                                'Response.write strSQLpDel
                                                                'Response.end
                                                                'end if
                                                                oConn.execute(strSQLpDel)
                                                                end if

                                                                if cint(p1_prg_on) = 1 AND cint(p1on) = 1 then
                                                                psKomm_1 = ""
                                                                psMin_1 = stPauseLic_1

                                                                '** p1 **
                                                                call tilfojPauser(0,strUsrId,LoginDato,psloginTidp,pslogudTidp,psMin_1,psKomm_1)

                                                                end if


                                                                'Response.write "her"
                                                                'Response.end

                                                                if cint(p2_prg_on) = 1 AND cint(p2on) = 1 then
                                                                psKomm_2 = ""
                                                                psMin_2 = stPauseLic_2

                                                                '** p2 **
                                                                call tilfojPauser(0,strUsrId,LoginDato,psloginTidp,pslogudTidp,psMin_2,psKomm_2)

                                                                end if


                                                                '** IF lto DENCKER
                                                                if cint(p3_prg_on) = 1 AND cint(p3on) = 1 then
                                                                psKomm_3 = ""
                                                                psMin_3 = stPauseLic_3

                                                                '** p3 **
                                                                call tilfojPauser(0,strUsrId,LoginDato,psloginTidp,pslogudTidp,psMin_3,psKomm_3)

                                                                end if

                                                                if cint(p4_prg_on) = 1 AND cint(p4on) = 1 then
                                                                psKomm_4 = ""
                                                                psMin_4 = stPauseLic_4

                                                                '** p3 **
                                                                call tilfojPauser(0,strUsrId,LoginDato,psloginTidp,pslogudTidp,psMin_4,psKomm_4)

                                                                end if

                                                   
    				                
				                                end if 'tidskonflikt 
				                
				                            end if 'LogudDateTimeDDMMYYY_TIME
				                
				                    fo_logud = 3   
				                    
				                    end if 'io
				                    
				                end if '20 timersDiff
                         
                        
                        
                    'oRec.movenext   
                    end if
                    oRec.close
				
				
				'Response.Write "<br>fo_logud" & fo_logud & "<br>"
				'Response.end
				
				
				'*** Hvis der ikke er et igangværende logind, oprettes et nyt. **'
				'** fo_logud = 0 (indlæs ny) **'
                if cint(fo_logud) <> 3 then
				
				
				'** Ignorer periode skal ikke kaldes, da man så 
				'** altid bliver tvunget til at angive en kommentar for ændret logind,
				'** ved logind i et interval der skal ignores. 
				
				lgi_sttid_hh = datepart("h", tid) 
				lgi_sttid_mm = datepart("n", tid)

				call stpauser_ignorePer(LoginDato,lgi_sttid_hh,lgi_sttid_mm, strUsrId)
			    
			    
			    if thisfile <> "stempelur_terminal_ns" then
			    ipn = request.servervariables("REMOTE_ADDR")
			    else
			    ipn = 0
			    end if
				
				
    				'if session("mid") <> 1 then
				            strSQL = "INSERT INTO login_historik "_
				            &"(dato, login, mid, stempelurindstilling, login_first, ipn, manuelt_afsluttet) "_
				            &" VALUES ('"& LoginDato &"', '"& LoginTid &"', "& strUsrId &", "_
				            &" "& intStempelur &", '"& LoginDateTime &"', '"& ipn &"', "& manuelt_afsluttet &")"

                            'if session("mid") = 1 then
				            'Response.write "<br>ins SQL: "& strSQL & "<br>"
		                    'end if		    

                            
         
                            oConn.execute(strSQL)
                    'end if

                 
				    fo_oprettetnytlogin = 1 'fo_logud <> 3
                  
                  '** 2012 nov. 21
                  '** Der SKAL ALTID tilføjes PAUSER, de bliver slettet igen ved afslut logind, 
                  '** De skal oprettes, så de ligger åbne i logind historikken, hvia folk vælger at afslutte manuelt via TO senere på dagen. Hvis ikke der fidnes pauser vil der ikke blive indlæst pauser på denne dato.
                  

                                                    '** PAUSER ***
                                                    psloginTidp = LoginDato
                                                    pslogudTidp = psloginTidp


                                                    
                                                    call stPauserFralicens(LoginDato)

                                                    '** tjekker om der skal tilføjes pause til projektgruppe ***' 
                                                    call stPauserProgrp(strUsrId, p1_grp, p2_grp, p3_grp, p4_grp)

                                                    
                                                    
                                                    if cint(p1_prg_on) = 1 OR cint(p2_prg_on) = 1 OR cint(p3_prg_on) = 1 OR cint(p4_prg_on) = 1 then
                                                    '**** Tømmer pauser så der er altid kun er indlæst 2 pauser pr. dag pr. medarb. ****
                                                    LoginDatoDelpau = year(LoginDato) &"/"& month(LoginDato) &"/"& day(LoginDato)
                                                    strSQLpDel = "DELETE FROM login_historik WHERE stempelurindstilling = -1  AND dato = '"& LoginDatoDelpau &"' AND mid = "& strUsrId
	                                                oConn.execute(strSQLpDel)
                                                    end if

                                                    if cint(p1_prg_on) = 1 AND cint(p1on) = 1 then
                                                    psKomm_1 = ""
                                                    psMin_1 = stPauseLic_1

                                                    
                                                    '** p1 **
                                                    call tilfojPauser(0,strUsrId,LoginDato,psloginTidp,pslogudTidp,psMin_1,psKomm_1)

                                                    end if


                                                    if cint(p2_prg_on) = 1 AND cint(p2on) = 1 then
                                                    psKomm_2 = ""
                                                    psMin_2 = stPauseLic_2

                                                    '** p2 **
                                                    call tilfojPauser(0,strUsrId,LoginDato,psloginTidp,pslogudTidp,psMin_2,psKomm_2)

                                                    end if


                                                     if cint(p3_prg_on) = 1 AND cint(p3on) = 1 then
                                                    psKomm_3 = ""
                                                    psMin_3 = stPauseLic_3

                                                    '** p3 **
                                                    call tilfojPauser(0,strUsrId,LoginDato,psloginTidp,pslogudTidp,psMin_3,psKomm_3)

                                                    end if



                                                     if cint(p4_prg_on) = 1 AND cint(p4on) = 1 then
                                                    psKomm_4 = ""
                                                    psMin_4 = stPauseLic_4

                                                    '** p4 **
                                                    call tilfojPauser(0,strUsrId,LoginDato,psloginTidp,pslogudTidp,psMin_4,psKomm_4)

                                                    end if

    				
				  
				
				end if


                'if io = 2 then
                'response.write "fo_logud" & fo_logud & "p1_prg_on: " & p1_prg_on & "p2_prg_on: " & p2_prg_on
                'Response.end
                'end if

end function
















public showkgtim, showkgpau, showkgtil, showkgtot, showkgnor, showkgsal, showkguds, showkgsaa, showextended
function stempelur_kolonne(lto, showextended)


    showextended = 0


    select case lcase(lto)
    case "intranet - local"

    showkgtim = 1
    showkgpau = 1
    showkgtil = 1
    showkgtot = 1
    showkgnor = 1
    showkgsal = 1
    showkguds = 0
    showkgsaa = 1

    case "epi2017"

    showkgtim = 0
    showkgpau = 0
    showkgtil = 0
    showkgtot = 0
    showkgnor = 0
    showkgsal = 0
    showkguds = 0
    showkgsaa = 0

    case "fk", "fk_bpm"

    showkgtim = 1
    showkgpau = 0
    showkgtil = 1
    showkgtot = 0
    showkgnor = 1
    showkgsal = 1
    showkguds = 0
    showkgsaa = 1
    
    case "kejd_pb", "kejd_pb2"

    showkgtim = 1
    showkgpau = 0
    showkgtil = 0
    showkgtot = 0
    showkgnor = 1
    showkgsal = 1
    showkguds = 0
    showkgsaa = 1

    case "dencker", "jttek", "tec", "esn"

    showkgtim = 1
    showkgpau = 1
    showkgtil = 1
    showkgtot = 1
    showkgnor = 1
    showkgsal = 1
    showkguds = 0
    showkgsaa = 1

    case "cflow"

    showkgtim = 1
    showkgpau = 1
    showkgtil = 1
    showkgtot = 1
    showkgnor = 1
    showkgsal = 1
    showkguds = 0
    showkgsaa = 1

    case "cisu", "intranet - local"

    showkgtim = 0
    showkgpau = 0
    showkgtil = 0
    showkgtot = 0
    showkgnor = 0
    showkgsal = 0
    showkguds = 0
    showkgsaa = 0

    case "tia"

    showkgtim = 0
    showkgpau = 0
    showkgtil = 0
    showkgtot = 0
    showkgnor = 0
    showkgsal = 0
    showkguds = 0
    showkgsaa = 0

    case else

    showkgtim = 1
    showkgpau = 1
    showkgtil = 1
    showkgtot = 1
    showkgnor = 1
    showkgsal = 1
    showkguds = 1
    showkgsaa = 1

    end select


end function





public totaltimerPer, totalpausePer, totalTimerPer100, loginTimerTot
public manMin, tirMin, onsMin, torMin, freMin, lorMin, sonMin, totMan, totTir, totOns, totTor, totFre, totLor, totSon
public manMinPause, tirMinPause, onsMinPause, torMinPause, freMinPause, lorMinPause, sonMinPause
Public manFraTimer, tirFraTimer, onsFraTimer, torFraTimer, freFraTimer, lorFraTimer, sonFraTimer
function fLonTimerPer(stDato, periode, visning, medid)

'Response.Write "////////her " & stDato & " Periode: " & periode & " visning: "& visning & " medid: "& medid & " rdir:"& rdir & "<hr>"
'Response.end

          manMin = 0
    manMinPause = 0
          tirMin = 0
    tirMinPause = 0
          onsMin = 0
    onsMinPause = 0
          torMin = 0
    torMinPause = 0
          freMin = 0
    freMinPause = 0
          lorMin = 0
    lorMinPause = 0
          sonMin = 0
    sonMinPause = 0



    totMan = 0
	totTir = 0
	totOns = 0
	totTor = 0
	totFre = 0
	totLor = 0
	totSon = 0


totaltimerPer = 0
totalpausePer = 0
ugeIaltFraTilTimer = 0

slutDato = dateadd("d", periode, stDato)
'weekDiff = datediff("ww", stDato, slutDato, 2, 2)

         call meStamdata(medid)

         '** Er ansat dato efter statdato i interval ****'
         if cdate(meAnsatDato) <= cDate(stdato) then
                
                weekDiff = dateDiff("ww", stdato, slutDato, 2, 2)

         else

                weekDiff = dateDiff("ww", meAnsatDato, slutDato, 2, 2)
	
	     end if

        
            if len(weekDiff) <> 0 AND weekDiff <> 0 then
	        weekDiff = cint(weekDiff)
	        else
	        weekDiff = 1
	        end if





if visning = 0 then '** Stempelur faneblad på timereg siden
    call akttyper2009(2)
end if


'*** Finder navne på til/fra typer **'
		                
akttyperTFh = replace(aty_sql_tilwhours, "t.tfaktim = 0", "")
akttyperTF = split(akttyperTFh, " OR t.tfaktim = ")
for tf = 0 to UBOUND(akttyperTF)

'Response.Write akttyperTF(tf) & "<br>"

if len(trim(akttyperTF(tf))) <> 0 then
call akttyper(akttyperTF(tf),0)

if tf = 0 then
akttypenavnTil = akttypenavn
else
akttypenavnTil = akttypenavnTil &", "& akttypenavn
end if
'Response.Write akttyperTF(tf) & "<br>"
end if

next


		                
akttyperTFr = replace(aty_sql_frawhours, "t.tfaktim = 0", "")
akttyperTF = split(akttyperTFr, " OR t.tfaktim = ")
for tf = 0 to UBOUND(akttyperTF)

'Response.Write akttyperTF(tf) & "<br>"

if len(trim(akttyperTF(tf))) <> 0 then
call akttyper(akttyperTF(tf),0)

if tf = 0 then
akttypenavnFra = akttypenavn
else
akttypenavnFra = akttypenavnFra & ", "& akttypenavn
end if
'Response.Write akttyperTF(tf) & "<br>"
end if

next

'*******


'Response.Write stDato & ", "& periode & "meid: " & medid & " weekdiff: "& weekdiff & "<hr>"

	'**** login historik (denne uge/ Periode) ****

    

	for intcounter = 0 to periode
	
                    'if visning <> 21 then

					select case intcounter
					case 0
					useSQLd = stDato
					case else
					tDat = dateadd("d", 1, useSQLd)
					useSQLd = year(tDat) &"/"& month(tDat) & "/"& day(tDat)
					end select

                    'else
                
                    'select case intcounter
					'case 0
					'useSQLd = stDato
                    'case 4 'tilbage til sidste fredag
					'tDat = dateadd("d", -6, useSQLd)
					'useSQLd = year(tDat) &"/"& month(tDat) & "/"& day(tDat)
					'case else
					'tDat = dateadd("d", 1, useSQLd)
					'useSQLd = year(tDat) &"/"& month(tDat) & "/"& day(tDat)
					'end select


                    'end if
					
					
		            useSQLd = year(useSQLd) & "/"& month(useSQLd) &"/"& day(useSQLd)			

					strSQL = "SELECT l.id AS lid, l.mid AS lmid, l.login, l.logud, l.minutter, "_
					&" s.navn AS stempelurnavn, s.faktor, s.minimum, stempelurindstilling FROM login_historik l"_
					&" LEFT JOIN stempelur s ON (s.id = l.stempelurindstilling) WHERE "_
					&" l.dato = '"& useSQLd &"' AND l.mid = " & medid &""_
					&" ORDER BY l.login" 
					
					'select case lto 
                    'case "dencker"
                    'Response.Write "<br>SQL:"& strSQL & "<br><br>"
					'case else
                    'end select
					
					f = 0
					oRec.open strSQL, oConn, 3 
					while not oRec.EOF 
					
						timerThis = 0
						timerThisDIFF = 0
						timerThisPause = 0

						if len(oRec("login")) <> 0 AND len(oRec("logud")) <> 0 then
						
						'loginTidAfr = left(formatdatetime(oRec("login"), 3), 5)
						'logudTidAfr = left(formatdatetime(oRec("logud"), 3), 5)
						
						        if cint(oRec("stempelurindstilling")) = -1 then
						    
						            timerThisDIFF = oRec("minutter")
						            useFaktor = 0
						    
						            timerThisPause = timerThisDIFF
						            timerThis = 0
                                    'Response.write oRec("minutter") & "<br>"
						    
						        else 
						    
						            timerThisDIFF = oRec("minutter") 'datediff("s", loginTidAfr, logudTidAfr)/60
						
						            if timerThisDIFF < oRec("minimum") then
							            timerThisDIFF = oRec("minimum")
						            end if
						
							
							        if oRec("faktor") > 0 then
							        useFaktor = oRec("faktor")
							        else
							        useFaktor = 0
							        end if
							
							        timerThisPause = 0
							        timerThis = (timerThisDIFF * useFaktor)
							
						        end if
						
						
						
						totaltimerPer = totaltimerPer + timerThis
						totalpausePer = totalpausePer + timerThisPause
						'Response.write oRec("lid") & ": " &  timerThis &" - "
						else
						
						totaltimerPer = totaltimerPer
						totalpausePer = totalpausePer 
						
						end if
						
						
						
						
						if visning = 0 OR visning = 20 OR visning = 21 then
					    
						select case intcounter
						case 0
						manMin = manMin + timerThis 
						manMinPause = manMinPause + timerThisPause 
						case 1
						tirMin = tirMin + timerThis 
						tirMinPause = tirMinPause + timerThisPause  
						case 2
						onsMin = onsMin + timerThis
						onsMinPause = onsMinPause + timerThisPause   
						case 3
						torMin = tormin + timerThis
						torMinPause = torminPause + timerThisPause   
						case 4
						freMin = freMin + timerThis
						freMinPause = freMinPause + timerThisPause   
						case 5
						lorMin = lorMin + timerThis
						lorMinPause = lorMinPause + timerThisPause  
						case 6
						sonMin = sonMin + timerThis
						sonMinPause = sonMinPause + timerThisPause   
						end select 
						
						
						end if
					    
					    'Response.write "tot:"& intcounter &" - "& totaltimer &":"& totalpause
						
						
					oRec.movenext
					wend
					oRec.close 
					
					f = 0
					
					
					
					    '*** Tillæg / Fradrag via Realtimer **'
					    if visning = 0 OR visning = 20 then
                	
		                
                		
		                tiltimer = 0
		                strSQL2 = "SELECT sum(t.timer) AS tiltimer FROM timer t WHERE "_
		                &" (tmnr = "& medid &" AND tdato = '"& useSQLd &"' AND ("& aty_sql_tilwhours &")) GROUP BY tmnr " 
		                
                        'if session("mid") = 1 then
                        'Response.Write strSQL2 & "<br>###<br>"
		                'Response.flush
                        'end if
                		
		                oRec2.open strSQL2, oConn, 3 
		                if not oRec2.EOF then
			                tiltimer = oRec2("tiltimer") * 60
		                end if
		                oRec2.close 
                		
		                if len(trim(tiltimer)) <> 0 then
		                tiltimer = tiltimer
		                else
		                tiltimer = 0
		                end if
                		
                		 
                		
                		
                		fradtimer = 0
		                strSQL2 = "SELECT sum(timer) AS fratimer FROM timer t WHERE "_
		                &" (tmnr = "& medid &" AND tdato = '"& useSQLd &"' AND ("& aty_sql_frawhours &")) GROUP BY tmnr "
		                
                        'if session("mid") = 1 then
                        'Response.Write strSQL2
		                'Response.flush
                        'end if
                		
		                oRec2.open strSQL2, oConn, 3 
		                if not oRec2.EOF then
			                fradtimer = oRec2("fratimer") * 60
		                end if
		                oRec2.close 
                		
		                if len(trim(fradtimer)) <> 0 then
		                fradtimer = fradtimer
		                else
		                fradtimer = 0
		                end if
		                
		                
		                select case intcounter
						case 0
						manFraTimer = (tiltimer - fradtimer)
						case 1
						tirFraTimer = (tiltimer - fradtimer)
						case 2
						onsFraTimer = (tiltimer - fradtimer)
						case 3
						torFraTimer = (tiltimer - fradtimer)
						case 4
						freFraTimer = (tiltimer - fradtimer)
						case 5
						lorFraTimer = (tiltimer - fradtimer)
						case 6
						sonFraTimer = (tiltimer - fradtimer)
						end select 
						
		                
		                
		                '*** Fleks ***'
                		call akttyper2009prop(7)
                		aty_fleks_on = aty_on
                		aty_fleks_tf = aty_tfval
                		if cint(aty_fleks_on) = 1 then
		                '*** Fleks ****'   
                	    flekstimer = 0
		                strSQL2 = "SELECT sum(t.timer) AS flekstimer FROM timer t WHERE "_
		                &" (tmnr = "& medid &" AND tdato = '"& useSQLd &"' AND tfaktim = 7) GROUP BY tmnr " 
		                'Response.Write strSQL2
		                'Response.flush
                		
		                oRec2.open strSQL2, oConn, 3 
		                if not oRec2.EOF then
			                flekstimer = oRec2("flekstimer") * 60
		                end if
		                oRec2.close 
                		
                		
                		
		                if len(trim(flekstimer)) <> 0 then
		                flekstimer = flekstimer
		                else
		                flekstimer = 0
		                end if
		                
		                select case intcounter
						case 0
						manflekstimer = (flekstimer)
						case 1
						tirflekstimer = (flekstimer)
						case 2
						onsflekstimer = (flekstimer)
						case 3
						torflekstimer = (flekstimer)
						case 4
						freflekstimer = (flekstimer)
						case 5
						lorflekstimer = (flekstimer)
						case 6
						sonflekstimer = (flekstimer)
						end select 
						
						
						end if
						
						
						'*** Ferie ***'
                		call akttyper2009prop(14)
                		aty_Ferie_on = aty_on
                		aty_Ferie_tf = aty_tfval
                		if cint(aty_Ferie_on) = 1 then
		                '*** Ferie ****'   
                	    Ferietimer = 0
		                strSQL2 = "SELECT sum(t.timer) AS Ferietimer FROM timer t WHERE "_
		                &" (tmnr = "& medid &" AND tdato = '"& useSQLd &"' AND tfaktim = 14) GROUP BY tmnr " 
		                'Response.Write strSQL2
		                'Response.flush
                		
		                oRec2.open strSQL2, oConn, 3 
		                if not oRec2.EOF then
			                Ferietimer = oRec2("Ferietimer") * 60
		                end if
		                oRec2.close 
                		
                		
                		
		                if len(trim(Ferietimer)) <> 0 then
		                Ferietimer = Ferietimer
		                else
		                Ferietimer = 0
		                end if
		                
		                select case intcounter
						case 0
						manFerietimer = (Ferietimer)
						case 1
						tirFerietimer = (Ferietimer)
						case 2
						onsFerietimer = (Ferietimer)
						case 3
						torFerietimer = (Ferietimer)
						case 4
						freFerietimer = (Ferietimer)
						case 5
						lorFerietimer = (Ferietimer)
						case 6
						sonFerietimer = (Ferietimer)
						end select 
						
						
						end if
						
						
						
						'*** Syg ***'
                		call akttyper2009prop(20)
                		aty_Syg_on = aty_on
                		aty_Syg_tf = aty_tfval
                		if cint(aty_Syg_on) = 1 then
		                '*** Syg ****'   
                	    Sygtimer = 0
		                strSQL2 = "SELECT sum(t.timer) AS Sygtimer FROM timer t WHERE "_
		                &" (tmnr = "& medid &" AND tdato = '"& useSQLd &"' AND tfaktim = 20) GROUP BY tmnr " 
		                'Response.Write strSQL2
		                'Response.flush
                		
		                oRec2.open strSQL2, oConn, 3 
		                if not oRec2.EOF then
			                Sygtimer = oRec2("Sygtimer") * 60
		                end if
		                oRec2.close 
                		
                		
                		
		                if len(trim(Sygtimer)) <> 0 then
		                Sygtimer = Sygtimer
		                else
		                Sygtimer = 0
		                end if
		                
		                select case intcounter
						case 0
						manSygtimer = (Sygtimer)
						case 1
						tirSygtimer = (Sygtimer)
						case 2
						onsSygtimer = (Sygtimer)
						case 3
						torSygtimer = (Sygtimer)
						case 4
						freSygtimer = (Sygtimer)
						case 5
						lorSygtimer = (Sygtimer)
						case 6
						sonSygtimer = (Sygtimer)
						end select 
						
						
						end if
						
						
						'*** BarnSyg ***'
                		call akttyper2009prop(21)
                		aty_BarnSyg_on = aty_on
                		aty_BarnSyg_tf = aty_tfval
                		if cint(aty_BarnSyg_on) = 1 then
		                '*** BarnSyg ****'   
                	    BarnSygtimer = 0
		                strSQL2 = "SELECT sum(t.timer) AS BarnSygtimer FROM timer t WHERE "_
		                &" (tmnr = "& medid &" AND tdato = '"& useSQLd &"' AND tfaktim = 21) GROUP BY tmnr " 
		                'Response.Write strSQL2
		                'Response.flush
                		
		                oRec2.open strSQL2, oConn, 3 
		                if not oRec2.EOF then
			                BarnSygtimer = oRec2("BarnSygtimer") * 60
		                end if
		                oRec2.close 
                		
                		
                		
		                if len(trim(BarnSygtimer)) <> 0 then
		                BarnSygtimer = BarnSygtimer
		                else
		                BarnSygtimer = 0
		                end if
		                
		                select case intcounter
						case 0
						manBarnSygtimer = (BarnSygtimer)
						case 1
						tirBarnSygtimer = (BarnSygtimer)
						case 2
						onsBarnSygtimer = (BarnSygtimer)
						case 3
						torBarnSygtimer = (BarnSygtimer)
						case 4
						freBarnSygtimer = (BarnSygtimer)
						case 5
						lorBarnSygtimer = (BarnSygtimer)
						case 6
						sonBarnSygtimer = (BarnSygtimer)
						end select 
						
						
						end if
						
						
						'*** Lage ***'
                		call akttyper2009prop(81)
                		aty_Lage_on = aty_on
                		aty_Lage_tf = aty_tfval
                		if cint(aty_Lage_on) = 1 then
		                '*** Lage ****'   
                	    Lagetimer = 0
		                strSQL2 = "SELECT sum(t.timer) AS Lagetimer FROM timer t WHERE "_
		                &" (tmnr = "& medid &" AND tdato = '"& useSQLd &"' AND tfaktim = 81) GROUP BY tmnr " 
		                'Response.Write strSQL2
		                'Response.flush
                		
		                oRec2.open strSQL2, oConn, 3 
		                if not oRec2.EOF then
			                Lagetimer = oRec2("Lagetimer") * 60
		                end if
		                oRec2.close 
                		
                		
                		
		                if len(trim(Lagetimer)) <> 0 then
		                Lagetimer = Lagetimer
		                else
		                Lagetimer = 0
		                end if
		                
		                select case intcounter
						case 0
						manLagetimer = (Lagetimer)
						case 1
						tirLagetimer = (Lagetimer)
						case 2
						onsLagetimer = (Lagetimer)
						case 3
						torLagetimer = (Lagetimer)
						case 4
						freLagetimer = (Lagetimer)
						case 5
						lorLagetimer = (Lagetimer)
						case 6
						sonLagetimer = (Lagetimer)
						end select 
						
						
						end if
						
						
						'*** Sund ***'
                		call akttyper2009prop(8)
                		aty_Sund_on = aty_on
                		aty_Sund_tf = aty_tfval
                		if cint(aty_Sund_on) = 1 then
		                '*** Sund ****'   
                	    Sundtimer = 0
		                strSQL2 = "SELECT sum(t.timer) AS Sundtimer FROM timer t WHERE "_
		                &" (tmnr = "& medid &" AND tdato = '"& useSQLd &"' AND tfaktim = 8) GROUP BY tmnr " 
		                'Response.Write strSQL2
		                'Response.flush
                		
		                oRec2.open strSQL2, oConn, 3 
		                if not oRec2.EOF then
			                Sundtimer = oRec2("Sundtimer") * 60
		                end if
		                oRec2.close 
                		
                		
                		
		                if len(trim(Sundtimer)) <> 0 then
		                Sundtimer = Sundtimer
		                else
		                Sundtimer = 0
		                end if
		                
		                select case intcounter
						case 0
						manSundtimer = (Sundtimer)
						case 1
						tirSundtimer = (Sundtimer)
						case 2
						onsSundtimer = (Sundtimer)
						case 3
						torSundtimer = (Sundtimer)
						case 4
						freSundtimer = (Sundtimer)
						case 5
						lorSundtimer = (Sundtimer)
						case 6
						sonSundtimer = (Sundtimer)
						end select 
						
						
						end if
                	    
                	    
                	    '*** Frokost ***'
                		call akttyper2009prop(10)
                		aty_Frokost_on = aty_on
                		aty_Frokost_tf = aty_tfval
                		if cint(aty_Frokost_on) = 1 then
		                '*** Frokost ****'   
                	    Frokosttimer = 0
		                strSQL2 = "SELECT sum(t.timer) AS Frokosttimer FROM timer t WHERE "_
		                &" (tmnr = "& medid &" AND tdato = '"& useSQLd &"' AND tfaktim = 10) GROUP BY tmnr " 
		                'Response.Write strSQL2
		                'Response.flush
                		
		                oRec2.open strSQL2, oConn, 3 
		                if not oRec2.EOF then
			                Frokosttimer = oRec2("Frokosttimer") * 60
		                end if
		                oRec2.close 
                		
                		
                		
		                if len(trim(Frokosttimer)) <> 0 then
		                Frokosttimer = Frokosttimer
		                else
		                Frokosttimer = 0
		                end if
		                
		                select case intcounter
						case 0
						manFrokosttimer = (Frokosttimer)
						case 1
						tirFrokosttimer = (Frokosttimer)
						case 2
						onsFrokosttimer = (Frokosttimer)
						case 3
						torFrokosttimer = (Frokosttimer)
						case 4
						freFrokosttimer = (Frokosttimer)
						case 5
						lorFrokosttimer = (Frokosttimer)
						case 6
						sonFrokosttimer = (Frokosttimer)
						end select 
						
						
						end if
	              
	                end if 'visning fradrag / tillæg
					
					
					
	
	next
	
	   
	

    'if session("mid") = 1 then
    'Response.write "HER: " & tirMin & " - ( " & tirMinPause & " - ( " & tirFraTimer  &"))"
    'end if

    totMan = manMin - (manMinPause - (manFraTimer))
	totTir = tirMin - (tirMinPause - (tirFraTimer))
	totOns = onsMin - (onsMinPause - (onsFraTimer))
	totTor = torMin - (torMinPause - (torFraTimer))
	totFre = freMin - (freMinPause - (freFraTimer))
	totLor = lorMin - (lorMinPause - (lorFraTimer))
	totSon = sonMin - (sonMinPause - (sonFraTimer))
	
    showextended = 0

    if cint(intEasyreg) <> 1 then
	    call stempelur_kolonne(lto, showextended)
    end if
	


    if instr(request.servervariables("PATH_TRANSLATED"), "to_2015") <> 0 then
		call fLonTimerPer_html_2018(stDato, periode, visning, medid)
	else
		call fLonTimerPer_html(stDato, periode, visning, medid)
	end if
	



end function







'********************************************************************************************************
'**** Fløntimer HTML visning
'********************************************************************************************************

function fLonTimerPer_html(stDato, periode, visning, medid)

         call thisWeekNo53_fn(stDato) 
           thisWeekNo53_stDato = thisWeekNo53

            call thisWeekNo53_fn(now) 
           thisWeekNo53_now = thisWeekNo53


select case visning 
	case 0, 20%>


	<%if visning <> 20 then %>
	
	<h4><span style="font-size:11px; font-weight:lighter;">

         <%if len(trim(meInit)) <> 0 then %>
            <%=meNavn & "  ["& meInit &"]"%>
            <%else %>
            <%=meTxt%>
            <%end if %>

        

	    </span><br />Saldo <%=tsa_txt_340 %>&nbsp;- 
		<%=tsa_txt_005 %>: <%=thisWeekNo53_stDato%></h4>
		
		<!-- Denne uge, nuværende login -->
         <%call erStempelurOn() 
               
        if cint(stempelur_hideloginOn) = 0 then 'skriv ikke login, men åben komme/gå tom

		if cint(thisWeekNo53_stDato) = cint(thisWeekNo53_now) then%>
		<%=tsa_txt_134 %>:
		<%
		
		sLoginTid = "00:00:00"
		
		strSQL = "SELECT l.id AS lid, l.login "_
		&" FROM login_historik l WHERE "_
		&" l.mid = " & medid &" AND stempelurindstilling <> -1"_
		&" ORDER BY l.id DESC" 
					
		'Response.write strSQL
		'Response.flush
		
		
		oRec.open strSQL, oConn, 3 
		if not oRec.EOF then 
        
        sLoginTid = oRec("login") 
        
        end if
        oRec.close
		
		%>
		<b><%=formatdatetime(sLoginTid, 3) %></b>
		
		<% 
		logindiffSidste = datediff("n", sLoginTid, now, 2, 2) 
		
            
        %>
		<br /><%=tsa_txt_340 %>
		<%call timerogminutberegning(logindiffSidste) %>
		<b><%=thoursTot &":"& left(tminTot, 2) %>&nbsp;t.</b>
		
		
		<%end if '*** Denne uge / nuværende login **'
          end if ' if cint(stempelur_hideloginOn) = 1 then 'skriv ikke login, men åben komme/gå tom %>
		
	<table cellspacing=1 cellpadding=4 border=0 width=100% bgcolor="#CCCCCC">
    
    

	<tr bgcolor="#FFFFFF">
		<td style="width:100px;">
            &nbsp;</td>
		
		<td  valign=bottom align=center style="white-space:nowrap; color:#000000; font-size:11px;" ><%=tsa_txt_128 %> d. <%=formatdatetime(stDato, 2) %></td>
		<td  valign=bottom align=center style="white-space:nowrap; color:#000000; font-size:11px;"><%=tsa_txt_129 %> d. <%=formatdatetime(dateadd("d", 1, stDato), 2) %></td>
		<td  valign=bottom align=center style="white-space:nowrap; color:#000000; font-size:11px;"><%=tsa_txt_130 %> d. <%=formatdatetime(dateadd("d", 2, stDato), 2) %></td>
		<td  valign=bottom align=center style="white-space:nowrap; color:#000000; font-size:11px;"><%=tsa_txt_131 %> d. <%=formatdatetime(dateadd("d", 3, stDato), 2) %></td>
		<td  valign=bottom align=center style="white-space:nowrap; color:#000000; font-size:11px;"><%=tsa_txt_132 %> d. <%=formatdatetime(dateadd("d", 4, stDato), 2) %></td>
	    <td  valign=bottom align=center style="white-space:nowrap; color:#000000; font-size:11px; background-color:#F7F7F7;"><%=tsa_txt_133 %> d. <%=formatdatetime(dateadd("d", 5, stDato), 2) %></td>
		<td  valign=bottom align=center style="white-space:nowrap; color:#000000; font-size:11px; background-color:#F7F7F7;"><%=tsa_txt_127 %> d. <%=formatdatetime(dateadd("d", 6, stDato), 2) %></td>
		<td  valign=bottom align=right><b><%=global_txt_167 %></b></td>
	</tr>
	<%
    cspsStur = 1
    else 
    cspsStur = 2%>

    <tr bgcolor="#FFFFFF"><td colspan=10 style="border-bottom:1px #CCCCCC solid;"><br /><br /><b><%=tsa_txt_340 %></b></td></tr>

    <%end if 'visning %>

    <%if cint(showkgtim) = 1 then %>
    <tr bgcolor="#ffffff">
		<td colspan="<%=cspsStur %>"><%=tsa_txt_137 %>:</td>
		<td valign=top align=right><%call timerogminutberegning(manMin)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIalt = ugeIalt + manMin/1%>
		<td valign=top align=right><%call timerogminutberegning(tirMin)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIalt = ugeIalt + tirMin/1%>
		<td valign=top align=right><%call timerogminutberegning(onsMin)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIalt = ugeIalt + onsMin/1%>
		<td valign=top align=right><%call timerogminutberegning(torMin)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIalt = ugeIalt +  torMin/1%>
		<td valign=top align=right><%call timerogminutberegning(freMin)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIalt = ugeIalt +  freMin/1%>
		<td valign=top align=right style="background-color:#F7F7F7;"><%call timerogminutberegning(lorMin)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIalt = ugeIalt +  lorMin/1%>
		<td valign=top align=right style="background-color:#F7F7F7;"><%call timerogminutberegning(sonMin)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIalt = ugeIalt +  sonMin/1%>
		<td valign=top align=right>= 
		<%call timerogminutberegning(ugeIalt)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
	</tr>
    <%
     loginTimerTot = ugeIalt - (ugeIaltPause - ugeIaltFraTilTimer)     
    end if %>


    <%
       



    if cint(showkgpau) = 1 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="<%=cspsStur %>"><%=tsa_txt_138 %>:</td>
		<td valign=top align=right><%call timerogminutberegning(-manMinPause)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIaltPause = ugeIaltPause + manMinPause%>
		<td valign=top align=right><%call timerogminutberegning(-tirMinPause)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIaltPause = ugeIaltPause + tirMinPause%>
		<td valign=top align=right><%call timerogminutberegning(-onsMinPause)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIaltPause = ugeIaltPause + onsMinPause%>
		<td valign=top align=right><%call timerogminutberegning(-torMinPause)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIaltPause = ugeIaltPause + torMinPause%>
		<td valign=top align=right><%call timerogminutberegning(-freMinPause)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIaltPause = ugeIaltPause + freMinPause%>
		<td valign=top align=right style="background-color:#F7F7F7;"><%call timerogminutberegning(-lorMinPause)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIaltPause = ugeIaltPause + lorMinPause%>
		<td valign=top align=right style="background-color:#F7F7F7;"><%call timerogminutberegning(-sonMinPause)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIaltPause = ugeIaltPause + sonMinPause%>
		<td valign=top align=right>= <%call timerogminutberegning(-ugeIaltPause)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
	</tr>
	<%end if %>

	<!-- Fradrag / Tillæg via Realtimer -->
	
    <%if cint(showkgtil) = 1 then %>
    <tr bgcolor="#ffffff">
		<td colspan="<%=cspsStur %>"><%=global_txt_168 %>:*</td>
		<td valign=top align=right><%call timerogminutberegning(manFraTimer)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIaltFraTilTimer = ugeIaltFraTilTimer + (manFraTimer)%>
		<td valign=top align=right><%call timerogminutberegning(tirFraTimer)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIaltFraTilTimer = ugeIaltFraTilTimer + (tirFraTimer)%>
		<td valign=top align=right><%call timerogminutberegning(onsFraTimer)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIaltFraTilTimer = ugeIaltFraTilTimer + (onsFraTimer)%>
		<td valign=top align=right><%call timerogminutberegning(torFraTimer)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIaltFraTilTimer = ugeIaltFraTilTimer + (torFraTimer)%>
		<td valign=top align=right><%call timerogminutberegning(freFraTimer)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIaltFraTilTimer = ugeIaltFraTilTimer + (freFraTimer)%>
		<td valign=top align=right style="background-color:#F7F7F7;"><%call timerogminutberegning(lorFraTimer)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIaltFraTilTimer = ugeIaltFraTilTimer + (lorFraTimer)%>
		<td valign=top align=right style="background-color:#F7F7F7;"><%call timerogminutberegning(sonFraTimer)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIaltFraTilTimer = ugeIaltFraTilTimer + (sonFraTimer)%>
		<td valign=top align=right>= <%call timerogminutberegning(ugeIaltFraTilTimer)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
	</tr>
	<%end if %>

 


    <%if cint(showkgtot) = 1 then %>
	<!-- total -->
	
	
	 <tr bgcolor="#cccccc">
		<td colspan="<%=cspsStur %>"><b><%=global_txt_167%>:</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totMan)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totTir)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totOns)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totTor)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </b></td>
	    <td valign=top align=right><b><%call timerogminutberegning(totFre)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totLor)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totSon)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </b></td>
		<td valign=top align=right>= <b><%call timerogminutberegning(ugeIalt - (ugeIaltPause - (ugeIaltFraTilTimer)))
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </b></td>
		
	</tr>
	<%end if %>
 
	<!-- Normtimer -->
	
    <%if cint(showkgnor) = 1 then %>
    <tr bgcolor="#ffffff">
		<td colspan="<%=cspsStur %>"><%=tsa_txt_259 %>:</td>
		
		<%call normtimerper(medid, varTjDatoUS_man, 6, 0) %>
		
		<td valign=top align=right><%call timerogminutberegning(ntimMan*60)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		
		<td valign=top align=right><%call timerogminutberegning(ntimTir*60)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		
		<td valign=top align=right><%call timerogminutberegning(ntimOns*60)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		
		<td valign=top align=right><%call timerogminutberegning(ntimTor*60)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		
		<td valign=top align=right><%call timerogminutberegning(ntimFre*60)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		
		<td valign=top align=right style="background-color:#F7F7F7;"><%call timerogminutberegning(ntimLor*60)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		
		
		<td valign=top align=right style="background-color:#F7F7F7;"><%call timerogminutberegning(ntimSon*60)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		
		<%
		NormTimerWeekTot = 0
		NormTimerWeekTot = (ntimMan + ntimTir + ntimOns + ntimTor + ntimFre + ntimLor + ntimSon) * 60 %>
		
		<td valign=top align=right>= <%call timerogminutberegning(NormTimerWeekTot)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		</tr>

      <%end if %>
	    
	  <!-- Saldo -->
      
      <%if cint(showkgsal) = 1 then %>  
	  <tr bgcolor="#DCF5BD">
		<td colspan="<%=cspsStur %>" style="height:20px;"><b><%=global_txt_163 %>:</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totMan - (ntimMan*60))
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totTir - (ntimTir*60))
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totOns - (ntimOns*60))
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totTor - (ntimTor*60))
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </b></td>
	    <td valign=top align=right><b><%call timerogminutberegning(totFre - (ntimFre*60))
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totLor - (ntimLor*60))
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totSon - (ntimSon*60))
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </b></td>
		<td valign=top align=right>=
        <%
        call timerogminutberegning((ugeIalt - (ugeIaltPause - (ugeIaltFraTilTimer))) - NormTimerWeekTot)
        %>

        <%
        stDatoUS = year(stDato) & "/" & month(stDato) &"/"& day(stDato)  
        slDatoUS = year(slutDato) & "/" & month(slutDato) &"/"& day(slutDato)
        %>
        <a href="afstem_tot.asp?usemrn=<%=usemrn%>&show=5&varTjDatoUS_man=<%=stDatoUS%>&varTjDatoUS_son=<%=slDatoUS%>" class=vmenu><%=thoursTot &":"& left(tminTot, 2)%></a></td>
		
	</tr>
    <%end if %>


    <%if visning <> 20 then '** Ikke fra timereg. siden da det er for tungt  %>
        <%if cint(showkgsaa) = 1 then '*** IKKE aktiv - se aftamte tot istedet for
        
        
        %>


        <%end if %>
    <%end if %>

  


    <%if visning <> 20 then %>
	</table>
	
    <%if cint(stempelur_hideloginOn) = 0 then 'skriv ikke login, men åben komme/gå tom
	
    if cint(thisWeekNo53_stDato) = cint(thisWeekNo53_now) then%>
	<!-- Denne uge incl. nuværende  -->
	<%=tsa_txt_139 %>: <% 
	call timerogminutberegning(logindiffSidste+(loginTimerTot))
	%>
	<b><%=thoursTot &":"& left(tminTot, 2) %>&nbsp;t.</b>

 
	
    <%end if ' *** Denne uge / Nuværende login **'
    end if    
    %>

	

    
	<br /><br />
	<table><tr><td >
	
	</td></tr></table>

    <%
	if cint(showkguds) = 1 then

	Response.Write "*) <b> Tillægs typer:</b> " & akttypenavnTil
	Response.Write "<br><b>Fradrags typer:</b> " & akttypenavnFra
	
	%>

	
    <%if media <> "print" then %>    
	<br /><br /><a href="#" id="udspec" class="vmenu">+ Udspecificering</a> (fraværs typer)
    <%end if 
    
    
    end if%>
    
    
    <div id="udspecdiv" style="position:relative; width:600px; visibility:hidden; display:none;">
	<table cellspacing=1 cellpadding=2 border=0 width=100% bgcolor="#c4c4c4">    
	    <tr bgcolor="#FFFFFF">
	    <td colspan=9><br /><b>Udspecificering på fraværstyper</b> <br />
	    Ikke medregnet i saldo, med mindre de er en del af <%=global_txt_168 %> typerne*.</td>
	</tr>
	<tr bgcolor="#d6dff5">
		<td style="width:100px;">
            &nbsp;</td>
		
		<td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_128 %> d. <%=formatdatetime(stDato, 2) %></b></td>
		<td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_129 %> d. <%=formatdatetime(dateadd("d", 1, stDato), 2) %></b></td>
		<td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_130 %> d. <%=formatdatetime(dateadd("d", 2, stDato), 2) %></b></td>
		<td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_131 %> d. <%=formatdatetime(dateadd("d", 3, stDato), 2) %></b></td>
		<td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_132 %> d. <%=formatdatetime(dateadd("d", 4, stDato), 2) %></b></td>
	    <td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_133 %> d. <%=formatdatetime(dateadd("d", 5, stDato), 2) %></b></td>
		<td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_127 %> d. <%=formatdatetime(dateadd("d", 6, stDato), 2) %></b></td>
		<td width=50 bgcolor="#ffdfdf" class=lille valign=bottom align=right><b><%=global_txt_167 %></b></td>
	</tr>
	
	    
	    <%if cint(aty_fleks_on) = 1 then %>
	<!-- Fleks Realtimer -->
	<tr bgcolor="#ffffff">
		<td><%=global_txt_147 &" "& aty_fleks_tf%></td>
		<td valign=top align=right><%call timerogminutberegning(manFlekstimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFlekstimer = ugeIaltFlekstimer + (manFlekstimer)%>
		<td valign=top align=right><%call timerogminutberegning(tirFlekstimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFlekstimer = ugeIaltFlekstimer + (tirFlekstimer)%>
		<td valign=top align=right><%call timerogminutberegning(onsFlekstimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFlekstimer = ugeIaltFlekstimer + (onsFlekstimer)%>
		<td valign=top align=right><%call timerogminutberegning(torFlekstimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFlekstimer = ugeIaltFlekstimer + (torFlekstimer)%>
		<td valign=top align=right><%call timerogminutberegning(freFlekstimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFlekstimer = ugeIaltFlekstimer + (freFlekstimer)%>
		<td valign=top align=right><%call timerogminutberegning(lorFlekstimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFlekstimer = ugeIaltFlekstimer + (lorFlekstimer)%>
		<td valign=top align=right><%call timerogminutberegning(sonFlekstimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFlekstimer = ugeIaltFlekstimer + (sonFlekstimer)%>
		<td valign=top align=right>= <%call timerogminutberegning(ugeIaltFlekstimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	</tr>
	
	<%end if %>
	
	
	
	<%if cint(aty_Ferie_on) = 1 then %>
	<!-- Ferie Realtimer -->
	<tr bgcolor="#ffffff">
		<td><%=global_txt_135 &" "& aty_Ferie_tf %></td>
		<td valign=top align=right><%call timerogminutberegning(manFerietimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFerietimer = ugeIaltFerietimer + (manFerietimer)%>
		<td valign=top align=right><%call timerogminutberegning(tirFerietimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFerietimer = ugeIaltFerietimer + (tirFerietimer)%>
		<td valign=top align=right><%call timerogminutberegning(onsFerietimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFerietimer = ugeIaltFerietimer + (onsFerietimer)%>
		<td valign=top align=right><%call timerogminutberegning(torFerietimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFerietimer = ugeIaltFerietimer + (torFerietimer)%>
		<td valign=top align=right><%call timerogminutberegning(freFerietimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFerietimer = ugeIaltFerietimer + (freFerietimer)%>
		<td valign=top align=right><%call timerogminutberegning(lorFerietimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFerietimer = ugeIaltFerietimer + (lorFerietimer)%>
		<td valign=top align=right><%call timerogminutberegning(sonFerietimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFerietimer = ugeIaltFerietimer + (sonFerietimer)%>
		<td valign=top align=right>= <%call timerogminutberegning(ugeIaltFerietimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	</tr>
	
	<%end if %>
	
	<%if cint(aty_Syg_on) = 1 then %>
	<!-- Syg Realtimer -->
	<tr bgcolor="#ffffff">
		<td><%=global_txt_138 &" "& aty_Syg_tf %></td>
		<td valign=top align=right><%call timerogminutberegning(manSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSygtimer = ugeIaltSygtimer + (manSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(tirSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSygtimer = ugeIaltSygtimer + (tirSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(onsSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSygtimer = ugeIaltSygtimer + (onsSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(torSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSygtimer = ugeIaltSygtimer + (torSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(freSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSygtimer = ugeIaltSygtimer + (freSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(lorSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSygtimer = ugeIaltSygtimer + (lorSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(sonSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSygtimer = ugeIaltSygtimer + (sonSygtimer)%>
		<td valign=top align=right>= <%call timerogminutberegning(ugeIaltSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	</tr>
	
	<%end if %>
	
	
	<%if cint(aty_BarnSyg_on) = 1 then %>
	<!-- BarnSyg Realtimer -->
	<tr bgcolor="#ffffff">
		<td><%=global_txt_139 &" "& aty_BarnSyg_tf %></td>
		<td valign=top align=right><%call timerogminutberegning(manBarnSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltBarnSygtimer = ugeIaltBarnSygtimer + (manBarnSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(tirBarnSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltBarnSygtimer = ugeIaltBarnSygtimer + (tirBarnSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(onsBarnSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltBarnSygtimer = ugeIaltBarnSygtimer + (onsBarnSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(torBarnSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltBarnSygtimer = ugeIaltBarnSygtimer + (torBarnSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(freBarnSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltBarnSygtimer = ugeIaltBarnSygtimer + (freBarnSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(lorBarnSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltBarnSygtimer = ugeIaltBarnSygtimer + (lorBarnSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(sonBarnSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltBarnSygtimer = ugeIaltBarnSygtimer + (sonBarnSygtimer)%>
		<td valign=top align=right>= <%call timerogminutberegning(ugeIaltBarnSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	</tr>
	
	<%end if %>
	
	
	<%if cint(aty_Lage_on) = 1 then %>
	<!-- Lage Realtimer -->
	<tr bgcolor="#ffffff">
		<td><%=global_txt_160 &" "& aty_Lage_tf %></td>
		<td valign=top align=right><%call timerogminutberegning(manLagetimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltLagetimer = ugeIaltLagetimer + (manLagetimer)%>
		<td valign=top align=right><%call timerogminutberegning(tirLagetimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltLagetimer = ugeIaltLagetimer + (tirLagetimer)%>
		<td valign=top align=right><%call timerogminutberegning(onsLagetimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltLagetimer = ugeIaltLagetimer + (onsLagetimer)%>
		<td valign=top align=right><%call timerogminutberegning(torLagetimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltLagetimer = ugeIaltLagetimer + (torLagetimer)%>
		<td valign=top align=right><%call timerogminutberegning(freLagetimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltLagetimer = ugeIaltLagetimer + (freLagetimer)%>
		<td valign=top align=right><%call timerogminutberegning(lorLagetimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltLagetimer = ugeIaltLagetimer + (lorLagetimer)%>
		<td valign=top align=right><%call timerogminutberegning(sonLagetimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltLagetimer = ugeIaltLagetimer + (sonLagetimer)%>
		<td valign=top align=right>= <%call timerogminutberegning(ugeIaltLagetimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	</tr>
	
	<%end if %>
	
		<%if cint(aty_Sund_on) = 1 then %>
	<!-- Sund Realtimer -->
	<tr bgcolor="#ffffff">
		<td><%=global_txt_148 &" "& aty_Sund_tf %></td>
		<td valign=top align=right><%call timerogminutberegning(manSundtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSundtimer = ugeIaltSundtimer + (manSundtimer)%>
		<td valign=top align=right><%call timerogminutberegning(tirSundtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSundtimer = ugeIaltSundtimer + (tirSundtimer)%>
		<td valign=top align=right><%call timerogminutberegning(onsSundtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSundtimer = ugeIaltSundtimer + (onsSundtimer)%>
		<td valign=top align=right><%call timerogminutberegning(torSundtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSundtimer = ugeIaltSundtimer + (torSundtimer)%>
		<td valign=top align=right><%call timerogminutberegning(freSundtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSundtimer = ugeIaltSundtimer + (freSundtimer)%>
		<td valign=top align=right><%call timerogminutberegning(lorSundtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSundtimer = ugeIaltSundtimer + (lorSundtimer)%>
		<td valign=top align=right><%call timerogminutberegning(sonSundtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSundtimer = ugeIaltSundtimer + (sonSundtimer)%>
		<td valign=top align=right>= <%call timerogminutberegning(ugeIaltSundtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	</tr>
	
	<%end if %>    
	
	<%if cint(aty_Frokost_on) = 1 then %>
	<!-- Frokost Realtimer -->
	<tr bgcolor="#ffffff">
		<td><%=global_txt_133 &" "& aty_Frokost_tf %></td>
		<td valign=top align=right><%call timerogminutberegning(manFrokosttimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFrokosttimer = ugeIaltFrokosttimer + (manFrokosttimer)%>
		<td valign=top align=right><%call timerogminutberegning(tirFrokosttimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFrokosttimer = ugeIaltFrokosttimer + (tirFrokosttimer)%>
		<td valign=top align=right><%call timerogminutberegning(onsFrokosttimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFrokosttimer = ugeIaltFrokosttimer + (onsFrokosttimer)%>
		<td valign=top align=right><%call timerogminutberegning(torFrokosttimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFrokosttimer = ugeIaltFrokosttimer + (torFrokosttimer)%>
		<td valign=top align=right><%call timerogminutberegning(freFrokosttimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFrokosttimer = ugeIaltFrokosttimer + (freFrokosttimer)%>
		<td valign=top align=right><%call timerogminutberegning(lorFrokosttimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFrokosttimer = ugeIaltFrokosttimer + (lorFrokosttimer)%>
		<td valign=top align=right><%call timerogminutberegning(sonFrokosttimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFrokosttimer = ugeIaltFrokosttimer + (sonFrokosttimer)%>
		<td valign=top align=right>= <%call timerogminutberegning(ugeIaltFrokosttimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	</tr>
	
	<%end if %>    
	    
	    
	</table>

   
	</div>
	<%

    end if 'vis <> 3
	
	case 2
	
	totalTimerPer100 = 0
	totalTimerPer100 = (totaltimerPer-totalpausePer)
	
    call timerogminutberegning(totalTimerPer100) %>
	<%=thoursTot &":"& left(tminTot, 2) %>
	
	<%
	case 3
	
	totalTimerPer100 = 0
	totalTimerPer100 = (totaltimerPer-totalpausePer)
	
	case 4 ''** Avg. Week ***' 
	
	if weekDiff <> 0 then
	weekDiff = weekDiff
	else
	weekDiff = 1
	end if
	
	totalTimerPer100 = 0
	totalTimerPer100 = (totaltimerPer-totalpausePer)/weekDiff
	
	
	call timerogminutberegning(((totaltimerPer-totalpausePer)/weekDiff)) %>
	<!-- formatnumber(totalTimerPer100, 2) -->
	
	<%=thoursTot &":"& left(tminTot, 2) %>
	<%

    case 21 'smiley på timereg.
	

     totalTimerPer100 = 0
	totalTimerPer100 = (totaltimerPer-totalpausePer)
	
    %>
	
     
    <% 
    case else
	
	
	totalTimerPer100 = 0
	totalTimerPer100 = (totaltimerPer-totalpausePer)
	
	
	call timerogminutberegning(totaltimerPer-totalpausePer) %>
	<!-- formatnumber(totalTimerPer100, 2) -->
	
	<%=thoursTot &":"& left(tminTot, 2) %> 
	
	<%
	end select 
	'** Visinng ***




end function

'******* END '***







'********************************************************************************************************
'*** Tilføj Pauser
'********************************************************************************************************

function tilfojPauser(io,psMid,psDato,psloginTidp,pslogudTidp,psMin,psKomm)

                            
                            
                            'Response.Write "psloginTidp: " & psloginTidp & " psDato: "& psDato &"<br>"
                            'response.write "<br>psMin: "& psMin
    
                            
                            if len(trim(pslogudTidp)) = 0 then
                            pslogudTidp = psloginTidp
                            end if



                            '*** Indlæser ikke pauser på 0 min.
                            if psMin <> 0 then
                            
                            psDato = year(psDato) &"/"& month(psDato) &"/"& day(psDato)

                            strSQLp = "INSERT INTO login_historik SET dato = '"& psDato &"', "_
	                        &" login = '"& psloginTidp &"', "_
	                        &" logud = '"& pslogudTidp &"', "_
					        &" stempelurindstilling = -1, minutter = "& psMin &", "_
					        &" manuelt_afsluttet = 0, kommentar = '"& psKomm &"', mid = "& psMid
					        
                            'if session("mid") = 1 then
					        'Response.Write "pause: "& strSQLp & "<br>"
					        'Response.flush
                            'end if                        

                            oConn.execute(strSQLp)


                            end if

                           

end function





public stPauseLic_1, stPauseLic_2, p1on, p2on, p1_grp, p2_grp, stPauseLic_3, stPauseLic_4, p3on, p4on, p3_grp, p4_grp
function stPauserFralicens(psDt)

            
            'Response.write "##"& psDt
            'Response.end

            psDt = day(psDt)&"/"&month(psDt)&"/"&year(psDt)
            dagidag = datepart("w", psDt, 2,2) 'weekday(psDt, 2)
            p1on = 0
            p2on = 0
            p3on = 0
            p4on = 0


            'response.write "dagidag: "& dagidag
            feltNavne = ""
            select case dagidag
            case 1
            feltNavne = "p1_man AS p1on, p2_man AS p2on"
            feltNavne = feltNavne &", p3_man AS p3on, p4_man AS p4on"
            case 2
            feltNavne = "p1_tir AS p1on, p2_tir AS p2on"
            feltNavne = feltNavne &", p3_tir AS p3on, p4_tir AS p4on"
            case 3
            feltNavne = "p1_ons AS p1on, p2_ons AS p2on"
            feltNavne = feltNavne &", p3_ons AS p3on, p4_ons AS p4on"
            case 4
            feltNavne = "p1_tor AS p1on, p2_tor AS p2on"
            feltNavne = feltNavne &", p3_tor AS p3on, p4_tor AS p4on"
            case 5
            feltNavne = "p1_fre AS p1on, p2_fre AS p2on"
            feltNavne = feltNavne &", p3_fre AS p3on, p4_fre AS p4on"
            case 6
            feltNavne = "p1_lor AS p1on, p2_lor AS p2on"
            feltNavne = feltNavne &", p3_lor AS p3on, p4_lor AS p4on"
            case 7
            feltNavne = "p1_son AS p1on, p2_son AS p2on"
            feltNavne = feltNavne &", p3_son AS p3on, p4_son AS p4on"
            end select
        
            '*** Henter forvalgte standard pauser ***

            p1_grp = 0 
            p2_grp = 0
            p3_grp = 0
            p4_grp = 0

            strSQLstp = "SELECT stpause, stpause2, stpause3, stpause4, "& feltNavne &", p1_grp, p2_grp, p3_grp, p4_grp"_
            &" FROM  "_
            &" licens WHERE id = 1" 
            

            'Response.write strSQLstp
            'Response.end
            oRec5.open strSQLstp, oConn, 3
            if not oRec5.EOF then
            
            stPauseLic_1 = oRec5("stpause")
            stPauseLic_2 = oRec5("stpause2")
            stPauseLic_3 = oRec5("stpause3")
            stPauseLic_4 = oRec5("stpause4")        

            p1on = oRec5("p1on")
            p2on = oRec5("p2on")
            p3on = oRec5("p3on")
            p4on = oRec5("p4on")

            p1_grp = oRec5("p1_grp")
            p2_grp = oRec5("p2_grp")
            p3_grp = oRec5("p3_grp")
            p4_grp = oRec5("p4_grp")
             
            end if
            oRec5.close
        


end function


public p1_prg_on, p2_prg_on, p3_prg_on, p4_prg_on
function stPauserProgrp(useMid, p1_grp, p2_grp, p3_grp, p4_grp)

'p1_grp Hvilke projektgrupper tilføejt pause 1, komma sep
'p2_grp Hvilke projektgrupper tilføejt pause 2, komma sep




        p1_prg_on = 0
        p2_prg_on = 0
        p3_prg_on = 0
        p4_prg_on = 0
        negativPauseGruppeFundet = 0
        tilfojikkepauserpagrp = 0

        strSQLmansat = " mansat <> 0 AND mansat <> 4 "


        '*** Adgang til Pause 1 ****'
        if len(trim(p1_grp)) = 0 OR isNull(p1_grp) = true then 'ingen værdi angivet i kontrolpanel
        p1_prg_on = 1
        else
            
            
            p1_grpArr = split(replace(p1_grp, " ", ""), ",") 
            for g = 0 to UBOUND(p1_grpArr)
                
                'Response.Write "DER: p1_grpArr(g)"& p1_grpArr(g) & " useMid: "&  useMid&"<br>"
                'Response.flush

                'negativPauseGruppeFundet = 0

                if instr(p1_grpArr(g), "-") <> 0 then 'Hvis evt minus betyder at denne gruppe er udelukket fra pauser, uanset om medarbejder er med i andre projektgrupper
                negativPauseGruppeFundet = 1
                p1_grpArr(g) = replace(p1_grpArr(g), "-", "")
                end if

                if len(trim(p1_grpArr(g))) <> 0 then
                call erdetint_st(trim(p1_grpArr(g)))
                if isInt_st = 0 then
                        
                       
                        call medarbiprojgrp(p1_grpArr(g), useMid, 0, -1)
                        
                        if cint(negativPauseGruppeFundet) = 1 then                    

                        if (instr(instrMedidProgrpThisGrp, "#"& trim(useMid) &"#,")) <> 0 then 'medarbejder er med i negativgruppe
                        tilfojikkepauserpagrp = 1
                        end if

                        end if

                end if
                isInt_st = 0
                end if


               
                
            next

                if cint(negativPauseGruppeFundet) = 1 then 
    
                if cint(tilfojikkepauserpagrp) = 0 then
                p1_prg_on = 1
                else
                p1_prg_on = 0
                end if   
    
                else   

                if instr(instrMedidProgrp, "#"& trim(useMid) &"#,") <> 0 then
                p1_prg_on = 1
                else
                p1_prg_on = 0
                end if

                end if

        end if

        'Response.write "HER: p1_grp:"& p1_grp &" p1_prg_on " & p1_prg_on & "  string: "& instrMedidProgrp & " instr " & instr(instrMedidProgrp, "#"& trim(useMid) &"#,") &" tilfojikkepauserpagrp: "& tilfojikkepauserpagrp ' & " # mid: " & useMid & " |<br>"

       
         '*** Adgang til Pause 2 ****'
        negativPauseGruppeFundet = 0
        tilfojikkepauserpagrp = 0


        if len(trim(p2_grp)) = 0 OR isNull(p2_grp) = true then
        p2_prg_on = 1
        else
            
            p2_grpArr = split(replace(p2_grp, " ", ""), ",")
            for g = 0 to UBOUND(p2_grpArr)
                
                'Response.Write "p1_grpArr(g)"& p1_grpArr(g) & " useMid: "&  useMid&"<br>"

                if instr(p2_grpArr(g), "-") <> 0 then 'Hvis evt minus betyder at denne gruppe er udelukket fra pauser, uanset om medarbejder er med i andre porjektgrupper
                negativPauseGruppeFundet = 1
                p2_grpArr(g) = replace(p2_grpArr(g), "-", "")
                end if

                    if len(trim(p2_grpArr(g))) <> 0 then
                    call erdetint_st(trim(p2_grpArr(g)))
                            if isInt_st = 0 then
                                        
    
                                            call medarbiprojgrp(p2_grpArr(g), useMid, 0, -1)

                                    
                                            if cint(negativPauseGruppeFundet) = 1 then                    

                                            if (instr(instrMedidProgrpThisGrp, "#"& trim(useMid) &"#,")) <> 0 then 'medarbejder er med i negativgruppe
                                            tilfojikkepauserpagrp = 1
                                            end if

                                            end if


                            end if
                    isInt_st = 0
                    end if

            

            next

                
    
    
                if cint(negativPauseGruppeFundet) = 1 then 
    
                if cint(tilfojikkepauserpagrp) = 0 then
                p2_prg_on = 1
                else
                p2_prg_on = 0
                end if   
    
                else   

                if instr(instrMedidProgrp, "#"& trim(useMid) &"#,") <> 0 then
                p2_prg_on = 1
                else
                p2_prg_on = 0
                end if

                end if

        end if


        'Response.write "<br><br>HER: p2_grp:"& p2_grp &" p2_prg_on " & p2_prg_on & "  string: "& instrMedidProgrp & " instr " & instr(instrMedidProgrp, "#"& trim(useMid) &"#,") & " # mid: " & useMid & " |<br>"
        'Response.end


         '*** Adgang til Pause 3 ****'
        negativPauseGruppeFundet = 0
        tilfojikkepauserpagrp = 0


        if len(trim(p3_grp)) = 0 OR isNull(p3_grp) = true then
        p3_prg_on = 1
        else
            
            p3_grpArr = split(replace(p3_grp, " ", ""), ",")
            for g = 0 to UBOUND(p3_grpArr)
                
                'Response.Write "p1_grpArr(g)"& p1_grpArr(g) & " useMid: "&  useMid&"<br>"

                if instr(p3_grpArr(g), "-") <> 0 then 'Hvis evt minus betyder at denne gruppe er udelukket fra pauser, uanset om medarbejder er med i andre porjektgrupper
                negativPauseGruppeFundet = 1
                p3_grpArr(g) = replace(p3_grpArr(g), "-", "")
                end if

                    if len(trim(p3_grpArr(g))) <> 0 then
                    call erdetint_st(trim(p3_grpArr(g)))
                            if isInt_st = 0 then
                                        
    
                                            call medarbiprojgrp(p3_grpArr(g), useMid, 0, -1)

                                    
                                            if cint(negativPauseGruppeFundet) = 1 then                    

                                            if (instr(instrMedidProgrpThisGrp, "#"& trim(useMid) &"#,")) <> 0 then 'medarbejder er med i negativgruppe
                                            tilfojikkepauserpagrp = 1
                                            end if

                                            end if


                            end if
                    isInt_st = 0
                    end if

            

            next

                
    
    
                if cint(negativPauseGruppeFundet) = 1 then 
    
                if cint(tilfojikkepauserpagrp) = 0 then
                p3_prg_on = 1
                else
                p3_prg_on = 0
                end if   
    
                else   

                if instr(instrMedidProgrp, "#"& trim(useMid) &"#,") <> 0 then
                p3_prg_on = 1
                else
                p3_prg_on = 0
                end if

                end if

        end if


        'Response.write "<br><br>HER: p3_grp:"& p3_grp &" p3_prg_on " & p3_prg_on & "  string: "& instrMedidProgrp & " instr " & instr(instrMedidProgrp, "#"& trim(useMid) &"#,") & " # mid: " & useMid & " |<br>"
        'Response.end


         '*** Adgang til Pause 4 ****'
        negativPauseGruppeFundet = 0
        tilfojikkepauserpagrp = 0


        if len(trim(p4_grp)) = 0 OR isNull(p4_grp) = true then
        p4_prg_on = 1
        else
            
            p4_grpArr = split(replace(p4_grp, " ", ""), ",")
            for g = 0 to UBOUND(p4_grpArr)
                
                'Response.Write "p1_grpArr(g)"& p1_grpArr(g) & " useMid: "&  useMid&"<br>"

                if instr(p4_grpArr(g), "-") <> 0 then 'Hvis evt minus betyder at denne gruppe er udelukket fra pauser, uanset om medarbejder er med i andre porjektgrupper
                negativPauseGruppeFundet = 1
                p4_grpArr(g) = replace(p4_grpArr(g), "-", "")
                end if

                    if len(trim(p4_grpArr(g))) <> 0 then
                    call erdetint_st(trim(p4_grpArr(g)))
                            if isInt_st = 0 then
                                        
    
                                            call medarbiprojgrp(p4_grpArr(g), useMid, 0, -1)

                                    
                                            if cint(negativPauseGruppeFundet) = 1 then                    

                                            if (instr(instrMedidProgrpThisGrp, "#"& trim(useMid) &"#,")) <> 0 then 'medarbejder er med i negativgruppe
                                            tilfojikkepauserpagrp = 1
                                            end if

                                            end if


                            end if
                    isInt_st = 0
                    end if

            

            next

                
    
    
                if cint(negativPauseGruppeFundet) = 1 then 
    
                if cint(tilfojikkepauserpagrp) = 0 then
                p4_prg_on = 1
                else
                p4_prg_on = 0
                end if   
    
                else   

                if instr(instrMedidProgrp, "#"& trim(useMid) &"#,") <> 0 then
                p4_prg_on = 1
                else
                p4_prg_on = 0
                end if

                end if

        end if


        'Response.write "<br><br>HER: p4_grp:"& p4_grp &" p4_prg_on " & p4_prg_on & "  string: "& instrMedidProgrp & " instr " & instr(instrMedidProgrp, "#"& trim(useMid) &"#,") & " # mid: " & useMid & " |<br>"
        'Response.end



end function


Public isInt_st
function erDetInt_st(FM_felt) 
isInt_st = instr(lcase(FM_felt), "a")
isInt_st = isInt_st + (instr(lcase(FM_felt), "b"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "c"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "d"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "e"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "f"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "g"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "h"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "i"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "j"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "k"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "l"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "m"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "n"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "o"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "p"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "q"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "r"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "s"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "t"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "u"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "v"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "w"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "x"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "y"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "z"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "æ"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "ø"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "å"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "<"))
isInt_st = isInt_st + (instr(lcase(FM_felt), ">"))
'isInt_st = isInt_st + (instr(lcase(FM_felt), ","))
isInt_st = isInt_st + (instr(lcase(FM_felt), "?"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "!"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "#"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "%"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "&"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "/"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "("))
isInt_st = isInt_st + (instr(lcase(FM_felt), ")"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "["))
isInt_st = isInt_st + (instr(lcase(FM_felt), "]"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "="))
isInt_st = isInt_st + (instr(lcase(FM_felt), ";"))
isInt_st = isInt_st + (instr(lcase(FM_felt), ":"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "¤"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "@"))
isInt_st = isInt_st + (instr(lcase(FM_felt), "'"))
isInt_st = isInt_st + (instr(lcase(FM_felt), " "))
end function








public lonTimerGT, stempelUrEkspTxtShowTot   
function tottimer(lastMnavn, lastMnr, totalhours, totalmin, lastMid, sqlDatoStart, sqlDatoslut, vis, lastDato)
		

        'if session("mid") = 1 then
        'Response.write totalhours &"+"& totalmin & "<br>"
        'end if

        'vis 1: total pr. uge
        'vis 2: GT i bunden af hver medarbejder

		totalmin = (60 * totalhours) + totalmin
		lonTimer = totalmin
		call timerogminutberegning(totalmin)

        if vis = 2 OR (cint(showTot) = 1 AND vis = 1) then
        lonTimerGT = lonTimerGT + lonTimer
	    end if	

		if cint(showTot) = 1 then
            
                select case lto
                case "cst"
	            csp = 5
                case else
                csp = 6
                end select

		
		else
		csp = 9
		end if

        ugeNummer = day(lastDato) & "/" & month(lastDato) & "/" & year(lastDato)
        'ugeNummer = 

        if cint(vis) = 1 then 'total i bunden / else uge sum
         trBg = "#FFFFFF"
        else
        trBg = "lightpink"
        end if
        
      

     if cint(showTot) = 1 AND vis = 1 then

        if media <> "export" then
    %>
        <tr>
        <td>&nbsp;</td>
        <td><%=lastMnavn & " ["& lastMinit &"]" %></td>
        <td><%=formatdatetime(sqlDatoStart, 1) &" - "& formatdatetime(sqlDatoSlut, 1) %></td>
        <td align=right><%=thoursTot%>:<%=left(tminTot, 2)%></td>

      

    <%
        else

        stempelUrEkspTxtShowTot = stempelUrEkspTxtShowTot & ""& meNavn &";"& meInit & ";" & sqlDatoStart & ";" & sqlDatoSlut & ";" 
        stempelUrEkspTxtShowTot = stempelUrEkspTxtShowTot & ""& thoursTot&"."& left(tminTot, 2) &";"

        end if
      
        
                

    else 
        
        if vis <> 1 then%>

               <%call thisWeekNo53_fn(ugeNummer)%>

        <tr bgcolor="<%=trBg%>">
		        <td>&nbsp;</td>
                <td colspan=<%=csp-4%> style="padding:2px 2px 2px 2px; height:30px;">Løntimer (komme/gå) uge: <%=thisWeekNo53%>:</td>
                <td align=right><b><%=thoursTot%>:<%=left(tminTot, 2)%></b></td>

		
		        <td>&nbsp;</td>
		        <td>&nbsp;</td>
		        <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
	     </tr>
		<%end if
		end if
    
        if vis = 1 then
        
        trBg = "#FFFFFF"

        if cint(showTot) <> 1 then
        %>
        <tr>
		    <td colspan="<%=csp+3%>" style="height:10px;">&nbsp;</td>
	    </tr>
        <%end if %>


      <%if cint(showTot) <> 1 then%>
	 <tr bgcolor="<%=trbg %>">
		<td>&nbsp;</td>
		<td colspan=<%=csp-4%> style="padding:2px;">Fradrag til løntimer:</td>
        <%end if 
            
            
         if media <> "export" then%>
        <td align=right>
		<%
        end if

		call akttyper2009(2)
		
		tiltimer = 0
        
            stDtKriTL = year(sqlDatoStart)&"/"&month(sqlDatoStart)&"/"&day(sqlDatoStart)
            slDtKriTL = year(sqlDatoSlut)&"/"&month(sqlDatoSlut)&"/"&day(sqlDatoSlut)

		strSQL2 = "SELECT sum(t.timer) AS tiltimer FROM timer t WHERE "_
		&" (tmnr = "& lastMid &" AND tdato BETWEEN '"& stDtKriTL &"' AND '"& slDtKriTL &"' AND ("& aty_sql_tilwhours &")) GROUP BY tmnr " 
		
            'if session("mid") = 1 then
            'Response.Write strSQL2
		'Response.flush
		'end if

		oRec2.open strSQL2, oConn, 3 
		if not oRec2.EOF then
			tiltimer = oRec2("tiltimer")
		end if
		oRec2.close 
		
		if len(trim(tiltimer)) <> 0 then
		tiltimer = tiltimer
		else
		tiltimer = 0
		end if
		
		fradtimer = 0
		strSQL2 = "SELECT sum(timer) AS fratimer FROM timer t WHERE "_
		&" (tmnr = "& lastMid &" AND tdato BETWEEN '"& stDtKriTL &"' AND '"& slDtKriTL &"' AND ("& aty_sql_frawhours &")) GROUP BY tmnr "
		'Response.Write strSQL2
		'Response.flush
		
		oRec2.open strSQL2, oConn, 3 
		if not oRec2.EOF then
			fradtimer = oRec2("fratimer")
		end if
		oRec2.close 
		
		if len(trim(fradtimer)) <> 0 then
		fradtimer = fradtimer
		else
		fradtimer = 0
		end if
		
		'Response.Write "fradtimer" & fradtimer & "# tilt#"& tiltimer
		
		totalmin = (60 * (-(fradtimer) + (tiltimer)))
		fradragTil = totalmin
		call timerogminutberegning(totalmin)
		
        if media <> "export" then
		%>
		&nbsp;&nbsp;<%=thoursTot%>:<%=left(tminTot, 2)%>
	    </td>
         <%end if %>

		<%if showTot <> 1 then%>
		    <td>&nbsp;</td>
		    <td>&nbsp;</td>
		    <td>&nbsp;</td>
            <td>&nbsp;</td>
             <td>&nbsp;</td>
	     </tr>

        <%else
            
            if media = "export" then
            stempelUrEkspTxtShowTot = stempelUrEkspTxtShowTot & ""& thoursTot&"."& left(tminTot, 2) &";"
            end if
     
		end if %>
		
	 

     <%if showTot <> 1 then%>

	  <tr bgcolor="<%=trBg %>">
		<td style="padding:2px; border-bottom:1px #999999 solid;">&nbsp;</td>
		<td colspan=<%=csp-4%> style="padding:2px; border-bottom:1px #999999 solid;"><b>Grandtotal Løntimer (komme/gå - fradrag):</b></td>
     <%end if %>

         <%if cint(showTot) <> 1 then%>
		<td style="border-bottom:1px #999999 solid;" align=right>
        <%else 
            
            if media <> "export" then%>
        <td align=right>
        <%end if
            end if %>

	    <%'*** Løn Timer minus fradarg *** %>
		<%lonTimerBeregnet = lonTimerGT + (fradragTil)
		call timerogminutberegning(lonTimerBeregnet) %>

        <%if media <> "export" then %>
		&nbsp;&nbsp;<b><%=thoursTot%>:<%=left(tminTot, 2)%></b></td>
        <%end if

             lonTimerBeregnet = 0
             lonTimerGT = 0 
             fradragTil = 0 
        %>
		

          <%if showTot <> 1 then%>
        	<td style="padding:2px; border-bottom:1px #999999 solid;">&nbsp;</td>
		
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
         </tr>
        <%else 

            if media = "export" then
            stempelUrEkspTxtShowTot = stempelUrEkspTxtShowTot & ""& thoursTot&"."& left(tminTot, 2) &";"
            end if

          
		end if %>
	
	
	 
	 <%if lto <> "cst" AND lto <> "cflow" then %>
	 
      <%if showTot <> 1 then%>
	  <tr bgcolor="<%=trBg %>">
		<td>&nbsp;</td>
		<td colspan=<%=csp-4%> style="padding:2px;">Realiseret timer (projekt):</td>
        <%end if 
            
            
        if media <> "export" then %>
        <td align=right>
		<%end if
		
		
		regtimer = 0
		strSQL2 = "SELECT sum(timer) AS sumtimer FROM timer WHERE tmnr = "& lastMid &" AND tdato BETWEEN '"& stDtKriTL &"' AND '"& slDtKriTL &"' AND ("& aty_sql_realhours &") "
		'Response.Write strSQL2
		'Response.flush
		
		oRec2.open strSQL2, oConn, 3 
		if not oRec2.EOF then
			regtimer = oRec2("sumtimer")
		end if
		oRec2.close 
		
		totalmin = (60 * regtimer)
		call timerogminutberegning(totalmin)
		
		if media <> "export" then %>
		&nbsp;&nbsp;<%=thoursTot%>:<%=left(tminTot, 2)%>
		</td>
		<%end if
            
        if showTot <> 1 then%>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
	    </tr>
		<%
        else
            
             if media = "export" then
            stempelUrEkspTxtShowTot = stempelUrEkspTxtShowTot & ""& thoursTot&"."& left(tminTot, 2) &";"
            end if
          

        end if %>
		
	 <%end if %>

     <%
            
        '**** Overtid / Overarbejde *************    
        select case lto 
        case "cflow", "intranet - local"
        %>

         <%if showTot <> 1 then%>
         <tr bgcolor="<%=trBg %>">
		<td>&nbsp;</td>
		<td colspan=<%=csp-4%> style="padding:2px;">Overtid:</td>
          <%
         end if
         
              
        if media <> "export" then %>
        <td align=right>
        <%end if

        overtid = 0
		strSQL2 = "SELECT sum(timer) AS overtid FROM timer t WHERE "_
		&" (tmnr = "& lastMid &" AND tdato BETWEEN '"& stDtKriTL &"' AND '"& slDtKriTL &"' AND (tfaktim = 30 AND godkendtstatus = 1)) GROUP BY tmnr "
		
        'if session("mid") = 1 then     
        'Response.Write strSQL2
		'Response.flush
        'end if
		
		oRec2.open strSQL2, oConn, 3 
		if not oRec2.EOF then
			overtid = oRec2("overtid")
		end if
		oRec2.close 
		
		if len(trim(overtid)) <> 0 then
		overtid = overtid
		else
		overtid = 0
		end if
		
		'Response.Write "fradtimer" & fradtimer & "# tilt#"& tiltimer
		
		totalmin = (60 * overtid)
		'fradragTil = totalmin
		call timerogminutberegning(totalmin)
		
		if media <> "export" then %>
		&nbsp;&nbsp;<%=thoursTot%>:<%=left(tminTot, 2)%>
        </td>
        <%end if
            
            
        if showTot <> 1 then%>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
        <td>&nbsp;</td>
       <td>&nbsp;</td>
       </tr>

        <%else 
        
             if media = "export" then
            stempelUrEkspTxtShowTot = stempelUrEkspTxtShowTot & ""& thoursTot&"."& left(tminTot, 2) &";"
            end if
           
            
        end if %>
	
	 


        
            
        <%end select%>


       <%if showTot = 1 then
           
            if media = "export" then
           stempelUrEkspTxtShowTot = stempelUrEkspTxtShowTot & "xx99123sy#z"
           end if

        %>
        </tr>
        <%else %>
        <tr><td colspan="10">&nbsp;<br /><br /><br /><br />&nbsp;</td></tr>
       <%end if %>

        


	 
     <%end if 'VIS: uge 2 ell. tot 1 %>

	<%
    Response.flush
	end function







    public stempelUrEkspTxto, datoOvertid, minutterExpTxtOtid
    function fn_overtid(lto, datoOvertid, medid, showifnul)


        select case lto
        case "cflow", "intranet - local"
        
        datoSQLo = year(datoOvertid) &"/"& month(datoOvertid) &"/"& day(datoOvertid) 
        overtid = 0
		strSQL2 = "SELECT sum(timer) AS overtid FROM timer t WHERE "_
		&" (tmnr = "& medid &" AND tdato = '"& datoSQLo &"' AND (tfaktim = 30 AND godkendtstatus = 1)) GROUP BY tmnr "
		
        'if session("mid") = 1 then
        'Response.Write strSQL2
		'Response.flush
        'end if
		
		oRec2.open strSQL2, oConn, 3 
		if not oRec2.EOF then
			overtid = oRec2("overtid")
		end if
		oRec2.close 
		
		if len(trim(overtid)) <> 0 then
		overtid = overtid
		else
		overtid = 0
		end if
		
		'Response.Write "fradtimer" & fradtimer & "# tilt#"& tiltimer
		
		totalmin = (60 * overtid)
		'fradragTil = totalmin
		call timerogminutberegning(totalmin)
		
        if media <> "export" AND (showifnul = 0 OR (showifnul = 1 AND overtid > 0)) then   
		%>
        <tr>
            <td colspan="1" style="height:30px; background-color:beige;">&nbsp;</td>
            <td colspan="5" style="background-color:beige; padding:2px 2px 2px 2px;"><span style="font-size:8px; line-height:12px;"><%=left(weekdayname(weekday(datoOvertid, 1)), 3) & ". "& formatdatetime(datoOvertid, 2) %></span><br /> Overtid (godkendt):</td>
            <td style="background-color:beige;" align="right"><b><%=thoursTot%>:<%=left(tminTot, 2)%></b></td>
            <td style="background-color:beige;" colspan="5">&nbsp;</td>

        </tr>
		 <%
         end if


        if cint(showTot) <> 1 then

            if (showifnul = 0 OR (showifnul = 1 AND overtid > 0)) then
             minutterExpTxtOtid = thoursTot&":"&left(tminTot, 2)
             stempelUrEkspTxto = stempelUrEkspTxto & meNavn &";"& meInit & ";" & datoOvertid & ";;;"& minutterExpTxtOtid &";;;;Overtid;;;;xx99123sy#z"
            end if

        end if
         %>
               

        <%end select


    end function
   
   %>