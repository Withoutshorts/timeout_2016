<% 
function indlaesTillaeg(lto)

                'EPI2017 lufthavn
        

                '*** Indl�ser till�g fra AVI DK
               if ((session("mid") = 1 OR session("mid") = 2694 Or session("mid") = 2452 OR session("mid") = 32821 OR session("mid") = 2663) AND (lto = "epi2017" OR lto = "xintranet - local")) then    

                '*EPI
                ddTjkSupplementSL = year(now) & "/" & month(now) & "/"& day(now) 

                ddTjkSupplementST = dateAdd("d", -45, now)
                ddTjkSupplementST = year(ddTjkSupplementST) & "/" & month(ddTjkSupplementST) & "/"& day(ddTjkSupplementST) 
                strSQLepiAViDKsupplement = "SELECT tid, tdato, sttid, sltid, tmnr FROM timer WHERE taktivitetnavn = 'Data Collection' AND tdato BETWEEN '"& ddTjkSupplementST &"' AND '"& ddTjkSupplementSL &"' AND sttid <> '00:00:00' AND (origin = 11 OR origin = 12) AND overfort = 0"
                'Local
                'strSQLepiAViDKsupplement = "SELECT tid, tdato, sttid, sltid FROM timer WHERE tjobnavn = 'A-SK Restest simpel' AND taktivitetnavn = 'Support U/B' AND tdato BETWEEN '2017-06-15' AND '2017-07-01'"
               
                addSupplement15 = 0
                addTillaeg = 0
                oRec6.open strSQLepiAViDKsupplement, oConn, 3
                while not oRec6.EOF

                addSupplement15 = 0
                addTillaeg = 0
                addTillaeg8 = 0
                addTillaeg20 = 0
                addTillaegW = 0
                'addTillaeg20

                    tregDato = oRec6("tdato")
                    sttidthis = oRec6("tdato") &" "& formatdatetime(oRec6("sttid"), 3)
                    sltidthis = oRec6("tdato") &" "& formatdatetime(oRec6("sltid"), 3)

                    '**** F�R 08:00
                    if formatdatetime(oRec6("sttid"), 3) < "08:00:00" then
                        addSupplement15 = 1

                        if formatdatetime(oRec6("sltid"), 3) > "08:00:00" then
                            sluttidKri = oRec6("tdato") &" 08:00:00"
                        else
                            sluttidKri = sltidthis
                        end if

                        'Response.write "<br>"& oRec6("tid") &"sttidthis:"& sttidthis &" #"& sltidthis & "<br>"
                        addTillaeg8 = dateDiff("n", sttidthis, sluttidKri, 2,2)
                        addTillaeg8 = (addTillaeg8/60) 
                        
                      
                     end if

                     '**** Efter 20:00
                      if formatdatetime(oRec6("sltid"), 3) > "20:00:00" then
                        addSupplement15 = 1

                        if formatdatetime(oRec6("sttid"), 3)  > "20:00:00" then
                           starttidKri = sttidthis
                        else
                           starttidKri = oRec6("tdato") &" 20:00:00"
                        end if

                        'Response.write "<br>"& oRec6("tid") &"sttidthis:"& sttidthis &" #"& sltidthis & "<br>"
                        addTillaeg20 = dateDiff("n", starttidKri, sltidthis,2,2)
                        addTillaeg20 = (addTillaeg20/60) 
                        
                        

                     end if


               
                    '**** L�r / S�n ELLER Hellig ****'
                   
                     call helligdage(oRec6("tdato"), 0, lto, oRec6("tmnr"))
                     if datePart("w", oRec6("tdato"), 2,2) = "6" OR datePart("w", oRec6("tdato"), 2,2) = "7" OR cint(erHellig) = 1 then
                        addSupplement15 = 1

                        'Response.write "<br>"& oRec6("tid") &"sttidthis:"& sttidthis &" #"& sltidthis & "<br>"
                        addTillaegW = dateDiff("n", sttidthis, sltidthis,2,2)
                        addTillaegW = (addTillaegW/60) 
                        'addTillaegW = 5
                    end if


                  

                    addTillaeg = formatnumber(addTillaeg8/1 + addTillaeg20/1 + addTillaegW/1, 2)
                    addTillaeg = replace(addTillaeg, ".", "") 
                    addTillaeg = replace(addTillaeg, ",", ".")   
                    if cint(addSupplement15) = 1 then

                        strSQLTidCopy = "INSERT INTO timer_tmp_for_copy SELECT * FROM timer WHERE tid = "& oRec6("tid")
                        oConn.execute(strSQLTidCopy)

                        lastTid = 0
                        strSQLTidLast = "SELECT tid FROM timer WHERE tid > 0 ORDER BY tid DESC LIMIT 1"
                        oRec5.open strSQLTidLast, oConn, 3
                        if not oRec5.EOF then

                            lastTid = oRec5("tid") 

                        end if
                        oRec5.close

                        strSQLTidUpdateAddOne = "UPDATE timer_tmp_for_copy SET tid = "& lastTid + 1
                        oConn.execute(strSQLTidUpdateAddOne)
                        
                        strSQLTidInsert = "INSERT INTO timer SELECT * FROM timer_tmp_for_copy WHERE tid = " & lastTid + 1
                        oConn.execute(strSQLTidInsert)

                        strSQLTidInsertDEL = "DELETE FROm timer_tmp_for_copy"
                        oConn.execute(strSQLTidInsertDEL)

                        strSQLTidUpdOfort = "UPDATE timer SET overfort = 1, overfortdt = '"& ddTjkSupplementSL &"' WHERE tid = " & oRec6("tid")
                        oConn.execute(strSQLTidUpdOfort)

                        strSQLTidUpdNew = "UPDATE timer SET tfaktim = 51, taktivitetnavn = 'Supplement 15', timer = "& addTillaeg &", origin = 112, editor = 'Supplement Copy' WHERE tid = " & lastTid + 1
                        oConn.execute(strSQLTidUpdNew)


                    end if
            
                    
                oRec6.movenext
                wend 
                oRec6.close 

               end if

                'Response.write "Overf�rt OK"

		        'Response.end

end function



public timertilindlasning, dtIndRegDato, medidindlaes, aid_usefortillag
function fyldoptimertilnorm(lto, mid, forcerun, stdato, sldato)

        call akttyper2009(2)
        medidindlaes = mid

       
    
        
        timertilindlasningSum = 0
        
        if cint(forcerun) <> 0 then
            dd = day(sldato) & "/" & month(sldato) & "/" & year(sldato)
        else
            dd = day(now) & "/" & month(now) & "/" & year(now)
        end if
        
        
        
        call meStamdata(mid)
        

        strSQLmansat = " mansat <> 0 AND mansat <> 4"
        

        select case lto
        case "intranet - local", "lm"


                call medarbiprojgrp(4, mid, 0, -1)
       
                if instr(instrMedidProgrp, "#"& mid &"#,") <> 0 then
                mediFar = 1              
                else
                mediFar = 0               
                end if

                if cint(mediFar) = 1 then
                aid_usefortillag = "35402"
                else
                aid_usefortillag = "32023"
                end if

        runday = 0
        eachDayOrSum = 0
        stpoint = 1

        endpoint = 30
        case "dencker"


                call medarbiprojgrp(41, mid, 0, -1)
       
                if instr(instrMedidProgrp, "#"& mid &"#,") <> 0 then
                mediTimelon = 1              
                else
                mediTimelon = 0               
                end if

        runday = 1
        aid_usefortillag = "32657"
        stpoint = 3
        endpoint = stpoint + 13
        eachDayOrSum = 1
       
        case else
        eachDayOrSum = 0
        runday = 0
        aid_usefortillag = "0"
        stpoint = 1
        endpoint = 30
        end select

        aktid = aid_usefortillag
        ddthisWeekDay = datepart("w",dd,2,2)

        '** Skiftet ud med selvvalgt datointerval
        'if cint(forcerun) = 1 then 'k�rt manuelt fra HR siden - find fredag i ugen f�r + 14 dage
        '    select case ddthisWeekDay
        '    case 1
        '        stpoint = 3
        '    case 2
        '        stpoint = 4
        '    case 3
        '        stpoint = 5
        '    case 4
        '        stpoint = 6
        '    case 5
        '        stpoint = 7
        '    case 6
        '        stpoint = 8
        '    case 7
        '        stpoint = 9
        '    end select

         '   endpoint = stpoint + 13

        'end if

        if cint(forcerun) = 1 then 'k�rt manuelt fra HR siden - find fredag i ugen f�r + 14 dage
            stpoint = 0
            endpoint = datediff("d", stdato, sldato, 2,2)
        end if

        '** K�r alle dage K�r kun specifik dag
        '** Husk p�skemandag, 2 juledag osv.


      

        if cint(runday) = 0 OR cint(runday) = cint(ddthisWeekDay) OR cint(forcerun) = 1 then

            for d = stpoint to endpoint 'last 13 OR 30 days STARTER fra ig�r



            ddthis = dateadd("d", -d, dd)
            call thisWeekNo53_fn(ddthis)
            ddthisWeek = thisWeekNo53 'datepart("ww",ddthis,2,2)
       
            dtInd = year(ddthis) &"/"& month(ddthis) &"/"& day(ddthis)
           

            tidspunkt = formatdatetime(now, 3)

            if cint(eachDayOrSum) = 0 then
            dtIndRegDato = dtInd
            else
                if d = stpoint then 'fredag
                dtIndRegDato = year(ddthis) &"/"& month(ddthis) &"/"& day(ddthis)
                end if
            end if
        

            '*** Kun hvis uge ikke er godkendt
            erUgeGodkendt = 0
            strSQLlogind = "SELECT uge, ugegodkendt FROM ugestatus WHERE WEEK(uge, 3) = '"& ddthisWeek &"' AND YEAR(uge) = '"& year(ddthis) &"' AND mid = "& mid & " AND splithr = 0"
            oRec9.open strSQLlogind, oConn, 3
            if not oRec9.EOF then

            erUgeGodkendt = oRec9("ugegodkendt")

            end if
            oRec9.close

            'if mid = 17 AND year(ddthis) = "2019" AND ddthisWeek = "48" then
            'erUgeGodkendt = 0
            'end if


            if cint(erUgeGodkendt) <> 1 then

            '** Renser ud i tidligere opfyldning ***'
            strSQLdel = "DELETE FROM timer WHERE origin = 479 AND tmnr = "& mid &" AND TAktivitetId = "& aid_usefortillag &" AND tdato = '"& year(dtInd) &"/"& month(dtInd) & "/"& day(dtInd) &"'"
            oConn.execute(strSQLdel)

            '** logget ind
            minutterThisDay = 0
            strSQLlogind = "SELECT login, logud, SUM(minutter) AS summin FROM login_historik WHERE dato = '"& dtInd &"' AND mid = "& mid & " AND stempelurindstilling <> -1 AND logud IS NOT NULL AND minutter <> 0 GROUP BY mid ORDER BY id DESC"
            
            'if session("mid") = 1 then
            'Response.write "<br>strSQLlogind: " & strSQLlogind & "<br>minutterThisDay: "& minutterThisDay& "<br>"
            'end if

            oRec9.open strSQLlogind, oConn, 3
            if not oRec9.EOF then

            minutterThisDay = oRec9("summin")

            end if
            oRec9.close

            
            strSQLlogind = "SELECT login, logud, SUM(minutter) AS summin FROM login_historik WHERE dato = '"& dtInd &"' AND mid = "& mid & " AND stempelurindstilling = -1 AND logud IS NOT NULL AND minutter <> 0 GROUP BY mid ORDER BY id DESC"
           
            'if session("mid") = 21 then
            'Response.write "<br>strSQLlogind MINUS: " & strSQLlogind & "<br>minutterThisDay: "& minutterThisDay& "<br>"
            'end if

            oRec9.open strSQLlogind, oConn, 3
            if not oRec9.EOF then

            minutterThisDay = minutterThisDay - oRec9("summin")

            end if
            oRec9.close

            minutterThisDay = minutterThisDay/60
           


            '*** Till�g til komme/g�
            tillagtimerThisDay = 0
            strSQLtimer = "SELECT SUM(timer) AS timer FROM timer t WHERE tdato = '"& dtInd &"' AND tmnr = "& mid & " AND ("& aty_sql_tilwhours &") GROUP BY tmnr"
            
            'response.write "strSQLtimer: " & strSQLtimer
            'response.flush
    
            oRec9.open strSQLtimer, oConn, 3
            if not oRec9.EOF then

            tillagtimerThisDay = oRec9("timer")

            end if
            oRec9.close

            '*** Fradrag til komme/g�
            fradragtimerThisDay = 0
            strSQLtimer = "SELECT SUM(timer) AS timer FROM timer t WHERE tdato = '"& dtInd &"' AND tmnr = "& mid & " AND ("& aty_sql_frawhours &") GROUP BY tmnr"
            oRec9.open strSQLtimer, oConn, 3
            if not oRec9.EOF then

            fradragtimerThisDay = oRec9("timer")

            end if
            oRec9.close


            '***** till�g/fradgr til stempelur
            minutterThisDay = minutterThisDay + (tillagtimerThisDay*1) - (fradragtimerThisDay*1)


            select case lto
            case "dencker"

                    timertilindlasning  = 0
                    if cdbl(minutterThisDay) > 0 then

                        '*** Norm
                        call normtimerPer(mid, ddthis, 0, 0)
                        timertilindlasning = (ntimper - minutterThisDay)

                    end if

                


            case else


                    '*** Real
                    timerThisDay = 0
                    strSQLtimer = "SELECT SUM(timer) AS timer FROM timer t WHERE tdato = '"& dtInd &"' AND tmnr = "& mid & " AND ("& aty_sql_realhours &") GROUP BY tmnr"
                    oRec9.open strSQLtimer, oConn, 3
                    if not oRec9.EOF then

                    timerThisDay = oRec9("timer")

                    end if
                    oRec9.close


                   
                        '*** Norm
                        call normtimerPer(mid, ddthis, 0, 0)


                        'if session("mid") = 1 then
                        'ntimper = 7
                        'end if

                        '** LM
                        '** ALTID FYLD OP TIL Komme/G� 
                        'if minutterThisDay <= ntimper then 'Er komme g� tid st�rre end norm fyldes op hertil
                        '** Kun dage med NORM tid *
                      timertilindlasning  = 0
                     'if cdbl(ntimper) > 0 then 'cdbl(minutterThisDay) > 0 OR 

                        if cdbl(ntimper) > 0 then

                            if minutterThisDay > 0 then
                            timertilindlasning = (minutterThisDay - timerThisDay)
                            else
                            timertilindlasning = 0 'ntimper
                            end if

                        else
                        timertilindlasning = 0
                        end if
                        
                        'timertilindlasning = (ntimper - timerThisDay)
                        'else
                        'timertilindlasning = (minutterThisDay - timerThisDay)
                        'end if

                    'end if

            end select

      
        
           

                    '**** Overtid ********************
                    if cint(eachDayOrSum) = 0 then
                        if cdbl(timertilindlasning) > 0 then
                         call indlaesTimerOpTilNorm
                        end if 'timer til indl�sning
                    else
                        timertilindlasningSum = ((timertilindlasningSum*1) + (timertilindlasning*1))
                    end if
               
                end if 'ugegodkendt


                   'if session("mid") = 1 AND mid = 17 AND year(ddthis) = "2019" AND ddthisWeek = "48" then

                   'Response.write "<br><hr>timertilindlasning: " & timertilindlasning & " aid: " & aid & " aid_usefortillag: "& aid_usefortillag &"<br>ntimper: "& ntimper & "<br>"
                   'Response.write "strSQLtimer: " & strSQLtimer & "<br>"
                   'Response.write "strSQLlogind: " & strSQLlogind & "<br>"
                    'Response.write strSQLdel
                    'Response.end
        
                    'end if


              next


                    'if session("mid") = 1 then
                    'Response.end
                    'end if
    

                    '**** Overtid SUM ********************
                    if cint(eachDayOrSum) = 1 AND cint(mediTimelon) = 1 then
                        
                        if cdbl(timertilindlasningSum) <> 0 then
                            timertilindlasning = timertilindlasningSum
                            call indlaesTimerOpTilNorm
                        end if 'timer til indl�sning
                    
                    end if


             end if 'runday
      



end function


sub indlaesTimerOpTilNorm

    timertilindlasning = formatnumber(timertilindlasning, 2)
    timertilindlasning = replace(timertilindlasning, ".", "")
    timertilindlasning = replace(timertilindlasning, ",", ".")


    
                '*** Finder AKT data ***
                        
                    'aktid = aid
                    aktnavn = "Activity NOT found"
                    fakturerbar = 0 
                    kid = 0
                    kkundenavn = "Customer NOT found"
                    jobnavn = "Job NOT found"
                    jobnr = "0"
                    aktid = aid_usefortillag


                    strSQLakt = "SELECT a.id, a.navn, a.fakturerbar, k.kkundenavn, kid, jobnavn, jobnr FROM aktiviteter a "_
                    & "LEFT JOIN job j ON (j.id = a.job) "_
                    & "LEFT JOIN kunder k ON (k.kid = j.jobknr) WHERE a.id = " & aktid & " AND aktstatus = 1"

                    'response.write strSQLakt
                    'response.flush

                    oRec9.open strSQLakt, oConn, 3
                    if not oRec9.EOF then

                    aktid = oRec9("id")
                    aktnavn = oRec9("navn")
                    fakturerbar = oRec9("fakturerbar")
                    kid = oRec9("kid")
                    kkundenavn = oRec9("kkundenavn")
                    jobnavn = oRec9("jobnavn") 
                    jobnr = oRec9("jobnr")

                    end if
                    oRec9.close

                               
                    '*** Indl�ser Overtid til godkendelse ***'
		            strSQLKoins = "INSERT INTO timer "_
		            &"("_
		            &" timer, tfaktim, tdato, tmnavn, tmnr, tjobnavn, tjobnr, tknavn, tknr, "_
		            &" TAktivitetId, taktivitetnavn, Taar, TimePris, TasteDato, fastpris, tidspunkt, "_
		            &" editor, kostpris, seraft, sttid, sltid, valuta, kurs, origin "_
		            &") "_
		            &" VALUES "_
		            &" (" _
		            & timertilindlasning &", "& fakturerbar &", "_
		            & "'"& year(dtIndRegDato) &"/"& month(dtIndRegDato) & "/"& day(dtIndRegDato) &"', "_
		            & "'"& meNavn &"', "_
		            & medidindlaes &", "_
		            & "'"& jobnavn &"', "_
		            & "'"& jobnr &"', "_
		            & "'"& kkundenavn &"', "& kid &", "_
		            & aktid &", "_
		            & "'"& aktnavn &"', "_
		            & year(now) &", 0, "_
		            & "'"& year(now) &"/"& month(now) & "/"& day(now) &"', 0, "_
		            & "'"& tidspunkt &"', "_
		            & "'Timeout till�g indl�sning', 0, 0, '00:00:00', '00:00:00', 1, 100, 479)"

		                
            'if session("mid") = 1 then
            'Response.write strSQLKoins & "<br>"
            'Response.flush
            'Response.end
            'end if

            oConn.execute(strSQLKoins)

end sub



function overtidsTillaeg(stempelurindstilling, lto, login, logud, mid, timertilindlasning, tfaktim)
    
        'Cflow
        'WAP?
        '** Indl�ser �nsket beordret overarbejde ***'
    
        dtInd = year(login) &"/"& month(login) &"/"& day(login)
        tidspunkt = formatdatetime(now, 3)

      


        '**** Overtid ********************
        if cdbl(timertilindlasning) > 0 then
        timertilindlasning = replace(timertilindlasning, ",", ".")



        '*** KLARG�RING afrunding til 15 min ****
        'if lto = "cflow" then 'afrund 15
    
        'call timerRound15_fn(timertilindlasning)
        'timertilindlasning = timerRound15

        'end if


        '**** Tfaktim '**** 
        'if session("mid") = 1 then
        call normtimerPer(mid, login, 0, 0)

        call meStamdata(mid)

              

                                '*** Finder AKT data ***
                        
                                aktid = 0
                                aktnavn = "Activity NOT found"
                                fakturerbar = 0 
                                kid = 0
                                kkundenavn = "Customer NOT found"
                                jobnavn = "Job NOT found"
                                jobnr = "0"


                                strSQLakt = "SELECT a.id, a.navn, a.fakturerbar, k.kkundenavn, kid, jobnavn, jobnr FROM aktiviteter a "_
                                & "LEFT JOIN job j ON (j.id = a.job) "_
                                & "LEFT JOIN kunder k ON (k.kid = j.jobknr) WHERE a.fakturerbar = " & tfaktim & " AND aktstatus = 1"
                                oRec9.open strSQLakt, oConn, 3
                                if not oRec9.EOF then

                                aktid = oRec9("id")
                                aktnavn = oRec9("navn")
                                fakturerbar = oRec9("fakturerbar")
                                kid = oRec9("kid")
                                kkundenavn = oRec9("kkundenavn")
                                jobnavn = oRec9("jobnavn") 
                                jobnr = oRec9("jobnr")

                                end if
                                oRec9.close

                                '*** Indl�ser Overtid til godkendelse ***'
		                        strSQLKoins = "INSERT INTO timer "_
		                        &"("_
		                        &" timer, tfaktim, tdato, tmnavn, tmnr, tjobnavn, tjobnr, tknavn, tknr, "_
		                        &" TAktivitetId, taktivitetnavn, Taar, TimePris, TasteDato, fastpris, tidspunkt, "_
		                        &" editor, kostpris, seraft, sttid, sltid, valuta, kurs, origin "_
		                        &") "_
		                        &" VALUES "_
		                        &" (" _
		                        & timertilindlasning &", "& fakturerbar &", "_
		                        & "'"& year(dtInd) &"/"& month(dtInd) & "/"& day(dtInd) &"', "_
		                        & "'"& meNavn &"', "_
		                        & mid &", "_
		                        & "'"& jobnavn &"', "_
		                        & "'"& jobnr &"', "_
		                        & "'"& kkundenavn &"', "& kid &", "_
		                        & aktid &", "_
		                        & "'"& aktnavn &"', "_
		                        & year(now) &", 0, "_
		                        & "'"& year(now) &"/"& month(now) & "/"& day(now) &"', 0, "_
		                        & "'"& tidspunkt &"', "_
		                        & "'Timeout till�g indl�sning', 0, 0, '00:00:00', '00:00:00', 1, 100, 479)"

		                
                       'if session("mid") = 1 then
                       'Response.write strSQLKoins & "<br>"
                       'Response.flush
                       'Response.end
                       'end if

                        oConn.execute(strSQLKoins)
        
    
    
                         '***** Oprettter Mail object ***

                            'response.write request.servervariables("PATH_TRANSLATED")
                            'response.flush
                                    sm = 0
                                    if sm = 1 then

			                        if (request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\sesaba.asp") then

 
			                        Set myMail=CreateObject("CDO.Message")
                                    myMail.Subject="TimeOut - Medarbejder "& meNavn &" har �nsket reisetid/overtid"
                                    myMail.From = "timeout_no_reply@outzource.dk"
				      
                                    
			                        'Modtagerens navn og e-mail
			                        select case lto
			                        case "cflow" 

                                    myMail.To= "frode@cflow.no; rune@cflow.no"
                                    'myMail.To= "Support <support@outzource.dk>"
                                    'myMail.Cc= ""

                                    'if len(trim(meEmail)) <> 0 then
                                    'Mailer.AddCC ""& meNavn &"", ""& meEmail &""
                                    'myMail.Cc= ""& meNavn &"<"& meEmail &">"
                                    'end if
                                    
			                       
			                        case else
			                        'Mailer.AddRecipient "OutZourCE", "timeout_no_reply@outzource.dk"
                                    myMail.To= "support@outzource.dk"
			                        end select
                        			
			                       

                                    timertilindlasningMail = replace(timertilindlasning, ".", ",")
                        			
			                                ' Selve teksten
					                        txtEmailHtml = "<br>Medarbejder "& meNavn &" har �nsket reisetid/overtid/till�g d. "& weekdayname(weekday(login, 1)) &" d. "& formatdatetime(login, 2) &"<br><br>"
					                       
                                            if tfaktim = 51 OR tfaktim = 54 then 'OVertid
                                            txtEmailHtml = txtEmailHtml & "Der er �nsket: "& timertilindlasning &" timers overtid. <br>"
                                            end if

                                            if tfaktim = 52 OR tfaktim = 53 OR tfaktim = 55 then 'Reisetid
                                            txtEmailHtml = txtEmailHtml & "Der er �nsket: "& timertilindlasning &" timers reisetid. <br>"
                                            end if

                                            if tfaktim = 50 OR tfaktim = 60 OR tfaktim = 61 then 'Till�g
                                            txtEmailHtml = txtEmailHtml & "Der er �nsket: "& timertilindlasning &" stk. till�g. <br>"
                                            end if

                                            txtEmailHtml = txtEmailHtml & "Komme/G� tid er: "& left(formatdatetime(login, 3), 5) & " - "& left(formatdatetime(logud, 3), 5) &"<br><br><br>"_ 
                                            & "Du kan godkende timerne ved at klikke <a href=""http://timeout.cloud/timeout_xp/wwwroot/ver2_14/to_2015/godkend_request_timer.asp?usemrn="&mid&"&lto="&lto&""">her..</a><br><br><br>"_
					                        & "Med venlig hilsen<br>TimeOut Stempelur Service<br>" 
					                        
					                      
                                               myMail.htmlBody= "<html><head><title></title>"_
			                                   &"</head>"_
			                                   &"<body>"_ 
			                                   & txtEmailHtml & "</body></html>"

                                           'myMail.AddAttachment "d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_14\inc\log\data\"& file

                                            myMail.Configuration.Fields.Item _
                                            ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
                                            'Name or IP of remote SMTP server
                                   
                                           
                                            smtpServer = "formrelay.rackhosting.com" 
                                           
                    
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


                        				
                        			
			                        end if ''** Mail


                                    end if 'SM = 0


    
    
          
        end if



end function





function maksFlexControl(lto, mid, timertilindlasning)

        '*** 20190110 - T�nk over at gemme talle hver FY, da indl�sningern over tid vil blive langsom hvis den altid skal g� tilbage til 1.1.2018 / Licensstart
        'maksFlexControl(
        'fn_flexSaldoFYreal_norm(
        'maxflexControlInsert(

        timertilindlasning = 0

        if len(trim(mid)) <> "" AND instr(mid, ",") = 0 then
        mid = mid
        else
        mid = 0
        end if

        call licensStartDato()
        call meStamdata(mid)
      
        if cDate(meAnsatDato) > cDate(licensstdato) then
        stDatoNorm = meAnsatDato
        else
        stDatoNorm = licensstdato '"1/1/" & year(now) ''"1/1/2018" '& year(now)
        end if

        
        '** DATO interval *****
        '** Kun til DD. Norm bliver ogs� kun beregnet til DD.
        select case lto
        case "dencker"

        'meType = 40 
        '** Funktion�re pr. uge
        '** Timel�nnede pr. m�ned
        'call medariprogrpFn(mid)
      

        dd = now

                if cint(meType) = 40 then 'timel�nnede
                'dd = dateAdd("d", -1, dd)
                DayOfWeekNo = datepart("w", dd, 2,2)

                Select Case DayOfWeekNo
                    Case 7
                        thisUgeDagMinus = 6
                    Case 1
                        thisUgeDagMinus = 0
                    Case 2
                        thisUgeDagMinus = 1
                    Case 3
                        thisUgeDagMinus = 2
                    Case 4
                        thisUgeDagMinus = 3
                    Case 5
                        thisUgeDagMinus = 4
                    Case 6
                        thisUgeDagMinus = 5
                End Select

                'if session("mid") = 21 then
                'response.write "DayOfWeekNo: " & DayOfWeekNo
                'response.end
                'end if

                thisMonday = dateAdd("d", -(thisUgeDagMinus),dd)
        

                 stDatoSQL = year(thisMonday) &"/"& month(thisMonday) &"/"& day(thisMonday) 'year(now)&"/1/1" 'year(now)&"/1/1"  
                 'slDatoSQL = year(now)&"/12/31"  
        
                 slDatoSQL = dateAdd("d", 6, thisMonday)
                 slDatoSQL = year(slDatoSQL) &"/"& month(slDatoSQL) &"/"& day(slDatoSQL)

                 else

                        'Funktion�rer pr. m�ned
                         thismonth = datepart("m", dd, 2,2)

                         stDatoSQL = year(dd) &"/"& month(dd) &"/1"  
                 
                         slDatoSQL = dateAdd("m", 1, stDatoSQL)
                         slDatoSQL = dateAdd("d", -1, slDatoSQL)
                         slDatoSQL = year(slDatoSQL) &"/"& month(slDatoSQL) &"/"& day(slDatoSQL)


                 end if 'timel�nnede

         tDatothis = year(dd) & "/" & month(dd) & "/"& day(dd) 
         dtInd = tDatothis

        case else
      
         stDatoSQL = year(stDatoNorm) &"/"& month(stDatoNorm) &"/"& day(stDatoNorm) 'year(now)&"/1/1" 'year(now)&"/1/1"  
         'slDatoSQL = year(now)&"/12/31"  
        
         slDatoSQL = dateAdd("d", -1, now)
         slDatoSQL = year(slDatoSQL) &"/"& month(slDatoSQL) &"/"& day(slDatoSQL)

         dtInd = year(now) &"/1/1"

        end select

      
    
    




        select case lto
        case "wap"
        aktid = 34
        maxflxprdag = 10
        aktFaktimType = 114
        case "intranet - local"
        aktid = 6290
        maxflxprdag = 10
        aktFaktimType = 114
        case "wwf"
        aktid = 1243
        maxflxprdag = 24
        aktFaktimType = 114
        case "dencker"
        aktid = 32658
        maxflxprdag = 0
        aktFaktimType = 30

            '*** Nulstiller indl�st overarb. pr. uge **'
            strSqlDel = "DELETE FROM timer WHERE tmnr = "& mid &" AND tfaktim = "& aktFaktimType &" AND origin = 480 AND tdato BETWEEN '"& stDatoSQL &"' AND '"& slDatoSQL &"'" 

            'if session("mid") = 21 AND lto = "dencker" then
            'Response.write strSqlDel
            'Response.end
            'end if
        
            oConn.execute(strSqlDel) 
        

        case else
        aktid = 0
        maxflxprdag = 24
        aktFaktimType = 114
        end select

       

     


       


        '**** Overtid ********************
        'if cdbl(timertilindlasning) > 0 then
       
        
        call akttyper2009(2)

        
      '*** If maxflxprdag <> 0 then
       If maxflxprdag <> 0 then

       '*** Tjekker f�rst MAX 10 timer pr Dag. TJEKKER ALLE DAGE **'
       '*** Nulstiller korrektion. MAX 10 timer pr. dag korrektion **'
       strSqlDel = "DELETE FROM timer WHERE tmnr = "& mid &" AND tfaktim = "& aktFaktimType &" AND origin = 580" 'SLET KUN dag til ag korrektioner
       oConn.execute(strSqlDel)

        '*** tjekker flex saldo : slDatoSQL T.O.M iG�R *** 
        realTimerFYmax10 = 0 
        strSQLtimer = "SELECT sum(timer) AS realTimerFY, tdato, tmnr FROM timer WHERE tmnr = "& mid &" AND tdato BETWEEN '"& stDatoSQL &"' AND '"& slDatoSQL &"' AND ("& aty_sql_realhours &") AND tfaktim <> "& aktFaktimType &" GROUP BY tmnr, tdato" 
        
        'if session("mid") = 1 then
        'Response.Write "Her<br>" & strSQLtimer
        'Response.end
        'end if
    
    
        oRec8.open strSQLtimer, Oconn, 3
        while not oRec8.EOF 

           

            realTimerFYmax10 = oRec8("realTimerFY")

            if cdbl(realTimerFYmax10) > maxflxprdag then

            
                tDatothis = year(oRec8("tdato")) & "/" & month(oRec8("tdato")) & "/"& day(oRec8("tdato")) 
                dtInd = tDatothis

                flexStatusPrDag = (realTimerFYmax10 - maxflxprdag)
                origin = 580
                call maxflexControlInsert(flexStatusPrDag, origin, mid, dtInd, aktid)

            end if

        oRec8.movenext
        wend 
        oRec8.close


        end if' If maxflxprdag <> 0 then

    
        '********************************************************************
        '** Norm FY
        '********************************************************************    
        select case lto
        case "dencker"

            ntimper = 0

        case else

            
            'stDatoNorm = licensstdato '"1/1/"& year(now)  

            slDatoNorm = dateAdd("d", -1, dd) '"1/1/"& year(now)
            'slDatoNorm = dateAdd("d", 0, now)
            slDatoNormPrDdMinus1 = slDatoNorm
            dageDiff = dateDiff("d", stDatoNorm, slDatoNorm, 2,2) 
            'dageDiff = dateDiff("d",stDatoNorm, now, 2,2) 
            tidspunkt = formatdatetime(dd, 3)

            call normtimerPer(mid, stDatoNorm, dageDiff, 0)


        end select

        '** END NORM
    
    
    


        '***************************************************************************
         '*** tjekker GRaND flex saldo / Periode ***
        '***************************************************************************

        realTimerFY = 0 
        strSQLtimer = "SELECT sum(timer) AS realTimerFY FROM timer WHERE tmnr = "& mid &" AND tdato BETWEEN '"& stDatoSQL &"' AND '"& slDatoSQL &"' AND ("& aty_sql_realhours &") AND (tfaktim <> "& aktFaktimType &" ) GROUP BY tmnr" 
        oRec9.open strSQLtimer, Oconn, 3
        if not oRec9.EOF then 

            realTimerFY = oRec9("realTimerFY")

        end if
        oRec9.close


        if len(trim(realTimerFY)) <> 0 then
        realTimerFY = realTimerFY
        else
        realTimerFY = 0
        end if


        '*** Henter korrektion inden indtastning Pr. dag ***
        korrektionMaxFlexTimerFYforprDag = 0 
        strSQLtimer = "SELECT sum(timer) AS realTimerFY FROM timer WHERE tmnr = "& mid &" AND tdato BETWEEN '"& stDatoSQL &"' AND '"& slDatoSQL &"' AND (tfaktim = "& aktFaktimType &"  AND origin = 580) GROUP BY tmnr" 
        oRec9.open strSQLtimer, Oconn, 3
        if not oRec9.EOF then 

            korrektionMaxFlexTimerFYforprDag = oRec9("realTimerFY")

        end if
        oRec9.close


        if len(trim(korrektionMaxFlexTimerFYforprDag)) <> 0 then
        korrektionMaxFlexTimerFYforprDag = korrektionMaxFlexTimerFYforprDag
        else
        korrektionMaxFlexTimerFYforprDag = 0
        end if

        
        
        '*** Henter korrektion inden indtastning GT ALTID indl�st p� 1 jan ***
        korrektionMaxFlexTimerFYforGT = 0 
        strSQLtimer = "SELECT sum(timer) AS realTimerFY FROM timer WHERE tmnr = "& mid &" AND tdato BETWEEN '"& stDatoSQL &"' AND '"& slDatoSQL &"' AND (tfaktim = "& aktFaktimType &"  AND origin = 480) GROUP BY tmnr" 
        oRec9.open strSQLtimer, Oconn, 3
        if not oRec9.EOF then 

            korrektionMaxFlexTimerFYforGT = oRec9("realTimerFY")

        end if
        oRec9.close


        if len(trim(korrektionMaxFlexTimerFYforGT)) <> 0 then
        korrektionMaxFlexTimerFYforGT = korrektionMaxFlexTimerFYforGT
        else
        korrektionMaxFlexTimerFYforGT = 0
        end if
      
        'korrektionMaxFlexTimerFYforGT = 0
        
        select case lto
        case "dencker"

       
        '** beregnes indenfor en uge / M�ned  **'
        if cint(meType) = 40 then 'timel�nnede
        flexLoft = 37
        else
        flexLoft = 74
        end if
        flexmodregn = (realTimerFY*1 - flexLoft*1) 
       
       
                    if cdbl(flexmodregn) > 0 then

                        origin = 480
                        flexmodregn = flexmodregn * -1.5
                        call maxflexControlInsert(flexmodregn, origin, mid, dtInd, aktid)
                
                    end if

        case else

        'flexStatus = ((realTimerFY+timertilindlasning) - ntimPer) 'M� ikke blive h�jere end 10
        flexStatus = ((realTimerFY+timertilindlasning+(korrektionMaxFlexTimerFYforprDag)+(korrektionMaxFlexTimerFYforGT)) - ntimPer) 'M� ikke blive h�jere end 10    

        select case lto
        case "wap"
        flexLoft = 10
        case else
        flexLoft = 10
        end select 
    
        flexmodregn = 0
        if cdbl(flexStatus) > cdbl(flexLoft) then
        flexmodregn = (flexStatus*1-flexLoft*1) 
        end if

      

                        'flexmodregn = flexStatus * -1

                        'end if

                        'if session("mid") = 1 then
                        'Response.write strSQLtimer & "<br>FLEX. bergn: ("& realTimerFY &"+"& timertilindlasning & " + "& korrektionMaxFlexTimerFYforprDag & " + "& korrektionMaxFlexTimerFYforGT &") - "&  ntimPer
                        'Response.write "<br>flexmodregn:" & flexmodregn
                        'Response.end
                        'end if
        
       
                        'Response.end

                                '*** Nulstiller GT korrektion. MAX 10 timer total**'
                                'strSqlDel = "DELETE FROM timer WHERE tmnr = "& mid &" AND tfaktim = 114 AND origin = 480"
                                'oConn.execute(strSqlDel)
        
                                if cdbl(flexmodregn) > 0 then

                                    origin = 480
                                   'if cdbl(korrektionMaxFlexTimerFYforGT) <> 0 then
                                   'korrektionMaxFlexTimerFYforGT = (korrektionMaxFlexTimerFYforGT * -1)
                                   'flexmodregn = flexmodregn + korrektionMaxFlexTimerFYforGT
                                   'end if

                                   call maxflexControlInsert(flexmodregn, origin, mid, dtInd, aktid)
                
                                end if

                            'end if 'timertilindl�srning

                 end select 'lto

end function


function maxflexControlInsert(flexmodregn, origin, mid, dtInd, aktid)

                flexmodregn = flexmodregn * -1
                flexmodregn = replace(flexmodregn, ",", ".")
                'timertilindlasning = replace(timertilindlasning, ",", ".")

                call meStamdata(mid)


                                '*** Finder AKT data ***
                                strSQLakt = "SELECT a.navn, a.fakturerbar, k.kkundenavn, kid, jobnavn, jobnr FROM aktiviteter a "_
                                & "LEFT JOIN job j ON (j.id = a.job) "_
                                & "LEFT JOIN kunder k ON (k.kid = j.jobknr) WHERE a.id = " & aktid
                                oRec9.open strSQLakt, oConn, 3
                                if not oRec9.EOF then

                                aktnavn = oRec9("navn")
                                fakturerbar = oRec9("fakturerbar")
                                kid = oRec9("kid")
                                kkundenavn = oRec9("kkundenavn")
                                jobnavn = oRec9("jobnavn") 
                                jobnr = oRec9("jobnr")

                                end if
                                oRec9.close

                                '*** Indl�ser Overtid til godkendelse ***'
		                        strSQLKoins = "INSERT INTO timer "_
		                        &"("_
		                        &" timer, tfaktim, tdato, tmnavn, tmnr, tjobnavn, tjobnr, tknavn, tknr, "_
		                        &" TAktivitetId, taktivitetnavn, Taar, TimePris, TasteDato, fastpris, tidspunkt, "_
		                        &" editor, kostpris, seraft, sttid, sltid, valuta, kurs, origin "_
		                        &") "_
		                        &" VALUES "_
		                        &" (" _
		                        & flexmodregn &", "& fakturerbar &", "_
		                        & "'"& year(dtInd) &"/"& month(dtInd) & "/"& day(dtInd) &"', "_
		                        & "'"& meNavn &"', "_
		                        & mid &", "_
		                        & "'"& jobnavn &"', "_
		                        & "'"& jobnr &"', "_
		                        & "'"& kkundenavn &"', "& kid &", "_
		                        & aktid &", "_
		                        & "'"& aktnavn &"', "_
		                        & year(now) &", 0, "_
		                        & "'"& year(now) &"/"& month(now) & "/"& day(now) &"', 0, "_
		                        & "'"& tidspunkt &"', "_
		                        & "'Timeout till�g indl�sning', 0, 0, '00:00:00', '00:00:00', 1, 100, "& origin &")"
		                
                                'if session("mid") = 21 then
                                'Response.write strSQLKoins & "<br>"
                                'Response.end
                                'end if

                                oConn.execute(strSQLKoins)


end function


function indlasTimertilUdbetaling(lto, mid, timertilindlasning, akttype)
    
        if len(trim(timertilindlasning )) <> 0 then
     
        dtInd = year(now) &"/"& month(now) &"/"& day(now)
        tidspunkt = formatdatetime(now, 3)


        select case cint(akttype) 
        case 32

                select case lto
                case "cflow"
                aktid = 184
                case "intranet - local"
                aktid = 6291
                case "dencker"
                aktid = 79829
                case else
                aktid = 0
                end select

        case 10

              select case lto
                
                case "dencker"
                aktid = 102587
                case else
                aktid = 0
                end select    

        case 124

              select case lto
                
                case "dencker"
                aktid = 102588
                case else
                aktid = 0
                end select    

        case else

               aktid = 0        

        end select


        if aktid <> 0 then

        timertilindlasning = replace(timertilindlasning, ".", "")
        timertilindlasning = replace(timertilindlasning, ",", ".")
          
  
        call meStamdata(mid)


                                '*** Finder AKT data ***
                                strSQLakt = "SELECT a.navn, a.fakturerbar, k.kkundenavn, kid, jobnavn, jobnr FROM aktiviteter a "_
                                & "LEFT JOIN job j ON (j.id = a.job) "_
                                & "LEFT JOIN kunder k ON (k.kid = j.jobknr) WHERE a.id = " & aktid
                                oRec9.open strSQLakt, oConn, 3
                                if not oRec9.EOF then

                                aktnavn = oRec9("navn")
                                fakturerbar = oRec9("fakturerbar")
                                kid = oRec9("kid")
                                kkundenavn = oRec9("kkundenavn")
                                jobnavn = oRec9("jobnavn") 
                                jobnr = oRec9("jobnr")

                                end if
                                oRec9.close

                                '*** Indl�ser Overtid til godkendelse ***'
		                        strSQLKoins = "INSERT INTO timer "_
		                        &"("_
		                        &" timer, tfaktim, tdato, tmnavn, tmnr, tjobnavn, tjobnr, tknavn, tknr, "_
		                        &" TAktivitetId, taktivitetnavn, Taar, TimePris, TasteDato, fastpris, tidspunkt, "_
		                        &" editor, kostpris, seraft, sttid, sltid, valuta, kurs, origin "_
		                        &") "_
		                        &" VALUES "_
		                        &" (" _
		                        & timertilindlasning &", "& fakturerbar &", "_
		                        & "'"& year(dtInd) &"/"& month(dtInd) & "/"& day(dtInd) &"', "_
		                        & "'"& meNavn &"', "_
		                        & mid &", "_
		                        & "'"& jobnavn &"', "_
		                        & "'"& jobnr &"', "_
		                        & "'"& kkundenavn &"', "& kid &", "_
		                        & aktid &", "_
		                        & "'"& aktnavn &"', "_
		                        & year(now) &", 0, "_
		                        & "'"& year(now) &"/"& month(now) & "/"& day(now) &"', 0, "_
		                        & "'"& tidspunkt &"', "_
		                        & "'Timeout till�g indl�sning', 0, 0, '00:00:00', '00:00:00', 1, 100, 481)"
		                
                       
                                'if lto = "dencker" AND session("mid") = 21 then
                                'Response.write strSQLKoins & "<br>"
                                'Response.flush
                                'Response.end
                                'end if

                                oConn.execute(strSQLKoins)
        



                
                                '** Fradrag p� Overarbejde 1:5 minus
                                select case lto
                                case "dencker"
     
                                        if cint(akttype) = 32 then

                                            aktid = 32658
                                            timertilindlasning = timertilindlasning*-1.5
                                            timertilindlasning = replace(timertilindlasning, ".", "")
                                            timertilindlasning = replace(timertilindlasning, ",", ".")

                                            '*** Finder AKT data ***
                                            strSQLakt = "SELECT a.navn, a.fakturerbar, k.kkundenavn, kid, jobnavn, jobnr FROM aktiviteter a "_
                                            & "LEFT JOIN job j ON (j.id = a.job) "_
                                            & "LEFT JOIN kunder k ON (k.kid = j.jobknr) WHERE a.id = " & aktid
                                            oRec9.open strSQLakt, oConn, 3
                                            if not oRec9.EOF then

                                            aktnavn = oRec9("navn")
                                            fakturerbar = oRec9("fakturerbar")
                                            kid = oRec9("kid")
                                            kkundenavn = oRec9("kkundenavn")
                                            jobnavn = oRec9("jobnavn") 
                                            jobnr = oRec9("jobnr")

                                            end if
                                            oRec9.close

                                            '*** Indl�ser Overtid til godkendelse ***'
		                                    strSQLKoins = "INSERT INTO timer "_
		                                    &"("_
		                                    &" timer, tfaktim, tdato, tmnavn, tmnr, tjobnavn, tjobnr, tknavn, tknr, "_
		                                    &" TAktivitetId, taktivitetnavn, Taar, TimePris, TasteDato, fastpris, tidspunkt, "_
		                                    &" editor, kostpris, seraft, sttid, sltid, valuta, kurs, origin "_
		                                    &") "_
		                                    &" VALUES "_
		                                    &" (" _
		                                    & (timertilindlasning) &", "& fakturerbar &", "_
		                                    & "'"& year(dtInd) &"/"& month(dtInd) & "/"& day(dtInd) &"', "_
		                                    & "'"& meNavn &"', "_
		                                    & mid &", "_
		                                    & "'"& jobnavn &"', "_
		                                    & "'"& jobnr &"', "_
		                                    & "'"& kkundenavn &"', "& kid &", "_
		                                    & aktid &", "_
		                                    & "'"& aktnavn &"', "_
		                                    & year(now) &", 0, "_
		                                    & "'"& year(now) &"/"& month(now) & "/"& day(now) &"', 0, "_
		                                    & "'"& tidspunkt &"', "_
		                                    & "'Timeout till�g indl�sning - fradrag', 0, 0, '00:00:00', '00:00:00', 1, 100, 482)"
		                
                       
                                            'if lto = "dencker" AND session("mid") = 21 then
                                            'Response.write strSQLKoins & "<br>"
                                            'Response.flush
                                            'Response.end
                                            'end if

                                            oConn.execute(strSQLKoins)
        
                                        end if


                                end select


               end if
                   

            

            end if 'timertilindl�srning

end function





public timerthis_mtx
function matrixtimespan_2018(medid, idag, mtrx, sTtid, sLtid, datoThis, lto)

                            'if session("mid") = 1 then
                            'Response.write "<br>HER mtrx: "& mtrx &"<br>"
                            'end if

                            useDate = idag '"01-01-" & year(now)
                            sTtid_org = sTtid


                            'Response.write "useDate: "& useDate & " " & sTtid
                            'Response.flush

                            call helligdage(datoThis, 0, lto, medid)
                
                            '** Hvilken ugedag
                            thisDayWeekday_mtrx = datepart("w", useDate, 2,2)

                           

                    
                            select case lto
                            case "cflow", "intranet - local"


                                         '** Er medarbejder med i Aumatisjon Eller Enginnering
                                        call medariprogrpFn(medid)

                                       'lp = 9

                                       if instr(medariprogrpTxt, "#14#") <> 0 OR instr(medariprogrpTxt, "#16#") <> 0 OR instr(medariprogrpTxt, "#3#") <> 0 OR instr(medariprogrpTxt, "#19#") <> 0 OR instr(medariprogrpTxt, "#15#") <> 0 then
                                       'if lp = 0 then

                                              


                                                        select case mtrx
                                                        case 0 'Dag ordin�r type 1
                                                                select case thisDayWeekday_mtrx
                                                                case 6,7 'l�rdag og s�ndag skal v�re 100% overtid alle timer
                                                                lb = "00:00:01"
                                                                ub = "23:59:59"
                                                                case else
                                                                lb = "07:30:00"
                                                                ub = "16:00:00"
                                                                end select
                                                        case 1 'Dag 50% type 54
                                                                lb = "16:00:00"
                                                                ub = "20:00:00"
                                                        case 2 'Dag 100%
                                                                lb = "20:00:00"
                                                                ub = "23:59:59"
                                                        case 8 'Efter midnat 1 time til timebank
                                                                lb = "00:00:00"
                                                                ub = "06:00:00"
                                                        case else
                                                                lb = "00:00:00"
                                                                ub = "00:00:00"
                                                        end select

                                            
                                            

            
                                        else

                                                '** Produktion
                                                select case mtrx
                                                case 0 'Dag ordin�r type 1
                                                        select case thisDayWeekday_mtrx
                                                        case 6,7 'l�rdag og s�ndag skal v�re 100% overtid alle timer
                                                        lb = "00:00:01"
                                                        ub = "23:59:59"
                                                        case else
                                                        lb = "07:30:00"
                                                        ub = "15:30:00"
                                                        end select
                                                case 1 'Dag 50% type 54
                                                        lb = "15:30:00"
                                                        ub = "20:00:00"
                                                case 2 'Dag 100%
                                                        lb = "20:00:00"
                                                        ub = "23:59:59"
                                                case 8 'Efter midnat 1 time til timebank
                                                        lb = "00:00:00"
                                                        ub = "06:00:00"
                                                case else
                                                        lb = "00:00:00"
                                                        ub = "00:00:00"
                                                end select



                                        end if




                            case "nonstop"

                                        select case mtrx
                                        case 1 'dag
                                                lb = "07:00:00"
                                                ub = "19:00:00"
                                        case 2,4 'aften / Aften Weekend
                                                lb = "19:00:00"
                                                ub = "23:59:59"
                                        case 3 'weekend
                                               lb = "07:00:00"
                                               ub = "19:00:00"
                                       
                                        case 5 'helligdag
                                               lb = "00:00:01"
                                               ub = "23:59:59" 

                                        case 6,7 'NAT / Morgen + Weekend nat (f�r m�detid)
                                                lb = "00:00:01"
                                                ub = "07:00:00"
                                        end select

                        
                             end select

                                        '***************************
                                        '*** START TID
                                        '***************************

                                        'if session("mid") = 1 then
                                        'Response.write "<hr>" & cDate(useDate &" "& sTtid) &" > "& cDate(useDate &" "& lb)  & " datepart: " & datepart("w", datoThis, 2,2) & "<br>"
                                        'end if


                                        if (((mtrx = 0 OR mtrx = 1 OR mtrx = 2 OR mtrx = 6 OR mtrx = 8) AND (datepart("w", datoThis, 2,2) < 6) AND erHellig <> 1 AND lto <> "cflow") OR _
                                        ((mtrx = 0 OR mtrx = 1 OR mtrx = 2 OR mtrx = 6 OR mtrx = 8) AND lto = "cflow") OR _
                                        ((mtrx = 3 OR mtrx = 4 OR mtrx = 7) AND (datepart("w", datoThis, 2,2) = 6 OR datepart("w", datoThis, 2,2) = 7) AND erHellig <> 1) _
                                        OR (mtrx = 5 AND erHellig = 1)) then 'OR mtrx = 6

                                        

                                                        '*** Beregner timer og klokkeslet efter tarifgr�nser *****
                                                        '06:00-23:00
                                                        '05:00-17:00
                                                        '15:00-03:00
                                                        '09:30-15:00
                                                        '05:00-06:00
                                                        '06:00-07:00 
                                                        '21:00-02:00
                                                        '15:00-23:30
                                                        '21:00-16:00
                                                        'START TID
                                                         if (cDate(useDate &" "& sTtid) > cDate(useDate &" "& lb)) then 
                                                            
                                                                   
                                                                    '6 7 Nat: Start tid vil altid v�re st�rre en 00:00
                                                                    if (cDate(useDate &" "& sTtid) >= cDate(useDate &" "& sLtid) AND mtrx >= 6) then 'NAT                                             
                                                                    sTtid = lb
                                                                    else
                                                                    
                                                                        '21:00-16:00 DAG
                                                                        if (cDate(useDate &" "& sTtid) >= cDate(useDate &" "& sLtid) AND cDate(useDate &" "& sLtid) >= cDate(useDate &" "& lb) AND (mtrx = 1 OR mtrx = 3)) then 'DAG
                                                                        sTtid = lb
                                                                        else     
                                                                        sTtid = sTtid
                                                                        end if

                                                                    end if

                                                        
                                                        else
                                                                  
                                                                   '2 4 Aften
                                                                   if (cDate(useDate &" "& sTtid) >= cDate(useDate &" "& lb) AND (mtrx = 2 OR mtrx = 4)) then 'NAT                                             
                                                                   sTtid = sTtid
                                                                   else
                                                                   sTtid = lb
                                                                   end if
                                                                    
                                                       
                                                        end if




                                                        '***************************
                                                        'SLUT TID
                                                        '***************************
                                                    
                                                        if ((cDate(useDate &" "& sLtid) <= cDate(useDate &" "& ub)) AND (cDate(useDate &" "& sLtid) >= cDate(useDate &" "& lb))) OR _
                                                        (cDate(useDate &" "& ub) < cDate(useDate &" 07:30:00") AND mtrx < 5) then '< 5 IKKE NAT + helligdag aktiviteter 00 - 07
                                                            
                                                                sLtid = sLtid
                                                        
                                                        else
                                                                    
                                                                   '2 4 Aften
                                                                   '15-03
                                                                   if (mtrx = 2 OR mtrx = 4) then 'AFTEN

                                                                            'response.write "<br> = mtrx: "& mtrx &" matrixAktFundet: "& matrixAktFundet & " timerthis_mtx: "& timerthis_mtx & " ub: "& ub &" lb:"& lb &" sTtid: "& sTtid & " sLtid: "& sLtid & " datepart: " & datepart("w", datoThis, 2,2) & "##"& cDate(useDate &" "& sTtid) &"<="& cDate(useDate &" "& sLtid)
                                                                            'response.flush
                                            
                                                                       
                                                                       '09:00 - 15:00 < en start tid for aften
                                                                       '15:00 - 03:00
                                                                       '06:00 - 07:00
                                                                       '05:00 - 02:00
                                                                       '21:00 - 23:00
                                                                       '15:00 - 21:00
                                                                       '05:00 - 23:00
                                                                       if (cDate(useDate &" "& sLtid) <= cDate(useDate &" "& lb)) then
                                                                                
                                                                                    if (cDate(useDate &" "& sTtid_org) >= cDate(useDate &" "& sLtid)) then                                             
                                                                                    sLtid = ub 
                                                                                    else
                                                                                    sLtid =  "00:00" '"00:01" 'sLtid = "00:01" ' = 0 timer. ALTID NUL TIMER
                                                                                    end if
                                                                       else
                                                                       
                                                                                    '** Hvis sluttid > 23:59
                                                                                    if (cDate(useDate &" "& sLtid) => cDate(useDate &" "& ub)) then
                                                                                    sLtid = ud
                                                                                    else
                                                                                    sLtid = sLtid
                                                                                    end if                                                                   

                                                                       end if

                                                                   else
                        
                                                                        if (mtrx = 1 OR mtrx = 3) then 'DAG / WEEKEND DAG

                                                                                '** Hvis sluttid er f�r starttid p� aften 
                                                                                '06:00 - 02:00 --> 12 timer p� dag USE: UB :OK
                                                                                '15:00 - 02:00 --> 4 timer p� Dag  USE: UB :OK
                                                                                '21:00 - 02:00 --> 0 timer p� Dag  USE: UB :OK
                                                                                '06:00 - 07:00 --> 0 timer p� Dag  USE: UB/sl :OK
                                                                                '05:00 - 06:00 --> 0 timer p� Dag  USE: UB: OK
                                                                                '05:00 - 15:00 --> 8 timer p� Dag  USE: sl :OK
                                                                                '09:00 - 15:00 --> 6 timer p� Dag  USE: sl :OK
                                                                                '09:00 - 21:00 --> 10 timer p� Dag USE: UB: OK
                                                                                '21:00 - 16:00 --> 9 timer p� Dag USE: 



                                                                              
                                                                                if (cDate(useDate &" "& sLtid) < cDate(useDate &" "& lb)) AND (cDate(useDate &" "& sTtid_org) <= cDate(useDate &" "& sLtid)) then
                                                                                
                                            
                                                                                sLtid = sLtid
                                                                                else

                                                                                                                                                                      
                                                                                    '** NEW 20180223 CFLOW : AND (cDate(useDate &" "& sLtid) > cDate(useDate &" "& lb)) 
                                                                                    if (cDate(useDate &" "& sLtid) < cDate(useDate &" "& ub)) AND (cDate(useDate &" "& sLtid) > cDate(useDate &" "& lb)) then
                                                                                    sLtid = sLtid
                                                                                    else 
                                                                                    sLtid = ub
                                                                                    end if  
                                                                    
                                                                                end if
                                                                        else
                                                                        sLtid = ub
                                                                        end if 

                                                                   end if
                                                        
                                                        end if


                                                        'select case mtrx
                                                        'case 8 'cflow > 24:00
                                                        '    timerthis_mtx = 1
                                                        'case else
                                                            timerthis_mtx = dateDiff("h", useDate &" "& sTtid, useDate &" "& sLtid, 2,2)
                                                            totalmin = datediff("n", idag &" "& sTtid, idag &" "& sLtid)

                                                            'if session("mid") = 1 then
                                                            'Response.Write "<br>ST: " & idag &" "& sTtid &" SLUT:"& idag &" "& sLtid & " totalmin: " & totalmin
                                                            'end if
                                                                
                                                            if cdbl(totalmin) > 0 then 'Cflow 20180814 

						                                    call timerogminutberegning(totalmin)
						                                    timerthis_mtx = thoursTot&"."&tminProcent 'tminTot
                                                            else
                                                            timerthis_mtx = 0
                                                            end if

						                                'end select

                                                        '23:59:59 // 24:00:00 Problem
	                                                    if mtrx = 2 AND lto = "cflow" AND timerthis_mtx = "3.98" then				 
                                                        timerthis_mtx = 4
                                                        end if

					                                    if timerthis_mtx < 0 then '** Hen over kl 24:00 **'
                                                        timerthis_mtx = 0
						                                'timerthis_mtx = 24 + (replace(timerthis_mtx,".",","))
						                                'timerthis_mtx = replace(timerthis_mtx,",",".")
						                                end if


                                                        '*** Mandag morgen --> Normal dag

                                                        'if session("mid") = 1 then
                                                        'response.write "<br> = mtrx: "& mtrx &" matrixAktFundet: "& matrixAktFundet & " timerthis_mtx: "& timerthis_mtx & " ub: "& ub &" lb:"& lb &" sTtid: "& sTtid & " sLtid: "& sLtid & " datepart: " & datepart("w", datoThis, 2,2) & "##"& cDate(useDate &" "& sTtid) &"<="& cDate(useDate &" "& sLtid)
                                                        'response.flush
                                                        'end if

                                     else        

                                        timerthis_mtx = 0

                                     end if 'mtrx 1-7 + weekend/helligdage   


                            
    
end function   






function indlasFerieSaldo(lto, mid, timertilindlasning, tfaktim, ferieaar)
    
        '*** Ferie optjent
       
        'dtInd = year(now) &"/"& month(now) &"/"& day(now)
        'Er medarbejderen anstat ved indl�sning 1.5? 
         call meStamdata(mid)
     
         if cDate(meAnsatDato) > cDate("01-05-" & ferieaar) then
         dtInd = year(meAnsatDato) &"-"& month(meAnsatDato) &"-"& day(meAnsatDato)
         else
	     dtInd = ferieaar &"-05-01"
         end if
        
        ferieslutDato = (ferieaar + 1) & "-04-30"
        tidspunkt = formatdatetime(now, 3)

        '**** Overtid ********************
        if cdbl(timertilindlasning) > 0 then
        timertilindlasning = replace(timertilindlasning, ",", ".")
    
        call meStamdata(mid)

								aktid = 0
           

								if cint(tfaktim) <> 0 then


									'*** Finder AKT data ***
									strSQLakt = "SELECT a.id, a.navn, a.fakturerbar, k.kkundenavn, kid, jobnavn, jobnr FROM aktiviteter a "_
									& "LEFT JOIN job j ON (j.id = a.job) "_
									& "LEFT JOIN kunder k ON (k.kid = j.jobknr) WHERE a.fakturerbar = " & tfaktim 
									oRec9.open strSQLakt, oConn, 3
									if not oRec9.EOF then

									aktid = oRec9("id")
									aktnavn = oRec9("navn")
									fakturerbar = oRec9("fakturerbar")
									kid = oRec9("kid")
									kkundenavn = oRec9("kkundenavn")
									jobnavn = oRec9("jobnavn") 
									jobnr = oRec9("jobnr")

									end if
									oRec9.close

								end if


							
								if cdbl(aktid) <> 0 then


                                '*** Indl�ser Overtid til godkendelse ***'
		                        strSQLKoins = "INSERT INTO timer "_
		                        &"("_
		                        &" timer, tfaktim, tdato, tmnavn, tmnr, tjobnavn, tjobnr, tknavn, tknr, "_
		                        &" TAktivitetId, taktivitetnavn, Taar, TimePris, TasteDato, fastpris, tidspunkt, "_
		                        &" editor, kostpris, seraft, sttid, sltid, valuta, kurs, origin "_
		                        &") "_
		                        &" VALUES "_
		                        &" (" _
		                        & timertilindlasning &", "& fakturerbar &", "_
		                        & "'"& year(dtInd) &"/"& month(dtInd) & "/"& day(dtInd) &"', "_
		                        & "'"& meNavn &"', "_
		                        & mid &", "_
		                        & "'"& jobnavn &"', "_
		                        & "'"& jobnr &"', "_
		                        & "'"& kkundenavn &"', "& kid &", "_
		                        & aktid &", "_
		                        & "'"& aktnavn &"', "_
		                        & year(now) &", 0, "_
		                        & "'"& year(now) &"/"& month(now) & "/"& day(now) &"', 0, "_
		                        & "'"& tidspunkt &"', "_
		                        & "'"& session("user") &"', 0, 0, '00:00:00', '00:00:00', 1, 100, 779)"

		                
                                'if session("mid") = 1 then
								'Response.write strSQLKoins & "<br>"
								'Response.flush
                                'end if
								'Response.end

								oConn.execute(strSQLKoins)
    
							 end if 'aktid = 0
          
        end if



end function



function indlasTimerTfaktimAktid(lto, mid, timertilindlasning, tfaktim, aktid, insertUpdate, insUpdDate, extsysid, timerkom, koregnr, destination, usebopal)
    
       'insertUpdate = 1 alm, 2 altid p� dd., 3 tjek for tidligere indtasninger (Dencker faste till�g / fradrag for mad der kun m� tastwes en gang pr. dag)

       
        extsysid = extsysid
      
        if aktid <> 0 then
        sqltfaktimaktidKri = "a.id = " & aktid & ""
        else
        sqltfaktimaktidKri = "a.fakturerbar = " & tfaktim & ""
        end if

        if cint(insertUpdate) = 2 then
        dtInd = year(now) &"/"& month(now) &"/"& day(now)
        else
        dtInd = year(insUpdDate) &"/"& month(insUpdDate) &"/"& day(insUpdDate)
        end if

        tidspunkt = formatdatetime(now, 3)


        '**** Overtid ********************
        if cdbl(timertilindlasning) > 0 then
        timertilindlasning = replace(timertilindlasning, ",", ".")
    
        call meStamdata(mid)

                                jobid = 0

                                '*** Finder AKT data ***
                                strSQLakt = "SELECT a.id, a.navn, a.fakturerbar, k.kkundenavn, kid, jobnavn, jobnr, j.id AS jid FROM aktiviteter a "_
                                & "LEFT JOIN job j ON (j.id = a.job) "_
                                & "LEFT JOIN kunder k ON (k.kid = j.jobknr) WHERE " & sqltfaktimaktidKri


                                'response.write strSQLakt
                                'response.flush
                                'response.end

                                oRec9.open strSQLakt, oConn, 3
                                if not oRec9.EOF then

                                jobid = oRec9("jid")
                                aktid = oRec9("id")
                                aktnavn = oRec9("navn")
                                fakturerbar = oRec9("fakturerbar")
                                kid = oRec9("kid")
                                kkundenavn = oRec9("kkundenavn")
                                jobnavn = oRec9("jobnavn") 
                                jobnr = oRec9("jobnr")

                                end if
                                oRec9.close


                         

                        if cdbl(aktid) <> 0 AND cdbl(jobid) <> 0 then

        

                               
                                findesallerede = 0

                                if cint(insertUpdate) = 3 then

                                    '*** Findes indtastning allerede **** f.eks faste till�g Dencker, der ikke skal tilf�jes hver gang der tastes.
                                    strSQLfidnes = "SELECT timer FROM timer WHERE (origin = 781 OR timer = -1) AND TAktivitetId = "& aktid &" AND tmnr = "& mid &" AND tdato =  '"& year(dtInd) &"/"& month(dtInd) & "/"& day(dtInd) &"'"
                                    'response.write "<br>Findes allerede: "& strSQLfidnes &"<br>" & findesallerede & " aktid: "& aktid &" jobid: " & jobid
                                    'response.end
                                    oRec9.open strSQLfidnes, oConn, 3
                                    if not oRec9.EOF then
                                    findesallerede = 1
                                    end if
                                    oRec9.close

                                end if

    
                                 'response.write "<br>Findes allerede: " & findesallerede & " aktid: "& aktid &" jobid: " & jobid
                              


                                '*** Indl�ser Overtid til godkendelse ***'
                                if cint(findesallerede) = 0 then
		                        strSQLKoins = "INSERT INTO timer "_
		                        &"("_
		                        &" timer, tfaktim, tdato, tmnavn, tmnr, tjobnavn, tjobnr, tknavn, tknr, "_
		                        &" TAktivitetId, taktivitetnavn, Taar, TimePris, TasteDato, fastpris, tidspunkt, "_
		                        &" editor, kostpris, seraft, sttid, sltid, valuta, kurs, origin, extsysid, Timerkom, koregnr, destination, bopal "_
		                        &") "_
		                        &" VALUES "_
		                        &" (" _
		                        & timertilindlasning &", "& fakturerbar &", "_
		                        & "'"& year(dtInd) &"/"& month(dtInd) & "/"& day(dtInd) &"', "_
		                        & "'"& meNavn &"', "_
		                        & mid &", "_
		                        & "'"& jobnavn &"', "_
		                        & "'"& jobnr &"', "_
		                        & "'"& kkundenavn &"', "& kid &", "_
		                        & aktid &", "_
		                        & "'"& aktnavn &"', "_
		                        & year(now) &", 0, "_
		                        & "'"& year(now) &"/"& month(now) & "/"& day(now) &"', 0, "_
		                        & "'"& tidspunkt &"', "_
		                        & "'Timer indl�st', 0, 0, '00:00:00', '00:00:00', 1, 100, 781, "& extsysid &", '"& timerkom &"', '"& koregnr &"', '"& destination &"', "& usebopal &")"

		                
                                'Response.write strSQLKoins & "<br>"
                                'Response.flush
                                'Response.end

                                oConn.execute(strSQLKoins)

                                end if 'findesallerede = 0
    
                        end if

        'response.Cookies("monitor_job")(session("mid")) = jobid
        'response.Cookies("monitor_akt")(session("mid")) = aktid
          
        end if



end function



function sygdom120(lto)
    
       
        '** Indl�ser automatisk sygdom + 120 dages regel ***'
        dd = now
        dd_1 = dateadd("d", -1, dd)
        dd_121 = dateadd("m", -12, dd_1)

        dd_1_sql = year(dd_1) &"/"& month(dd_1) &"/"& day(dd_1)
        dd_121_sql = year(dd_121) &"/"& month(dd_121) &"/"& day(dd_121)


        dtInd = year(dd) &"/"& month(dd) &"/"& day(dd)
        tidspunkt = formatdatetime(now, 3)

      
                                '***** Er funktioen allerede k�rt idag
                                korantalSygdom = 1
                                strSQLsygdom = " SELECT timer FROM timer WHERE origin = 481 AND tfaktim = 20 AND tastedato = '"& dtInd &"'"
                                oRec10.open strSQLsygdom, oConn, 3
                                if not oRec10.EOF then
                                korantalSygdom = 0
                                end if
                                oRec10.close


              
          if cint(korantalSygdom) = 1 then


                                '*** Finder AKT data ***
                        
                                aktid = 0
                                aktnavn = "Activity NOT found"
                                fakturerbar = 0 
                                kid = 0
                                kkundenavn = "Customer NOT found"
                                jobnavn = "Job NOT found"
                                jobnr = "0"


                                strSQLakt = "SELECT a.id, a.navn, a.fakturerbar, k.kkundenavn, kid, jobnavn, jobnr FROM aktiviteter a "_
                                & "LEFT JOIN job j ON (j.id = a.job) "_
                                & "LEFT JOIN kunder k ON (k.kid = j.jobknr) WHERE a.fakturerbar = 20 AND aktstatus = 1"
                                oRec7.open strSQLakt, oConn, 3
                                if not oRec7.EOF then

                                aktid = oRec7("id")
                                aktnavn = oRec7("navn")
                                fakturerbar = oRec7("fakturerbar")
                                kid = oRec7("kid")
                                kkundenavn = oRec7("kkundenavn")
                                jobnavn = oRec7("jobnavn") 
                                jobnr = oRec7("jobnr")

                                end if
                                oRec7.close



                                strSQLM = "SELECT mid FROM medarbejdere WHERE mansat = 1"
                                oRec7.open strSQLM, oCOnn, 3
                                While not oRec7.EOF 


                                        call normtimerPer(oRec7("mid"), dd_1, 0, 0)
                                        timertilindlasning = ntimPer
                                        timertilindlasning = replace(timertilindlasning, ",", ".")

                                        call meStamdata(oRec7("mid"))



                                        erLastSygdom = 0
                                        '*** Er den sidste reg. en sygdom - indl�s da en mere
                                        strSQLsygdom = " SELECT timer, tfaktim FROM timer WHERE tmnr = "& oRec7("mid") &" ORDER BY tdato DESC"
                                        oRec10.open strSQLsygdom, oConn, 3
                                        if not oRec10.EOF then
                                        erLastSygdom = oRec10("tfaktim")
                                        end if
                                        oRec10.close
        
        
                                        'if session("mid") = 1 then
                                        'Response.write "<br>HER: "& meNavn &": erLastSygdom: " & erLastSygdom & " timertilindlasning: " & timertilindlasning
                                        'end if

 
                                        if cint(erLastSygdom) = 20 AND timertilindlasning > 0 then


                                        '*** Indl�ser Sygdom ***'
		                                strSQLKoins = "INSERT INTO timer "_
		                                &"("_
		                                &" timer, tfaktim, tdato, tmnavn, tmnr, tjobnavn, tjobnr, tknavn, tknr, "_
		                                &" TAktivitetId, taktivitetnavn, Taar, TimePris, TasteDato, fastpris, tidspunkt, "_
		                                &" editor, kostpris, seraft, sttid, sltid, valuta, kurs, origin "_
		                                &") "_
		                                &" VALUES "_
		                                &" (" _
		                                & timertilindlasning &", "& fakturerbar &", "_
		                                & "'"& dd_1_sql &"', "_
		                                & "'"& meNavn &"', "_
		                                & oRec7("mid") &", "_
		                                & "'"& jobnavn &"', "_
		                                & "'"& jobnr &"', "_
		                                & "'"& kkundenavn &"', "& kid &", "_
		                                & aktid &", "_
		                                & "'"& aktnavn &"', "_
		                                & year(now) &", 0, "_
		                                & "'"& year(now) &"/"& month(now) & "/"& day(now) &"', 0, "_
		                                & "'"& tidspunkt &"', "_
		                                & "'Timeout sygdom indl�sning', 0, 0, '00:00:00', '00:00:00', 1, 100, 481)"

		                
                                       'if session("mid") = 1 then
                                       'Response.write strSQLKoins & "<br>"
                                       'Response.flush
                                       'Response.end
                                       'end if

                                        oConn.execute(strSQLKoins)

                                        end if
        


                        
                                        '**** t�ller Sygdom *****'
                                        antalSygdom = 0
                                        strSQLsygdom = " SELECT count(tid) AS antal FROM timer WHERE tmnr = "& oRec7("mid") &" AND tfaktim = 20 AND tdato BETWEEN '"& dd_121_sql &"' AND '"& dd_1_sql &"' GROUP BY tmnr"
                                        oRec10.open strSQLsygdom, oConn, 3
                                        if not oRec10.EOF then
                                        antalSygdom = oRec10("antal")
                                        end if
                                        oRec10.close
                           
                                        'if session("mid") = 1 then
                                        'Response.write "antalSygdom: "& antalSygdom
                                        'end if
    
                                 '***** Oprettter Mail object ***

                                    'response.write request.servervariables("PATH_TRANSLATED")
                                    'response.flush
                                            sm = 1
                                            if sm = 1 AND antalSygdom > 1 then

			                                if (request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\sesaba.asp") then

 
			                                Set myMail=CreateObject("CDO.Message")
                                            myMail.Subject="TimeOut - Medarbejder "& meNavn &" n�rmer sig 120 sygedages regel"
                                            myMail.From = "timeout_no_reply@outzource.dk"
				      
                                    
			                                'Modtagerens navn og e-mail
			                                select case lto
			                                case "cflow" 

                                            myMail.To= "frode@cflow.no; rune@cflow.no"
                                            'myMail.To= "Support <support@outzource.dk>"
                                            'myMail.Cc= ""

                                            'if len(trim(meEmail)) <> 0 then
                                            'Mailer.AddCC ""& meNavn &"", ""& meEmail &""
                                            'myMail.Cc= ""& meNavn &"<"& meEmail &">"
                                            'end if
                                    
			                       
			                                case else
			                                'Mailer.AddRecipient "OutZourCE", "timeout_no_reply@outzource.dk"
                                            myMail.To= "support@outzource.dk"
                                            myMail.Cc= "jda@outzource.dk; sk@outzource.dk"
			                                end select
                        			
			                      
                        			
			                                        ' Selve teksten
					                                txtEmailHtml = "<br>Medarbejder "& meNavn &" har n�et "& antalSygdom &" sygedage d. "& weekdayname(weekday(dd_1, 1)) &" d. "& formatdatetime(dd_1, 2) 
                                                    txtEmailHtml = txtEmailHtml & "<br><br>Det er opn�et i perioden: "& dd_121 &" til "& dd_1 
					                                txtEmailHtml = txtEmailHtml & "<br><br>Du kan verificere timerne ved at logge ind i TimeOut <a href=""https://timeout.cloud/"&lto&""">her..</a><br><br><br>"_
					                                & "Med venlig hilsen<br>TimeOut Email Service<br>" 
					                        
					                      
                                                       myMail.htmlBody= "<html><head><title></title>"_
			                                           &"</head>"_
			                                           &"<body>"_ 
			                                           & txtEmailHtml & "</body></html>"

                                                   'myMail.AddAttachment "d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_14\inc\log\data\"& file

                                                    myMail.Configuration.Fields.Item _
                                                    ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
                                                    'Name or IP of remote SMTP server
                                   
                                           
                                                    smtpServer = "formrelay.rackhosting.com" 
                                           
                    
                                                    myMail.Configuration.Fields.Item _
                                                    ("http://schemas.microsoft.com/cdo/configuration/smtpserver")= smtpServer

                                                    'Server port
                                                    myMail.Configuration.Fields.Item _
                                                    ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25
                                                    myMail.Configuration.Fields.Update
                    
                                                    'if len(trim(meEmail)) <> 0 then
                                                    myMail.Send
                                                    'end if
                                                    set myMail=nothing


                        				
                        			
			                                end if ''** Mail path


                                            end if 'SM = 0


    

                                oRec7.movenext
                                wend
                                oRec7.close
  



                end if 'korantalSygdom = 0


                'if session("mid") = 1  then
                'Response.end
                'end if

end function


%>
