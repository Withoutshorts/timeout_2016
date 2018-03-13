<% 
function indlaesTillaeg(lto)

                'EPI2017 lufthavn
        

                '*** Indlæser tillæg fra AVI DK
               if ((session("mid") = 1 OR session("mid") = 2694 Or session("mid") = 2452 OR session("mid") = 32821 OR session("mid") = 2663) AND (lto = "epi2017" OR lto = "intranet - local")) then    

                '*EPI
                ddTjkSupplementSL = year(now) & "/" & month(now) & "/"& day(now) 

                ddTjkSupplementST = dateAdd("d", -45, now)
                ddTjkSupplementST = year(ddTjkSupplementST) & "/" & month(ddTjkSupplementST) & "/"& day(ddTjkSupplementST) 
                strSQLepiAViDKsupplement = "SELECT tid, tdato, sttid, sltid FROM timer WHERE taktivitetnavn = 'Data Collection' AND tdato BETWEEN '"& ddTjkSupplementST &"' AND '"& ddTjkSupplementSL &"' AND sttid <> '00:00:00' AND (origin = 11 OR origin = 12) AND overfort = 0"
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

                    '**** FØR 08:00
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


               
                    '**** Lør / Søn ELLER Hellig ****'
                   
                     call helligdage(oRec6("tdato"), 0, lto)
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

                'Response.write "Overført OK"

		        'Response.end

end function



function overtidsTillaeg(stempelurindstilling, lto, login, logud, mid, timertilindlasning, tfaktim)
    
        'Cflow
        'WAP?
        '** Indlæser ønsket beordret overarbejde ***'
    
        dtInd = year(now) &"/"& month(now) &"/"& day(now)
        'aktid = 183 '1798
        tidspunkt = formatdatetime(now, 3)

        '**** Overtid ********************
        if cdbl(timertilindlasning) > 0 then
        timertilindlasning = replace(timertilindlasning, ",", ".")


        '**** Tfaktim '**** 
        'if session("mid") = 1 then
        call normtimerPer(mid, login, 0, 0)

        call meStamdata(mid)

              

                                '*** Finder AKT data ***
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

                                '*** Indlæser Overtid til godkendelse ***'
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
		                        & "'Timeout tillæg indlæsning', 0, 0, '00:00:00', '00:00:00', 1, 100, 479)"

		                
                       
                        'Response.write strSQLKoins & "<br>"
                        'Response.flush
                        'Response.end

                        oConn.execute(strSQLKoins)
        
    
    
                         '***** Oprettter Mail object ***

                            'response.write request.servervariables("PATH_TRANSLATED")
                            'response.flush
                                    sm = 0
                                    if sm = 1 then

			                        if (request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\sesaba.asp") then

 
			                        Set myMail=CreateObject("CDO.Message")
                                    myMail.Subject="TimeOut - Medarbejder "& meNavn &" har ønsket reisetid/overtid"
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
					                        txtEmailHtml = "<br>Medarbejder "& meNavn &" har ønsket reisetid/overtid/tillæg d. "& weekdayname(weekday(login, 1)) &" d. "& formatdatetime(login, 2) &"<br><br>"
					                       
                                            if tfaktim = 51 OR tfaktim = 54 then 'OVertid
                                            txtEmailHtml = txtEmailHtml & "Der er ønsket: "& timertilindlasning &" timers overtid. <br>"
                                            end if

                                            if tfaktim = 52 OR tfaktim = 53 OR tfaktim = 55 then 'Reisetid
                                            txtEmailHtml = txtEmailHtml & "Der er ønsket: "& timertilindlasning &" timers reisetid. <br>"
                                            end if

                                            if tfaktim = 50 OR tfaktim = 60 OR tfaktim = 61 then 'Tillæg
                                            txtEmailHtml = txtEmailHtml & "Der er ønsket: "& timertilindlasning &" stk. tillæg. <br>"
                                            end if

                                            txtEmailHtml = txtEmailHtml & "Komme/Gå tid er: "& left(formatdatetime(login, 3), 5) & " - "& left(formatdatetime(logud, 3), 5) &"<br><br><br>"_ 
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
    

        timertilindlasning = 0

       

      

         stDatoSQL = year(now)&"/1/1"  
         'slDatoSQL = year(now)&"/12/31"  
        
         slDatoSQL = dateAdd("d", -1, now)
         slDatoSQL = year(slDatoSQL) &"/"& month(slDatoSQL) &"/"& day(slDatoSQL)

        '**Kun til DD. Norm bliver også kun beregnet til DD.
       

        select case lto
        case "wap"
        aktid = 34
        case "intranet - local"
        aktid = 6290
        case else
        aktid = 0
        end select


     


        '*** Nulstiller korrektion. **'
        'strSqlDel = "DELETE FROM timer WHERE tmnr = "& mid &" AND tfaktim = 31 AND tdato BETWEEN '"& stDatoSQL &"' AND '"& slDatoSQL &"'"
        'oConn.execute(strSqlDel)


        '**** Overtid ********************
        'if cdbl(timertilindlasning) > 0 then
       
        
        call akttyper2009(2)

        


       '*** Tjekker først MAX 10 timer pr Dag. TJEKKER ALLE DAGE **'
       '*** Nulstiller korrektion. MAX 10 timer pr. dag korrektion **'
       strSqlDel = "DELETE FROM timer WHERE tmnr = "& mid &" AND tfaktim = 114 AND origin = 580" 'SLET KUN dag til ag korrektioner
       oConn.execute(strSqlDel)

        '*** tjekker flex saldo ***
        realTimerFYmax10 = 0 
        strSQLtimer = "SELECT sum(timer) AS realTimerFY, tdato, tmnr FROM timer WHERE tmnr = "& mid &" AND tdato BETWEEN '"& stDatoSQL &"' AND '"& slDatoSQL &"' AND ("& aty_sql_realhours &") AND tfaktim <> 114 GROUP BY tmnr, tdato" 
        oRec8.open strSQLtimer, Oconn, 3
        while not oRec8.EOF 

           

            realTimerFYmax10 = oRec8("realTimerFY")

            if cdbl(realTimerFYmax10) > 10 then

            
                tDatothis = year(oRec8("tdato")) & "/" & month(oRec8("tdato")) & "/"& day(oRec8("tdato")) 
                dtInd = tDatothis

                flexStatusPrDag = (realTimerFYmax10 - 10)
                origin = 580
                call maxflexControlInsert(flexStatusPrDag, origin, mid, dtInd)

            end if

        oRec8.movenext
        wend 
        oRec8.close


    
    
    
    
    


        '***************************************************************************
         '*** tjekker GRaND flex saldo ***
        '***************************************************************************

          '** DD. ***'
        dtInd = year(now) &"/1/1" 'year(now) &"/"& month(now) &"/"& day(now)

        '** Norm FY
        stDatoNorm = "1/1/"& year(now)  
        slDatoNorm = dateAdd("d", -1, now) '"1/1/"& year(now)
        'slDatoNorm = dateAdd("d", 0, now)
        slDatoNormPrDdMinus1 = slDatoNorm
        dageDiff = dateDiff("d",stDatoNorm, slDatoNorm, 2,2) 
        'dageDiff = dateDiff("d",stDatoNorm, now, 2,2) 
        tidspunkt = formatdatetime(now, 3)

        call normtimerPer(mid, stDatoNorm, dageDiff, 0)


        realTimerFY = 0 
        strSQLtimer = "SELECT sum(timer) AS realTimerFY FROM timer WHERE tmnr = "& mid &" AND tdato BETWEEN '"& stDatoSQL &"' AND '"& slDatoSQL &"' AND ("& aty_sql_realhours &") AND (tfaktim <> 114) GROUP BY tmnr" 
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
        strSQLtimer = "SELECT sum(timer) AS realTimerFY FROM timer WHERE tmnr = "& mid &" AND tdato BETWEEN '"& stDatoSQL &"' AND '"& slDatoSQL &"' AND (tfaktim = 114 AND origin = 580) GROUP BY tmnr" 
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

        
        
        '*** Henter korrektion inden indtastning GT ALTID indlæst på 1 jan ***
        korrektionMaxFlexTimerFYforGT = 0 
        strSQLtimer = "SELECT sum(timer) AS realTimerFY FROM timer WHERE tmnr = "& mid &" AND tdato BETWEEN '"& stDatoSQL &"' AND '"& slDatoSQL &"' AND (tfaktim = 114 AND origin = 480) GROUP BY tmnr" 
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
        
        'flexStatus = ((realTimerFY+timertilindlasning) - ntimPer) 'Må ikke blive højere end 10
        flexStatus = ((realTimerFY+timertilindlasning+(korrektionMaxFlexTimerFYforprDag)+(korrektionMaxFlexTimerFYforGT)) - ntimPer) 'Må ikke blive højere end 10    

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

                   call maxflexControlInsert(flexmodregn, origin, mid, dtInd)
                
                end if

            'end if 'timertilindlæsrning

end function


function maxflexControlInsert(flexmodregn, origin, mid, dtInd)

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

                                '*** Indlæser Overtid til godkendelse ***'
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
		                        & "'Timeout tillæg indlæsning', 0, 0, '00:00:00', '00:00:00', 1, 100, "& origin &")"
		                
                                'if session("mid") = 1 then
                                'Response.write strSQLKoins & "<br>"
                                'Response.end
                                'end if

                                oConn.execute(strSQLKoins)


end function


function indlasTimertilUdbetaling(lto, mid, timertilindlasning)
    

     
        dtInd = year(now) &"/"& month(now) &"/"& day(now)
        tidspunkt = formatdatetime(now, 3)

        select case lto
        case "cflow"
        aktid = 184
        case "intranet - local"
        aktid = 6291
        case else
        aktid = 0
        end select


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

                                '*** Indlæser Overtid til godkendelse ***'
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
		                        & "'Timeout tillæg indlæsning', 0, 0, '00:00:00', '00:00:00', 1, 100, 481)"
		                
                       
                                'Response.write strSQLKoins & "<br>"
                                'Response.end

                                oConn.execute(strSQLKoins)
        
    
    
                   

            

            'end if 'timertilindlæsrning

end function





public timerthis_mtx
function matrixtimespan_2018(idag, mtrx, sTtid, sLtid, datoThis, lto)

                            useDate = idag '"01-01-" & year(now)
                            sTtid_org = sTtid


                            'Response.write "useDate: "& useDate & " " & sTtid
                            'Response.flush

                            call helligdage(datoThis, 0, lto)


                    
                            select case lto
                            case "cflow", "intranet - local"

                                        select case mtrx
                                        case 0 'Dag ordinær type 1
                                                lb = "07:30:00"
                                                ub = "15:30:00"
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

                                        case 6,7 'NAT / Morgen + Weekend nat (før mødetid)
                                                lb = "00:00:01"
                                                ub = "07:00:00"
                                        end select

                        
                             end select

                                        '***************************
                                        '*** START TID
                                        '***************************
                                        if (((mtrx = 0 OR mtrx = 1 OR mtrx = 2 OR mtrx = 6 OR mtrx = 8) AND (datepart("w", datoThis, 2,2) < 6) AND erHellig <> 1) OR _
                                        ((mtrx = 3 OR mtrx = 4 OR mtrx = 7) AND (datepart("w", datoThis, 2,2) = 6 OR datepart("w", datoThis, 2,2) = 7) AND erHellig <> 1) _
                                        OR (mtrx = 5 AND erHellig = 1)) then 'OR mtrx = 6

                                        

                                                        '*** Beregner timer og klokkeslet efter tarifgrænser *****
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
                                                            
                                                                   
                                                                    '6 7 Nat: Start tid vil altid være større en 00:00
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

                                                                                '** Hvis sluttid er før starttid på aften 
                                                                                '06:00 - 02:00 --> 12 timer på dag USE: UB :OK
                                                                                '15:00 - 02:00 --> 4 timer på Dag  USE: UB :OK
                                                                                '21:00 - 02:00 --> 0 timer på Dag  USE: UB :OK
                                                                                '06:00 - 07:00 --> 0 timer på Dag  USE: UB/sl :OK
                                                                                '05:00 - 06:00 --> 0 timer på Dag  USE: UB: OK
                                                                                '05:00 - 15:00 --> 8 timer på Dag  USE: sl :OK
                                                                                '09:00 - 15:00 --> 6 timer på Dag  USE: sl :OK
                                                                                '09:00 - 21:00 --> 10 timer på Dag USE: UB: OK
                                                                                '21:00 - 16:00 --> 9 timer på Dag USE: 



                                                                              
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
						                                    call timerogminutberegning(totalmin)
						                                    timerthis_mtx = thoursTot&"."&tminProcent 'tminTot
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

                                                        'response.write "<br> = mtrx: "& mtrx &" matrixAktFundet: "& matrixAktFundet & " timerthis_mtx: "& timerthis_mtx & " ub: "& ub &" lb:"& lb &" sTtid: "& sTtid & " sLtid: "& sLtid & " datepart: " & datepart("w", datoThis, 2,2) & "##"& cDate(useDate &" "& sTtid) &"<="& cDate(useDate &" "& sLtid)
                                                        'response.flush

                                     else        

                                        timerthis_mtx = 0

                                     end if 'mtrx 1-7 + weekend/helligdage   


                            
    
end function   






function indlasFerieSaldo(lto, mid, timertilindlasning, tfaktim, ferieaar)
    
        '*** Ferie optjent
       
        'dtInd = year(now) &"/"& month(now) &"/"& day(now)
	     dtInd = ferieaar &"-05-01"
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


                                '*** Indlæser Overtid til godkendelse ***'
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
		                        & "'Ferie optjent indlæst', 0, 0, '00:00:00', '00:00:00', 1, 100, 779)"

		                
                       
								'Response.write strSQLKoins & "<br>"
								'Response.flush
								'Response.end

								oConn.execute(strSQLKoins)
    
							 end if 'aktid = 0
          
        end if



end function



function indlasTimerTfaktimAktid(lto, mid, timertilindlasning, tfaktim, aktid)
    
      
        if aktid <> 0 then
        sqltfaktimaktidKri = "a.id = " & aktid & ""
        else
        sqltfaktimaktidKri = "a.fakturerbar = " & tfaktim & ""
        end if

        dtInd = year(now) &"/"& month(now) &"/"& day(now)
        tidspunkt = formatdatetime(now, 3)

        '**** Overtid ********************
        if cdbl(timertilindlasning) > 0 then
        timertilindlasning = replace(timertilindlasning, ",", ".")
    
        call meStamdata(mid)

              

                                '*** Finder AKT data ***
                                strSQLakt = "SELECT a.id, a.navn, a.fakturerbar, k.kkundenavn, kid, jobnavn, jobnr, j.id AS jid FROM aktiviteter a "_
                                & "LEFT JOIN job j ON (j.id = a.job) "_
                                & "LEFT JOIN kunder k ON (k.kid = j.jobknr) WHERE " & sqltfaktimaktidKri
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


                        if cdbl(aktid) <> 0 then

                                '*** Indlæser Overtid til godkendelse ***'
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
		                        & "'Timer indlæst', 0, 0, '00:00:00', '00:00:00', 1, 100, 781)"

		                
                       
                        'Response.write strSQLKoins & "<br>"
                        'Response.flush
                        'Response.end

                        oConn.execute(strSQLKoins)
    
                        end if

        'response.Cookies("monitor_job")(session("mid")) = jobid
        'response.Cookies("monitor_akt")(session("mid")) = aktid
          
        end if



end function



%>
