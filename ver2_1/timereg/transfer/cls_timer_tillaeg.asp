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



function overtidsTillaeg(stempelurindstilling, lto, login, logud, mid, timertilindlasning)
    
        'Cflow
        'WAP?
        '** Indlæser ønsket beordret overarbejde ***'
    
        dtInd = year(now) &"/"& month(now) &"/"& day(now)
        aktid = 183 '1798
        tidspunkt = formatdatetime(now, 3)

        '**** Overtid ********************
        if cdbl(timertilindlasning) > 0 then
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
		                        & "'Timeout tillæg indlæsning', 0, 0, '00:00:00', '00:00:00', 1, 100, 479)"

		                
                       
                        'Response.write strSQLKoins & "<br>"
                        'Response.flush
                        'Response.end

                        oConn.execute(strSQLKoins)
        
    
    
                         '***** Oprettter Mail object ***
			                        if (request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\sesaba.asp") then

 
			                        Set myMail=CreateObject("CDO.Message")
                                    myMail.Subject="TimeOut - Medarbejder "& meNavn &" har ønsket overtid"
                                    myMail.From = "timeout_no_reply@outzource.dk"
				      
                                    
			                        'Modtagerens navn og e-mail
			                        select case lto
			                        case "cflow" 

                                    myMail.To= "frode@cflow.no; rune@cflow.no; support@outzource.dk"
                                    'myMail.To= "support@outzource.dk"
                                    'myMail.Cc= ""

                                    'if len(trim(meEmail)) <> 0 then
                                    'Mailer.AddCC ""& meNavn &"", ""& meEmail &""
                                    'myMail.Cc= ""& meNavn &"<"& meEmail &">"
                                    'end if
                                    
			                       
			                        case else
			                        'Mailer.AddRecipient "OutZourCE", "timeout_no_reply@outzource.dk"
                                    myMail.To= "Support <support@outzource.dk>"
			                        end select
                        			
			                       

                                    timertilindlasningMail = replace(timertilindlasning, ".", ",")
                        			
			                                ' Selve teksten
					                        txtEmailHtml = "<br>Medarbejder "& meNavn &" har ønsket overtid d. "& weekdayname(weekday(login, 1)) &" d. "& formatdatetime(login, 2) &"<br><br>"_ 
					                        & "Der er ønsket: "& timertilindlasning &" timer. <br>"_
                                            & "Komme/Gå tid er: "& left(formatdatetime(login, 3), 5) & " - "& left(formatdatetime(logud, 3), 5) &"<br><br><br>"_ 
                                            & "Du kan godkende timerne ved at klikke <a href=""http://timeout.cloud/timeout_xp/wwwroot/ver2_14/to_2015/godkend_request_timer.asp?usemrn="&mid&"lto="&lto&""">her..</a><br><br><br>"_
					                        & "Med venlig hilsen<br>TimeOut Stempelur Service<br>" 
					                        
					                      
                                               myMail.htmlBody= "<html><head><title></title>"_
			                                   &"<LINK rel=""stylesheet"" type=""text/css"" href=""http://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/style/timeout_style_print.css""></head>"_
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



    
    
          
        end if



end function



function maksFlexControl(lto, mid, timertilindlasning)
    

        timertilindlasning = 0

         '** Nulstiller inde får FY
        stDatoNorm = "1/1/"& year(now)  
        dageDiff = dateDiff("d",stDatoNorm, now, 2,2) 
        tidspunkt = formatdatetime(now, 3)

        call normtimerPer(mid, stDatoNorm, dageDiff, 0)


        '** Indlæser ønsket beordret overarbejde ***'
        dtInd = year(now) &"/"& month(now) &"/"& day(now)

         stDatoSQL = year(now)&"/1/1"  
         slDatoSQL = year(now)&"/12/31"  
        '**Kun til DD. Norm bliver også kun beregnet til DD.
        'slDatoSQL = year(now)&"/"& month(now) &"/"& day(now)  


        select case lto
        case "wap"
        aktid = 34
        case "intranet - local"
        aktid = 6290
        case else
        aktid = 0
        end select


        '*** Nulstiller tillæg. F.eks hvis der slettes timer **'
        strSqlDel = "DELETE FROM timer WHERE tmnr = "& mid &" AND tfaktim = 31 AND tdato BETWEEN '"& stDatoSQL &"' AND '"& slDatoSQL &"'"
        oConn.execute(strSqlDel)


        '**** Overtid ********************
        'if cdbl(timertilindlasning) > 0 then
       
        
        call akttyper2009(2)

        '*** tjekker flex saldo ***
        realTimerFY = 0 
        strSQLtimer = "SELECT sum(timer) AS realTimerFY FROM timer WHERE tmnr = "& mid &" AND tdato BETWEEN '"& stDatoSQL &"' AND '"& slDatoSQL &"' AND ("& aty_sql_realhours &") GROUP BY tmnr" 
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

      
        
        flexStatus = ((realTimerFY+timertilindlasning) - ntimPer) 'Må ikke blive højere end 10
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

        'if session("mid") = 1 then
        'Response.write strSQLtimer & "<br>FLEX. bergn: "& realTimerFY &"+"& timertilindlasning & " - "&  ntimPer
        'Response.write "<br>flexmodregn:" & flexmodregn
        'Response.end
        'end if
        
       
        'Response.end
        

                if cdbl(flexmodregn) > 0 then
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
		                        & "'Timeout tillæg indlæsning', 0, 0, '00:00:00', '00:00:00', 1, 100, 480)"
		                
                       
                                'Response.write strSQLKoins & "<br>"
                                'Response.end

                                oConn.execute(strSQLKoins)
        
    
    
                   

            end if

            'end if 'timertilindlæsrning

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
%>