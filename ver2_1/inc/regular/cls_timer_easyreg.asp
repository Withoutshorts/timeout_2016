<% 
function indlaesEasyreg(lto)


               

                'Kun hvis admin logger p�

                'Tjekker om det er en hverdag

                'Tjekker hvorn�r der senest er aautomatisk overf�rt

                'Henter alle medarbejder emed easyretimer pr. dag <> 0
        
                'Henter alle easyreg. aktiviteter : med antal

                'Indl�ser timer


                if session("rettigheder") = 1 then

                a = 0
                    
               
                '*** Seneste Easyregdato
                lastEasyregDato = "1/12/2017"
                strSQLeasyregSenest = "SELECT tdato FROM timer WHERE origin = 7901 ORDER BY tdato DESC LIMIT 1"
                oRec7.Open strSQLeasyregSenest, oConn, 3
                if not oRec7.EOF then
            
                lastEasyregDato = oRec7("tdato") 

                end if
                oRec7.close

                'lastEasyregDato = "24/6/2018"

                antalDageDiff = dateDiff("d", lastEasyregDato, now)

               

                '*** SKAL Hentes fra kontrol panel
           
                treeweeksminus = now
                treeweeksminus = dateadd("d", -showeasyreg_per, treeweeksminus)
                treeweeksminus = year(treeweeksminus) &"/"& month(treeweeksminus) &"/"& day(treeweeksminus)
        

                if cint(showeasyreg_per) <> 0 then
                strSQLtreeweeksminus =  " AND jobstartdato >= '"& treeweeksminus &"'"
                else
                strSQLtreeweeksminus =  ""
                end if

                '**** Antal Easyreg aktiviteter ****
                antalEasyreg = 0
                strSQLeasyregA = "SELECT a.id AS antal, COALESCE(Sum(timer)) AS sumtimer, easyreg_max FROM job j "_
                &" LEFT JOIN aktiviteter a ON (a.job = j.id) "_
                &" LEFT JOIN timer t ON (t.taktivitetid = a.id) "_
                &" WHERE easyreg = 1 AND aktstatus = 1 AND job <> 0 AND jobstatus = 1 "& strSQLtreeweeksminus &" GROUP BY taktivitetid" 
                

                'if session("mid") = 21 then
                'response.Write strSQLeasyregA
                'response.flush
                'end if

                'Response.write "HER: " &  lastEasyregDato & " antalDageDiff: " & antalDageDiff
                'Response.end

                oRec7.Open strSQLeasyregA, oConn, 3
                while not oRec7.EOF '

                if ISNULL(oRec7("sumtimer")) <> true then
            
                    if cdbl(oRec7("sumtimer")) < cdbl(oRec7("easyreg_max")) then 
                    antalEasyreg = antalEasyreg + 1 
                    end if            

                else
                    antalEasyreg = antalEasyreg + 1 
                end if

                oRec7.movenext
                wend
                oRec7.close

                if antalEasyreg <> 0 then
                antalEasyreg = antalEasyreg
                else
                antalEasyreg = 0
                end if

                'if session("mid") = 21 then
                'Response.Write "EA: "& C & " antalEasyreg: "& antalEasyreg &" antalDageDiff: "& antalDageDiff &" lastEasyregDato: "& lastEasyregDato &"<br>"
                'end if        

                
    
    
                if antalEasyreg > 0 then

                    select case datepart("w", now, 2,2) 
                    case 1,3,5
                    sqlAktOrdeBy = "DESC"
                    case else
                    sqlAktOrdeBy = ""
                    end select
              
                '**** Aktive Easyreg aktiviteter ****
                strSQLeasyreg = "SELECT a.id As aktid, navn AS aktnavn, fakturerbar, jobnr, jobnavn, kkundenr, kid, "_
                & " kkundenavn, fasttp, fastkp, fasttp_val, easyreg_max, easyreg_timer_proc, SUM(timer) AS sumtimer FROM aktiviteter a"_
                & " LEFT JOIN job j ON (j.id = a.job) "_
                & " LEFT JOIN kunder k ON (kid = j.jobknr) "_
                & " LEFT JOIN timer t ON (taktivitetid = a.id) "_
                &"  WHERE easyreg = 1 AND aktstatus = 1 AND a.job <> 0 "& strSQLtreeweeksminus &" AND jobstatus = 1 GROUP BY taktivitetid ORDER BY taktivitetid "& sqlAktOrdeBy

                'if session("mid") = 21 then
                'response.write strSQLeasyreg & "<br>antalEasyreg: " & antalEasyreg
                'response.end
                'end if


                oRec8.Open strSQLeasyreg, oConn, 3
                while not oRec8.EOF 
            

                        '** MAX Easyreg timer tjk ***'
                        if ISNULL(oRec8("sumtimer")) <> true then
                        timerThisAkt = oRec8("sumtimer")
                        else
                        timerThisAkt = 0
                        end if

                        if ISNULL(oRec8("easyreg_max")) <> true then
                        easyreg_max = oRec8("easyreg_max")
                        else
                        easyreg_max = 0
                        end if

                        
                        
                    
                        if cdbl(easyreg_max) < cdbl(timerThisAkt) then


                        '*** Henter medarbejdere
                        measyregtimer = 0
                        strSQlmedarb = "SELECT mid, mnavn, measyregtimer FROM medarbejdere WHERE measyregtimer > 0 AND mansat = 1 ORDER BY mid" 'AND mid = 31 AND mid = 7 
                        oRec7.open strSQlmedarb, oConn, 3
                        while not oRec7.EOF 
        
                        dageC = 0

                        'Felter til TIMER tabel
                        origin = 7901
                        extsysid = 0
                        mtrx = 0
                        intValuta = oRec8("fasttp_val")
                        intKpValuta = intValuta            
                        kommthis = ""
                        tildeliheledage = 0
                        dage = 0
                        bopal = 0
                        destination = 0
                        stopur = 0
                        visTimerelTid = 0
                        sTtid = "00:00:00" 
                        sLtid = "00:00:00"
                        strYear = year(toDay)
                        intServiceAft = 0
                        offentlig = 0
    
                        dblkostpris = oRec8("fastkp")
                        intTimepris = oRec8("fasttp")
                    
                        kommthis = ""

                       
                        timerthis = formatnumber(oRec7("measyregtimer")/antalEasyreg, 2)
                        
                        measyregtimer = oRec7("measyregtimer")

                        'if cdbl(timerthis) < "0,01" then 'tjekker ned p� 6 decimaler
                        
                        'timerthis = formatnumber(oRec7("measyregtimer")/antalEasyreg, 6)
                        'end if

                      
                        timerthis = replace(timerthis, ".", "")
                        timerthis = replace(timerthis, ",", ".")          
                        
                        strMnavn = oRec7("mnavn")
                        medid = oRec7("mid")
    
                        aktid = oRec8("aktid") 
                        aktnavn = oRec8("aktnavn") 
                        tfaktimvalue = oRec8("fakturerbar") 
                        strFastpris = 0
                        jobnr = oRec8("jobnr")
                        strJobnavn = replace(oRec8("jobnavn"), "'", "")
                        strJobknr = oRec8("kid")
                        strJobknavn = replace(oRec8("kkundenavn"), "'", "")

                        'pmok = 0
                        'if session("mid") = 21 then 'AND instr(strMnavn, "Kristensen") <> 0
                        'Response.Write "jobnr: "& jobnr &" jobnavn: "& strJobnavn &" aktnavn: "& aktnavn &" mnavn: "& strMnavn &" timerthis: " & timerthis & "<br>"
                        'response.flush
                        'pmok = 0 '1
                        'end if

                        'if session("mid") = 21 then
                        'Response.Write "datothis: " & datothis & " antalDageDiff: "& antalDageDiff &"<br>"
                        'end if


                        for dageC = 0 TO antalDageDiff - 1

                        'if pmok = 1 then
                
                        'Response.Write "EA: "& C & "<br>"

                        'lastEasyregDato = 
                        if dageC = 0 then
                        datothisBeregn = dateAdd("d", 1, lastEasyregDato)
                        else    
                        datothisBeregn = dateAdd("d", 1, datothisBeregn)
                        end if

                        datothis = FormatDateTime(datothisBeregn,2) 'day(datothisBeregn) &"/"& month(datothisBeregn) &"/"& year(datothisBeregn)
                        
                        'Dobbeltjk timer p� dagen pr. emdarb.
                        medTimerPrDag = 0
                        sqlMtimerDato = year(datothisBeregn) & "/" & month(datothisBeregn) & "/" & day(datothisBeregn)
                        strSQlmedarbTimerprDag = "SELECT SUM(timer) AS medtimer FROM timer WHERE tdato = '" & sqlMtimerDato & "' AND tmnr = "& oRec7("mid") &" AND origin = 7901 AND taktivitetnavn = 'Easyreg' GROUP BY tmnr"

                        'Response.Write strSQlmedarbTimerprDag
                        'Response.Flush

                        oRec9.Open strSQlmedarbTimerprDag, oConn, 3
                        if not oRec9.EOF then

                         '** MAX MEdarb. timer pr. dag ***'
                        if ISNULL(oRec9("medtimer")) <> true then
                        medTimerPrDag = oRec9("medtimer")
                        else
                        medTimerPrDag = 0
                        end if
                           

                        end if
                        oRec9.close


                        toDay = datepart("w", datothis, 2,2)

                        'if session("mid") = 21 then
                        'Response.Write "datothis: " & datothis & " toDay: "& toDay &" strJobknr: "& strJobknr &"<br>"
                        'end if

                            if toDay < 6 AND (cdbl(medTimerPrDag) < cdbl(measyregtimer)) then 'Kun hverdage

                            
                
                                call opdaterTimer(aktid, aktnavn, tfaktimvalue, strFastpris, jobnr, strJobnavn, strJobknr, strJobknavn,_
                                medid, strMnavn, datothis, timerthis, kommthis, intTimepris,_
                                dblkostpris, offentlig, intServiceAft, strYear, sTtid, sLtid, visTimerelTid, stopur, intValuta, bopal, destination, dage, tildeliheledage, origin, extsysid, mtrx, intKpValuta, 0)



                            end if

                        'end if 'pmok

                        next

                
                        oRec7.movenext
                        wend
                        oRec7.close

                        end if


                a = a + 1
                oRec8.movenext
                wend 
                oRec8.close



                end if 'antalEasyreg
    
                end if 'level

        
                'if session("mid") = 21 then
                'Response.write "<br>Easyreg indl�st og klar! Antal: " & a
                'response.end
                'end if

end function

%>