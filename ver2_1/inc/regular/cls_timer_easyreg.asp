<% 
function indlaesEasyreg(lto)

                'Kun hvis admin logger på

                'Tjekker om det er en hverdag

                'Tjekker hvornår der senest er aautomatisk overført

                'Henter alle medarbejder emed easyretimer pr. dag <> 0
        
                'Henter alle easyreg. aktiviteter : med antal

                'Indlæser timer


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

                'lastEasyregDato = "16/5/2018"

                antalDageDiff = dateDiff("d", lastEasyregDato, now)

                
                treeweeksminus = now
                treeweeksminus = dateadd("d", -21, treeweeksminus)
                treeweeksminus = year(treeweeksminus) &"/"& month(treeweeksminus) &"/"& day(treeweeksminus)

                '**** Antal Easyreg aktiviteter ****
                antalEasyreg = 0
                strSQLeasyregA = "SELECT COUNT(a.id) AS antal FROM job j "_
                &" LEFT JOIN aktiviteter a ON (a.job = j.id) WHERE easyreg = 1 AND aktstatus = 1 AND job <> 0 AND jobstatus = 1 AND jobstartdato >= '"& treeweeksminus &"' GROUP BY easyreg" 

                'if session("mid") = 21 then
                'response.Write strSQLeasyregA
                'response.flush
                'end if

                oRec7.Open strSQLeasyregA, oConn, 3
                if not oRec7.EOF then
            
                antalEasyreg = oRec7("antal") 

                end if
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


               

                '**** Antal Easyreg aktiviteter ****
                strSQLeasyreg = "SELECT a.id As aktid, navn AS aktnavn, fakturerbar, jobnr, jobnavn, kkundenr, kkundenavn, fasttp, fastkp, fasttp_val FROM aktiviteter a"_
                & " LEFT JOIN job j ON (j.id = a.job) "_
                & " LEFT JOIN kunder k ON (kid = j.jobknr) "_
                &"  WHERE easyreg = 1 AND aktstatus = 1 AND a.job <> 0 AND jobstartdato >= '"& treeweeksminus &"' AND jobstatus = 1"

                'if session("mid") = 21 then
                'response.write strSQLeasyreg
                'response.flush
                'end if


                oRec8.Open strSQLeasyreg, oConn, 3
                while not oRec8.EOF 
            
              
              

                        '*** Henter medarbejdere
                        strSQlmedarb = "SELECT mid, mnavn, measyregtimer FROM medarbejdere WHERE measyregtimer > 0 AND mansat = 1 ORDER BY mid" 'AND mid = 31
                        oRec7.open strSQlmedarb, oConn, 3
                        while not oRec7.EOF 
        
                        dageC = 0


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

                        if cdbl(timerthis) < "0,01" then 'tjekker ned på 6 decimaler
                        'timerthis = "0,0003"
                        timerthis = formatnumber(oRec7("measyregtimer")/antalEasyreg, 6)
                        end if
                        
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
                        strJobknr = oRec8("kkundenr")
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
                        
                       

                        toDay = datepart("w", datothis, 2,2)

                        'if session("mid") = 21 then
                        'Response.Write "datothis: " & datothis & " toDay: "& toDay &"<br>"
                        'end if

                            if toDay < 6 then 'Kun hverdage

                            
                
                                call opdaterTimer(aktid, aktnavn, tfaktimvalue, strFastpris, jobnr, strJobnavn, strJobknr, strJobknavn,_
                                medid, strMnavn, datothis, timerthis, kommthis, intTimepris,_
                                dblkostpris, offentlig, intServiceAft, strYear, sTtid, sLtid, visTimerelTid, stopur, intValuta, bopal, destination, dage, tildeliheledage, origin, extsysid, mtrx, intKpValuta)



                            end if

                        'end if 'pmok

                        next

                
                        oRec7.movenext
                        wend
                        oRec7.close


                a = a + 1
                oRec8.movenext
                wend 
                oRec8.close



                end if 'antalEasyreg
    
                end if 'level

        
                'if session("mid") = 21 then
                'Response.write "<br>Easyreg indlæst og klar! Antal: " & a
                'response.end
                'end if

end function

%>