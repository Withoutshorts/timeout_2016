<% 
function indlaesEasyreg(lto)

                'Kun hvis admin logger på

                'Tjekker om de er en hverdag

                'Tjekker hvornår der senest er aautomatisk overført

                'Henter alle medarbejder emed easyretimer pr. dag <> 0
        
                'Henter alle easyreg. aktiviteter : med antal

                'Indlæser timer


                if session("rettigheder") = 1 then

            
                    
               
                '*** Seneste Easyregdato
                lastEasyregDato = "1/12/2017"
                strSQLeasyregSenest = "SELECT tdato FROM timer WHERE origin = 7901 ORDER BY tdato DESC LIMIT 1"
                oRec7.Open strSQLeasyregSenest, oConn, 3
                if not oRec7.EOF then
            
                lastEasyregDato = oRec7("tdato") 

                end if
                oRec7.close


                antalDageDiff = dateDiff("d", lastEasyregDato, now)

                
                '**** Antal Easyreg aktiviteter ****
                antalEasyreg = 0
                strSQLeasyregA = "SELECT COUNT(id) AS antal FROM aktiviteter WHERE easyreg = 1 AND aktstatus = 1 AND job <> 0 GROUP BY easyreg"
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


                'Response.Write "EA: "& C & " antalEasyreg: "& antalEasyreg &" antalDageDiff: "& antalDageDiff &" lastEasyregDato: "& lastEasyregDato &"<br>"
            
                if antalEasyreg > 0 then

                '**** Antal Easyreg aktiviteter ****
                strSQLeasyreg = "SELECT a.id As aktid, navn AS aktnavn, fakturerbar, jobnr, jobnavn, kkundenr, kkundenavn, fasttp, fastkp, fasttp_val FROM aktiviteter a"_
                & " LEFT JOIN job j ON (j.id = a.job) "_
                & " LEFT JOIN kunder k ON (kid = j.jobknr) "_
                &"  WHERE easyreg = 1 AND aktstatus = 1 AND a.job <> 0"

                'response.write strSQLeasyreg
                'response.flush

                oRec8.Open strSQLeasyreg, oConn, 3
                while not oRec8.EOF 
            
              
              

                        '*** Henter medarbejdere
                        strSQlmedarb = "SELECT mid, mnavn, measyregtimer FROM medarbejdere WHERE measyregtimer > 0 AND mansat = 1 ORDER BY mid"
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


                        for dageC = 0 TO antalDageDiff - 1
                
                        'Response.Write "EA: "& C & "<br>"

                        'lastEasyregDato = 
                        if dageC = 0 then
                        datothisBeregn = dateAdd("d", 1, lastEasyregDato)
                        else    
                        datothisBeregn = dateAdd("d", 1, datothisBeregn)
                        end if

                        datothis = FormatDateTime(datothisBeregn,2) 'day(datothisBeregn) &"/"& month(datothisBeregn) &"/"& year(datothisBeregn)
                        toDay = datepart("w", datothis, 2,2)

                            if toDay < 6 then 'Kun hverdage

                            
                
                                call opdaterTimer(aktid, aktnavn, tfaktimvalue, strFastpris, jobnr, strJobnavn, strJobknr, strJobknavn,_
                                medid, strMnavn, datothis, timerthis, kommthis, intTimepris,_
                                dblkostpris, offentlig, intServiceAft, strYear, sTtid, sLtid, visTimerelTid, stopur, intValuta, bopal, destination, dage, tildeliheledage, origin, extsysid, mtrx, intKpValuta)



                            end if

                        next

                
                        oRec7.movenext
                        wend
                        oRec7.close


                oRec8.movenext
                wend 
                oRec8.close



                end if 'antalEasyreg
    
                end if 'level

        


end function

%>