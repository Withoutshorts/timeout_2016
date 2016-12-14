
<%
    public fctimer, strAfgSQL, strAfgSQLtimer, feltTxtValFc
    public timerBrugt, timerTastet
    'public feltTxtValFc
    function ressourcefc_tjk(ibudgetaar, ibudgetmd, aar, md, usemrn, aktid, timerTastet)
    
                select case cint(ibudgetaar) 
                case 1 '** Afgræns indenfor budgetår

                    if ibudgetmd <> 1 then '7: 1 juli - 30 juni

                        'Beregn evt. days inmonth
                       
                        if md >= 7 then 'jul-dec sammeår
                        strAfgSQL = "AND ((aar = "& aar &" AND md >= "& ibudgetmd &") OR (aar = "& aar+1 &" AND md < "& ibudgetmd &"))"
                        strAfgSQLtimer = "AND (tdato BETWEEN '"& aar &"-"& ibudgetmd &"-01' AND '"& aar+1 &"-"& ibudgetmd-1 &"-30')"
                        else
                        strAfgSQL = "AND ((aar = "& aar-1 &" AND md >= "& ibudgetmd &") OR (aar = "& aar &" AND md < "& ibudgetmd &"))"
                        strAfgSQLtimer = "AND (tdato BETWEEN '"& aar-1 &"-"& ibudgetmd &"-01' AND '"& aar &"-"& ibudgetmd-1 &"-30')"
                        end if
                    
                      else

                        strAfgSQL = "AND (aar = "& aar &")"
                        strAfgSQLtimer = "AND (tdato BETWEEN '"& aar &"-01-01' AND '"& aar &"-12-31')"
                
                    end if

                case 2 'Afgræns indenfor MD

                        strAfgSQL = "AND (aar = "& aar &" AND md = "& ibudgetmd &")"
                        strAfgSQLtimer = "AND (tdato BETWEEN '"& aar &"-"& ibudgetmd &"-01' AND '"& aar+1 &"-"& ibudgetmd-1 &"-30')"
                        'ibudgetmd RET TIL

                case else

                    strAfgSQL = ""
                    strAfgSQLtimer = ""
        
                end select 

                feltTxtVal = 0
                
                '** Tjekker resouceforecast
                strSQLtimfc = "SELECT COALESCE(SUM(timer), 0) AS fctimer FROM ressourcer_md WHERE aktid = "& aktid & " AND medid = "& usemrn &" "& strAfgSQL
                
               'response.write strSQLtimfc
               'response.end


                fctimer = 0
                oRec5.open strSQLtimfc, oConn, 3
                while not oRec5.EOF
               
                fctimer = oRec5("fctimer")
                 
                oRec5.movenext
                wend
                oRec5.close


                
                timerBrugt = 0
               
        
                
                        '** Finder timeforbrug på AKTIVITET total på tværs
                        strSQLtimerReal = "SELECT COALESCE(SUM(timer), 0) AS timerbrugt FROM timer WHERE taktivitetid = "& aktid &" AND tmnr = "& usemrn &" "& strAfgSQLtimer &" GROUP BY taktivitetid"
                        t = 0

                        'response.write strSQLtimerReal
                        'response.end
                        
                        oRec5.open strSQLtimerReal, oConn, 3
                        while not oRec5.EOF
               
                        timerBrugt = oRec5("timerbrugt")
                 
                        t = 1
                        oRec5.movenext
                        wend
                        oRec5.close

        
               

               if len(trim(fctimer)) <> 0 AND fctimer <> 0 then
               fctimer = replace(fctimer, ".", ",") * 1
               else
               fctimer = 0
               end if

               
               if len(trim(timerBrugt)) <> 0 AND timerBrugt <> 0 then
               timerBrugt = replace(timerBrugt, ".", ",") * 1
               else
               timerBrugt = 0
               end if
               
               if len(trim(timerTastet)) <> 0 AND timerTastet <> 0 then
               timerTastet = replace(timerTastet, ".", ",") * 1
               else
               timerTastet = 0
               end if

               
              
               if timerTastet*1 > 0 then
               feltTxtValFc = (fctimer*1 - (timerTastet*1 + timerBrugt*1))
               else
               feltTxtValFc = (fctimer*1 - (timerBrugt*1))
               end if

                


      end function










    public forecastAktids, forecastAktidsTxt
    function allejobaktmedFC(viskunForecast, usemrn, jobid, risiko)

                    '*************************************************************************************************
                    '*** Henter kun akt med ressource timer for valgte medarb                      *******************
                    '*************************************************************************************************

                    'HENTER ALLE JOB OG AKT uanset dato afgrænsning. Alerts og luk linjer kommer først på på den nekelte linje.
                    '* Men linjerne bliver vis bare der findes et forecast 
                    '* == Diaglog hvis man kan se der er linjer der er ikke er mere forecast / timebudget på. 
                                            
                    '*** Henter kun aktiviteter med forecast på ****'
                    forecastAktids = " AND (a.id = 0"
                    forecastAktidsTxt = "#0#,"
                    if cint(viskunForecast) = 1 then
                                              
                        kforCastSQL = "SELECT aktid FROM ressourcer_md WHERE medid = " & usemrn & " AND jobid = "& jobid &" GROUP BY medid, aktid"

                        'response.write kforCastSQL
                        'response.flush

                        oRec5.open kforCastSQL, oConn, 3
                        while not oRec5.EOF 
                        forecastAktids = forecastAktids & " OR a.id = "&oRec5("aktid") 
                        forecastAktidsTxt = forecastAktidsTxt & ",#"& oRec5("aktid") &"#"
                        oRec5.movenext
                        wend
                        oRec5.close       

                                                
                        if cint(risiko) < 0 then '+ alle interne

                        kforCastSQL = "SELECT a.id AS aktid FROM aktiviteter AS a WHERE a.job = "& jobid &" GROUP BY id"
                        oRec5.open kforCastSQL, oConn, 3
                        while not oRec5.EOF 
                        forecastAktids = forecastAktids & " OR a.id = "&oRec5("aktid") 
                        forecastAktidsTxt = forecastAktidsTxt & ",#"& oRec5("aktid") &"#"
                        oRec5.movenext
                        wend
                        oRec5.close    


                        end if



                    end if

                    forecastAktids = forecastAktids & ")"

                    if cint(viskunForecast) = 1 then
                    forecastAktids = forecastAktids
                    else
                    forecastAktids = ""
                    end if

                                           

                    '*************************************************************************************************

    end function


%>