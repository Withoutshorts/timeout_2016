
<%
    public resTimerThisM, restimeruspecThisM 
    function ressourcetimerTildelt(useDateStSQL,job_jid, job_aid, usemrn)

      'call aktBudgettjkOn_fn()
     'if session("mid") = 1 then
     ' Response.write "aktBudgettjkOn: "& aktBudgettjkOn & " aktBudgettjkOn_afgr "& aktBudgettjkOn_afgr
     'end if

                        resTimerThisM = 0
                        restimeruspecThisM = 0
                        SQLdatoKriTimer = ""
                        if cint(aktBudgettjkOn) = 1 then
                        '*******************************************
                        '*** Ressource timer tildelt             **'
                        '*******************************************
                                                                       
                                                                                    
                                if cint(aktBudgettjkOn_afgr) = 1 then

                                    '*** Afgræans indenfor regnskabsår // Starter dwet md 1 eller 7 (kontrolpanel)
                                    if month(useDateStSQL) < month(aktBudgettjkOnRegAarSt) then                                                                          
                                        useRgnArr =  dateAdd("yyyy", -1, useDateStSQL)
                                        useRgnArrNext = useDateStSQL
                                    else
                                        useRgnArr = useDateStSQL
                                        useRgnArrNext =  dateAdd("yyyy", 1, useDateStSQL)

                                    end if
                                                                     
                                    else

                                                                                                            
                                        useRgnArr = useDateStSQL
                                        useRgnArrNext =  dateAdd("yyyy", 1, useDateStSQL)
                                                                                    
                                    end if



                        select case cint(aktBudgettjkOn_afgr) 
                        case 1 'regnskabsår
                        sqlBudgafg = " AND ((md >= "& month(aktBudgettjkOnRegAarSt) & " AND aar = "& year(useRgnArr) & ") OR (md < "& month(aktBudgettjkOnRegAarSt) & " AND aar = "& year(useRgnArrNext) & "))" 
                        case 2 'måned
                        sqlBudgafg = " AND (md = "& month(aktBudgettjkOnRegAarSt) & " AND aar = "& year(useRgnArr) & ")" 
                        case else
                        sqlBudgafg = ""
                        end select

                        strSQLres = "SELECT IF(aktid <> 0, SUM(timer), 0) AS restimer, IF(aktid = 0, SUM(timer), 0) AS restimeruspec FROM ressourcer_md WHERE ((jobid = "& job_jid &" AND aktid = "& job_aid &" AND medid = "& usemrn &") OR (jobid = "& job_jid &" AND aktid = 0 AND medid = "& usemrn &")) "& sqlBudgafg &"  GROUP BY aktid, medid" 
                        
                        'if session("mid") = 1 then
                        'Response.write "<br>"& strSQLres & " aktBudgettjkOnRegAarSt = "& aktBudgettjkOnRegAarSt &" : useDateStSQL : "& useDateStSQL  &" // useDateSlSQL "& useDateSlSQL & "<br>"
                        'Response.end 
                        'end if


                        oRec9.open strSQLres, oConn, 3
                        while not oRec9.EOF 
                        resTimerThisM = resTimerThisM + oRec9("restimer")
                        restimeruspecThisM = restimeruspecThisM + oRec9("restimeruspec")
                        oRec9.movenext
                        wend
                        oRec9.close
                                                                        

                        end if


            end function







    public fctimer, strAfgSQL, strAfgSQLtimer, feltTxtValFc, fcBelob, fctimerTot, fcBelobTot
    public timerBrugt, timerTastet
    'public feltTxtValFc
    function ressourcefc_tjk(ibudgetaar, ibudgetmd, aar, md, usemrn, aktid, timerTastet)
    
        if lto = "wap" then
        ibudgetaar = 2
        end if

        select case md
        case 2

            select case aar
            case "2000","2004","2008","2012","2016","2020","2024","2028","2032","2036","2040","2044"
            daysinmth = 29
            case else
            daysinmth = 28
            end select    

        case 1,3,5,7,8,10,12
        daysinmth = 31
        case else
        daysinmth = 30
        end select

                select case cint(ibudgetaar) 
                case 1 '** Afgræns indenfor budgetår

                    if ibudgetmd <> 1 then '7: 1 juli - 30 juni

                        'Beregn evt. days inmonth
                       
                        if md >= 7 then 'jul-dec sammeår
                        strAfgSQL = "AND ((aar = "& aar &" AND md >= "& ibudgetmd &") OR (aar = "& aar+1 &" AND md < "& ibudgetmd &"))"
                        strAfgSQLtimer = "AND (tdato BETWEEN '"& aar &"-"& ibudgetmd &"-01' AND '"& aar+1 &"-"& ibudgetmd-1 &"-"& daysinmth &"')"
                        else
                        strAfgSQL = "AND ((aar = "& aar-1 &" AND md >= "& ibudgetmd &") OR (aar = "& aar &" AND md < "& ibudgetmd &"))"
                        strAfgSQLtimer = "AND (tdato BETWEEN '"& aar-1 &"-"& ibudgetmd &"-01' AND '"& aar &"-"& ibudgetmd-1 &"-"& daysinmth &"')"
                        end if
                    
                      else

                        strAfgSQL = "AND (aar = "& aar &")"
                        strAfgSQLtimer = "AND (tdato BETWEEN '"& aar &"-01-01' AND '"& aar &"-12-31')"
                
                    end if

                case 2 'Afgræns indenfor MD

                        strAfgSQL = "AND (aar = "& aar &" AND md = "& md &")"
                        strAfgSQLtimer = "AND (tdato BETWEEN '"& aar &"-"& md &"-01' AND '"& aar &"-"& md &"-"& daysinmth &"')"
                        'ibudgetmd RET TIL

                case else

                    strAfgSQL = ""
                    strAfgSQLtimer = ""
        
                end select 

                feltTxtVal = 0
                
                '** Tjekker resouceforecast
                if usemrn <> 0 then ' forecest pr akt og medarb
                strSQLtimfc = "SELECT COALESCE(SUM(timer), 0) AS fctimer FROM ressourcer_md WHERE aktid = "& aktid &" AND medid = "& usemrn &" "& strAfgSQL
                else 'forecast GT på akt samlet for alle medarb incl beløb
                strSQLtimfc = "SELECT COALESCE(SUM(r.timer), 0) AS fctimer, COALESCE(SUM(tp.6timepris*r.timer)) AS fcBelob FROM ressourcer_md AS r" _
                &" LEFT JOIN timepriser AS tp ON (tp.aktid = r.aktid AND tp.medarbid = r.medid) "_  
                &" WHERE r.aktid = "& aktid &" "& strAfgSQL & " GROUP BY r.aktid, r.medid"
                end if            


               'if session("mid") = 1 then
               'response.write strSQLtimfc & "<br>"
               'response.flush
               'end if        

                fctimerTot = 0
                fcBelobTot = 0 
                fcBelob = 0
                fctimer = 0
                oRec5.open strSQLtimfc, oConn, 3
                while not oRec5.EOF
               
                    fctimer = oRec5("fctimer")

                    if usemrn = 0 then ' forecest pr akt og medarb
                    fcBelob = oRec5("fcBelob")
                    end if
                 
                    fctimerTot = fctimerTot + oRec5("fctimer")
                
                    if usemrn = 0 then ' forecest pr akt og medarb
                    fcBelobTot = fcBelobTot + oRec5("fcBelob")
                    end if

                oRec5.movenext
                wend
                oRec5.close

             

                if fctimerTot <> 0 then
                fctimerTot = fctimerTot
                else
                fctimerTot = 0
                end if

                if fcBelobTot <> 0 then
                fcBelobTot = fcBelobTot
                else
                fcBelobTot = 0
                end if

                if fcBelob <> 0 then
                fcBelob = fcBelob
                else
                fcBelob = 0
                end if

                if fctimer <> 0 then
                fctimer = fctimer
                else
                fctimer = 0
                end if
     
             
                
                timerBrugt = 0
               
        
                
                        '** Finder timeforbrug på AKTIVITET total på tværs
                        strSQLtimerReal = "SELECT COALESCE(SUM(timer), 0) AS timerbrugt FROM timer WHERE taktivitetid = "& aktid &" AND tmnr = "& usemrn &" "& strAfgSQLtimer &" GROUP BY taktivitetid"
                        t = 0

                        'response.write strSQLtimerReal
                        'response.end

                           'if session("mid") = 1 then
                           'response.write strSQLtimerReal & "<br>"
                           'response.flush
                           'end if   
                        
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

                        '** Tilføj kun på åben job og åbne aktiviteter
                        'kforCastSQL = "SELECT aktid FROM ressourcer_md r"
                        'kforCastSQL = kforCastSQL & " LEFT JOIN job j ON (j.id = r.jobid)"
                        'kforCastSQL = kforCastSQL & " WHERE r.medid = " & usemrn & " AND r.jobid = "& jobid 
                        'kforCastSQL = kforCastSQL & " GROUP BY r.medid, r.aktid"

                        'if session("mid") = 1 then
                        'response.write kforCastSQL
                        'response.flush
                        'end if

                        oRec5.open kforCastSQL, oConn, 3
                        while not oRec5.EOF 
                        forecastAktids = forecastAktids & " OR a.id = "&oRec5("aktid") 
                        forecastAktidsTxt = forecastAktidsTxt & ",#"& oRec5("aktid") &"#"
                        oRec5.movenext
                        wend
                        oRec5.close       

                                                
                        if cint(risiko) < 0 OR (lto = "oko" AND cint(risiko) = -3) then '+ alle interne

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




     '*******************************************************************************************
        '**** Planlæg / Akt_bookings
        '*******************************************************************************************
        function addresbooking(medid, aktid, jobid, startDato, startDatoTid, slutDato, slutDatoTid, lto)
    	                
                        select case lto
                        case "essens", "intranet - local", "hidalgo", "outz", "jttek", "dencker"
                            show_planlag = 1
                        case else
                            show_planlag = 0
                        end select

                        if cint(show_planlag) = 1 then


                                        '**** VARIABLE ***'
                                        startDato = startDato
                                        startDatoTid = startDatoTid
                                        slutDato = slutDato
                                        slutDatoTid = slutDatoTid
                                        strEditor = session("user")
                                        strEditorDato = year(now) &"/"& month(now) &"/"& day(now) 

                                            aktnavn = ""
                                            strProjektgr1 = 1
						                    strProjektgr2 = 1
                                            strProjektgr3 = 1
                                            strProjektgr4 = 1
                                            strProjektgr5 = 1
                                            strProjektgr6 = 1
                                            strProjektgr7 = 1
                                            strProjektgr8 = 1
                                            strProjektgr9 = 1
                                            strProjektgr10 = 1

                                        '*** SLUT VAR ****


                                        strSQlaktprojgrp = "SELECT navn as aktnavn, projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, "_
						                &" projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM aktiviteter WHERE id = "& aktid
                                        
                                        oRec6.open strSQlaktprojgrp, oConn, 3
                                        if not oRec6.EOF then

                                            aktnavn = oRec6("aktnavn")
                                            strProjektgr1 = oRec6("projektgruppe1") 
						                    strProjektgr2 = oRec6("projektgruppe2") 
                                            strProjektgr3 = oRec6("projektgruppe3") 
                                            strProjektgr4 = oRec6("projektgruppe4") 
                                            strProjektgr5 = oRec6("projektgruppe5") 
                                            strProjektgr6 = oRec6("projektgruppe6") 
                                            strProjektgr7 = oRec6("projektgruppe7") 
                                            strProjektgr8 = oRec6("projektgruppe8") 
                                            strProjektgr9 = oRec6("projektgruppe9") 
                                            strProjektgr10 = oRec6("projektgruppe10")     
                                      

                                        end if
                                        oRec6.close

                                      
                                        ab_id = 0
                                        akt_boooking_findes = 0
                                        strSQLfindes_booking = "SELECT ab_id FROM akt_bookings WHERE ab_aktid = "& aktid
                                        
                                        'Response.write strSQLfindes_booking
                                        'Response.flush

                                        oRec6.open strSQLfindes_booking, oConn, 3
                                        if not oRec6.EOF then

                                        akt_boooking_findes = 1
                                        ab_id = oRec6("ab_id")

                                        end if
                                        oRec6.close

                                    if cint(akt_boooking_findes) = 1 then

                                    strSQLupd_akt_booking = "UPDATE akt_bookings "_
                                    &"SET ab_medid = "& medid &", ab_aktid = "& aktid &", "_
                                    &" ab_jobid = "& jobid &", ab_date = '"& startDato &"', ab_startdate  ='"& startDato &" "& startDatoTid &"', ab_enddate = '"& slutDato &" "& slutDatoTid &"', "_
                                    &" ab_serie = 0, ab_end_after = 0, ab_important = 0, ab_editor = '"& strEditor &"', ab_editor_date = '"& strEditorDato &"'"_
                                    & " WHERE ab_aktid = "& aktid
                                   

                                    'Response.write strSQLupd_akt_booking
                                    'Response.flush
                                    oConn.execute(strSQLupd_akt_booking)


                                    else

                                    strSQLins_akt_booking = "INSERT INTO akt_bookings (ab_name, ab_medid, ab_aktid, ab_jobid, ab_date, ab_startdate, ab_enddate, "_
                                    &" ab_serie, ab_end_after, ab_important, ab_editor, ab_editor_date)"_
                                    &" VALUES (0, "& session("mid") &", "& aktid &", "& jobid &", "_
                                    &"'"& startDato &"', "_
                                    &"'"& startDato &" "& startDatoTid &"', "_ 
						            &"'"& startDato &" "& slutDatoTid &"', "_
                                    &"0,0,0,"_
						            &"'"& strEditor &"', "_
						            &"'"& strEditorDato &"'"_
                                    &") "

                                   
                                    'Response.write strSQLins_akt_booking
                                    'Response.Flush

                                    oConn.execute(strSQLins_akt_booking)


                                    '*** Henter det netop oprettede akt-id ***
						            strSQLid = "SELECT ab_id FROM akt_bookings ORDER BY ab_id DESC"
						            oRec3.open strSQLid, oConn, 3
						            if not oRec3.EOF then
						            ab_id = oRec3("ab_id")
						            end if
						            oRec3.close
						
						            if len(ab_id) <> 0 then
						            ab_id = ab_id
						            else
						            ab_id = 0
						            end if


                                    end if 'akt_boooking_findes

                                    '** AKT HEADING
                                    '"UPDATE aktiviteter " +
                                    '"SET pl_header = " + model.Heading + " " +
                                    '"WHERE id = " + model.ActivityId + "; " +
                                    
                                 

                                    '* SKAL IKKE SLETTES Hver gang ved rediger - KUN NÅR DER VÆLGES
                                    if cint(akt_boooking_findes) = 1 then
                                    strSQLbookdel = "DELETE FROM akt_bookings_rel WHERE abl_bookid = " & aktid
                                    oConn.execute(strSQLbookdel)
                                    end if


                                     '** Ved opret tilføjes projekgrupper **'
                                    if cint(akt_boooking_findes) = 0 then

                                        multitildelmedarb = 0
                                        if cint(multitildelmedarb) = 1 then

                                        for p = 1 TO 10

                                            select case p
                                            case 1
                                                progrpID = strProjektgr1
                                            case 2
                                                progrpID = strProjektgr2
                                            case 3
                                                progrpID = strProjektgr3
                                            case 4
                                                progrpID = strProjektgr4
                                            case 5
                                                progrpID = strProjektgr5
                                            case 6
                                                progrpID = strProjektgr6
                                            case 7
                                                progrpID = strProjektgr7
                                            case 8
                                                progrpID = strProjektgr8
                                            case 9
                                                progrpID = strProjektgr9
                                            case 10
                                                progrpID = strProjektgr10
                                            end select
                                    
                                            strSQLmedarbiprogrp = "SELECT medarbejderid FROM progrupperelationer WHERE projektgruppeid = "& progrpID
                                            oRec6.open strSQLmedarbiprogrp, oConn, 3
                                            While not oRec6.EOF 
                            
                                               strSQLinsAkt_bookrel = "INSERT INTO akt_bookings_rel (abl_bookid, abl_medid) VALUES ("& ab_id &", "& oRec6("medarbejderid") &")"
                                                'Response.write strSQLinsAkt_bookrel & "<br>"
                                                'Response.flush
                                                oConn.execute(strSQLinsAkt_bookrel)


                                            oRec6.movenext
                                            wend 
                                            oRec6.close 

                                    
                                  
                                        next
                

                                     
                                            else 'Single tildel

                                                    select case lto
                                                    case "essens"
                                                    strMSQL = " mnavn = 'uspecificeret'"
                                                    case else
                                                    strMSQL = " mid = "& session("mid") &""
                                                    end select

                                                    strSQLm = "SELECT mid FROM medarbejdere WHERE "& strMSQL &" LIMIT 1"
                                                    oRec6.open strSQLm, oConn, 3   
                                                    if not oRec6.EOF then
                            
                                                     strSQLinsAkt_bookrel = "INSERT INTO akt_bookings_rel (abl_bookid, abl_medid) VALUES ("& ab_id &", "& oRec6("mid") &")"
                                                    'Response.write strSQLinsAkt_bookrel & "<br>"
                                                    'Response.flush
                                                    oConn.execute(strSQLinsAkt_bookrel)


                                                    end if
                                                    oRec6.close 

                                            

                                            end if 'multi

                                        end if 'findes

                        end if 'showplanlæg

    end function
%>