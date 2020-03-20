
<%
    thisfile = "timetag_mobile" 
%>

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->

<%'**** Søgekriterier AJAX **'

        'section for ajax calls
        if Request.Form("AjaxUpdateField") = "true" then
        Select Case Request.Form("control")
            case "CheckHoursInDietPeriod"

                strStdato = request("strStdatoDiet")
                strSldato = request("strSldatoDiet")

                'response.Write " strStdato " & strStdato & " strSldato " & strSldato
                'response.End

                call akttyper2009(2)
                call traveldietexp_fn()

                antalDagArr = dateDiff("d", strStdato, strSldato, 2,2)
                timerForbrugprDagOverskreddet = 0

                for d = 0 TO antalDagArr

                if d = 0 then
                tjkThisDatp = year(strStdato) & "-" & month(strStdato) & "-" & day(strStdato)
                else
                tjkThisDatp = dateAdd("d", 1, tjkThisDatp)
                tjkThisDatp = year(tjkThisDatp) & "-" & month(tjkThisDatp) & "-" & day(tjkThisDatp)
                end if 

                strMedid = session("mid")

                timerBrugt = 0
                strSQLtimreal = "SELECT COALESCE(SUM(timer), 0) AS timerbrugt FROM timer WHERE tmnr = "& strMedid &" AND tdato = '"& tjkThisDatp &"' AND ("& aty_sql_realhours &") GROUP BY tdato, tmnr"
        
                'response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;antalDagArr: "& antalDagArr &" strSQLtimreal: " & strSQLtimreal & "<br>"
                      
                oRec5.open strSQLtimreal, oConn, 3
                while not oRec5.EOF
               
                timerBrugt = oRec5("timerbrugt")
        
                if cdbl(timerBrugt) > cdbl(traveldietexp_maxhours) then
                timerForbrugprDagOverskreddet = 1
                end if         
        
                oRec5.movenext
                wend
                oRec5.close

                next      

                response.Write timerForbrugprDagOverskreddet

            case "GetKundeNavn"
                jobid = request("jobid")
        
                kundetxt = ""
                strSQl = "SELECT jobknr, kkundenavn, kkundenr FROM job LEFT JOIN kunder ON (kid = jobknr) WHERE id = "& jobid
                oRec.open strSQl, oConn, 3
                if not oRec.EOF then
                    kundetxt = oRec("kkundenavn") &" ("& oRec("kkundenr") & ")"
                end if
                oRec.close

                call jq_format(kundetxt)
                kundetxt = jq_formatTxt

                response.Write kundetxt

            case "KmReg"

                medid = session("mid")
                antal = request("antalkm")
                destination = request("distination")
                koregnr = request("koregnr")
                kmdato = request("kmdato")
                if kmdato <> "" then
                    kmdato = year(kmdato) &"-"& month(kmdato) &"-"& day(kmdato)
                else
                    kmdato = year(now) &"-"& month(now) &"-"& day(now)
                end if

                if len(trim(request("usebopal"))) <> 0 then
                usebopal = request("usebopal")
                else
                usebopal = 0
                end if

                aktid = 0
                select case lto
                    case "outz"
                        aktid = 3046
                    case "lm"
                        aktid = 58468
                end select

                jobid = request("jobid")
                strkundenavn = ""
                strSQL = "SELECT kkundenavn FROM job LEFT JOIN kunder ON (jobknr = kid) WHERE id = "& jobid
                oRec.open strSQL, oConn, 3
                if not oRec.EOF then
                    strkundenavn = oRec("kkundenavn")
                end if
                oRec.close
                
                if strkundenavn <> "" then
                    timerkom = "Kunde: " & strkundenavn
                end if

                if aktid <> 0 then
                    call indlasTimerTfaktimAktid(lto, medid, antal, 0, aktid, 0, kmdato, 0, timerkom, koregnr, destination, usebopal)
                end if

                'response.Write "Done " & aktid

            case "GetExpenceId"
                    
                databaseName = ""
                select case lto
                    case "outz"
                        databaseName = "timeout_intranet"    
                    case else
                        databaseName = "timeout_" & lto
                end select

                nextId = 1
                lastid = 0
                strSQL = "SELECT AUTO_INCREMENT FROM information_schema.tables WHERE table_name = 'materiale_forbrug' AND table_schema = '"& databaseName &"'"
                oRec.open strSQL, oConn, 3 
                if not oRec.EOF then
                    nextid = oRec("AUTO_INCREMENT")
                    nextid = cint(nextid) - 1
                end if
                oRec.close

                'nextId = lastid + 1

                response.Write nextid

            case "GetCurrencies"
                strOpt = ""
                strSQLValuta = "SELECT id, valutakode FROM valutaer"
                oRec.open strSQLValuta, oConn, 3
                while not oRec.EOF
                strOpt = strOpt & "<option value='"& oRec("id") &"' data-valutakode='"& oRec("valutakode") &"' >"& oRec("valutakode") &"</option>"   
                oRec.movenext
                wend
                oRec.close
     
                response.Write strOpt

            case "GetExpenecTypes"
                strOpt = ""
                select case lto
                case "nt", "lm"
                strSQLMatgrpSQLkri = " WHERE matgrp_konto <> 0"
                case else
                strSQLMatgrpSQLkri = ""
                end select

               select case lto
                    case "lm" 
                        matidKri = " AND id <> 9" ' 9 = "Andet" og "Andet" skal sta til sidst
                    case else
                        matidKri = ""
                end select

                strSQLMatgrp = "SELECT id, navn FROM materiale_grp " & strSQLMatgrpSQLkri & matidKri & " ORDER BY navn"
                oRec.open strSQLMatgrp, oConn, 3
                while not oRec.EOF
                    strOpt = strOpt & "<option value='"& oRec("id") &"'>"& oRec("navn") &"</option>"
                oRec.movenext
                wend
                oRec.close

                if lto = "lm" then
                    strOpt = strOpt & "<option value='9'>"& expence_txt_092 &"</option>" 'Andet gruppen skal l;gge til sidst
                end if

                call jq_format(strOpt)
                strOpt = jq_formatTxt
        
                response.Write strOpt
        
            case "FN_sogjobogkunde"



                'jq_lto = request("jq_lto")
                

                '*** SØG kunde & Job            
                
                if len(trim(request("jq_newfilterval"))) <> 0 then
                filterVal = 1 
                jobkundesog = request("jq_newfilterval")

                else
                filterVal = 0
                jobkundesog = "6xxxxxfsdf554"
                end if
        
                medid = request("jq_medid")

                'if medid = 1 AND lto = "hestia" then
                '    medid = 164
                'end if

                if len(trim(request("thisJobid"))) <> 0 then
                jq_jobid = request("thisJobid")
                else
                jq_jobid = 0
                end if
                        

                        call positiv_aktivering_akt_fn()
                        '** PA settings // Overrule
                        if instr(lto, "epi") <> 0 then
                            pa_aktlist = 1
                        else
                            pa_aktlist = pa_aktlist
                        end if

                        if cint(pa_aktlist) = 0 then 'PA = 0 kan søge i jobbanken / PA = 1 kan kun søge på aktivjobliste
                        strSQLPAkri =  ""
                        else
                        strSQLPAkri =  " AND tu.forvalgt = 1" 
                        end if


                varTjDatoUS_man = request("varTjDatoUS_man")
                varTjDatoUS_son = dateAdd("d", 6, varTjDatoUS_man)

                varTjDatoUS_man = year(varTjDatoUS_man) &"/"& month(varTjDatoUS_man) &"/"& day(varTjDatoUS_man)
                varTjDatoUS_son = year(varTjDatoUS_son) &"/"& month(varTjDatoUS_son) &"/"& day(varTjDatoUS_son)

            
                '*** Datospærring Vis først job når stdato er oprindet
                call lukaktvdato_fn()
                ignJobogAktper = lukaktvdato

                select case lto
                case "nt"

                    strSQLDatokri = ""

                case else

                    select case ignJobogAktper
                    case 0,1
                    strSQLDatokri = " AND ((j.jobstartdato <= '"& varTjDatoUS_son &"' AND j.jobstatus = 1) OR (j.jobstatus = 3))"
                    case 3
                    strSQLDatokri = " AND ((j.jobstartdato <= '"& varTjDatoUS_son &"' AND j.jobslutdato >= '"& varTjDatoUS_man &"' AND j.jobstatus = 1) OR (j.jobstatus = 3))"
                    case else
                    strSQLDatokri = ""
                    end select

                end select


                if filterVal <> 0 then
            
                 if jobkundesog <> "-1" then

                    if jobkundesog = "all" then 
                        strSQLSogKri = " AND (jobnr <> '' AND kkundenavn <> '') "
                    else
                        strSQLSogKri = " AND (jobnr LIKE '"& jobkundesog &"%' OR jobnavn LIKE '%"& jobkundesog &"%' OR "_
                        &" kkundenavn LIKE '"& jobkundesog &"%' OR kkundenr = '"& jobkundesog &"' OR k.kinit = '"& jobkundesog &"') AND kkundenavn <> ''"
                    end if            

                 lmt = 250
                 else
                 strSQLSogKri = ""
                 lmt = 250
                 end if            


                lastKid = 0
                
                if jobkundesog <> "-1" then
                strJobogKunderTxt = strJobogKunderTxt &"<span style=""color:red; font-size:9px; float:right;"" id=""luk_jobsog"">[X]</span>"    
                strJobogKunderTxt = strJobogKunderTxt &"<ul>"
                end if

            
                select case lto
                    case "lm"
                        strSQLjobidKri = " AND (j.id <> 3)"
                    case else
                        strSQLjobidKri = ""
                end select
                

                select case lto
                case "nt"

                    strSQL = "SELECT j.id AS jid, j.jobnavn, j.jobnr, j.jobstatus, k.kkundenavn, k.kkundenr, k.kid, j.jobstartdato FROM job AS j "_
                    &" LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "_
                    &" WHERE (j.jobstatus = 1 OR j.jobstatus = 3) "& strSQLPAkri & " AND j.id = 12839" 

                    strSQL = strSQL & strSQLSogKri
                    strSQL = strSQL &" GROUP BY j.id ORDER BY kkundenavn, jobnavn LIMIT "& lmt     
                    
                 case "xlm"

                    strSQL = "SELECT j.id AS jid, j.jobnavn, j.jobnr, j.jobstatus, k.kkundenavn, k.kkundenr, k.kid, j.jobstartdato FROM job AS j "_
                    &" LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "_
                    &" WHERE (j.jobstatus = 1 OR j.jobstatus = 3) "& strSQLPAkri & " AND j.id = 6730" 

                    strSQL = strSQL & strSQLSogKri
                    strSQL = strSQL &" GROUP BY j.id ORDER BY kkundenavn, jobnavn LIMIT "& lmt 

                case else

                    strSQL = "SELECT j.id AS jid, j.jobnavn, j.jobnr, j.jobstatus, k.kkundenavn, k.kkundenr, k.kid, j.jobstartdato FROM timereg_usejob AS tu "_ 
                    &" LEFT JOIN job AS j ON (j.id = tu.jobid) "_
                    &" LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "_
                    &" WHERE tu.medarb = "& medid &" AND (j.jobstatus = 1 OR j.jobstatus = 3) "& strSQLPAkri &" "& strSQLDatokri & strSQLjobidKri
            
                    strSQL = strSQL & strSQLSogKri

                    if lto = "hestia" then 
                    strSQL = strSQL &" GROUP BY j.id ORDER BY jobnavn LIMIT "& lmt     
                    else
                    strSQL = strSQL &" GROUP BY j.id ORDER BY kkundenavn, jobnavn LIMIT "& lmt     
                    end if
              

                end select

                'if lto = "nt" then
                'response.write "<option>"& strSQL &"</option>"
                'response.flush 
                'end if

                'if lto <> "tbg" then '1 job forvalgt i DD

                    if (jobkundesog = "-1") then 'kunde job sog DD

                        if cdbl(jq_jobid) = 0 then
                        jobSEL0 = "SELECTED"
                        else
                        jobSEL0 = ""
                        end if

                        strJobogKunderTxt = strJobogKunderTxt & "<option value=""-1"" "& jobSEL0 &">:</option>"
                    end if   

                'end if
                           

                c = 0
                oRec.open strSQL, oConn, 3
                while not oRec.EOF
        
                if lastKnavn <> oRec("kkundenavn") AND (lto <> "hestia" AND lto <> "intranet - local")  then
                
                        if jobkundesog <> "-1" then 'kunde job sog DD
                            
                            if c <> 0 then
                            strJobogKunderTxt = strJobogKunderTxt &"<br><br>"
                            end if            
                            strJobogKunderTxt = strJobogKunderTxt &"<li class=""span_job""><b>"& oRec("kkundenavn") &" "& oRec("kkundenr") &"</b></li>"
        
                        else

                            if lto <> "tbg" then '1 job forvalgt i DD
                            strJobogKunderTxt = strJobogKunderTxt & "<option DISABLED>"& left(oRec("kkundenavn"), 20) & "</option>"
                            end if

                        end if
                
                end if 


                    if jobkundesog <> "-1" then 'kunde job sog DD
                       
                        if lto = "hestia" OR lto = "intranet - local" then
                        strJobogKunderTxt = strJobogKunderTxt & "<br><li class=""span_job"" id=""chbox_job_"& oRec("jid") &"""><input type=""hidden"" id=""hiddn_job_"& oRec("jid") &""" value="""& oRec("jobnavn") & " ("& oRec("jobnr") &")"">"
                        strJobogKunderTxt = strJobogKunderTxt & "<input type=""hidden"" id=""hiddn_jid_"& oRec("jid") &""" value="& oRec("jid") &"><b>"& oRec("jobnavn") & "</b> ("& oRec("jobnr") &")"
                        strJobogKunderTxt = strJobogKunderTxt &" ..... "& oRec("jobstartdato") & ""
                        else
                        strJobogKunderTxt = strJobogKunderTxt & "<li class=""span_job"" id=""chbox_job_"& oRec("jid") &"""><input type=""hidden"" id=""hiddn_job_"& oRec("jid") &""" value="""& oRec("jobnavn") & " ("& oRec("jobnr") &")"">"
                        strJobogKunderTxt = strJobogKunderTxt & "<input type=""hidden"" id=""hiddn_jid_"& oRec("jid") &""" value="& oRec("jid") &">"& left(oRec("jobnavn"), 30) & " ("& oRec("jobnr") &")"
                        strJobogKunderTxt = strJobogKunderTxt 
                        end if
    
                        strJobogKunderTxt = strJobogKunderTxt &"</li>"
      
                     else
    
                        if cdbl(jq_jobid) = cdbl(oRec("jid")) then
                        jobSEL = "SELECTED"
                        else
                        jobSEL = ""
                        end if

                      
                        strJobogKunderTxt = strJobogKunderTxt & "<option value="& oRec("jid") &" "& jobSEL &">"& oRec("jobnavn") & " ("& oRec("jobnr") &")"
                         

                        if lto = "hestia" OR lto = "intranet - local" then
                        strJobogKunderTxt = strJobogKunderTxt &" ..... "& oRec("jobstartdato") & ""
                        else
                        strJobogKunderTxt = strJobogKunderTxt 
                        end if
                        
                        strJobogKunderTxt = strJobogKunderTxt &"</option>"

                     end if     
                 



                lastKnavn = oRec("kkundenavn") 
                c = c + 1
                oRec.movenext
                wend
                oRec.close

                 
                    if cint(c) = 0 then
                    
                         if jobkundesog <> "-1" then 'kunde job sog DD
                         strJobogKunderTxt = strJobogKunderTxt & "<li class=""span_job"">"&expence_txt_060&"</li>"
                         else
                         strJobogKunderTxt = strJobogKunderTxt & "<option value=""-1"" DISABLED>"&expence_txt_060&"</option>"
                         end if

                    else

                          strJobogKunderTxt = strJobogKunderTxt & "<option value=""-1"" DISABLED>"&expence_txt_061&": "& c &"</option>"

                    end if
       
                

                    '*** ÆØÅ **'
                    call jq_format(strJobogKunderTxt)
                    strJobogKunderTxt = jq_formatTxt


                    response.write strJobogKunderTxt

                end if




                case "FN_sogakt"

               
                '*** Søg Aktiviteter 
                keycode = session("lto")
                ltouse = ""
                if keycode = "9K2016-0410-TO171" then
                ltouse = "epi2017"
                end if

                if len(trim(request("jq_newfilterval"))) <> 0 then
                filterVal = 1 
                aktsog = request("jq_newfilterval")
                else
                filterVal = 0
                aktsog = "6xxxxxfsdf554"
                end if
        
                medid = request("jq_medid")
                aktid = request("jq_aktid")
    
                if len(trim(request("jq_jobid"))) <> 0 then        
                jobid = request("jq_jobid")
                else
                jobid = 0
                end if

                'positiv aktivering
                'pa = request("jq_pa") 
                call positiv_aktivering_akt_fn()
                pa = pa_aktlist
                pa_only_specifikke_akt = positiv_aktivering_akt_val


                '*** HER ***

                varTjDatoUS_man = request("varTjDatoUS_man")
                varTjDatoUS_son = dateAdd("d", 6, varTjDatoUS_man)

                varTjDatoUS_man = year(varTjDatoUS_man) &"/"& month(varTjDatoUS_man) &"/"& day(varTjDatoUS_man)
                varTjDatoUS_son = year(varTjDatoUS_son) &"/"& month(varTjDatoUS_son) &"/"& day(varTjDatoUS_son)


                '*** Vis kun aktiviteter med forecast på
                call aktBudgettjkOn_fn()
                '*** Skal akt lukkes for timereg. hvis forecast budget er overskrddet..?
                '** MAKS budget / Forecast incl. peridoe afgrænsning
                call akt_maksbudget_treg_fn()

                if cint(aktBudgettjkOnViskunmbgt) = 1 then
                viskunForecast = 1
                else
                viskunForecast = 0
                end if


                timerTastet = 0 'request("timer_tastet")
                ibudgetaar = aktBudgettjkOn_afgr
                ibudgetmd = datePart("m", aktBudgettjkOnRegAarSt, 2,2) 
                aar = datepart("yyyy", varTjDatoUS_man, 2,2)
                md = datepart("m", varTjDatoUS_man, 2,2)


                '*** Forecast tjk & Jobstatus
                risiko = 0
                jobstatusTjk = 1
                strSQLjobRisisko = "SELECT j.risiko, jobstatus FROM job AS j WHERE id = "& jobid
                oRec5.open strSQLjobRisisko, oConn, 3
                if not oRec5.EOF then
                risiko = oRec5("risiko")
                jobstatusTjk = oRec5("jobstatus")
                end if
                oRec5.close 

                call allejobaktmedFC(viskunForecast, medid, jobid, risiko)

                
                '*** Datospærring Vis først job når stdato er oprindet
                call lukaktvdato_fn()
                ignJobogAktper = lukaktvdato

               
	            if (cint(ignJobogAktper) = 1 OR cint(ignJobogAktper) = 2 OR cint(ignJobogAktper) = 3) then
	            strSQLDatoKri = " AND ((a.aktstartdato <= '"& varTjDatoUS_son &"' AND a.aktslutdato >= '"& varTjDatoUS_man &"') OR (a.fakturerbar = 6))" 
	            end if


                 if lto = "mpt" OR session("lto") = "9K2017-1121-TO178" then
                  onlySalesact = ""

                else

                if cint(jobstatusTjk) = 3 then 'tilbud
                onlySalesact = " AND a.fakturerbar = 6"
                else
                onlySalesact = ""
                end if

                end if


                call akttyper2009(2)
                
                
                'response.write "aty_sql_realhours: "& aty_sql_realhours & " aty_sql_hide_on_treg: " & replace(aty_sql_hide_on_treg, )           
                'response.end


                if filterVal <> 0 then
            
                 
    
                
                         
               'pa = 0 '***ÆNDRES til variable
               'if cint(pa) = 1 then
               'strSQL= "SELECT a.id AS aid, navn AS aktnavn FROM timereg_usejob LEFT JOIN aktiviteter AS a ON (a.id = tu.aktid) "_
               '&" WHERE tu.medarb = "& medid &" AND tu.jobid = "& jobid &" AND aktid <> 0 AND a.navn LIKE '%"& aktsog &"%' AND aktstatus = 1 AND ("& replace(aty_sql_realhours, "tfaktim", "a.fakturerbar") &") AND ("& aty_sql_hide_on_treg &") ORDER BY navn LIMIT 20"   

                    
                    '** Select eller søgeboks
                    call mobil_week_reg_dd_fn()
                    
                    
                  
                    if cint(mobil_week_reg_akt_dd) <> 1 then 'AND aktsog <> "-1" 

                    strAktTxt = strAktTxt &"<span style=""color:red; font-size:9px; float:right;"" class=""luk_aktsog"">[X]</span><ul>"    
                    strSQlAktSog = "AND navn LIKE '%"& aktsog &"%'"
                    
                    else
                    strSQlAktSog = ""

                            '** Forvalgt 1 aktivitet
                            if cint(mobil_week_reg_akt_dd_forvalgt) <> 1 AND cint(mobil_week_reg_akt_dd) = 1 then
                            strAktTxt = strAktTxt & "<option value=""-1"">"& ttw_txt_024 &"</option>" 
                            end if

                    end if


                strSQLkunSpecialAkt = ""

                if ltouse = "epi2017" then

                call meStamdata(medid)
                 
                    if cint(meType) = 14 then
                    strSQLkunSpecialAkt = " AND (a.navn = 'Data Collection' OR a.navn = 'Break (CPH Airport)')"
                    end if

                end if



               if cint(pa) = 1 then '**Kun på Personlig aktliste
    
    
                   'Positiv aktivering
                   if cint(pa_only_specifikke_akt) then

                   strSQL = "SELECT a.id AS aid, navn AS aktnavn "_
                   &" FROM timereg_usejob AS tu LEFT JOIN aktiviteter AS a ON (a.id = tu.aktid) "_
                   &" WHERE tu.medarb = "& medid &" AND tu.jobid = "& jobid &" AND aktid <> 0 "& strSQlAktSog &" AND aktstatus = 1 AND ("& aty_sql_hide_on_treg &") "& forecastAktids &" AND a.navn IS NOT NULL "& strSQLkunSpecialAkt &" "& onlySalesact &" ORDER BY fase, sortorder, navn LIMIT 150"   
                   'AND ("& replace(aty_sql_realhours, "tfaktim", "a.fakturerbar") &")
                   else 

                   strSQL = "SELECT a.id AS aid, navn AS aktnavn "_
                   &" FROM timereg_usejob AS tu LEFT JOIN aktiviteter AS a ON (a.job = tu.jobid) "_
                   &" WHERE tu.medarb = "& medid &" AND tu.jobid = "& jobid &" AND aktid = 0 "& strSQlAktSog &" AND aktstatus = 1  AND ("& aty_sql_hide_on_treg &") "& forecastAktids &" AND a.navn IS NOT NULL "& strSQLkunSpecialAkt &" "& onlySalesact &" ORDER BY fase, sortorder, navn LIMIT 150" 
                   'AND ("& replace(aty_sql_realhours, "tfaktim", "a.fakturerbar") &")
                   end if


                  

               else


               select case lto
               case "nt"
               jobid = "12839"
               case "xlm"
               jobid = "6730"            

               case else
               jobid = jobid

                         '*** Finder medarbejders projektgrupper 
                        '** Medarbejder projektgrupper **'
                        medarbPGrp = "#0#" 
                        strMpg = "SELECT projektgruppeId, medarbejderId, teamleder FROM progrupperelationer WHERE medarbejderId = "& medid & " GROUP BY projektgruppeId"

                     
                        'response.write "strMpg " & strMpg
                        'response.end

                        oRec5.open strMpg, oConn, 3
                        while not oRec5.EOF
                        medarbPGrp = medarbPGrp & ",#"& oRec5("projektgruppeId") &"#"         
        
                        oRec5.movenext
                        wend
                        oRec5.close 

               end select
               
            

               strSQL = "SELECT a.id AS aid, navn AS aktnavn, projektgruppe1, projektgruppe2, projektgruppe3, "_
               &" projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM aktiviteter AS a "_
               &" WHERE a.job = " & jobid & " AND a.job <> 0 "& strSQlAktSog &" AND aktstatus = 1 AND ("& aty_sql_hide_on_treg &") "& strSQLkunSpecialAkt &" "& onlySalesact &" ORDER BY fase, sortorder, navn LIMIT 150"      
    

              end if


                 'if lto = "essens" then
                'response.write "<option>"& strSQL &"</option>"
                'response.flush 
                'end if

                'response.write "strSQL " & strSQL
                'response.end
            

                afundet = 0
                oRec.open strSQL, oConn, 3
                while not oRec.EOF

                if cint(pa) = 1 OR lto = "nt" then 'Positiv aktivering

                    showAkt = 1

                else
        
                    showAkt = 0
                    if instr(medarbPGrp, "#"& oRec("projektgruppe1") &"#") <> 0 _
                    OR instr(medarbPGrp, "#"& oRec("projektgruppe2") &"#") <> 0 _
                    OR instr(medarbPGrp, "#"& oRec("projektgruppe3") &"#") <> 0 _
                    OR instr(medarbPGrp, "#"& oRec("projektgruppe4") &"#") <> 0 _
                    OR instr(medarbPGrp, "#"& oRec("projektgruppe5") &"#") <> 0 _
                    OR instr(medarbPGrp, "#"& oRec("projektgruppe6") &"#") <> 0 _
                    OR instr(medarbPGrp, "#"& oRec("projektgruppe7") &"#") <> 0 _
                    OR instr(medarbPGrp, "#"& oRec("projektgruppe8") &"#") <> 0 _
                    OR instr(medarbPGrp, "#"& oRec("projektgruppe9") &"#") <> 0 _
                    OR instr(medarbPGrp, "#"& oRec("projektgruppe10") &"#") <> 0 then
                    showAkt = 1
                    end if 

                end if
               
                  

                
                '** Forecast peridore afgrænsning
                'if cint(akt_maksforecast_treg) = 1 then
                if cint(aktBudgettjkOn) = 1 then
                    call ressourcefc_tjk(ibudgetaar, ibudgetmd, aar, md, medid, oRec("aid"), timerTastet)
                end if
               
                
                
                if cint(showAkt) = 1 then 
                 
              
                    if cint(aktBudgettjkOn) = 1 then


                        if len(trim(feltTxtValFc)) <> 0 then
                        fcsaldo_txt = "<span style=""font-weight:lighter; font-size:9px;""> ("&expence_txt_068&": "& formatnumber(feltTxtValFc, 2) & " / "& formatnumber(fctimer,2) &" "&expence_txt_069&")</span>"
                        end if

                            optionFcDis = ""
                            if cint(akt_maksforecast_treg) = 1 then
                                if feltTxtValFc <= 0 then
                                      optionFcDis = "DISABLED"
                                end if
                            end if

                    else

                    fcsaldo_txt = ""

                    end if
                 

                
                       if cint(mobil_week_reg_akt_dd) <> 1 then

                            if optionFcDis <> "DISABLED" then
                                strAktTxt = strAktTxt & "<li class=""span_akt"" id=""span_akt_"& oRec("aid") &"""><input type=""hidden"" id=""hiddn_akt_"& oRec("aid") &""" value="""& oRec("aktnavn") &""">" 
                                strAktTxt = strAktTxt & ""& oRec("aktnavn") &" "& fcsaldo_txt &"</li>" 
                            else
                                strAktTxt = strAktTxt & "<li style=""background-color:#CCCCCC; color:#999999;"">"& oRec("aktnavn") &" "& fcsaldo_txt &"</li>" 
                            end if

                        else

                            strAktTxt = strAktTxt & "<option value="& oRec("aid") &" "& optionFcDis &">"& oRec("aktnavn") &" "& fcsaldo_txt &"</option>"

                        end if            


                end if
                

                afundet = afundet + 1
                oRec.movenext
                wend
                oRec.close

              
                if cint(mobil_week_reg_akt_dd) <> 1 then
                    strAktTxt = strAktTxt &"</ul>"  
                else
    
                    if afundet = 0 then
                    strAktTxt = strAktTxt & "<option value=""-1"" DISABLED>"& ttw_txt_025 &"</option>" 
                    end if 

    
                end if      


                    '*** ÆØÅ **'
                    call jq_format(strAktTxt)
                    strAktTxt = jq_formatTxt


                    response.write strAktTxt

                end if




            case "FN_createOutlay"
                'response.Write "HEJ "
                'response.End
                strEditor = session("user")
		        strDato = year(now)&"/"&month(now)&"/"&day(now)
                matregid = 0
                otf = 1
                medid = session("mid")

                jobid = request("jq_jobid")
                aktid = request("aktid")
                if aktid <> -1 then
                aktid = aktid
                else
                aktid = 0
                end if
                aftid = 0
                matId = 0
                
                if len(trim(request("antal"))) <> 0 then
                intAntal = request("antal")
                else
                intAntal = 1 'Antal findes ikke paa mobil, saa altid en
                end if

                regdato = year(request("FM_datoer"))&"/"&month(request("FM_datoer"))&"/"&day(request("FM_datoer")) 'fornu
                valuta = request("FM_udlaeg_valuta")
                
                'select case lto 
                'case "synergi1", "intranet - local", "epi", "epi_as", "epi_se", "epi_os", "epi_uk", "alfanordic"
                'intKode = 1 'Intern
                'case else 
	            'intKode = 2 'Ekstern = viderefakturering
                'end select

                intKode = request("FM_udlaeg_faktbar")

                personlig = request("FM_udlaeg_form")
                bilagsnr = "" 'for nu
                pris = request("FM_udlaeg_belob")
                pris = replace(pris,".",",")
                salgspris = pris 'Fordi det er mobilen, dvs. udlaeg derfor ingen salgspris
                navn = request("FM_udlaeg_navn")
                gruppe = request("FM_udlaeg_gruppe")
                varenr = 0
                opretlager = 0
                betegnelse = ""
                mat_func = "dbopr"
                matreg_opdaterpris = 0
                matava = 0
                mattype = 1

                basic_valuta = request("basic_valuta")
                basic_kurs = request("basic_kurs")
                basic_belob = request("basic_belob")

                'response.Write basic_valuta &" "& basic_kurs &" "& basic_belob
                'response.End
                'call MatTest(matregid, otf, medid, jobid, aktid, aftid, matId, strEditor, strDato, intAntal, regdato, valuta, intkode, personlig, bilagsnr, pris, salgspris, navn, gruppe, varenr, opretlager, betegnelse, mat_func, matreg_opdaterpris, matava)
                call indlaes_mat(matregid, otf, medid, jobid, aktid, aftid, matId, strEditor, strDato, intAntal, regdato, valuta, intkode, personlig, bilagsnr, pris, salgspris, navn, gruppe, varenr, opretlager, betegnelse, mat_func, matreg_opdaterpris, matava, mattype, basic_valuta, basic_kurs, basic_belob)


             
        
        end select
        response.end
        end if

%>



<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

    
    <!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
    <!--#include file="../inc/regular/top_menu_mobile.asp"-->


 <%
        if len(session("user")) = 0 then
	    errortype = 5
	    call showError(errortype)
	    response.End
        end if
        %>


<head>

<!-- <meta name="viewport" content="width=device-width, initial-scale=1.0"> -->
<meta name="viewport" content='width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=0'>

<title>TimeOut mobile</title>


       

    <script src="js/timetag_web_jav_2018_4.js" type="text/javascript"></script>
    <script type="text/javascript" src="js/expence_jav10.js"></script>
<script src="../to_2015/js/plugins/datepicker/bootstrap-datepicker.js"></script>

<!--<script src="js/timetag_web_jav_2018.js" type="text/javascript"></script>
    
    <script type="text/javascript" src="../to_2015/js/plugins/datepicker/bootstrap-datepicker.js"></script>-->




    <style type="text/css">

        input[type="text"] 
        {
          height:125%;
          font-size:125%;
        }
        input[type="number"] 
        {
          height:125%;
          font-size:125%;
        }
        input[type="button"] 
        {
          height:125%;
          font-size:125%;
        }

        .span_job {
        list-style:none;
        font-size:125%;
        }

         .span_akt {
        list-style:none;
        font-size:125%;
        }

        .span_mat {
        list-style:none;
        font-size:125%;
        }

        

    </style>

</head>
    

    <%
    if request("func") = "delexpence" then 
        '*** Her spørges om det er ok at der slettes et udlæg ***

        dato = request("dato")
        matid = request("matid")
        matnavn = request("matnavn")
        matpris = request("pris")

        %>
        <div class="container">
            <div class="portlet">
                <div class="row">
                    <h5 class="col-lg-12" style="text-align:center"><%=expence_txt_062 %> <b><%=expence_txt_063 %></b> <%=expence_txt_064 %></h5>
                </div>
                <div class="row">
                    <h6 class="col-lg-12" style="text-align:center"><%=matnavn &" d. "& dato &" til "& formatnumber(matpris, 2) %></h6>
                </div>

                <br /><br />

                <div class="row">
                    <div class="col-lg-12" style="text-align:center">
                        <a class="btn btn-primary" href="expence.asp?func=delexpenceOK&matid=<%=matid %>&matnavn=<%=matnavn %>" style="width:75px; font-size:115%;"><%=expence_txt_049 %></a>
                        &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp
                        <a class="btn btn-default" href="Javascript:history.back()" style="width:75px; font-size:115%;"><%=expence_txt_050 %></a>
                    </div>
                </div>

            </div>
        </div>
        <%
	    'oskrift = "Udlæg" 
        'slttxta = "Du er ved at <b>SLETTE</b> et udlæg. Er dette korrekt?"
        'slttxtb = dato &" "& matnavn &" "& formatnumber(matpris, 2)
        'slturl = "expence.asp?func=delexpenceOK&id="& matid

       
        'call sletcnf_2015(oskrift, slttxta, slttxtb, slturl)

        response.End
    end if



    if request("func") = "delexpenceOK" then

            matid = request("matid")
            matnavn = request("matnavn")

            strSQLdel = "DELETE FROM materiale_forbrug WHERE id = " & matid
            oConn.execute(strSQLdel) 

            matnavn = replace(matnavn, "'", "")
	        call insertDelhist("mat", matid, 0, matnavn, session("mid"), session("user"))

            response.Redirect "expence.asp"
    end if            
    %>

    <%
    if request("func") = "oprdiets" then

        if len(trim(request("strStdatoDiet"))) = 0 OR len(trim(request("strSldatoDiet"))) = 0 then
            response.Redirect "expence.asp"
            response.End
        end if

        strDist = "Uangivet"
        strMedid = session("mid")

        strStdato = request("strStdatoDiet")
        strSldato = request("strSldatoDiet")

        strJobid = request("FM_jobid")
        strkontoid = 0

        strMorgenAntal = request("FM_morgenantal")
        strMiddagAntal = request("FM_frokostantal") 
        strAftenAntal = request("FM_aftenantal")    

        strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)

        diet_daypriceVal = request("FM_diet_dayprice")
        if len(trim(diet_daypriceVal)) <> 0 then
        diet_daypriceVal = diet_daypriceVal
        else
        diet_daypriceVal = 0
        end if

        if len(trim(request("FM_diet_dayprice"))) <> 0 then
        diet_dayprice = replace(request("FM_diet_dayprice"), ",", ".")
        else
        diet_dayprice = 0
        end if

        diet_dayprice_halfVal = request("FM_diet_dayprice_delvis")
        if len(trim(diet_dayprice_halfVal)) <> 0 then
        diet_dayprice_halfVal = diet_dayprice_halfVal
        else
        diet_dayprice_halfVal = 0
        end if

        if len(trim(request("FM_diet_dayprice_delvis"))) <> 0 then
        diet_dayprice_half = replace(request("FM_diet_dayprice_delvis"), ",", ".")
        else
        diet_dayprice_half = 0
        end if


        'Prices
        diet_morgenpriceVal = request("FM_diet_morgenprice")
        if len(trim(diet_morgenpriceVal)) <> 0 then
        diet_morgenpriceVal = diet_morgenpriceVal
        else
        diet_morgenpriceVal = 0
        end if

        diet_middagpriceVal = request("FM_diet_middagprice")
        if len(trim(diet_middagpriceVal)) <> 0 then
        diet_middagpriceVal = diet_middagpriceVal
        else
        diet_middagpriceVal = 0
        end if

        diet_aftenpriceVal = request("FM_diet_aftenprice")
        if len(trim(diet_aftenpriceVal)) <> 0 then
        diet_aftenpriceVal = diet_aftenpriceVal
        else
        diet_aftenpriceVal = 0
        end if

        if len(trim(request("FM_diet_morgenprice"))) <> 0 then
        diet_morgenprice = replace(request("FM_diet_morgenprice"), ",", ".")
        else
        diet_morgenprice = 0
        end if

        
        if len(trim(request("FM_diet_middagprice"))) <> 0 then
        diet_middagprice = replace(request("FM_diet_middagprice"), ",", ".")
        else
        diet_middagprice = 0
        end if

        
        if len(trim(request("FM_diet_aftenprice"))) <> 0 then
        diet_aftenprice = replace(request("FM_diet_aftenprice"), ",", ".")
        else
        diet_aftenprice = 0
        end if

        strBilag = 0

        call akttyper2009(2)
        call traveldietexp_fn()

        'Opret Diet Reg

        if (isDate(strStdato) = true AND isDate(strSldato) = true) AND len(trim(strDist)) = 0 then
            useleftdiv = "to_2015"
			errortype = 206
			call showError(errortype)
		    Response.End
        end if

        if isDate(strStdato) = true then
            strStdato = year(strStdato) &"-"& month(strStdato) &"-"& day(strStdato)
        end if


        if cint(request("brug_tidspunkt")) = 1 then
            if len(trim(request("FM_diet_sttime"))) <> 0 then
                strSttime = request("FM_diet_sttime")
            else       
                if (isDate(strStdato) = true AND isDate(strSldato) = true) then
                    useleftdiv = "to_2015"
			        errortype = 205
			        call showError(errortype)
		            Response.End
                end if
            end if
        else
            strSttime = "00:00:00"
        end if

        strStdato = strStdato &" "& strSttime

        if isDate(strSldato) = true then
            strSldato = year(strSldato) &"-"& month(strSldato) &"-"& day(strSldato)
        end if

         if cint(request("brug_tidspunkt")) = 1 then
            if len(trim(request("FM_diet_sltime"))) <> 0 then
                strSltime = request("FM_diet_sltime")
            else
                if (isDate(strStdato) = true AND isDate(strSldato) = true) then
                    useleftdiv = "to_2015"
			        errortype = 205
			        call showError(errortype)
		            Response.End
                end if
            end if
        else
            strSltime = "00:00:00"
        end if

        strSldato = strSldato &" "& strSltime



        strDelVis = 0


        antalDagArr = dateDiff("d", strStdato, strSldato, 2,2)
        timerForbrugprDagOverskreddet = 0

        for d = 0 TO antalDagArr

        if d = 0 then
        tjkThisDatp = year(strStdato) & "-" & month(strStdato) & "-" & day(strStdato)
        else
        tjkThisDatp = dateAdd("d", 1, tjkThisDatp)
        tjkThisDatp = year(tjkThisDatp) & "-" & month(tjkThisDatp) & "-" & day(tjkThisDatp)
        end if 

        timerBrugt = 0
        strSQLtimreal = "SELECT COALESCE(SUM(timer), 0) AS timerbrugt FROM timer WHERE tmnr = "& strMedid &" AND tdato = '"& tjkThisDatp &"' AND ("& aty_sql_realhours &") GROUP BY tdato, tmnr"
        
        
        'response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;antalDagArr: "& antalDagArr &" strSQLtimreal: " & strSQLtimreal & "<br>"
                      
        oRec5.open strSQLtimreal, oConn, 3
        while not oRec5.EOF
               
        timerBrugt = oRec5("timerbrugt")
        
        if cdbl(timerBrugt) > cdbl(traveldietexp_maxhours) then
        timerForbrugprDagOverskreddet = 1
        end if         
        
        oRec5.movenext
        wend
        oRec5.close

        next

        select case lto 
        case "plan", "lm"
        timerForbrugprDagOverskreddet = 0
        case else
        timerForbrugprDagOverskreddet = timerForbrugprDagOverskreddet
        end select

        
        if cint(timerForbrugprDagOverskreddet) = 1 then
            useleftdiv = "to_2015"
			errortype = 178
			call showError(errortype)
		    Response.End
        end if


        diet_dageIalt = dateDiff("h", strStdato, strSldato, 2,2) / 24
        if diet_dageIalt <> 0 then
            if instr(diet_dageIalt, ",") <> 0 then
            komma = instr(diet_dageIalt, ",")
            diet_dageIalt = left(diet_dageIalt, komma+2)
            else
            diet_dageIalt = diet_dageIalt
            end if

        else
        diet_dageIalt = 0
        end if

        'Lægger en dag til for tia, fordi de altid får penge uanset om de sover der eller ej
        if lto = "tia" then
            diet_dageIalt = diet_dageIalt + 1
        end if


        if len(trim(strJobids)) <> 0 then
        strJobid = strJobid 
        else
        strJobid = 0
        end if

        if len(trim(strMorgenAntal)) <> 0 then
        strMorgenAntal = strMorgenAntal
        else
        strMorgenAntal = 0
        end if

        if len(trim(strMiddagAntal)) <> 0 then
        strMiddagAntal = strMiddagAntal
        else
        strMiddagAntal = 0
        end if

        if len(trim(strAftenAntal)) <> 0 then
        strAftenAntal = strAftenAntal
        else
        strAftenAntal = 0
        end if

        strMorgenAntal = replace(strMorgenAntal, ",", "")
        strMorgenAntal = replace(strMorgenAntal, ".", "")

        strMiddagAntal = replace(strMiddagAntal, ",", "")
        strMiddagAntal = replace(strMiddagAntal, ".", "")

        strAftenAntal = replace(strAftenAntal, ",", "")
        strAftenAntal = replace(strAftenAntal, ".", "")

        diet_total = (strMorgenAntal/1+strMiddagAntal/1+strAftenAntal/1)
        if len(trim(diet_total)) <> 0 then
        diet_total = diet_total
        else
        diet_total = 0
        end if

        if lto = "plan" or lto = "outz" then
            if cint(strDelvis) = 1 then
            makskost = (diet_dayprice_halfVal/1)*diet_dageIalt
            else
            makskost = (diet_daypriceVal/1)*diet_dageIalt
            end if
        else
        makskost = (diet_daypriceVal/1)*diet_dageIalt
        end if

        totalefterReduktion = makskost/1 - ((strMorgenAntal*diet_morgenpriceVal/1) + (strMiddagAntal*diet_middagpriceVal/1) + (strAftenAntal*diet_aftenpriceVal/1))

        makskost = replace(makskost, ",", ".")
        diet_dageIalt = replace(diet_dageIalt, ",", ".")
        totalefterReduktion = replace(totalefterReduktion, ",", ".")

        if isDate(strStdato) = true AND isDate(strSldato) = true then
            
            strSQlins = ("INSERT INTO traveldietexp (diet_namedest, diet_stdato, diet_sldato, diet_jobid, diet_konto, diet_morgen, diet_middag, diet_aften, diet_total, diet_dayprice, diet_dayprice_half, "_
            &" diet_morgenamount, diet_middagamount, diet_aftenamount, diet_mid, diet_bilag, diet_maksamount, diet_rest, diet_traveldays, diet_delvis) "_
            &" VALUES ('"& strDist &"', '"& strStdato &"', '"& strSldato &"', "& strJobid &", "& strkontoid &", "_
            &" "& strMorgenAntal &", "& strMiddagAntal &", "& strAftenAntal &", "& diet_total &", "& diet_dayprice &", "& diet_dayprice_half &", "_
            &" "& diet_morgenprice &","& diet_middagprice &","& diet_aftenprice &", "& strMedid &", "& strBilag &", "& makskost &", "& totalefterReduktion &", "& diet_dageIalt &", "& strDelVis &")")

            'response.write strSQlins
            'response.flush

            oConn.execute(strSQlins)

        end if

        'response.Write "<br> Antal Dage " & antalDagArr & "OVER skredet " & timerForbrugprDagOverskreddet


        'response.Write "Diet Reg"
        response.Redirect "expence.asp"
        response.End


    end if
    %>







    
    <%

       

        select case lto
        case "tbg"
        tomobjid = 4
        case else
        tomobjid = session("tomobjid")
        end select


        call mobil_week_reg_dd_fn()
    %>

       


    <%call mobile_header %>


<body style="background-color:white;">
<div class="container" style="height:100%;">

    <input type="hidden" id="nameTxt" value="<%=expence_txt_006 %>" />
    <input type="hidden" id="priceTxt" value="<%=expence_txt_007 %>" />
    <input type="hidden" id="notbillableTxt" value="<%=expence_txt_070 %>" />
    <input type="hidden" id="billableTxt" value="<%=expence_txt_013 %>" />
    <input type="hidden" id="companypaidTxt" value="<%=expence_txt_012 %>" />
    <input type="hidden" id="personalTxt" value="<%=expence_txt_011 %>" />
    <input type="hidden" id="selectimageTxt" value="<%=expence_txt_014 %>" />
    
    <input type="hidden" id="thisfile" value="expence" />
    

    <div class="portlet">
       
        <div class="portlet-body">
            
            <input type="hidden" id="varTjDatoUS_man" value="<%=year(now) &"-"& month(now) &"-"& day(now) %>" />
            <input type="hidden" id="FM_medid_k" name="FM_medid_k" value="<%=session("mid") %>"/>
            <input type="hidden" id="FM_medid" name="FM_medid_k" value="<%=session("mid") %>"/>
            <input type="hidden" id="jq_lto" name="jq_lto" value="<%=lto %>"/>
            <input type="hidden" id="tomobjid" name="tomobjid" value="<%=tomobjid %>"/>
            <input type="hidden" id="FM_pa" name="FM_pa" value="<%=pa_aktlist %>"/>
            <input type="hidden" id="mobil_week_reg_job_dd" name="" value="<%=mobil_week_reg_job_dd %>"/> 
            <input type="hidden" id="mobil_week_reg_akt_dd" name="" value="<%=mobil_week_reg_akt_dd %>"/>

            <div class="row">
                <div class="col-lg-12"><span id="errorMessage" style="color:darkred;"></span></div>
            </div>


            <%
        
                        if day(now) < 10 then
                        todayDay = "0"& day(now)
                        else
                        todayDay = day(now)
                        end if
              
                        if month(now) < 10 then
                        todayMonth = "0"& month(now)
                        else
                        todayMonth = month(now)
                        end if
           
                        todayYear = year(now)      
              
                        select case lto 'mulighed for VÆLG DATO 
                        case "xplan", "xoutz", "xintranet - local", "tbg"
                              %> <input type="hidden" id="FM_datoer" name="FM_datoer" value="<%=todayDay &"-"& todayMonth &"-"& todayYear%>"/><%
                        case else
                       
                        %>  
                        
                        <style>
                            .datepicker-dropdown
                            {
                                margin-top:-10px;
                            }
                        </style>

                        <div class="row">
                        
                            <div class="col-lg-12">
                                <div class='input-group date' id='datepicker3'>
                                    <input type="text" class="form-control" id="FM_datoer" name="FM_datoer" value="<%=todayDay &"-"& todayMonth &"-"& todayYear%>" placeholder="dd-mm-yyyy"  readonly />
                                    <span class="input-group-addon input-small">
                                            <span class="fa fa-calendar">
                                            </span>
                                    </span>
                                </div>
                            </div>

                        </div>
                        <%end select %>     



            <div class="row">
                <div class="col-lg-12">

                    <%select case lto
                    case "nt"
                        %>
                        <input type="text" id="FM_job" name="FM_job" value="" placeholder="North-Tex" class="form-control" autocomplete="off" disabled />
                         <input type="hidden" id="FM_jobid" value="12839"/>
                         <input type="hidden" id="dv_job" name="FM_jobid" value="12839"/>
                        
                        <%
                     case "xlm"
                        %>
                        <input type="text" id="FM_job" name="FM_job" value="" placeholder="Landsforeningen" class="form-control" autocomplete="off" disabled />
                         <input type="hidden" id="FM_jobid" value="6730"/>
                         <input type="hidden" id="dv_job" name="FM_jobid" value="6730"/>
                        
                        <%

                    case else%>

                        <%if lto = "lm" then %>
                            <span style="font-size:9px;" id="kundenavnholder"></span>
                        <%end if %>

                        <%if mobil_week_reg_job_dd = 1 then %>
                             <input type="hidden" id="FM_job" value="-1"/>
                                <select id="dv_job" name="FM_jobid" class="form-control">
                                 
                                </select>
                        <%else %>
                            <input type="text" id="FM_job" name="FM_job" value="" placeholder="<%=ttw_txt_014 %>" class="form-control" autocomplete="off" />
                            <input type="hidden" id="FM_jobid" name="FM_jobid" value="0"/>
                            <div id="dv_job" style="padding:5px 5px 5px 5px; display:none; visibility:hidden;"></div>
                        <%end if 
                        
                        
                    end select    
                    %>



                </div>
            </div>
             
            
            
            <%select case lto
                    case "nt"
                        %>
                        <input type="hidden" id="FM_akt" value="-1"/>
                        <input type="hidden" id="FM_aktid" value="38"/><!-- Expence -->
                        <input type="hidden" id="dv_akt" name="FM_aktivitetid" value="38"/><!-- Expence -->
                        <%

                      case "xlm"
                        %>
                        <input type="hidden" id="FM_akt" value="-1"/>
                        <input type="hidden" id="FM_aktid" value="57814"/><!-- Expence -->
                        <input type="hidden" id="dv_akt" name="FM_aktivitetid" value="57814"/><!-- Expence -->
                        <%

                    case else%>
            
            
            <div class="row">
                <div class="col-lg-12">

                        <%if cint(mobil_week_reg_akt_dd) = 1 then %>
                        <input type="hidden" id="FM_akt" value="-1"/>
                        <!--<textarea id="dv_akt_test"></textarea>-->
                        <select id="dv_akt" name="FM_aktivitetid" class="form-control">
                            <option value="0">..</option>
                        </select>
                        <%else %>
                        <input type="text" id="FM_akt" name="activity" value="" class="form-control" placeholder="<%=ttw_txt_004 %>" autocomplete="off" />
                        <input type="hidden" id="FM_aktid" name="FM_aktivitetid" value="0"/>
                        <div id="dv_akt" class="dv-closed" style="padding:5px 5px 5px 5px;"></div> 
                        <%end if %>

                </div>
            </div>

             <%end select %>

            <br />

            <style>
                td {
                    padding:8px;
                }
            </style>


            <%
            kmpris = 0
            strSQl = "SELECT kmpris FROM licens WHERE id = 1"
            oRec.open strSQl, oConn, 3
            if not oRec.EOF then
                kmpris = oRec("kmpris")
            end if
            oRec.close

            mkoregnr = ""
            strSQL = "SELECT mkoregnr FROM medarbejdere WHERE mid = "& session("mid")
            oRec.open strSQL, oConn, 3
            if not oRec.EOF then
                mkoregnr = oRec("mkoregnr")
            end if
            oRec.close

            'madr = ""
            'mpostnr = ""
            'mcity = ""
            'strSQL = "SELECT madr, mpostnr, mcity FROM medarbejdere WHERE mid = "& session("mid")
            'oRec.open strSQL, oConn, 3
            'if not oRec.EOF then
                'madr = oRec("madr")
                'mpostnr = oRec("mpostnr")
                'mcity = oRec("mcity")
            'end if
            'oRec.close

            if lto = "lm" then
                lmdisplay = "normal"
            else
                lmdisplay = "none"
            end if

            %>

            <input type="hidden" id="prisprkm" value="<%=kmpris %>" />
            <table style="width:100%; border:solid 1px #c9c9c9; margin-top:10px; display:<%=lmdisplay%>;" id="dietTable">
                <tr>
                    <td colspan="2"><input type="text" class="form-control" placeholder="<%=expence_txt_095 %>" id="kmregsted" /></td>
                </tr>
                <tr>
                    <td colspan="2">
                        <span style="font-size:9px;"><%=expence_txt_082 %></span>
                        <input type="text" class="form-control" placeholder="<%=expence_txt_085 %>" id="fradistination" />
                        <div id="dv_fradistination" style="padding:5px 5px 5px 5px; display:none;">
                            <%'if madr <> "" AND mpostnr <> "" AND mcity <> "" then  %>
                                <li><a class="fradistination_forvalgt" id="bopal" ><%=expence_txt_094 %></a></li>
                            <%
                            'end if

                            strSQL = "SELECT adresse, postnr, city FROM kunder WHERE useasfak = 1"
                            oRec.open strSQL, oConn, 3
                            while not oRec.EOF

                                response.Write "<li><a  class='fradistination_forvalgt'>"& oRec("adresse") &" "& oRec("postnr") &" "& oRec("city") &"</a></li>"
                            
                            oRec.movenext
                            wend
                            oRec.close
                            %>
                        </div>
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
                        <span style="font-size:9px;"><%=expence_txt_083 %></span>
                        <input type="text" class="form-control" placeholder="<%=expence_txt_085 %>" id="tildistination" />
                        <div id="dv_tildistination" style="padding:5px 5px 5px 5px; display:none;">
                            <%'if madr <> "" AND mpostnr <> "" AND mcity <> "" then  %>
                                <li><a class="tildistination_forvalgt" id="bopal" ><%=expence_txt_094 %></a></li>
                            <%
                            'end if

                            strSQL = "SELECT adresse, postnr, city FROM kunder WHERE useasfak = 1"
                            oRec.open strSQL, oConn, 3
                            while not oRec.EOF

                                response.Write "<li><a  class='tildistination_forvalgt'>"& oRec("adresse") &" "& oRec("postnr") &" "& oRec("city") &"</a></li>"
                            
                            oRec.movenext
                            wend
                            oRec.close
                            %>
                        </div>
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
                        <span style="font-size:9px;"><%=expence_txt_084 %></span>
                        <input type="text" class="form-control" placeholder="<%=expence_txt_085 %>" id="returdistination" />
                        <div id="dv_returdistination" style="padding:5px 5px 5px 5px; display:none;">
                            <%'if madr <> "" AND mpostnr <> "" AND mcity <> "" then  %>
                                <li><a class="returdistination_forvalgt" id="bopal" ><%=expence_txt_094 %></a></li>
                            <%
                            'end if

                            strSQL = "SELECT adresse, postnr, city FROM kunder WHERE useasfak = 1"
                            oRec.open strSQL, oConn, 3
                            while not oRec.EOF

                                response.Write "<li><a  class='returdistination_forvalgt'>"& oRec("adresse") &" "& oRec("postnr") &" "& oRec("city") &"</a></li>"
                            
                            oRec.movenext
                            wend
                            oRec.close
                            %>
                        </div>
                    </td>
                </tr>

                <tr>
                    <td colspan="2">
                        <span style="font-size:9px;"><%=expence_txt_086 %></span>
                        <input type="text" class="form-control" id="koregnr" placeholder="<%=expence_txt_087 %>" value="<%=mkoregnr %>" />
                    </td>
                </tr>
                <tr>
                    <td style="width:50%;"><input type="number" min="1" step="1" onkeypress='return event.charCode >= 48 && event.charCode <= 57' id="antalkm" placeholder="<%=expence_txt_088 %>" class="form-control" /></td>
                    <td><%=expence_txt_091 &" "& formatnumber(kmpris, 2) &" "& expence_txt_089 %> = <b id="kmtotalpris">-</b></td>
                </tr>
            </table>

            <%if lto = "lm" then %>
            <br />
            <%end if %>

            <%
            dagsfradrag = 0
            halvdagsfradrag = 0
            morgenDiet = 0
            middagDiet = 0
            aftendDiet = 0

            strSQLDiet = "SELECT tdf_diet_name, tdf_diet_morgenamount, tdf_diet_middagamount, tdf_diet_aftenamount, tdf_diet_dayprice, tdf_diet_dayprice_half FROM travel_diet_tariff WHERE tdf_diet_current = 1"
            oRec.open strSQLDiet, oConn, 3
                if not oRec.EOF then
                    morgenDiet = oRec("tdf_diet_morgenamount")
                    middagDiet = oRec("tdf_diet_middagamount")
                    aftendDiet = oRec("tdf_diet_aftenamount")
                    dagsfradrag = oRec("tdf_diet_dayprice")
                    halvdagsfradrag = oRec("tdf_diet_dayprice_half")
                end if
            oRec.close

            'dietSTSQL = year(now) &"-"& month(now) &"-"& day(now)
            'dietSTSQL
            
            select case lto
            case "plan", "intranet - local"
            vis_reduktion = 1
            hide_klokkeslet = 0
            brug_tidspunkt = 1
            case "tia"
            vis_reduktion = 0
            hide_klokkeslet = 0
            brug_tidspunkt = 0
            case else    
            vis_reduktion = 1 
            hide_klokkeslet = 0
            brug_tidspunkt = 1
            end select  

            %>

            <form action="expence.asp?func=oprdiets" method="post" id="dietsform" style="display:<%=lmdisplay%>;">
                <input type="hidden" value="<%=dagsfradrag %>" name="FM_diet_dayprice" id="dagsfradrag" />        
                <input type="hidden" value="<%=halvdagsfradrag %>" name="FM_diet_dayprice_delvis" />     
                <input type="hidden" value="<%=morgenDiet %>" name="FM_diet_morgenprice" />
                <input type="hidden" value="<%=middagDiet %>" name="FM_diet_middagprice" />
                <input type="hidden" value="<%=aftendDiet %>" name="FM_diet_aftenprice" />

                <input type="hidden" name="brug_tidspunkt" value="<%=brug_tidspunkt %>" />
                
                <div id="dietStacker" style="border:solid 1px #c9c9c9;">
                    <table style="width:100%; margin-top:10px;" id="dietTable">
                        <tr>
                            <td>
                                <div class='input-group date' >
                                    <input type="text" class="form-control" id="strStdatoDiet" name="strStdatoDiet" value="" placeholder="<%=expence_txt_078 %>" readonly />
                                    <span class="input-group-addon input-small">
                                            <span class="fa fa-calendar">
                                            </span>
                                    </span>
                                </div>
                            </td>

                            <td><input type="time" id="FM_diet_sttime" name="FM_diet_sttime" class="form-control" /></td>

                        </tr>

                        <tr>
                            <td>
                                <div class='input-group date' >
                                    <input type="text" class="form-control" id="strSldatoDiet" name="strSldatoDiet" value="" placeholder="<%=expence_txt_079 %>" readonly />
                                    <span class="input-group-addon input-small">
                                            <span class="fa fa-calendar">
                                            </span>
                                    </span>
                                </div>
                            </td>

                            <td><input type="time" id="FM_diet_sltime" name="FM_diet_sltime" class="form-control" /></td>
                        </tr>

                    </table>

                    <table style="width:100%; margin-top:10px;" id="dietTable">

                        <tr>
                            <td colspan="3" style="font-size:75%;"><i><b><%=expence_txt_072 %></b></i></td>
                        </tr>

                        <tr>
                            <td style="font-size:125%; width:100px;"><%=expence_txt_074 %></td>
                            <td style="width:75px;"><input readonly type="number" class="form-control" name="FM_morgenantal" id="antalmorgenmad" /></td>
                            <td style="font-size:125%;"><%=expence_txt_073 & " " %><%=morgenDiet %> kr.</td>
                            <input type="hidden" value="<%=morgenDiet %>" id="morgenmadprice" />
                        </tr>

                        <tr>
                            <td style="font-size:125%; width:100px;"><%=expence_txt_075 %></td>
                            <td style="width:75px;"><input readonly type="number" class="form-control" name="FM_frokostantal" id="antalfrokost" /></td> 
                            <td style="font-size:125%;"><%=expence_txt_073 & " " %> <%=middagDiet %> kr.</td>
                            <input type="hidden" value="<%=middagDiet %>" id="frokostprice" />
                        </tr>

                        <tr>
                            <td style="font-size:125%; width:100px;"><%=expence_txt_076 %></td>
                            <td style="width:75px;"><input readonly type="number" class="form-control" name="FM_aftenantal" id="antalaftensmad" /></td>
                            <td style="font-size:125%;"><%=expence_txt_073 & " " %> <%=aftendDiet %> kr.</td>
                            <input type="hidden" value="<%=aftendDiet %>" id="aftensmadprice" />
                        </tr>

                        <tr>
                            <td colspan="3" style="font-size:75%;"><i><b><%=expence_txt_081 %>: <span id="dietTotal"></span></b></i></td>
                        </tr>

                    </table>

                </div>
            </form>


            <%prisprovernatning = 300 %>
            <div id="overnatning" style="display:<%=lmdisplay%>;">
                <table style="width:100%; border:solid 1px #c9c9c9; margin-top:10px;" >
                    <tr>
                        <input type="hidden" id="prisprovernatning" value="<%=prisprovernatning %>" />
                        <td style="width:50%;"><input type="number" min="1" step="1" onkeypress='return event.charCode >= 48 && event.charCode <= 57' class="form-control" placeholder="<%=expence_txt_090 %>" id="antalovernatninger" /></td>
                        <td><%=expence_txt_091 %> 300 kr = <b id="overnatningerTotalPrice"></b></td>
                    </tr>
                </table>
            </div>


            <%if lto = "lm" then %>
            <br /><br />
            <%end if %>

            <div class="row">	
                <div class="col-lg-12"><%=expence_txt_001 %></div>	
            </div>

            <div id="expencestacker">
            <input type="hidden" id="lto" value="<%=lto %>" />
            <table style="width:100%; border:solid 1px #c9c9c9; margin-top:10px;" id="expenceTable_0">
                <tr>
                    <td colspan="2"><input type="text" name="FM_udlaeg_navn_0" id="FM_udlaeg_navn_0" placeholder="<%=expence_txt_067 %>" class="form-control"/></td>
                </tr>
                <tr>
                    <td>
                        <input type="number" name="FM_udlaeg_belob_0" id="FM_udlaeg_belob_0" placeholder="<%=expence_txt_007 %>" class="form-control" value="" />
                    </td>
                    <td>
                        <select class="form-control" name="FM_udlaeg_valuta_0" id="FM_udlaeg_valuta_0">
                            <%
                                strSQLValuta = "SELECT id, valutakode FROM valutaer"
                                oRec.open strSQLValuta, oConn, 3
                                while not oRec.EOF
                                %><option value="<%=oRec("id") %>" data-valutakode="<%=oRec("valutakode") %>" ><%=oRec("valutakode") %></option><%
                                oRec.movenext
                                wend
                                oRec.close
                            %>
                        </select>
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
                        <select class="form-control" name="FM_udlaeg_gruppe_0" id="FM_udlaeg_gruppe_0">
                            <%

                                select case lto
                                case "nt", "lm"
                                strSQLMatgrpSQLkri = " WHERE matgrp_konto <> 0"
                                case else
                                strSQLMatgrpSQLkri = ""
                                end select

                                select case lto
                                    case "lm"
                                        matidKri = " AND id <> 9" ' 9 = "Andet" og "Andet" skal sta til sidst
                                    case else
                                        matidKri = ""
                                end select

                                strSQLMatgrp = "SELECT id, navn FROM materiale_grp " & strSQLMatgrpSQLkri & " "& matidKri &" ORDER BY navn"
                                oRec.open strSQLMatgrp, oConn, 3
                                while not oRec.EOF
                                %><option value="<%=oRec("id") %>"><%=oRec("navn") %></option><%
                                oRec.movenext
                                wend
                                oRec.close

                                if lto = "lm" then
                                    %>
                                    <option value="9"><%=expence_txt_092 %></option>
                                    <%
                                end if

                            %>
                        </select>
                    </td>
                </tr>
                <%
                if lto <> "lm" then 
                %>
                <tr>
                    <td colspan="2">
                        <select class="form-control" name="FM_udlaeg_faktbar_0" id="FM_udlaeg_faktbar_0">

                            <%if lto = "nt" then %>

                                <option value="1" selected><%=expence_txt_070 %></option>
                                <option value="2"><%=expence_txt_013 & " " %>100 %</option>
                                <option value="5"><%=expence_txt_013 & " " %>50 %</option>

                            <%else %>

                                <option value="1"><%=expence_txt_070 %></option>
                                <option value="2"><%=expence_txt_013 %></option>

                            <%end if %>

                        </select>
                    </td>
                </tr>
                <%else %>
                <input type="hidden" name="FM_udlaeg_faktbar_0" id="FM_udlaeg_faktbar_0" value="1" />
                <%end if %>

                <tr>
                    <td colspan="2">
                        <select class="form-control" name="FM_udlaeg_form_0" id="FM_udlaeg_form_0">
                            <%select case lto
                              case "nt"%>
                              <option value="0"><%=expence_txt_012 %></option>
                              <option value="1"><%=expence_txt_042 %></option>
                             <%case else%>
                                <option value="1"><%=expence_txt_042 %></option>

                                <%if lto <> "lm" then %>
                                <option value="0"><%=expence_txt_012%></option> 
                                <%end if %>

                            <%end select %>

                            <%if lto = "lm" then %>
                                <option value="20">M.Card</option>
                                <option value="21">Workplus</option>
                                <option value="22">BroBizz</option>
                                <option value="23">Rejsekort</option>
                            <%end if %>

                        </select>
                    </td>
                </tr>

                <tr>
                    <td colspan="2" style="text-align:center;">
                        <%
                            'strSQL = "SELECT id FROM materiale_forbrug"
                            'oRec.open strSQL, oConn, 3 
                            'if not oRec.EOF then
                                'lastid = oRec("id")
                            'end if                         
                            'oRec.close

                            'matId = lastid '+ 1
                        %>

                        <form ENCTYPE="multipart/form-data" id="image_upload_0">

                            <label class="btn btn-default btn-lg"><INPUT id="0" NAME="fileupload1" TYPE="file" style="width:400px; display:none;" onchange="readURL(this);"><b> <%=expence_txt_014 %> </b></label>
                      
                        </form>

                    </td>
                </tr>

                <tr>
                    <td colspan="2" style="text-align:center;">
                        <img id="imageholder_0" src="<%=fileLink %>" alt='' border='0' style="width:50px;">
                    </td>
                </tr>


            </table>
            </div>

            <br />
            <div class="row">
                <div class="col-lg-6">
                    <span class="btn btn-default pull-left" id="addExpence">+</span>
                </div>
                 <div class="col-lg-6">
                    <span class="btn btn-default pull-right" id="removeExpence">-</span>
                </div>
            </div>
            

            <br /><br />


           <!-- <div style="border:1px solid black;">
                <div class="row">
                    <div class="col-lg-12"><input type="text" name="FM_udlaeg_navn" id="FM_udlaeg_navn" placeholder="<%=expence_txt_067 %>" class="form-control"/></div>
                </div>

                <div class="row">
                    <div class="col-lg-12">
                        <table style="width:100%">
                            <tr>
                                <td><input type="number" name="FM_udlaeg_belob" id="FM_udlaeg_belob" placeholder="<%=expence_txt_007 %>" class="form-control" value="" /></td>
                                <td style="padding-left:5px;">
                                    <select class="form-control" name="FM_udlaeg_valuta" id="FM_udlaeg_valuta">
                                    <%
                                        strSQLValuta = "SELECT id, valutakode FROM valutaer"
                                        oRec.open strSQLValuta, oConn, 3
                                        while not oRec.EOF
                                        %><option value="<%=oRec("id") %>" data-valutakode="<%=oRec("valutakode") %>" ><%=oRec("valutakode") %></option><%
                                        oRec.movenext
                                        wend
                                        oRec.close
                                    %>
                                </select>
                                </td>
                            </tr>
                        </table>
                    </div>
                </div>

                <div class="row">                                     
                    <div class="col-lg-12">
                        <select class="form-control" name="FM_udlaeg_gruppe" id="FM_udlaeg_gruppe">
                            <%

                                select case lto
                                case "nt", "lm"
                                strSQLMatgrpSQLkri = " WHERE matgrp_konto <> 0"
                                case else
                                strSQLMatgrpSQLkri = ""
                                end select

                                strSQLMatgrp = "SELECT id, navn FROM materiale_grp " & strSQLMatgrpSQLkri & " ORDER BY navn"
                                oRec.open strSQLMatgrp, oConn, 3
                                while not oRec.EOF
                                %><option value="<%=oRec("id") %>"><%=oRec("navn") %></option><%
                                oRec.movenext
                                wend
                                oRec.close
                            %>
                        </select>
                    </div>
                </div>

                <div class="row">                                     
                    <div class="col-lg-12">
                        <select class="form-control" name="FM_udlaeg_faktbar" id="FM_udlaeg_faktbar">
                            <option value="0"><%=expence_txt_070 %></option>
                        
                            <%select case lto 
                             case "nt" %>
                             <option value="1"><%=expence_txt_013 %> 100%</option>
                             <option value="5"><%=expence_txt_013 %> 50%</option>
                         
                            <%case else %>
                            <option value="1"><%=expence_txt_013 %></option>
                            <%end select %>
                        </select>
                    </div>
                </div>

                <div class="row">                                     
                    <div class="col-lg-12">
                        <select class="form-control" name="FM_udlaeg_form" id="FM_udlaeg_form">
                            <%select case lto
                              case "nt"%>
                              <option value="0"><%=expence_txt_012 %></option>
                              <option value="1"><%=expence_txt_042 %></option>
                             <%case else%>
                            <option value="1"><%=expence_txt_042 %></option>
                            <option value="0"><%=expence_txt_012%></option> 
                            <%end select %>
                        </select>
                    </div>
                </div>

                <%
                    strSQL = "SELECT id FROM materiale_forbrug"
                    oRec.open strSQL, oConn, 3 
                    while not oRec.EOF
                        lastid = oRec("id")
                    oRec.movenext
                    wend
                    oRec.close

                    matId = lastid + 1
                %>
                                        
                <div class="row">
                    <div class="col-lg-12" style="text-align:center;">
                        <form ENCTYPE="multipart/form-data" action="../timereg/upload_bin.asp?matUpload=1&matId=<%=matId %>&thisfile=timetag_mobile" method="post" id="image_upload">
                            <label class="btn btn-default btn-lg"><INPUT id="image-file" NAME="fileupload1" TYPE="file" style="width:400px; display:none;" onchange="readURL(this);"><b> <%=expence_txt_014 %> </b></label>
                        </form>
                    </div>
                </div>

                <div class="row">
                    <div class="col-lg-12" style="text-align:center;">
                        <img id="imageholder" src="<%=fileLink %>" alt='' border='0' style="width:50px;">
                    </div>
                </div>
            </div> -->

            <script type="text/javascript">
                function readURL(input) {
                    // alert("show pic")

                    thisid = input.id;

                    if (input.files && input.files[0]) {
                        var reader = new FileReader();

                        reader.onload = function (e) {
                            $('#imageholder_' + thisid)
                                .attr('src', e.target.result);
                        };

                        reader.readAsDataURL(input.files[0]);
                    }
                }
            </script>

            <br /><br />

            <div class="row">
                <div class="col-lg-12">
                    <button type="submit" id="sbmExpence" class="btn btn-success" style="text-align:center; width:100%"><b><%=ttw_txt_011 %> >></b></button>
                </div>
            </div>


            <!-- Uploader bar -->
            <style>
                /* The Modal (background) */
                .modal {
                    display: none; /* Hidden by default */
                    position: fixed; /* Stay in place */
                    z-index: 2; /* Sit on top */
                    padding-top: 100px; /* Location of the box */
                    left: 0;
                    top: 0;
                    width: 100%; /* Full width */
                    height: 100%; /* Full height */
                    overflow: auto; /* Enable scroll if needed */
                    background-color: rgb(0,0,0); /* Fallback color */
                    background-color: rgba(0,0,0,0.4); /* Black w/ opacity */
                }

                /* Modal Content */
                .modal-content {
                    background-color: #fefefe;
                    margin: auto;
                    padding: 20px;
                    border: 1px solid #888;
                    width: 350px;
                    height: 250px;
                }

                .ui-datepicker { left:250px; } 
 
            </style>

            <div id="uploadbar" class="modal">
                <div class="modal-content">
                    <div class="row">
                        <div class="col-lg-12" style="text-align:center;"><h3><%=expence_txt_093 %></h3></div>
                    </div>

                    <div class="row">
                        <div class="col-lg-12" style="text-align:center;">
                            <div style="width:100%; background-color:#ddd;">
                                <div style="width:1%; background-color:#4CAF50; height:30px;" id="progressbar"></div>
                            </div>
                        </div>
                    </div>

                </div>
            </div>



            <%if lto <> "lm" then %>
            <br /><br />
            <div class="row">
                <div class="col-lg-12" style="text-align:center"><b><%=expence_txt_065 %></b></div>
            </div>
            <div class="row">
                <div class="col-lg-12">
                    <table class="table">
                        <tr>
                            <th><%=expence_txt_039 %></th>
                            <th><%=expence_txt_006 %></th>
                            <th><%=expence_txt_007 %></th>
                            <th style="text-align:center;"><%=expence_txt_066 %></th>
                        </tr>

                        <%
                        strSQL = "SELECT mf.id as id, matnavn, matkobspris, mf.dato as dato, v.valutakode as valutakode FROM materiale_forbrug mf LEFT JOIN valutaer as v ON (mf.valuta = v.id) WHERE usrid = "& session("mid") & " AND godkendt = 0 AND afregnet = 0 AND mattype = 1 ORDER BY dato DESC LIMIT 5"
                        'response.Write strSQL
                        oRec.open strSQL, oConn, 3
                        ialt = 0
                        while not oRec.EOF
                            %>
                            <tr>
                                <td><%=oRec("dato") %></td>
                                <td><%=oRec("matnavn") %></td>
                                <td><%=formatnumber(oRec("matkobspris"), 2) &" "& oRec("valutakode") %></td>
                                <td style="text-align:center;"><a href="expence.asp?func=delexpence&matid=<%=oRec("id") %>&dato=<%=oRec("dato") %>&matnavn=<%=oRec("matnavn") %>&pris=<%=oRec("matkobspris") %>"><span style="color:darkred;" class="fa fa-times"></span></a></td>
                            </tr>
                            <%
                        ialt = ialt + 1
                        oRec.movenext
                        wend
                        oRec.close
                        %>

                    </table>
                </div>
            </div>
            <%else %>
            <br /> <br />
            <%end if %>
        </div>

    </div>
</div>
</body>

<!--#include file="../inc/regular/footer_inc.asp"-->
