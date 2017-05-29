
<%
    thisfile = "timetag_mobile" 
%>

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->

<%'**** Søgekriterier AJAX **'

        'section for ajax calls
        if Request.Form("AjaxUpdateField") = "true" then
        Select Case Request.Form("control")
        case "FN_geolocation"

             '*** Tilføjer longitude og latitude til seneste åbne loginf for dem der har stempelur slået til *********
            
             call erStempelurOn()
             
            longitude = request("jq_longitude")
            latitude = request("jq_latitude")

             if cint(stempelurOn) = 1 then
             strSQLlatestlogin = "SELECT id FROM login_historik WHERE mid =  " & session("mid") & " AND stempelurindstilling > 0 AND lh_latitude = 0 AND logud IS NULL ORDER BY id DESC LIMIT 1 "
			 oRec.open strSQLlatestlogin, oConn, 3
             if not oRec.EOF then	   

         
	                    strSQLAddLong = "UPDATE login_historik SET "_
	                    &" lh_longitude = "& longitude &", lh_latitude = "& latitude &""_
	                    &" WHERE id = "& oRec("id")
    				   
				       oConn.execute(strSQLAddLong)  
            
              
               end if
               oRec.close
              
             end if
        



        case "FN_akttyper"
    
    
    
        if len(trim(request("id"))) then
        id = request("id")
        else
        id = 0
        end if

        strSQLakttype = "SELECT fakturerbar FROM aktiviteter WHERE id = "& id

        oRec.open strSQLakttype, oConn, 3 
            if not oRec.EOF then
            akt_type = oRec("fakturerbar")
            end if
        oRec.close

        response.Write akt_type

        case "FN_jobbesk"
                
                if len(trim(request("id"))) then
                jobid = request("id")
                else
                jobid = 0
                end if


                jobbesk = "- Jobbeskrivelse ikke fundet"
                'jobnavn = "- Jobnavn ikke fundet"
                strSQLjobbesk = "SELECT jobnavn, beskrivelse FROM job WHERE id = "& jobid
                oRec.open strSQLjobbesk, oConn, 3 
                if not oRec.EOF then
        
                jobbesk = oRec("beskrivelse")
                'jobnavn = oRec("jobnavn")

                end if
                oRec.close   


                      '*** ÆØÅ **'
                    call jq_format(jobbesk)
                    jobbeskTxt = jq_formatTxt


                    response.write jobbeskTxt


        case "FN_tjktimer_forecast_tt" 
                '** FORECAST ALERT 
         
                

                aktid = request("aktid")
                timerTastet = request("timer_tastet")
                usemrn = request("treg_usemrn")
                ibudgetaar = request("ibudgetaar")
                ibudgetmd = request("ibudgetaarMd")  
                aar = request("ibudgetUseAar")
                md = request("ibudgetUseMd")


             
                'response.write "aktid: "+ aktid + " timerTastet: "+ timerTastet
                'response.end

                call ressourcefc_tjk(ibudgetaar, ibudgetmd, aar, md, usemrn, aktid, timerTastet)

                response.write feltTxtValFc


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

                select case ignJobogAktper
                case 0,1
                strSQLDatokri = " AND ((j.jobstartdato <= '"& varTjDatoUS_son &"' AND j.jobstatus = 1) OR (j.jobstatus = 3))"
                case 3
                strSQLDatokri = " AND ((j.jobstartdato <= '"& varTjDatoUS_son &"' AND j.jobslutdato >= '"& varTjDatoUS_man &"' AND j.jobstatus = 1) OR (j.jobstatus = 3))"
                case else
                strSQLDatokri = ""
                end select



                if filterVal <> 0 then
            
                 if jobkundesog <> "-1" then

                    if jobkundesog = "all" then 
                        strSQLSogKri = " AND (jobnr <> '' AND kkundenavn <> '') "
                    else
                        strSQLSogKri = " AND (jobnr LIKE '"& jobkundesog &"%' OR jobnavn LIKE '%"& jobkundesog &"%' OR "_
                        &" kkundenavn LIKE '"& jobkundesog &"%' OR kkundenr = '"& jobkundesog &"' OR k.kinit = '"& jobkundesog &"') AND kkundenavn <> ''"
                    end if            

                 lmt = 50
                 else
                 strSQLSogKri = ""
                 lmt = 250
                 end if            


                lastKid = 0
                
                if jobkundesog <> "-1" then
                strJobogKunderTxt = strJobogKunderTxt &"<span style=""color:red; font-size:9px; float:right;"" id=""luk_jobsog"">[X]</span>"    
                strJobogKunderTxt = strJobogKunderTxt &"<ul>"
                end if


                strSQL = "SELECT j.id AS jid, j.jobnavn, j.jobnr, j.jobstatus, k.kkundenavn, k.kkundenr, k.kid, j.jobstartdato FROM timereg_usejob AS tu "_ 
                &" LEFT JOIN job AS j ON (j.id = tu.jobid) "_
                &" LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "_
                &" WHERE tu.medarb = "& medid &" AND (j.jobstatus = 1 OR j.jobstatus = 3) "& strSQLPAkri &" "& strSQLDatokri 
            
                strSQL = strSQL & strSQLSogKri

                if lto = "hestia" then 
                strSQL = strSQL &" GROUP BY j.id ORDER BY jobnavn LIMIT "& lmt     
                else
                strSQL = strSQL &" GROUP BY j.id ORDER BY kkundenavn, jobnavn LIMIT "& lmt     
                end if
              
                'if lto = "essens" then
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
                         strJobogKunderTxt = strJobogKunderTxt & "<li class=""span_job"">Ingen kunder/job fundet</li>"
                         else
                         strJobogKunderTxt = strJobogKunderTxt & "<option value=""-1"" DISABLED>Ingen kunder/job fundet</option>"
                         end if
                    end if
       
                

                    '*** ÆØÅ **'
                    call jq_format(strJobogKunderTxt)
                    strJobogKunderTxt = jq_formatTxt


                    response.write strJobogKunderTxt

                end if    


    
    
        
          case "FN_sogakt"

               
                '*** Søg Aktiviteter 
                

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


                '*** Forecast tjk
                risiko = 0
                strSQLjobRisisko = "SELECT j.risiko FROM job AS j WHERE id = "& jobid
                oRec5.open strSQLjobRisisko, oConn, 3
                if not oRec5.EOF then
                risiko = oRec5("risiko")
                end if
                oRec5.close 

                call allejobaktmedFC(viskunForecast, medid, jobid, risiko)

                
                '*** Datospærring Vis først job når stdato er oprindet
                call lukaktvdato_fn()
                ignJobogAktper = lukaktvdato

               
	            if (cint(ignJobogAktper) = 1 OR cint(ignJobogAktper) = 2 OR cint(ignJobogAktper) = 3) then
	            strSQLDatoKri = " AND ((a.aktstartdato <= '"& varTjDatoUS_son &"' AND a.aktslutdato >= '"& varTjDatoUS_man &"') OR (a.fakturerbar = 6))" 
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




               if cint(pa) = 1 then '**Kun på Personlig aktliste
    
    
                   'Positiv aktivering
                   if cint(pa_only_specifikke_akt) then

                   strSQL = "SELECT a.id AS aid, navn AS aktnavn "_
                   &" FROM timereg_usejob AS tu LEFT JOIN aktiviteter AS a ON (a.id = tu.aktid) "_
                   &" WHERE tu.medarb = "& medid &" AND tu.jobid = "& jobid &" AND aktid <> 0 "& strSQlAktSog &" AND aktstatus = 1 AND ("& aty_sql_hide_on_treg &") "& forecastAktids &" AND a.navn IS NOT NULL ORDER BY fase, sortorder, navn LIMIT 150"   
                   'AND ("& replace(aty_sql_realhours, "tfaktim", "a.fakturerbar") &")
                   else 

                   strSQL = "SELECT a.id AS aid, navn AS aktnavn "_
                   &" FROM timereg_usejob AS tu LEFT JOIN aktiviteter AS a ON (a.job = tu.jobid) "_
                   &" WHERE tu.medarb = "& medid &" AND tu.jobid = "& jobid &" AND aktid = 0 "& strSQlAktSog &" AND aktstatus = 1  AND ("& aty_sql_hide_on_treg &") "& forecastAktids &" AND a.navn IS NOT NULL ORDER BY fase, sortorder, navn LIMIT 150" 
                   'AND ("& replace(aty_sql_realhours, "tfaktim", "a.fakturerbar") &")
                   end if


                  

               else


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


            
                   
               
            

               strSQL = "SELECT a.id AS aid, navn AS aktnavn, projektgruppe1, projektgruppe2, projektgruppe3, "_
               &" projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM aktiviteter AS a "_
               &" WHERE a.job = " & jobid & " AND a.job <> 0 "& strSQlAktSog &" AND aktstatus = 1 AND ("& aty_sql_hide_on_treg &") ORDER BY fase, sortorder, navn LIMIT 150"      
    

              
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

                if cint(pa) = 1 then 'Positiv aktivering

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
                        fcsaldo_txt = "<span style=""font-weight:lighter; font-size:9px;""> (fc. Saldo: "& formatnumber(feltTxtValFc, 2) & " / "& formatnumber(fctimer,2) &" t.)</span>"
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

          


           case "FN_sogmat"


                

                '*** SØG Mat            
                
                if len(trim(request("jq_newfilterval"))) <> 0 then
                filterVal = 1 
                jobkundesog = request("jq_newfilterval")
                else
                filterVal = 0
                jobkundesog = "6xxxxxfsdf554"
                end if
        
                

                if filterVal <> 0 then
            
                
                
                strMat = strMat &"<span style=""color:darkred; font-size:11px; float:right;"" id=""luk_matsog"">X</span>"    
                         
                strMat = strMat &"<ul>"

                strSQL = "SELECT m.navn AS matnavn, m.id AS matid, m.matgrp, m.varenr, m.enhed, mg.navn AS grpnavn, mg.id AS mgid, m.salgspris FROM materialer AS m "_
                &" LEFT JOIN materiale_grp AS mg ON (mg.id = m.id) WHERE m.id <> 0 AND (m.navn LIKE '%"& jobkundesog &"%') ORDER BY m.navn LIMIT 100"   
    

                'response.write strSQL
                'response.end            

                c = 0
                oRec.open strSQL, oConn, 3
                while not oRec.EOF
        
                if lastgrp <> oRec("grpnavn") then
                
                    if c <> 0 then
                    strMat = strMat &"<br>"
                    'strMat = strMat & "<option value=""0"" DISABLED></option>"
                    end if            

                strMat = strMat &"<b><u>"& oRec("grpnavn") &"</b></u><br>"
                'strMat = strMat & "<option value=""0"" DISABLED>"& oRec("grpnavn") &"</option>"
                end if 
                 
                strMat = strMat & "<li class=""span_mat"" id=""chbox_mat_"& oRec("matid") &"""><input type=""hidden"" id=""hiddn_mat_"& oRec("matid") &""" value="""& oRec("matnavn") &""">"
                strMat = strMat & "<input type=""hidden"" id=""hiddn_matid_"& oRec("matid") &""" value="& oRec("matid") &"> "& oRec("matnavn") &" ("& oRec("enhed") &") </li>" 
                strMat = strMat & "<input type=""hidden"" id=""hiddn_matid_salgspris_"& oRec("matid") &""" value="& oRec("salgspris") &">"

                'strMat = strMat & "<option value="& oRec("matid") &">"& oRec("matnavn") &"</option>"

                
                lastgrp = oRec("mgid") 
                c = c + 1
                oRec.movenext
                wend
                oRec.close

               
    
              
                    if c = 0 then
                    strMat = strMat & "<li class=""span_mat"">Ingen materialer fundet</li>"
                    'strMat = strMat & "<option value=""-1"" DISABLED>Ingen materialer fundet</option>" 
                    end if 

                    strMat = strMat &"</ul>"

                    '*** ÆØÅ **'
                    call jq_format(strMat)
                    strMat = jq_formatTxt


                    response.write strMat

                end if    

        
      
          case "FN_tomobjid"


                if len(trim(request("jq_jobid"))) <> 0 then
                jq_jobid = request("jq_jobid")
                else
                jq_jobid = 0
                end if

                
                strJobnavn = ""
                strSQL = "SELECT jobnavn, jobnr FROM job WHERE id = "& jq_jobid
                

                oRec.open strSQL, oConn, 3
                if not oRec.EOF then
        
                strJobnavn = oRec("jobnavn") & " ("& oRec("jobnr") &")"           

                end if
                oRec.close
                
            
                  '*** ÆØÅ **'
                    call jq_format(strJobnavn)
                    strJobnavn = jq_formatTxt


                    response.write strJobnavn


        end select
        response.end
        end if

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

    
    <!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
    <!--#include file="../inc/regular/top_menu_mobile.asp"-->

<% 

    

    if len(trim(request("flushsessionuser"))) <> 0 then ' flushsessionuser = kommer fra logindsiden
     flushsessionuser = 1
    else
     flushsessionuser = 0
    end if

    '*** Hvis der ligger coookie, skal telefonen blive på / flushsessionuser = kommer fra logind ****'
    if request.Cookies("timeouttimeoutcloud")("mobileuser") <> "" AND cint(flushsessionuser) <> 1 then

    session("mid") = request.Cookies("timeoutcloud")("mobilemid")
    session("user") = request.Cookies("timeoutcloud")("mobileuser")
    session("rettigheder") = request.Cookies("timeoutcloud")("mobilerettigheder")

    end if




    if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/errors/error_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	response.End
    end if

    call meStamdata(session("mid"))

    '** Special settings medarbejdertype
    mt_mobil_visstopur = 0
    strSQLmeType = "SELECT mt_mobil_visstopur FROM medarbejdertyper WHERE id = " & meType
    oRec4.open strSQLmeType, oConn, 3
    if not oRec4.EOF then
        
    mt_mobil_visstopur = oRec4("mt_mobil_visstopur")

    end if 
    oRec4.close


    '** Select eller søgeboks
    call mobil_week_reg_dd_fn()
    'mobil_week_reg_job_dd = mobil_week_reg_job_dd
    'mobil_week_reg_akt_dd = 1
    'mobil_week_reg_akt_dd = mobil_week_reg_akt_dd    

    call positiv_aktivering_akt_fn()
     '*** Vis kun aktiviteter med forecast på
    call aktBudgettjkOn_fn()
    '*** Skal akt lukkes for timereg. hvis forecast budget er overskrddet..?
    '** MAKS budget / Forecast incl. peridoe afgrænsning
    call akt_maksbudget_treg_fn()

    call erStempelurOn()

    if len(trim(request("indlast"))) <> 0 AND session("timetag_web_indMsgShown") <> "1" then
    indlast = 1
    else
    indlast = 0
    end if

    '** Frvalgt job eller job fra mail
    select case lto
    case "tbg"
    tomobjid = 4
    case else
    tomobjid = session("tomobjid")
    end select

    select case lto
    case "hestia", "xoutz", "micmatic", "intranet - local"
        showAfslutJob = 1
        showMatreg = 1
        showStop = 0
        showDetailDayResumeOrLink = 0
        ststopbtnTxt = "St. / Stop"
    case "xintranet - local", "sdutek", "nonstop", "cc" ', "epi", "epi_uk", "epi_ab", "epi_no"
        showAfslutJob = 0
        showMatreg = 0
        showStop = 1
        showDetailDayResumeOrLink = 1
        ststopbtnTxt = "St. / Stop"
    case "epi2017"
        showAfslutJob = 0
        showMatreg = 0
        
        if cint(mt_mobil_visstopur) = 1 then
        showStop = 1
        mobil_week_reg_akt_dd = 1
        mobil_week_reg_job_dd = 1
        mobil_week_reg_akt_dd_forvalgt = 1
        else
        showStop = 0
        end if
        
        showDetailDayResumeOrLink = 1
        ststopbtnTxt = "St. / Stop"
    case "tbg", "hidalgo"
        showAfslutJob = 0
        showMatreg = 1
        showStop = 0
        showDetailDayResumeOrLink = 0
        ststopbtnTxt = "St. / Stop"
    case "eniga"
        showAfslutJob = 0
        showMatreg = 0
        showStop = 1
        showDetailDayResumeOrLink = 0
        ststopbtnTxt = "St. / Stop"
    case else
        showAfslutJob = 0
        showMatreg = 0
        showStop = 0
        showDetailDayResumeOrLink = 0
        ststopbtnTxt = "St. / Stop"
    end select


   


    '** Hvis st_stop er vagt på medarbejder overruler det standard settings ovenfor
    if cint(timer_ststop) = 1 then
        showStop = 1
    else
        showStop = showStop
    end if


    ddDato = year(now) &"/"& month(now) &"/"& day(now)
    ddDato_ugedag = day(now) &"/"& month(now) &"/"& year(now)
    ddDato_ugedag_w = datepart("w", ddDato_ugedag, 2, 2)
            
    varTjDatoUS_man_tt = dateAdd("d", -(ddDato_ugedag_w-1), ddDato_ugedag)

%>

<head>

<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>TimeOut mobile</title>



<script src="js/timetag_web_jav_2017_023.js" type="text/javascript"></script>




    <style type="text/css">

        input[type="text"] 
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
    

    <!--
     <button onclick="getLocation()">Try It</button>
    <p id="demo"></p>
        -->


       


        <%call mobile_header %>


           
        <div class="container" style="height:100%">
            <div class="portlet">
                <div class="portlet-body">
                    
                    <div id="dvindlaes_msg" style="position:absolute; top:0px; left:0px; height:100%; width:100%; background-color:#cccccc; visibility:hidden; display:none;"><%=ttw_txt_002 %></div>
                   
                    <%if cint(indlast) = 1 then %>
                    <div class="row">
                                <div class="col-lg-12">
                                     <div id="timer_indlast" style="text-align:center; background-color:greenyellow; padding:4px;"><%=ttw_txt_003 %></div>
                                    </div>
                    </div>
                    <%session("timetag_web_indMsgShown") = "1"
                    end if %>

                    <form id="container" action="../timereg/timereg_akt_2006.asp?func=db&rdir=timetag_web" method="post">

                        
                        <input type="hidden" id="Hidden5" name="year" value="<%=year(now) %>"/>       
                        <!--<input type="hidden" id="Hidden3" name="FM_dager" value=""/>-->
                        <input type="hidden" id="Hidden4" name="FM_dager" value="0"/>
                        <input type="hidden" id="Hidden2" name="FM_feltnr" value="0"/>
                        <input type="hidden" id="FM_pa" name="FM_pa" value="<%=pa_aktlist %>"/>
                        <input type="hidden" id="FM_medid" name="FM_medid" value="<%=session("mid") %>"/>
                        <input type="hidden" id="FM_medid_k" name="FM_medid_k" value="<%=session("mid") %>"/>
                        <input type="hidden" id="jq_lto" name="jq_lto" value="<%=lto %>"/>
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
                              %> <input type="hidden" id="jq_dato" name="FM_datoer" value="<%=todayDay &"-"& todayMonth &"-"& todayYear%>"/><%
                        case else
                       
                        %>                        
                        <div class="row">
                            <div class="col-lg-12"><input type="text" name="FM_datoer" id="jq_dato" value="<%=todayDay &"-"& todayMonth &"-"& todayYear%>" class="form-control" /> <!-- placeholder="dd-mm-yyyy" --></div>
                        </div>
                        <%end select %>     
                        
                        
                         
                        <input type="hidden" id="tomobjid" name="tomobjid" value="<%=tomobjid %>"/>
                        <input type="hidden" id="showstop" name="showstop" value="<%=showStop %>"/>
                        <input type="hidden" id="" name="FM_vistimereltid" value="<%=showStop %>"/>
                        <input type="hidden" id="varTjDatoUS_man" name="varTjDatoUS_man" value="<%=varTjDatoUS_man_tt %>"/>

                        <input type="hidden" id="mobil_week_reg_akt_dd" name="" value="<%=mobil_week_reg_akt_dd %>"/>
                        <input type="hidden" id="mobil_week_reg_job_dd" name="" value="<%=mobil_week_reg_job_dd %>"/> 

                        <input type="hidden" id="jq_longitude" name="" value="0"/>
                        <input type="hidden" id="jq_latitude" name="" value="0"/>
                        <input type="hidden" id="jq_stempeluron" name="" value="<%=stempelurOn %>"/>  
                        
                                              
                        <!-- Forecast max felter -->
                        <input type="hidden" id="aktnotificer_fc" name="" value="0"/>
                        <input type="hidden" id="akt_maksbudget_treg" value="<%=akt_maksbudget_treg%>">  
                        <input type="hidden" id="akt_maksforecast_treg" value="<%=akt_maksforecast_treg%>">
                        <input type="hidden" id="aktBudgettjkOn_afgr" value="<%=aktBudgettjkOn_afgr%>">
                        <input type="hidden" id="regskabsaarStMd" value="<%=datePart("m", aktBudgettjkOnRegAarSt, 2,2)%>">
                        <input type="hidden" id="regskabsaarUseAar" value="<%=datepart("yyyy", varTjDatoUS_man_tt, 2,2)%>">
                        <input type="hidden" id="regskabsaarUseMd" value="<%=datepart("m", varTjDatoUS_man_tt, 2,2)%>">
                        <input type="hidden" id="dv_akttype" value="0">


                         
             
        

            

                        <%
                        '*** GL 
                        if cint(mobil_week_reg_job_dd) = 1 then %>
                         <div class="row">
                                <div class="col-lg-12">
                                <input type="hidden" id="FM_job" value="-1"/>
                                <select id="dv_job" name="FM_jobid" class="form-control">
                                 
                                </select>
                                </div>
                           </div>
                        <%else %>
                         <div class="row">
                            <div class="col-lg-12">
                                <input type="text" id="FM_job" name="FM_job" value="" placeholder="<%=ttw_txt_014 %>" class="form-control"/>
                                <input type="hidden" id="FM_jobid" name="FM_jobid" value="0"/>
                                <div id="dv_job" style="padding:5px 5px 5px 5px; display:none; visibility:hidden;"></div> 
                            </div>
                        </div>
                        <%end if%>

                   


                        





                        <%select case lto
                        case "hestia", "intranet - local"
                            %><br />
                             <div class="panel-group accordion-panel" id="accordion-paneled1">
                            <div class="panel panel-default">
                                <div class="panel-heading">
                                    <h4 class="panel-title"><a class="accordion-toggle" data-toggle="collapse" data-target="#collapseTwo" id="dv_jobbesk_header"><%=ttw_txt_022 %></a></h4></div>                                    
                                <div id="collapseTwo" class="panel-collapse collapse">
                                    <div class="panel-body">
                                         <div class="row">
                                                <div class="col-lg-12" id="dv_jobbesk" style="padding:5px 5px 5px 25px;">
                                               <!-- <a href="#" id="jq_jobbesk" target="_blank">Jobbeskrivelse +</a>-->
                                                </div>
                                         </div>
                                     </div>
                                 </div>
                                 </div>
                            </div>
                            </div>
                            <%
                        end select%>

                    

                       
                       

                        <%if cint(mobil_week_reg_akt_dd) = 1 then %>

                            

                            <div class="row">
                                    <div class="col-lg-12">
                                        <input type="hidden" id="FM_akt" value="-1"/>
                                        <!--<textarea id="dv_akt_test"></textarea>-->
                                        <select id="dv_akt" name="FM_aktivitetid" class="form-control">
                                            <option>..</option>
                                        </select>
                                    </div>
                             </div>
                         <%else %>
                        <div class="row">
                            <div class="col-lg-12">
                                <input type="text" id="FM_akt" name="activity" value="" class="form-control" placeholder="<%=ttw_txt_004 %>"/>
                                <input type="hidden" id="FM_aktid" name="FM_aktivitetid" value="0"/>
                                <div id="dv_akt" class="dv-closed" style="padding:5px 5px 5px 5px;"></div> 
                            </div>
                        </div>            
                        <!-- dv_akt -->
                        <%end if %>

                       

                        <div class="row">
                            <div class="col-lg-12"><input type="text" id="FM_kom" name="FM_kom_0" placeholder="<%=ttw_txt_007 %>" class="form-control" /></div>
                        </div>



                        <%if cint(showStop) = 1 then%>
                           
                        <div class="row">
                            <div class="col-lg-12">
                            <table>
                                <tr>
                                    <td><input type="button" id="bt_stst" value="<%=ststopBtnTxt %>" class="btn btn-secondary btn-sm"/></td>
                                    <td>&nbsp</td>
                                    <td><input type="text" id="FM_sttid" name="FM_sttid" value="00:00" style="color:#cccccc;" class="form-control"/></td>
                                    <td>&nbsp</td>
                                    <td><input type="text" id="FM_sltid" name="FM_sltid" value="00:00" style="color:#cccccc;" class="form-control"/></td>
                                    <td>&nbsp</td>
                                    <td style="width:15%; text-align:right">= <span id="FM_timerlbl">0</span></td>
                                    <input type="hidden" id="FM_timer" name="FM_timer" value="0" style="color:#cccccc; width:65px;"/>
                                </tr>
                            </table>
                            </div>
                        </div>
                        <%else %>

                        

                            <input type="hidden" id="FM_sttid" name="FM_sttid" value="00:00"/>
                            <input type="hidden" id="FM_sltid" name="FM_sltid" value="00:00"/>
                            <div class="row">
                                <div class="col-lg-12"><input type="text" id="FM_timer" name="FM_timer" value="" placeholder="<%=ttw_txt_023 %>" class="form-control"/></div>
                            </div>
                        <%end if %>
                       

                       

                        
                        <%if cint(showMatreg) = 1 then
                            
                            if lto = "tbg" then
                            ttw_txt_008 = "Products"
                            ttw_txt_009 = "Add Products"
                            end if
                            
                            %>
                         <br />
                        <div class="panel-group accordion-panel" id="accordion-paneled">
                            <div class="panel panel-default">
                                <div class="panel-heading">
                                    <h4 class="panel-title"><a class="accordion-toggle" data-toggle="collapse" data-target="#collapseOne"><%=ttw_txt_008 %></a></h4></div>                                    
                                <div id="collapseOne" class="panel-collapse collapse">
                                    <div class="panel-body">
                                         <div class="row">
                                          <div class="col-lg-12" style="padding:0px 2px 0px 4px;">
                                          <table style="width:100%;"><tr><td style="width:60%;">    
                                          <input type="text" id="FM_matnavn" name="FM_matnavn" placeholder="<%=ttw_txt_009 %>" class="form-control"/></td>
                                          <%if lto = "tbg" OR lto = "intranet - local" then %>
                                              <td style="width:20%; padding:0px 2px 0px 2px;"><input type="number" id="FM_matantal_stkpris" name="FM_matantal_stkpris" value="" placeholder="Price" class="form-control"/></td>
                                              <td style="width:20%; padding:0px 2px 0px 2px;"><input type="number" id="FM_matantal" name="FM_matantal" value="" placeholder="<%=ttw_txt_010 %>." class="form-control"/></td>
                                          
                                              </tr>  
                                              <td colspan="3" align="right" style="padding:0px 4px 0px 2px;"><input type="button" value=">>" id="sbmmat" class="btn btn-secondary"/>
                                           </td></tr>
                                            <%else %>
                                              <td style="width:25%; padding:0px 4px 0px 4px;"><input type="number" id="FM_matantal" name="FM_matantal" value="" placeholder="<%=ttw_txt_010 %>." class="form-control"/></td>
                                          
                                               <td><input type="button" value=">>" id="sbmmat" class="btn btn-secondary"/>
                                                </td></tr>
                                              <%end if %>
                                         </table>
                                              </div>
                                        </div>
                                        

                                             <div class="row">
                                                  <div class="col-lg-4" style="padding:5px 2px 5px 2px;">
                                                      
                                                                 <div id="dv_mat" class="dv-closed" style="padding:5px 2px 5px 2px;"></div>
                                                            
                                                      </div>
                                             </div>

                                              
                                             <div class="row">
                                                  <div class="col-lg-8">
                                                    <div id="dv_mat_sbm" style="position:relative; border:0px; padding-left:10px; height:40px; overflow:auto;"></div>
                                                      </div>
                                               </div>  

                                        </div>



                                        <input type="hidden" id="FM_matid" name="FM_matid" value="0"/>
                                       
                                        <input type="hidden" id="FM_matids" name="FM_matids" value=""/>
                                        <input type="hidden" id="FM_matantals" name="FM_matantals" value=""/>
                                        <input type="hidden" id="FM_matantals_stkpris" name="FM_matantals_stkpris" value=""/>
                                    </div>
                                </div>
                            </div>
                        </div>

                        <%end if %>


                        <%if cint(showAfslutJob) = 1 then %>
                        <div class="row">
                            <div class="col-lg-12"><span><input type="checkbox" value="2" name="FM_lukjobstatus" id="FM_lukjobstatus" /> <%=ttw_txt_012 %></span></div>
                        </div>
                        <%end if %>

                        <br /><br />
                        <div class="row">
                            <div class="col-lg-12">
                                <button type="submit" id="sbmtimer" class="btn btn-success" style="text-align:center; width:100%"><b><%=ttw_txt_011 %> >></b></button>
                            </div>
                        </div>

                        <br />

                        <%
                            call akttyper2009(2)

             
           
            
                            timerIdagTxt = ""
                            timerIdag = 0

            
                            if cint(showDetailDayResumeOrLink) = 1 then

                             timerIdagTxt = "<table cellpadding=1 cellspacing=2 border=0 width=""100%"">"

                             timerIdagTxt = timerIdagTxt & "<tr><td colspan=4><b>"& ttw_txt_007 &"</b></td></tr>"
           
            
                             strSQLtimer = "SELECT timer, tjobnavn, taktivitetnavn, sttid, sltid, timer FROM timer WHERE tmnr = "& session("mid") &" AND tdato = '"& ddDato &"' AND ("& aty_sql_realhours &")"
            
                            oRec.open strSQLtimer, oConn, 3
                            while not oRec.EOF 

                            if cint(showStop) = 1 then
                            timerIdagTxt = timerIdagTxt & "<tr><td>"& left(formatdatetime(oRec("sttid"), 3), 5) & " - "& left(formatdatetime(oRec("sltid"), 3), 5) &"</td><td>"& left(oRec("tjobnavn"), 15) &"</td><td>"& left(oRec("taktivitetnavn"), 10) &"</td><td align=right>"&  formatnumber(oRec("timer"), 2) & "</td></tr>"
                            else
                            timerIdagTxt = timerIdagTxt & "<tr><td>"& oRec("timer") &"</td><td>"& left(oRec("tjobnavn"), 15) &"</td><td>"& left(oRec("taktivitetnavn"), 10) &"</td><td align=right>"&  formatnumber(oRec("timer"), 2) & "</td></tr>"
                            end if


                            timerIdag = timerIdag + oRec("timer")
                            oRec.movenext
                            wend
                            oRec.close 

           


                             if len(trim(timerIdag)) <> 0 then
                            timerIdag = timerIdag
                            else
                            timerIdag = 0 
                            end if

                            timerIdag = formatnumber(timerIdag, 2)

                            timerIdagTxt = timerIdagTxt & "<tr><td colspan=4 align=right><b><u>"& timerIdag &"</u></b></td></tr>"

                            timerIdagTxt = timerIdagTxt & "</table>"

                            %> 
                                 <div class="row">
                                <div class="col-lg-12">
                                    <div id="dv_timeridag" style="font-size:12px; float:left; text-align:left; width:100%; border:0px; padding:5px 5px 5px 5px; color:#999999; white-space:nowrap;">
                                    <%=meNavn%><br /> 
                                    <%=timerIdagTxt %>
                                    </div>
                                </div>
                                 </div>
                            <%

                            else

            
                            strSQLtimer = "SELECT SUM(timer) AS timer FROM timer WHERE tmnr = "& session("mid") &" AND tdato = '"& ddDato &"' AND ("& aty_sql_realhours &")"
            
                            oRec.open strSQLtimer, oConn, 3
                            if not oRec.EOF then

                            timerIdag = oRec("timer")

                            end if
                            oRec.close 

                            if len(trim(timerIdag)) <> 0 then
                            timerIdag = timerIdag
                            else
                            timerIdag = 0 
                            end if

                            timerIdag = formatnumber(timerIdag, 2)

                            %>
                            <div class="row">
                                <div class="col-lg-12">
                            <div id="dv_timeridag" style="width:200px; float:right; text-align:right; border:0px; color:#999999;"><span style="font-size:12px; line-height:13px; vertical-align:baseline; color:#999999;"><%=meNavn &" d. "& formatdatetime(now, 2)  %><br /></span> <a href="../to_2015/ugeseddel_2011.asp?usemrn=<%=session("mid")%>&varTjDatoUS_man=<%=varTjDatoUS_man_tt %>" style="font-size:32px;"><%=timerIdag%></a> </div>
                            </div>
                                  </div>  
                            <br /><br />&nbsp;
                            <%

                            end if

           
            
                            %>


       
                        <!--
                        <span style="font-size:10px; color:#999999;">[<%=lto %>:<%=meNavn %>]</span>
                        -->

                        <!--
                        <input type="submit" class="inactive" value="Tilføj timer & materialer"/>
                        -->



                    </form>

                    
                       
                

                </div>
            </div>
        </div>

</div>
</div>


<!--#include file="../inc/regular/footer_inc.asp"-->