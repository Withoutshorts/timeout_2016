<!--#include file="../inc/connection/conn_db_inc.asp"-->

     
<%
 '**** S�gekriterier AJAX **'
        'section for ajax calls

        if Request.Form("AjaxUpdateField") = "true" then
        Select Case Request.Form("control")


            case "FN_nyhederlaest"

            sqlDTD = year(now) &"-"& month(now) &"-"& day(now)
            strSQL = "SELECT id FROM info_screen WHERE ('"& sqlDTD &"' >= datofra AND '"& sqlDTD &"' <= datotil) AND vigtig = 1 ORDER BY vigtig DESC, id DESC"
            oRec.open strSQL, oConn, 3 
            while not oRec.EOF

                oConn.execute("INSERT INTO news_rel SET medarbid = "& session("mid") &", newsid = "& oRec("id") &", newsread = 1")

            oRec.movenext
            wend
            oRec.close



      '*************************************************************************
        '**** K�rsel ***'
        '*************************************************************************
        case "FM_get_destinations"

                                    


            	if len(trim(request("kid"))) <> 0 then
				kid = request("kid") 
				else
				kid = 0
				end if
				
				
				
				
				    if len(trim(request("visalle"))) <> 0 then
				    visalle = request("visalle")
				    else
				    visalle = 0
				    end if
				    
				 
                                    if len(trim(request("jqmid"))) <> 0 then
                                    jqmrn = request("jqmid")
                                    else
                                    jqmrn = 0
                                    end if
				
			                       strFil_Kon_Txt = ""

                    '**** Valgt medarbejer *****'
                    strSQLmnradr = "SELECT mnavn, mnr, init, madr, mpostnr, mcity, mland FROM medarbejdere WHERE mid = "& jqmrn 
                    'Response.write "<option>"& strSQLmnradr &"</option>"
                    
                                    
                     oRec2.open strSQLmnradr, oConn, 3
                    if not oRec2.EOF then
                    
                   
                    
                                strFil_Kon_Txt = strFil_Kon_Txt & "<option value='0'>"& oRec2("mnavn") &" "_
                                & left(oRec2("madr"), 30) &" "_
                                & oRec2("mpostnr") &" "& oRec2("mcity") & " "& oRec2("mland")
                                strFil_Kon_Txt = strFil_Kon_Txt &"</option>"
                           

                    end if
                    oRec2.close
                                   

                  


                        strSQLklicens = "SELECT k.kid, k.kkundenavn, k.adresse, k.postnr, k.city, k.land FROM kunder AS k WHERE useasfak = 1"
                        oRec.open strSQLklicens, oConn, 3
                        if not oRec.EOF then
                        
                                if len(trim(oRec("kkundenavn"))) <> 0 then
                                
                                strFil_Kon_Txt = strFil_Kon_Txt & "<option value='k_"& oRec("kid") &"'>"& left(oRec("kkundenavn"), 50) &" "_
                                & left(oRec("adresse"), 30) &" "_
                                & oRec("postnr") &" "& oRec("city") &" "_
                                & oRec("land")
                                strFil_Kon_Txt = strFil_Kon_Txt & "</option>"

                                end if

                        end if
                        oRec.close
                      
                            
    
                            

                          
                strSQLkperskm = "SELECT kp.id AS kpid, kp.navn AS kpersnavn, kp.email, kp.dirtlf, kp.mobiltlf, "_
				&" kp.adresse As kpadr, kp.postnr As kpzip, kp.town As kptown, kp.land As kpland, kp.afdeling AS kpafd, "_
				&" k.kid, k.kkundenavn, k.kkundenr, k.adresse, k.postnr, k.city, k.telefon, k.land FROM kunder k "_ 
				&" LEFT JOIN kontaktpers kp ON (kp.kundeid = k.kid AND kp.navn IS NOT NULL AND kp.adresse IS NOT NULL AND kp.postnr IS NOT NULL AND kp.town IS NOT NULL)"_
				&" WHERE useasfak <> 1 "_
                &" AND k.kkundenavn IS NOT NULL AND k.adresse IS NOT NULL ORDER BY k.kkundenavn, kp.navn"
			   
			    'Response.Write "<option>"& strSQLkperskm & "</option>"
			    'Response.end
			   
			
                lastKid = 0
                lastAlfabet = ""
			   
			    oRec2.open strSQLkperskm, oConn, 3
                x = xvalbegin
                k = 0
                while not oRec2.EOF
                
                   'strFil_Kon_Txt = strFil_Kon_Txt & "<option>"& oRec2("kkundenavn") &" </option>"
                       
                                 if lastAlfabet <> left(oRec2("kkundenavn"), 1) then
                                 strFil_Kon_Txt = strFil_Kon_Txt & "<option DISABLED></option>"
                                 strFil_Kon_Txt = strFil_Kon_Txt & "<option DISABLED></option>"
                                 strFil_Kon_Txt = strFil_Kon_Txt & "<option DISABLED style=""font-size:20px; background-color:#F7F7F7;"">"& left(ucase(oRec2("kkundenavn")), 1) &"</option>"
                                 end if
                              

                                if len(trim(oRec2("kkundenavn"))) <> 0 AND lastKid <> oRec2("kid") then
                                strFil_Kon_Txt = strFil_Kon_Txt & "<option DISABLED></option>"
                                strFil_Kon_Txt = strFil_Kon_Txt & "<option value='k_"& oRec2("kid") &"'>"& oRec2("kkundenavn") &" "_
                                & left(oRec2("adresse"), 30) &" "_
                                & oRec2("postnr") &" "& oRec2("city")&" "_
                                & oRec2("land")
                                strFil_Kon_Txt = strFil_Kon_Txt &  "</option>"
                                end if
                                


                  x = x + 1   
                  k = 1 
                       
                        
                
                               if len(trim(oRec2("kpid"))) <> 0 then
                                
                                               
                                        if len(trim(oRec2("kpersnavn"))) <> 0 then
                                        strFil_Kon_Txt = strFil_Kon_Txt & "<option value='kp_"& oRec2("kpid") &"'>.........."& oRec2("kpersnavn") &" "
                                       
                                        if len(trim(oRec2("kpafd"))) <> 0 then
                                        strFil_Kon_Txt = strFil_Kon_Txt &  "Afd:"&  oRec2("kpafd") &" "
                                        end if
                                        
                                        strFil_Kon_Txt = strFil_Kon_Txt & left(oRec2("kpadr"), 20) &" "_
                                        & oRec2("kpzip") &" "& oRec2("kptown") &" "& oRec2("kpland")
                                       
                                        
                                        strFil_Kon_Txt = strFil_Kon_Txt &  "</option>"
                                        end if
                                        

                                  x = x + 1
                                end if 'oRec2("id") <> 0    
                
                              
                lastAlfabet = left(oRec2("kkundenavn"), 1)
                lastknavn = oRec2("kkundenavn") 
                lastKid = oRec2("kid")
                oRec2.movenext
                wend
                oRec2.close


                

                                    'Response.Write strFil_Kon_Txt
                                    'Response.end

                           
                                    
                                   
                     
                
                
                '*** ��� **'
                call jq_format(strFil_Kon_Txt)
                strFil_Kon_Txt = jq_formatTxt
                Response.write strFil_Kon_Txt  
                response.end

    
    
    
    
    case "FN_sogjobogkunde" 

        
                '*** S�G kunde & Job            
                
                if len(trim(request("jq_newfilterval"))) <> 0 then
                filterVal = 1 
                jobkundesog = request("jq_newfilterval")
                else
                filterVal = 0
                jobkundesog = "6xxxxxfsdf554"
                end if
        
                medid = request("jq_medid")

                if len(trim(request("jq_ign_projgrp"))) <> 0 then
                ign_projgrp = request("jq_ign_projgrp")
                else
                ign_projgrp = 0
                end if

                if filterVal <> 0 then

                    
                    varTjDatoUS_man = request("varTjDatoUS_man")
                    varTjDatoUS_son = dateAdd("d", 6, varTjDatoUS_man)

                    varTjDatoUS_man = year(varTjDatoUS_man) &"/"& month(varTjDatoUS_man) &"/"& day(varTjDatoUS_man)
                    varTjDatoUS_son = year(varTjDatoUS_son) &"/"& month(varTjDatoUS_son) &"/"& day(varTjDatoUS_son)                              

            
                 lastKid = 0
                
                  'strJobogKunderTxt = strJobogKunderTxt &"<span style=""color:red; font-size:9px; float:right;"" class=""luk_jobsog"">[X]</span>"    
            
                    'call positiv_aktivering_akt_fn()
                    'positiv_aktivering_akt_val = positiv_aktivering_akt_val

                    'if cint(positiv_aktivering_akt_val) = 1 OR cint(pa_aktlist) = 1 then 'Positiv aktivering / M� kun s�ge i aktiv jobliste
                    'positiv_aktivering_akt_val_SQL = " AND ((tu.forvalgt = 1 "
            
                    '        if cint(positiv_aktivering_akt_val) = 1 then
                    '        positiv_aktivering_akt_val_SQL = positiv_aktivering_akt_val_SQL & " AND tu.aktid <> 0 AND risiko >= 0) OR risiko < 0)"
                    '        else
                    '        positiv_aktivering_akt_val_SQL = positiv_aktivering_akt_val_SQL & "))"
                    '        end if

                    'else
                    'positiv_aktivering_akt_val_SQL = ""
                    'end if
 
                


                             '***** POSITIV aktivitering ***'
                    call positivakt_forecast_sqlkri(medid)
                    strSQLPAkri = positiv_aktivering_akt_val_SQL

                      

                         '*** Datosp�rring Vis f�rst job n�r stdato er oprindet
                        'call lukaktvdato_fn()
                        'ignJobogAktper = lukaktvdato

                        'select case lto
                        'case "mpt"
                        'jobstatusExtra = " OR (j.jobstatus = 2) OR (j.jobstatus = 4)"
                        'case else
                        'jobstatusExtra = ""
                        'end select


                        'select case ignJobogAktper
                        'case 0,1
                        'strSQLDatoStatuskri = " AND ((j.jobstartdato <= '"& varTjDatoUS_son &"' AND j.jobstatus = 1) OR (j.jobstatus = 3) "& jobstatusExtra &")"
                        'case 3
                        'strSQLDatoStatuskri = " AND ((j.jobstartdato <= '"& varTjDatoUS_son &"' AND j.jobslutdato >= '"& varTjDatoUS_man &"' AND j.jobstatus = 1) OR (j.jobstatus = 3) "& jobstatusExtra &")"
                        'case else
                        'strSQLDatoStatuskri = " AND ((j.jobstatus = 1) OR (j.jobstatus = 3) "& jobstatusExtra &")"
                        'end select


                call datosparrings_sglkri(lto, varTjDatoUS_man, varTjDatoUS_son)
                strSQLDatoStatuskri = strSQLDatoStatuskri


                if cint(ign_projgrp) = 0 then' alm

                    strSQL = "SELECT j.id AS jid, j.jobnavn, j.jobnr, j.jobstatus, k.kkundenavn, k.kkundenr, k.kid FROM timereg_usejob AS tu "_ 
                    &" LEFT JOIN job AS j ON (j.id = tu.jobid) "_
                    &" LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "_
                    &" WHERE tu.medarb = "& medid &" "& positiv_aktivering_akt_val_SQL &" "& strSQLDatoStatuskri &" AND "_
                    &" (jobnr LIKE '"& jobkundesog &"%' OR jobnavn LIKE '%"& jobkundesog &"%' OR "_
                    &" kkundenavn LIKE '"& jobkundesog &"%' OR kkundenr = '"& jobkundesog &"' OR k.kinit = '"& jobkundesog &"')  AND kkundenavn <> ''"_
                    &" GROUP BY j.id ORDER BY kkundenavn, jobnavn LIMIT 50"       
    
                else 'ignorer projektgrupper

                
                    strSQL = "SELECT j.id AS jid, j.jobnavn, j.jobnr, j.jobstatus, k.kkundenavn, k.kkundenr, k.kid FROM job AS j "_
                    &" LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "_
                    &" WHERE "& strSQLDatoStatuskri &" AND "_
                    &" (jobnr LIKE '"& jobkundesog &"%' OR jobnavn LIKE '%"& jobkundesog &"%' OR "_
                    &" kkundenavn LIKE '"& jobkundesog &"%' OR kkundenr = '"& jobkundesog &"' OR k.kinit = '"& jobkundesog &"') AND kkundenavn <> ''"_
                    &" GROUP BY j.id ORDER BY kkundenavn, jobnavn LIMIT 50"      


                end if


                 'response.write "strSQL " &strSQL
                 'response.end
                antalJob = 0
                oRec.open strSQL, oConn, 3
                while not oRec.EOF
        
               ' if lastKid <> oRec("kid") then
                'strJobogKunderTxt = strJobogKunderTxt &"<br><br><b>"& oRec("kkundenavn") &" "& oRec("kkundenr") &"</b><br>"
               ' end if 
                 
                
                if lastKid <> oRec("kid") then
                    if antalJob <> 0 then
                    strJobogKunderTxt = strJobogKunderTxt & "<option value=""0"" DISABLED ></option>"
                    end if
                strJobogKunderTxt = strJobogKunderTxt & "<option value=""0"" DISABLED >"& oRec("kkundenavn") & "</option>" '&"<br><br><b>"& oRec("kkundenavn") &" "& oRec("kkundenr") &"</b><br>"
                end if

                'strJobogKunderTxt = strJobogKunderTxt & "<input type=""hidden"" id=""hiddn_job_"& oRec("jid") &""" value="""& oRec("jobnavn") & " ("& oRec("jobnr") &")"">"
                'strJobogKunderTxt = strJobogKunderTxt & "<input type=""checkbox"" class=""chbox_job"" id=""chbox_job_"& oRec("jid") &""" value="& oRec("jid") &"> "& oRec("jobnavn") & " ("& oRec("jobnr") &")" &"<br>" 
                select case lto
                case "tia"
                strJobogKunderTxt = strJobogKunderTxt & "<option value="& oRec("jid") &" "& jobSEL &">"& oRec("jobnr") &" - "& oRec("jobnavn") & "</option>"
                case else
                strJobogKunderTxt = strJobogKunderTxt & "<option value="& oRec("jid") &" "& jobSEL &">"& oRec("jobnavn") & " ("& oRec("jobnr") &")</option>"
                end select            

                antalJob = antalJob + 1 
                lastKid = oRec("kid") 
                oRec.movenext
                wend
                oRec.close


                    strJobogKunderTxt = strJobogKunderTxt & "<option value=0 disabled>Jobs: "& antalJob &"</option>" 'ign_projgrp: ("& ign_projgrp &")
              
                     'if lto = "cflow" then
                     'strJobogKunderTxt = strJobogKunderTxt & "<option value=0>"& strSQL & " "&  varTjDatoUS_man & "</option>"
                     'end if


                    '*** ��� **'
                    call jq_format(strJobogKunderTxt)
                    strJobogKunderTxt = jq_formatTxt

                    response.write strJobogKunderTxt

                end if



     case "FN_sogakt"

               
                '*** S�g Aktiviteter 
                

                if len(trim(request("jq_newfilterval"))) <> 0 then
                filterVal = 1 
                aktsog = request("jq_newfilterval")
                else
                filterVal = 0
                aktsog = "6xxxxxfsdf554"
                end if
        
                medid = request("jq_medid")
                aktid = request("jq_aktid")
                FM_medid = request("FM_medid")
                FM_jobid = request("FM_jobid")
    
                if len(trim(request("jq_jobid"))) <> 0 then        
                jobid = request("jq_jobid")
                else
                jobid = 0
                end if


                ign_projgrp = 0
                if len(trim(request("jq_ign_projgrp"))) <> 0 then
                ign_projgrp = request("jq_ign_projgrp")
                else
                ign_projgrp = 0
                end if


                 'if session("mid") = 1 then
                'response.write "strSQL " & strSQL
                 'response.end
                'strAktTxt = strAktTxt & "<option value=0>ign_projgrp: "& ign_projgrp &"</option>"
                'response.write strAktTxt
                'response.end
                'end if

                 call akttyper2009(2)

                'positiv aktivering
                call positiv_aktivering_akt_fn()
                pa = pa_aktlist
                positiv_aktivering_akt_val = positiv_aktivering_akt_val

	            'pa = positiv_aktivering_akt_val 
                'pa = 0
               ' if len(trim(request("jq_pa") )) <> 0 then
                'pa = request("jq_pa") 
                'else
                'pa = 0
               ' end if
        
                'pa = 0
            
                '*** Sales / tilbud kun Salgsaktiviteter
                '(a.fakturerbar = 6 AND j.jobstatus = 3)
                jobstatusTjk = 1
                strSQLtilbud = "SELECT jobstatus, jobnavn, risiko FROM job WHERE id = "& jobid
                oRec5.open strSQLtilbud, oConn, 3
                if not oRec5.EOF then

                jobstatusTjk = oRec5("jobstatus")
                jobNavnTjk = oRec5("jobnavn")
                jobRisikoTjk = oRec5("risiko")

                end if
                oRec5.close

                
    

                if lto = "mpt" OR session("lto") = "9K2017-1121-TO178" then
                  onlySalesact = ""

                else

                if cint(jobstatusTjk) = 3 then 'tilbud
                onlySalesact = " AND a.fakturerbar = 6"
                else
                onlySalesact = ""
                end if

                end if



                if filterVal <> 0 then
            
                 
    
                'strAktTxt = strAktTxt &"<span style=""color:#999999; font-size:9px; float:right;"" class=""luk_aktsog"">X</span>"    
                 

                '************************************************************************************************
                '*** FINDES AKT i firvejen p� favoritlisten 
                '************************************************************************************************
                 findesaktfav = "#0#"
                'strSQLfavakt = "SELECT aktid FROM timereg_usejob WHERE favorit <> 1 AND medarb = "& FM_medid &" AND jobid = "& FM_jobid
                strSQLfavakt = "SELECT aktid, favorit FROM timereg_usejob WHERE favorit = 1 AND medarb = "& FM_medid &" AND jobid = "& FM_jobid
                oRec6.open strSQLfavakt, oConn, 3
                while not oRec6.EOF
                findesaktfav = findesaktfav & ",#"& oRec6("aktid") &"#"                  
                oRec6.movenext
                wend          
                oRec6.close  


                '*** Praktikant m� ikke se sygdom **
                strSQLaidmaikkesetype = ""
                if (lto = "plan" AND session("mid") = 274 AND medid <> 274) OR (lto = "plan" AND session("mid") = 1 AND medid <> 1) then
                    strSQLaidmaikkesetype = " AND (a.fakturerbar <> 20 AND a.fakturerbar <> 21) "
                end if


               if pa = "1" OR cint(positiv_aktivering_akt_val) = 1 AND jobRisikoTjk >= 0 then 'WWF speciel + ALLE Interne g�r altid p� projektgrupper rettigheder

                    if cint(positiv_aktivering_akt_val) = 1 then 'Positiv aktivering
                    positiv_aktivering_akt_val_SQL = " AND tu.forvalgt = 1 AND tu.aktid <> 0 "
                    positiv_aktivering_akt_val_SQL_join = " AND a.id = tu.aktid"
                    
                    else
                    positiv_aktivering_akt_val_SQL = ""
                    positiv_aktivering_akt_val_SQL_join = ""
                    end if

               strSQL = "SELECT a.id AS aid, navn AS aktnavn, a.fase FROM timereg_usejob tu LEFT JOIN aktiviteter AS a ON (a.job = tu.jobid "& positiv_aktivering_akt_val_SQL_join &" "& onlySalesact &") "_
               &" WHERE tu.medarb = "& medid &" AND tu.jobid = "& FM_jobid &" "& positiv_aktivering_akt_val_SQL &" AND aktstatus = 1 AND ("& aty_sql_hide_on_treg &") "& onlySalesact &" GROUP BY a.id ORDER BY a.fase, sortorder, navn"   

                'AND aktid <> 0

               else

                        if cint(ign_projgrp) = 0 then' alm


                            '*** Finder medarbejders projektgrupper 
                            '** Medarbejder projektgrupper **'
                            medarbPGrp = "#0#" 
                            strMpg = "SELECT projektgruppeId, medarbejderId, teamleder FROM progrupperelationer WHERE medarbejderId = "& medid & " GROUP BY projektgruppeId"

                            oRec5.open strMpg, oConn, 3
                            while not oRec5.EOF
                            medarbPGrp = medarbPGrp & ",#"& oRec5("projektgruppeId") &"#"         
  
                            oRec5.movenext
                            wend
                            oRec5.close 


                         end if

                          
                          

                        strSQL= "SELECT a.id AS aid, navn AS aktnavn, a.fase, projektgruppe1, projektgruppe2, projektgruppe3, "_
                        &" projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM aktiviteter AS a "_
                        &" WHERE a.job = " & jobid & " AND aktstatus = 1 AND ("& aty_sql_hide_on_treg &") "& onlySalesact &" ORDER BY a.fase, sortorder, navn"      
                          
                        'AND fakturerbar <> 90
            
                end if

                'if session("mid") = 1 then
                'response.write "strSQL " & strSQL
                'response.end
                'strAktTxt = strAktTxt & "<option value=0>PA: "& pa &"HER: "& strSQL &" findesaktfav: "& findesaktfav &"</option>"
                'response.write strAktTxt
                'response.end
                'end if

                strAktTxt = strAktTxt & "<option SELECTED disabled value=0>..</option>"             
                a = 0
                lastFase = ""
                thisFase = ""
                oRec.open strSQL, oConn, 3
                while not oRec.EOF

                


                findesaktfavshowakt = 0
                if instr(findesaktfav, "#"& oRec("aid") &"#") then
                findesaktfavshowakt = 1
                end if
        
                
                'if session("mid") = 1 then
                '   strAktTxt = strAktTxt & "<option value="& oRec("aid") &">"& oRec("aktnavn") &" "& fcsaldo_txt &" ("& findesaktfavshowakt &")</option>" 
                'end if

                
                showAkt = 0
            
                if cint(findesaktfavshowakt) = 0 then

                if pa = "1" OR cint(positiv_aktivering_akt_val) = 1 then
            
                     
                     showAkt = 1
                    
            
                else

                    if cint(ign_projgrp) = 0 then' alm

                            if instr(medarbPGrp, "#"& oRec("projektgruppe1") &"#") <> 0 AND findesaktfavshowakt = 0 _
                            OR instr(medarbPGrp, "#"& oRec("projektgruppe2") &"#") <> 0 AND findesaktfavshowakt = 0 _
                            OR instr(medarbPGrp, "#"& oRec("projektgruppe3") &"#") <> 0 AND findesaktfavshowakt = 0 _
                            OR instr(medarbPGrp, "#"& oRec("projektgruppe4") &"#") <> 0 AND findesaktfavshowakt = 0 _
                            OR instr(medarbPGrp, "#"& oRec("projektgruppe5") &"#") <> 0 AND findesaktfavshowakt = 0 _
                            OR instr(medarbPGrp, "#"& oRec("projektgruppe6") &"#") <> 0 AND findesaktfavshowakt = 0 _
                            OR instr(medarbPGrp, "#"& oRec("projektgruppe7") &"#") <> 0 AND findesaktfavshowakt = 0 _
                            OR instr(medarbPGrp, "#"& oRec("projektgruppe8") &"#") <> 0 AND findesaktfavshowakt = 0 _
                            OR instr(medarbPGrp, "#"& oRec("projektgruppe9") &"#") <> 0 AND findesaktfavshowakt = 0 _
                            OR instr(medarbPGrp, "#"& oRec("projektgruppe10") &"#") <> 0 AND findesaktfavshowakt = 0 then
                            
                            showAkt = 1

                            end if 


                    else
                    showAkt = 1

                    end if

                end if 'pa
                end if 'if cint(findesaktfavshowakt) = 0 then
                                              
                if cint(showAkt) = 1 then 


                        if isnull(oRec("fase")) = false then

                                if len(trim(oRec("fase"))) <> 0 then
                                thisFase = oRec("fase")
                                else
                                thisFase = ""
                                end if

                        else
                        thisFase = ""
                        end if

                                if len(trim(thisFase)) <> 0 AND (lastFase <> thisFase OR len(trim(lastFase)) = 0) then

                                 thisFase = replace(thisFase, "_", " ")
                                 strAktTxt = strAktTxt & "<option value=0 disabled></option>"
                                 strAktTxt = strAktTxt & "<option value=0 disabled>[ "& thisFase &" ]</option>"
                                 
                                 

                                end if
                 
                    'strAktTxt = strAktTxt & "<input type=""hidden"" id=""hiddn_akt_"& oRec("aid") &""" value="""& oRec("aktnavn") &""">"
                    'strAktTxt = strAktTxt & "<input type=""checkbox"" class=""chbox_akt"" id=""chbox_akt_"& oRec("aid") &""" value="& oRec("aid") &"> "& oRec("aktnavn") &"<br>" 
                    
                    strAktTxt = strAktTxt & "<option value="& oRec("aid") &">"& oRec("aktnavn") &" "& fcsaldo_txt &"</option>"             

                  
               
                end if 'showakt

                if isnull(oRec("fase")) = false then
                lastFase = oRec("fase")
                else
                lastFase = ""
                end if
                
                 'strAktTxt = strAktTxt & "<option value="& oRec("aid") &" "& optionFcDis &">"& oRec("aktnavn") &" "& fcsaldo_txt &"</option>"        
                 'strAktTxt = strAktTxt & "<option value="& a &" "& optionFcDis &">"& a &"</option>"

                a = a + 1
                oRec.movenext
                wend
                oRec.close

                 if a = 0 then
                 strAktTxt = strAktTxt & "<option disabled value=""0"">No activities</option>"
                 end if


                    '*** ��� **'
                    call jq_format(strAktTxt)
                    strAktTxt = jq_formatTxt


                    response.write strAktTxt

               
            end if 'filerval




        case "tilfoj_akt"

        jobid = request("FM_jobid")
        aktid = request("FM_aktid")
        splitaktid = split(aktid, ",")
        medid = request("FM_medid_id")

        for a = 0 TO UBOUND(splitaktid)
            
            strSQL = "SELECT id FROM timereg_usejob WHERE jobid = "& jobid &" AND aktid = "& splitaktid(a) &" AND medarb = "& medid
            oRec.open strSQL, oConn, 3
            if not oRec.EOF then
            Strtilfojakt = "UPDATE timereg_usejob SET favorit = 1 WHERE aktid = "& splitaktid(a) & " AND medarb = "& medid &" AND jobid = "& jobid
            else
            Strtilfojakt = "INSERT INTO timereg_usejob SET jobid = "& jobid &", favorit = 1, aktid = "& splitaktid(a) & ", medarb = "& medid
            end if
            oRec.close 

        
            oConn.execute(Strtilfojakt)

        next
        
        'response.write "jobid "& jobid  & "medid: "& medid & " aktid: "& aktid
         'response.end

        

        'response.write "HER: "&  Strtilfojakt
        'response.end


        'oConn.execute("UPDATE timereg_usejob SET favorit = 1 WHERE aktid = "&aktid&"")
        'Strtilfojakt = "INSERT INTO timereg_usejob SET jobid = "& jobid &", favorit = 1, aktid = "& aktid & ", medarb = "& medid
        'oConn.execute(Strtilfojakt)

        'response.end

       'akt_jobid = jobid
        'akt_aktid = aktid


        'response.Write Strtilfojakt
        'response.end 

        'Response.redirect "favorit.asp"

        'strSQLUpdjWiphist = "INSERT INTO wip_historik (dato, restestimat, stade_tim_proc, medid, jobid) VALUES ('"& ddDato &"', "& rest &", "& timerproc &", "& session("mid") &", "& jobid &")"
        'oConn.Execute(strSQLUpdjWiphist)




        case "tilfoj_mat"
    
        mat_jobid = request("mat_jobid")
        mat_aktid = request("mat_aktid")
        mat_antal = request("mat_antal")
        mat_navn = request("mat_navn")
        mat_kobpris = request("mat_kobpris")
        mat_enhed = request("mat_enhed")
        mat_dato = request("mat_dato")
        mat_forbrugsdato = request("mat_forbrugsdato")
        mat_editor = request("mat_editor")
        mat_usrid = request("mat_usrid")
        mat_gruppe = request("mat_gruppe")
        mat_salgspris = request("mat_salgspris")
        mat_bilagsnr = request("mat_bilagsnr")
        mat_valuta = request("mat_valuta")
        mat_varenr = request("mat_varenr") 

        select case lto
        case "sdeo" 
        mf_konto = 1
        case else
        mf_konto = 0
        end select
        
        strsqlmat = "INSERT INTO materiale_forbrug SET jobid ="& mat_jobid & ", aktid ="& mat_aktid & ", matantal ="& mat_antal & ", matnavn ='"& mat_navn &"', matkobspris ="& mat_kobpris & ", matenhed ='"& mat_enhed &"'" & ", dato ='"& mat_dato &"'" & ", forbrugsdato ='"& mat_forbrugsdato & "'" & ", editor ='"& mat_editor &"'" & ", usrid =" & mat_usrid & ", matgrp ="& mat_gruppe & ", "_
        &" matsalgspris ="& mat_salgspris & ", bilagsnr ='"& mat_bilagsnr & "', valuta ="& mat_valuta & ", matvarenr ='" & mat_varenr & "', mf_konto = "& mf_konto &""   
        oConn.execute(strsqlmat)
        

        end select                  
        response.end
        end if

        
    %>


<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../timereg/inc/convertDate.asp"-->

<%
    thisfile = "favorit.asp"
%>

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<SCRIPT src="../timereg/inc/smiley_jav.js"></script>
<script src="js/favorit_jav_20200205.js"></script>
<!-- <link rel="stylesheet" href="../../bower_components/bootstrap-datepicker/css/datepicker3.css"> -->




<style>
    .kommodal:hover,
    .kommodal:focus {
    text-decoration: none;
    cursor: pointer;
    } 
</style>

<%call browsertype()  %>

<%if (browstype_client <> "ip") then %>
<div class="wrapper">
<div class="content">
<%end if %>

<%if (browstype_client = "ip") then %>
<meta name="viewport" content="width=device-width, minimum-scale=1.0, maximum-scale=1.0, user-scalable=no">
<%end if %>


    <%
    if len(session("user")) = 0 then
	%>
	
	<%
	errortype = 5
	call showError(errortype)
       response.End
	end if

    

    'if len(trim(request("FM_medid"))) <> 0 then
    'medid = request("FM_medid")
    'else
    'medid = session("mid")
    'end if


      if len(trim(request("usemrn"))) <> 0 then 'from ugeseddel, komme/g� minilinks
      usemrn = request("usemrn")
      medid = usemrn

     else

        if len(trim(request("FM_medid"))) <> 0 then
        medid = request("FM_medid")
        else
        medid = session("mid") 
        end if

        usemrn = medid

    end if


    '*** TIA Test for taste for andre medarbejdere
    'if session("mid") = 1 then
    '     medid = 1592
    '     usemrn = medid
    '     level = 6

    '    Response.write "YOU ARE TESTING: " & medid & " level: " & level
    'end if

    if len(trim(request("FM_ign_projgrp"))) <> 0 then
    ign_projgrp = 1
    ign_projgrpCHK = "CHECKED"
    else
    ign_projgrp = 0
    ign_projgrpCHK = ""
    end if

    struSQLusemedarb = "SELECT mnavn FROM medarbejdere WHERE mid ="& medid
    oRec.open struSQLusemedarb, oConn, 3
    if not oRec.EOF then
    medarbejdernavn = oRec("mnavn")
    end if
    oRec.close

    call aktBudgettjkOn_fn()

    if (browstype_client <> "ip") then
        call menu_2014
    else
        %><!--#include file="../inc/regular/top_menu_mobile.asp"--><%


        ddDato = year(now) &"/"& month(now) &"/"& day(now)
        ddDato_ugedag = day(now) &"/"& month(now) &"/"& year(now)
        ddDato_ugedag_w = datepart("w", ddDato_ugedag, 2, 2)
            
        varTjDatoUS_man_tt = dateAdd("d", -(ddDato_ugedag_w-1), ddDato_ugedag)
        thisfile = "favorit_mob"

        call mobile_header
    end if

    func = request("func")


	if func = "db" then


        aktids = request("FM_aktivitetids")
        strAktNavn = request("FM_aktnavn")

        'response.Write "test tekst: " & aktids & strAktNavn
        response.End


    end if


    stDato = request("stDato")

    redim tjekdag(7)
    tjekdag(1) = stDato
            
    for x = 2 to 7
    tjekdag(x) = dateAdd("d", x-1, stDato)
    next



    
    'response.Write "medid: " & medid

    
    if (browstype_client <> "ip") then
        varTjDatoUS_man = request("varTjDatoUS_man")
    else
        varTjDatoUS_man = request("FM_choosendate")
    end if


    if varTjDatoUS_man = "" then
        varTjDatoUS_man = Day(now) &"-"& Month(now) &"-"& Year(now)
    end if
    mondayofweek = DateAdd("d", -((Weekday(varTjDatoUS_man) + 7 - 2) Mod 7), varTjDatoUS_man)
                  

    varTjDatoUS_man = mondayofweek
    varTjDatoUS_son = dateAdd("d", 6, varTjDatoUS_man)

    datoMan = day(varTjDatoUS_man) &"/"& month(varTjDatoUS_man) &"/"& year(varTjDatoUS_man)
    datoSon = day(varTjDatoUS_son) &"/"& month(varTjDatoUS_son) &"/"& year(varTjDatoUS_son)

    next_varTjDatoUS_man = dateadd("d", 7, varTjDatoUS_man)
	next_varTjDatoUS_son = dateadd("d", 7, varTjDatoUS_son)
    next_varTjDatoUS_man = day(next_varTjDatoUS_man) & "-" & month(next_varTjDatoUS_man) &"-"& year(next_varTjDatoUS_man)
	next_varTjDatoUS_son = day(next_varTjDatoUS_son) &"-"& month(next_varTjDatoUS_son) &"-"& year(next_varTjDatoUS_son)


    prev_varTjDatoUS_man = dateadd("d", -7, varTjDatoUS_man)
	prev_varTjDatoUS_son = dateadd("d", -7, varTjDatoUS_son)
    prev_varTjDatoUS_man = day(prev_varTjDatoUS_man) &"-"& month(prev_varTjDatoUS_man) &"-"& year(prev_varTjDatoUS_man) 
	prev_varTjDatoUS_son = day(prev_varTjDatoUS_son) &"-"& month(prev_varTjDatoUS_son) &"-"& year(prev_varTjDatoUS_son) 

    weeknumber = year(varTjDatoUS_man) & "-" & month(varTjDatoUS_man) & "-" & day(varTjDatoUS_man)

    select case func 

    case "fjernfavorit"

        medid = request("FM_medid")
        varTjDatoUS_man = request("varTjDatoUS_man")
        id = request("id")

        
        if id <> 0 then
        oConn.execute("UPDATE timereg_usejob SET favorit = 0 WHERE aktid = "& id &" AND medarb = "& medid)
        end  if

        response.Redirect "favorit.asp?FM_medid="&medid&"&varTjDatoUS_man="&varTjDatoUS_man


    case "XXXXXtilfojfavorit"

        id = request("id")
                            
            %>
                <div class="container">
                <div class="portlet">
                    <h3 class="portlet-title"><u><%=favorit_txt_001 %></u> </h3>
                    <div class="portlet-body">
                        <%response.Write "favor" & id  %>
                    </div>

                </div>
                </div>
            <%

        oConn.execute("UPDATE timereg_usejob SET favorit = 1 WHERE aktid = "&id&"")

        response.Redirect "favorit.asp"




    case else

   



    'Response.write "<br><br><br><br><br><br><br><div style='left:200px; top:100px; position:relative;'>MEDID: " & request("FM_medid") & "</div>"
    call ersmileyaktiv()
    call smileyAfslutSettings()

    if cint(SmiWeekOrMonth) = 0 then

    call thisWeekNo53_fn(varTjDatoUS_man)

    usePeriod = thisWeekNo53 'datePart("ww", varTjDatoUS_man, 2,2)
    useYear = year(varTjDatoUS_man)
    else
    usePeriod = datePart("m", varTjDatoUS_man, 2,2)
    useYear = year(varTjDatoUS_man)
    end if

    call erugeAfslutte(useYear, usePeriod, medid, SmiWeekOrMonth, 0, varTjDatoUS_man)

    call visAktSimpel_fn()

%>

            <%
            if browstype_client = "ip" then
                if len(trim(request("FM_choosenDate"))) <> 0 then
                    choosenDate = year(request("FM_choosenDate"))&"-"&month(request("FM_choosenDate"))&"-"&day(request("FM_choosenDate"))
                    dayNumberChoosen =  weekday(choosenDate , 2)
                else
                    choosenDate = year(now)&"-"&month(now)&"-"&day(now)
                    dayNumberChoosen = weekday(choosenDate , 2)
                end if
            else
                dayNumberChoosen = -1
            end if
            
            
            '**** NYHEDER ELLER KALENDER
            select case lto
            case "dencker", "xintranet - local"
            %>

            <!--#include file="../timereg/inc/timereg_akt_2006_nyheder_inc.asp"-->
            <!-- Nyheds DEL -->
    
                <%
                    call HentNyhedsVariabler()
                %>

                <div style="position:absolute; right:0; top:0px; z-index:1000; margin-right:20px;">
                <div style="width:250px;">
                <div class="panel-group accordion-panel" id="accordion-paneled1">

                    <div class="panel">

                    <div class="panel-heading" id="nyhederdropdown">
                        <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapse_news">
                                Nyheder <!--<b><span id="nyheds_icon" class="fa fa-angle-up"></span></b>-->
                            </a>
                        </h4>
                    </div>

                    <div id="collapse_news" class="panel-collapse collapse in">
                        <div class="panel-body">

                           <table style="width:100%; background-color:white;">  
                            <%for i = 0 TO UBOUND(nyhedsid) %>
                             <%if nyhedoverskrift(i) <> "" then %>
                             <tr>
                                <td>
                                    <%if nyhedvigtig(i) = 1 then  %>
                                    <h4 style="color:#f44842;"><%=nyhedoverskrift(i) %> <br /> <span style="font-size:9px; color:#9e9e9e;">Oprettet: <%=datofra(i) %> af <%=newsEditor(i) %></span></h4>
                                    <%else %>
                                    <h4><%=nyhedoverskrift(i) %> <br /> <span style="font-size:9px; color:#9e9e9e;">Oprettet: <%=datofra(i) %> af <%=newsEditor(i) %></span></h4>
                                    <%end if %>

                                    <%if nyhedbrodtext(i) <> "" then  %>
                                    <strong><%=nyhedbrodtext(i) %></strong>
                                    <%else %>
                                    <strong>Ingen beskrivelse</strong>
                                    <%end if %>


                                    <%if filnavn(i) <> "" then %>
                                        <br /><span id="modal_<%=nyhedsid(i) %>" class="nyhedsbillede"><img src="../inc/upload/<%=lto%>/<%=filnavn(i) %>" alt='' border='0' style="width:100px; height:100px;"></span>
                                        <div id="nyhedsbillede_<%=nyhedsid(i) %>" class="modal">
                                            <!-- Modal content -->
                                            <div class="modal-content" style="text-align:center; width:1600px; height:850px">
                                                <img src="../inc/upload/<%=lto%>/<%=filnavn(i) %>" alt='' border='0' style="max-width:100%; max-height:100%;">
                                            </div>
                                         </div>
                                    <%end if %> 

                                    
                                </td>                                
                            </tr>
                            <tr>
                                <td>&nbsp</td>
                            </tr>
                             <%end if %>
                            <%
                            next
                            %>

                            <%if antalnyheder = 0 then %>
                               <tr>
                                   <td>Ingen nyheder i dag</td>
                               </tr>
                            <%end if %>
                         </table>

                        </div> 
                    </div> 

                    </div> </div>
                    </div>
                </div>

            

            <div class="container">


  
               
                <%call NyhedsPopUp() %>

                <script src="../timereg/inc/timereg_nyheder_jav.js"></script>

        <%case else %>

                <style>

    
                    /* The Modal (background) */
                    .modal {
                        display: none; /* Hidden by default */
                        position: fixed; /* Stay in place */
                        z-index: 3; /* Sit on top */
                        padding-top: 200px; /* Location of the box */
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
                        width: 500px;
                        height: 450px;
                    }

                    .picmodal:hover,
                    .picmodal:focus {
                    text-decoration: none;
                    cursor: pointer;
                    }

                    .kommodal:hover,
                    .kommodal:focus {
                    text-decoration: none;
                    cursor: pointer;
                    }
   
                </style>

          
          

            <%
            'response.Cookies("calendar2020") = "0"

            if browstype_client <> "ip" then 
                select case request.Cookies("calendar2020")
                case "1"
                maincontainerstyle = "style='margin-left:115px; width:1150px;'" 
                case else
                maincontainerstyle = ""
                end select
            else
                maincontainerstyle = ""
            end if %>


            <%if request.Cookies("calendar2020") = "1" AND browstype_client <> "ip" then
                
                response.Write "<div style='position:absolute; left:1270px;'>"
                call calender_2015
                response.Write "</div>"

            end if %>



            <div class="container" <%=maincontainerstyle %> >
             <%end select%>





                <div class="portlet">
                    <%if browstype_client <> "ip" then  %>
                    <h3 class="portlet-title"><u><%=favorit_txt_001 %></u>

                        <%if lto = "cflow" then
                            %>
                          <span style="font-size:50%;">(Bruk denne n�r Du arbeider ute / ikke hos Cflow)</span>

                         <%end if%>

                    </h3>
                    <%end if %>

                    <div class="portlet-body">
                        
                          
                        
                        <%
                        if browstype_client <> "ip" then
                        %>
                                        <%   
                            
                        select case lto 
                        case "bf"
            
                        case else%>

                        <%   
                            
                            
                            select case lto 
                            case "ddc"
                            wdth = 140
                            lft = 960 
                            case else
                            wdth = 240
                            lft = 860 
                            end select


                            if cint(stempelurOn) = 1 then 
                                wdth = wdth + 60
                                lft = lft - 60
                          
                            end if

                                if cint(vis_favorit) = 1 then
                                wdth = wdth + 60
                                lft = lft - 60
                                end if

                        %>
          
           
                            <div style="position:relative; background-color:#ffffff; border:1px #cccccc solid; border-bottom:0; padding:4px; width:<%=wdth%>px; top:-60px; left:<%=lft%>px; z-index:0; font-size:11px;">
           
                                  <%select case lto
                                    case "ddc", "cflow", "foa", "care", "kongeaa"
                                    case else
                                    %>
                                    <a href="../timereg/<%=lnkTimeregside%>" class="vmenu"><%=replace(tsa_txt_031, " ", "")%></a>
                                    &nbsp;&nbsp;|&nbsp;&nbsp;
                                    <%end select %>

                           <%  
                            select case lto
                            case "cflow"
                            call medariprogrpFn(session("mid"))

                                if instr(medariprogrpTxt, "#14#") <> 0 OR instr(medariprogrpTxt, "#16#") <> 0 OR instr(medariprogrpTxt, "#3#") <> 0 OR instr(medariprogrpTxt, "#19#") <> 0 OR level = 1 then 
                               %>
                                 <a href="<%=lnkUgeseddel%>" class="vmenu"><%=tsa_txt_337 %></a>&nbsp;&nbsp;|&nbsp;&nbsp;
                                <%
                                end if

                            case else
                            %>
                            <a href="<%=lnkUgeseddel%>" class="vmenu"><%=tsa_txt_337 %></a>&nbsp;&nbsp;|&nbsp;&nbsp;
                            <%end select %>


                         <%if cint(stempelurOn) = 1 then
                
               
                            select case lto
                            case "cflow"
                            call medariprogrpFn(session("mid"))

                                if instr(medariprogrpTxt, "#14#") <> 0 OR instr(medariprogrpTxt, "#16#") <> 0 OR instr(medariprogrpTxt, "#3#") <> 0 OR instr(medariprogrpTxt, "#19#") <> 0 OR level = 1 then 
                                %>
                                <a href="<%=lnkLogind%>" class="vmenu"><%=tsa_txt_340 %></a>&nbsp;&nbsp;|&nbsp;&nbsp;
                                <%
                                end if
                            case else
                                    %>

                         <a href="<%=lnkLogind%>" class="vmenu"><%=tsa_txt_340 %></a>&nbsp;&nbsp;|&nbsp;&nbsp;
                        <%end select %>
               
            
                      <%end if


                                   if cint(vis_favorit) = 1 then

                                        select case lto
                                        case "cflow"
                                        call medariprogrpFn(session("mid"))

                                        if instr(medariprogrpTxt, "#14#") <> 0 OR instr(medariprogrpTxt, "#16#") <> 0 OR instr(medariprogrpTxt, "#3#") <> 0 OR instr(medariprogrpTxt, "#19#") <> 0 OR instr(medariprogrpTxt, "#21#") <> 0 then 
                                           %>
                                              <a href="<%=lnkFavorit%>" style="background-color:azure"><%=favorit_txt_001 %></a>
                                            <%
                                            end if

                                        case else
                                        %>
                                        <a href="<%=lnkFavorit%>" style="background-color:azure"><%=favorit_txt_001 %></a>
                                        <%end select
                                     
                                   end if %>


                                   <%if lto = "foa" OR lto = "care" then %>
                                        &nbsp;&nbsp;|&nbsp;&nbsp;
                                        <a href="../timereg/<%=lnkAfstem%>"><%=afstem_txt_011 %></a>
                                   <%end if %>
           
                            </div>
            

                        <%end select %>
                        <%end if %>

                        <!--<br />-->

                        <%if browstype_client <> "ip" then %>

                            <div class="row">
                                          <div class="col-lg-10">
                                            <%call thisWeekNo53_fn(weeknumber) %>
                                            <h4><a href="favorit.asp?FM_medid=<%=medid %>&varTjDatoUS_man=<%=prev_varTjDatoUS_man %>" ><</a>&nbsp <%=favorit_txt_005 &" "& thisWeekNo53%> &nbsp<a href="favorit.asp?FM_medid=<%=medid %>&varTjDatoUS_man=<%=next_varTjDatoUS_man %>" >></a></h4>
                                          </div>


                                        <%if (request.Cookies("calendar2020") = "" OR request.Cookies("calendar2020") = "0") AND browstype_client <> "ip" then %>
                                          <div class="col-lg-2" style="text-align:right;">
                                                <h4 class="panel-title"><a class="accordion-toggle" data-toggle="collapse" href="#calender_div"><span class="fa fa-calendar-o"></span></a></h4>

                                               
                                                
                                                     <div id="calender_div" style="position:absolute; border:1px #d9d9d9 solid; background-color:#FFFFFF; z-index:10; margin-right:150px; left:-180px;" class="panel-collapse collapse">
                               
                                                             <%call calender_2015 %>
             
                                                    </div>

                                               
                                          </div>
                                         <%end if %>
                            </div>

                        <!--  <div class="panel-group accordion-panel" id="accordion-paneled">
                                       <div class="panel panel-default">
                                                <div class="panel-heading">

                                                  <h4 class="panel-title"><a class="accordion-toggle" data-toggle="collapse" href="#calender_div">CALENDER <span class="glyph icon-home"></span>
                                                       
                                                 </h4>

                                                </div><!-- Panel - heading 

                         </div>
                        </div><!-- panel group -->
                                           

                             

                                <!-- <div id="calender_div" style="<%=stylepos%> padding:5px; background-color:#FFFFFF;" class="panel-collapse collapse">
                                
                                       <div class="panel-body">
                                             <div class="row">
                                          <div class="col-lg-12">
                                            <%if browstype_client <> "ip" then %>
                                         <%'call calender_2015 %>
                                        <%end if %></div>

                                        </div>
                                     </div>
                                </div> -->

                           

                        <%end if %> 

                        <form action="favorit.asp?" method="post" name="favorit" id="favorit">
                                              
                            <%
                                showTd = "normal"
                                if browstype_client = "ip" then
                                showTd = "none"
                                end if
                            %>

                            <table style="display:<%=showTd%>;">
                                <tr>
                                    <td>
                                        <%'SKAL UDBYGGES TIL ALLE man er projektelder for
                                        select case cint(level) 
                                        case 1 
                                        sqgMw = " mansat <> 2 AND mansat <> 4 "
                                        case 2,6

                                            call medarb_teamlederfor

                                            sqgMw = " mansat <> 2 AND mansat <> 4 "& medarbgrpIdSQLkri
                                        case else
                                            sqgMw = " mid = "& session("mid") &" AND mansat <> 2 AND mansat <> 4"
                                        end select

                                        strSQL = "SELECT Mid, Mnavn, Mnr, Brugergruppe, init FROM medarbejdere WHERE "& sqgMw &" GROUP BY mid ORDER BY Mnavn" 
                                        %>
                                        <select name="FM_medid" id="FM_medid" <%=progrpmedarbDisabled  %> class="form-control input-small"  onchange="submit();" style="width:240px">
                                            <%

                                            oRec.open strSQL, oConn, 3
                                            while not oRec.EOF
                                
                                    
                                            StrMnavn = oRec("Mnavn")
                                            StrMinit = oRec("init")
	
                                            if cdbl(medid) = cdbl(oRec("Mid")) then
				                            isSelected = "SELECTED"
                                            else
				                            isSelected = ""
				                            end if

				                            %>
                                             <option value="<%=oRec("Mid")%>" <%=isSelected%>><%=StrMnavn &" ["& StrMinit & "]"%></option>
                                            <%
                                            oRec.movenext
                                            wend
                                            oRec.close  
                                            %>
                                        </select>
                                    </td>
                                    <td>
                                        <div class="col-lg-2" style="z-index:1;">
                                            <div class='input-group date' style="width:135px">
                                                <input type="text" class="form-control input-small" name="varTjDatoUS_man" id="varTjDatoUS_man" value="<%=varTjDatoUS_man %>" />
                                                <span class="input-group-addon input-small">
                                                <span class="fa fa-calendar">
                                                </span>
                                                </span>
                                            </div>
                                        </div>                                                                                                                           
                                    </td>
                                    
                                    
                                    <%if browstype_client <> "ip" then %>
                                    <td>                                       
                                        <button type="submit" class="btn btn-sm btn-default"><b><%=favorit_txt_004 %></b></button>                                       
                                    </td>

                                    
                                    <%call thisWeekNo53_fn(weeknumber) %>
                                   
                                    <%end if %> 

                                    <td style="text-align:right; width:100%;"><span id="submitreg" class="btn btn-success btn-sm"><b><%=favorit_txt_018 %></b></span></td>
                                </tr>
                            </table>

                        </form>                       

                        <%'response.Write "id: " & medid%>
                        <%if browstype_client = "ip" then %>

                         

                        <div class="row">
                            <div class="col-lg-2"><h1><a href="favorit.asp?FM_choosenDate=<%=DateAdd("d", -1, choosenDate) %>"><</a> &nbsp <a href="favorit.asp?FM_choosenDate=<%=DateAdd("d", 1, choosenDate) %>">></a></h1></div>
                          
                        </div>
                        <%end if %>

                        <%
                        if browstype_client <> "ip" then
                            FMlink = "rdir=favorit&varTjDatoUS_man="&varTjDatoUS_man
                        else
                            FMlink = "rdir=favorit_mob&FM_choosendate="&choosendate
                        end if
                        %>

                        <%select case lto
                        case "cflow", "xintranet - local"
                            %><br /><br /><%
                            sqlDatoStart = year(varTjDatoUS_man) &"/"& month(varTjDatoUS_man) & "/" & day(varTjDatoUS_man)
                            sqlDatoSlut = year(varTjDatoUS_son) &"/"& month(varTjDatoUS_son) & "/" & day(varTjDatoUS_son)
                            call cflow_hl_timer(usemrn, sqlDatoStart, sqlDatoSlut)
                        end select%>

                        <%
                           showKommega = 1
                           if cint(stempelurOn) = 1 AND (lto = "foa" OR lto = "xintranet - local") AND showKommega = 1 then

                             'if session("mid") = 1 then%>
                            <br />
                                  <div class="panel-group accordion-panel" id="accordion-paneled">
                                       <div class="panel panel-default">
                                            <div class="panel-heading">

                                                  <h4 class="panel-title"><a class="accordion-toggle" data-toggle="collapse" href="#K0"><b>Komme/G� tid</b> 
                                                <span style="font-size:85%; width:60%;">- <%=meTxt %></span>
                            
                                                 <span style="float:right;"> 

                                                        <% 
                                                        showKommega = 1 'sl� den til n�r den er klar til CFLOW    
                                                        if cint(stempelurOn) = 1 AND cint(showKommega) = 1 then 'AND session("mid") = 1%>
                                                        <span id="sp_sumlontimer_dag_<%=datepart("w", varTjDatoUS_man, 2,2)%>">&nbsp;</span> vs.
                                                        <%end if %>


                                                <span id="sp_sumtimer_dag_<%=datepart("w", varTjDatoUS_man, 2,2)%>">&nbsp;</span></a>
                         
                                                </span>
                                                </h4>

                                                </div><!-- Panel - heading -->

                              
                                                
                              <div id="K0" class="panel-collapse collapse">
                                
                                <div class="panel-body">

                                <div class="col-lg-5 well">
                                <b><a href="logindhist_2011.asp?usemrn=<%=usemrn%>&varTjDatoUS_man=<%=varTjDatoUS_man %>" target="_blank">Komme/G� tid:</a></b>
                                <%
                                stdatoSQL = year(varTjDatoUS_man) &"/"& month(varTjDatoUS_man) & "/" & day(varTjDatoUS_man)
                                sldatoSQL = year(varTjDatoUS_son) &"/"& month(varTjDatoUS_son) & "/" & day(varTjDatoUS_son)
                                call logindhistorik_week_60_100(usemrn, 2, stdatoSQL, sldatoSQL) %>

                                </div>

                             
                                <% 
                                call thisWeekNo53_fn(now) 
                                thisWeekNo53_now = thisWeekNo53

                                call thisWeekNo53_fn(st_dato) 
                                thisWeekNo53_st_dato = thisWeekNo53

                                if nowdate = dateid AND (cint(thisWeekNo53_now) = cint(thisWeekNo53_st_dato)) then %>

                                <div class="col-lg-5 well-white">
                                <%


                                    if lto <> "tec" AND lto <> "esn" AND lto <> "cflow" AND lto <> "intranet - local" AND lto <> "foa" then 'SKAL MATCHE menu LOGUD %>
                                    <a class="btn btn-danger" href="<%=toSubVerPath14 %>stempelur.asp?func=redloginhist&medarbSel=<%=session("mid")%>&showonlyone=1&hidemenu=1&id=0&rdir=sesaba" target="_top" style="color:#FFFFFF;"><%=tsa_txt_435 %></a>
                                    <%else%>
                                    <a class="btn btn-danger" href="<%=toSubVerPath14 %>../sesaba.asp?fromweblogud=1" target="_top" style="color:#FFFFFF;"><%=tsa_txt_435 %></a> <br />
                                    <%end if%>

                                </div>
           
                                <%end if 'nowdate
                                
                                
                         %></div></div>
                                </div></div><%'Panel - Panel body
                
                        end if   
                        
                         %>

                        <style>
                          /*  #maintable td {
                                border-left:1px solid #8c8c8c !important;
                                
                            } */

                            #maintable tr td:last-child {
                               border-right:1px solid #d9d9d9;
                            } 

                            #maintable tr th:last-child {
                               border-right:1px solid #d9d9d9;
                            } 

                             #maintable tr td:first-child {
                               border-left:1px solid #d9d9d9;
                            } 

                            #maintable tr th:first-child {
                               border-left:1px solid #d9d9d9;
                            }

                            .sideborder {
                                border-left:1px solid #d9d9d9 !important;
                            }
             
                        </style> 
                           

                        <form action="../timereg/timereg_akt_2006.asp?func=db&<%=FMlink %>" method="post" id="regform">
                            
                           
                            <input type="hidden" id="jq_varTjDatoUS_man" value="<%=varTjDatoUS_man%>" />
                            <input type="hidden" name="FM_medid" value="<%=medid %>" />
                            <input id="usemrn" value="<%=usemrn %>" type="hidden" />
                            <input type="hidden" id="Hidden4" name="FM_dager" value="7"/>
                           
                            <input type="hidden" id="" name="FM_vistimereltid" value="0"/>
                            <input type="hidden" id="Hidden5" name="year" value="<%=year(now) %>"/>
                            <input type="hidden" value="0" name="extsysId" />

                        <script src="js/fav_mat.js" type="text/javascript"></script>
                        <table id="maintable" class="table table-striped table-bordered dataTable ui-datatable" style="border:none; border-top:1px solid #d9d9d9; border-bottom:1px solid #d9d9d9;">
                            
                            <%
                            showTd = "normal"
                            headerSize = ""
                            jobtdSize = "200px"
                            if browstype_client = "ip" then
                                showTd = "none"
                                headerSize = "125%"
                                jobtdSize = "150px"
                            end if
                            %>

                            <thead>
                                <tr>
                                    <th style="font-size:<%=headerSize%>;"><%=favorit_txt_002 %></th>
                                    <th style="display:<%=showTd%>;"><%=favorit_txt_003 %></th>

                                    <%
                                            perInterval = 6 'dateDiff("d", varTjDatoUS_man, varTjDatoUS_son, 2,2) 
                                            perIntervalLoop = perInterval
                                            showTd = "normal"
                                            dim thisDayHellig, nomrtimerdagfavorit

                                            redim thisDayHellig(perInterval), nomrtimerdagfavorit(perInterval) 

                                            for l = 0 to perIntervalLoop 
        

                                            if browstype_client = "ip" then
                                                if l = (dayNumberChoosen - 1) then
                                                    showTd = "normal"
                                                else
                                                    showTd = "none"
                                                end if
                                            end if



                                            thisDayHellig(l) = 0

                                            if l = 0 then
                                            varTjDatoUS_use = varTjDatoUS_man
                                            else
                                            varTjDatoUS_use = dateAdd("d", l, varTjDatoUS_man)
                                            end if 

                                            showdate = Right("0" & DatePart("d",varTjDatoUS_use,2,2), 2) & "-" & Right("0" & DatePart("m",varTjDatoUS_use,2,2), 2) & "-" & DatePart("yyyy",varTjDatoUS_use,2,2)

                                            'showweekdayname = weekdayname(weekday(varTjDatoUS_use, 1))
                                            'daynamenum = weekday(varTjDatoUS_use,1)
                                            bgcol = ""
                                            select case l 
                                            case 0
                                            daynameword = favorit_txt_006
                                            case 1
                                            daynameword = favorit_txt_007
                                            case 2
                                            daynameword = favorit_txt_008
                                            case 3
                                            daynameword = favorit_txt_009
                                            case 4
                                            daynameword = favorit_txt_010
                                            case 5
                                            daynameword = favorit_txt_011
                                            bgcol = "#CCCCCC"
                                            case 6
                                            daynameword = favorit_txt_012
                                            bgcol = "#CCCCCC"
                                            end select

                                            'daynameword = WeekDayName(daynamenum,true)

                                            'response.Write weekdayname(weekday(varTjDatoUS_use, 1))

                                            if formatdatetime(now, 2) = formatdatetime(varTjDatoUS_use, 2) then
                                            bgcol = "lightpink"
                                            else
                                            bgcol = bgcol
                                            end if

                                            call helligdage(varTjDatoUS_use, 0, lto, usemrn)

                                            if cint(erHellig) = 1 then 
                                            bgcol = "#CCCCCC"
                                            thisDayHellig(l) = 1
                                            else
                                            bgcol = bgcol
                                            end if

                                            '*** Norm ***'
                                            call normtimerPer(usemrn, varTjDatoUS_use, 0, 0)
                                            nomrtimerdagfavorit(l) = ntimper 
                                            %>
                                                <th style="width:75px; vertical-align:bottom; background-color:<%=bgcol%>; display:<%=showTd%>; font-size:<%=headerSize%>;"><%=UCase(Left(daynameword,1)) & Mid(daynameword,2) %> <br />
                                                    <%=showdate %>
                                                    <%if cint(erHellig) = 1 then %>
                                                    <br /><span style="font-size:75%;"><%=helligdagnavnTxt %></span>
                                                    <%end if %>
                                                </th>
                                            <%

                                            next



                                           showTd = "normal"
                                           if browstype_client = "ip" then
                                                showTd = "none"
                                           end if
                                    %>
                                    <th style="text-align:center; width:55px; display:<%=showTd%>;"><%=favorit_txt_013 %></th>
                                    <th style="display:<%=showTd%>;"></th>
                                </tr>
                            </thead>

                            <tbody>

                                <%
                                     
                                     favoriter = 0
                                     lastaktid = 0
                                     i = 1250 'antal aktiviteter ialt.

                                     'Dim jobid, aktid, medarb, afakbar
                                     Redim jobid(i), aktid(i), medarb(i), afakbar(i), ajobnavn(i), ajobnr(i), akundenavn(i), akundenr(i), ajobknr(i), ajobans1(i)
                                     Redim ajobstartdato(i), ajobslutdato(i), abudgettimer(i), abeskrivelse(i)

                                     select case lto
                                     case "wap"
                                      strAkrORderBy = "kkundenavn, jobnavn, jobnr, a.sortorder, a.navn"

                                            'if session("mid") = 1 or session("mid") = 21 or session("mid") = 8 then 'LK skal kunne se korrektion p� alle medarbejdere
                                            'strSQlaktspecial = " OR tu.id = 200"
                                            'else
                                            strSQlaktspecial = ""
                                            'end if

                                     case "cflow"

                                       strAkrORderBy = "kkundenavn, a.fase, risiko DESC, jobnavn, jobnr, a.sortorder, a.navn"
                                      strSQlaktspecial = ""

                                     case else
                                     
                                      strAkrORderBy = "kkundenavn, jobnavn, jobnr, a.fase, a.navn" 'a.sortorder
                                      strSQlaktspecial = ""
                                     end select

                                     
                                    '*** Dobbelttjekker der er rettigheder til job og aktiviteter stadigv�k (man kan v�re fjernet fra projektgruppen efterf�lgende)
                                     call hentbgrppamedarb(medid)
                                     

                              

                                     '*** MAIN ***
                                     StrSqlfav = "SELECT j.id AS jid, medarb, jobid, aktid, forvalgt_af, jobstatus, fakturerbar, jobnavn, jobnr, "_
                                     &" j.jobknr, jobans1, jobstartdato, jobslutdato, j.budgettimer, j.beskrivelse FROM timereg_usejob AS tu "_
                                     &" LEFT JOIN job j ON (j.id = jobid "& strPgrpSQLkri &") "_
                                     &" LEFT JOIN aktiviteter a ON (a.id = aktid AND "& strSQLkri3 &") "_
                                     &" LEFT JOIN kunder k ON (kid = jobknr) "_
                                     &" WHERE (medarb = "& medid & " AND favorit <> 0) "& strSQlaktspecial &" AND (jobstatus = 1 OR jobstatus = 3) AND a.aktstatus = 1 "& strPgrpSQLkri &" AND "& strSQLkri3 &" GROUP BY a.id ORDER BY "& strAkrORderBy &"" 
                                     
                                        'if session("mid") = 4 then
                                        'Response.write StrSqlfav
                                        'Response.flush
                                        'end if

                                    

                                     i = 0
                                     hlfundet = 0
                                     strAktIswrt = " AND aktid <> 0"
                                     oRec.open StrSqlfav, oConn, 3
                                     while not oRec.EOF


                                        '*** Dobbelttjekekr om det er et tilbud - kun salgs aktiviteter vises
                                        jobstatusTjk = 1
                                        jobstatusTjk = oRec("jobstatus")
                                        'strSQLtilbud = "SELECT jobstatus FROm job WHERE id = "& oRec("jobid")
                                        'oRec5.open strSQLtilbud, oConn, 3
                                        'if not oRec5.EOF then

                                        'jobstatusTjk = oRec5("jobstatus")

                                        'end if
                                        'oRec5.close

                                        afakturerbar = 0
                                        
                                        'if cint(jobstatusTjk) = 3 then 'tilbud
                                        if isNULL(oRec("fakturerbar")) <> true then
                                        afakturerbar = oRec("fakturerbar")
                                        else
                                        afakturerbar = 0
                                        end if
                                        
                                        'strSQLtilbudA = "SELECT fakturerbar FROM aktiviteter WHERE id = "& oRec("aktid")
                                        'oRec5.open strSQLtilbudA, oConn, 3
                                        'if not oRec5.EOF then

                                        'afakturerbar = oRec5("fakturerbar")

                                        'end if
                                        'oRec5.close
                                        'end if

                                                '** dobbeltjek KUN aktivejob + tilbud && salgsaktivitet
                                                if cint(jobstatusTjk) = 1 OR (cint(jobstatusTjk) = 3 AND cint(afakturerbar) = 6) then  

                                                 jobid(i) = oRec("jobid")
                                                 aktid(i) = oRec("aktid")
                                                 medarb(i) = oRec("medarb")
                                                 'response.Write jobid(i)
                                                 afakbar(i) = oRec("fakturerbar")
                                                 ajobnavn(i) = oRec("jobnavn") 
                                                 ajobnr(i) = oRec("jobnr")  
                                                 ajobknr(i) = oRec("jobknr")
                                                 ajobans1(i) = oRec("jobans1")

                                                ajobstartdato(i) = oRec("jobstartdato") 
                                                 ajobslutdato(i) = oRec("jobslutdato")
                                                 abudgettimer(i) = oRec("budgettimer")
                                                abeskrivelse(i) = oRec("beskrivelse")

                                                 if lastaktid <> oRec("aktid") then


                                                 strAktIswrt = strAktIswrt &" AND aktid <> " & oRec("aktid")
                                                 i = i + 1
                                                 end if

                                                 lastaktid = oRec("aktid")
                                     
                                                 favoriter = favoriter + 1


                                                 end if
                                     

                                     oRec.movenext
                                     wend
                                     oRec.close







                                            '*************************************************************************************************
                                            '*** Henter ALLE med ressource forecase Oko, WWF timer for valgte medarb    *******************
                                            '*************************************************************************************************
                                            

                                            if cint(aktBudgettjkOnViskunmbgt) = 1 then 'AND (lto = "sduuas" OR lto = "xwwf" OR lto = "intranet - local") then 'viser alle med forecast p�

                                            datoKrionly = 1
                                            call ressourcetimerTildelt(now, 0, 0, usemrn, datoKrionly)

                                            i = i
                                            hlfundet = hlfundet 
                                            strSQLr = "SELECT r.jobid, SUM(r.timer) AS restimer, j.jobnavn, j.jobnr, j.jobknr, "_
                                            & "jobstatus, fakturerbar, a.id AS aktid, r.medid AS medarb, jobans1, jobstartdato, jobslutdato, j.budgettimer, j.beskrivelse FROM ressourcer_md AS r "_
                                            &" LEFT JOIN job AS j ON (j.id = r.jobid) "_
                                            &" LEFT JOIN aktiviteter AS a ON (a.id = r.aktid) "_
                                            &" WHERE r.medid = "& usemrn & " AND r.aktid <> 0 AND a.navn IS NOT NULL AND (jobstatus = 1) AND aktstatus = 1 "& sqlBudgafg &" "& strAktIswrt &" GROUP BY r.jobid ORDER BY jobnavn LIMIT 50" 
                                            


                                            'if session("mid") = 1 then

                                            '    Response.write strSQLr
                                            '    Response.flush

                                            'end if


                                         
                                            oRec.open strSQLr, oConn, 3
                                            while not oRec.EOF

                                            '*** Dobbelttjekekr om det er et tilbud - kun salgs aktiviteter vises
                                                jobstatusTjk = 1
                                                jobstatusTjk = oRec("jobstatus")
                                      
                                                afakturerbar = oRec("fakturerbar")
                                       

                                                '** dobbeltjek KUN aktivejob + tilbud && salgsaktivitet
                                                if cint(jobstatusTjk) = 1 OR (cint(jobstatusTjk) = 3 AND cint(afakturerbar) = 6) then  

                                                 jobid(i) = oRec("jobid")
                                                 aktid(i) = oRec("aktid")
                                                 medarb(i) = oRec("medarb")
                                                 'response.Write jobid(i)
                                                 afakbar(i) = oRec("fakturerbar")

                                                 ajobnavn(i) = oRec("jobnavn") 
                                                 ajobnr(i) = oRec("jobnr")  
                                                 ajobknr(i) = oRec("jobknr")
                                                  ajobans1(i) = oRec("jobans1")

                                                 ajobstartdato(i) = oRec("jobstartdato") 
                                                 ajobslutdato(i) = oRec("jobslutdato")
                                                 abudgettimer(i) = oRec("budgettimer")
                                                 abeskrivelse(i) = oRec("beskrivelse")

                                                 if lastaktid <> oRec("aktid") then
                                                 i = i + 1
                                                 end if

                                                 lastaktid = oRec("aktid")
                                     
                                                 favoriter = favoriter + 1


                                                 end if
                                     

                                     oRec.movenext
                                     wend
                                     oRec.close  

                                            
                                     end if 'vis automatisik dem med forecast

                                   
                                    
                                    '*** MAIN Tjek for positiv - SKAL ALTID vises
                                    '***  POSITIV ligger i DROPDOWN og hendetes inde derfra
                                       





                                    '**** HTML MAIN Array loop

                                    i_end = i
                                    i = 0
                                    'response.write "<br> d" & i_end
                                    'response.Write "lastid: " & lastjobid 
                                    
                                     y = 0      
                                     alljobids = "0"
                                     lastJobnavn = ""
                                     lastjobids = 0
                                     for i = 0 to i_end - 1



                                        resTimerThisM = 0
                                        jobids = jobid(i)
                                        jobnavn = ajobnavn(i) & " ("& ajobnr(i) & ")"

                                        'response.Write "aktid1: " & aktid(i)
                                       
                                        'StrSQLjob = "SELECT id, jobnavn, jobnr, jobstartdato, jobslutdato, jobans1, jobans2, jobans3, jobans4, jobans5, jobknr, beskrivelse, budgettimer FROM job WHERE id ="& jobid(i)
                                        'oRec3.open StrSQLjob, oConn, 3
                                        'if not oRec3.EOF then
                                        'jobids = oRec3("id")
                                        'jobnavn = oRec3("jobnavn") & " ("& oRec3("jobnr") &")"
                                        'end if
                                        'oRec3.close
                                        

                                        strSQLkunde = "SELECT kundeans1, kkundenavn FROM kunder WHERE kid = "& ajobknr(i)
                                        oRec4.open strSQLkunde, oConn, 3
                                        if not oRec4.EOF then
                                        kundeans = oRec4("kundeans1")
                                        kundenavn = oRec4("kkundenavn")
                                        end if
                                        oRec4.close

                                        aktbudgettimer = 0
                                        StrSQLakt = "SELECT id, navn, beskrivelse, budgettimer, fakturerbar, fase FROM aktiviteter WHERE id ="& aktid(i)
                                        oRec2.open StrSqlakt, oConn, 3
                                        if not oRec2.EOF then
                                        
                                        aktNavn = oRec2("navn")
                                        TaktId = oRec2("id")
                                        aktbudgettimer = oRec2("budgettimer")
                                        fakturerbar = oRec2("fakturerbar")
                                        thisfase = oRec2("fase")
                                        'response.Write jobnavn

                                        end if
                                        oRec2.close



                                        
                                       



                                     select case lto
                                     case "wap"
                                     
                                        if aktNavn = "Kontor arbejdstid" then 'L�ge/tandl�ge
                                        %>
                                        <tr><td colspan="11" class="sideborder"><br />&nbsp;</td></tr>
                                        <%
                                        end if

                                     end select

                                    select case lto
                                     case "cflow"
                                     
                                        if fakturerbar = 30 then
                                        %>
                                        <tr><td colspan="11" style="background-color:snow;" class="sideborder"><br /><br />&nbsp;<b>Huldt & Lillevik timer</b></td></tr>
                                        <%
                                            'hlfundet = 1
                                        end if

                                     end select


                                         select case lto
                                         case "tia", "wap", "xcflow"
                                         usecollapsed = 0
                                         case else
                                         usecollapsed = 1
                                         end select
                                       
                                            
                                       if cdbl(jobids) = cdbl(lastjobids) AND cint(usecollapsed) = 1 then
                                         
                                         select case lto
                                         case "tia"
                                         collapsed = 0
                                         cls = "tr_jobid_0"
                                         case else
                                         collapsed = 1
                                         cls = "tr_jobid_"& jobids
                                         end select

                                        else
                                         collapsed = 0
                                         cls = "tr_jobid_0"

                                       end if


                                        if cint(collapsed) = 1 then
                                        tr_vzb = "hidden"
                                        tr_dsp = "none"
                                        else
                                        tr_vzb = "visible"
                                        tr_dsp = ""
                                        end if
                                        %>
                                       
                                        <tr style="visibility:<%=tr_vzb%>; display:<%=tr_dsp%>;" class="<%=cls%>">
                                             <input type="hidden" value="<%=jobids %>" name="FM_jobid" />
                                            <%if cint(hlfundet) <> 1 then %>
                                            <td style="vertical-align:top; width:200px;" class="sideborder">
                                            
                                               
                                                <%if (browstype_client <> "ipx") then %>


                                                    <%if lastJobnavn <> jobnavn then %>
                                                    <span style="font-size:10px;"><b><%=kundenavn %></b></span><br />
                                                
                                                    <%

                                                            if i < i_end AND cint(usecollapsed) = 1 then    
                                                
                                                                if jobid(i) = jobid(i+1) then %>
                                                                <!--  -->
                                                                <span id="sp_tr_jobid_<%=jobids %>" class="sp_tr_jobid fa fa-plus" style="color:#5582d2;"></span>&nbsp;<%=jobnavn %>
                                                                <%else %>
                                                                <%=jobnavn %>
                                                                <%end if %>

                                                            <%else %>
                                                            <%=jobnavn %>
                                                            <%end if %>

                                                    <a data-toggle="modal" href="#styledModalSstGrp20"><span id="jobinfo_<%=jobids %>" style="color:#8c8c8c; font-size:75%;" class="fa fa-file-text jobinfo"></span></a> 

                                                       
                                                
                                                    <%end if %>

                                                <%else %>
                                                <span style="font-size:10px;"><%=kundenavn %></span> <br />
                                                <span style="font-size:12px;"><%=jobnavn %></span> <br />
                                                <%=aktNavn %>

                                                <%end if %>


                                                 <%if (browstype_client = "ip") then %>
                                                 <br /><%=aktNavn %>
                                                 <%end if %>

                                                

                                                 <div id="jobmodal_<%=jobids %>" class="modal">
                                                     <div class="modal-content" style="width:400px; height:500px;">
                                                         <%
                                                            

                                                            strSQLjobnas = "SELECT mnavn FROM medarbejdere WHERE mid = "& ajobans1(i)
                                                            oRec4.open strSQLjobnas, oConn, 3
                                                            if not oRec4.EOF then
                                                            jobansvarlig = oRec4("mnavn")
                                                            end if
                                                            oRec4.close

                                                            strSQLkundeans = "SELECT mnavn FROM medarbejdere WHERE mid="& kundeans  
                                                            oRec4.open strSQLkundeans, oConn, 3
                                                            if not oRec4.EOF then
                                                            kundeansvarlig = oRec4("mnavn")
                                                            end if
                                                            oRec4.close

                                                            'strSQLrealiseret = "SELECT sum(timer) as timer FROM timer WHERE tjobnr ="& jobids & " GROUP BY tjobnr"
                                                            'oRec4.open strSQLrealiseret, oConn, 3
                                                            'if not oRec4.EOF then
                                                            ' realiserettimer = oRec4("timer")
                                                            'end if 
                                                            'oRec4.close

                                                         %>
                                                   

                                                        <div class="row">
                                                            <div class="col-lg-12">
                                                                <b><%=kundenavn %></b> - <%=jobnavn %>
                                                            </div>
                                                        </div>
                                                         <br /><br />

                                                 

                                                        <div class="row">
                                                             <div class="col-lg-4"><b><%=favorit_txt_022 %>:</b></div>
                                                             <div class="col-lg-5"><%=ajobstartdato(i)%></div>
                                                        </div>
                                                        <div class="row">
                                                             <div class="col-lg-4"><b><%=favorit_txt_023 %>:</b></div>
                                                             <div class="col-lg-5"><%=ajobslutdato(i)%></div>
                                                        </div>
                                                        <br />
                                                        <div class="row">
                                                            <div class="col-lg-4"><b><%=favorit_txt_024 %>:</b></div>
                                                            <div class="col-lg-5"><%=jobansvarlig %></div>
                                                        </div>
                                                        <div class="row">
                                                            <div class="col-lg-4"><b><%=favorit_txt_025 %>:</b></div>
                                                            <div class="col-lg-5"><%=kundeansvarlig %></div>
                                                        </div>
                                                        <br />

                                                        <div class="row">
                                                            <div class="col-lg-4"><b><%=favorit_txt_026 %>:</b></div>
                                                            <div class="col-lg-5"><%=abudgettimer(i)%> t.</div>
                                                        </div>
                                                         <!--
                                                        <div class="row">
                                                            <div class="col-lg-4"><b><%=favorit_txt_027 %>:</b></div>
                                                            <div class="col-lg-5"><%=realiserettimer %> t.</div>
                                                        </div>
                                                         -->
                                                        <br />

                                                        <div class="row">
                                                            <div class="col-lg-4"><b><%=favorit_txt_028 %>:</b></div>
                                                        </div>
                                                        <div class="row">
                                                            <div class="col-lg-12"><div class="form-control input-small" style="height:100px; overflow-y:auto;"><%=abeskrivelse(i) %></div></div>
                                                        </div>                                                    

                                                     </div>
                                                 </div>

                                            </td>
                                            <%end if 'hlfundet = 0 %>




                                            <%if cint(hlfundet) = 0 then %>
                                            <td style="vertical-align:middle; width:195px; display:<%=showTd%>;">
                                            <%else %>
                                            <td style="vertical-align:middle;" colspan="2">
                                            <%end if %>

                                                <input type="hidden" value="<%=TaktId %>" name="FM_aktivitetid" />
                                                <%if len(trim(thisfase)) <> 0 then %>
                                                <span style="font-size:9px;"><%=replace(thisfase, "_", "") %></span><br />
                                                <%end if %>

                                                <%=aktNavn %>

                                                <%if cint(hlfundet) <> 1 AND cint(visAktlinjerSimpel) <> 1 AND ( cint(afakbar(i)) = 1 OR cint(afakbar(i)) = 2 ) then %>
                                                <span id="modal_<%=TaktId %>" class="fa fa-chevron-down pull-right picmodal" style="color:#8c8c8c"></span>                                               
                                                <div id="myModal_<%=TaktId %>" style="display:none">
                                                <!--<div>-->
                                                    <%
                                                        timerforalle = 0
                                                        if cint(visAktlinjerSimpel_realtimer) = 1 OR cint(visAktlinjerSimpel_medarbrealtimer) = 1 then
                                                            StrSqltimerialt = "SELECT sum(Timer) as timer FROM timer WHERE TAktivitetId = "& TaktId & " GROUP BY TAktivitetId"
                                                            oRec6.open StrSqltimerialt, oConn, 3
                                                            if not oRec6.EOF then
                                                            timerforalle = oRec6("timer")
                                                            end if
                                                            oRec6.close
                                                        end if

                                                        resTimerThisM = 0
                                                        'if cint(aktBudgettjkOnViskunmbgt) = 1 then 
                                                        useDateStSQL = year(varTjDatoUS_man) &"/"& month(varTjDatoUS_man) &"/"& day(varTjDatoUS_man)
                                                        datoKrionly = 0

                                                        'if session("mid") = 1 then
                                                        'Response.write "HER aktBudgettjkOnViskunmbgt: "& aktBudgettjkOnViskunmbgt & " aktBudgettjkOn: " & aktBudgettjkOn
                                                        'Response.write "j:"& jobids &",a "& TaktId &", m " & medid & " visAktlinjerSimpel_restimer: " & visAktlinjerSimpel_restimer
                                                        'end if
                                                        
                                                        if cint(visAktlinjerSimpel_restimer) = 1 then
                                                            call ressourcetimerTildelt(useDateStSQL, jobids, TaktId, medid, datoKrionly)
                                                        end if

                                                        'end if
                                                        timertotal = 0
                                                        if cint(visAktlinjerSimpel_medarbrealtimer) = 1 then
                                                            StrSqltotaltimer = "SELECT TAktivitetId, sum(timer) as timer FROM timer WHERE TAktivitetId = "& TaktId &" AND tmnr ="& medid & " GROUP BY TAktivitetId "
                                                            oRec5.open StrSqltotaltimer, oConn, 3
                                                            if not oRec5.EOF then
                                                
                                                            timertotal = oRec5("timer")
                                                       
                                                            end if
                                                            oRec5.close 
                                                        end if


                                                        if timertotal <> 0 then
                                                        timertotal = timertotal
                                                        else
                                                        timertotal = 0
                                                        end if

                                                        if resTimerThisM <> 0 then
                                                        resTimerThisM = resTimerThisM
                                                        else
                                                        resTimerThisM = 0
                                                        end if

                                                        if aktbudgettimer <> 0 then
                                                        aktbudgettimer = aktbudgettimer
                                                        else
                                                        aktbudgettimer = 0
                                                        end if

                                                        if timerforalle <> 0 then
                                                        timerforalle = timerforalle
                                                        else
                                                        timerforalle = 0
                                                        end if

                                                        if timerforalle > aktbudgettimer then
                                                           txtcolor = "red"
                                                        else
                                                           txtcolor = "green"
                                                        end if


                                                        if lto = "wap" AND TaktId = 28 then

                                                                'if session("mid") = 1 then

                                                                    'datoMan = now
                                                                    'datoSon = now

                                                                    maxflextimertotal = 0
                                                              
                                                                    mthstart_dato = year(datoMan) & "-" & month(datoMan) & "-1" 

                                                                    

                                                                    antaldageiM = dateadd("m", 1, datoMan)
                                                                    antaldageiM = "1-" & month(antaldageiM) &"-"& year(antaldageiM)
                                                                    antaldageiM = dateadd("d", -1, antaldageiM)
                                                                    mthslut_dato = year(antaldageiM) & "-" & month(antaldageiM) & "-"& day(antaldageiM) 

                                                                    StrSqltotaltimer = "SELECT TAktivitetId, sum(timer) as timer FROM timer WHERE TAktivitetId = 28 AND tmnr ="& medid & " AND tdato BETWEEN '"& mthstart_dato &"' AND '"& mthslut_dato &"' GROUP BY TAktivitetId "
                                                                    'if session("mid") = 1 then
                                                                    'response.write StrSqltotaltimer
                                                                    'response.flush
                                                                    'end if
                                                                    oRec5.open StrSqltotaltimer, oConn, 3
                                                                    if not oRec5.EOF then
                                                
                                                                    maxflextimertotal = oRec5("timer")
                                                       
                                                                    end if
                                                                    oRec5.close 
                                                       

                                                            if maxflextimertotal <> 0 then
                                                            maxflextimertotal = maxflextimertotal
                                                            else
                                                            maxflextimertotal = 0
                                                            end if

                                                            'end if

                                                        end if

                                                    %>

                                                         <%if lto = "wap" AND TaktId = 28 then 'flexhome %>
                                                            
                                                    
                                                           <%if formatnumber(resTimerThisM) <> 0 then%> 
                                                            
                                                            <span style="font-size:75%; line-height:9px;">
                                                            <%=monthname(month(mthstart_dato)) %><br />    
                                                            Forecast: <%=formatnumber(resTimerThisM, 2) %></span>
                                                            <%end if %>

                                                            <br /><span style="font-size:75%;">Afholdt: <%=formatnumber(maxflextimertotal, 2) %><br />Saldo:<b> <%=formatnumber(resTimerThisM-maxflextimertotal, 2) %></b></span>

                                                        <%else %>

                                                    


                                                        <% if cint(visAktlinjerSimpel_timebudget) = 1 AND formatnumber(aktbudgettimer) <> 0 then %>
                                                        <span style="font-size:75%; line-height:9px;"><%=favorit_txt_029 %>: 
                                                        <%=formatnumber(aktbudgettimer, 2) %></span><br />
                                                        <%end if %>
                                                        
                                                        <%if formatnumber(resTimerThisM) <> 0 then%> 
                                                        <span style="font-size:75%; line-height:9px;">Forecast: <%=formatnumber(resTimerThisM, 2) %></span><br /> 
                                                        <!-- (<%=visAktlinjerSimpel_restimer %> / <%=useDateStSQL %> ) -->
                                                        <%end if %>
                                                       
                                                        <%if cint(visAktlinjerSimpel_realtimer) = 1 OR cint(visAktlinjerSimpel_medarbrealtimer) = 1 then %>
                                                        <span style="font-size:75%; line-height:9px; color:<%=txtcolor%>;"><%=favorit_txt_030 %>: <%=formatnumber(timerforalle, 2) %></span>&nbsp;
                                                        <%end if %>

                                                         <%if cint(visAktlinjerSimpel_medarbrealtimer) = 1 then %>
                                                         <span style="font-size:75%;">(<%=favorit_txt_031 %>: <%=formatnumber(timertotal, 2) %>)</span>
                                                        <%end if %>

                                                       
                                                    
                                                        <%end if %>
                                                        
                                                    

                                                </div>
                                                <%end if'cint(hlfundet) <> 1 %>
                                            </td>
                                            
                                            <%
                                                
                                                for l = 0 to 6
                                                
                                                     
                                                    timerKom = ""
                                                    fmborcl = ""
                                                    godkendtstatus = 0
                                                    timerdag = ""
                                                    overfort = 0
                                                    origin = 0
                                                    showTd = "normal"
                                                    y = y + 1
                                                    'response.Write y

                                                    if browstype_client = "ip" then
                                                        if l = (dayNumberChoosen - 1) then
                                                            showTd = "normal"
                                                            verticalAlign = "middle"
                                                        else
                                                            showTd = "none"
                                                            verticalAlign = ""
                                                        end if
                                                    end if

                                                    if l = 0 then
                                                        timerdato = varTjDatoUS_man
                                                    else
                                                        timerdato = dateAdd("d",1,timerdato)
                                                    end if




                                                     varTjDato_ugedag = day(timerdato) & "-" & month(timerdato) & "-" & year(timerdato)
                                                    

                                                     'if session("mid") = 1 then
                                                     '*** Er uge lukket ***
                                                        'varTjDato_ugedag_erugeAfslutte = formatdatetime(varTjDato_ugedag, 2)
                                                        call thisWeekNo53_fn(varTjDato_ugedag)
                                                        thisW = thisWeekNo53 'datepart("ww", varTjDato_ugedag, 2,2)
                                                        thisY = datepart("yyyy", varTjDato_ugedag, 2,2) 
                                                        'Response.write "<br>AHR FAv: "& varTjDato_ugedag &"<br> " '& thisW & " varTjDato_ugedag:  "& varTjDato_ugedag & " SmiWeekOrMonth: " & SmiWeekOrMonth
                                                        call erugeAfslutte(thisY, thisW, medid, SmiWeekOrMonth, 0, varTjDato_ugedag)
                                                     'end if 


                                                     '**** Er periode lukket via l�nk�rsel **''
                                                     call lonKorsel_lukketPer(varTjDato_ugedag, job_internt, medid) 'Hr -2 job bliver kun lukket ifh.- l�nperiode

                                                  
                                                    '*** tjekker om uge er afsluttet / lukket / l�nk�rsel
                                                    call tjkClosedPeriodCriteria(varTjDato_ugedag, ugeNrAfsluttet, usePeriod, SmiWeekOrMonth, splithr, smilaktiv, autogk, autolukvdato, lonKorsel_lukketIO, ugegodkendt)
                       


                                                    '***** LAST faktdato ***************************************
		                                            lastFakdato = "01/01/2002"
		                                            strSQL = "SELECT fakdato FROM fakturaer WHERE jobid = "& jobids &" AND faktype = 0 ORDER BY fakdato DESC LIMIT 0,1"
		        
		       
		                                            oRec.open strSQL, oConn, 3
		                                            if not oRec.EOF then
		                                            lastFakdato = oRec("fakdato")
		                                            end if
		                                            oRec.close

                                                    '***********************************************************


                                                    timerdatoSQL = year(timerdato) & "-" & month(timerdato) & "-" & day(timerdato) 
                                                    
                                                     StrSQLtimer = "SELECT TAktivitetId, sum(timer) as Timer, extsysId, tdato, Timerkom, origin, tjobnr, j.risiko, godkendtstatus, overfort FROM timer t "_
                                                     &" LEFT JOIN job j ON (j.jobnr = t.tjobnr) WHERE TAktivitetId = "& TaktId & " AND tdato = "& "'" & timerdatoSQL & "' AND tmnr ="& medid & " GROUP BY tmnr"
                                                

                                                     oRec4.open StrSQLtimer, oConn, 3
                                                     if not oRec4.EOF then
                                                
                                                     Stryear = oRec4("tdato")
                                                     if cdbl(oRec4("Timer")) <> 0 then
                                                     timerdag = oRec4("Timer")
                                                     else
                                                     timerdag = ""
                                                     end if

                                                     extsysid = oRec4("extsysId")
                                                     timerkcoment = oRec4("Timerkom")
                                                     origin = oRec4("origin")
                                                     godkendtstatus = oRec4("godkendtstatus")
                                                     job_internt = oRec4("risiko")
                                                     overfort = oRec4("overfort")
                                                     

                                                     timerKom = oRec4("Timerkom")    

                                                      
                                                     end if
                                                      oRec4.close

                                                      todaydate = DatePart("yyyy",Date, 2,2) &"-"& Right("0" & DatePart("m",Date,2,2), 2) &"-"& Right("0" & DatePart("d",Date,2,2), 2)


                                                     

                                                     '** enkelt timer godkendt/tentativ/afvist
                                                        select case cint(godkendtstatus)
                                                        case 0
                                                        fmborcl = "1px #CCCCCC solid" 
                                                        case 1
                                                        fmborcl = "1px yellowgreen solid"
                                                        case 3
                                                        fmborcl = "1px orange solid" 
                                                        case 2
                                                            'maxl = 5
	                                                        fmbgcol = "#FFFFFF"
	                                                        fmborcl = "1px red solid"
                                                        end select

                                                     
                                                            if cint(level) <> 1 then
	                                                            'maxl = 0
                                                                fmbgcol = fmbgcol
	                                                        else
	                                                            'maxl = maxl
                                                                fmbgcol = "#FFFFFF"
	                                                        end if


                                                        bgcol = ""
                                                        if (l = 5 OR l = 6 OR thisDayHellig(l) = 1) then 
                                                        bgcol = "#f1f1f1"
                                                        else
                                                        bgcol = bgcol
                                                        end if

                                                        if formatdatetime(now, 2) = formatdatetime(timerdato, 2) then
                                                        bgcol = "lavenderblush"
                                                        end if

                                                    %>
                                                        <td style="background-color:<%=bgcol%>; display:<%=showTd%>; vertical-align:<%=verticalAlign%>;">  
                                                               <div class="row">   

                                                                    <%
                                                                    'if session("mid") = 1 then
                                                                    'Response.write "origin:" & origin & " overfort:" & overfort & "<br>u_autogk:" & ugeerAfsl_og_autogk_smil & "<br><br>godkendtstatus:"&  godkendtstatus & " level:" & level & "<br>maxl:" & maxl 
                                                                    'end if
                                                                     %>

                                                            <!-- Til k�rsel -->
                                                            <input type="hidden" id="jq_kid_<%=y %>" value="<%=ajobknr(i) %>" />
                                                            <input type="hidden" id="jq_row_<%=y %>" value="<%=y %>" />
                                                            <input type="hidden" id="jq_flt_<%=y %>" value="<%=l %>" />
                                                            <input type="hidden" id="jq_tfa_<%=y %>" value="<%=fakturerbar %>" />
                                                            
                                                            
                                                            


                                                            <input type="hidden" name="FM_feltnr" value="<%=y %>" />
                                                            <input type="hidden" value="<%=timerdato %>" name="FM_datoer" />
                                                            <input type="hidden" value="dist" name="FM_destination_<%=y %>" />
                                                              <input type="hidden" name="FM_sttid" value="00:00"/>
                                                            <input type="hidden" name="FM_sltid" value="00:00"/>

                                                            <%
                                                            '*** �ben / Sp�rret for indtastninger hvis:
                                                            'Or cint(godkendtstatus) = 3
                                                            '*** Der m� redigeres hvis:
                                                            '*** Ugen ikke er godkendt
                                                            '*** Origin = Favorit eller gl. timereg.
                                                            '*** Timer ikke er overf�rt
                                                            '*** Timer ikke er godkendt dvs. afvist eller tentative m� gerne redigeres.
                                                            '*** Admin m� gerne redigere i godkendte. MEN ikke overf�rte.

                                                            '*** Lukket GL
                                                            'if (origin <> 0 AND origin <> 20 AND (lto <> "wap" AND cdbl(TaktId) <> 34)) OR _
                                                            'cint(overfort) = 1 OR _
                                                            '(((cint(ugeerAfsl_og_autogk_smil) = 1 AND cint(godkendtstatus) <> 2) OR (cint(godkendtstatus) = 1)) AND level <> 1) _
                                                            'then 
                                                                
                                                                

                                                                '** �ben:
                                                                if cint(overfort) = 0 AND (origin = 0 OR origin = 20) AND cdate(lastfakdato) < cdate(timerdato) AND media <> "print" AND _
                                                                (( (ugeerAfsl_og_autogk_smil = 0 AND cint(godkendtstatus) <> 1) _
                                                                OR (ugeerAfsl_og_autogk_smil = 1 AND cint(godkendtstatus) = 2) ) _
                                                                OR ((ugeerAfsl_og_autogk_smil = 1 OR ugeerAfsl_og_autogk_smil = 0) AND level = 1) _
                                                                OR (lto = "wap" AND cdbl(TaktId) = 34)) then%>

                                                                  

                                                                <%if browstype_client <> "ip" then %>
                                                                <div class="col-lg-9" style="padding-right:2px!important"><input type="text" class="form-control input-small timer_flt" id="FM_timer_<%=y%>" name="FM_timer" value="<%=timerdag %>" style="width:55px; border:<%=fmborcl%>;" /></div>
                                                                  
                                                                   <input type="hidden" id="FM_norm_<%=y%>" value="<%=nomrtimerdagfavorit(l)%>" />
                                                                <div class="col-lg-2" style="padding-left:3px!important;"><span id="modal_<%=y%>" class="kommodal">+</span></div>
                                                                <%else %>
                                                                <div class="col-lg-12">
                                                                    <div class="form-inline">
                                                                        <input type="text" class="form-control timer_flt" id="FM_timer_<%=y%>" name="FM_timer" value="<%=timerdag %>" style="width:80%; display:inline-block; border:<%=fmborcl%>;" />
                                                                        <span id="modal_<%=y%>" style="font-size:150%" class="kommodal">+</span>
                                                                    </div>
                                                                </div>
                                                                <%end if %>       


                                                           

                                                            <%else '** Lukket for redigering
                                                                %>
                                                                <input type="hidden" name="FM_timer" value=""/>

                                                                <%if browstype_client <> "ip" then %>
                                                                    <div class="col-lg-9" style="padding-right:2px!important"><input type="text" class="form-control input-small" style="width:55px; border:<%=fmborcl%>;" value="<%=timerdag %>" readonly /></div>
                                                                   
                                                                    <%if cint(ugeerAfsl_og_autogk_smil) = 0 AND origin <> 0 AND origin <> 20 then 'Er tastest ind via andre medier og UGE ikke godkendt %>
                                                                    <br /><span style="font-size:75%; overflow:hidden; white-space:nowrap; padding-left:15px;"><a style="color:dimgrey;" href="ugeseddel_2011.asp?usemrn=<%=medid %>&varTjDatoUS_man=<%=varTjDatoUS_man %>"><u><%=favorit_txt_014 %></u></a></span>
                                                                    <%end if %>

                                                                <%else  %>

                                                                   <div class="col-lg-12" style="padding-right:2px!important"><input type="text" class="form-control" style="width:80%; border:<%=fmborcl%>;" value="<%=timerdag %>" readonly /></div>
                                                                   <!--
                                                                   <div class="col-lg-12" style="text-align:center; font-size:50%"><span><a style="color:dimgrey;" href="ugeseddel_2011.asp?usemrn=<%=medid %>&varTjDatoUS_man=<%=varTjDatoUS_man %>"><%=favorit_txt_014 %></a></span></div>
                                                                   -->
                                                                <%end if %>                            

                                                            <%end if %>
                                                            </div>



                                                                <%'** KOMMENTAR **'
                                                                if browstype_client <> "ip" then 
                                                                    kommentarWidth = ""
                                                                    kommentarHeight = ""
                                                                    kommentarTxtSize = ""
                                                                    kommentarRows = "18"
                                                                else
                                                                    kommentarWidth = "300px !important"
                                                                    kommentarHeight = "300px !important"
                                                                    kommentarTxtSize = "125%"
                                                                    kommentarRows = "9"
                                                                end if
                                                                %>
                                                            
                                                                <div id="kommentarmodal_<%=y%>" class="modal">
                                                                        <div class="modal-content" style="width:<%=kommentarWidth%>; height:<%=kommentarHeight%>;">                                                                                                                        
                                                                            <div class="row">
                                                                                <div class="col-lg-2"><b style="font-size:<%=kommentarTxtSize%>;"><%=favorit_txt_032 %>:</b></div>
                                                                            </div>
                                                                            <div class="row">
                                                                                <div class="col-lg-12"><textarea rows="<%=kommentarRows %>" id="FM_kom_<%=y %>" name="FM_kom_<%=y %>" class="form-control input-small"><%=timerKom %></textarea></div>
                                                                            </div>


                                                                            <%
                                                                            showExpences = 0    
                                                                            if cint(showExpences) = 1 then %>
                                                                            <br /><br />

                                                                            <div class="panel-group accordion-panel" id="accordion-paneled" style="display:none">
                                                                    
                                                                            <div class="panel panel-default">
                                                                                <div class="panel-heading">
                                                                                    <h6 class="panel-title"><a class="accordion-toggle" data-toggle="collapse" data-target="#collapse_<%=y %>">Udl�g</a></h6>
                                                                                </div>
                                                                                <div id="collapse_<%=y %>" class="panel-collapse collapse">  
                                                                                    
                                                                                  

                                                                                    <input type="hidden" id="mat_id" value="0" />
                                                                                    <input type="hidden" id="mat_jobid_<%=y %>" value="<%=jobids %>" />
                                                                                    <input type="hidden" id="mat_dato_<%=y %>" value="<%=todaydate%>" />
                                                                                    <input type="hidden" id="mat_editor" value="<%=medarbejdernavn %>" />                                                                                  
                                                                                    <input type="hidden" id="mat_userid" value="<%=medid %>" />
                                                                                    <input type="hidden" id="mat_forbrugsdato_<%=y %>" value="<%=timerdato %>" />
                                                                                    <input type="hidden" id="mat_serviceaft" value="0" />
                                                                                    <input type="hidden" id="mat_endhed_<%=y %>" value="Stk." />
                                                                                    <input type="hidden" id="mat_aktid_<%=y %>" value="<%=Taktid %>" />
                                                                                    <input type="hidden" id="mat_bilagsnr" value="" />
                                                                                    <input type="hidden" id="mat_varenr" value="0" /> 

                                                                                                                                                          
                                                                                    <div class="panel-body">
                                                                                        
                                                                                        <div class="row" id="error_felt_<%=y %>" style="visibility:hidden">
                                                                                            <span id="error_txt_<%=y %>" class="col-lg-12" style="color:red"></span>
                                                                                        </div>
                                                                                        
                                                                                        <div class="row">
                                                                                            <div class="col-lg-4">Antal:</div>
                                                                                            <div class="col-lg-4"><input type="text" value="1" id="mat_antal_<%=y %>" class="form-control input-small" /></div>
                                                                                        </div>
                                                                                        <div class="row">
                                                                                            <div class="col-lg-4">Mat. navn:</div>
                                                                                            <div class="col-lg-4"><input type="text" value="" id="mat_navn_<%=y %>" class="form-control input-small" /></div>
                                                                                        </div>
                                                                                        <div class="row">
                                                                                            <div class="col-lg-4">Indk�bspris:</div>
                                                                                            <div class="col-lg-3"><input type="text" value="" id="mat_kobpris_<%=y %>" class="form-control input-small" /></div>
                                                                                            <div class="col-lg-4">
                                                                                                <select name="FM_valuta" id="mat_valuta<%=y %>" class="form-control input-small">
		                                                                                            <!--<option value="0"><=tsa_txt_229 %></option>-->
		                                                                                            <%
		                                                                                            strSQL3 = "SELECT id, valutakode, grundvaluta FROM valutaer ORDER BY valutakode"
    		
    		
		                                                                                            oRec5.open strSQL3, oConn, 3 
		                                                                                            while not oRec5.EOF 
    		
		                                                                                            if cint(valuta) = oRec5("id") then
		                                                                                            valGrpCHK = "SELECTED"
		                                                                                            else
		                                                                                            valGrpCHK = ""
		                                                                                            end if
		    
		   
		                                                                                            %>
		                                                                                            <option value="<%=oRec5("id")%>" <%=valGrpCHK %>><%=oRec5("valutakode")%></option>
		                                                                                            <%
		                                                                                            oRec5.movenext
		                                                                                            wend
		                                                                                            oRec5.close %>
		                                                                                        </select>
                                                                                            </div>                                                      
                                                                                        </div>
                                                                                        <div class="row">
                                                                                            <div class="col-lg-4">Gruppe:</div>
                                                                                            <div class="col-lg-4">                                                                                               
                                                                                                <select class="form-control input-small" name="gruppe" id="mat_gruppe_<%=y %>"><!-- onchange="beregnsalgsprisOTF(0)" -->
		                                                                                            <option value="0"><%=tsa_txt_200 %></option>
		                                                                                            <%
		                                                                                            strSQL = "SELECT id, navn, av FROM materiale_grp ORDER BY navn"
		                                                                                            oRec.open strSQL, oConn, 3 
		
		                                                                                                    while not oRec.EOF 
		
		                                                                                                    if cint(matgrp) = oRec("id") then
		                                                                                                    matgrpSel = "SELECTED"
		                                                                                                    else
		                                                                                                    matgrpSel = ""
		                                                                                                    end if
		
		
    		                                                                                                'matgrpVal = matgrpVal &  "<input id=""avagrpval_"&oRec("id")&""" name=""avagrpval_"&oRec("id")&""" type=""hidden"" value="& oRec("av") &" />"
    		
		
		                                                                                                %>
		                                                                                                <option value="<%=oRec("id")%>" <%=matgrpSel %>><%=oRec("navn")%>
		                                                                                                <%if level <= 2 OR level = 6 then %>
		                                                                                                &nbsp;(<%=oRec("av") %>%)
		                                                                                                <%end if %></option>

		                                                                                            <%
		                                                                                            oRec.movenext
		                                                                                            wend
		                                                                                            oRec.close %>
		                                                                                        </select>
                                                                                            </div>
                                                                                        </div>

                                                                                        <div class="row">
                                                                                            <div class="col-lg-4">Salgspris:</div>
                                                                                            <div class="col-lg-4"><input type="text" id="mat_salgspris_<%=y %>" class="form-control input-small" /></div>
                                                                                        </div>
                                                                                       
                                                                                        
                                                                                        
                                                                                        
                                                                                    <!------------- Henter allerede udl�g ---------------->                                                                  
                                                                                    <%
                                                                                    strSQLudlag = "SELECT matnavn, m.forbrugsdato, m.matsalgspris, v.valutakode FROM materiale_forbrug as m "_ 
                                                                                    & "LEFT JOIN valutaer as v ON (v.id = m.valuta) WHERE aktid ="& Taktid &" AND forbrugsdato ='"& timerdato &"'"
                                                                                    oRec8.open strSQLudlag, oConn, 3
                                                                                    'response.write strSQLudlag
                                                                                    'response.Flush
                                                                                    'if oRec8("forbrugsdato") <> 0 then
                                                                                    %>
                                                                                        <div class="row">
                                                                                            <div class="col-lg-6"><span style="font-size:75%">Indl�ste materialer:</span></div>
                                                                                        </div> 
                                                                                    <%
                                                                                    'end if

                                                                                    while not oRec8.EOF
                                                                                    strforbrugsdato = oRec8("Forbrugsdato")
                                                                                    matnavn = oRec8("matnavn")
                                                                                    matsalgspris = oRec8("matsalgspris")
                                                                                    valutakode = oRec8("valutakode")                                                                                                                                                                                                                        
                                                                                        %>                                                                                                                                
                                                                                        <div class="row"><div class="col-lg-12" style="font-size:75%"><%response.Write matnavn & " " & matsalgspris & " " & valutakode %></div></div>
                                                                                        <% 
                                                                                        oRec8.movenext
                                                                                        wend                                                                                 
                                                                                        oRec8.close 
                                                                                        %>                           
                                                                                        <div class="row">
                                                                                            <div class="col-lg-12 pull-right">
                                                                                              <!--  <a id="<%=Taktid %>" class="btn btn-sm btn-default pull-right mat_save"><b>Gem</b></a> -->
                                                                                                <a class="btn btn-sm btn-default pull-right mat_save matreg_sb" id="<%=y %>"><b>Gem</b></a>
                                                                                            </div>
                                                                                        </div>
                                                                              
                                                                                    </div>


                                                                                </div>
                                                                                </div>
                                                                                </div> 
                                                                            <%end if 'ShowExpences %>

                                                                         
                                                                                    <%if browstype_client = "ip" then  %>
                                                                                    <div class="row">
                                                                                        <div class="col-lg-12" style="text-align:center"><a class="btn btn-secondary closeCom"><b>Gem</b></a></div>
                                                                                    </div>
                                                                                    <%end if %>

                                                                               </div>
                                                                            </div>

                                                            <input type="hidden" name="FM_timer" value="xx"/>
                                                          
                                                        </td>
                                                 <%
                                                    

                                                    strforbrugsdato = 0


                                                next

                                                 'StrSQLtimer = "SELECT TAktivitetId, Timer FROM timer WHERE TAktivitetId ="& aktId
                                                
                                                 'oRec4.open StrSQLtimer, oConn, 3
                                                 'if not oRec4.EOF then
                                                
                                                    'timerdag = oRec4("Timer")

                                                 'end if 
                                                 'oRec4.close



                                                showTd = "normal"
                                                if browstype_client = "ip" then
                                                    showTd = "none"
                                                end if
                                                 
                                            %>

                                            <td style="text-align:center; vertical-align:middle; display:<%=showTd%>;">
                                                <%
                                                    ugestart_dato = year(datoMan) & "-" & month(datoMan) & "-" & day(datoMan)
                                                    ugeslut_dato = year(datoSon) & "-" & month(datoSon) & "-" & day(datoSon)

                                                    'response.Write ugestart_dato & "S�NDAG: " & ugeslut_dato
                                                    
                                                    StrSqlweektotal = "SELECT sum(timer) as timer FROM timer WHERE TAktivitetId ="& TaktId & " AND tdato BETWEEN '"& ugestart_dato &"' AND '"& ugeslut_dato &"' AND tmnr ="&medid 
                                                
                                                    oRec7.open StrSqlweektotal, oConn, 3
                                                    if not oRec7.EOF then
                                                    timerweektotal = oRec7("timer")
                                                    
                                                   
                                                    end if
                                                    oRec7.close  
                                                    
                                                    if timerweektotal <> 0 then
                                                    timerweektotal = timerweektotal
                                                    else
                                                    timerweektotal = 0
                                                    end if
  
                                                    
                                                %>
                                                <%=formatnumber(timerweektotal, 2) %>
                                            </td>

                                            <td style="vertical-align:middle; display:<%=showTd%>;">
                                                <%'*** DELETE / SLET / Fjern %>
                                                <%if cint(aktBudgettjkOnViskunmbgt) = 1 AND cint(resTimerThisM) <> 0 then %>
                                                &nbsp;
                                                <%else %>
                                                <a href="favorit.asp?id=<%=aktid(i)%>&FM_medid=<%=medid %>&varTjDatoUS_man=<%=varTjDatoUS_man%>&func=fjernfavorit"><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></a>
                                                <%end if %>
                                            </td>

                                        </tr>
                                        <%
                                       

                                        if cdbl(jobids) <> cdbl(lastjobids) then 
                                        alljobids = alljobids & "," & jobids 
                                        end if

                                        lastJobnavn = jobnavn
                                        lastjobids = jobids

                                    

                                    next


                                     
                                    %>
                                <input type="hidden" id="alljobids" value="<%=alljobids %>" />


                                  <!--  <tr>
                                        <td colspan="10"><input type="text" id="aktiviteter_sog" class="aktivitet_sog form-control input-small" />
                                            <select id="aktiviteterfelt" class="form-control input-small chbox_job" size="10" style="visibility:hidden; display:none;">
                                            <option><%=week_txt_007 %>..</option>
                                            </select>
                                        </td>
                                        
                                    </tr> -->

                                    <%
                                        next_akt_id = 0 
                                        for a = 0 to 10
                                        next_akt_id = next_akt_id + 1

                                        'response.Write "next_akt_id :" & next_akt_id
                                    %>
                                    <tr class="next_akt_id" style="visibility:hidden; display:none;" id="FN_akt_tilfojed_<%=next_akt_id %>">
                                        <td class="sideborder"><input type="text" value="" id="next_akt_jobid_<%=next_akt_id %>" class="form-control input-small" readonly /></td>
                                        <td><input type="text" value="" id="next_akt_aktid_<%=next_akt_id %>" class="form-control input-small" readonly /></td>


                                        <%for i = 0 to 6 %>
                                        <td><input type="text" style="width:75px;" class="form-control input-small" name="FM_timerdag" value="0" /></td>
                                        <%next %>
                                        <td>&nbsp</td>
                                        <td>&nbsp</td>
                                    </tr>    
                                    <%
                                        next 
                                    %>

                                   <!-- <tr>
                                        <td colspan="100" style="background-color:white; border-right:none; border-left:none;"></td>
                                    </tr> -->

                                    <tr style="border-bottom:inherit; display:<%=showTd%>;">  

                                        <!--<input type="hidden" value="0" name="FM_pa" />-->
                                        <input type="hidden" id="FM_jobid" value=""/>

                                        <td class="sideborder">
                                            
                                             
                                            
                                            <input type="text" class="FM_job form-control input-small" id="FM_job" value="" placeholder="<%=favorit_txt_015 %>" autocomplete="off" style="width:225px;"/>
                                            <!-- <div id="dv_job"></div> -->
                                            <select id="dv_job" class="form-control input-small chbox_job" size="10" style="visibility:hidden; display:none; width:225px;">
                                                <option><%=week_txt_007 %>..</option>
                                            </select>

                                            <%
                                            'e = 0
                                            'if e = 0 then
                                            if (level = 1 AND pa = 0) OR (level = 2 OR level = 6 ANd lto = "dencker") then %>
                                            <input type="checkbox" value="1" id="ign_projgrp" name="FM_ign_projgrp" <%=ign_projgrpCHK %> /> Ignorer projektgrupper
                                           
                                            <%
                                            findesIgProgrp = 1    
                                            else 
                                            findesIgProgrp = 0
                                            end if 
                                            %>

                                            <input type="hidden" value="<%=findesIgProgrp %>" id="ign_projgrp_hd" />

                                        </td>                                           
                             

                                        <td>
                                           <!--<input type="hidden" name="FM_aktivitetid" id="FM_aktid" value=""/>-->
                                           <!-- <input type="text" class="aktivitet_sog form-control input-small" id="FM_akt" value="" placeholder="<%=favorit_txt_016 %>" />-->
                                            <!--<div id="dv_akt"></div>--> 
                                            
                                          
                                            <select size="1" id="dv_akt" name="FM_aktivitetid" multiple class="form-control input-small chbox_akt" style="width:195px;">
                                                <option value="0" DISABLED style="height:23px;">..</option>
                                            </select> 
                                           
                                        </td>
                                        
                                        <td style="text-align:center;" colspan="9">
                                           <input type="hidden" id="FM_medid_id" class="form-control input-small" value="<%=medid %>" />
                                            <input type="hidden" value="1" id="next_akt_id" />                                                                                      
                                            <a class="tilfoj_akt btn btn-default btn-sm" id="1" style="width:100%"><b><%=favorit_txt_017 %></b></a>
                                            <div id="dv_akttil"></div>
                                        </td>                                     
                                    </tr>                                   

                                    <tr>
                                        <td colspan="100" style="background-color:white; border-right:none; border-left:none;"></td>
                                    </tr>

                                    <!-- Total -->
                                    <%call akttyper2009(2) %>

                                    <%

                                    '*** uge total / week total i bunden ****
                                    select case lto
                                    case "wap"
                                        visWeekTotals = 0
                                    case else
                                        visWeekTotals = 1
                                    end select



                                    if cint(visWeekTotals) = 1 then



                                         select case lto
                                            case "tec", "esn"
                                            aty_sql_realhours = " tfaktim <> 0"
                                            case "xdencker", "xintranet - local"
                                            aty_sql_realhours = " tfaktim = 1"
                                            case else
                                            aty_sql_realhours = aty_sql_realhours 
                                            '&""_
		                                    '& "tfaktim = 30 OR tfaktim = 31 OR tfaktim = 7 OR tfaktim = 11" 'sp�rg s�ren om det er rigtigt at  tfaktim skal v�re noget her
                                          end select

                                        timerdenneuge_dothis = 0
                                        SmiWeekOrMonth = 0
                                        'response.Write SmiWeekOrMonth
                                        call timerDenneUge(medid, lto, varTjDatoUS_man, aty_sql_realhours, timerdenneuge_dothis, SmiWeekOrMonth)
                            

                                        'if session("mid") = 1 then
                                        'response.Write "freTimer: "&  freTimer
                                        'end if
                            
                            
                                        'henter norm timer minus dencker
                                        Select case lto
                                        case "dencker"
                            
                                        case else
                                        'response.Write "medid: " & varTjDatoUS_man
                                        call normtimerPer(medid, varTjDatoUS_man, 6, 0)

                                        end select


                                        'response.Write "<br> norm:" & ntimMan
                             

                                    %>




                        


                                    <input type="hidden" id="timerdagman" value="<%=replace(manTimer, ",", ".") %>" />
                                    <input type="hidden" id="timerdagtir" value="<%=replace(tirTimer, ",", ".") %>" />
                                    <input type="hidden" id="timerdagons" value="<%=replace(onsTimer, ",", ".") %>" />
                                    <input type="hidden" id="timerdagtor" value="<%=replace(torTimer, ",", ".") %>" />
                                    <input type="hidden" id="timerdagfre" value="<%=replace(freTimer, ",", ".") %>" />
                                    <input type="hidden" id="timerdaglor" value="<%=replace(lorTimer, ",", ".") %>" />
                                    <input type="hidden" id="timerdagson" value="<%=replace(sonTimer, ",", ".") %>" />


                                    <input type="hidden" id="normdagman" value="<%=replace(ntimMan, ",", ".") %>" />
                                    <input type="hidden" id="normdagtir" value="<%=replace(ntimTir, ",", ".") %>" />
                                    <input type="hidden" id="normdagons" value="<%=replace(ntimOns, ",", ".") %>" />
                                    <input type="hidden" id="normdagtor" value="<%=replace(ntimTor, ",", ".") %>" />
                                    <input type="hidden" id="normdagfre" value="<%=replace(ntimFre, ",", ".") %>" />
                                    <input type="hidden" id="normdaglor" value="<%=replace(ntimLor, ",", ".") %>" />
                                    <input type="hidden" id="normdagson" value="<%=replace(ntimSon, ",", ".") %>" />


                                   <!--<div class="row">
                                       <div class="col-lg-5"><div id="stacked-vertical-chart" class="chart-holder-200"></div></div>
                                    </div>
                                    -->

                                    <%
                                        'maxHeight = "200px"
                                        'Width = "9%"
                                        dd = now
                                        dateMan = DateAdd("d",0,varTjDatoUS_man)
                                        dateTir = DateAdd("d",1,varTjDatoUS_man)
                                        dateOns = DateAdd("d",2,varTjDatoUS_man)
                                        dateTor = DateAdd("d",3,varTjDatoUS_man)
                                        dateFre = DateAdd("d",4,varTjDatoUS_man)
                                        dateLor = DateAdd("d",5,varTjDatoUS_man)
                                        dateSon = DateAdd("d",6,varTjDatoUS_man)

                                        'timerHeight_man = manTimer * 30
                                        'timerHeight_tir = tirTimer * 30
                                        'timerHeight_ons = onsTimer * 30
                                        'timerHeight_tor = torTimer * 30
                                        'timerHeight_fre = freTimer * 30
                                        'timerHeight_lor = lorTimer * 30
                                        'timerHeight_son = sonTimer * 30

                                        if cDate(dd) >= cDate(dateMan) OR manTimer <> 0 then
                                        balMan = manTimer - ntimMan
                                        else
                                        balMan = 0
                                        end if

                                        if cDate(dd) >= cDate(dateTir) OR tirTimer <> 0 then
                                        balTir = tirTimer - ntimTir
                                        else
                                        balTir = 0
                                        end if

                                        if cDate(dd) >= cDate(dateOns) OR onsTimer <> 0 then
                                        balOns = onsTimer - ntimOns
                                        else
                                        balOns = 0
                                        end if

                                        if cDate(dd) >= cDate(dateTor) OR torTimer <> 0 then
                                        balTor = torTimer - ntimTor
                                        else
                                        balTor = 0
                                        end if

                                        if cDate(dd) >= cDate(dateFre) OR freTimer <> 0 then
                                        balFre = freTimer - ntimFre
                                        else
                                        balFre = 0
                                        end if

                                        if cDate(dd) >= cDate(dateLor) OR lorTimer <> 0 then
                                        balLor = lorTimer - ntimLor
                                        else
                                        balLor = 0
                                        end if

                                        if cDate(dd) >= cDate(dateSon) OR sonTimer <> 0 then
                                        balSon = sonTimer - ntimSon
                                        else
                                        balSon = 0
                                        end if

                                        weekhourstotal = manTimer + tirTimer + onsTimer + torTimer + freTimer + lorTimer + sonTimer
                                        normtotal = ntimMan + ntimTir + ntimOns + ntimTor + ntimFre + ntimLor + ntimSon

                                        baltotal = (balMan + balTir + balOns + balTor + balFre + balLor + balSon) 'weekhourstotal - normtotal

                                        if baltotal < 0 then 
                                            balcolor = "red;"
                                        else
                                            balcolor = "green;"
                                        end if

                                        if balMan < 0 then 
                                            mancolor = "red;"
                                        else
                                            mancolor = "green;"
                                        end if

                                        if balTir < 0 then 
                                            tircolor = "red;"
                                        else
                                            tircolor = "green;"
                                        end if

                                        if balOns < 0 then 
                                            onscolor = "red;"
                                        else
                                            onscolor = "green;"
                                        end if

                                        if balTor < 0 then 
                                            torcolor = "red;"
                                        else
                                            torcolor = "green;"
                                        end if

                                        if balFre < 0 then 
                                            frecolor = "red;"
                                        else
                                            frecolor = "green;"
                                        end if

                                        'response.Write balTir 
                                        'response.Write timerHeight_man

                                        if browstype_client = "ip" then
                                            oskriftcolspan = 1

                                            visMan = 0
                                            visTir = 0
                                            visOns = 0
                                            visTor = 0
                                            visFre = 0
                                            visLor = 0
                                            visSon = 0

                                            if dayNumberChoosen = 1 then
                                                visMan = 1
                                            end if

                                            if dayNumberChoosen = 2 then
                                                visTir = 1
                                            end if

                                            if dayNumberChoosen = 3 then
                                                visOns = 1
                                            end if

                                            if dayNumberChoosen = 4 then
                                                visTor = 1
                                            end if

                                            if dayNumberChoosen = 5 then
                                                visFre = 1
                                            end if

                                            if dayNumberChoosen = 6 then
                                                visLor = 1
                                            end if

                                            if dayNumberChoosen = 7 then
                                                visSon = 1
                                            end if
                                                
                                        else
                                            oskriftcolspan = 2
                                            visMan = 1
                                            visTir = 1
                                            visOns = 1
                                            visTor = 1
                                            visFre = 1
                                            visLor = 1
                                            visSon = 1
                                        end if

                                        

                             
                                    %>


                                    <tr>
                                        <td colspan="<%=oskriftcolspan %>"><b><%=favorit_txt_019 %>:</b></td>

                                        <%if visMan = 1 then %>
                                        <td style="text-align:center"><%=formatnumber(manTimer, 2) %></td>
                                        <%end if %>

                                        <%if visTir = 1 then %>
                                        <td style="text-align:center"><%=formatnumber(tirTimer, 2) %></td>
                                        <%end if %>

                                        <%if visOns = 1 then %>
                                        <td style="text-align:center"><%=formatnumber(onsTimer, 2) %></td>
                                        <%end if %>

                                        <%if visTor = 1 then %>
                                        <td style="text-align:center"><%=formatnumber(torTimer, 2) %></td>
                                        <%end if %>

                                        <%if visFre = 1 then %>
                                        <td style="text-align:center"><%=formatnumber(freTimer, 2) %></td>
                                        <%end if %>

                                        <%if visLor = 1 then %>
                                        <td style="text-align:center"><%=formatnumber(lorTimer, 2) %></td>
                                        <%end if %>

                                        <%if visSon = 1 then %>
                                        <td style="text-align:center"><%=formatnumber(sonTimer, 2) %></td>
                                        <%end if %>

                                        <%if browstype_client <> "ip" then %>
                                        <td colspan="2" style="text-align:center"><%=formatnumber(weekhourstotal, 2) %></td>
                                        <%end if %>
                                    </tr>

                                    <tr>
                                        <td colspan="<%=oskriftcolspan %>"><b><%=favorit_txt_020 %>:</b></td>

                                        <%if visMan = 1 then %>
                                        <td style="text-align:center"><%=formatnumber(ntimMan, 2) %></td>
                                        <%end if %>

                                        <%if visTir = 1 then %>
                                        <td style="text-align:center"><%=formatnumber(ntimTir, 2) %></td>
                                        <%end if %>

                                        <%if visOns = 1 then %>
                                        <td style="text-align:center"><%=formatnumber(ntimOns, 2) %></td>
                                        <%end if %>

                                        <%if visTor = 1 then %>
                                        <td style="text-align:center"><%=formatnumber(ntimTor, 2) %></td>
                                        <%end if %>

                                        <%if visFre = 1 then %>
                                        <td style="text-align:center"><%=formatnumber(ntimFre, 2) %></td>
                                        <%end if %>

                                        <%if visLor = 1 then %>
                                        <td style="text-align:center"><%=formatnumber(ntimLor, 2) %></td>
                                        <%end if %>

                                        <%if visSon = 1 then %>
                                        <td style="text-align:center"><%=formatnumber(ntimSon, 2) %></td>
                                        <%end if %>

                                        <%if browstype_client <> "ip" then %>
                                        <td colspan="2" style="text-align:center"><%=formatnumber(normtotal, 2) %></td>
                                        <%end if %>
                                    </tr>

                                    <tr>
                                        <td colspan="<%=oskriftcolspan %>"><b><%=favorit_txt_021 %>:</b></td>

                                        <%if visMan = 1 then %>
                                        <td style="text-align:center; color:<%=mancolor%>"><%=formatnumber(balMan, 2) %></td>
                                        <%end if %>

                                        <%if visTir = 1 then %>
                                        <td style="text-align:center; color:<%=tircolor%>"><%=formatnumber(balTir, 2) %></td>
                                        <%end if %>

                                        <%if visOns = 1 then %>
                                        <td style="text-align:center; color:<%=onscolor%>"><%=formatnumber(balOns, 2) %></td>
                                        <%end if %>

                                        <%if visTor = 1 then %>
                                        <td style="text-align:center; color:<%=torcolor%>"><%=formatnumber(balTor, 2) %></td>
                                        <%end if %>

                                        <%if visFre = 1 then %>
                                        <td style="text-align:center; color:<%=frecolor%>"><%=formatnumber(balFre, 2) %></td>
                                        <%end if %>

                                        <%if visLor = 1 then %>
                                        <td style="text-align:center;"><%=formatnumber(balLor, 2) %></td>
                                        <%end if %>

                                        <%if visSon = 1 then %>
                                        <td style="text-align:center;"><%=formatnumber(balSon, 2) %></td>
                                        <%end if %>

                                        <%if browstype_client <> "ip" then %>
                                        <td colspan="2" style="text-align:center; color:<%=balcolor%>"><%=formatnumber(baltotal, 2) %></td>
                                        <%end if %>
                                    </tr>

                                    <%if lto = "wap" OR lto = "intranet - local" then
                                        aty_sql_modregn = "tfaktim = 114" '31
                                        call timerDenneUge(medid, lto, varTjDatoUS_man, aty_sql_modregn, timerdenneuge_dothis, SmiWeekOrMonth)
                                        weekhourstotal = manTimer + tirTimer + onsTimer + torTimer + freTimer + lorTimer + sonTimer
                                    %>

                                    <tr>
                                        <td colspan="<%=oskriftcolspan %>">Korrektion Flex:</td>

                                        <%if visMan = 1 then %>
                                        <td style="text-align:center">(<%=formatnumber(manTimer, 2) %>)</td>
                                        <%end if %>

                                        <%if visTir = 1 then %>
                                        <td style="text-align:center">(<%=formatnumber(tirTimer, 2) %>)</td>
                                        <%end if %>

                                        <%if visOns = 1 then %>
                                        <td style="text-align:center">(<%=formatnumber(onsTimer, 2) %>)</td>
                                        <%end if %>

                                        <%if visTor = 1 then %>
                                        <td style="text-align:center">(<%=formatnumber(torTimer, 2) %>)</td>
                                        <%end if %>

                                        <%if visFre = 1 then %>
                                        <td style="text-align:center">(<%=formatnumber(freTimer, 2) %>)</td>
                                        <%end if %>

                                        <%if visLor = 1 then %>
                                        <td style="text-align:center">(<%=formatnumber(lorTimer, 2) %>)</td>
                                        <%end if %>

                                        <%if visSon = 1 then %>
                                        <td style="text-align:center">(<%=formatnumber(sonTimer, 2) %>)</td>
                                        <%end if %>

                                        <%if browstype_client <> "ip" then %>
                                        <td colspan="2" style="text-align:center">(<%=formatnumber(weekhourstotal, 2) %>)</td>
                                        <%end if %>
                                    </tr>

                                    <%end if %>


                                    <%if lto = "wap" OR lto = "intranet - local" AND browstype_client <> "ip" then 
                                            
                                           call fn_flexSaldoFYreal_norm(medid)
                                            
                                           %>

                                         <tr>
                                            <td>Flexsaldo: (opgjort pr. <b><%=formatdatetime(slDatoNormPrDdMinus1, 2) %></b> t.o.m ig�r)</td>
                                            <td colspan="7">&nbsp;</td>
                                         
                                            <td style="text-align:center;"><b><%=formatnumber(flexSaldoFYreal_norm, 2) %></b></td>
                                        </tr>
                                            
                                            
                                    <%end if%>

                                    <%end if %>

                            </tbody>

                        </table>

                            <%
                            if browstype_client <> "ip" then
                                btncls = "btn-sm"
                            else
                                btncls = ""
                            end if
                            %>

                            <div class="row">
                            <div class="col-lg-10">&nbsp</div>
                            <div class="col-lg-2 pad-b10">
                                <button type="submit" class="btn btn-success <%=btncls %> pull-right"><b><%=favorit_txt_018 %></b></button>
                            </div>
                        </div>
                              <!--<input type="hidden" value="<%=ign_projgrp%>" id="ign_projgrp" />-->     
                            
                        </form>

                        

                
                           
                            <%'** NOT IN USE 20190329 ****
                            fm_tilfoj = 0
                            if cint(fm_tilfoj) = 1 then
                                strSQL = "SELECT id, medarb, aktid, favorit FROM timereg_usejob WHERE medarb ="& medid & " AND aktid <> 0"
                            %>
                               <!-- <select name="FM_favorittilfoj" class="form-control input-small"  onchange="submit();"> -->
                            <%
                              

                                  oRec.open strSQL, oConn, 3

                                  While not oRec.EOF

                                  aktid = oRec("aktid")
                                %>
                                   <!-- <a href="favorit.asp?func=tilfojfavorit&id=<%=aktid %>"><%=aktid %></a> -->
                                <%
                                  oRec.movenext
                                  wend                     
                                  oRec.close
                              %>                               
                           <!-- </select> -->
                            <%end if %>
                        
                        

                        <!-- Godkend / Afslut uge -->
                          <%if browstype_client <> "ip" then
                                showElements = "normal"
                              else
                              showElements = "none"
                         end if%>
                          <div class="row" style="display:<%=showElements%>;">
                            <div class="col-lg-10">
                                 
                                 <%
                                 
                                 dothis = 1
                                 SmiWeekOrMonth = 1

                                

                                  'usemrn
                                  'varTjDatoUS_man
                                  'smilaktiv
                                  'media
                                  'meType
                                  'afslutugekri
                                  timerdenneuge_dothis = 0
                                  aty_sql_typer = aty_sql_realhours

                                 call showafslutuge_ugeseddel %>

                  

                            </div></div>
                    

                    </div>
                </div>
            </div>

               



    <%end select %>


    <%if (browstype_client <> "ip") then %>
        </div>
    </div>



    <%end if %>
 <!--
  <div id="fajl" style="width:400px; height:400px; border:2px #999999 solid; left:200px;">..</div>
     -->
 

                <!-- 
                '********************************************************
                Kontakt K�rsel info div 
                
                
                
                '********************************************************
                -->
	            <div id="korseldiv" style="position:absolute; left:125px; top:300px; background-color:#ffffff; width:600px; border:10px #CCCCCC solid; padding:10px; visibility:hidden; display:none; z-index:20000; height:400px; overflow:auto;">
				 
                    
                  
                    <!--
                    <a href="javascript:popUp('kontaktpers.asp?func=opr&id=0&kid=<%=kid%>&rdir=treg','550','500','20','20')" target="_self" class=remnu>+ <%=timereg_txt_094 %></a>
                    -->

				<table border=0 cellspacing=0 cellpadding=0 width=100%>
				<form action="#" method="post">
				<tr><td align=right><span id="jq_hideKpersdiv" style="color:#999999; font-size:14px;">X</span></td></tr>
				
				<tr>
				    <td><h3><%=tsa_txt_360 %>:</h3>

                    <div class="row">
                       <div class="col-lg-12">
                    <%=tsa_txt_365 %><br /><%=tsa_txt_361 %>
                             </div>
                        </div>

                        <div class="row">
                       <div class="col-lg-6">
                        <select class="form-control input-small" id="ko0" style="width:500px;"><option value="0">....</option></select><br /><br />
                        </div>
                        </div>

                        <div class="row">
                       <div class="col-lg-6">
                        <h4><%=timereg_txt_095 %>:</h4>
                        <select class="form-control input-small" id="ko1" style="width:500px;"><option value="0">....</option></select>
                           </div>
                        </div>



                        
     



                        
				   </td>
				</tr>
				
				
			
			
				<tr><td>
				<%call erkmDialog()  'skal den vises
				%>

                    
                    
                    <input id="FM_sog_kpers_dist_kid" value="0" type="hidden" />
				    <input id="koKmDialog" value="<%=kmDialogOnOff%>" type="hidden" />
                    <input id="koFlt" value="0" type="text" /><!-- row -->
                    <input id="koFltx" value="1" type="hidden" /><!-- x dag -->
                    <!--<input id="kperfil_fundet" type="text" value="0">-->

                    <!--<input id="indlaes_koadr_2" type="button" value=" Indl�s >> " /><br />&nbsp;-->

                      <div class="row">
                           
                            <div class="col-lg-12">
                                <br />
                                <button type="submit" id="indlaes_koadr_2" class="btn btn-success btn-sm pull-right"><b><%=favorit_txt_018 %></b></button>
                            </div>
                        </div>

                    </td></tr>
				</form>
				</table>
				</div>
				<!--KM dialog div -->
				


<!--#include file="../inc/regular/footer_inc.asp"-->