<!--#include file="../inc/connection/conn_db_inc.asp"-->

     
<%
 '**** S�gekriterier AJAX **'
        'section for ajax calls

        if Request.Form("AjaxUpdateField") = "true" then
        Select Case Request.Form("control")


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
            
                 lastKid = 0
                
                  'strJobogKunderTxt = strJobogKunderTxt &"<span style=""color:red; font-size:9px; float:right;"" class=""luk_jobsog"">[X]</span>"    
                         
                if cint(ign_projgrp) = 0 then' alm

                    strSQL = "SELECT j.id AS jid, j.jobnavn, j.jobnr, j.jobstatus, k.kkundenavn, k.kkundenr, k.kid FROM timereg_usejob AS tu "_ 
                    &" LEFT JOIN job AS j ON (j.id = tu.jobid) "_
                    &" LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "_
                    &" WHERE tu.medarb = "& medid &" AND (j.jobstatus = 1 OR j.jobstatus = 3) AND "_
                    &" (jobnr LIKE '"& jobkundesog &"%' OR jobnavn LIKE '%"& jobkundesog &"%' OR "_
                    &" kkundenavn LIKE '"& jobkundesog &"%' OR kkundenr = '"& jobkundesog &"' OR k.kinit = '"& jobkundesog &"')  AND kkundenavn <> ''"_
                    &" GROUP BY j.id ORDER BY kkundenavn, jobnavn LIMIT 50"       
    
                else 'ignorer projektgrupper

                
                    strSQL = "SELECT j.id AS jid, j.jobnavn, j.jobnr, j.jobstatus, k.kkundenavn, k.kkundenr, k.kid FROM job AS j "_
                    &" LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "_
                    &" WHERE (j.jobstatus = 1 OR j.jobstatus = 3) AND "_
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
                strSQLtilbud = "SELECT jobstatus FROm job WHERE id = "& jobid
                oRec5.open strSQLtilbud, oConn, 3
                if not oRec5.EOF then

                jobstatusTjk = oRec5("jobstatus")

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


               if pa = "1" then
               strSQL = "SELECT a.id AS aid, navn AS aktnavn, a.fase FROM timereg_usejob tu LEFT JOIN aktiviteter AS a ON (a.job = tu.jobid "& onlySalesact &") "_
               &" WHERE tu.medarb = "& medid &" AND tu.jobid = "& FM_jobid &" AND aktstatus = 1 AND ("& aty_sql_hide_on_treg &") "& onlySalesact &" ORDER BY a.fase, sortorder, navn"   

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
                'strAktTxt = strAktTxt & "<option value=0>HER: "& strSQL &" findesaktfav: "& findesaktfav &"</option>"
                'response.write strAktTxt
                'response.end
                'end if

                strAktTxt = strAktTxt & "<option SELECTED disabled value=0>..</option>"             

                lastFase = ""
                thisFase = ""
                oRec.open strSQL, oConn, 3
                while not oRec.EOF

                findesaktfavshowakt = 0
                if instr(findesaktfav, "#"& oRec("aid") &"#") then
                findesaktfavshowakt = 1
                end if
        
                
                showAkt = 0
                if pa = "1" then
            
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
                    strAktTxt = strAktTxt & "<option value="& oRec("aid") &" "& optionFcDis &">"& oRec("aktnavn") &" "& fcsaldo_txt &"</option>"             

               
                 end if 'showakt

                if isnull(oRec("fase")) = false then
                lastFase = oRec("fase")
                else
                lastFase = ""
                end if
                
                oRec.movenext
                wend
                oRec.close

              


                    '*** ��� **'
                    call jq_format(strAktTxt)
                    strAktTxt = jq_formatTxt


                    response.write strAktTxt

               
            end if 'filerval




        case "tilfoj_akt"

        jobid = request("FM_jobid")
        aktid = request("FM_aktid")
        medid = request("FM_medid_id")

       
       
       
        
        'response.write "jobid "& jobid  & "medid: "& medid & " aktid: "& aktid
         'response.end

        strSQL = "SELECT id FROM timereg_usejob WHERE jobid = "& jobid &" AND aktid = "& aktid &" AND medarb = "& medid
        oRec.open strSQL, oConn, 3
        if not oRec.EOF then
        Strtilfojakt = "UPDATE timereg_usejob SET favorit = 1 WHERE aktid = "& aktid & " AND medarb = "& medid &" AND jobid = "& jobid
        else
        Strtilfojakt = "INSERT INTO timereg_usejob SET jobid = "& jobid &", favorit = 1, aktid = "& aktid & ", medarb = "& medid
        end if
        oRec.close 

        'response.write "HER: "&  Strtilfojakt
        'response.end

        oConn.execute(Strtilfojakt)



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
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<SCRIPT src="../timereg/inc/smiley_jav.js"></script>
<script src="js/favorit_jav_20171128.js"></script>
<!-- <link rel="stylesheet" href="../../bower_components/bootstrap-datepicker/css/datepicker3.css"> -->




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


<div class="wrapper">
<div class="content">

    <%
    if len(session("user")) = 0 then
	%>
	
	<%
	errortype = 5
	call showError(errortype)
       response.End
	end if

    thisfile = "favorit.asp"

    if len(trim(request("FM_medid"))) <> 0 then
    medid = request("FM_medid")
    else
    medid = session("mid")
    end if


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

    call menu_2014

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

    varTjDatoUS_man = request("varTjDatoUS_man")
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

    if len(trim(request("FM_medid"))) <> 0 then
    medid = request("FM_medid")
    else
    medid = session("mid") 
    end if

    'Response.write "<br><br><br><br><br><br><br><div style='left:200px; top:100px; position:relative;'>MEDID: " & request("FM_medid") & "</div>"
    call ersmileyaktiv()
    call smileyAfslutSettings()

    if cint(SmiWeekOrMonth) = 0 then
    usePeriod = datePart("ww", varTjDatoUS_man, 2,2)
    useYear = year(varTjDatoUS_man)
    else
    usePeriod = datePart("m", varTjDatoUS_man, 2,2)
    useYear = year(varTjDatoUS_man)
    end if

    call erugeAfslutte(useYear, usePeriod, medid, SmiWeekOrMonth, 0)

%>



            <div class="container">
                <div class="portlet">
                    <h3 class="portlet-title"><u><%=favorit_txt_001 %></u> </h3>
                    <div class="portlet-body">                                                                                   
                        <form action="favorit.asp?" method="post" name="favorit" id="favorit">
                                              
                        <!-- <input type="hidden" name="varTjDatoUS_man" id="varTjDatoUS_man" value="<%=varTjDatoUS_man %>"> -->

                      <!--  <div class="row">
                            <div class="col-lg-2">
                                 
                                <%
                                    strSQL = "SELECT Mid, Mnavn, Mnr, Brugergruppe, init FROM medarbejdere WHERE mansat <> 2 GROUP BY mid ORDER BY Mnavn" 
                                %>
                                <select name="FM_medid" id="FM_medid" <%=progrpmedarbDisabled  %> class="form-control input-small"  onchange="submit();" style="width:210px">
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
                                     <option value="<%=oRec("Mid")%>" <%=isSelected%>><%=StrMnavn &" "& StrMinit%></option>
                                    <%
                                    oRec.movenext
                                    wend
                                    oRec.close  
                                    %>
                                </select>
                            </div>
                           
                           <div class="col-lg-1" style="z-index:1;">
                               <div class='input-group date' style="padding-left:30px; width:50px">
                                    <input type="text" class="form-control input-small" name="varTjDatoUS_selectedday" id="varTjDatoUS_selectedday" value="<%=varTjDatoUS_selectedday %>" />
                                    <span class="input-group-addon input-small">
                                    <span class="fa fa-calendar">
                                    </span>
                                    </span>
                                </div>
                            </div>

                            <div class="col-lg-2">
                                <button type="submit" class="btn btn-sm btn-default"><b>G�</b></button>
                            </div>

                            <div class="col-lg-3"></div>
                            <h4 class="col-lg-2" style="text-align:right"><a href="favorit.asp?FM_medid=<%=medid %>&varTjDatoUS_man=<%=prev_varTjDatoUS_man %>" ><</a>&nbsp Uge <%=datepart("ww",weeknumber)  %> &nbsp<a href="favorit.asp?FM_medid=<%=medid %>&varTjDatoUS_man=<%=next_varTjDatoUS_man %>" >></a></h4>
                            
                        </div> -->

                            <table>
                                <tr>
                                    <td>
                                        <%'SKAL UDBYGGES TIL ALLE man er projektelder for
                                        if cint(level) = 1 then
                                            sqgMw = " mansat <> 2"
                                        else
                                            sqgMw = " mid = "& session("mid") &" AND mansat <> 2"
                                        end if

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
                                    
                                    <td>                                       
                                        <button type="submit" class="btn btn-sm btn-default"><b><%=favorit_txt_004 %></b></button>                                       
                                    </td>

                                   <td style="text-align:right; width:100%"><h4><a href="favorit.asp?FM_medid=<%=medid %>&varTjDatoUS_man=<%=prev_varTjDatoUS_man %>" ><</a>&nbsp <%=favorit_txt_005 & " " %> <%=datepart("ww",weeknumber)  %> &nbsp<a href="favorit.asp?FM_medid=<%=medid %>&varTjDatoUS_man=<%=next_varTjDatoUS_man %>" >></a></h4></td>
                                </tr>
                            </table>

                        </form>
                        <%'response.Write "id: " & medid%>


                        <form action="../timereg/timereg_akt_2006.asp?func=db&rdir=favorit&varTjDatoUS_man=<%=varTjDatoUS_man%>" method="post">
                            
                            <%usemrn = medid %>
                            <input type="hidden" name="FM_medid" value="<%=medid %>" />
                            <input id="usemrn" value="<%=usemrn %>" type="hidden" />
                            <input type="hidden" id="Hidden4" name="FM_dager" value="7"/>
                           
                            <input type="hidden" id="" name="FM_vistimereltid" value="0"/>
                            <input type="hidden" id="Hidden5" name="year" value="<%=year(now) %>"/>
                            <input type="hidden" value="0" name="extsysId" />

                        <script src="js/fav_mat.js" type="text/javascript"></script>
                        <table class="table table-striped dataTable table-bordered ui-datatable">
                            
                            

                            <thead>
                                <tr>
                                    <th><%=favorit_txt_002 %></th>
                                    <th><%=favorit_txt_003 %></th>

                                    <%
                                            perInterval = 6 'dateDiff("d", varTjDatoUS_man, varTjDatoUS_son, 2,2) 
                                            perIntervalLoop = perInterval

                                            for l = 0 to perIntervalLoop 
        
                                            if l = 0 then
                                            varTjDatoUS_use = varTjDatoUS_man
                                            else
                                            varTjDatoUS_use = dateAdd("d", l, varTjDatoUS_man)
                                            end if 

                                            showdate = Right("0" & DatePart("d",varTjDatoUS_use,2,2), 2) & "-" & Right("0" & DatePart("m",varTjDatoUS_use,2,2), 2) & "-" & DatePart("yyyy",varTjDatoUS_use,2,2)

                                            'showweekdayname = weekdayname(weekday(varTjDatoUS_use, 1))
                                            'daynamenum = weekday(varTjDatoUS_use,1)

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
                                            case 6
                                            daynameword = favorit_txt_012
                                            end select

                                            'daynameword = WeekDayName(daynamenum,true)

                                            'response.Write weekdayname(weekday(varTjDatoUS_use, 1))

                                            %>
                                                <th style="width:75px"><%=UCase(Left(daynameword,1)) & Mid(daynameword,2) %> <br />
                                                    <%=showdate %>
                                                </th>
                                            <%

                                            next
                                    %>
                                    <th style="text-align:center; width:55px;"><%=favorit_txt_013 %></th>
                                    <th></th>
                                </tr>
                            </thead>

                            <tbody>
                                <%
                                     
                                     favoriter = 0
                                     lastaktid = 0
                                     i = 150

                                     'Dim jobid, aktid, medarb
                                     Redim jobid(i), aktid(i), medarb(i)

                                     select case lto
                                     case "wap"
                                      strAkrORderBy = "a.sortorder, a.navn"

                                            if session("mid") = 1 or session("mid") = 21 then 'LK skal kunne se korrektion p� alle medarbejdere
                                            strSQlaktspecial = " OR tu.id = 200"
                                            else
                                            strSQlaktspecial = ""
                                            end if

                                     case else
                                      strAkrORderBy = "a.navn" 
                                      strSQlaktspecial = ""
                                     end select

                                     

                                     StrSqlfav = "SELECT medarb, jobid, aktid, forvalgt_af FROM timereg_usejob AS tu "_
                                     &" LEFT JOIN job j ON (j.id = jobid) "_
                                     &" LEFT JOIN aktiviteter a ON (a.id = aktid) "_
                                     &" LEFT JOIN kunder k ON (kid = jobknr) "_
                                     &" WHERE (medarb = "& medid & " AND favorit <> 0) "& strSQlaktspecial &" AND a.aktstatus = 1 GROUP BY a.id ORDER BY kkundenavn, jobnavn, jobnr, "& strAkrORderBy &"" 
                                     
                                        'if session("mid") = 1 or session("mid") = 9 then
                                        'Response.write StrSqlfav
                                        'end if

                                     oRec.open StrSqlfav, oConn, 3

                                      i = 0

                                     while not oRec.EOF

                                        '*** Dobbelttjekekr om det er et tilbud - kun salgs aktiviteter vises
                                        jobstatusTjk = 1
                                        strSQLtilbud = "SELECT jobstatus FROm job WHERE id = "& oRec("jobid")
                                        oRec5.open strSQLtilbud, oConn, 3
                                        if not oRec5.EOF then

                                        jobstatusTjk = oRec5("jobstatus")

                                        end if
                                        oRec5.close

                                        afakturerbar = 0
                                        if cint(jobstatusTjk) = 3 then 'tilbud
                                        
                                        strSQLtilbudA = "SELECT fakturerbar FROM aktiviteter WHERE id = "& oRec("aktid")
                                        oRec5.open strSQLtilbudA, oConn, 3
                                        if not oRec5.EOF then

                                        afakturerbar = oRec5("fakturerbar")

                                        end if
                                        oRec5.close
                                        end if

                                                '** dobbeltjek KUN aktivejob + tilbud && salgsaktivitet
                                                if cint(jobstatusTjk) = 1 OR (cint(jobstatusTjk) = 3 AND cint(afakturerbar) = 6) then  

                                                 jobid(i) = oRec("jobid")
                                                 aktid(i) = oRec("aktid")
                                                 medarb(i) = oRec("medarb")
                                                 'response.Write jobid(i)

                                                 if lastaktid <> oRec("aktid") then
                                                 i = i + 1
                                                 end if

                                                 lastaktid = oRec("aktid")
                                     
                                                 favoriter = favoriter + 1


                                                 end if
                                     

                                     oRec.movenext
                                     wend
                                     oRec.close

                                    i_end = i
                                    i = 0
                                    'response.write "<br> d" & i_end
                                    'response.Write "lastid: " & lastjobid 
                                    
                                    y = 0                                
                                     for i = 0 to i_end -1
                                                          
                                        'response.Write "aktid1: " & aktid(i)

                                        StrSQLjob = "SELECT id, jobnavn, jobnr, jobstartdato, jobslutdato, jobans1, jobans2, jobans3, jobans4, jobans5, jobknr, beskrivelse, budgettimer FROM job WHERE id ="& jobid(i)
                                        oRec3.open StrSQLjob, oConn, 3
                                        if not oRec3.EOF then
                                        jobids = oRec3("id")
                                        jobnavn = oRec3("jobnavn") & " ("& oRec3("jobnr") &")"

                                        strSQLkunde = "SELECT kundeans1, kkundenavn FROM kunder WHERE kid = "& oRec3("jobknr")
                                        oRec4.open strSQLkunde, oConn, 3
                                        if not oRec4.EOF then
                                        kundeans = oRec4("kundeans1")
                                        kundenavn = oRec4("kkundenavn")
                                        end if
                                        oRec4.close

                                        
                                        StrSQLakt = "SELECT id, navn, beskrivelse, budgettimer, fakturerbar, fase FROM aktiviteter WHERE id ="& aktid(i)
                                        oRec2.open StrSqlakt, oConn, 3
                                        if not oRec2.EOF then
                                        aktNavn = oRec2("navn")
                                        TaktId = oRec2("id")
                                        aktbudgettimer = oRec2("budgettimer")
                                        fakturerbar = oRec2("fakturerbar")
                                        thisfase = oRec2("fase")
                                        'response.Write jobnavn

                                     select case lto
                                     case "wap"
                                     
                                        if aktNavn = "L�ge/tandl�ge" then
                                        %>
                                        <tr><td colspan="11"><br />&nbsp;</td></tr>
                                        <%
                                        end if

                                     end select

                                        %>
                                        <tr>
                                            <td style="vertical-align:top; width:200px"><input type="hidden" value="<%=jobids %>" name="FM_jobid" />
                                                <span style="font-size:9px;"><%=kundenavn %></span><br /><%=jobnavn %><a data-toggle="modal" href="#styledModalSstGrp20"><span id="jobinfo_<%=jobids %>" style="color:#8c8c8c" class="fa fa-file-text pull-right jobinfo"></span></a>                                                                                              
                                                 <div id="jobmodal_<%=jobids %>" class="modal">
                                                     <div class="modal-content" style="width:400px; height:500px;">
                                                         <%
                                                            

                                                            strSQLjobnas = "SELECT mnavn FROM medarbejdere WHERE mid="& oRec3("jobans1")  
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

                                                            strSQLrealiseret = "SELECT sum(timer) as timer FROM timer WHERE tjobnr ="& jobids
                                                            oRec4.open strSQLrealiseret, oConn, 3
                                                            if not oRec4.EOF then
                                                             realiserettimer = oRec4("timer")
                                                            end if 
                                                            oRec4.close

                                                         %>
                                                       <!--  <div class="row">
                                                            <div class="col-lg-4">
                                                                <table>
                                                                     <tr>
                                                                         <td><b>Start & slut dato</b>
                                                                             <br />
                                                                             <%=oRec3("jobstartdato") & " - " & oRec3("jobslutdato")  %>
                                                                            <br /><br />
                                                                         </td>
                                                                     </tr>
                                                                    <tr>
                                                                         <td><b>Jobansvarlig:</b> <%=jobansvarlig %><br />
                                                                             <b>Kundeansvarlig:</b> <%=kundeansvarlig  %>                                                                             
                                                                         </td>
                                                                       
                                                                     </tr>                                                                                                                                    
                                                                </table>
                                                            </div>
                                                             <div class="col-lg-4">
                                                                <table>
                                                                    <tr>
                                                                        <td><b>Forkalk.</b>: <%=oRec3("budgettimer") %> t.<br />
                                                                            <b>Realiseret:</b> <%=realiserettimer %> t.
                                                                        </td>
                                                                    </tr>
                                                                </table>
                                                            </div>
                                                            <div class="col-lg-4">
                                                                <table>
                                                                    <tr>
                                                                        <td><b>Jobbeskrivelse:</b>
                                                                            <br />
                                                                            <textarea rows="5" class="form-control input-small"><%=oRec3("beskrivelse") %></textarea>                                                                            
                                                                        </td>
                                                                    </tr>
                                                                </table>
                                                            </div>
                                                        </div> -->

                                                        <div class="row">
                                                            <div class="col-lg-12">
                                                                <b><%=kundenavn %></b> - <%=jobnavn %>
                                                            </div>
                                                        </div>
                                                         <br /><br />

                                                        <div class="row">
                                                             <div class="col-lg-4"><b><%=favorit_txt_022 %>:</b></div>
                                                             <div class="col-lg-5"><%=oRec3("jobstartdato")%></div>
                                                        </div>
                                                        <div class="row">
                                                             <div class="col-lg-4"><b><%=favorit_txt_023 %>:</b></div>
                                                             <div class="col-lg-5"><%=oRec3("jobslutdato")  %></div>
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
                                                            <div class="col-lg-5"><%=oRec3("budgettimer") %> t.</div>
                                                        </div>
                                                        <div class="row">
                                                            <div class="col-lg-4"><b><%=favorit_txt_027 %>:</b></div>
                                                            <div class="col-lg-5"><%=realiserettimer %> t.</div>
                                                        </div>
                                                        <br />

                                                        <div class="row">
                                                            <div class="col-lg-4"><b><%=favorit_txt_028 %>:</b></div>
                                                        </div>
                                                        <div class="row">
                                                            <div class="col-lg-12"><div class="form-control input-small" style="height:100px; overflow-y:auto;"><%=oRec3("beskrivelse") %></div></div>
                                                        </div>                                                    

                                                     </div>
                                                 </div>

                                            </td>

                                            <td style="vertical-align:middle; width:195px;">
                                                <input type="hidden" value="<%=TaktId %>" name="FM_aktivitetid" />
                                                <%if len(trim(thisfase)) <> 0 then %>
                                                <span style="font-size:9px;"><%=replace(thisfase, "_", "") %></span><br />
                                                <%end if %>

                                                <%=aktNavn %>

                                                <span id="modal_<%=TaktId %>" class="fa fa-chevron-down pull-right picmodal" style="color:#8c8c8c"></span>                                               
                                                <div id="myModal_<%=TaktId %>" style="display:none">
                                                
                                                    <%
                                                        StrSqltimerialt = "SELECT sum(Timer) as timer FROM timer WHERE TAktivitetId = "& TaktId
                                                        oRec6.open StrSqltimerialt, oConn, 3
                                                        if not oRec6.EOF then
                                                        timerforalle = oRec6("timer")
                                                        end if
                                                        oRec6.close

                                                        useDateStSQL = year(varTjDatoUS_man) &"/"& month(varTjDatoUS_man) &"/"& day(varTjDatoUS_man)
                                                        call ressourcetimerTildelt(useDateStSQL, jobids, TaktId, medid)



                                                        StrSqltotaltimer = "SELECT TAktivitetId, sum(timer) as timer FROM timer WHERE TAktivitetId = "& TaktId &" AND tmnr ="& medid
                                                        oRec5.open StrSqltotaltimer, oConn, 3
                                                        if not oRec5.EOF then
                                                
                                                        timertotal = oRec5("timer")
                                                       
                                                        end if
                                                        oRec5.close 
                                                        

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

                                                    %>


                                                       
                                                        <span style="font-size:75%; color:#5582d2"><%=favorit_txt_029 %>: <%=formatnumber(aktbudgettimer, 2) %></span>
                                                        <span style="font-size:75%; color:#5582d2">Forecast: <%=formatnumber(resTimerThisM, 2) %></span><br />
                                                        <span style="font-size:75%; color:<%=txtcolor%>;"><%=favorit_txt_030 %>: <%=formatnumber(timerforalle, 2) %></span>
                                                        &nbsp;<span style="font-size:75%; color:#5582d2;">(<%=favorit_txt_031 %>: <%=formatnumber(timertotal, 2) %>)</span>
                                                   
                                                </div>
                                            </td>
                                            
                                            <%
                                                
                                                for l = 0 to 6
                                                
                                                     
                                                    timerKom = ""
                                                    fmborcl = ""
                                                    godkendtstatus = 0
                                                    timerdag = ""
                                                    overfort = 0
                                                    origin = 0
                                                    y = y + 1
                                                    'response.Write y

                                                    if l = 0 then
                                                        timerdato = varTjDatoUS_man
                                                    else
                                                        timerdato = dateAdd("d",1,timerdato)
                                                    end if


                                                     varTjDato_ugedag = day(timerdato) & "-" & month(timerdato) & "-" & year(timerdato)

                                                     '**** Er periode lukket via l�nk�rsel **''
                                                     call lonKorsel_lukketPer(varTjDato_ugedag, job_internt) 'Hr -2 job bliver kun lukket ifh.- l�nperiode

                                                  
                                                    '*** tjekker om uge er afsluttet / lukket / l�nk�rsel
                                                    call tjkClosedPeriodCriteria(varTjDato_ugedag, ugeNrAfsluttet, usePeriod, SmiWeekOrMonth, splithr, smilaktiv, autogk, autolukvdato, lonKorsel_lukketIO)
                       

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

                                                    %>
                                                        <td>  
                                                               <div class="row">   

                                                            <!-- Til k�rsel -->
                                                            <input type="hidden" id="jq_kid_<%=y %>" value="<%=oRec3("jobknr") %>" />
                                                            <input type="hidden" id="jq_row_<%=y %>" value="<%=y %>" />
                                                            <input type="hidden" id="jq_flt_<%=y %>" value="<%=l %>" />
                                                            <input type="hidden" id="jq_tfa_<%=y %>" value="<%=fakturerbar %>" />
                                                            
                                                            
                                                            


                                                            <input type="hidden" name="FM_feltnr" value="<%=y %>" />
                                                            <input type="hidden" value="<%=timerdato %>" name="FM_datoer" />
                                                            <input type="hidden" value="dist" name="FM_destination_<%=y %>" />
                                                              <input type="hidden" name="FM_sttid" value="00:00"/>
                                                            <input type="hidden" name="FM_sltid" value="00:00"/>

                                                            <%if (origin <> 0 AND (lto <> "wap" AND cdbl(TaktId) <> 34)) OR cint(overfort) = 1 OR (((cint(ugeerAfsl_og_autogk_smil) = 1 AND cint(godkendtstatus) <> 2) OR (cint(godkendtstatus) = 1 Or cint(godkendtstatus) = 3)) AND level <> 1) then %>
                                                            <input type="hidden" name="FM_timer" value=""/>
                                                                <div class="col-lg-9" style="padding-right:2px!important"><input type="text" class="form-control input-small" style="width:55px; border:<%=fmborcl%>;" value="<%=timerdag %>" readonly /></div>
                                                                   
                                                                <%if cint(ugeerAfsl_og_autogk_smil) = 0 AND origin <> 0 then 'Er tastest ind via andre medier og UGE ikke godkendt %>
                                                                <span style="font-size:75%"><a style="color:dimgrey;" href="ugeseddel_2011.asp?usemrn=<%=medid %>&varTjDatoUS_man=<%=varTjDatoUS_man %>"><%=favorit_txt_014 %></a></span>
                                                                <%end if %>

                                                            <%else 
                                                                
                                                                %>
                                                            <div class="col-lg-9" style="padding-right:2px!important"><input type="text" class="form-control input-small timer_flt" id="FM_timer_<%=y%>" name="FM_timer" value="<%=timerdag %>" style="width:55px; border:<%=fmborcl%>;" /></div>
                                                            <div class="col-lg-2" style="padding-left:3px!important;"><span id="modal_<%=y%>" class="kommodal">+</span></div>
                                                          

                                                            <%end if %>
                                                            </div>
                                                            
                                                                <div id="kommentarmodal_<%=y%>" class="modal">
                                                                        <div class="modal-content">                                                                                                                        
                                                                            <div class="row">
                                                                                <div class="col-lg-2"><b><%=favorit_txt_032 %>:</b></div>
                                                                            </div>
                                                                            <div class="row">
                                                                                <div class="col-lg-12"><textarea rows="20" name="FM_kom_<%=y %>" class="form-control input-small"><%=timerKom %></textarea></div>
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
                                                 
                                            %>

                                            <td style="text-align:center; vertical-align:middle;">
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

                                            <td style="vertical-align:middle;">
                                                <a href="favorit.asp?id=<%=oRec2("id") %>&FM_medid=<%=medid %>&varTjDatoUS_man=<%=varTjDatoUS_man%>&func=fjernfavorit"><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></a>
                                            </td>

                                        </tr>
                                        <%
                                        end if
                                        oRec2.close

                                        end if
                                        oRec3.close

                                    next

                                     
                                    %>


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
                                        <td><input type="text" value="" id="next_akt_jobid_<%=next_akt_id %>" class="form-control input-small" readonly /></td>
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

                                    <tr style="border-bottom:inherit">  

                                        <!--<input type="hidden" value="0" name="FM_pa" />-->
                                        <input type="hidden" id="FM_jobid" value=""/>

                                        <td>
                                            
                                             
                                            
                                            <input type="text" class="FM_job form-control input-small" id="FM_job" value="" placeholder="<%=favorit_txt_015 %>" style="width:225px;"/>
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
                                            <!--<div id="dv_akt"></div> -->
                                            <select id="dv_akt" name="FM_aktivitetid" class="form-control input-small chbox_akt" style="width:195px;">
                                                <option SELECTED value="0" DISABLED>..</option>
                                            </select>                                     
                                        </td>
                                        
                                        <td style="text-align:center;" colspan="9">
                                           <input type="hidden" id="FM_medid_id" class="form-control input-small" value="<%=medid %>" />
                                            <input type="hidden" value="1" id="next_akt_id" />                                                                                      
                                            <a class="tilfoj_akt btn btn-default btn-sm" id="1" style="width:100%"><b><%=favorit_txt_017 %></b></a>
                                            <div id="dv_akttil"></div>
                                        </td>                                     
                                    </tr>
                                    

                            </tbody>

                        </table>
                            <div class="row">
                            <div class="col-lg-10">&nbsp</div>
                            <div class="col-lg-2 pad-b10">
                                <button type="submit" class="btn btn-success btn-sm pull-right"><b><%=favorit_txt_018 %></b></button>
                            </div>
                        </div>
                              <!--<input type="hidden" value="<%=ign_projgrp%>" id="ign_projgrp" />-->     
                            
                        </form>


                        <!-- Godkend / Afslut uge -->
                          <div class="row">
                            <div class="col-lg-10">
                                 
                                 <%
                                 
                                 dothis = 1
                                 SmiWeekOrMonth = 1

                                call akttyper2009(2)

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


                        <br /><br />
                           
                            <%
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
                        


                        <%
                             select case lto
                                case "tec", "esn", "xintranet - local", "outz"
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
                            
                            'response.Write manTimer
                            
                            
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
                             
                        %>


                        <div class="row">
                            <div class="col-lg-12">

                                <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                                    <thead>
                                        <tr>
                                            <th style="width:405px;"></th>
                                            <!--<th style="width:10%; text-align:center"><%response.Write Month(dateMan) & "-" & Day(dateMan)%></th>
                                            <th style="width:10%; text-align:center"><%response.Write Month(dateTir) & "-" & Day(dateTir) %></th>
                                            <th style="width:10%; text-align:center"><%response.Write Month(dateOns) & "-" & Day(dateOns) %></th>
                                            <th style="width:10%; text-align:center"><%response.Write Month(dateTor) & "-" & Day(dateTor) %></th>
                                            <th style="width:10%; text-align:center"><%response.Write Month(dateFre) & "-" & Day(dateFre) %></th>
                                            <th style="width:10%; text-align:center"><%response.Write Month(dateLor) & "-" & Day(dateLor) %></th>
                                            <th style="width:10%; text-align:center"><%response.Write Month(dateSon) & "-" & Day(dateSon) %></th>-->
                                            <th style="width:75px; text-align:center"><%=favorit_txt_006 %></th>
                                            <th style="width:75px; text-align:center"><%=favorit_txt_007 %></th>
                                            <th style="width:75px; text-align:center"><%=favorit_txt_008 %></th>
                                            <th style="width:75px; text-align:center"><%=favorit_txt_009 %></th>
                                            <th style="width:75px; text-align:center"><%=favorit_txt_010 %></th>
                                            <th style="width:75px; text-align:center"><%=favorit_txt_011 %></th>
                                            <th style="width:75px; text-align:center"><%=favorit_txt_012 %></th>
                                            <th style="width:75px; text-align:center"><%=favorit_txt_013 %></th>
                                        </tr>
                                    </thead>
                                    <tbody>

                                         <tr>
                                            <td><%=favorit_txt_020 %>:</td>
                                            <td style="text-align:center"><%=formatnumber(ntimMan, 2) %></td>
                                            <td style="text-align:center"><%=formatnumber(ntimTir, 2) %></td>
                                            <td style="text-align:center"><%=formatnumber(ntimOns, 2) %></td>
                                            <td style="text-align:center"><%=formatnumber(ntimTor, 2) %></td>
                                            <td style="text-align:center"><%=formatnumber(ntimFre, 2) %></td>
                                            <td style="text-align:center"><%=formatnumber(ntimLor, 2) %></td>
                                            <td style="text-align:center"><%=formatnumber(ntimSon, 2) %></td>
                                            <td style="text-align:center"><%=formatnumber(normtotal, 2) %></td>
                                        </tr>

                                        <tr>
                                            <td><%=favorit_txt_019 %>:</td>
                                            <td style="text-align:center"><%=formatnumber(manTimer, 2) %></td>
                                            <td style="text-align:center"><%=formatnumber(tirTimer, 2) %></td>
                                            <td style="text-align:center"><%=formatnumber(onsTimer, 2) %></td>
                                            <td style="text-align:center"><%=formatnumber(torTimer, 2) %></td>
                                            <td style="text-align:center"><%=formatnumber(freTimer, 2) %></td>
                                            <td style="text-align:center"><%=formatnumber(lorTimer, 2) %></td>
                                            <td style="text-align:center"><%=formatnumber(sonTimer, 2) %></td>
                                            <td style="text-align:center"><%=formatnumber(weekhourstotal, 2) %></td>

                                        </tr>
                                    
                                        


                                       

                                       
                                        <tr>
                                            <td><%=favorit_txt_021 %>:</td>
                                            <td style="text-align:center; color:<%=mancolor%>"><%=formatnumber(balMan, 2) %></td>
                                            <td style="text-align:center; color:<%=tircolor%>"><%=formatnumber(balTir, 2) %></td>
                                            <td style="text-align:center; color:<%=onscolor%>"><%=formatnumber(balOns, 2) %></td>
                                            <td style="text-align:center; color:<%=torcolor%>"><%=formatnumber(balTor, 2) %></td>
                                            <td style="text-align:center; color:<%=frecolor%>"><%=formatnumber(balFre, 2) %></td>
                                            <td style="text-align:center;"><%=formatnumber(balLor, 2) %></td>
                                            <td style="text-align:center;"><%=formatnumber(balSon, 2) %></td>
                                            <td style="text-align:center; color:<%=balcolor%>"><%=formatnumber(baltotal, 2) %></td>
                                        </tr>


                                         <%if lto = "wap" OR lto = "intranet - local" then 
                                          aty_sql_modregn = "tfaktim = 114" '31
                                          call timerDenneUge(medid, lto, varTjDatoUS_man, aty_sql_modregn, timerdenneuge_dothis, SmiWeekOrMonth)
                                           weekhourstotal = manTimer + tirTimer + onsTimer + torTimer + freTimer + lorTimer + sonTimer

                                            'balMan = balMan + manTimer
                                            'balTir = balTir + tirTimer 
                                            'balOns = balOns + onsTimer 
                                            'balTor = balTor + torTimer
                                            'balFre = balFre + freTimer 
                                            'balLor = balLor + lorTimer
                                            'balSon = balSon + sonTimer
                                            
                                            %>
                                         <tr>
                                            <td>Korrektion Flex:</td>
                                             <td style="text-align:center">(<%=formatnumber(manTimer, 2) %>)</td>
                                            <td style="text-align:center">(<%=formatnumber(tirTimer, 2) %>)</td>
                                            <td style="text-align:center">(<%=formatnumber(onsTimer, 2) %>)</td>
                                            <td style="text-align:center">(<%=formatnumber(torTimer, 2) %>)</td>
                                            <td style="text-align:center">(<%=formatnumber(freTimer, 2) %>)</td>
                                            <td style="text-align:center">(<%=formatnumber(lorTimer, 2) %>)</td>
                                            <td style="text-align:center">(<%=formatnumber(sonTimer, 2) %>)</td>
                                            <td style="text-align:center">(<%=formatnumber(weekhourstotal, 2) %>)</td>
                                        </tr>

                                        <%end if %>


                                        <%if lto = "wap" OR lto = "intranet - local" then 
                                            
                                           call fn_flexSaldoFYreal_norm(medid)
                                            
                                           %>

                                         <tr>
                                            <td>Flexsaldo: (opgjort pr. <b><%=formatdatetime(slDatoNormPrDdMinus1, 2) %></b> t.o.m ig�r)</td>
                                            <td colspan="7">&nbsp;</td>
                                         
                                            <td style="text-align:center;"><b><%=formatnumber(flexSaldoFYreal_norm, 2) %></b></td>
                                        </tr>
                                            
                                            
                                        <%end if%>

                                    </tbody>
                                </table>

                            </div>
                        </div>

                        


                    

                    </div>
                </div>
            </div>

            



    <%end select %>



        </div>
    </div>

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