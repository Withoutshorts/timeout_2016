
<%
    thisfile = "timetag_mobile" 
    uniksessionid = Session.SessionID & formatdatetime(now, 3)
    uniksessionid = replace(uniksessionid, ":", "")
%>

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->

<%'**** Søgekriterier AJAX **'

        'section for ajax calls
        if Request.Form("AjaxUpdateField") = "true" then
        Select Case Request.Form("control")

        
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

                    else

                          strJobogKunderTxt = strJobogKunderTxt & "<option value=""-1"" DISABLED>Antal job: "& c &"</option>"

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




            case "FN_createOutlay"
                
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
                intAntal = 1 'Antal findes ikke paa mobil, saa altid en
                regdato = year(now)&"/"&month(now)&"/"&day(now) 'fornu
                valuta = request("FM_udlaeg_valuta")
                
                select case lto 
                case "synergi1", "intranet - local", "epi", "epi_as", "epi_se", "epi_os", "epi_uk", "alfanordic"
                intKode = 1 'Intern
                case else 
	            intKode = 2 'Ekstern = viderefakturering
                end select

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

                unikId = request("unikId")

                'call MatTest(matregid, otf, medid, jobid, aktid, aftid, matId, strEditor, strDato, intAntal, regdato, valuta, intkode, personlig, bilagsnr, pris, salgspris, navn, gruppe, varenr, opretlager, betegnelse, mat_func, matreg_opdaterpris, matava)
                call indlaes_mat(matregid, otf, medid, jobid, aktid, aftid, matId, strEditor, strDato, intAntal, regdato, valuta, intkode, personlig, bilagsnr, pris, salgspris, navn, gruppe, varenr, opretlager, betegnelse, mat_func, matreg_opdaterpris, matava, unikId)


             
        
        end select
        response.end
        end if

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

    
    <!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
    <!--#include file="../inc/regular/top_menu_mobile.asp"-->


<head>

<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>TimeOut mobile</title>



<script src="js/timetag_web_jav_2018.js" type="text/javascript"></script>




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
    

    <!--
     <button onclick="getLocation()">Try It</button>
    <p id="demo"></p>
        -->
    <%

        if len(session("user")) = 0 then
	    %>
	    <!--#include file="../inc/errors/error_inc.asp"-->
	    <%
	    errortype = 5
	    call showError(errortype)
	    response.End
        end if

        select case lto
        case "tbg"
        tomobjid = 4
        case else
        tomobjid = session("tomobjid")
        end select


        call mobil_week_reg_dd_fn()
    %>

       


    <%call mobile_header %>


<script type="text/javascript" src="js/expence_jav2.js"></script>
<div class="container" style="height:100%">
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
            <input type="hidden" id="unikId" value="<%=uniksessionid %>" />

            <div class="row">
                <div class="col-lg-12"><span id="errorMessage" style="color:darkred;"></span></div>
            </div>

            <div class="row">
                <div class="col-lg-12">
                    <%if mobil_week_reg_job_dd = 1 then %>
                         <input type="hidden" id="FM_job" value="-1"/>
                            <select id="dv_job" name="FM_jobid" class="form-control">
                                 
                            </select>
                    <%else %>
                        <input type="text" id="FM_job" name="FM_job" value="" placeholder="<%=ttw_txt_014 %>" class="form-control" autocomplete="off" />
                        <input type="hidden" id="FM_jobid" name="FM_jobid" value="0"/>
                        <div id="dv_job" style="padding:5px 5px 5px 5px; display:none; visibility:hidden;"></div>
                    <%end if %>

                </div>
            </div>

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

            <div class="row">
                <div class="col-lg-12"><input type="text" name="FM_udlaeg_navn" id="FM_udlaeg_navn" placeholder="Navn" class="form-control"/></div>
            </div>

            <div class="row">
                <div class="col-lg-12">
                    <table style="width:100%">
                        <tr>
                            <td><input type="number" name="FM_udlaeg_belob" id="FM_udlaeg_belob" placeholder="Beløb" class="form-control" value="" /></td>
                            <td style="padding-left:5px;">
                                <select class="form-control" name="FM_udlaeg_valuta" id="FM_udlaeg_valuta">
                                <%
                                    strSQLValuta = "SELECT id, valutakode FROM valutaer"
                                    oRec.open strSQLValuta, oConn, 3
                                    while not oRec.EOF
                                    %><option value="<%=oRec("id") %>"><%=oRec("valutakode") %></option><%
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

           <!-- <div class="row">
                <div class="col-lg-12"><input type="number" name="FM_udlaeg_belob" id="FM_udlaeg_belob" placeholder="Beløb" class="form-control" value="" /></div>
            </div>

            <div class="row">
                <div class="col-lg-12">
                    <select class="form-control" name="FM_udlaeg_valuta" id="FM_udlaeg_valuta">
                        <%
                            strSQLValuta = "SELECT id, valutakode FROM valutaer"
                            oRec.open strSQLValuta, oConn, 3
                            while not oRec.EOF
                            %><option value="<%=oRec("id") %>"><%=oRec("valutakode") %></option><%
                            oRec.movenext
                            wend
                            oRec.close
                        %>
                    </select>
                </div>
            </div> -->

            <div class="row">                                     
                <div class="col-lg-12">
                    <select class="form-control" name="FM_udlaeg_gruppe" id="FM_udlaeg_gruppe">
                        <%
                            strSQLMatgrp = "SELECT id, navn FROM materiale_grp"
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
                        <option value="0">Ikke fakturerbar</option>
                        <option value="1">Fakturerbar</option>                                             
                    </select>
                </div>
            </div>

            <div class="row">                                     
                <div class="col-lg-12">
                    <select class="form-control" name="FM_udlaeg_form" id="FM_udlaeg_form">
                        <option value="1">Personlig udlæg</option>
                        <option value="0">Firmabetalt</option>                                             
                    </select>
                </div>
            </div>
                                        
            <div class="row">
                <div class="col-lg-12" style="text-align:center;">
                    <form ENCTYPE="multipart/form-data" action="../timereg/upload_bin.asp?matUpload=1&unikId=<%=uniksessionid %>&thisfile=timetag_mobile" method="post" id="image_upload">
                        <label class="btn btn-default btn-sm"><INPUT id="image-file" NAME="fileupload1" TYPE="file" style="width:400px; display:none;" onchange="readURL(this);"><b> Vælg billede </b></label>
                    </form>
                </div>
            </div>

            <div class="row">
                <div class="col-lg-12" style="text-align:center;">
                    <img id="imageholder" src="<%=fileLink %>" alt='' border='0' style="width:50px;">
                </div>
            </div>

            <script type="text/javascript">
                function readURL(input) {
                    // alert("show pic")

                    if (input.files && input.files[0]) {
                        var reader = new FileReader();

                        reader.onload = function (e) {
                            $('#imageholder')
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

        </div>

    </div>
</div>


<!--#include file="../inc/regular/footer_inc.asp"-->
