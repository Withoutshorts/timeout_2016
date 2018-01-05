<%response.buffer = true%>


<!--#include file="../inc/connection/conn_db_inc.asp"-->



<%
 '**** Søgekriterier AJAX **'
        'section for ajax calls
        if Request.Form("AjaxUpdateField") = "true" then
        Select Case Request.Form("control")

        case "FN_tjktimer_forecast"
                '** FORECAST ALERT 
         

                aktid = request("aktid")
                timerTastet = request("timer_tastet")
                usemrn = request("treg_usemrn")
                ibudgetaar = request("ibudgetaar")
                ibudgetmd = request("ibudgetaarMd")  
                aar = request("ibudgetUseAar")
                md = request("ibudgetUseMd")

               

                call ressourcefc_tjk(ibudgetaar, ibudgetmd, aar, md, usemrn, aktid, timerTastet)

                
             
                response.write feltTxtValFc
               


        case "FN_sogjobogkunde"


                '*** SØG kunde & Job       
                 lto = request("lto")
    
               
                if len(trim(request("jq_newfilterval"))) <> 0 then
                filterVal = 1 
                jobkundesog = request("jq_newfilterval")
                

                'if session("mid") = 1 then
                '*** ÆØÅ **'
                'call jq_format(jobkundesog)
                'jobkundesog = jq_formatTxt
                'end if
                

                else
                filterVal = 0
                jobkundesog = "6xxxxxfsdf554"
                end if
        
                medid = request("jq_medid")

                call positiv_aktivering_akt_fn()
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
                 strSQLSogKri = " AND (jobnr LIKE '"& jobkundesog &"%' OR jobnavn LIKE '%"& jobkundesog &"%' OR "_
                 &" kkundenavn LIKE '"& jobkundesog &"%' OR kkundenr = '"& jobkundesog &"' OR k.kinit = '"& jobkundesog &"') AND kkundenavn <> ''"
                 lmt = 50
                 else
                 strSQLSogKri = ""
                 lmt = 250
                 end if
            
                 lastKid = 0
                
                strSQL = "SELECT j.id AS jid, j.jobnavn, j.jobnr, j.jobstatus, k.kkundenavn, k.kkundenr, k.kid FROM timereg_usejob AS tu "_ 
                &" LEFT JOIN job AS j ON (j.id = tu.jobid) "_
                &" LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "_
                &" WHERE tu.medarb = "& medid &" AND (j.jobstatus = 1 OR j.jobstatus = 3) "& strSQLPAkri &" "& strSQLDatokri 
                
                strSQL = strSQL & strSQLSogKri
                strSQL = strSQL &" GROUP BY j.id ORDER BY kkundenavn, jobnavn LIMIT "& lmt       
    
                'if session("mid") = 1 then
                ' response.write "<option>strSQL " & strSQL & "</option>"
                ' response.end
                'end if
    
                select case lto 
                case "xcflow", "xintranet - local"
    

                case else

                    if (jobkundesog = "-1") then
                        
                        if len(trim(week_txt_009)) = 0 then
                        week_txt_009 = "Select Job"
                        end if

                        strJobogKunderTxt = strJobogKunderTxt & "<option value=""-1"" SELECTED>"& week_txt_009 &":</option>"
                    end if            

                end select

                 
                k = 0
                oRec.open strSQL, oConn, 3
                while not oRec.EOF
        
             

                if lastKnavn <> oRec("kkundenavn") then
    

                    if k <> 0 then
                    ' strJobogKunderTxt = strJobogKunderTxt &"<br>"
                    strJobogKunderTxt = strJobogKunderTxt & "<option DISABLED></option>"
                    end if
    
                'strJobogKunderTxt = strJobogKunderTxt & oRec("kkundenavn") &" "& oRec("kkundenr") &"<br>"

                 strJobogKunderTxt = strJobogKunderTxt & "<option DISABLED>"& oRec("kkundenavn") &" "& oRec("kkundenr") &"</option>"
    
                end if 
                 
               ' strJobogKunderTxt = strJobogKunderTxt & "<input type=""hidden"" id=""hiddn_job_"& oRec("jid") &""" value="""& oRec("jobnavn") & " ("& oRec("jobnr") &")"">"
               ' strJobogKunderTxt = strJobogKunderTxt & "<a class=""chbox_job"" id=""chbox_job_"& oRec("jid") &""" value="& oRec("jid") &">"& oRec("jobnavn") & " ("& oRec("jobnr") &")" &"</a><br>" 
                
                strJobogKunderTxt = strJobogKunderTxt & "<option value="& oRec("jid") &">"& oRec("jobnavn") & " ("& oRec("jobnr") &")" &"</option>"

                k = k + 1
                lastKnavn = oRec("kkundenavn") 
                oRec.movenext
                wend
                oRec.close

              
                if cint(k) = 0 then
                strJobogKunderTxt = strJobogKunderTxt & "<option value=""-1"" DISABLED>"& week_txt_010 &"</option>"
                end if


                    '*** ÆØÅ **'
                    call jq_format(strJobogKunderTxt)
                    strJobogKunderTxt = jq_formatTxt


                    response.write strJobogKunderTxt

                end if    


        
          case "FN_sogakt"

               
                '*** Søg Aktiviteter 
                'strAktFcSaldoTxt = ""            
                strAktTxt = ""

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

                 '*** Sales / tilbud kun Salgsaktiviteter
                '(a.fakturerbar = 6 AND j.jobstatus = 3)
                jobstatusTjk = 1
                strSQLtilbud = "SELECT jobstatus FROm job WHERE id = "& jobid
                oRec5.open strSQLtilbud, oConn, 3
                if not oRec5.EOF then

                jobstatusTjk = oRec5("jobstatus")

                end if
                oRec5.close

                if cint(jobstatusTjk) = 3 then 'tilbud
                onlySalesact = " AND a.fakturerbar = 6"
                else
                onlySalesact = ""
                end if


                'positiv aktivering
                pa = 0
                'if len(trim(request("jq_pa") )) <> 0 then
                'pa = request("jq_pa") 
                'else
                'pa = 0
                'end if
        
                call positiv_aktivering_akt_fn()
                pa = pa_aktlist
                pa_only_specifikke_akt = positiv_aktivering_akt_val
            
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

              

                if filterVal <> 0 then
            
                 
    
                 'strAktTxt = strAktTxt &"<span style=""color:#999999; font-size:9px; float:right;"" class=""luk_aktsog"">X</span>"    
                         

                     '** Select eller søgeboks
                    call mobil_week_reg_dd_fn()
                    
                    
                    if cint(mobil_week_reg_akt_dd) <> 1 then 'AND aktsog <> "-1" 
                    strSQlAktSog = "AND navn LIKE '%"& aktsog &"%'"
                    else
                    strSQlAktSog = ""

                            '** Forvalgt 1 aktivitet
                            if cint(mobil_week_reg_akt_dd_forvalgt) <> 1 AND cint(mobil_week_reg_akt_dd) = 1 then
                            strAktTxt = strAktTxt & "<option value=""-1"">Choose..</option>" 
                            end if

                    end if



                   if cint(pa) = 1 then '**Kun på Personlig aktliste
    
    
                       'Positiv aktivering
                       if cint(pa_only_specifikke_akt) then

                       strSQL = "SELECT a.id AS aid, navn AS aktnavn "_
                       &" FROM timereg_usejob AS tu LEFT JOIN aktiviteter AS a ON (a.id = tu.aktid "& onlySalesact &") "_
                       &" WHERE tu.medarb = "& medid &" AND tu.jobid = "& jobid &" AND aktid <> 0 "& strSQlAktSog &" AND aktstatus = 1 AND ("& aty_sql_hide_on_treg &") "& forecastAktids &" "& onlySalesact &" ORDER BY fase, sortorder, navn LIMIT 250"   
                       'AND ("& replace(aty_sql_realhours, "tfaktim", "a.fakturerbar") &")
                       else 

                       strSQL = "SELECT a.id AS aid, navn AS aktnavn "_
                       &" FROM timereg_usejob AS tu LEFT JOIN aktiviteter AS a ON (a.job = tu.jobid) "_
                       &" WHERE tu.medarb = "& medid &" AND tu.jobid = "& jobid &" AND aktid = 0 "& strSQlAktSog &" AND aktstatus = 1 AND ("& aty_sql_hide_on_treg &") "& forecastAktids &" "& onlySalesact &" ORDER BY fase, sortorder, navn LIMIT 250"
                       'AND ("& replace(aty_sql_realhours, "tfaktim", "a.fakturerbar") &")
                       end if


                   else


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


           

                   

                
                   

                   strSQL = "SELECT a.id AS aid, navn AS aktnavn, projektgruppe1, projektgruppe2, projektgruppe3, "_
                   &" projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM aktiviteter AS a "_
                   &" WHERE a.job = " & jobid & " "& strSQLDatoKri &" "& strSQlAktSog &" AND aktstatus = 1 AND ("& aty_sql_hide_on_treg &") "& forecastAktids &" "& onlySalesact &" ORDER BY fase, sortorder, navn LIMIT 250"      
    

               
            
                end if

                'response.write "strSQL " & strSQL & ""
                'response.write "<option>strSQL " & strSQL & "</option>"
                'response.flush

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
                 
                'strAktTxt = strAktTxt & "<input type=""hidden"" id=""hiddn_akt_"& oRec("aid") &""" value="""& oRec("aktnavn") &""">"
                'strAktTxt = strAktTxt & "<a class=""chbox_akt"" id=""chbox_akt_"& oRec("aid") &""" value="& oRec("aid") &">"& oRec("aktnavn") &"</a><br>" 

                if cint(aktBudgettjkOn) = 1 then


                if len(trim(feltTxtValFc)) <> 0 then
                fcsaldo_txt = " (fc. Saldo: "& formatnumber(feltTxtValFc, 2) & " / "& formatnumber(fctimer,2) &" t.)"
                else
                feltTxtValFc = 0
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
                
                
                'strAktFcSaldoTxt = strAktFcSaldoTxt & "<input type=""text"" value="& feltTxtValFc &" id=""FM_fcs_"& oRec("aid") &">"
                strAktTxt = strAktTxt & "<option value="& oRec("aid") &" "& optionFcDis &">"& oRec("aktnavn") &" "& fcsaldo_txt &"</option>" 
                
                end if
                
                afundet = afundet + 1
                oRec.movenext
                wend
                oRec.close

                
                if afundet = 0 then
                strAktTxt = strAktTxt & "<option value=""-1"" DISABLED>"& week_txt_011 &"</option>" 
                end if          



                    '*** ÆØÅ **'
                    call jq_format(strAktTxt)
                    strAktTxt = jq_formatTxt


                    response.write strAktTxt & strAktFcSaldoTxt

                end if    




        end select
        response.end
        end if




tloadA = now
    
   %>

<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../timereg/inc/convertDate.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/treg_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<% 

if len(session("user")) = 0 then
	%>
	
	<%
	errortype = 5
	call showError(errortype)
	
	else
	

        'SET LOCALE for denne side
        select case sprog
        case 1
        Session.LCID = 1030
        case 2
        Session.LCID = 2057
        case 3
        Session.LCID = 1053
        case 4
        Session.LCID = 2068
        case 5
        Session.LCID = 1034
        case 6
        Session.LCID = 1031
        case 7
        Session.LCID = 1036
        case else
        Session.LCID = 1030
        end select


	function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
	end function
	
	func = request("func")
	id = request("id")
	thisfile = "ugeseddel_2011.asp"
	media = request("media")

    if len(request("touchMonitor")) <> 0 then
        touchMonitor = request("touchMonitor")

        %>
        <meta http-equiv="refresh" content="120;url=monitor.asp?func=startside" />
        <%
    else
        touchMonitor = 0
    end if


    if len(trim(request("nomenu"))) <> 0 AND request("nomenu") <> 0 then
    nomenu = 1
    else
    nomenu = 0
    end if

    if len(trim(request("usemrn"))) <> 0 then
    usemrn = request("usemrn")
    else
    usemrn = session("mid")
    end if
    


            
         

    level = session("rettigheder")
 
    rdir = request("rdir")
    select case rdir
    case "logindhist"
    rdirfile = "logindhist_2011.asp"
    case "godkenduge"
    rdirfile = "../timereg/godkenduge.asp"
    case else
	rdirfile = "ugeseddel_2011.asp"
	end select
    

    if len(trim(request("medarbsel_form"))) <> 0 then

        'response.write "HER"
        'response.end

    if len(trim(request("FM_visAlleMedarb"))) <> 0 then
    visAlleMedarbCHK = "CHECKED"
    visAlleMedarb = 1
    else
    visAlleMedarbCHK = ""
    visAlleMedarb = 0
    end if

    else

                    if request.cookies("tsa")("visAlleMedarb") = "1" then
                    visAlleMedarbCHK = "CHECKED"
                    visAlleMedarb = 1
                    else
                    visAlleMedarbCHK = ""
                    visAlleMedarb = 0
                    'usemrn = session("mid")
                    end if
    end if
    response.cookies("tsa")("visAlleMedarb") = visAlleMedarb



    
    call ersmileyaktiv()
	
    call medarb_teamlederfor
    
    '** Select eller søgeboks
    call mobil_week_reg_dd_fn()
    'mobil_week_reg_job_dd = mobil_week_reg_job_dd
    'mobil_week_reg_akt_dd = 1
    'mobil_week_reg_akt_dd = mobil_week_reg_akt_dd    
    
        
    varTjDatoUS_man = request("varTjDatoUS_man")

    'if len(trim(request("varTjDatoUS_son"))) <> 0 then
	'varTjDatoUS_son = request("varTjDatoUS_son")
    'else
    varTjDatoUS_son = dateAdd("d", 6, varTjDatoUS_man)
    'end if

    'Response.Write varTjDatoUS_man
    'Response.end

    datoMan = day(varTjDatoUS_man) &"/"& month(varTjDatoUS_man) &"/"& year(varTjDatoUS_man)
    datoSon = day(varTjDatoUS_son) &"/"& month(varTjDatoUS_son) &"/"& year(varTjDatoUS_son)

    next_varTjDatoUS_man = dateadd("d", 7, varTjDatoUS_man)
	next_varTjDatoUS_son = dateadd("d", 7, varTjDatoUS_son)
    next_varTjDatoUS_man = year(next_varTjDatoUS_man) &"/"& month(next_varTjDatoUS_man) &"/"& day(next_varTjDatoUS_man)
	next_varTjDatoUS_son = year(next_varTjDatoUS_son) &"/"& month(next_varTjDatoUS_son) &"/"& day(next_varTjDatoUS_son)


    prev_varTjDatoUS_man = dateadd("d", -7, varTjDatoUS_man)
	prev_varTjDatoUS_son = dateadd("d", -7, varTjDatoUS_son)
    prev_varTjDatoUS_man = year(prev_varTjDatoUS_man) &"/"& month(prev_varTjDatoUS_man) &"/"& day(prev_varTjDatoUS_man)
	prev_varTjDatoUS_son = year(prev_varTjDatoUS_son) &"/"& month(prev_varTjDatoUS_son) &"/"& day(prev_varTjDatoUS_son)





    call positiv_aktivering_akt_fn()
    call smileyAfslutSettings()

	'Response.Write "media " & media
	
	
	'*** Sætter lokal dato/kr format. *****
   'call lcid_sprog(session("mid"))
   'Session.LCID = lcid_sprog_Val
	lcid_sprog_val = Session.LCID

    select case func 
	case "-"

    case "xxxopdaterstatus" 'NOT IN USE 20170927


        
        'response.write "request(ids)" & request("ids")
        
          ujid = split(request("ids"), ",")

                if len(trim(request("FM_godkendt"))) <> 0 then
                uGodkendt = request("FM_godkendt")
                else
                uGodkendt = 0
                end if
	
    for u = 0 to UBOUND(ujid)
	
	            editor = session("user")


                'if len(trim(request("FM_godkendt_"& trim(ujid(u))))) <> 0 then
                'uGodkendt = request("FM_godkendt_"& trim(ujid(u)))
                'else
                'uGodkendt = 0
                'end if
              

				strSQL = "UPDATE timer SET godkendtstatus = "& uGodkendt &", "_
				&"godkendtstatusaf = '"& editor &"' WHERE tid = " & ujid(u) & " AND overfort = 0 "
				
			    'Response.write strSQL &"<br>"
				
				oConn.execute(strSQL)
				
	next

       'response.end

    'Response.Redirect "ugeseddel_2011.asp?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&varTjDatoUS_son="&varTjDatoUS_son&"&nomenu="&nomenu

    Response.Redirect rdirfile & "?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&nomenu="&nomenu '&"&varTjDatoUS_son="&varTjDatoUS_son
	
	case "godkendugeseddel"
	
	'*** Godkender ugeseddel ***'



	
	    varTjDatoUS_manSQL = year(varTjDatoUS_man)&"/"&month(varTjDatoUS_man)&"/"&day(varTjDatoUS_man)
        varTjDatoUS_sonSQL = year(varTjDatoUS_son)&"/"&month(varTjDatoUS_son)&"/"&day(varTjDatoUS_son)
	    
	   
        
         '**** Godkender timer i dbne uge der bnliver godkendt
        call godkenderTimeriUge(usemrn, varTjDatoUS_manSQL, varTjDatoUS_sonSQL, SmiWeekOrMonth)
       

        '*** Godkend uge status ****'
        call godekendugeseddel(thisfile, session("mid"), usemrn, varTjDatoUS_man)
	    
	
	
	Response.Redirect rdirfile & "?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&nomenu="&nomenu '&"&varTjDatoUS_son="&varTjDatoUS_son
	

    case "afvisugeseddel"
	
	'*** Afviser ugeseddel ***'
    if len(trim(request("FM_afvis_grund"))) <> 0 then
    txt = replace(request("FM_afvis_grund"), "'", "")
    else
    txt = ""
    end if

	call afviseugeseddel(thisfile, session("mid"), usemrn, varTjDatoUS_man, varTjDatoUS_son, txt)
	    
	
	Response.Redirect rdirfile & "?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&nomenu="&nomenu '&"&varTjDatoUS_son="&varTjDatoUS_son
	
    
    
    case "adviserugeafslutning"

    call afslutugereminder(thisfile, session("mid"), usemrn, varTjDatoUS_man, varTjDatoUS_son, txt)

  
	Response.Redirect rdirfile & "?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&showadviseringmsg=1&nomenu="&nomenu '&"&varTjDatoUS_son="&varTjDatoUS_son
	

    case "slet_tip"

        %>
        	
        <%

        id = request("id")

        slttxt = ""& week_txt_001 &"<br> "& week_txt_002 &""
        slturl = rdirfile & "?func=slet_tip_ok&id="&id&"&usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&showadviseringmsg=1&nomenu="&nomenu '&"&varTjDatoUS_son="&varTjDatoUS_son

       call sltque(slturl,slttxt,slturlalt,slttxtalt,110,90)

    

    case "slet_tip_ok"

        id = request("id")

        sqlInTimerImpTemp = "DELETE FROM timer_import_temp WHERE id = "& id
        oConn.Execute(sqlInTimerImpTemp)

    Response.Redirect rdirfile & "?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&showadviseringmsg=1&nomenu="&nomenu '&"&varTjDatoUS_son="&varTjDatoUS_son

    case else



     if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
    <!--#include file="../inc/regular/top_menu_mobile.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	response.end
    end if

   
    '*** Vis kun aktiviteter med forecast på
    call aktBudgettjkOn_fn()
    '*** Skal akt lukkes for timereg. hvis forecast budget er overskrddet..?
    '** MAKS budget / Forecast incl. peridoe afgrænsning
    call akt_maksbudget_treg_fn()

    call akttyper2009(2)
	

    if len(trim(request("FM_datoer"))) <> 0 then 'redirect fra timereg_akt_2006 HUSK Dato
        tregDato = day(request("FM_datoer")) &"/"& month(request("FM_datoer")) &"/"& year(request("FM_datoer"))

        ddDato_ugedag = day(tregDato) &"/"& month(tregDato) &"/"& year(tregDato)
        ddDato_ugedag_w = datepart("w", ddDato_ugedag, 2, 2)
            
        varTjDatoUS_man = dateAdd("d", -(ddDato_ugedag_w-1), ddDato_ugedag)

    else
        
        if (datepart("ww", varTjDatoUS_man, 2, 2) = datepart("ww", day(now) &"/"& month(now) &"/"& year(now), 2, 2)) then  'Indeværende uge vises DD
        tregDato = day(now) &"/"& month(now) &"/"& year(now)
        else
        tregDato = day(varTjDatoUS_man) &"/"& month(varTjDatoUS_man) &"/"& year(varTjDatoUS_man)
        end if

    end if
	
	
	if media <> "export" then
	
	
	if media <> "print" then
	
	
	%>

	
    <SCRIPT src="js/ugeseddel_2011_jav.js"></script>
    <SCRIPT src="../timereg/inc/smiley_jav.js"></script>

    <%call browsertype() 
    if (browstype_client <> "ip") AND (touchMonitor <> 1) then
        
        call menu_2014()

    else

        ddDato_ugedag = day(now) &"/"& month(now) &"/"& year(now)
        ddDato_ugedag_w = datepart("w", ddDato_ugedag, 2, 2)
            
        varTjDatoUS_man_tt = dateAdd("d", -(ddDato_ugedag_w-1), ddDato_ugedag)
        
        if touchMonitor <> 1 then
        call mobile_header 
        end if

    end if
    %>    


   <%      
         
    
    else 
	

    end if 'print%>
	
	<%end if 'eksport
        
        


   
        %>

 

 <script type="text/javascript" src="js/plugins/flot/jquery.flot.js"></script>
    <script type="text/javascript" src="js/demos/flot/stacked-vertical_ugetotal.js"></script>
   
   
	
	<!-------------------------------Sideindhold------------------------------------->

   
 


   <%if (browstype_client <> "ip") AND (touchMonitor <> 1) then %>
<div class="wrapper">
 <div class="content">   
     <%end if %>

 
     <%if cint(touchMonitor) = 1 then %>
     <div class="container" style="width:100%; height:100%">
     <%else %>
     <div class="container">
    <%end if %>

      <div class="portlet">
        <% if browstype_client <> "ip" then %>
          <%if cint(touchMonitor) <> 1 then %>
          <h3 class="portlet-title"><u><%=tsa_txt_337%></u><!-- ugeseddel --></h3>
          <%else 
          medarb_navn = session("user")'request("medarb_navn")
          medid = session("mid") 'request("medid")    
          %>
          <div class="portlet-title">

              <table style="width:100%;">
                  <tr>
                      <td><h3 style="color:black"><%=medarb_navn %> <!--(<%=medid %>)-->

                          <%
                            '******* Tjekker logindtid ********'
                            strSQL = "SELECT id, login FROM login_historik WHERE mid = "& medid &" AND stempelurindstilling > 0 ORDER BY id DESC"
                            oRec.open strSQL, oConn, 3
                            if not oRec.EOF then
                            login = oRec("login")
                            
                            end if
                            oRec.close
                            
                              
                          %><br />
                          <span style="font-size:11px;">Logget på kl. <%=login %></span>


                          </h3></td>
                      <td style="text-align:right"><img style="width:15%; margin-bottom:2px; margin-left:55px" src="img/cflow_logo.png" /></td>
                  </tr>
              </table>

           <!--   <div class="row">
                  <div class="col-lg-4"><h3><%=medarb_navn %> <!--(<%=medid %>)--></h3></div>
                 <!-- <div class="col-lg-6"></div>
                  <div class="col-lg-2"><img style="max-width:50%; margin-bottom:2px; margin-left:55px" src="img/cflow_logo.png" /></div>
              </div> -->
          </div>
          <%end if %>
        <%end if %>

        <div class="portlet-body">
            
      
        <%   
                                
            timerdenneuge_dothis = 0
            call timerDenneUge(usemrn, lto, varTjDatoUS_man, aty_sql_realhours, timerdenneuge_dothis, SmiWeekOrMonth)

                ugetotal = manTimer + tirTimer + onsTimer + torTimer + freTimer + lorTimer + sonTimer
                manTimer = replace(manTimer, ",",".")
                tirTimer = replace(tirTimer, ",",".")
                onsTimer = replace(onsTimer, ",",".")
                torTimer = replace(torTimer, ",",".")
                lorTimer = replace(lorTimer, ",",".")
                sonTimer = replace(sonTimer, ",",".")

                call normtimerPer(usemrn, varTjDatoUS_man, 6, 0)

            norm_ugetotal = ntimMan + ntimTir + ntimOns + ntimTor + ntimFre + ntimLor + ntimSon

            if norm_ugetotal <> 0 then
            'weeknormpro = ugetotal / norm_ugetotal * 100
            weeknormpro = ugetotal / norm_ugetotal * 100
            weeknormpro = replace(weeknormpro, ",",".")
            else
            weeknormpro = 0
            end if

            %>

            <input type="hidden" id="dagman" value="<%=manTimer %>" />
            <input type="hidden" id="dagtir" value="<%=tirTimer %>" />
            <input type="hidden" id="dagons" value="<%=onsTimer %>" />
            <input type="hidden" id="dagtor" value="<%=torTimer %>" />
            <input type="hidden" id="dagfre" value="<%=freTimer %>" />
            <input type="hidden" id="daglor" value="<%=lorTimer %>" />
            <input type="hidden" id="dagson" value="<%=sonTimer %>" />

        
                

            
            <%
          

           
    if media <> "print" AND len(trim(strSQLmids)) > 0 then 'Hvis man er level 1 eller teamleder vil len(trim(strSQLmids)) ALTID VÆRE > 16 %>
    <%
        
        if browstype_client <> "ip" then
    %>
    <%
    if touchMonitor <> 1 then
    %>
    <div class="well well-white">
    <%end if %>
        
        <%select case lto 
        case "bf"
            
        case else%>

        <%if cint(stempelurOn) = 1 then 
            wdth = 225
        else
            wdth = 120
        end if

        %>
          
            <%if touchMonitor <> 1 then %>
            <div style="position:relative; background-color:#ffffff; border:1px #cccccc solid; border-bottom:0; padding:4px; width:<%=wdth%>px; top:-70px; left:880px; z-index:0; font-size:11px;">
           
            <a href="../timereg/<%=lnkTimeregside%>" class="vmenu"><%=replace(tsa_txt_031, " ", "")%> >></a>
          

            <%if cint(stempelurOn) = 1 then %>
            &nbsp;|&nbsp;<a href="../timereg/<%=lnkLogind%>" class="vmenu"><%=tsa_txt_340 %> >></a>
            <%end if%>

           
            </div>
            <%end if %>

         

        <%end select %>
      
        <%select case lto
         case "xepi2017", "xintranet - local"

            thisEpiDay = datepart("w", now, 2,2)  
            visGuide = 0
            strSQLvisguide = "SELECt visguide FROM medarbejdere WHERE mid = "& session("mid")
            oRec.open strSQLvisguide, oConn, 3
            if not oRec.EOF then
             visGuide = oRec("visguide")
            end if
            oRec.close

            'session("show_skiftStSide_vist") <> "11" AND
            if (thisEpiDay = 1 OR thisEpiDay = 2) AND visGuide <> 99 then
            show_skiftStSide = 1
            'session("show_skiftStSide_vist") = 1
            else
            show_skiftStSide = 0
            end if

         case else
            show_skiftStSide = 0
         end select 
            
            
         if cint(show_skiftStSide) = 1 then%>
        <form action="../timereg/timereg_akt_2006.asp">
            <input type="hidden" name="func" value="epitregpage" />
            <div id="d_epinote" style="position:absolute; top:70px; left:850px; width:400px; border:10px #999999 solid; z-index:10000000; background-color:#ffffff; padding:20px;">
                <span style="float:right;"><a href="#" id="a_epinote" style="color:red;">[X]</a></span>
                <h4>Welcome to Epinion version 2017</h4> 
                We have put together all Epinion versions Of TimeOut in one solution: Epi 2017.<br /><br />
                You will be able to search for all your projects across DK, NO and UK in this database. <br /><br />
                If You wish to keep your old timecording page and use it as your default, you can choose it heere. 
                <br /><input type="checkbox" name="FM_oldtreg" value="1" /> I want the old timecording page.<br /> 
                <input type="checkbox" name="FM_dontshowvisguide" value="1" /> Don't show this message again. <input type="submit" class="btn btn-sm btn-default" value=" >> " /><br /><br>
                You can always change it in your userprofile or select it form the main menu.<br /><br />
                
                If you are missing a project don't hesitate to write to Epinion Finance.<br /><br />
                If You experience any troubles with the new version send an email and a screenshot to support@outzource.dk  
                


            </div>

        </form>
        <%end if%>

        <form id="filterkri" method="post" action="ugeseddel_2011.asp">
            <input type="hidden" name="FM_sesMid" id="FM_sesMid" value="<%=session("mid") %>">
            <input type="hidden" name="medarbsel_form" id="medarbsel_form" value="1">
            <input type="hidden" name="nomenu" id="nomenu" value="<%=nomenu %>">
            <input type="hidden" name="varTjDatoUS_man" id="varTjDatoUS_man" value="<%=varTjDatoUS_man %>">
           
            <%if touchMonitor <> 1 then %>
            <div class="row">
                <div class="col-lg-6">
                      <table style="font-size:100%; color:black">
                    
                    <tr>
                        <!--<td style="padding-right:5px; vertical-align:text-top; width:80px;"><input type="CHECKBOX" name="FM_visallemedarb" id="FM_visallemedarb" value="1" <%=visAlleMedarbCHK %> /> <%=tsa_txt_357%> </td>-->
                        <td style="padding:0px 0px 4px 12px;">
                            <input type="CHECKBOX" name="FM_visallemedarb" id="FM_visallemedarb" value="1" <%=visAlleMedarbCHK %> /> <%=tsa_txt_357%><br />
                    
                   
               
                    <% 
				    call medarb_vaelgandre
                    %>
                            </td>
                        <td>&nbsp;</td>
                        </tr>
                        </table>
                </div>
            </div>
            <%else %>
            <input type="hidden" name="usemrn" id="usemrn" value="<%=usemrn %>" />
            <%end if %>
                  
        </form>

	
       
	
	<%
   
       end if 'media
	
        'if datePart("d", now, 2,2) < 4 OR datePart("d", now, 2,2) > 29 then 

	   if media <> "print" then
        %> 
        <%
            
            if browstype_client <> "ip" then
        %>

        <% 
        if touchMonitor = 1 then
        txtClass = "large"
        inputHeight = "150%"
        inputFont = "150%"
        paddingTop = "30px"
        'response.Write "<br><br><br>"
        rdir_timereg = "touchMonitor"
        else
        txtClass = "small"
        inputWidth = "100%"
        inputFont = "100%"
        paddingTop = "10px"
        rdir_timereg = "ugeseddel_2011"
        end if
        %>

      
            <div class="row"> 
     
           
                     <%if cint(touchMonitor) <> 1 then 
                    %><div class="col-lg-7"> <%
                    else
                    %><div class="col-lg-7"> <%
                    end if%>
                     
     
          
                      <form id="container" action="../timereg/timereg_akt_2006.asp?func=db&rdir=<%=rdir_timereg %>" method="post">
      
                    <%if cint(touchMonitor) <> 1 then 
                    mgnRight = 250    
                    wdtTbl = 60
                    else
                    mgnRight = 0
                    wdtTbl = 60
                    end if%>

                    <table style="font-size:100%; display:inline-table; margin-right:<%=mgnRight%>px; width:<%=wdtTbl%>%;">
                    
                    <%if touchMonitor <> 1 then %>
                    <tr>
                        <!--<td style="padding-right:5px; vertical-align:text-top; width:80px;"><b><%=tsa_txt_183%>:</b></td>-->
                        <td style="padding-left:10px">
                            <input type="hidden" id="Hidden5" name="year" value="<%=year(now) %>"/>
                               
                              <div class='input-group date'>
                                      <input type="text" style="width:300px;" class="form-control input-small" name="FM_datoer" id="jq_dato" value="<%=tregDato %>" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>

                           
                        </td>
                        <td>&nbsp;</td>
                    </tr>
                    <%else %>
                    <input type="hidden" id="Hidden5" name="year" value="<%=year(now) %>"/>
                    <input type="hidden" name="FM_datoer" id="jq_dato" value="<%=tregDato %>" />
                    <%end if %>
                    
                   
                        <!--<td style="padding-right:5px; padding-top:10px; vertical-align:text-top;"><b><%=left(tsa_txt_066, 5) %>/<%=tsa_txt_236 %>:</b></td>-->
                        <%if cint(touchMonitor) = 1 then 
                        rwspan = 1 '3
                        %><tr><td style="padding-left:10px; font-size:15px; text-align:left;">Komponent/Aktivitet</td></tr><%
                        else
                        rwspan = 1
                        end if%>

                        
                     <tr>
                        <td style="padding-top:<%=paddingTop%>; padding-left:10px;" rowspan="<%=rwspan %>">
                             <input type="hidden" name="varTjDatoUS_man" value="<%=varTjDatoUS_man %>">
        
                            <input type="hidden" name="usemrn" id="treg_usemrn" value="<%=usemrn%>">
                            <input type="hidden" id="lto" value="<%=lto%>">

                            <!--<input type="hidden" id="Hidden3" name="FM_dager" value=""/>-->
                            <input type="hidden" id="Hidden4" name="FM_dager" value="0"/><!-- , xx -->
                            <input type="hidden" id="Hidden2" name="FM_feltnr" value="0"/>
                            <input type="hidden" id="FM_pa" name="FM_pa" value="<%=pa_aktlist %>"/>
                            
                            <input type="hidden" id="FM_medid" name="FM_medid" value="<%=usemrn %>"/>
                            
                            <input type="hidden" id="FM_medid_k" name="FM_medid_k" value="<%=usemrn%>"/>
                            <input type="hidden" id="" name="FM_vistimereltid" value="0"/>
                            <input type="hidden" id="lcid_sprog" name="" value="<%=lcid_sprog_val %>"/>

                            <input type="hidden" id="mobil_week_reg_akt_dd" name="" value="<%=mobil_week_reg_akt_dd %>"/>
                            <input type="hidden" id="mobil_week_reg_job_dd" name="" value="<%=mobil_week_reg_job_dd %>"/>
                            
                           
                            <!-- Forecast max felter -->
                            <input type="hidden" id="aktnotificer_fc" name="" value="0"/>
                            <input type="hidden" id="akt_maksbudget_treg" value="<%=akt_maksbudget_treg%>">  
                            <input type="hidden" id="akt_maksforecast_treg" value="<%=akt_maksforecast_treg%>">
                            <input type="hidden" id="aktBudgettjkOn_afgr" value="<%=aktBudgettjkOn_afgr%>">
                            <input type="hidden" id="regskabsaarStMd" value="<%=datePart("m", aktBudgettjkOnRegAarSt, 2,2)%>">
                            <input type="hidden" id="regskabsaarUseAar" value="<%=datepart("yyyy", varTjDatoUS_man, 2,2)%>">
                            <input type="hidden" id="regskabsaarUseMd" value="<%=datepart("m", varTjDatoUS_man, 2,2)%>">     
                            
                            
                            <%if cint(mobil_week_reg_job_dd) = 1 then %>

                          
                             <input type="hidden" id="FM_job_0" value="-1"/>
                             <select id="dv_job_0" name="FM_jobid" style="font-size:<%=inputFont %>; height:<%=inputHeight%>;" class="form-control input-<%=txtClass %> chbox_job">
                                 <option value="-1"><%=left(tsa_txt_145, 4) %>..</option>
                                 <!--<option value="0">..</option>-->
                             </select>

                            <%else %>
                            <input type="text" style="font-size:<%=inputFont%>; height:<%=inputHeight%>" id="FM_job_0" name="FM_job" placeholder="<%=tsa_txt_066 %>/<%=tsa_txt_236 %>" class="FM_job form-control input-<%=txtClass %>"/>
                           <!-- <div id="dv_job_0" class="dv-closed dv_job" style="border:1px #cccccc solid; padding:10px; visibility:hidden; display:none;"></div>--> <!-- dv_job -->

                             <select id="dv_job_0" class="form-control input-small chbox_job" size="10" style="visibility:hidden; display:none;">
                                 <option><%=tsa_txt_534%>..</option>
                             </select>

                            <input type="hidden" id="FM_jobid_0" name="FM_jobid" value="0"/>
                            <%end if %>

                        </td>
                     
                       <%'if cint(touchMonitor) = 1 then 
                        
                        'else
                        %>
                        </tr><tr>
                        <%
                        'end if%>
                    
                        <!--<td style="padding-right:5px; padding-top:10px; vertical-align:text-top;"><b><%=tsa_txt_068%>:</b></td>-->
                         <td style="padding-top:<%=paddingTop%>; padding-left:10px; width:225px">
                            <%if cint(mobil_week_reg_akt_dd) = 1 then %>
                                 <input type="hidden" id="FM_akt_0" value="-1"/>
                                 <!--<textarea id="dv_akt_test"></textarea>-->
                                 <select id="dv_akt_0" name="FM_aktivitetid" class="form-control input-<%=txtClass %> chbox_akt" style="font-size:<%=inputFont%>; height:<%=inputHeight%>" DISABLED>
                                      <option>..</option>
                                  </select>

                             <%else %>
                              <input style="font-size:<%=inputFont%>; height:<%=inputHeight%>" type="text" id="FM_akt_0" name="activity" placeholder="<%=tsa_txt_068%>" class="FM_akt form-control input-<%=txtClass %>"/>
                                <!--<div id="dv_akt_0" class="dv-closed dv_akt" style="border:1px #cccccc solid; padding:10px; visibility:hidden; display:none;"></div>--> <!-- dv_akt -->

                                  <select id="dv_akt_0" class="form-control input-small chbox_akt" size="10" style="visibility:hidden; display:none;"> <!-- chbox_akt -->
                                      <option><%=tsa_txt_534%>..</option>
                                  </select>
                              <input type="hidden" id="FM_aktid_0" name="FM_aktivitetid" value="0"/>
                             <%end if %>
                           
                        </td>
                       
                    </tr>

                    <tr>
                        <!--<td style="padding-right:5px; padding-top:10px; vertical-align:text-top;"><b><%=tsa_txt_137%>:</b></td>-->
                         <td style="padding-top:<%=paddingTop%>; padding-left:10px">
                             <input type="hidden" id="FM_sttid" name="FM_sttid" value="00:00"/>
                             <input type="hidden" id="FM_sltid" name="FM_sltid" value="00:00"/>
                             <input type="text" id="FM_timer" name="FM_timer" placeholder="<%=tsa_txt_137%>" class="form-control input-<%=txtClass %>" style="font-size:<%=inputFont%>; height:<%=inputHeight%>" /><!-- brug type number for numerisk tastatur -->
                        </td>
                        
                    </tr>

                    <tr>
                        <!--<td style="padding-right:5px; padding-top:10px; vertical-align:text-top;"><b><%=tsa_txt_051%>:</b></td>-->
                        <td style="padding-top:<%=paddingTop%>; padding-left:10px; text-align:right;">
                            <input type="text" id="FM_kom" name="FM_kom_0" placeholder="<%=tsa_txt_051%>" class="form-control input-<%=txtClass %>" style="font-size:<%=inputFont%>; height:<%=inputHeight%>" /><br />
                            <%if cint(touchMonitor) <> 1 then 
                                btnSz = "sm"
                            else
                                btnSz = "lg"
                            end if%>
                            <button class="btn btn-<%=btnSz %> btn-success"><b><%=tsa_txt_085%> >></b></button>
                            <%'end if %>
                        </td>
                       
                    </tr>

                    <%'if cint(touchMonitor) = 1 then %>
                    <!--<tr>
                         <td style="padding-top:30px; padding-left:10px; text-align:right" colspan="2"><button class="btn btn-lg btn-success"><b><%=tsa_txt_085%> >></b></button></td>                         
                    </tr>
                        -->
                  
      
                </table>
                </form>
                         
                     <%'end if %>

          </div>
           
            <div class="col-lg-5"> 
            <%
            '**** Norm Chart ****'    
            if cint(touchMonitor) <> 1 then %>
            <div id="stacked-vertical-chart" class="chart-holder-200"></div>
            <%else %>
                <!--#include file="../inc/regular/cls_kpi.asp"-->
                <%call keyfigures(touchMonitor, login) %>
        <%end if %>
                </div>
     
        </div><!-- ROW --> 
        <%end if %>
        
                <%if cint(touchMonitor) = 1 then %>
                <div class="row">
                    <div class="col-lg-5"> 
                        <a href="monitor.asp?func=startside" class="btn btn-lg btn-default"><< <b>Tilbage / Ny bruger</b></a>   
                        </div>
            </div><!-- ROW --> 
                <%end if %>

        </div>
        


        <%end if  %>


        <%
            
            if browstype_client = "ip" then
            %>
            <div class="row">
            <div class="col-lg-12"><div id="stacked-vertical-chart" class="chart-holder-200"></div></div>
            </div>
            <%end if %>

            
        <%
            end if

        'end if
    

   if touchMonitor <> 1 then
       ugeseddelvisning = 1

       perInterval = 6 'dateDiff("d", varTjDatoUS_man, varTjDatoUS_son, 2,2) 
       perIntervalLoop = perInterval

       'response.write "perIntervalLoop " & perIntervalLoop
       for l = 0 to perIntervalLoop 
        
            if l = 0 then
            varTjDatoUS_use = varTjDatoUS_man
            showheader = 1
            showtotal = 0
            else
            varTjDatoUS_use = dateAdd("d", l, varTjDatoUS_man)
            showheader = 0
            showtotal = 0
             end if 

           if l = perIntervalLoop then
           showtotal = 1
           end if

             varTjDatoUS_son = varTjDatoUS_man

            call ugeseddel(usemrn, varTjDatoUS_use, varTjDatoUS_use, ugeseddelvisning, showtotal, showheader)  
   
       next
   end if
	
	
        select case lto
        case "tec", "esn"
        aty_sql_realhours = " tfaktim <> 0"
        case "xdencker", "xintranet - local"
        aty_sql_realhours = " tfaktim = 1"
        case else
        aty_sql_realhours = aty_sql_realhours &""_
		& " OR tfaktim = 30 OR tfaktim = 31 OR tfaktim = 7 OR tfaktim = 11"
        end select 
   
                                

            'response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>aty_sql_realhours "& aty_sql_realhours
            'response.flush
            timerdenneuge_dothis = 0
            call timerDenneUge(usemrn, lto, varTjDatoUS_man, aty_sql_realhours, timerdenneuge_dothis, SmiWeekOrMonth)

                ugetotal = manTimer + tirTimer + onsTimer + torTimer + freTimer + lorTimer + sonTimer
                manTimer = replace(manTimer, ",",".")
                tirTimer = replace(tirTimer, ",",".")
                onsTimer = replace(onsTimer, ",",".")
                torTimer = replace(torTimer, ",",".")
                lorTimer = replace(lorTimer, ",",".")
                sonTimer = replace(sonTimer, ",",".")

                '*** NORM eller Komme/Gå som grundlag
                select case lto
                case "dencker", "intranet - local"

                    call fLonTimerPer(varTjDatoUS_man, 6, 21, usemrn)
                    ntimMan = ((manMin-manMinPause)/60)
                    ntimTir = ((tirMin-tirMinPause)/60)
                    ntimOns = ((onsMin-onsMinPause)/60)
                    ntimTor = ((torMin-torMinPause)/60)
                    ntimFre = ((freMin-freMinPause)/60)
                    ntimLor = ((lorMin-lorMinPause)/60)
                    ntimSon = ((sonMin-sonMinPause)/60)

                case else
                    call normtimerPer(usemrn, varTjDatoUS_man, 6, 0)
               
                end select

             norm_ugetotal = ntimMan + ntimTir + ntimOns + ntimTor + ntimFre + ntimLor + ntimSon

            'weeknormpro = ugetotal / norm_ugetotal * 100
            if cdbl(norm_ugetotal) <> 0 then
                
           
                    weeknormpro = ugetotal / norm_ugetotal * 100
                    weeknormpro = formatnumber(weeknormpro, 0)
                    weeknormpro = replace(weeknormpro, ",",".")
                

                 if ugetotal <= norm_ugetotal then
                 weeknormproWdt = weeknormpro
                 else
                 weeknormproWdt = norm_ugetotal / norm_ugetotal * 100 
                 end if

            else
                weeknormpro = 0
                norm_ugetotal = 0
            end if

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

        
          <%if touchMonitor <> 1 then %>
          <div class="row">
                <div class="col-lg-12">
                    <h4 style="text-align:right"><%=tsa_txt_173 &" "& tsa_txt_137%>: <%=formatnumber(norm_ugetotal, 2) %> &nbsp <%=tsa_txt_535%>: <%=ugetotal %> &nbsp&nbsp</h4>
                    <table class="table table">
                        <tr>
                            <th colspan="11"><div style="background-color:lightgray; height:30px; width: 100%;"><div style="background-color:#207800; height:30px; width:<%=weeknormproWdt %>%; color:#ffffff; text-align:right; padding:5px;"><%=weeknormpro %> %</div></div></th>
                        </tr>
                        <tr>
                            <th style="border-top:none; width:10%;"><h6>00</h6></th>
                            <th style="border-top:none; width:10%;"><h6></h6></th>
                            <th style="border-top:none; width:10%;"><h6>20</h6></th>
                            <th style="border-top:none; width:10%;"><h6></h6></th>
                            <th style="border-top:none; width:10%;"><h6>40</h6></th>
                            <th style="border-top:none; width:10%;"><h6></h6></th>
                            <th style="border-top:none; width:10%;"><h6>60</h6></th>
                            <th style="border-top:none; width:10%;"><h6></h6></th>
                            <th style="border-top:none; width:10%;"><h6>80</h6></th>
                            <th style="border-top:none; width:10%;"><h6></h6></th>
                            <th style="border-top:none; width:10%;"><h6>100</h6></th>
                                        
                        </tr>
                       <!-- <tr>
                            <th style="text-align:left; border-top:none; font-size:85%" colspan="11">Norm timer: <%=norm_ugetotal %></th>
                        </tr>
                        <tr>
                            <th style="text-align:left; border-top:none; font-size:85%" colspan="11">Opgjort: <%=weeknormpro %>%</th>
                        </tr> -->
                    </table>
                </div>
            </div>
            <%end if %>






            <%if media <> "print" then %>

                <!--
                <form action="ugeseddel_2011.asp?media=export" method="post" target="_blank">
                    
                 <div class="row">
                     <div class="col-lg-12 pad-r30">
                         <input id="Submit6" type="submit" value="Eksport til .csv simpel" class="btn btn-sm" />
                      </div>


                </div>
                    </form>
        -->

          
            <br /><br />


         <%if browstype_client = "ip" then %>
           
            <br /><br /><br /><br />
            <form action="../timetag_web/timetag_web.asp">
                <div class="row">
                    <div class="col-lg-12">
                       <!-- <a href="../timetag_web/timetag_web.asp?flushsessionuser=1" class="btn btn-sm btn-default" style="text-align:center; width:100%"><b>Tilbage</b></a> -->
                        <button class="btn btn-sm btn-default" style="text-align:center; width:50%"><b><< <%=tsa_txt_437 %></b></button>
                    </div>
                </div>
            </form>
           
        <%end if %>

             
        <%
            
            if (browstype_client <> "ip") AND (touchMonitor <> 1) then
        %>
    <div class="well" style="width:35%;">
      <div class="portlet">
        <h3 class="portlet-title">
          <u><%=week_txt_003 %>:</u><!-- ugeseddel -->
        </h3>

          

           

              <form action="ugeseddel_2011.asp?media=print&usemrn=<%=usemrn %>&varTjDatoUS_man=<%=varTjDatoUS_man%>" method="post" target="_blank">
                    
                 <div class="row">
                     <div class="col-lg-12 pad-r30">
                         <input id="Submit6" type="submit" value="<%=tsa_txt_536 %> >>" class="btn btn-sm btn-default" />
                      </div>


                </div>
                    </form>


         

        


          
<table>

    <tr><td colspan="2">

           <!-- hvis level = 1 OR teamleder -->
        <%if (level <=2 OR level = 6) AND cint(smilaktiv) <> 0 AND SmiWeekOrMonth <> 2 then %>
          

        <%
        '** Tjekker for Uge 53. SKAL i virkeligheden være om søndag er i et andet år end mandag - da år så skifter.
        if datepart("ww", varTjDatoUS_son, 2,2) = 53 then
        tjkAar = year(varTjDatoUS_son) + 1
        else
        tjkAar = year(varTjDatoUS_son)
        end if    
            
        '**
        call erugeAfslutte(tjkAar, datepart("ww", varTjDatoUS_son, 2,2), usemrn, SmiWeekOrMonth, 0) 
        
        call godkendugeseddel(fmlink, usemrn, varTjDatoUS_man, rdir) %>
    
        <%end if 'level %>
    

             </td>
        </tr>



</table>
  </div> <!-- well -->
  </div> <!-- portlet title -->
            <%end if %>
      

<%else

    
Response.Write("<script language=""JavaScript"">window.print();</script>")
    
end if 

      %>


    </div>

    <%
       '** Wrapper / Content
        if browstype_client <> "ip" then %>
      </div></div>
     <%end if %>

     
   
    <br /><br />&nbsp;
    <%
	
	
	end select
	
	
	end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
