
<%

    public strSQLDatoStatuskri
    function datosparrings_sglkri(lto, varTjDatoUS_man, varTjDatoUS_son)

                    call lukaktvdato_fn()
                    ignJobogAktper = lukaktvdato

                    select case lto
                    case "mpt"
                    jobstatusExtra = " OR (j.jobstatus = 2) OR (j.jobstatus = 4)"
                    case else
                    jobstatusExtra = ""
                    end select


                    select case ignJobogAktper
                    case 0,1
                    strSQLDatoStatuskri = " AND ((j.jobstartdato <= '"& varTjDatoUS_son &"' AND j.jobstatus = 1) OR (j.jobstatus = 3) "& jobstatusExtra &")"
                    case 3
                    strSQLDatoStatuskri = " AND ((j.jobstartdato <= '"& varTjDatoUS_son &"' AND j.jobslutdato >= '"& varTjDatoUS_man &"' AND j.jobstatus = 1) OR (j.jobstatus = 3) "& jobstatusExtra &")"
                    case else
                    strSQLDatoStatuskri = " AND ((j.jobstatus = 1) OR (j.jobstatus = 3) "& jobstatusExtra &")"
                    end select

                    'strSQLDatoStatuskri = " AND ((j.jobstatus = 1) OR (j.jobstatus = 3) "& jobstatusExtra &")"

    end function


    public strmedarbjobans 
    function GetMedarbJobAns(medid)
        
        if level <> 1 then     
            if level = 2 OR level = 6 then
                strmedarbjobans = ""
                strSQL = "SELECT id FROM job WHERE (jobans1 = "& medid &" OR jobans2 = "& medid &" OR jobans3 = "& medid &" OR jobans4 = "& medid &" OR jobans5 = "& medid &")"
                oRec.open strSQL, oConn, 3
                j = 0
                while not oRec.EOF

                    if j <> 0 then
                        strmedarbjobans = strmedarbjobans & "," & oRec("id")
                    else
                        strmedarbjobans = oRec("id")
                    end if

                j = j + 1
                oRec.movenext
                wend
                oRec.close

            else
            strmedarbjobans = "ingen"
            end if
        else
            strmedarbjobans = "alle"
        end if

    end function


    public positiv_aktivering_akt_val_SQL
    function positivakt_forecast_sqlkri(medid)

                    call positiv_aktivering_akt_fn()


                    positiv_aktivering_akt_val = positiv_aktivering_akt_val

                    if cint(positiv_aktivering_akt_val) = 1 OR cint(pa_aktlist) = 1 then 'Positiv aktivering / Må kun søge i aktiv jobliste
                    positiv_aktivering_akt_val_SQL = " AND ((tu.forvalgt = 1 "
            
                            if cint(positiv_aktivering_akt_val) = 1 then
                            positiv_aktivering_akt_val_SQL = positiv_aktivering_akt_val_SQL & " AND tu.aktid <> 0 AND risiko >= 0) OR risiko < 0"


                                            'Ved Positiv aktivering + PA = 1 søg i personlig jobliste, skal de aktiviteter med forecast på altid være tilgengelige
                                            call aktBudgettjkOn_fn()
                                            '*************************************************************************************************
                                            '*** Henter kun akt med ressource forecase Oko, WWF timer for valgte medarb    *******************
                                            '*************************************************************************************************
                                            if cint(aktBudgettjkOnViskunmbgt) = 1 then
                                                
                                               
                                               
                                                datoKrionly = 1
                                                call ressourcetimerTildelt(varTjDatoUS_man, 0, 0, medid, datoKrionly)
                                                
                                                strSQLaktIdFc = "SELECT jobid FROM ressourcer_md WHERE jobid <> 0 AND medid = "& medid &" "& sqlBudgafg &" GROUP BY jobid"
                                                
                                                         'Response.Write "<option>OK - "& strSQLaktIdFc &"</option>"
                                                         'Response.end

                                                jobidsMtimSQL = ""
                                                oRec6.open strSQLaktIdFc, oConn, 3 
                                                While not oRec6.EOF

                                                jobidsMtimSQL = jobidsMtimSQL & " OR j.id = " & oRec6("jobid")
                                            
                                                oRec6.movenext
                                                wend
                                                oRec6.close

                                               
                                            
                                                positiv_aktivering_akt_val_SQL = positiv_aktivering_akt_val_SQL & jobidsMtimSQL

                                          

                                            end if
                                            '*************************************************************************************************
                                    

                                    positiv_aktivering_akt_val_SQL = positiv_aktivering_akt_val_SQL & ")"

                            else
                            positiv_aktivering_akt_val_SQL = positiv_aktivering_akt_val_SQL & "))"
                            end if

                    

                    else
                    positiv_aktivering_akt_val_SQL = ""
                    end if
 

                    
                  


    end function

    
    function selectJobogKunde_jq()
                
                '*** SØG kunde & Job 

                 medid = request("jq_medid")

                 if len(trim(request("jq_jobidc"))) <> 0 then
                 jq_jobidc = request("jq_jobidc")
                 else
                 jq_jobidc = "-1"
                 end if
    
                 

                 lto = request("lto") 

              
    
               
                if len(trim(request("jq_newfilterval"))) <> 0 then
                filterVal = 1 
                jobkundesog = request("jq_newfilterval")
                
                if len(trim(request("searchmode"))) <> 0 then 'Searchmode bruges af Monitor Cflow Oliver

                    searchmode = request("searchmode")
                    sogSort = request("sogSort")

                    select case cint(sogSort)
                    case 1
                    strsogSort = "ORDER BY jobnr, jobnavn"
                    case 2
                    strsogSort = "ORDER BY jobnr DESC, jobnavn"
                   end select          
    
                else

                    searchmode = 0

                end if

                'if session("mid") = 1 then
                '*** ÆØÅ **'
                'call jq_format(jobkundesog)
                'jobkundesog = jq_formatTxt
                'end if
                

                else
                filterVal = 0
                jobkundesog = "6xxxxxfsdf554"
                end if
        
                


                if filterVal <> 0 then


                    varTjDatoUS_man = request("varTjDatoUS_man")
                    varTjDatoUS_son = dateAdd("d", 6, varTjDatoUS_man)

                    varTjDatoUS_man = year(varTjDatoUS_man) &"/"& month(varTjDatoUS_man) &"/"& day(varTjDatoUS_man)
                    varTjDatoUS_son = year(varTjDatoUS_son) &"/"& month(varTjDatoUS_son) &"/"& day(varTjDatoUS_son)                  
                
    
                    '***** POSITIV aktivitering ***'
                    call positivakt_forecast_sqlkri(medid)
                    strSQLPAkri = positiv_aktivering_akt_val_SQL
                    'strSQLPAkri = ""
          
               
                  '*** Datospærring Vis først job når stdato er oprindet
                  call datosparrings_sglkri(lto, varTjDatoUS_man, varTjDatoUS_son)

                   
             

                 if jobkundesog <> "-1" then
                 strSQLSogKri = " AND (jobnr LIKE '"& jobkundesog &"%' OR jobnavn LIKE '%"& jobkundesog &"%' OR "_
                 &" kkundenavn LIKE '"& jobkundesog &"%' OR kkundenr = '"& jobkundesog &"' OR k.kinit = '"& jobkundesog &"') AND kkundenavn <> ''"
                 lmt = 50
                 else
                 strSQLSogKri = ""
                 lmt = 500 '250
                 end if
            
                 lastKid = 0
                

               
                strSQL = "SELECT j.id AS jid, j.jobnavn, j.jobnr, j.jobstatus, k.kkundenavn, k.kkundenr, k.kid FROM timereg_usejob AS tu "_ 
                &" LEFT JOIN job AS j ON (j.id = tu.jobid) "_
                &" LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "_
                &" WHERE tu.medarb = "& medid &" "& strSQLPAkri &" "& strSQLDatoStatuskri 
        
                'AND ("& jobstatusSQL &")
                
                strSQL = strSQL & strSQLSogKri

                if lto = "cflow" then
                    if cint(searchmode = 1) then
                    strSQL = strSQL &" GROUP BY j.id "& strsogSort &" LIMIT "& lmt       
                    else
                    strSQL = strSQL &" GROUP BY j.id ORDER BY jobnr, jobnavn LIMIT "& lmt
                    end if
                else
                strSQL = strSQL &" GROUP BY j.id ORDER BY kkundenavn, jobnavn LIMIT "& lmt
                end if

                'if session("mid") = 1 OR session("mid") = 35 then
                'response.write "strSQL " & strSQL & ""
                'response.flush
                'end if
    
                select case lto 
                case "xcflow", "xintranet - local"
    

                case else

                    if (jobkundesog = "-1") then
                        
                        if len(trim(week_txt_009)) = 0 then
                        week_txt_009 = "Select Job"
                        end if
                        if cint(searchmode) <> 1 then
                        strJobogKunderTxt = strJobogKunderTxt & "<option value=""-1"" SELECTED>"& week_txt_009 &":</option>"
                        else
                        strJobogKunderTxt = strJobogKunderTxt & ""
                        end if
                end if            

                end select


                dim jobtd
                redim jobtd(lmt)

                antalA = 0
                antalB = 0
                antalC = 0
                antalD = 0
                antalE = 0
                antalF = 0
                antalG = 0
                antalH = 0
                antalI = 0
                antalJ = 0
                antalK = 0
                antalL = 0
                antalM = 0
                antalN = 0
                antalO = 0
                antalP = 0
                antalQ = 0
                antalR = 0
                antalS = 0
                antalT = 0
                antalU = 0
                antalV = 0
                antalX = 0
                antalY = 0
                antalZ = 0
                antalAE = 0
                antalOO = 0
                antalAA = 0
                antalW = 0

                alabetColor = ""
                alabetDisplay = ""

                j = 0
                k = 0
                oRec.open strSQL, oConn, 3
                while not oRec.EOF
        
    
                'if lto = "cflow" then
                'strtdtest = strtdtest & "<td>j: "& j & "</td>"
                'end if

                'if session("mid") = 1 then
                '   strJobogKunderTxt = strJobogKunderTxt & "<option>"& strSQL &"</option>"
                'response.end
                'end if


                if lastKnavn <> oRec("kkundenavn") then
    
                    if cint(searchmode) <> 1 then
                        if k <> 0 then
                            ' strJobogKunderTxt = strJobogKunderTxt &"<br>"
                            strJobogKunderTxt = strJobogKunderTxt & "<option DISABLED></option>"
                        end if
                    end if
    
                'strJobogKunderTxt = strJobogKunderTxt & oRec("kkundenavn") &" "& oRec("kkundenr") &"<br>"

                   if cint(searchmode) <> 1 then
                        strJobogKunderTxt = strJobogKunderTxt & "<option DISABLED>"& oRec("kkundenavn") &" "& oRec("kkundenr") &"</option>"
                    end if
                end if 
                 
               ' strJobogKunderTxt = strJobogKunderTxt & "<input type=""hidden"" id=""hiddn_job_"& oRec("jid") &""" value="""& oRec("jobnavn") & " ("& oRec("jobnr") &")"">"
               ' strJobogKunderTxt = strJobogKunderTxt & "<a class=""chbox_job"" id=""chbox_job_"& oRec("jid") &""" value="& oRec("jid") &">"& oRec("jobnavn") & " ("& oRec("jobnr") &")" &"</a><br>" 
                
                if cdbl(jq_jobidc) = cdbl(oRec("jid")) then
                 jobidSEL = "SELECTED"
                else
                 jobidSEL = ""
                end if

                select case lto
                case "intranet - local", "cflow"
                    if cint(searchmode) <> 1 then

    
                        strJobogKunderTxt = strJobogKunderTxt & "<option value="& oRec("jid") &" "& jobidSEL &">"& oRec("jobnr") &" "& oRec("jobnavn") &"</option>"
                    
    
                     else

                        strjobnavn = oRec("jobnavn")
                        if len(strjobnavn) > 14 then
                        strjobnavn = left(strjobnavn, 14) & "..."
                        else
                        strjobnavn = strjobnavn
                        end if

                        'strjobnavnFirstLetter = left(strjobnavn,1)
                        'strjobnavn = right(strjobnavn, (len(strjobnavn) - 1))

                        strjobnFirstLetter = left(oRec("jobnr"),1)
                        
                        select case strjobnFirstLetter
                            case "1"
                                if antalA < 1 then
                                    alabetColor = "#db2b2b"
                                    alabetDisplay = ""
                                    antalA = antalA + 1
                                else
                                    alabetColor = ""
                                    alabetDisplay = "hidden"
                                end if
                            case "2"
                                if antalB < 1 then
                                    alabetColor = "#db2b2b"
                                    alabetDisplay = ""
                                    antalB = antalB + 1
                                else
                                    alabetColor = ""
                                    alabetDisplay = "hidden"
                                end if
                            case "3"
                                if antalC < 1 then
                                    alabetColor = "#db2b2b"
                                    alabetDisplay = ""
                                    antalC = antalC + 1
                                else
                                    alabetColor = ""
                                    alabetDisplay = "hidden"
                                end if
                            case "4"
                                if antalD < 1 then
                                    alabetColor = "#db2b2b"
                                    alabetDisplay = ""
                                    antalD = antalD + 1
                                else
                                    alabetColor = ""
                                    alabetDisplay = "hidden"
                                end if
                            case "5"
                                if antalE < 1 then
                                    alabetColor = "#db2b2b"
                                    alabetDisplay = ""
                                    antalE = antalE + 1
                                else
                                    alabetColor = ""
                                    alabetDisplay = "hidden"
                                end if
                            case "6"
                                if antalF < 1 then
                                    alabetColor = "#db2b2b"
                                    alabetDisplay = ""
                                    antalF = antalF + 1
                                else
                                    alabetColor = ""
                                    alabetDisplay = "hidden"
                                end if
                            case "7"
                                if antalG < 1 then
                                    alabetColor = "#db2b2b"
                                    alabetDisplay = ""
                                    antalG = antalG + 1
                                else
                                    alabetColor = ""
                                    alabetDisplay = "hidden"
                                end if
                            case "8"
                                if antalH < 1 then
                                    alabetColor = "#db2b2b"
                                    alabetDisplay = ""
                                    antalH = antalH + 1
                                else
                                    alabetColor = ""
                                    alabetDisplay = "hidden"
                                end if
                            case "9"
                                if antalI < 1 then
                                    alabetColor = "#db2b2b"
                                    alabetDisplay = ""
                                    antalI = antalI + 1
                                else
                                    alabetColor = ""
                                    alabetDisplay = "hidden"
                                end if                            
                            case else
                                alabetColor = ""
                                alabetDisplay = "hidden"
                        end select

                          'strtd = strtd & "<td>j: "& j & " "& oRec("jobnavn") & " ("& oRec("jobnr") &") </td>"


                        if len(oRec("jobnavn")) > 40 then
                        jobnavn_nr = left(oRec("jobnavn"), 40) & ".. ("& oRec("jobnr") &")"
                        else
                        jobnavn_nr = oRec("jobnavn") &" ("& oRec("jobnr") &")"
                        end if

                        'jobtd(j)
                        'thisjobtd = ""
                        jobtd(j) = "<td> <div style='display:inline-block;'> <span style='margin-left:15px; font-size:110%; width:35px; visibility:"&alabetDisplay&"; background-color:#db2b2b;' class='btn btn-default btn-sm'><b style='color:white'>"&strjobnFirstLetter&"<b></span> <br> <span style='width:200px; margin-top:1px; margin-left:15px; font-size:200%; background-color:#f9f9f9;' class='btn btn-default jobsel' id="& oRec("jid") &"><b><span>"&oRec("jobnr")&"</span> <br> <span style='font-size:17px;'>"& strjobnavn &"</span>" &" <input type='hidden' id='jobnavn_sel_"&oRec("jid")&"' value='"& jobnavn_nr &"' /></b></span> </div> </td>"
                            
                            'if j = 0 then
                            '    strtd = strtd & "<tr>"
                            'end if
                            
                            'strtd = strtd & thisjobtd 

                            'if right(j, 1) = 5 then
                            '    strtd = strtd & "</tr>"
                            'end if

                            'totaljobs = totaljobs + 1
                

                    end if
    
                case else
                strJobogKunderTxt = strJobogKunderTxt & "<option value="& oRec("jid") &" "& jobidSEL &">"& oRec("jobnavn") & " ("& oRec("jobnr") &")" &"</option>"
                end select

                j = j + 1
                k = k + 1
                lastKnavn = oRec("kkundenavn") 
                oRec.movenext
                wend
                oRec.close

                if cint(searchmode) <> 1 then
                    if cint(k) = 0 then
                        strJobogKunderTxt = strJobogKunderTxt & "<option value=""-1"" DISABLED>"& week_txt_010 &"</option>"
                    end if
                end if

                if cint(searchmode) = 1 then


                    strtd = ""
                    row = 1
                    totaljobs = 0
                    for i = 0 TO UBOUND(jobtd)
    
                        if len(jobtd(i)) <> 0 then
                           if row = 1 then
                                strtd = strtd & "<tr>"
                           end if
                            
                            strtd = strtd & jobtd(i)

                            if row = 5 then
                                strtd = strtd & "</tr>"
                                row = 0
                            end if

                            row = row + 1
                            totaljobs = totaljobs + 1
                        end if

                        
                    next
        
                    if cint(totaljobs) < 5 then
                    'if cint(j) < 5 then
                        td = totaljobs
                        for td = totaljobs TO 4
                            strtd = strtd & "<td style='display:none;'>&nbsp;</td>"
                        next
                    end if
                    
                    'if session("mid") = 1 then
                    'strtd = strtd & "<td>"& strSQL &"</td>" & "<td>HER DER</td>" & strtdtest
                    'end if

                    strJobogKunderTxt = strJobogKunderTxt & strtd
                end if


                    '*** ÆØÅ **'
                    call jq_format(strJobogKunderTxt)
                    strJobogKunderTxt = jq_formatTxt


                    response.write strJobogKunderTxt

                end if    'filterval
    
    
    
    
    end function






    function lukkedato(jobId, jobStatus)

strLukkeDato = year(now) &"/"& month(now) &"/"& day(now)

 if jobStatus = 0 then 'Kun ved luk job
    strSQLjs = "SELECT jobstatus FROM job WHERE id = "& jobId
    oRec5.open strSQLjs, oConn, 3
    if not oRec5.EOF then
                                        
        if oRec5("jobstatus") <> 0 then 'job bliver lukket nu

        strSQLjlukkedt = "UPDATE job SET lukkedato = '"& strLukkeDato &"' WHERE id = " & jobId
        oConn.execute(strSQLjlukkedt)

        end if


    end if
    oRec5.close 
end if


end function




    '***** Sync Job slutDato ***********'
    public useSyncDato
    function syncJobSlutDato(jobid, jobnr, syncslutdato)

    fod = 0

     '***** slutdato på job følger altid sidste timereg. / sidste fakturadato ellers dagsdato ****'
    if syncslutdato = 1 then

    '** Sidste timeregdato **
    lastTregDato = "1-1-2002"
    strSQLsidstetreg = "SELECT tdato FROM timer WHERE tjobnr = '"& jobnr &"' ORDER BY tdato DESC"
    Response.write strSQLsidstetreg & "<br>"
    oRec5.open strSQLsidstetreg, oConn, 3
    if not oRec5.EOF then
                                
    lastTregDato = oRec5("tdato")

    fod = 1
    end if
    oRec5.close 

    '** Sidste fakdato **
    lastFakDato = "1-1-2002"
    strSQLsidstefak = "SELECT if(brugfakdatolabel = 1, labeldato, fakdato) AS lastfaktdato FROM fakturaer WHERE jobid = "& jobid &" AND shadowcopy <> 1 AND faktype = 0 ORDER BY fid DESC"
    oRec5.open strSQLsidstefak, oConn, 3
    if not oRec5.EOF then
                                
    lastFakDato = oRec5("lastfaktdato")

    fod = 1
    end if
    oRec5.close 


    'Response.write "fod " & fod & " lastTregDato: "& lastTregDato & " lastFakDato: "& lastFakDato &"<br>" 

    '*** Er der fundet en ****
    if fod = 1 then

            if cDate(lastTregDato) > cDate(lastFakDato) then
            useSyncDato = lastTregDato
            else
            useSyncDato = lastFakDato
            end if 

    else
    '*** else dd **'
    useSyncDato = year(now) & "/" & month(now) & "/" & day(now)

    end if

    useSyncDato = year(useSyncDato) & "/" & month(useSyncDato) & "/" & day(useSyncDato) 

    strSQLjobSyncDato = "UPDATE job SET jobslutdato = '"& useSyncDato &"' WHERE id = "& jobid    

    'Response.write strSQLjobSyncDato

    oConn.execute(strSQLjobSyncDato)

    'Response.end

    end if
    '****

    end function


    public jobstatus_fn_txt, jobstatus_fn_options, jobstatus_fn_txt_col
    function jobstatus_fn(jobid, jobstatus, io)

        jobstatus_fn_txt = ""
        jobstatus_fn_options = ""


         if len(trim(jobstatus)) <> 0 then
        jobstatus = jobstatus
        else
        jobstatus = -1
        end if

        if cint(io) = 1 then 'liste options
        sqlWh = " js_id <> 100 "
        else
        sqlWh = " js_id = "& jobstatus
        end if
         

        strSQLjobstatus = "SELECT js_id, js_navn FROM job_status WHERE "& sqlWh &" ORDER BY js_id"
        'response.write strSQLjobstatus
        'response.flush 
    
        oRec6.open strSQLjobstatus, oConn, 3
        while not oRec6.EOF 

            select case oRec6("js_id")
            case 1
		    jobstatus_fn_txt = job_txt_094 'Aktiv
            jobstatus_fn_txt_col  = "yellowgreen"
			case 2
			jobstatus_fn_txt = job_txt_095 'Passiv
            jobstatus_fn_txt_col  = "#CCCCCC"
			case 0
			jobstatus_fn_txt = job_txt_096 'Lukket
            jobstatus_fn_txt_col  = "red"
            case 3
			jobstatus_fn_txt = job_txt_097 'Tilbud
            jobstatus_fn_txt_col  = "#5582d2"
            case 4
			jobstatus_fn_txt = job_txt_098 'Gennemsyn
            jobstatus_fn_txt_col  = "yellow"
            case 5
			jobstatus_fn_txt = "Evaluering" 'Evaluering
            jobstatus_fn_txt_col  = "orange"
			end select

           
            if cint(io) = 1 then
            
            'if cint(jobstatus) = cint(oRec6("js_id")) then
            'statjobSEL = "SELECTED"
            'else
            'statjobSEL = ""
            'end if

            '"& statjobSEL &"

            jobstatus_fn_options = jobstatus_fn_options & "<option value="& oRec6("js_id") &">"& jobstatus_fn_txt &"</option>"

            end if
            	
    
          

        oRec6.movenext
        wend
        oRec6.close

    end function


     public totTerminBelobJob, totTerminBelobGrand
    function terminbelob(jobid, showhtml)

    

        totTerminBelobJob = 0
		strSQLm = "SELECT milepale.id AS id, milepale.navn AS navn, milepal_dato, "_
		&" milepale_typer.navn AS type, ikon, beskrivelse, milepale.editor, belob FROM milepale "_
		&" LEFT JOIN milepale_typer ON (milepale_typer.id = milepale.type) "_
		&" WHERE jid = "& jobid &" GROUP BY milepale.id ORDER BY milepal_dato"
		x = 0

        'response.write strSQLm
        'response.flush

		oRec4.open strSQLm, oConn, 3
		while not oRec4.EOF 
		
        if showhtml = 1 then

		select case right(x, 1)
		case 0,2,4,6,8
		bgthis = "#FFFFFF"
		case else
		bgthis = "#EFF3FF"
		end select

        if isNull(oRec4("milepal_dato")) <> true then
        mpDate = formatdatetime(oRec4("milepal_dato"), 2)
        else
        mpDate = "-"
        end if

		%>
		<tr bgcolor="<%=bgthis %>">
			<td valign="top" style="border-bottom:1px #CCCCCC solid; font-size:8px; white-space:nowrap;">
			<a href="javascript:popUp('milepale.asp?func=red&id=<%=oRec4("id")%>&jid=<%=jobid%>&rdir=wip&type=1','650','500','250','120');" target="_self" class=rmenu><%=left(oRec4("navn"),3)%></a> <%=mpDate%><br />
		    </td>
            <td valign="top" style="border-bottom:1px #CCCCCC solid;" align=right class=lille><%=formatnumber(oRec4("belob"),0) %></td>
              <td valign="top" style="border-bottom:1px #CCCCCC solid;" align=right class=lille> <a href="javascript:popUp('milepale.asp?func=slet&id=<%=oRec4("id")%>&jid=<%=oRec("id")%>&rdir=wip&type=1','650','500','250','120');" target="_self" class=red>x</a></td>
			
			
		</tr>
		
		<%
        end if

        totTerminBelobJob = totTerminBelobJob + oRec4("belob")
        totTerminBelobGrand = totTerminBelobGrand + oRec4("belob")

        Response.flush

		x = x + 1
		oRec4.movenext
		wend
		oRec4.close
		

    end function



 function lukjobmail(jstatus, lukjob, medarbejderid)
 'for j = 0 to UBOUND(lukjob)
            				
                            'if lukjob(j) <> 0 then
                            
                            if jstatus = 0 then
                            lukjobSQL = ", lukkedato = '"& year(now)&"/"& month(now)&"/"& day(now) &"'"
                            else
                            lukjobSQL = ""
                            end if

                            '*** Skifter status ****'
                            strSQLst = "UPDATE job SET jobstatus = "& jstatus &" "& lukjobSQL &" WHERE id = "& lukjob  'jobnr = "& intJobnr
                            'Response.write strSQLst 
                            'Response.end
		                    oConn.execute(strSQLst)

                            '** Ny jobstatus 
                            'select case jstatus
                            'case 0
                            'jstatusTxt = "Lukket"
                            'case 1
                            'jstatusTxt = "Aktiv"
                            'case 2
                            'jstatusTxt = "Passiv / Til Fakturering"
                            'case 3
                            'jstatusTxt = "Tilbud"
                            'case 4
                            'jstatusTxt = "Til gennemsyn"
                            'end select

                            call jobstatus_fn(lukjob, jstatus, 0)
                            jstatusTxt = jobstatus_fn_txt

				            '**** Finder jobansvarlige *****
				            strSQL = "SELECT job.id AS jid, jobnavn, jobnr, jobans1, jobans2, m1.mnavn AS m1mnavn, m1.email AS m1email,"_
				            &" m2.mnavn AS m2mnavn, m2.email AS m2email, kkundenavn, kkundenr FROM job "_
				            &" LEFT JOIN medarbejdere AS m1 ON (m1.mid = jobans1)"_
				            &" LEFT JOIN medarbejdere AS m2 ON (m2.mid = jobans2)"_
                            &" LEFT JOIN kunder AS k ON (k.kid = job.jobknr)"_
				            &" WHERE job.id = "& lukjob

                            'Response.Write strSQL
                            'Response.flush

				            oRec.open strSQL, oConn, 3
				            x = 0
				            if not oRec.EOF then
            				
				            jobid = oRec("jid")
				            jobans1 = oRec("m1mnavn")
				            jobans2 = oRec("m2mnavn")
				            jobans1email = oRec("m1email")
				            jobans2email = oRec("m2email")
				            jobnavnThis = oRec("jobnavn")
                            intJobnr = oRec("jobnr")
                            strkkundenavn = oRec("kkundenavn")
                            'kkundenr = oRec("kkundenr")
            				
				            x = x + 1
				            end if
				            oRec.close
            				
            				
				            '*** Henter afsender **
				            strSQL = "SELECT mnavn, email FROM medarbejdere"_
				            &" WHERE mid = "& medarbejderid
				            oRec.open strSQL, oConn, 3
            				
				            if not oRec.EOF then
            				
				            afsNavn = oRec("mnavn")
				            afsEmail = oRec("email")
            				
				            end if
				            oRec.close


                            select case lto
                            case "dencker", "intranet - local", "outz"

                       

                             '*** Henter timeforbrug fakturerbare **

                            strTime_forbrug = "Timeforbrug:<br>"
				            fakturerbaretimer = 0
                            strSQL = "SELECT SUM(timer) AS sumtimer FROM timer WHERE tfaktim = 1 AND tjobnr = '"& intJobnr &"' GROUP BY tjobnr, tfaktim"
				            oRec.open strSQL, oConn, 3
            				
				            if not oRec.EOF then
            				
                            fakturerbaretimer = oRec("sumtimer")
                            strTime_forbrug = strTime_forbrug & "Fakturerbare timer: " & formatnumber(fakturerbaretimer, 2) & "<br>"

                            end if
				            oRec.close

                             '*** Henter timeforbrug ikke fakturerbare **
				            ikkefakturerbaretimer = 0
                            strSQL = "SELECT SUM(timer) AS sumtimer FROM timer WHERE tfaktim = 2 AND tjobnr = '"& intJobnr &"' GROUP BY tjobnr, tfaktim"
				            oRec.open strSQL, oConn, 3
            				
				            if not oRec.EOF then
            				
                            ikkefakturerbaretimer = oRec("sumtimer")
                            strTime_forbrug = strTime_forbrug & "Ikke fakturerbare timer: " & formatnumber(ikkefakturerbaretimer, 2) & "<br>"

                            end if
				            oRec.close

                            
                           
            				

                            '*** Henter timeforbrug E1 ubemandet **
				            e1 = 0
                            strSQL = "SELECT SUM(timer) AS sumtimer FROM timer WHERE tfaktim = 90 AND tjobnr = '"& intJobnr &"' GROUP BY tjobnr, tfaktim"
				            oRec.open strSQL, oConn, 3
            				
				            if not oRec.EOF then
            				
                            e1 = oRec("sumtimer")
                            strTime_forbrug = strTime_forbrug & "Ubemandet: " & formatnumber(e1, 2) & "<br>"

                            end if
				            oRec.close


                            

                             '*** Henter materialeforbrug**
				            matforbrug = 0
                            matsalgspris = 0
                            matnavn = ""
                            strMat_forbrug = ""

                            strMat_forbrug = "<br><br>Materialeforbrug: <br>"
                            strSQL = "SELECT matnavn, SUM(matantal) as matantal, SUM(matsalgspris) AS salgspris FROM materiale_forbrug WHERE jobid = "& lukjob &" GROUP BY jobid"
				            oRec.open strSQL, oConn, 3
            				
				            if not oRec.EOF then
            				
                            matsalgspris = oRec("salgspris") 
                            matforbrug = oRec("matantal")
                            'matnavn = oRec("matnavn")

                            'strMat_forbrug = strMat_forbrug & matnavn & ": "& matforbrug &" stk., pris: "& formatnumber(matsalgspris,2) & "<br>"  
                            strMat_forbrug = strMat_forbrug & matforbrug &" stk., pris: "& formatnumber(matsalgspris,2) & "<br>"               

                            'oRec.movenext
                            end if
				            oRec.close

                            'Response.write strTime_forbrug
                            'Response.write strMat_forbrug & "<br><br>"



                            end select

                            'Sum af bemandet timer (fakturerbare)
                            'Sum af ikke fakturerbare
                            'Sum af ubemandet maskintid (E1)
                            'Sum af registreret materiale    
            					
            						
            					
				            if x <> 0 AND request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\timereg_akt_2006.asp" then
					  	            'Sender notifikations mail
		                            'Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		                            ' Sætter Charsettet til ISO-8859-1
		                            'Mailer.CharSet = 2

                                    Set myMail=CreateObject("CDO.Message")
                                    myMail.From="timeout_no_reply@outzource.dk" 'TimeOut Email Service 

		                            'Mailer.FromName = "TimeOut | " & afsNavn 
            
                                    'select case lto
                                    'case "hestia"
                                    'Mailer.FromAddress = "timeout_no_reply@outzource.dk"
                                    'case else
		                            'Mailer.FromAddress = afsEmail
                                    'end select
		                            
                                    'Mailer.RemoteHost = "webout.smtp.nu" '"webmail.abusiness.dk" '"pasmtp.tele.dk"
						            'Mailer.AddRecipient "" & jobans1 & "", "" & jobans1email & ""
		                            'Mailer.AddRecipient "" & jobans2 & "", "" & jobans2email & ""

                                    'myMail.To= ""& jobans1 &"<"& jobans1email &">"
                                    'myMail.Cc= ""& jobans2 &"<"& jobans2email &">"

                                    
                                    myMail.To= jobans1email
                                    myMail.CC= jobans2email 
            						
                                    if lto = "synergi1" then
                                    'Mailer.AddRecipient "Bettina Schou", "bs@synergi1.dk"
                                    'myMail.Bcc= "Bettina Schou<bs@synergi1.dk>"
                                    myMail.Bcc= "Louise Frandsen<lf@synergi1.dk>"
                                    'Mailer.AddRecipient "TimeOut Support", "support@outzource.dk"
                                    end if

                                    if lto = "outz" then
                                    'Mailer.AddRecipient "TimeOut Support", "support@outzource.dk"
                                    myMail.Bcc= "TimeOut Support<support@outzource.dk>"
                                    end if

                                    if lto = "dencker" then
                                    'Mailer.AddRecipient "Anders Dencker", "ad@dencker.net"
                                    'myMail.Bcc= "Anders Dencker<ad@dencker.net>"
                                     'myMail.Bcc= "TimeOut Support<support@outzource.dk>"
                                     myMail.Bcc= "Dencker Ordrer<ordre@dencker.net>"
            
                                    end if

                                    'if lto = "fe" then
                                    'Mailer.AddRecipient "Lone P", "lrp@feteknik.dk"
                                    'end if

                                    if lto = "essens" then
                                    'Mailer.AddRecipient "Astrid", "astrid@essens.info"
                                    myMail.Bcc= "Bolette Westh<bolette@essens.info>"
                                    end if

            						    
						            'Mailer.Subject = strkkundenavn &", "& jobnavnThis &" ("& intJobnr &") - Afsluttet."
                                    myMail.Subject= strkkundenavn &", "& jobnavnThis &" ("& intJobnr &") - Afsluttet af "& session("user")
		                           
                                    strBody = "Hej jobansvarlige.<br><br>"
                                    strBody = strBody &"Mit arbejde er afsluttet / Jobbet har skiftet status.<br><br>"
                                    strBody = strBody &"Kunde: "& strkkundenavn &"<br>"
						            strBody = strBody &"Job: "& jobnavnThis &" ("& intJobnr &")<br><br> "
                                    strBody = strBody &"Jobstatus er: "& jstatusTxt &"<br><br>"

                                    select case lto 
                                    case "dencker", "outz", "intranet - local"

                                    strBody = strBody & "<hr>"& strTime_forbrug
		                            strBody = strBody & strMat_forbrug & "<hr><br><br>"

                                    end select

		                            strBody = strBody &"Med venlig hilsen<br>"
		                            strBody = strBody & session("user") & "<br><br>"
            		                
            		
                                    myMail.HTMLBody= "<html><head></head><body>" & strBody & "</body>"

                                    myMail.Configuration.Fields.Item _
                                    ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
                                    'Name or IP of remote SMTP server
                                    
                                                if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
                                                   smtpServer = "webout.smtp.nu"
                                                else
                                                   smtpServer = "formrelay.rackhosting.com" 
                                                end if
                    
                                                myMail.Configuration.Fields.Item _
                                                ("http://schemas.microsoft.com/cdo/configuration/smtpserver")= smtpServer

                                    'Server port
                                    myMail.Configuration.Fields.Item _
                                    ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25
                                    myMail.Configuration.Fields.Update

                                    if len(trim(jobans1email)) <> 0 then
                                    myMail.Send
                                    end if

                                    set myMail=nothing



            		
		                            'Mailer.BodyText = strBody
            		
		                            'Mailer.sendmail()
		                            'Set Mailer = Nothing

                                    Response.Write "mail afsendt - " & request.form("jq_lukjob")
				            
                            end if' x



    end function

 %>