<%
response.buffer = true
'timeA = now
level = session("rettigheder")
    %>
        <!--#include file="../inc/connection/conn_db_inc.asp"-->
        <!--#include file="../inc/errors/error_inc.asp"-->
        <!--#include file="../inc/regular/global_func.asp"-->
        <!--#include file="../inc/regular/stat_func.asp"-->
        <%
        if Request.Form("AjaxUpdateField") = "true" then
            Select Case Request.Form("control")

                case "HentMedarbejdere"

                    jobid = request("jobid")
                    aktid = request("aktid")
                    startmd = request("startMD")
                    startyy = request("startYY")
                    slutmd = request("slutMD")
                    slutyy = request("slutYY")

                    'response.Write startmd & " " & startyy & " til " & slutmd & " " & slutYY
                    
                    visalle = 0

                    strSQL = "SELECT projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM aktiviteter WHERE id = "& aktid
                   ' response.Write strSQL
                    oRec.open strSQL, oConn, 3
                    if not oRec.EOF then
                        if cint(oRec("projektgruppe1")) = 10 OR cint(oRec("projektgruppe2")) = 10 OR cint(oRec("projektgruppe3")) = 10 OR cint(oRec("projektgruppe4")) = 10 OR cint(oRec("projektgruppe5")) = 10 OR cint(oRec("projektgruppe6")) = 10 OR cint(oRec("projektgruppe7")) = 10 OR cint(oRec("projektgruppe8")) = 10 OR cint(oRec("projektgruppe9")) = 10 OR cint(oRec("projektgruppe10")) = 10 then
                            visalle = 1
                        end if

                       prgids = oRec("projektgruppe1") &", "_
                        & oRec("projektgruppe2") &", "_
                        & oRec("projektgruppe3") &", "_
                        & oRec("projektgruppe4") &", "_
                        & oRec("projektgruppe5") &", "_
                        & oRec("projektgruppe6") &", "_
                        & oRec("projektgruppe7") &", "_
                        & oRec("projektgruppe8") &", "_
                        & oRec("projektgruppe9") &", "_
                        & oRec("projektgruppe10")

                    end if
                    oRec.close 
                                                           
                    strSQL = "SELECT jobans1, jobans2, jobans3, jobans4, jobans5 FROM job WHERE id = "& jobid
                    oRec.open strSQL, oConn, 3
                    if not oRec.EOF then

                        jobansids = oRec("jobans1") & ", "_
                            & oRec("jobans2") & ", "_
                            & oRec("jobans3") & ", "_
                            & oRec("jobans4") & ", "_
                            & oRec("jobans5")

                    end if
                    oRec.close

                    erjobans = 0
                    jobansidsSplit = split(jobansids, ", ")
                    for a = 0 TO UBOUND(jobansidsSplit)
                        if cdbl(jobansidsSplit(a)) = cdbl(session("mid")) then
                            erjobans = 1
                        end if
                    next

                    if visalle = 1 then
                        sqlWHERE = "mansat = 1"
                    else
                        
                        if erjobans = 0 then
                            call fTeamleder(session("mid"), 0, 1)
                            if strPrgids <> "ingen" AND strPrgids <> "alle" then
                                prgsidsToShow = ""
                                
                                prgidsSplit = split(prgids, ", ")
                                strPrgidsSplit = split(strPrgids, ", ")

                                for p = 0 TO UBOUND(prgidsSplit)
                                    for tp = 0 TO UBOUND(strPrgidsSplit)
                                        if cint(prgidsSplit(p)) = cint(strPrgidsSplit(tp)) then
                                            if prgsidsToShow <> "" then
                                                prgsidsToShow = prgsidsToShow & ", " & prgidsSplit(p)
                                            else
                                                prgsidsToShow = prgidsSplit(p)
                                            end if
                                        end if
                                    next
                                next

                            end if

                            if prgsidsToShow = "" then
                            prgsidsToShow = "99999"
                            end if

                            if level = 1 then
                                prgsidsToShow = prgids
                            end if

                        else
                            prgsidsToShow = prgids
                        end if
                                                
                        sqlWHERE = "projektgruppeid IN ("& prgsidsToShow &")"
                    end if

                    'response.Write sqlWHERE
                    'response.End
                    
                    if cint(startyy) = cint(slutyy) then
                        orandsign = "AND"
                    else
                        orandsign = "OR"
                    end if
          
                    strOpt = "<option value='0'>Vælg</option>"
                    strSQL = "SELECT medarbejderid, mnavn, init FROM progrupperelationer LEFT JOIN medarbejdere ON (medarbejderid = mid) WHERE "& sqlWHERE & " GROUP BY medarbejderid ORDER BY mnavn"
                    oRec.open strSQL, oConn, 3
                    while not oRec.EOF
                        findesFC = 0
                        strSQL2 = "SELECT timer FROM ressourcer_md WHERE medid ="& oRec("medarbejderid") & " AND ((md >= "& startmd &" AND aar = "& startyy &") "& orandsign &" ( md <= "& slutmd &" AND aar = "& slutyy &")) AND aktid = "& aktid
                        oRec2.open strSQL2, oConn, 3
                        if not oRec2.EOF then
                                findesFC = 1
                        end if
                        oRec2.close

                        if findesFC = 0 then
                            strOpt = strOpt & "<option value='"& oRec("medarbejderid") &"'>"& oRec("mnavn") &" ("&oRec("init")&")</option>"
                        end if
                    oRec.movenext
                    wend
                    oRec.close

                    call jq_format(stropt)
                    stropt = jq_formatTxt

                    response.Write strOpt

                    case "HentAktiviteter"

                        jobid = request("jobid")
                        ' Hvis vis kun en aktivitet er slået til tager den kun den første, og bruger den til at forecast på, ellers skal den hente alle hvor efter man vælger hvilken man vil bruge
                        aktid = 0
                        strSQL = "SELECT id FROM aktiviteter WHERE job = "& jobid
                        oRec.open strSQL, oConn, 3 
                        if not oRec.EOF then
                            aktid = oRec("id")
                        end if
                        oRec.close

                        response.Write aktid


                    case "HentJobs"                       
                        medid = request("medid")
                        startmd = request("startMD")
                        startyy = request("startYY")
                        slutmd = request("slutMD")
                        slutyy = request("slutYY")

                        call medariprogrpFn(medid)
                        medariprogrpTxt = replace(medariprogrpTxt, "#", "")
                        
                        'Hvem er den loggede ind teamleder for
                        call fTeamleder(session("mid"), 0, 1) 
                        'response.Write "er teamleder for " & strPrgids
                        if strPrgids <> "ingen" AND strPrgids <> "alle" then
                            
                        end if

                        if cint(startyy) = cint(slutyy) then
                        orandsign = "AND"
                        else
                        orandsign = "OR"
                        end if

                        prgSQLSEL = "projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10"
                        
                        stropt = "<option value='0'>Vælg</option>"
                        strSQL = "SELECT id, jobnavn, kkundenavn, kid, jobans1, jobans2, jobans3, jobans4, jobans5, "& prgSQLSEL &" FROM job LEFT JOIN kunder ON (jobknr = kid) WHERE jobstatus = 1 "_
                            & "AND (projektgruppe1 IN ("& medariprogrpTxt &")"_
                            & "OR projektgruppe2 IN ("& medariprogrpTxt &")"_
                            & "OR projektgruppe3 IN ("& medariprogrpTxt &")"_
                            & "OR projektgruppe4 IN ("& medariprogrpTxt &")"_
                            & "OR projektgruppe5 IN ("& medariprogrpTxt &")"_
                            & "OR projektgruppe6 IN ("& medariprogrpTxt &")"_
                            & "OR projektgruppe7 IN ("& medariprogrpTxt &")"_
                            & "OR projektgruppe8 IN ("& medariprogrpTxt &")"_
                            & "OR projektgruppe8 IN ("& medariprogrpTxt &")"_
                            & "OR projektgruppe10 IN ("& medariprogrpTxt &"))"_
                            & " ORDER BY kkundenavn, jobnavn" 

                        oRec.open strSQL, oConn, 3
                        lastkid = -1
                        while not oRec.EOF
                    
                            'Tjekker om der allerede er FC på jobbet på medarbejderen, og dermed allerede er i listen
                            findesFC = 0
                            strSQL2 = "SELECT timer FROM ressourcer_md WHERE medid = "& medid & " AND ((md >= "& startmd &" AND aar = "& startyy &") "& orandsign &" ( md <= "& slutmd &" AND aar = "& slutyy &")) AND jobid = "& oRec("id")
                            oRec2.open strSQL2, oConn, 3
                            if not oRec2.EOF then
                                findesFC = 1
                            end if
                            oRec2.close

                            if findesFC = 0 then

                                erjobans = 0
                                if cdbl(oRec("jobans1")) = cdbl(session("mid")) OR cdbl(oRec("jobans2")) = cdbl(session("mid")) OR cdbl(oRec("jobans3")) = cdbl(session("mid")) OR cdbl(oRec("jobans4")) = cdbl(session("mid")) OR cdbl(oRec("jobans5")) = cdbl(session("mid")) then
                                erjobans = 1
                                end if

                                if erjobans = 1 OR level = 1 then
                                    if lastkid <> oRec("kid") then                
                                        stropt = stropt & "<option value='"& oRec("kid") &"' disabled>"& oRec("kkundenavn") &"</option>"
                                    end if

                                    stropt = stropt & "<option value='"& oRec("id") &"'>"& oRec("jobnavn") &"</option>"
                                else
            
                                    if strPrgids <> "ingen" then

                                        prgIds = oRec("projektgruppe1") & ", "_
                                            & oRec("projektgruppe2") & ", "_
                                            & oRec("projektgruppe3") & ", "_
                                            & oRec("projektgruppe4") & ", "_
                                            & oRec("projektgruppe5") & ", "_
                                            & oRec("projektgruppe6") & ", "_
                                            & oRec("projektgruppe7") & ", "_
                                            & oRec("projektgruppe8") & ", "_
                                            & oRec("projektgruppe9") & ", "_
                                            & oRec("projektgruppe10") & ", "

                                        strPrgidsSplit = split(strPrgids, ",")
                                        prgIdsSplit = split(prgIds, ",")                
                                
                                        visJob = 0

                                        for tp = 0 TO UBOUND(strPrgidsSplit)
                                            for p = 0 TO UBOUND(prgIdsSplit)
                                                if strPrgidsSplit(tp) = prgIdsSplit(p) then
                                                    visJob = 1
                                                end if
                                            next
                                        next

                                        if visJob = 1 then                                       

                                            if lastkid <> oRec("kid") then                
                                                stropt = stropt & "<option value='"& oRec("kid") &"' disabled>"& oRec("kkundenavn") &"</option>"
                                            end if

                                            stropt = stropt & "<option value='"& oRec("id") &"'>"& oRec("jobnavn") &"</option>"
                                        
                                        end if
                                    end if

                                end if

                                lastkid = oRec("kid")

                            end if 'FC = 0

                        oRec.movenext
                        wend
                        oRec.close

                        call jq_format(stropt)
                        stropt = jq_formatTxt

                        response.Write stropt                        

            end select
        response.End
        end if
        %>





        <!--#include file="../inc/regular/webblik_func.asp"-->
        <!--#include file="../timereg/inc/ressource_belaeg_jbpla_inc.asp"-->
        <!--#include file="../timereg/inc/isint_func.asp"-->
       
        <!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
        <!--#include file="../inc/regular/topmenu_inc.asp"-->

       
        <%' GIT 20-08-2018 - osn


    call positiv_aktivering_akt_fn()

    public visAktiv, visRamme
    public strSQLa
    function hentaktiviteter(positiv_aktivering_akt_val, jid, mid, aty_sql_realhoursAkt)
                    
                
                    '** Fjerner akt der allerede er i bruge for denne medarbejder **'
                    aktivIbrug = "OR (a.id = 0 "

                    if len(trim(mid)) <> 0 then
                    mid = mid
                    else
                    mid = 0
                    end if

                    if len(trim(jid)) <> 0 then
                    jid = jid
                    else
                    jid = 0
                    end if

                    strSQLaib = "SELECT aktid FROM ressourcer_md AS rmd WHERE rmd.medid = "& mid &" AND rmd.jobid = "& jid & " GROUP BY aktid" 'AND ((rmd.md >= "& startdatoMD &" AND rmd.aar = "& startdatoYY &") "& orandval &" (rmd.md <= "& slutdatoMD &" AND rmd.aar = "& slutdatoYY &")) GROUP BY aktid"
                    'response.write "strSQLaib " & strSQLaib
                    'response.flush            
        
                    oRec9.open strSQLaib, oConn, 3
                    while not oRec9.EOF 
                            
                    aktivIbrug = aktivIbrug & " OR a.id = "& oRec9("aktid")

                    oRec9.movenext
                    wend
                    oRec9.close                
            
                    aktivIbrug = aktivIbrug & ")"


                  

                    'aktivIbrug = ""
    
                    if cint(positiv_aktivering_akt_val) = 1 then
                    strSQLa = "SELECT navn AS aktnavn, a.id AS aid, fase FROM timereg_usejob AS tu "_
                    &" LEFT JOIN aktiviteter AS a ON (a.id = tu.aktid) "_
                    &" WHERE tu.jobid = "& jid &" AND tu.medarb = "& mid &" AND aktstatus = 1 "& aktivIbrug &" GROUP BY a.id ORDER BY fase, sortorder, navn"
                    else
                    strSQLa = "SELECT navn AS aktnavn, id AS aid, fase FROM aktiviteter AS a WHERE job = "& jid &" AND aktstatus = 1 AND ("& aty_sql_realhoursAkt &") "& aktivIbrug &" GROUP BY id ORDER BY fase, sortorder, navn"
                    end if


                    'if session("mid") = 1 then
                    'response.write strSQLa
                    'response.flush
                    'end if
                  

    end function


    function jobmedtimeruforecast(mid,vis)

    


        if len(trim(jobisWrt(mid))) <> 0 then
	    jobisWrt(mid) = jobisWrt(mid) & ")"
	    jobIdisWrt(mid) = jobIdisWrt(mid) & ")"
        else
	    jobisWrt(mid) = ""
	    jobIdisWrt(mid) = ""
        end if

    '********************************************************************************
    '*** finder job, medarb og kundeoplysninger, Real timer bliver fundet senere ***'
    '*** +Aktive job på personlig aktiv listen *************************************'
    '********************************************************************************

  

    strSQLt = "SELECT jobnavn, m.mnavn, m.mnr, tmnr, j.id AS jobid, j.jobnr, m.medarbejdertype, m.mid, "
    strSQLt = strSQLt & "SUM(t.timer) AS realtimer, "
    strSQLt = strSQLt & " k.kkundenavn, k.kkundenr, m.init, m.forecaststamp, jobans1, jobans2, "_
	&" m2.mnavn AS m2mnavn, m2.init AS m2init, m2.mnr AS m2mnr, m3.mnavn AS m3mnavn, m3.init AS m3init, m3.mnr AS m3mnr "_
    &" FROM timer t  "
	
	
	strSQLt = strSQLt &" LEFT JOIN medarbejdere m ON (m.mid = " & mid &")"_
	&" LEFT JOIN job j ON (j.jobnr = t.tjobnr AND "& jobSQLkri &" AND (j.risiko > -1 OR j.risiko = -3)) AND j.jobnavn IS NOT NULL"
	
	
    strSQLt = strSQLt &" LEFT JOIN kunder k ON (kid = j.jobknr)"_
	&" LEFT JOIN medarbejdere m2 ON (m2.mid = j.jobans1)"_
	&" LEFT JOIN medarbejdere m3 ON (m3.mid = j.jobans2)"
	
  

    strSQLt = strSQLt &" WHERE t.tmnr = "& mid &" "& jobisWrt(mid)
   
    strSQLt = strSQLt &" AND j.jobstatus = 1 AND t.tdato BETWEEN '"& tStDatoJobUforecast &"' AND '"& tSlDatoJobUforecast &"'"
	strSQLt = strSQLt &" GROUP BY t.tjobnr, t.tmnr ORDER BY t.tjobnavn"
    
    'if lto = "essens" then
    'Response.write "<br>Job med timre uden forecast: <b>"& strSQLt &"</b><br><br>"
	'Response.flush
    'end if
    
    
    jobMrealtimerPa = jobIdisWrt(mid)
    jobMrealtimerPa = replace(jobMrealtimerPa, ")", "")
    oRec2.open strSQLt, oConn, 3
    while not oRec2.EOF 

            if oRec2("realtimer") <> 0 then

                        medarbKundeoplysX(x, 0) = oRec2("jobid")
	                    medarbKundeoplysX(x, 1) = oRec2("jobnr")
                        medarbKundeoplysX(x, 2) = oRec2("jobnavn")
                        medarbKundeoplysX(x, 3) = oRec2("mid")
                        medarbKundeoplysX(x, 4) = oRec2("mnavn")
    
	                    medarbKundeoplysX(x, 5) = oRec2("kkundenavn")
	                    medarbKundeoplysX(x, 6) = oRec2("kkundenr")
	                    medarbKundeoplysX(x, 7) = oRec2("mnr")
	                    medarbKundeoplysX(x, 8) = oRec2("init")
	
	       
	                    medarbKundeoplysX(x, 9) = "2002-01-01" 'oRec("forecaststamp")
	                    medarbKundeoplysX(x, 10) = 0
	                    resid = -1
	       
	
	                    medarbKundeoplysX(x, 11) = oRec2("jobans1")
	                    medarbKundeoplysX(x, 12) = oRec2("jobans2")
	
	                    if len(trim(oRec2("m2mnavn"))) <> 0 then
                        medarbKundeoplysX(x, 13) = oRec2("m2mnavn") &" ("&oRec2("m2mnr")&")"
                        medarbKundeoplysX(x, 14) = oRec2("m2init") 
                        else
                        medarbKundeoplysX(x, 13) = ""
                        medarbKundeoplysX(x, 14) = ""
                        end if
    
                        if len(trim(oRec2("m3mnavn"))) <> 0 then
                        medarbKundeoplysX(x, 15) = oRec2("m3mnavn") &" ("&oRec2("m3mnr")&")"
                        medarbKundeoplysX(x, 16) = oRec2("m3init")
                        else
                        medarbKundeoplysX(x, 15) = ""
                        medarbKundeoplysX(x, 16) = ""
                        end if 
            

            
                        medarbKundeoplysX(x, 17) = 0
                        medarbKundeoplysX(x, 18) = ""

                        jobMrealtimerPa = jobMrealtimerPa & " AND jobid <> '"&  oRec2("jobid") & "'"
    
                    x = x + 1

            end if
    oRec2.movenext
    wend
    oRec2.close



    '***** Henter job der er med på aktivjoblisten for den enkelte medarbejder ****'
     if len(trim(jobMrealtimerPa)) <> 0 then
     jobMrealtimerPa = jobMrealtimerPa & ")"
     end if

    'PA = 0
    call positiv_aktivering_akt_fn()
    if cint(pa_aktlist) = 1 then
    paSQLkri = " AND forvalgt = 1 "
    else
    paSQLkri = ""
    end if

    'jobIswrtKri = jobisWrt(mid)
    strSQLtua = "SELECT aktid, medarb AS mid, m.mnavn, jobid, j.jobnr, j.jobnavn, "_
    &" kkundenavn, kkundenr, m.mnr, m.init, jobans1, jobans2, "_
    &" m2.mnavn AS m2mnavn, m2.init AS m2init, m2.mnr AS m2mnr, m3.mnavn AS m3mnavn, m3.init AS m3init, m3.mnr AS m3mnr FROM timereg_usejob AS tu "
    
    strSQLtua = strSQLtua &" LEFT JOIN medarbejdere m ON (m.mid = " & mid &")"_
	&" LEFT JOIN job j ON (j.id = tu.jobid AND (j.risiko > -1 OR j.risiko = -3))"
	
    strSQLtua = strSQLtua &" LEFT JOIN kunder k ON (kid = j.jobknr)"_
	&" LEFT JOIN medarbejdere m2 ON (m2.mid = j.jobans1)"_
	&" LEFT JOIN medarbejdere m3 ON (m3.mid = j.jobans2)"_
    &" WHERE medarb = "& mid &" "& jobMrealtimerPa & " "& paSQLkri &" AND aktid = 0 AND "& jobSQLkri &" AND (j.risiko > -1 OR j.risiko = -3) AND jobstatus = 1"
    
    'if lto = "essens" then
    'Response.write strSQLtua
    'Response.flush
    'end if
    oRec2.open strSQLtua, oConn, 3
    while not oRec2.EOF 

            

                        medarbKundeoplysX(x, 0) = oRec2("jobid")
	                    medarbKundeoplysX(x, 1) = oRec2("jobnr")
                        medarbKundeoplysX(x, 2) = oRec2("jobnavn")
                        medarbKundeoplysX(x, 3) = oRec2("mid")
                        medarbKundeoplysX(x, 4) = oRec2("mnavn")
    
	                    medarbKundeoplysX(x, 5) = oRec2("kkundenavn")
	                    medarbKundeoplysX(x, 6) = oRec2("kkundenr")
	                    medarbKundeoplysX(x, 7) = oRec2("mnr")
	                    medarbKundeoplysX(x, 8) = oRec2("init")
	
	       
	                    medarbKundeoplysX(x, 9) = "2002-01-01" 'oRec("forecaststamp")
	                    medarbKundeoplysX(x, 10) = 0
	                    resid = -1
	       
	
	                    medarbKundeoplysX(x, 11) = oRec2("jobans1")
	                    medarbKundeoplysX(x, 12) = oRec2("jobans2")
	
	                    if len(trim(oRec2("m2mnavn"))) <> 0 then
                        medarbKundeoplysX(x, 13) = oRec2("m2mnavn") &" ("&oRec2("m2mnr")&")"
                        medarbKundeoplysX(x, 14) = oRec2("m2init") 
                        else
                        medarbKundeoplysX(x, 13) = ""
                        medarbKundeoplysX(x, 14) = ""
                        end if
    
                        if len(trim(oRec2("m3mnavn"))) <> 0 then
                        medarbKundeoplysX(x, 15) = oRec2("m3mnavn") &" ("&oRec2("m3mnr")&")"
                        medarbKundeoplysX(x, 16) = oRec2("m3init")
                        else
                        medarbKundeoplysX(x, 15) = ""
                        medarbKundeoplysX(x, 16) = ""
                        end if 
            

            
                        medarbKundeoplysX(x, 17) = 0
                        medarbKundeoplysX(x, 18) = ""

           
    
                    x = x + 1

         
    oRec2.movenext
    wend
    oRec2.close


     
    'jobisWrt(mid) = " AND (t.tjobnr <> 0 "
    
    end function




%>

<div class="wrapper">
    <div class="content">

        <%
            if len(session("user")) = 0 then
	
	        errortype = 5
	        call showError(errortype)
	
	        response.End
            end if


            '*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	        Session.LCID = 1030
	
	        '*** Akt. typer der tælles med i ugereg. ***'
	        call akttyper2009(2)
	
	
	        func = request("func")
	
	        thisfile = "res_belaeg"
	        'Response.write "func: " & func
	
	        if func = "dbopr" then
	        id = 0
	        else
	        id = request("id")
	        end if


            if len(trim(request("FM_progrp"))) <> 0 then
	        progrp = request("FM_progrp")
	        else
            progrp = 0
	        end if


            if len(request("FM_medarb")) <> 0 then 'form submitted

            if len(trim(request("FM_expvisreal"))) <> 0 then
            expvisreal = request("FM_expvisreal")
            else
            expvisreal = 0
            end if

            response.cookies("tsa")("respreal") = expvisreal

            else
        
            if request.cookies("tsa")("respreal") <> "" then
            expvisreal = request.cookies("tsa")("respreal")
            else
            expvisreal = 0
            end if


            end if

            if cint(expvisreal) = 0 then
            expvisrealCHK = ""
            else
            expvisrealCHK = "CHECKED"
            end if


            'Response.Write "medid first: "& left(request("FM_medarb"), 1)
	        'Response.end
	
	        '*** Rettigheder på den der er logget ind **'
	        medarbid = session("mid")
	
	        vis = 0
	        if len(trim(request("vis"))) <> 0 then
            vis = request("vis")
            else
            vis = 0
            end if

            '**** vis som timer ell. procent ***'
            if len(trim(request("FM_showasproc"))) <> 0 then
                showasproc = request("FM_showasproc")
            else
        
                if request.cookies("tsa")("showasproc") <> "" then
                showasproc = request.cookies("tsa")("showasproc")
                else
                showasproc = 0
                end if

            end if 

            if cint(showasproc) = 1 then
                showasproc0Sel = ""
                showasproc1Sel = "SELECTED"
            else
                 showasproc0Sel = "SELECTED"
                showasproc1Sel = ""
            end if
            response.cookies("tsa")("showasproc") = showasproc 


            '** Job uden forecast
             if len(request("FM_medarb")) <> 0 then 'form submitted
    
                if len(trim(request("FM_vis_job_u_fc"))) <> 0 AND request("FM_vis_job_u_fc") <> 0 then
                    vis_job_u_fc = 1 'request("FM_vis_job_u_fc")
                    response.Cookies("resbel")("ufc") = vis_job_u_fc
                else
                    vis_job_u_fc = 0
                end if

                else 

                if request.cookies("resbel")("ufc") <> "" then
                        vis_job_u_fc = request.cookies("resbel")("ufc")
                    else
                        select case lto 
                        case "bf"
                        vis_job_u_fc = 1
                        case else
                        vis_job_u_fc = 0
                        end select
                end if
    
            end if


            if cint(vis_job_u_fc) = 0 then
            vis_job_u_fcCHK = ""
            else
            vis_job_u_fcCHK = "CHECKED"
            end if


             '** Vis simpel (uden susbtotaler)
            if len(request("FM_medarb")) <> 0 then 'form submitted
    
                if len(trim(request("FM_vis_simpel"))) <> 0 then
                vis_simpel = request("FM_vis_simpel")
                else
                vis_simpel = 0
                end if

            else 

                if request.cookies("resbel")("vsi") <> "" then
                 vis_simpel = request.cookies("resbel")("vsi")
                else

                    select case lto 
                    case "bf"
                    vis_simpel = 1
                    case else
                    vis_simpel = 0
                    end select          
                end if
    
            end if

            if len(trim(vis_simpel)) <> 0 then
            vis_simpel = vis_simpel
            else
            vis_simpel = 0
            end if


            'vis_simpel = 0
            vis_simpelCHK0 = ""
            vis_simpelCHK1 = ""
            vis_simpelCHK2 = ""


        'response.write "vis_simpel: " & vis_simpel
        'response.flush
        if lto = "ddc" OR lto = "outzx" OR lto = "hidalgo" OR lto = "care" then
                ddcView = 1
                vis_simpel = 0
        else
                ddcView = 0
        end if


        select case cint(vis_simpel) 
        case 1
            vis_simpelCHK1 = "SELECTED"
        case 2
            vis_simpelCHK2 = "SELECTED"
        case else
            vis_simpelCHK0 = "SELECTED"            
        end select

        response.Write "<input type='hidden' value='"&ddcView&"' id='ddcView' />"

        if ddcView = 1 then
            if len(trim(request("FM_sortering"))) <> 0 then
                sortering = request("FM_sortering")
                response.Cookies("to_2015")("resbel") = sortering
            else
                if len(request.Cookies("to_2015")("resbel")) <> 0 then
                sortering = request.Cookies("to_2015")("resbel")
                else
                sortering = 0
                end if
            end if

            select case sortering
                case 0 
                sortJOB = "SELECTED"
                sortMED = ""
                case 1
                sortJOB = ""
                sortMED = "SELECTED"
            end select
        else
            sortering = 1
        end if

        '** Job hvor forecast er overskreddet
        if len(request("FM_medarb")) <> 0 then 'form submitted
    
            if len(trim(request("FM_vis_job_fc_neg"))) <> 0 AND request("FM_vis_job_fc_neg") <> 0 then
            vis_job_fc_neg = 1 'request("FM_vis_job_fc_neg")
            else
            vis_job_fc_neg = 0
            end if

        else 

            if request.cookies("resbel")("fcneg") <> "" then
            vis_job_fc_neg = request.cookies("resbel")("fcneg")
            else
            vis_job_fc_neg = 0
            end if
    
        end if


        if cint(vis_job_fc_neg) = 0 then
        vis_job_fc_negCHK = ""
        else
        vis_job_fc_negCHK = "CHECKED"
        end if


        response.cookies("resbel")("fcneg") = vis_job_fc_neg
        response.cookies("resbel")("fcneg") = vis_job_fc_neg
        response.cookies("resbel")("ufc") = vis_job_u_fc
        response.cookies("resbel")("vsi") = vis_simpel

        if len(request("FM_medarb")) <> 0 then 'OR func = "export"
	
	        if left(request("FM_medarb"), 1) = "0" then
	        
                if media <> "print" then
	            thisMiduse = replace(request("FM_medarb_hidden"), "-", "0")
    	        else
    	        thisMiduse = replace(request("FM_medarb"), "-", "0")
    	        end if
    	
                intMids = split(thisMiduse, ", ")
	        else
	            thisMiduse = replace(request("FM_medarb"), "-", "0")
	            'Response.write "<br>thisMiduse "& thisMiduse
                intMids = split(thisMiduse, ", ")
	        end if
	
        else
	    
	    thisMiduse = session("mid") 
	    intMids = split(thisMiduse, ", ")
	   
        end if

        'Response.Write "FM_medarb:" &  request("FM_medarb")
	

	    media = request("media")
	    'print = request("print")
	
	    '*** Periode ***
	    dim timerTotalTotal, maanedY, medarbTotalTimer, medarbTimerReal, realTotalTotal
        dim timerMJobTotal, realMJobTotal, fc_lastjobnrWrt, fc_lastaktidWrt


        if len(request("periodeselect")) <> 0 then
	        periodeSel = request("periodeselect")
	    'periodeShow = ""& periodeSel &" Måneder frem"
	        perBeregn = periodeSel - 1
	    else	
	        if len(request.cookies("resbel")("periodesel")) <> 0 then
	        periodeSel = request.cookies("resbel")("periodesel")
	        'periodeShow = ""& periodeSel &" Måneder frem"
	        perBeregn = periodeSel - 1
	        else
	        periodeSel = 6
	        'periodeShow = "6 Måneder frem"
	        perBeregn = 5
	        end if
	    end if

        if ddcView = 1 then
            periodeSel = 6
        end if


        select case periodeSel
	    'case 1
	    'leftwdt = 35
	    'rdimnb = 31
	    'periodeShow = periodeShow & " dage"
	    case 3
	    stDatoDag = 1 
	    rdimnb = 15
	    leftwdt = 350
	    periodeShow = periodeShow & " " & resbelaeg_txt_036
	    case 6
	    stDatoDag = 14 '** for at være sikker på at der ikke bliver skrvet en uge ind, der hører til forrige måned.
	    rdimnb = 7
	    leftwdt = 500
	    periodeShow = resbelaeg_txt_038
	    case 12
	    stDatoDag = 14
	    rdimnb = 13
	    leftwdt = 850
	    periodeShow = resbelaeg_txt_037
	    end select


        redimNo = 9500

	    Redim timerTotalTotal(redimNo), maanedY(redimNo), medarbTotalTimer(redimNo), medarbTimerReal(redimNo), realTotalTotal(redimNo) 
        Redim timerMJobTotal(redimNo), realMJobTotal(redimNo), fcGrandtotal(redimNo), realGrandtotal(redimNo), jobTimerGrandGrand(redimNo), jobBgtGrandGrand(redimNo)
        Redim procMJobTotal(redimNo), procTotalTotal(redimNo),  medarbTotalProc(redimNo), fpGrandtotal(redimNo) 

        realignPerGrandGrand = 0
        realGrandGrand = 0
        fcignPerGrandGrand = 0 
        fcGrandGrand = 0
        jobTimerGrandGrand = 0
        jobBgtGrandGrand = 0

        response.cookies("resbel")("periodesel") = periodeSel


        '************** Indlæser variable til sidevisning *************

	    '*** Hvis plus eller minus knap er brugt ***
	    if len(request("selectplus")) <> 0 then
	    selectplus = request("selectplus")
	    else
	    selectplus = 0
	    end if
	
	    if len(request("selectminus")) <> 0 then
	    selectminus = request("selectminus")
	    else
	    selectminus = 0
	    end if

        'response.write "request "& request("mdselect") & "minus: "& selectminus 


        if len(request("mdselect")) <> 0 then
			monththis = request("mdselect")
				
			if selectminus <> 0 then
			monththis = monththis - 1
			end if
				
			if selectplus <> 0 then
			monththis = monththis + 1
			end if
				
				
			select case monththis 
			case 0
			monththis = 12
			yearthis = request("aarselect")
			yearthis = yearthis - 1
			case 13
			monththis = 1
			yearthis = request("aarselect")
			yearthis = yearthis + 1
			case else
			yearthis = request("aarselect")
			end select

            response.cookies("resbel")("mm") = monththis
            response.cookies("resbel")("yyyy") = yearthis
				
		else

            if request.cookies("resbel")("mm") <> "" then

                monththis = request.cookies("resbel")("mm")
                yearthis = request.cookies("resbel")("yyyy")

            else
               

			monththis = month(now)
			yearthis = year(now)

            end if
		
        end if


        jobSlutTEMP = dateadd("m", perBeregn, "1/"& monththis &"/"& yearthis) '** år ?

        monthUse = datepart("m", jobSlutTEMP)
	    yearUse = datepart("yyyy", jobSlutTEMP)
	
	    call dageimd(monthUse,yearUse)	
	    mthDaysUse = mthDays

        select case periodeSel
        case 3
            jobStartKri = yearthis&"/"&monththis&"/"&stDatoDag

            dayOfWeek = datePart("w",jobStartKri,2,2)
        
            '1 = (dayOfWeek - d) 
            d = (dayOfWeek - 1)
            firstdayOfFirstWeek = dateAdd("d", -d, jobStartKri)
            if month(firstdayOfFirstWeek) <> monththis then
            firstdayOfFirstWeek = dateAdd("d", 7, firstdayOfFirstWeek)
            end if

            firstThursdayOfFirstWeek = firstdayOfFirstWeek 'dateAdd("d", +3, firstdayOfFirstWeek)

            jobStartKri = firstThursdayOfFirstWeek 'firstdayOfFirstWeek 'year(firstdayOfFirstWeek) &"-"& month(firstdayOfFirstWeek) &"-"& year(firstdayOfFirstWeek)

        case else
	        jobStartKri = yearthis&"/"&monththis&"/"&stDatoDag
        end select

        jobSlutKri = yearUse&"/"&monthUse&"/"&mthDaysUse


        '*** Bruges til at tjekke det nøjagtige normtoimer pr. medarb. ved tilføje ny linie.
	    datoInterval = datediff("d",jobStartKri,jobSlutKri,2,2)


        '**** Tilføjer en md hvis der er en uge går ind i ny måned.
        if periodeSel = 3 then
        jobSlutKri = dateAdd("d", 1, jobSlutKri)
        'jobSlutKri = dateAdd("d")	
        jobSlutKri = year(jobSlutKri)&"/"&month(jobSlutKri)&"/"&day(jobSlutKri)
        end if


        'response.write "jobSlutKri: "& jobSlutKri

	    daynow = day(now)&"/"&month(now)&"/"&year(now)

        

            if len(trim(request("FM_jobsel"))) <> 0 then 		
	            jobidsel = request("FM_jobsel")
	            response.cookies("resbel")("jobidsel") = jobidsel
	        else
	            if len(request.cookies("resbel")("jobidsel")) <> 0 then
		            jobidsel = request.cookies("resbel")("jobidsel")
		        else
		            jobidsel = 0
		        end if
	        end if

        if ddcView <> 1 then

            if cint(jobidsel) <> 0 then
	        jobSQLkri = " j.id = " & jobidsel
	        else
	        jobSQLkri = " j.id <> 0 "
            'jobSQLkri = " m.mid <> 0 " 'hvis ingen job er ibrug må der ikke tjekkes på job id i wh clause da listen så bliver tom.
	        end if



            if len(trim(request("FM_sog"))) <> 0 then
            sogTxt = request("FM_sog")
            else
                if trim(request.cookies("tsa")("sogval")) <> "" AND len(trim(request("mdselect"))) = 0 then
                sogTxt = request.cookies("tsa")("sogval")
                else
                sogTxt = ""
                end if
            end if

            if ddcView = 1 then
                sogTxt = ""
            end if


            if len(trim(sogTxt)) <> 0 then
                jobSogVal = sogTxt
                jobSQLkri = ""

                if instr(jobSogVal, ",") <> 0 then '** Komma **'
	            
                    jobSQLkri = " jobnr = 0 "
	                jobSogValuse = split(jobSogVal, ",")
	                for i = 0 TO UBOUND(jobSogValuse)
	                    jobSQLkri = jobSQLkri & " OR jobnr = '"& jobSogValuse(i) &"'"   
	                next
	    
	                jobSQLkri = "("&jobSQLkri&")"
	    
	            else


                    if instr(jobSogVal, "-") <> 0 then '** Interval
	                jobSogValuse = split(jobSogVal, "-")
	                jobSQLkri = "(jobnr >= '"& trim(jobSogValuse(0)) &"' AND jobnr <= '" & trim(jobSogValuse(1)) & "')"   
	                else 
                    jobSQLkri = "(jobnavn LIKE '" & jobSogVal &"%' OR jobnr = '"& jobSogVal &"')"
                    end if

	            end if
	
	
	        Response.cookies("tsa")("sogval") = sogTxt	
            end if

        end if 'ddcView

        if ddcView = 1 then
            if jobidsel = "0" then
            jobSQLkri = " j.id <> 0"
            else
            jobSQLkri = " j.id IN ("&jobidsel&")"
            end if

            jobids = split(jobidsel, ",")
        end if


        if len(trim(request("viskunjobiper"))) <> 0 then
	        viskunjobiper = request("viskunjobiper")

	    else
	    
	        if len(request.cookies("resbel")("viskunjobiper")) <> 0 then
		    viskunjobiper = request.cookies("resbel")("viskunjobiper")
		    else
		    viskunjobiper = 0
		    end if	
	    end if 

        if ddcView = 1 then
            viskunjobiper = 0
        end if
	
        response.cookies("resbel")("viskunjobiper") = viskunjobiper

        if ddcView <> 1 then
            if len(trim(request("FM_hideEmptyEmplyLines"))) <> 0 then
                hideEmptyEmplyLines = request("FM_hideEmptyEmplyLines")
            else

                if len(request.cookies("resbel")("hideEmptyEmplyLines")) <> 0 AND len(trim(request("mdselect"))) = 0 then
		        hideEmptyEmplyLines = request.cookies("resbel")("hideEmptyEmplyLines")
		        else
		        hideEmptyEmplyLines = 0        
		        end if
          
            end if
        else
            hideEmptyEmplyLines = 1
        end if


        if cint(hideEmptyEmplyLines) = 1 then
        hideEmptyEmplyLinesCHK = "CHECKED"
        else
        hideEmptyEmplyLinesCHK = ""
        end if

        response.cookies("resbel")("hideEmptyEmplyLines") = hideEmptyEmplyLines




        if len(trim(request("mdselect"))) <> 0 OR len(trim(request("FM_vis_kolonne_simpel"))) <> 0 then
            if len(trim(request("FM_vis_kolonne_simpel"))) <> 0 then
            vis_kolonne_simpel = request("FM_vis_kolonne_simpel")
            else
            vis_kolonne_simpel = 0
            end if
        else

            if len(request.cookies("resbel")("vis_kolonne_simpel")) <> 0 then
            vis_kolonne_simpel = request.cookies("resbel")("vis_kolonne_simpel")
		    else
                vis_kolonne_simpel = 0
		    end if          
        end if


        response.cookies("resbel")("vis_kolonne_simpel") = vis_kolonne_simpel


        if len(trim(request("mdselect"))) <> 0 OR len(trim(request("FM_vis_kolonne_simpel_akt"))) <> 0 then
           if len(trim(request("FM_vis_kolonne_simpel_akt"))) <> 0 then
            vis_kolonne_simpel_akt = request("FM_vis_kolonne_simpel_akt")
           else
            vis_kolonne_simpel_akt = 0
            end if
        else

            if len(request.cookies("resbel")("vis_kolonne_simpel_akt")) <> 0 then
            vis_kolonne_simpel_akt = request.cookies("resbel")("vis_kolonne_simpel_akt")
		    else
             vis_kolonne_simpel_akt = 0
		    end if         
        end if

        if ddcView = 1 then
            vis_kolonne_simpel_akt = 1
        end if
            

        'response.write "vis_kolonne_simpel_akt: "& vis_kolonne_simpel_akt

        response.cookies("resbel")("vis_kolonne_simpel_akt") = vis_kolonne_simpel_akt


        '*** Licens indstillinger (overruler) ****'
        select case lto
        case "kejd_pb", "xintranet - local"     
            vis_kolonne_simpelDisabled = "DISABLED"
            vis_kolonne_simpel_aktDisabled = "DISABLED"
        case else		  
            vis_kolonne_simpelDisabled = ""
            vis_kolonne_simpel_aktDisabled = ""
        end select


        select case lto
        case "kejd_pb", "xintranet - local"
        vis_kolonne_simpel = 1
        case else
		vis_kolonne_simpel = vis_kolonne_simpel
        end select

        select case lto
        case "kejd_pb", "xintranet - local"
        vis_kolonne_simpel_akt = 1
        case else
		vis_kolonne_simpel_akt = vis_kolonne_simpel_akt
        end select

        if cint(vis_kolonne_simpel) = 1 then
        vis_kolonne_simpelCHK = "CHECKED"
        else
        vis_kolonne_simpelCHK = ""
        end if


        if cint(vis_kolonne_simpel_akt) = 1 then
        vis_kolonne_simpel_aktCHK = "CHECKED"
        else
        vis_kolonne_simpel_aktCHK = ""
        end if


        select case vis_kolonne_simpel_akt
        case 1
        visAktiv = 0
        case else
        visAktiv = 1
        end select 

        select case vis_kolonne_simpel
        case 1
        visRamme = 0
        case else
        visRamme = 1
        end select 


        '** Tjekekr variable ****'
        'response.write "vis_simpel: "& vis_simpel & " -- vis: "& vis & " -- media = "& media & " -- func: "& func
        'response.write  " -- viskunjobiper: "& viskunjobiper 
        'response.write  " -- hideEmptyEmplyLines: "&  hideEmptyEmplyLines
        'response.write  " -- vis_kolonne_simpel: "&  vis_kolonne_simpel 
        'response.write " -- vis_kolonne_simpel_akt: "& vis_kolonne_simpel_akt & " -- FM_vis_job_u_fc:" & vis_job_u_fc & " -- FM_vis_job_fc_neg:" & vis_job_fc_neg



        SELECT case viskunjobiper 
        case 1
	    viskunjobiper1 = "CHECKED"
	    viskunjobiper0 = ""
        'viskunjobiper2 = ""
	    jobDatoKri = " AND jobslutdato >= '"& jobStartKri &"' "
   
	    case else
	    jobDatoKri = ""
	    viskunjobiper1 = ""
	    viskunjobiper0 = "CHECKED"
        'viskunjobiper2 = ""
	    end select
	
	    'if len(request("showonejob")) <> 0 then
	    'showonejob = request("showonejob")
	    'else
	    showonejob = 0
	    'end if


        if (level <= 2 OR level = 6) then 
	        whSQLkri = " m.mid <> 0 "
	        multiUse = "multiple"
	
	        call licensStartDato()
	
	        if licensklienter < 20 then
	            if licensklienter < 5 then
	            szi = 5
	            else
	            szi = 10
	            end if
	        else
	            if licensklienter > 36 then
	            szi = 20
	            else
	            szi = 15
	            end if
	        end if
	
	    else
	        szi = 1
	        multiUse = ""
	        whSQLkri = " m.mid = " & medarbSel
	    end if


        'Response.Write "her" & medarbSel & "<br>"
	
	    'if medarbSel = "0" Then
	    'strMedidSQL = " m.mid <> 0 "
	    'strMedidSel = 0
	    'else
	     '   strMedidSQL = " (m.mid = 0"
	     '   medarbSelArr = split(trim(medarbSel), ",")
	     '   for m = 0 to UBOUND(medarbSelArr)
	     '   strMedidSQL = strMedidSQL & " OR m.mid = " & trim(medarbSelArr(m))
	     '   strMedidSel = strMedidSel & "#"& trim(medarbSelArr(m)) &"#"
	     '   next
	     '   strMedidSQL = strMedidSQL & ")"
         'end if
	


        '*** finder udfra valgte projektgrupper og medarbejdere
	    'medarbSQlKri 
	    'kundeAnsSQLKri

        medarbSQlKri2 = ""
        ddcMid = 0
        for m = 0 to UBOUND(intMids)
			    
		    if m = 0 then
		    medarbSQlKri = "(m.mid = " & intMids(m)
			ddcMid = intMids(m)
            medarbSQlKri2 = intMids(m)
		    else
		    medarbSQlKri = medarbSQlKri & " OR m.mid = " & intMids(m)
			medarbSQlKri2 = medarbSQlKri2 & ", " & intMids(m)
		    end if
			
	    next
			 
        if ddcView = 1 then 
            if request("FM_medarb") <> "0" then
	            medarbSQlKri = medarbSQlKri & ")"
            end if
        else
            medarbSQlKri = medarbSQlKri & ")"
        end if

        'medarbSQlKri = strMedidSQL    
	
	
	    response.cookies("resbel").expires = date + 30
	
	    'Response.write strMedidSQL & "<br>"& strMedidSel
	    'Response.flush



        medarbIswrt = ""
	    if func = "redalle" then
	
						
						 'firstLoop = 1
						'*** Datoer ****
						datoer = split(request("FM_dato"), ",")
						
                        'Response.Write request("FM_jobid") & "<hr>"
                        'Response.Write request("FM_aktid") & "<hr>"
						'Response.flush

						tmedid = split(request("FM_medarbid"), ",")
						tjobid = split(request("FM_jobid"), ",")
                        taktid = split(request("FM_aktid"), ",")
                        taktid_old = split(request("FM_aktid_old"), ",")
                      
                        
						
						tTimertildelt_temp = replace(request("FM_timer"), ", #,", ";")
						tTimertildelt_temp2 = replace(tTimertildelt_temp, ", #", "")
						
						'Response.write request("FM_jobid") & "<br>"
						'Response.write tTimertildelt_temp2
						'Response.flush
						
						tTimertildelt = split(tTimertildelt_temp2, ";") 'split(request("FM_timer"), ";")

                        select case periodeSel
                        case 3 
		                mdorweek = 1
                        case else 
                        mdorweek = 0
                        end select				

						'*** Validerer ***
						antalErr = 0
						for y = 0 to UBOUND(tTimertildelt)
						
			                'response.Write "<br> --------------------- " & tTimertildelt(y)			

							call erDetInt(SQLBless(trim(tTimertildelt(y))))
							if isInt > 0 then
								antalErr = 1
								errortype = 67
								'useleftdiv = "t"
								
								call showError(errortype)
								response.end
							end if	
							isInt = 0
						
						next


                       
                      			
						for y = 0 to Ubound(datoer)
						
						'Response.write "her<br>"
                        if y > 0 then
                                                
                        'if tmedid(y-1) <> tmedid(y) OR taktid_old(y-1) <> taktid_old(y) OR tjobid(y-1) <> tjobid(y) then 
                        '            firstLoop = 1
						'end if

                        end if

                        response.Write "DAte " & datoer(y)
                        if tmedid(y) <> 0 AND tjobid(y) <> 0 then 'Jobid og medarb SKAL være udfyldt
                        response.Write "<br> " & tmedid(y) & " - " & tjobid(y)
						if cint(instr(medarbIswrt, "#"&tmedid(y)&"#")) = 0 then
								
                                '*** Opdaterer medarbejder TimeStamp ***
								medarbStamp = year(now)&"/"&month(now)&"/"&day(now)
								strSQlmed = "UPDATE medarbejdere SET forecaststamp = '"& medarbStamp &"' WHERE mid = "& tmedid(y)		
								
								oConn.execute(strSQLmed)
								
								medarbIswrt = medarbIswrt & "#"& tmedid(y)&"#,"
						
						end if
							
							
												
												monththis = datepart("m", datoer(y), 2,2)
												'monththis = datepart("m", datoer(y))
												yearthis = datepart("yyyy", datoer(y))
                                                call thisWeekNo53_fn(datoer(y))
												weekThis = thisWeekNo53 'datepart("ww", datoer(y), 2,2)
												
                                                'Response.write "datoer(y): "& datoer(y) & " tTimertildelt(y) : "& tTimertildelt(y) &"<br>"
												
                                            

													'*** Opretter nye records ***
												
													if len(trim(tTimertildelt(y))) <> 0 then
													
													
															if tTimertildelt(y) <> 0 then
																
																'*** Renser ud i tabellen 
                                                                 
																select case periodeSel
																case 3
																strSQLperKri = " AND uge = "& weekThis
                                                                'weekFaktor = 1 'gns antal uger pr. md
                                                                'normPer = 1
																case 6,12
                                                                'normPer = 30
														        'weekFaktor = 4 'gns antal uger pr. md
                                                                strSQLperKri = " AND md = "& monththis
																end select
																
																strSQLdel = "DELETE FROM ressourcer_md WHERE jobid = "& tjobid(y) &" AND medid = "& tmedid(y) &" AND aktid = "& taktid_old(y) &""_
																&"  AND aar =  "& yearthis & strSQLperKri
															    
															    'AND md = "& monththis &"
															    'Response.Write "<br>periodeSel"& periodeSel &" strSQLdel:" & strSQLdel &"<br>"
																'Response.flush
                                                                oConn.execute strSQLdel
															    
                                                                jan1 = "1-1-"& yearthis
                                                                nu = dateAdd("ww", weekThis-1, jan1)
                                                                'firstdayoffWeek = dateadd("d",(-datepart("w",jan1)),nu)
                                                                ugeDag = datePart("w",nu,2,2)

																if cint(showasproc) <> 1 then
																
                                                                    if len(trim(tTimertildelt(y))) <> 0 then
                                                                    timerThisMD = tTimertildelt(y) 'sqlbless(tTimertildelt(y))
                                                                    else
                                                                    timerThisMD = 0
                                                                    end if
                                                                   

                                                                    call normtimerPer(tmedid(y), nu, 6, 1)

                                                                    if len(trim(ntimPer)) <> 0 AND ntimPer > 0 then
                                                                    ntimPer = cint(ntimPer)
                                                                    else
                                                                    ntimPer = 1
                                                                    end if

                                                                    procThisMD = (timerThisMD/ntimPer) * 100	
                                                                    procThisMD = replace(formatnumber(procThisMD, 0), ",",".")
                                                                
                                    
                                                                else
									                            
                                                                    if len(trim(tTimertildelt(y))) <> 0 then
                                                                    procThisMD = tTimertildelt(y) 'sqlbless(tTimertildelt(y))
                                                                    else
                                                                    procThisMD = 1
                                                                    end if
                                                                    
                                                                    

                                                                
                                                                    call normtimerPer(tmedid(y), nu, 6, 1)
                                                                    timerThisMD = replace(formatnumber((ntimPer * (procThisMD/100)), 0), ",",".")						    
                                                                        'nTimerPerIgnHellig

                                                                     'response.write "------------------------------------------------ tTimertildelt(y) " & tTimertildelt(y) & "timerThisMD: "& timerThisMD & " nTimerPerIgnHellig: "& nTimerPerIgnHellig &"<br>"

                                                                end if

                                                                'response.write "<br>dt: "& nu &"  ugeDag: "& ugeDag & " nTimerPerIgnHellig: "& nTimerPerIgnHellig & " formatnumber(tTimertildelt(y), 0): "& formatnumber(tTimertildelt(y), 0)

                                                                if tjobid(y) <> 0 then

                                                                '*** Renser ud i tabellen 
																select case periodeSel
																case 3
                                                                timerThisMD = replace(timerThisMD, ".", "")
                                                                timerThisMD = replace(timerThisMD, ",", ".")
                                                                procThisMD = replace(procThisMD, ".", "")
                                                                procThisMD = replace(procThisMD, ",", ".")
															   
                                                                strSQL = "INSERT INTO ressourcer_md"_
																& " (jobid, medid, md, aar, timer, uge, aktid, proc) VALUES ("_
																&" "& tjobid(y) &", "& tmedid(y) &", "& monththis &", "& yearthis &", "& timerThisMD &", "& weekThis &", "& taktid(y) &", "& procThisMD &")" 
																
                                                                
                                                                'if lto = "kejd_pb" then
                                                                'Response.Write "<br>strSQL A"& strSQL & "<br>"
                                                                'Response.flush
                                                                'end if
                                                                oConn.execute strSQL

																case 6,12
                                                                
                                                                firstOfmonth = "1-"& month(datoer(y)) &"-"& year(datoer(y))
                                                                thisMth = datepart("m", firstOfmonth, 2,2)
                                                                
                                    
                                                                nextMonth = dateAdd("m", 1, firstOfmonth)
                                                                nextMonthFirst = "1-"& month(nextMonth) & "-"& year(nextMonth)   
                                                                lastDInMonth = dateAdd("d", -1, nextMonthFirst) 'for at sikere vi kommer tilbage midt i ugen før. Der er den sidste uge i måneden før.
                                                                
                                                                dayOfLastWeekThis = datePart("w",lastDInMonth,2,2) - 1
                                                                firstdayOfLastWeekThis = dateAdd("d", -dayOfLastWeekThis, lastDInMonth)
                                    
                                                                dividerThisMth = datediff("w", firstOfmonth, firstdayOfLastWeekThis,2,2)
                                                        
                                                               

                                                                dayOfWeek = datePart("w",firstOfmonth,2,2)
                                                                d = (dayOfWeek - 1)
                                                                firstdayOfFirstWeek = dateAdd("d", -d, firstOfmonth)
                                                                firstThursdayOfFirstWeek = firstdayOfFirstWeek 'dateAdd("d", +3, firstdayOfFirstWeek)
                                                                
                                                                'if year(nextMonthFirst) <> year(firstOfmonth) then 'sidste uge i år ræekker ind i det nye år
                                                                'dividerThisMth = dividerThisMth - 1
                                                                'firstThursdayOfFirstWeek = dateAdd("d", 7, firstThursdayOfFirstWeek)
                                                                'end if

                                                             
                                                             
                                                                if cint(showasproc) <> 1 then
                                                                timerThisMD = formatnumber(timerThisMD/((dividerThisMth/1)+1),2)
                                                                
                                                                procThisMD = formatnumber(procThisMD/((dividerThisMth/1)+1),2)
                                                                
                                                
                                                                end if            


                                                                timerThisMD = replace(timerThisMD, ".", "") 
                                                                timerThisMD = replace(timerThisMD, ",", ".") 


                                                                procThisMD = replace(procThisMD, ".", "") 
                                                                procThisMD = replace(procThisMD, ",", ".")   
                                            
                                                                
                                                                for weekno = 0 to dividerThisMth 
                                                                
                                                            
                                                                thisDate = dateAdd("d", 7*weekno, firstThursdayOfFirstWeek)
                                                                
                                                                if datePart("m", thisDate, 2,2) < datePart("m", firstOfmonth, 2,2) OR datePart("yyyy", thisDate, 2,2) < datePart("yyyy", firstOfmonth, 2,2) then
                                                                thisDate = dateAdd("d", 7, thisDate)
                                                                firstThursdayOfFirstWeek = dateAdd("d", 7, firstThursdayOfFirstWeek)
                                                                end if
                                                    
                                                                call thisWeekNo53_fn(thisDate)
                                                                weekThis = thisWeekNo53 'datepart("ww", thisDate, 2,2)

                                                                monththis = datepart("m", thisDate, 2,2)
												                yearthis = datepart("yyyy", thisDate)
												               
                                                
                                                                            

                                                                'response.write "dateDiff(m, thisDateGlobal, datoer(y), 2,2): "& dateDiff("m", thisDateGlobal, datoer(y), 2,2)&" thisMth: "& thisMth &" monththis: "& monththis &"<br>   Y: "& y &" datoer(y): "& datoer(y) &" thisDateGlobal: " & thisDateGlobal& " DateDiff M: "& dateDiff("m", thisDateGlobal, datoer(y),2,2)  &" dividerThisMth: "& dividerThisMth & " thisDate: "& thisDate & " firstOfmonth: "& firstOfmonth &" lastDInMonth = "& lastDInMonth &"<br>"
                                                                '** Ved månedsindlæsning må der kune læses timer ind i den valgte måned. Derfor tjekkes om uge om torsdagen går over i næste måned
                                                                '** F.eks slut apr 2014 går over i maj og må ikke indlæses ***
                                                        
                                                                'if cint(thisMth) = cint(monththis) then

                                                                strSQL = "INSERT INTO ressourcer_md"_
																& " (jobid, medid, md, aar, timer, uge, aktid, proc) VALUES ("_
																&" "& tjobid(y) &", "& tmedid(y) &", "& thisMth &", "& yearthis &", "& timerThisMD &", "& weekThis &", "& taktid(y) &", "& procThisMD &")" 
																
                                                                'if lto = "wwf" then
                                                                'response.write "dividerThisMth" & dividerThisMth & " showasproc: "& showasproc & "<br>"
                                                                'Response.Write "strSQL A  "& strSQL & "<br><br>"
                                                                'end if                            

                                                                oConn.execute strSQL

                                                                'end if

                                                                'lastweekwrt = weekThis

                                                                
                                    
                                                                next
                                                                'thisDateGlobal = dateAdd("d", 7, thisDate)

																end select

															
																
															   'if lto = "wwf" then
															   'Response.end
                                                               'end if
																
																

                                                                end if
																
															
															else
																		
																'**** Sletter evt. gamle registreringer (Hvis 0 er angivet) ***	
																'*** Renser ud i tabellen 
																select case periodeSel
																case 3
																strSQLperKri = " AND uge = "& weekThis
																case 6,12
														        strSQLperKri = " AND md = "& monththis
																end select
																
																
																strSQLdel = "DELETE FROM ressourcer_md WHERE jobid = "& tjobid(y) &" AND medid = "& tmedid(y) &" AND aktid = "& taktid_old(y) &""_
																&" AND aar =  "& yearthis & strSQLperKri
															    'AND md = "& monththis &"
															    'Response.Write "<br>strSQldel B"& strSQldel & "<br>"
															    'Response.flush
															    
																oConn.execute strSQLdel
														
															end if
													end if
												
							
								end if 'len(tmedid(y)) <> 0 AND tjobid(y)				
						next
	                    
                    
	                    'Response.end

                        '***** Års rammer ****'

                        'Response.Write request("FM_aarsramme_medarb")
                        'Response.end

                        arr_rids = split(request("FM_aarsramme_rid"), ",")
                        arr_medids = split(request("FM_aarsramme_medarb"), ",")
                        arr_jobids = split(request("FM_aarsramme_jobid"), ",") 
                        arr_aarids = split(request("FM_aarsramme_aar"), ",") 

                        arr_timerids_temp = replace(request("FM_aarsramme_timer"), ", #,", ";")
						arr_timerids_temp2 = replace(arr_timerids_temp, ", #", "")
						
						arr_timerids = split(arr_timerids_temp2, ";") 'split(request("FM_timer"), ";")
                        

                        'arr_aarids = split(request("FM_aarsramme_timer"), "# ,") 

                        
                        for a = 0 to UBOUND(arr_medids)
                            
                            arr_timerids(a) = replace(arr_timerids(a), ",", ".")

                            

                            if arr_rids(a) <> 0 then
                            strSQLaar = "UPDATE ressourcer_ramme SET timer = "& arr_timerids(a) &" WHERE id = " & arr_rids(a)
                            oConn.execute(strSQLaar)
                            else

                            if arr_timerids(a) <> 0 then
                            strSQLaar = "INSERT INTO ressourcer_ramme (jobid, medid, aar, timer) VALUES ("& arr_jobids(a) &", "& arr_medids(a) &", "& arr_aarids(a) &", "& arr_timerids(a) &")" 
                            'Response.Write strSQLaar & "<br>"
                            oConn.execute(strSQLaar)
                            end if

                            end if


                        next
       
            'Response.end

	
	
	


	
	    'Response.end
	    response.Redirect "ressource_belaeg_jbpla.asp?FM_progrp="&progrp&"&FM_medarb="&thisMiduse&"&FM_medarb_hidden="&thisMiduse&"&showonejob="&showonejob&"&id="&id&"&periodeselect="&periodeSel&"&jobselect="&jobnavn&"&mdselect="&request("mdselect")&"&aarselect="&request("aarselect")&"&selectplus="&selectplus&"&selectminus="&selectminus&"&FM_vis_job_u_fc="& vis_job_u_fc &"&FM_vis_job_fc_neg="&vis_job_fc_neg&"&FM_hideEmptyEmplyLines="&hideEmptyEmplyLines&"&FM_vis_simpel="&vis_simpel&"&FM_vis_kolonne_simpel="&vis_kolonne_simpel&"&FM_vis_kolonne_simpel_akt="&vis_kolonne_simpel_akt&"&FM_sog="& sogTxt 
	
	
	    end if




         if func = "overfor" then


                                   


            if len(trim(request("FM_overfor"))) <> 0 then
            overfor = 1
            else
            overfor = 0
            end if 


            if cint(overfor) = 1 then

            overforfra = request("FM_overforfra") 
            overforfraAar = request("FM_overforfraAar")


           lastjob = 0

            select case periodeSel
            case 3
            perSQL = " uge = "& overforfra & ""
            perSQL2 = "WEEK(tdato) = " & overforfra

            if cint(overforfra) < 52 then
            overfortil = overforfra + 1
            overfortilAar = overforfraAar
            else
            overfortil = 1
            overfortilAar = overforfraAar + 1
            end if
    
                                                                                          
            perSQLtil = " uge = "& overfortil & ""
            perSQLtil2 = "WEEK(tdato) = " & overfortil

            case else
            perSQL = " md = "& overforfra & ""
            perSQL2 = "MONTH(tdato) = " & overforfra


            if cint(overforfra) < 12 then
            overfortil = overforfra + 1
            overfortilAar = overforfraAar
            else
            overfortil = 1
            overfortilAar = overforfraAar + 1
            end if

            perSQLtil = " md = "& overfortil & ""
            perSQLtil2 = "MONTH(tdato) = " & overfortil
            end select

    
            call akttyper2009(2)
  


            '**** Overfør alle ressourcetimer der ligger i denne uge / md til næste uge / måned
            '**** Overfør kun diff
            '**** Hvis der allerede er timer i den næste periode lægges disse til
            '**** Uspecificeret hentes altid først, ORDER BY aktid, da de først overføres, da specifikke aktiviteter ellers vil blive overskrevet


                                    strSQL = "SELECT SUM(timer) AS restimer, medid, aktid, jobid, j.jobnr AS jobnr, rmd.id AS rid, uge, md, aar FROM ressourcer_md AS rmd "_
                                    &" LEFT JOIN job AS j ON (j.id = jobid) WHERE " & perSQL & " AND aar = "& overforfraAar &" AND jobstatus = 1 GROUP BY medid, jobid, aktid ORDER BY medid, jobid, aktid DESC"
                                    oRec.open strSQL, oConn, 3
                                    while not oRec.EOF 

                                    realtimerthisjobAkt = 0

                                    if lastjob <> oRec("jobid") then  
                                    aktidsThisJob = " AND taktivitetid <> 0"
                                    aktidsThisJobSQL = ""
                                    else
                                    aktidsThisJob = aktidsThisJob
                                    end if

                                    'response.Write strSQL & "<br>"
                                    'response.flush

                                        '*** Henter timer og beregner diff

                                        if oRec("aktid") <> 0 then
                                        aktKri = " AND taktivitetid = " & oRec("aktid")
                                        aktidsThisJob = aktidsThisJob & " AND taktivitetid <> " & oRec("aktid") 
                                        aktidsThisJobSQL = ""
                                        else
                                        aktKri = ""
                                        aktidsThisJobSQL = aktidsThisJob
                                        end if
                                            
                                                
                                                diffTimer = oRec("restimer")
                                                realtimerthisjobAkt = 0
                                                
                                                strSQLtimer = "SELECT SUM(timer) AS realtimertot FROM timer WHERE tjobnr = "& oRec("jobnr") &" AND " & perSQL2 &" AND YEAR(tdato) = "& overforfraAar &" AND tmnr = "& oRec("medid") & aktKri & " "& aktidsThisJobSQL &" AND ("& aty_sql_realhours &") GROUP BY tmnr"
                                                
                                                'response.write strSQLtimer & "<br>"
                                                'response.flush

                                                oRec2.open strSQLtimer, oConn, 3
                                                if not oRec2.EOF then

                                                realtimerthisjobAkt = oRec2("realtimertot")
                                                diffTimer = (diffTimer - realtimerthisjobAkt)
                                               
                                                end if
                                                oRec2.close


                                                '** Overfører kun hvis diffTimer > 0, dvs. hvis der mangler at blive udførst arbejde
                                                '** Finder tim i dne måned der skal overføres til
                                                nyeResTimer = 0    
                                                if cdbl(diffTimer) > 0 then
                                            

                                                 'response.write "MEdid: "& oRec("medid") & " jobnr:"& oRec("jobnr") &" aktid: "& oRec("aktid") & " timer: "& oRec("restimer") & " Real timer: "& realtimerthisjobAkt &" diffTimer: "& diffTimer &"<hr>"

                                                strSQL = "SELECT timer AS restimer, medid, aktid, jobid, rmd.id AS rid FROM ressourcer_md AS rmd "_
                                                &" WHERE " & perSQLtil & " AND aar = "& overfortilAar &" AND medid = "& oRec("medid") &" AND jobid = "& oRec("jobid") &" AND aktid = "& oRec("aktid") 
                                                oRec3.open strSQL, oConn, 3
                                                if not oRec3.EOF then 

                                                    nyeResTimer = (diffTimer + oRec3("restimer"))
                                                    nyeResTimer = replace(nyeResTimer, ",", ".")

                                                    strSQLUpdateResTimer = "UPDATE ressourcer_md SET timer = "& nyeResTimer &" WHERE id = "& oRec3("rid") 
                                                    
                                                    'response.write "strSQLUpdateResTimer: "& strSQLUpdateResTimer & "<br>"
                                                    'response.flush
                                                    oConn.execute(strSQLUpdateResTimer)

                                                
                                                else

                                                    nyeResTimer = diffTimer
                                                    nyeResTimer = replace(nyeResTimer, ",", ".")

                                                    select case periodeSel
                                                    case 3 'uge
                                                    ugetil = overfortil
                                                    
                                                    
                                                    weekdayy = datepart("w","1/1/"&overfortilAar, 2,2)                
                                            
                                                    if weekdayy > 4 then 'fredag, lør, søn == uge 53 dvs ekstra uge i diff
                                                    mdtil = dateAdd("w", ugetil + 1, "1/1/"&overfortilAar)
                                                    else
                                                    mdtil = dateAdd("w", ugetil, "1/1/"&overfortilAar)
                                                    end if                     
                                        

                                                    mdtil = month(mdtil)
                                                    proctil = 0
                                                    case else 'md
                                                    ugetil = datepart("ww", "15/"&overfortil&"/"&overfortilAar, 2,2) 'midt i måned
                                                    mdtil = overfortil            
                                                    proctil = 0
                                                    end select

                                                    strSQLInsertResTimer = "INSERT INTO ressourcer_md (timer, medid, aktid, jobid, md, aar, uge, proc) VALUES "_
                                                    &" ("& nyeResTimer &", "& oRec("medid") &", "& oRec("aktid") &", "& oRec("jobid") &", "& mdtil &", "& overfortilAar &","& ugetil &", "& proctil &")"
                                                    
                                                    'response.write "strSQLInsertResTimer: "& strSQLInsertResTimer & "<br>"
                                    
                                                    oConn.execute(strSQLInsertResTimer)

                                                end if 
                                                oRec3.close  



                                                    '*** Nedskriver den måned der overføres fra
                                                    nedskrivResTimer = (realtimerthisjobAkt)
                                                    nedskrivResTimer = replace(nedskrivResTimer, ",", ".")               
                                        
                                    
                                                   
                                    
                                                    if cint(realtimerthisjobAkt) > 0 then          

                                                   

                                                    strSQLUpdateFraResTimer = "UPDATE ressourcer_md SET timer = ("& nedskrivResTimer &") WHERE id = "& oRec("rid") 

                                                   
                                                    else

                                                    strSQLUpdateFraResTimer = "UPDATE ressourcer_md SET timer = 0 WHERE id = "& oRec("rid") 

                                                    end if


                                                     'renser ud i de andre uger ved månedsoverførsel
                                                    select case periodeSel
                                                    case 6,12 'måned
                                                    strSQLUpdateFraResTimerDel = "DELETE FROM ressourcer_md "_
                                                    &" WHERE " & perSQL & " AND aar = "& overforfraAar &" AND medid = "& oRec("medid") &" AND jobid = "& oRec("jobid") &" AND aktid = "& oRec("aktid") &" AND id <> "& oRec("rid")
                                                
                                                    'response.write "<br>strSQLUpdateFraResTimerDel" & strSQLUpdateFraResTimerDel
                                                    

                                                    oConn.execute(strSQLUpdateFraResTimerDel)

                                                  
                                                    end select

                                                    

                                                    oConn.execute(strSQLUpdateFraResTimer)


                                                end if 'Difftimer > 0

                                    lastjob = oRec("jobid")

                                    oRec.movenext
                                    wend
                                    oRec.close




        end if


        'response.write "<br><br><br>HER"
        'response.write "periodeSel: " & periodeSel
        'response.write request("FM_overforfra")
        'response.write request("FM_overfor")



        strLink = "ressource_belaeg_jbpla.asp?FM_progrp="&progrp&"&FM_medarb="&thisMiduse&"&FM_medarb_hidden="&thisMiduse&"&showonejob="&showonejob&"&id="&id&"&periodeselect="&periodeSel&"&jobselect="&jobnavn&"&mdselect="&request("mdselect")&"&aarselect="&request("aarselect")&"&selectplus="&selectplus&"&selectminus="&selectminus&"&FM_vis_job_u_fc="& vis_job_u_fc &"&FM_vis_job_fc_neg="&vis_job_fc_neg&"&FM_hideEmptyEmplyLines="&hideEmptyEmplyLines&"&FM_vis_simpel="&vis_simpel&"&FM_vis_kolonne_simpel="&vis_kolonne_simpel&"&FM_vis_kolonne_simpel_akt="&vis_kolonne_simpel_akt&"&FM_sog="& sogTxt 
	
        response.write "<br><br><div style=""padding:20px;"">"& resbelaeg_txt_132 &"<br><a href="& strLink &">"&resbelaeg_txt_133&" >></a></div>"                                
        response.end                           
        end if



        if func = "opduge" then
	
	    strSQL = "SELECT md, aar, id FROM ressourcer_md WHERE id <> 0"
	    oRec.open strSQL, oConn, 3
	    While not oRec.EOF
	            
	            thisdate = "1/"& oRec("md")&"/"&oRec("aar")
                call thisWeekNo53_fn(thisdate)
	            thisWeek = thisWeekNo53 'datepart("ww", thisdate, 2,2)
	            
	            
	            strSQLu = "UPDATE ressourcer_md SET uge = "& thisWeek & " WHERE id = "& oRec("id")
	            'Response.write "<b>"& thisdate &"</b>: "& strSQLu & "<br>"
	            'Response.flush
	            'oConn.execute strSQLu
	            
	    
	    oRec.movenext
	    wend 
	    oRec.close
	
	
	    'Response.Write "ok"
	     Response.end
	
	    end if






        intMid = 0
	    intJid = 0
	
	    startdato = jobStartKri
	    slutdato =  jobSlutKri
	
	
	    startdatoMD = month(jobStartKri)
	    startdatoYY = year(jobStartKri)
	    slutdatoMD =  month(jobSlutKri)
	    slutdatoYY = year(jobSlutKri)
	
	    '** Forgående 3 md. (totaler) ***
	    startdatoMDtot = month(dateadd("m", -3, jobStartKri))
	    startdatoYYtot = year(dateadd("m", -3, jobStartKri))
	    slutdatoMDtot =  month(dateadd("m", -1, jobStartKri))
	    slutdatoYYtot = year(dateadd("m", -1, jobStartKri))
	
	    timregStdato = year(dateadd("m", -3, jobStartKri))&"/"& month(dateadd("m", -3, jobStartKri)) &"/"& day(dateadd("m", -3, jobStartKri))
	    timregSldato = year(dateadd("m", -0, jobStartKri))&"/"& month(dateadd("m", -0, jobStartKri)) &"/"& day(dateadd("m", -0, jobStartKri))


        %>
        <script src="../timereg/inc/header_lysblaa_stat_20170808.js" type="text/javascript"></script>
        <script src="../timereg/inc/res_bl_jav6.js"></script>

       <%if media <> "print" then
           
           ldTop = 100
           ldLft = 250
           
         %>
	    <div id="load" style="position:absolute; display:; visibility:visible; top:<%=ldTop%>px; left:<%=ldLft%>px; width:300px; background-color:#ffffff; border:10px #6CAE1C solid; padding:20px; z-index:100000000;">
        <table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	    <img src="../ill/outzource_logo_200.gif" /><br />
	    &nbsp;&nbsp;&nbsp;&nbsp;<%=resbelaeg_txt_001 %>
	    </td><td align=right style="padding-right:40px;">
	    <img src="../inc/jquery/images/ajax-loader.gif" />
	    </td></tr></table>
	
	    </div>
	    <% end if
            
            
            Response.Flush    

            if media <> "print" AND media <> "eksport" AND media <> "chart" then
                call menu_2014
            end if
        %>

        <div class="container" style="width:1500px;">
            <div class="portlet">
                <h3 class="portlet-title"><u><%=resbelaeg_txt_002 %></u></h3>

                <div class="portlet-body">

                    <%if media <> "print" ANd media <> "eksport" AND media <> "chart" then %>

                    <form action="ressource_belaeg_jbpla.asp?showonejob=<%=showonejob%>&id=<%=id%>" method="post" name="mdselect" id="mdselect">
                    <div class="well">
                        <h4 class="panel-title-well"><%=resbelaeg_txt_146 %></h4>

                        <div class="row">
                            <%if ddcView <> 1 then %>
                            <div class="col-lg-3"><%=resbelaeg_txt_012 %>:</div>
                            <%else %>
                            <div class="col-lg-3">Medarbejdere:</div>
                            <%end if %>
                            <div class="col-lg-3"><%if ddcView <> 1 then %> <%=resbelaeg_txt_147 %><%else %>Job<%end if %>:</div>
                            <div class="col-lg-2"><%=resbelaeg_txt_148 %>:</div>
                            <div class="col-lg-2">
                                <%if ddcView <> 1 then %>
                                <%=resbelaeg_txt_035%>:
                                <%else %>
                                Sortering
                                <%end if %>
                            </div>
                            <div class="col-lg-2">
                                <%if ddcView <> 1 then %>
                                <%=resbelaeg_txt_047 %>:
                                <%end if %>
                            </div>
                        </div>

                        <div class="row">
                            <%if ddcView <> 1 then %>
                            <div class="col-lg-3"><input id="sogtxt" name="FM_sog" value="<%=sogTxt %>" type="text" class="form-control input-small" /></div>
                            <%else %>

                            <%
                            call fTeamleder(session("mid"), 0, 1)
                            'response.Write "herher " & strPrgids
                             
                            %>

                            <div class="col-lg-3">
                             <select class="form-control input-small" id="FM_medarb2" name="FM_medarb" multiple size="10">
                                    <%if strPrgids = "alle" then %>

                                        <%
                                            if thisMiduse = "0" then
                                            mallSel = "SELECTED"
                                            else
                                            mallSel = ""
                                            end if
                                        %>

                                        <option value="0" <%=mallSel %> >Alle</option>
                                   
                                        <%
                                        strSQL = "SELECT mnavn, mid FROM medarbejdere WHERE mansat = 1"
                                        oRec.open strSQL, oConn, 3
                                        while not oRec.EOF
                                            'if cdbl(ddcMid) = cdbl(oRec("mid")) then
                                            'medSEL = "SELECTED"
                                            'else
                                            'medSEL = 0
                                            'end if

                                            medSEL = ""
                                            for j = 0 TO UBOUND(intMids)
                                                if cdbl(intMids(j)) = cdbl(oRec("mid")) then
                                                    medSEL = "SELECTED"
                                                end if
                                            next

                                            response.Write "<option value='"&oRec("mid")&"' "&medSEL&" >"& oRec("mnavn") &"</option>"
                                        oRec.movenext
                                        wend
                                        oRec.close
                                        %>
                                    <%end if %>

                                    <%if strPrgids = "ingen" then %>
                                        <option value="<%=session("mid") %>"><%=session("user") %></option>
                                    <%end if %>

                                    <%if strPrgids <> "ingen" AND strPrgids <> "alle" AND len(strPrgids) <> 0 then
                                      
                                        response.Write "<option value='0'>Alle</option>"

                                        strSQL = "SELECT p.medarbejderid as medid, m.mnavn as mnavn FROM progrupperelationer p LEFT JOIN medarbejdere m ON (p.medarbejderid = m.mid) WHERE p.projektgruppeid IN ("&strPrgids&") GROUP BY p.medarbejderid ORDER BY m.mnavn"
                                        oRec.open strSQL, oConn, 3
                                        while not oRec.EOF

                                            medSEL = ""
                                            for j = 0 TO UBOUND(intMids)                                               
                                                if cdbl(intMids(j)) = cdbl(oRec("medid")) then
                                                    medSEL = "SELECTED"
                                                end if
                                            next

                                            response.Write "<option value="&oRec("medid")&" "& medSEL &">"&oRec("mnavn")&"</option>"
                                        oRec.movenext
                                        wend
                                        oRec.close
                                    end if %>
                                 </select>                     
                            </div>
                            <%end if 'ddcView %>
                            <!-- Job selection -->
                            <div class="col-lg-3">
                                <%
                                if ddcView = 1 then
                                    jobanskri = " AND (jobans1 = "& session("mid") &" OR jobans2 = "& session("mid") &" OR jobans3 = "& session("mid") & " OR jobans4 = "& session("mid") &" OR jobans5 = "& session("mid") & ")"
                                else
                                    jobanskri = ""
                                end if
                                %>

                                <%if ddcView = 1 then
                                    jobSELmultiple = "multiple"    
                                    jobSELsize = "10"
                                else
                                    jobSELmultiple = ""
                                    jobSELsize = "1"
                                end if %>
                                <select class="form-control input-small" name="FM_jobsel" id="FM_jobsel" <%=jobSELmultiple %> size="<%=jobSELsize %>" >
                                    <%
		                                strSQL = "SELECT jobnavn, jobnr, j.id, k.kkundenavn, k.kkundenr, k.kid"_
		                                &" FROM job j "_
		                                &" LEFT JOIN kunder k ON (k.kid = j.jobknr) "_
	                                    &" WHERE jobstatus = 1 "& jobDatoKri &" AND j.fakturerbart = 1 AND j.risiko > -1 "& jobanskri &" ORDER BY k.kkundenavn, j.jobnavn"
		
		                                '3 = tilbud
                                        'if lto = "hvk_bbb" AND session("mid") = 1 then
		                                ''Response.write strSQL
		                                'Response.flush
		                                'end if

                                        if ddcView <> 1 then
                                            if cint(jobidsel) = 0 then
                                            jallSel = "SELECTED"
                                            else
                                            jallSel = ""
                                            end if
                                        else
                                            if jobidsel = "0" then
                                            jallSel = "SELECTED"
                                            else
                                            jallSel = ""
                                            end if
                                        end if

                                    %>
                                    <option value="0" <%=jallSel %>><%=resbelaeg_txt_006 %></option>
                                    <%				
				                    oRec.open strSQL, oConn, 3
                                    k = 0
				                    while not oRec.EOF
				
                                        if ddcView <> 1 then
				                            if cint(jobidsel) = cint(oRec("id")) then
				                            isSelected = "SELECTED"
				                            else
				                            isSelected = ""
				                            end if
                                        else
                                            isSelected = ""
                                            for j = 0 TO UBOUND(jobids)
                                                if cdbl(jobids(j)) = oRec("id") then
                                                    isSelected = "SELECTED"
                                                end if
                                            next
                                        end if

                                    if cdbl(lastKid) <> oRec("kid") then
                                    %>
                                    <option value="<%=oRec("id")%>" disabled></option>
                                    <option value="<%=oRec("id")%>" disabled><%=oRec("kkundenavn")%>&nbsp;(<%=oRec("kkundenr")%>)</option>
                                    <%
				                    end if
				
				                    %>
				                    <option value="<%=oRec("id")%>" <%=isSelected%>><%=oRec("jobnavn")%>&nbsp;(<%=oRec("jobnr")%>) <%=aftnavnognr %></option>
				                    <%

                                    lastKid = oRec("kid")
                                    k = k + 1
				                    oRec.movenext
				                    wend
				                    oRec.close
				
				                    %>
                                    
                                </select>
                                <%if ddcView <> 1 then %>
                                <input id="viskunjobiper0" name="viskunjobiper" value="0" type="radio" <%=viskunjobiper0 %> onclick="submit();" /> <%=resbelaeg_txt_007 %> <b><%=resbelaeg_txt_008 %></b> <%=resbelaeg_txt_009 %><br />    
		                        <input id="viskunjobiper1" name="viskunjobiper" value="1" type="radio" <%=viskunjobiper1 %> onclick="submit();" /> <%=resbelaeg_txt_007 %> <b><%=resbelaeg_txt_010 %></b> <%=resbelaeg_txt_011 %>
                                <%end if %>

                            </div>

                            <div class="col-lg-2">

                                <%
                                monththisNumber = month("2018-"&monththis&"-01")
 
                                select case monththisNumber
                                    case 1
                                        monththisTxt = resbelaeg_txt_023
                                    case 2
                                        monththisTxt = resbelaeg_txt_024
                                    case 3
                                        monththisTxt = resbelaeg_txt_025
                                    case 4
                                        monththisTxt = resbelaeg_txt_026
                                    case 5
                                        monththisTxt = resbelaeg_txt_027
                                    case 6
                                        monththisTxt = resbelaeg_txt_028
                                    case 7
                                        monththisTxt = resbelaeg_txt_029
                                    case 8
                                        monththisTxt = resbelaeg_txt_030
                                    case 9
                                        monththisTxt = resbelaeg_txt_031
                                    case 10
                                        monththisTxt = resbelaeg_txt_032
                                    case 11
                                        monththisTxt = resbelaeg_txt_033
                                    case 12
                                        monththisTxt = resbelaeg_txt_034
                                end select

                                %>

                                <select name="mdselect" class="form-control input-small" style="display:inline-block; width:49%;">
	                                <option value="<%=monththis%>" SELECTED><%=monththisTxt %></option>
	                                <option value="1"><%=resbelaeg_txt_023 %></option>
	                                <option value="2"><%=resbelaeg_txt_024 %></option>
	                                <option value="3"><%=resbelaeg_txt_025 %></option>
	                                <option value="4"><%=resbelaeg_txt_026 %></option>
	                                <option value="5"><%=resbelaeg_txt_027 %></option>
	                                <option value="6"><%=resbelaeg_txt_028 %></option>
	                                <option value="7"><%=resbelaeg_txt_029 %></option>
	                                <option value="8"><%=resbelaeg_txt_030 %></option>
	                                <option value="9"><%=resbelaeg_txt_031 %></option>
	                                <option value="10"><%=resbelaeg_txt_032 %></option>
	                                <option value="11"><%=resbelaeg_txt_033 %></option>
	                                <option value="12"><%=resbelaeg_txt_034 %></option>
	                            </select>
                                <select name="aarselect" class="form-control input-small" style="display:inline-block; width:49%;">
	                                <option value="<%=yearthis%>" SELECTED><%=yearthis%></option>

                                    <%for yThis = 0 to 30  

                                    yUse = 2007 + yThis
    
                                    %>
                                    <option value="<%=yUse %>"><%=yUse %></option>
                                    <%
                                    next
                                    %> 
	                            </select>
                            </div>
                            <div class="col-lg-2">
                                <%if ddcView <> 1 then %>
                                <select name="periodeselect" class="form-control input-small" onchange="submit()">
	                                <option value="<%=periodeSel%>" SELECTED><%=periodeShow%></option>
	                                <!--<option value="1">1 Måned (dag/dag)</option>-->
	                                <option value="3"><%=resbelaeg_txt_036 %></option>
	                                <!--<option value="4">4 Måneder frem (måned/måned)</option>-->
                                    <%select case lto 
                                        case "kejd_pb"
                                        case else %>
	                                <option value="12"><%=resbelaeg_txt_037 %></option>
	                                <option value="6"><%=resbelaeg_txt_038 %></option>
                                        <%end select %>
    
                                </select>
                                <%else %>
                                <select name="FM_sortering" id="FM_sortering" class="form-control input-small">
                                    <option value="0" <%=sortJOB %>>Job</option>
                                    <option value="1" <%=sortMED %>>Medarbejdere</option>
                                </select>
                                <%end if %>
                            </div>

                            <%if ddcView = 1 then %>
                            <div class="col-lg-2"><input type="checkbox" value="1" id="FM_vis_job_u_fc" name="FM_vis_job_u_fc" <%=vis_job_u_fcCHK %> /> Ignoere forecast</div>
                            <%end if %>

                            <div class="col-lg-2">
                                <%if ddcView <> 1 then %>
                                <select class="form-control input-small" name="FM_showasproc" onchange="submit();"><option value="0" <%=showasproc0Sel %>><%=resbelaeg_txt_045 %></option><option value="1" <%=showasproc1Sel %>><%=resbelaeg_txt_046 %></option></select>
                                <%end if %>
                            </div>

                        </div>

                        <br />

                        <div class="row">
                            <div class="col-lg-3">
                                <%if ddcView <> 1 then %>
                                <%=resbelaeg_txt_003 %>:
                                <%end if %>
                            </div>
                            <div class="col-lg-3">
                                <%if ddcView <> 1 then %>
                                <%=resbelaeg_txt_004 %>:
                                <%end if %>
                            </div>
                            <div class="col-lg-2">
                                <%if ddcView <> 1 then %>
                                <%=resbelaeg_txt_015 %>:
                                <%end if %>
                            </div>
                            <div class="col-lg-2">
                                <%if ddcView <> 1 then %>
                                <%=resbelaeg_txt_200 %>:
                                <%end if %>
                            </div>
                            <div class="col-lg-2">
                                <%if ddcView <> 1 then %>
                                <%=resbelaeg_txt_161 %>:
                                <%end if %>
                            </div>
                        </div>

                        <div class="row">
                            <%if ddcView <> 1 then %>
                            <%call progrpmedarb_2018 %>
                            <%end if %>

                            <div class="col-lg-2">
                                <%if ddcView <> 1 then %>
                                <input type="checkbox" value="1" id="FM_vis_job_fc_neg" name="FM_vis_job_fc_neg" <%=vis_job_fc_negCHK%> /><%=" " & resbelaeg_txt_016 %> <br />
                                <input type="checkbox" value="1" id="FM_vis_job_u_fc" name="FM_vis_job_u_fc" <%=vis_job_u_fcCHK %> /><%=" " & resbelaeg_txt_017 %>
                                <%call positiv_aktivering_akt_fn()
                                  if cint(pa_aktlist) = 1 then %>
                                  <br /> <span style="color:#999999;">(<%=resbelaeg_txt_019 & " " %>+/-5<%=" " & resbelaeg_txt_020 %>)</span> 
                                 <%end if %>
                                <br />
                                <input type="checkbox" value="1" id="Checkbox1" name="FM_hideEmptyEmplyLines" <%=hideEmptyEmplyLinesCHK %> /><%=" " & resbelaeg_txt_162 %>
                                <%else %>
                                <input type="hidden" value="0" name="FM_vis_job_fc_neg" />
                               <!-- <input type="hidden" value="0" name="FM_vis_job_u_fc" /> -->
                                <input type="hidden" value="0" name="FM_hideEmptyEmplyLines" />
                                <%end if %>
                            </div>

                            <div class="col-lg-2">

                                <%if ddcView <> 1 then %>
                                <input type="checkbox" value="1" id="Checkbox2" <%=vis_kolonne_simpelDisabled %> name="FM_vis_kolonne_simpel" <%=vis_kolonne_simpelCHK %> /><%=" " & resbelaeg_txt_201 %> <br />
                                <%else %>
                                <input type="hidden" name="FM_vis_kolonne_simpel" value="0" />
                                <%end if %>

                                <%if ddcView <> 1 then %>
                                <input type="checkbox" value="1" id="Checkbox3" <%=vis_kolonne_simpel_aktDisabled %> name="FM_vis_kolonne_simpel_akt" <%=vis_kolonne_simpel_aktCHK %> /><%=" " & resbelaeg_txt_202 %>
                                <%else %>
                                <input type="hidden" name="FM_vis_kolonne_simpel_akt" value="1" />
                                <%end if %>
                            </div>

                            <div class="col-lg-2">
                                <%if ddcView <> 1 then %>
                                <select class="form-control input-small" id="FM_vis_simpel" name="FM_vis_simpel" onchange="submit();">
                                    <option value="0" <%=vis_simpelCHK0 %>><%=resbelaeg_txt_042 %></option>
                                    <option value="1" <%=vis_simpelCHK1 %>><%=resbelaeg_txt_043 %></option>
                                    <option value="2" <%=vis_simpelCHK2 %>><%=resbelaeg_txt_044 %></option>
                                </select>
                                <%else %>
                                <input type="hidden" name="FM_vis_simpel" value="0" />
                                <%end if %>
                            </div>

                        </div>

                        <div class="row">
                             <div class="col-lg-12">
                                 <%if ddcView = 1 then %>
                                <span id="loadfilter" class="btn btn-secondary btn-sm pull-right"><b><%=resbelaeg_txt_163 %></b></span>
                                 <%else %>
                                 <button type="submit" class="btn btn-secondary btn-sm pull-right"><b><%=resbelaeg_txt_163 %></b></button>
                                 <%end if %>
                             </div>
                        </div>

                    </div>

                    </form>
                    <%else %>

                        <input type="hidden" id="FM_vis_simpel" value="<%=vis_simpel %>" />

                    <%end if %>

                    <%
                        aarsRamme = datepart("yyyy",startdato, 2,2)

                        if cint(visRamme) = 1 then


                            select case periodeSel
                            case 3
                            tWdth = 1794
                            case 12
                            tWdth = 1794
                            case else
                            tWdth = 1694
	                        end select


                        else

                            select case periodeSel
                            case 3
                            tWdth = 1294
                            case 12
                            tWdth = 1394
                            case else
                            tWdth = 1194
	                        end select

                        end if



                        '**** Akt typer der er med i ddagligt timeregnskab ***'
                        aty_sql_realhoursAKt = replace(aty_sql_realhours, "tfaktim", "fakturerbar")
	
	                    isArsNormWrtEksp = 0
	                    skrivCsvFil = 1 
	                    public csvTxtTop, csvTxt
	                    public csvDatoerMD(35), csvDatoerAR(35), csvDatoerUge(35)
   
	                    public lastxmid
	                    lastxmid = 0
	
	                    mids = 1150
	                    dt = 300 '1000
	                    ji = 6

	                    dim midjiddato 
	                    redim midjiddato(mids, dt, ji) 

   
                        dim n_arr
                        redim n_arr(mids)
	
	                    antalX = 6000
	                    xId = 25
	
	                    'Response.flush
	
	                    public medarbKundeoplysX
	                    redim medarbKundeoplysX(antalX, xId) 
	
	                    'dim testarr
	                    '***  Md, Aar, Jobid, Mid **
	                    'select case lto
	                    'case "bminds", "execon"
	                    'redim testarr(53,12,44,900,110)
	                    'case else
	                    'redim testarr(53,12,500,40)
	                    'end select
	
	                    '*** Unique job id brugs i array, så hvis de når job id nr 901 fejler siden. **'
	
	                    lastYear = 0
	                    lastWeek = 0
	                    lastMid = 0
	                    lastJid = 0
                        lastJid_Aid = 0
	                    'public n
                        n = 0
                        m = 0
	                    'wrtNewRow = 0
	
	                    lastmedarbId = 0
	                    totFC_job_med_IsWrt = "##-1##"
	
   
	                    if cint(startdatoYY) = cint(slutDatoYY)  then
	                    orandval = " AND "
	                    else
	                    orandval = " OR "
	                    end if
	
                        public jobisWrt, jobIdisWrt
                        redim  jobisWrt(4000), jobIdisWrt(4000) 
	
                        medarbwrt = "#0#"
                        fc_lastjobnrWrt = 0
                        fc_lastaktidWrt = 0
	
	                    '************************************************************************
	                    '*** Main SQL ***
	                    '************************************************************************
	                    't = 0 alle med ress. timer i forecast
	                    't = 1 alle med realiserede timer og som ikke er med i forecast
	                    public x
	                    x = 0
	                    't1first = 1


                        tStDato = year(jobStartKri) &"-"& month(jobStartKri) & "-1" 
	                    tSlDato = jobSlutKri
	
                        tStDatoJobUforecast = dateAdd("m", - 6, jobStartKri)
                        tStDatoJobUforecast = year(tStDatoJobUforecast) & "-" & month(tStDatoJobUforecast) & "-1"
	
                        tSlDatoJobUforecast = dateAdd("m", 1, now)
                        tSlDatoJobUforecast = year(tSlDatoJobUforecast) &"-"& month(tSlDatoJobUforecast) & "-" & day(tSlDatoJobUforecast)

                        strSQL = "SELECT jobnavn, m.mnavn, m.mnr, j.id AS jobid, j.jobnr, j.jobstatus, m.medarbejdertype, j.jo_bruttooms, j.jo_udgifter_ulev, j.stade_tim_proc, j.budgettimer, j.restestimat, "_
                        &" m.mid, a.navn AS aktnavn, "
                        strSQL = strSQL & "sum(rmd.timer) AS restimer, SUM(rmd.proc) AS proc, rmd.md AS resmd, rmd.aar AS resaar, rmd.id AS resid, rmd.uge AS resuge, rmd.aktid AS aktid, "
	                    strSQL = strSQL & " k.kkundenavn, k.kkundenr, m.init, m.forecaststamp, jobans1, jobans2, "_
	                    &" m2.mnavn AS m2mnavn, m2.init AS m2init, m2.mnr AS m2mnr, m3.mnavn AS m3mnavn, m3.init AS m3init, m3.mnr AS m3mnr "_
                        &" FROM medarbejdere m "
	
                        strSQL = strSQL &" LEFT JOIN ressourcer_md rmd ON (rmd.medid = m.mid AND "_
	                    &" ((rmd.md >= "& startdatoMD &" AND rmd.aar = "& startdatoYY &") "& orandval &" (rmd.md <= "& slutdatoMD &" AND rmd.aar = "& slutdatoYY &")))"_
	                    &" LEFT JOIN job j ON (j.id = rmd.jobid)"

                        'tu.jobid
                        '&" LEFT JOIN timereg_usejob AS tu ON (tu.medarb = m.mid)"_

	                    strSQL = strSQL &" LEFT JOIN kunder k ON (kid = j.jobknr)"_
	                    &" LEFT JOIN medarbejdere m2 ON (m2.mid = j.jobans1)"_
	                    &" LEFT JOIN medarbejdere m3 ON (m3.mid = j.jobans2)"_
                        &" LEFT JOIN aktiviteter a ON (a.id = rmd.aktid)"

                        strSQL = strSQL &" WHERE "& medarbSQlKri &" "& deakSQL & "" 'strMedidSQL 
	
                        if ddcView = 1 then
                            if request("FM_medarb") = "0" then
	                        strSQL = strSQL &" m.mid <> 0 AND "& jobSQLkri &""
                            else
                            strSQL = strSQL &" AND m.mid <> 0 AND "& jobSQLkri &""
                            end if
                        else
                            strSQL = strSQL &" AND m.mid <> 0 AND "& jobSQLkri &"" ' AND j.jobnavn IS NOT NULL"                        
                        end if

	                    strSQL = strSQL &" AND (j.risiko > -1 OR j.risiko = -3)"
    
                        strSQL = strSQL &" GROUP BY m.mid, rmd.jobid, rmd.aktid, rmd.aar, rmd.md"
            
                                if periodeSel = 3 then
                                strSQL = strSQL &", rmd.uge"
                                end if

                        strSQL = strSQL &" ORDER BY m.mid, k.kkundenavn, rmd.jobid, rmd.aktid DESC, a.fase, a.sortorder, a.navn" 'rmd.aktid
   
	                    strSQL = strSQL &", rmd.aar, rmd.md"

                        if periodeSel = 3 then
                        strSQL = strSQL &", rmd.uge"
                        end if

                        'AND (m.mansat <> 2 AND m.mansat <> 3)
                        'if lto = "essens" then
	                    'Response.write strSQL &"<br><br>"
	                    'Response.flush
                        'end if
	                    'response.Write "23 " & strSQL
	                    oRec.open strSQL, oConn, 3 

                        while not oRec.EOF 

                        if x = 0 then
                            jobisWrt(oRec("mid")) = " AND (t.tjobnr <> '0' "
                            jobIdisWrt(oRec("mid")) = " AND (jobid <> '0' "
                        end if


                        '** Afslutter forrige og finder job med timer på udenne forecast pr. medarb **'
                        if lastMid <> oRec("mid") AND x <> 0 then
        
       
                            lastJid = 0
                            n_arr(lastMid) = n-1 

                            '*** Henter job uden forecast ***'
                            if cint(vis_job_u_fc) = 1 then
                            call jobmedtimeruforecast(lastMid, 1)
        
                            jobisWrt(oRec("mid")) = " AND (t.tjobnr <> '0' "
                            jobIdisWrt(oRec("mid")) = " AND (jobid <> '0' "
                            medarbwrt = medarbwrt & ",#" & lastMid & "#"

                            end if

                            n = 0
     

                        end if


                        'if oRec("jobnr") <> "" AND instr(jobisWrt(oRec("mid")), " AND t.tjobnr <> "& oRec("jobnr")) = 0 then
                        if lastJid <> oRec("jobid") then 
	                    jobisWrt(oRec("mid")) = jobisWrt(oRec("mid")) & " AND t.tjobnr <> '"& oRec("jobnr") &"'"
                        jobIdisWrt(oRec("mid")) = jobIdisWrt(oRec("mid")) & " AND jobid <> '"& oRec("jobid") &"'" 
	                    end if
    
   
                        if len(trim(oRec("resmd"))) <> "" then
	                    showresdato = "1/"&oRec("resmd")&"/"&oRec("resaar")
	                    showresuge = oRec("resuge")
                        else
                        showresdato = "1/1/2002"
	                    showresuge = 1
                        end if


                        medarbKundeoplysX(x, 0) = oRec("jobid")
	                    medarbKundeoplysX(x, 1) = oRec("jobnr")
                        medarbKundeoplysX(x, 2) = oRec("jobnavn")
                        medarbKundeoplysX(x, 3) = oRec("mid")
                        medarbKundeoplysX(x, 4) = oRec("mnavn")
    
	                    medarbKundeoplysX(x, 5) = oRec("kkundenavn")
	                    medarbKundeoplysX(x, 6) = oRec("kkundenr")
	                    medarbKundeoplysX(x, 7) = oRec("mnr")
	                    medarbKundeoplysX(x, 8) = oRec("init")


                        'if t = 0 then
	                    medarbKundeoplysX(x, 9) = oRec("forecaststamp")

                        'if cint(showasproc) = 1 then
                        'medarbKundeoplysX(x, 10) = oRec("proc")
                        'else
	                    medarbKundeoplysX(x, 10) = oRec("restimer")
	                    'end if

                        '** procent altid indtastet i forhold til normtid i den valgete periode. **'
                        medarbKundeoplysX(x, 23) = oRec("proc") 
                        'medarbKundeoplysX(x, 24) = oRec("mdorweek")
   
                                'response.write "JOB: "& oRec("jobnr") &" PROC: "& oRec("proc") & "<br>"


                        resid = oRec("resid")
	                    'else
	                    'medarbKundeoplysX(x, 9) = "2002-01-01" 'oRec("forecaststamp")
	                    'medarbKundeoplysX(x, 10) = 0
	                    'resid = -1
	                    'end if

                        medarbKundeoplysX(x, 11) = oRec("jobans1")
	                    medarbKundeoplysX(x, 12) = oRec("jobans2")
	
	                    if len(trim(oRec("m2mnavn"))) <> 0 then
                        medarbKundeoplysX(x, 13) = oRec("m2mnavn") &" ("&oRec("m2mnr")&")"
                        medarbKundeoplysX(x, 14) = oRec("m2init") 
                        else
                        medarbKundeoplysX(x, 13) = ""
                        medarbKundeoplysX(x, 14) = ""
                        end if
    
                        if len(trim(oRec("m3mnavn"))) <> 0 then
                        medarbKundeoplysX(x, 15) = oRec("m3mnavn") &" ("&oRec("m3mnr")&")"
                        medarbKundeoplysX(x, 16) = oRec("m3init")
                        else
                        medarbKundeoplysX(x, 15) = ""
                        medarbKundeoplysX(x, 16) = ""
                        end if 

                        medarbKundeoplysX(x, 17) = oRec("aktid")
    
                         medarbKundeoplysX(x, 18) = oRec("aktnavn")
                         'Response.Write medarbKundeoplysX(x, 0) &"_"& medarbKundeoplysX(x, 17) & "<br>"


                        medarbKundeoplysX(x, 19) = oRec("jo_bruttooms")
                        medarbKundeoplysX(x, 20) = oRec("budgettimer")
                        medarbKundeoplysX(x, 21) = oRec("stade_tim_proc")
                        medarbKundeoplysX(x, 22) = oRec("restestimat")

                        select case oRec("jobstatus")
                            case 0
                            medarbKundeoplysX(x, 25) = tsa_txt_020
                            case 1
                            medarbKundeoplysX(x, 25) = tsa_txt_019
                            case 2
                            medarbKundeoplysX(x, 25) = resbelaeg_txt_164
                            case 3
                            medarbKundeoplysX(x, 25) = resbelaeg_txt_165
                            case 4
                            medarbKundeoplysX(x, 25) = resbelaeg_txt_166
                        end select


                        if len(resid) <> 0 AND len(oRec("jobid")) <> 0 then
		         		         
		                 'if periodeSel = 3 OR periodeSel = 6 OR periodeSel = 12 then   
		            
				     
					                'Response.Write "N: " & n & "<br>"
					                'Response.flush
					                midjiddato(oRec("mid"),n, 0) = medarbKundeoplysX(x, 10) 'oRec("restimer") 'showrestimer(x) 
                                    midjiddato(oRec("mid"),n, 6) = medarbKundeoplysX(x, 23) 'procent
					                midjiddato(oRec("mid"),n, 1) = oRec("jobid") 'useJobid(x) 
					                midjiddato(oRec("mid"),n, 2) = oRec("mid") 'xmid(x) 
					                midjiddato(oRec("mid"),n, 3) = showresdato 
					                midjiddato(oRec("mid"),n, 4) = showresuge
                                    midjiddato(oRec("mid"),n, 5) = oRec("aktid")
                           
				      
				        
				          ' else
				           '*** MD / MD tjkker ikke om uge passer **'
				            
				            
				     

                                    'if t = 0 then
					                'Response.Write "<br>dt: "& formatdatetime(showresdato, 2) &"n:"& n &" x:"& x &" arr: md: "& oRec("resmd") &",AAr: "& right(oRec("resaar"),2) &", Job: "& oRec("jobid") &", MId: "& oRec("mid") & ":Navn: "& medarbKundeoplysX(x, 4) &" Timer: "& medarbKundeoplysX(x, 10) &"<br>"
	                                'Response.Flush
        					        'end if
					                'testarr(oRec("resuge"),right(oRec("resaar"),2),oRec("jobid"),oRec("mid")) = oRec("restimer")
        	                
        	                
        	                
					        
					                'Response.Write "N: " & n & "<br>"
					                'Response.flush
					        '        midjiddato(oRec("mid"),n, 0) = medarbKundeoplysX(x, 10) 'timer 
                            '       midjiddato(oRec("mid"),n, 6) = medarbKundeoplysX(x, 23) 'procent

					        '        midjiddato(oRec("mid"),n, 1) = oRec("jobid") 'useJobid(x) 
					        '        midjiddato(oRec("mid"),n, 2) = oRec("mid") 'xmid(x) 
                                    'if t = 0 then
					                'midjiddato(oRec("mid"),n, 3) = "1/"&oRec("resmd")&"/"&oRec("resaar") 'showresdato 
                                    'else
                            '        midjiddato(oRec("mid"),n, 3) = showresdato 
                                    'end if
					        '        midjiddato(oRec("mid"),n, 4) = showresuge
                            '        midjiddato(oRec("mid"),n, 5) = oRec("aktid")

                        
				   
				           'end if

                        end if

                        lastWeek = showresuge	
			            lastYear = datepart("yyyy", showresdato) 
			            lastMonth = datepart("m", showresdato, 2,2) 
			            lastMid = oRec("mid")
			            lastJid = oRec("jobid")
                        lastJid_Aid = oRec("jobid") &"_"& oRec("aktid")
                        lastx = x


                         n = n + 1


                        'Response.Write "N:" & n & "<br>"
                        'n_arr(oRec("mid")) = n
	
	                    x = x + 1
	                    oRec.movenext
	                    wend
	                    oRec.close 


                        n_arr(lastMid) = n-1 


                        '**** Henter timer på job udne forecast på sidste medarbejer
                        '**** Kun hvis der kun er 1 medarb. med forecast på, så bliver den ikke hentet i loopet 

                        '*** Henter job uden forecast ***'
                        if cint(vis_job_u_fc) = 1 then

                         call jobmedtimeruforecast(lastMid, 2)
                         medarbwrt = medarbwrt & ",#" & lastMid & "#"


                        '*** Henter resten af e valgte medarbejdere uden forecast **'
	
   

                        for m = 0 to UBOUND(intMids)
            
                 

                            if instr(medarbwrt, "#"& intMids(m) &"#") = 0 then
                            jobisWrt(intMids(m)) = " AND (t.tjobnr <> '0'"
                            jobIdisWrt(intMids(m)) = " AND (jobid <> '0'"

                            'Response.Write medarbwrt & " " & intMids(m) & " // " & instr(medarbwrt, "#"& intMids(m) &"#") & "<br>" 
                            call jobmedtimeruforecast(intMids(m), 2)

                            '** Hvis der er valgt alle job skal denne være slået til, da loopet ellers finder medarbejdere dobbelt. ***'
                            'if jobidsel = 0 then
                            'medarbwrt = medarbwrt & "1538,#" & intMids(m) & "#"
                            'end if
            
                            end if
                        next

                        end if  'vis_job_u_fc ***'

                        '*** vis kolonner ****'
                        if cint(vis_kolonne_simpel) = 1 then
                        visRamme = 0
                        visStatus = 0
                        else
                        visRamme = 1
                        visStatus = 1
                        end if

                         '*** vis akt. kolonner ****'
                        if cint(vis_kolonne_simpel_akt) = 1 then
                        visAkt = 0
                        else
                        visAkt = 1
                        end if

                    %>
                    <form action="ressource_belaeg_jbpla.asp?func=redalle&FM_progrp=<%=progrp%>&FM_medarb=<%=thisMiduse%>&FM_medarb_hidden=<%=thisMiduse%>&showonejob=<%=showonejob%>&id=<%=id%>&FM_vis_job_u_fc=<%=vis_job_u_fc %>&FM_vis_job_fc_neg=<%=vis_job_fc_neg %>&FM_vis_simpel=<%=vis_simpel%>&FM_hideEmptyEmplyLines=<%=hideEmptyEmplyLines %>&FM_vis_kolonne_simpel=<%=vis_kolonne_simpel %>&FM_vis_kolonne_simpel_akt=<%=vis_kolonne_simpel_akt%>" method="post" name="indlas" id="indlas">
                        <%if media <> "print" AND media <> "eksport" AND media <> "chart" then%>
                    
                            <input id="Text1" name="FM_sog" value="<%=sogTxt %>" type="hidden"/>
                            <input type="hidden" name="mdselect" id="Hidden2" value="<%=monththis%>">
                            <input type="hidden" name="aarselect" id="aarselect" value="<%=yearthis%>">
                            <input type="hidden" name="periodeselect" id="periodeselect" value="<%=periodeSel%>">
                            <input type="hidden" name="jobselect" id="Hidden1" value="<%=jobnavn%>">

                            <%if ddcView = 1 then %>
                            <input type="hidden" id="FM_job_new" name="FM_jobsel" value="<%=request("FM_jobsel") %>" /> 
                            <%end if %>
                            <div class="row">
                                <input type="hidden" name="selectminus" id="selectminus" value="0">
	                            <input type="hidden" name="selectplus" id="selectplus" value="0">

                                <%
                                    monththisNumber = month("2018-"&monththis&"-01")
                                    select case monththisNumber
                                        case 1
                                            monthTxtAdd = resbelaeg_txt_024
                                            monthTxtSubtract = resbelaeg_txt_034
                                        case 2
                                            monthTxtAdd = resbelaeg_txt_025
                                            monthTxtSubtract = resbelaeg_txt_023
                                        case 3
                                            monthTxtAdd = resbelaeg_txt_026
                                            monthTxtSubtract = resbelaeg_txt_024
                                        case 4
                                            monthTxtAdd = resbelaeg_txt_027
                                            monthTxtSubtract = resbelaeg_txt_025
                                        case 5
                                            monthTxtAdd = resbelaeg_txt_028
                                            monthTxtSubtract = resbelaeg_txt_026
                                        case 6
                                            monthTxtAdd = resbelaeg_txt_029
                                            monthTxtSubtract = resbelaeg_txt_027
                                        case 7
                                            monthTxtAdd = resbelaeg_txt_030
                                            monthTxtSubtract = resbelaeg_txt_028
                                        case 8
                                            monthTxtAdd = resbelaeg_txt_031
                                            monthTxtSubtract = resbelaeg_txt_029
                                        case 9
                                            monthTxtAdd = resbelaeg_txt_032
                                            monthTxtSubtract = resbelaeg_txt_030
                                        case 10
                                            monthTxtAdd = resbelaeg_txt_033
                                            monthTxtSubtract = resbelaeg_txt_031
                                        case 11
                                            monthTxtAdd = resbelaeg_txt_034
                                            monthTxtSubtract = resbelaeg_txt_032
                                        case 12
                                            monthTxtAdd = resbelaeg_txt_023
                                            monthTxtSubtract = resbelaeg_txt_033
                                    end select
                                %>

                                <div class="col-lg-1"><input type="button" class="btn btn-default btn-sm" style="width:100%" name="submitplus" id="submitminus" value="<< <%=monthTxtSubtract %>" onClick="selminus()"></div>
                                <div class="col-lg-1"><input type="button" class="btn btn-default btn-sm" style="width:100%" name="submitplus" id="submitplus" value="<%=monthTxtAdd %> >>" onClick="seladd()"></div>
                               <!-- <div class="col-lg-10"><button type="submit" class="btn btn-success btn-sm pull-right"><b><%=resbelaeg_txt_049 %></b></button></div> -->
                            </div>
                    
                        <%end if %>

                        <%
                        select case periodeSel
                        'case 1
                            'numoffdaysorweeksinperiode = datediff("d", startdato, slutdato)
                        case 3
                            numoffdaysorweeksinperiode = datediff("ww", startdato, slutdato)
                        case 6, 12
                            numoffdaysorweeksinperiode = datediff("m", startdato, slutdato)
                        end select
                        %>

                        <input type="hidden" value="<%=numoffdaysorweeksinperiode %>" id="numoffdaysorweeksinperiode" />
                        <input type="hidden" value="<%=startdatoMD %>" id="startdatoMD" />
                        <input type="hidden" value="<%=startdatoYY %>" id="startdatoYY" />
                        <input type="hidden" value="<%=slutdatoMD %>" id="slutdatoMD" />
                        <input type="hidden" value="<%=slutdatoYY %>" id="slutdatoYY" />

                        <%if sortering <> 0 then %>
                        <table class="table table-bordered table-hover">
                            <!--<thead> -->

                                <%
                                    MedIdSQLKri = ""
                                    MedTotSQLKri = ""	
                                    'MedJobFcastSQLKri = ""

                                    
                                    fob = 0
                                    y = 0 
                                    nday = startdato
                                    weekrest = 0		
                                    antalxx = x
                                    xx = 0
                                    'foundone = 0
                                    daynowrpl = replace(daynow, "/", "-")
                                    lastMonth = monththis

                                    thisMTotRamme = 0

                                    antalJobAktlinier = 0
                                    antalJobAktMedlinier = 0
                                    antalJobAktlinierGrand = 0
                                    'antanlJoblinierprM = 0
                                    thisMedarbJoblist = ""

                                    lastAktJidExpVis_1 = "0_0_0"
                                    lastAktJidExpVis_1Total = "0_0_0"

                                    dagsdato = now

                                    if media <> "eksport" AND media <> "chart" then
                                        call datoeroverskrift(1)
                                    else
                                        call datoeroverskriftCSV(1)
                                    end if

                                    'public MedTotSQLKri
                                    'strjIdWrt = "0#," 


                                    'Response.write "HER: x: "& x &" antalx: "& antalxx
                                    'Response.flush

                                    x = 0
                                    for x = 0 to antalxx - 1 'UBOUND(medarbKundeoplysX, 1) 'medarbKundeoplysX(x, 0)

                                        '**** Ny overksift for de linier er ligger timer på men der ikke ligger forecast på ***'

                                        if lastxmid <> medarbKundeoplysX(x, 3) then	'xmid(x) 

	                                        tdcolspan = 5
	                                        strDageChkboxOne = "<input type=text class='form-control input-small' name=FM_sel_dage id=FM_sel_dage value="&daynowrpl&">&nbsp;dd-mm-aaaa"

                                            if x > 0 then				

                                                '*** Total pr. medarb linije ********************

                                                if media <> "eksport" AND media <> "chart" AND cint(vis_job_fc_neg) <> 1 then

                                                    'if ddcView <> 1 then
                                                    call jobtotalprmedarb
                                                    'end if

                                                    'if cint(jobidsel) = 0 OR ddcView = 1 then
                                                    if jobidsel = "0" OR ddcView = 1 then
                                                        'if cint(vis_simpel) <> 1 then
                                                        mTotthisMid = lastxmid
                                                        'if ddcView <> 1 then
	                                                    call medarbtotal
                                                        'end if
                                                        'end if
	                                                end if
	    
	                                            end if

                                                'if (cint(jobidsel) = 0 OR ddcView = 1) AND cint(vis_job_fc_neg) <> 1 then
                                                if (jobidsel = "0" OR ddcView = 1) AND cint(vis_job_fc_neg) <> 1 then
                    
                                                    if media = "chart" then
                                
                                                            call meStamdata(medarbKundeoplysX(x-1, 3))
                                                            csvTxtA = csvTxtA & meTxt 

                                                            for y = 0 TO numoffdaysorweeksinperiode
                                                            csvTxtA = csvTxtA & ";"& medarbTotalTimer(y)
                                                            medarbTotalTimer(y) = 0 
                                                            next

                                                            csvTxtA = csvTxtA & "xx99123sy#z"
                                                    end if

                                                end if

                                            end if 'x


                                            if media <> "eksport" AND media <> "chart" then 
        
                                                'Response.Write "#"&medarbwrt&"#, "& medarbKundeoplysX(x, 3) &"<br>"
	                                            'call medarbejderLinje(medarbKundeoplysX(x, 3),jobStartKri,datoInterval,rdimnb)
                                                medarbwrt = medarbwrt & ",#" & medarbKundeoplysX(x, 3) & "#"
         
	                                        end if


                                            if media <> "print" AND media <> "eksport" AND media <> "chart" then 'AND (instr(strjIdWrt, "#"&jobidSel&"#") = 0) then 
				    
				                                call tilfojNylinie(medarbKundeoplysX(x, 3), visRamme, visAktiv)
				   				    
				                            end if 'Print

                                            MedIdSQLKri = medarbKundeoplysX(x, 3)
	                                        csvTmnavn = medarbKundeoplysX(x, 4)
	                                        csvTmnr = medarbKundeoplysX(x, 7)
	                                        csvTinit = medarbKundeoplysX(x, 8)
	                                        csvTnormuge = ntimPer

                                        end if '*** lastXMid / MedarbejderID


                                        '***** Total pr. job indenfor medarbejder
                                        if lastJid <> medarbKundeoplysX(x, 0) AND len(medarbKundeoplysX(x, 0)) <> 0 AND fob = 1 AND cint(vis_job_fc_neg) <> 1 then
                                            'if ddcView <> 1 then
                                            call jobtotalprmedarb
                                            'end if
                                            'antanlJoblinierprM = 0
                                        end if


                                        '*** Tjekker om linje på medarb er skrevet **'
                                    if (lastJid_Aid <> medarbKundeoplysX(x, 0)&"_"& medarbKundeoplysX(x, 17) OR cdbl(lastxmid) <> cdbl(medarbKundeoplysX(x, 3))) AND len(medarbKundeoplysX(x, 0)) <> 0 then
		
                                        isArsNormWrtEksp = 0

                                        '**** Total forecast pr. medarb. pr. job / akt. ***'
                                        '**** Toal Real pr. medarb. pr. job / akt. ********'

                                        if inStr(totFC_job_med_IsWrt, "##"&x&"##") = 0 then
                        
                                             forcastTimerTotAktMedignPer = 0
                                             realTimerTotAktMedignPer = 0

                                            '*** Forecast, total ign. periode ***
                                            strSQLforcTotjobMid = "SELECT SUM(timer) AS sumTimerAktMedIgnPer FROM ressourcer_md WHERE medid = "& medarbKundeoplysX(x, 3) &" AND aktid = "& medarbKundeoplysX(x, 17) & " AND jobid = "& medarbKundeoplysX(x, 0) &" GROUP BY medid, jobid"
                                            'Response.write strSQLforcTotjobMid & "<br><br>"
                                            'Response.flush

                                            oRec3.open strSQLforcTotjobMid, oConn, 3
                                            if not oRec3.EOF then
                    
                                                forcastTimerTotAktMedignPer = oRec3("sumTimerAktMedIgnPer") 

                                            end if
                                            oRec3.close 

                 

                                            '*** Real. tot ign periode **'
                                            if medarbKundeoplysX(x, 17) <> 0 then
                                                aktSQLkri = " AND taktivitetid = "& medarbKundeoplysX(x, 17) 
                                            else
                                                aktSQLkri = ""
                                            end if

                                            strSQLRealTotjobMid = "SELECT SUM(timer) AS sumTimerAktMedIgnPer FROM timer WHERE tmnr = "& medarbKundeoplysX(x, 3) &" "& aktSQLkri &" AND tjobnr = '"& medarbKundeoplysX(x, 1) &"' AND ("& aty_sql_realhours &") GROUP BY tmnr, tjobnr"
                                            'Response.write "<hr>medarbKundeoplysX(x, 17):"& medarbKundeoplysX(x, 17)  &" "& medarbKundeoplysX(x, 4) &":" & strSQLRealTotjobMid & "<br>"
                        
                                            'Response.write strSQLRealTotjobMid & "<br>"
                                            'Response.flush

                                            oRec3.open strSQLRealTotjobMid, oConn, 3
                                            if not oRec3.EOF then
                    
                                                realTimerTotAktMedignPer = oRec3("sumTimerAktMedIgnPer") 

                                            end if
                                            oRec3.close 


                      
                      
                                            realTimerTotJobMedignPerGrand = realTimerTotJobMedignPerGrand + realTimerTotAktMedignPer
                                            realTimerTotJobMedignPer = realTimerTotAktMedignPer

                                            realTimIngPerTxt = "<span style=""font-size:9px; color:#999999;"">"& formatnumber(realTimerTotAktMedignPer, 2) &"</span>"
                        
                                            totFC_job_med_IsWrt = totFC_job_med_IsWrt & ",##"&x&"##"
                    
                                        end if

                                        '*******************************************************************
                                        '***************** Ny job linje     ********************************
                                        '*******************************************************************

                                        '*** Vis kun akt. / job hvor forecast er overskrddet 
                                        'z = 900
                                        'if z = 900 then

                                        if cint(vis_job_fc_neg) = 0 OR (cint(vis_job_fc_neg) = 1 AND (cdbl(realTimerTotAktMedignPer) > cdbl(forcastTimerTotAktMedignPer))) then
                                            '**efter medarbejder sum er skrevet           
                                            antalJobAktlinier = antalJobAktlinier + 1 
                                            'antalJobAktlinierPrJob = antalJobAktlinierPrJob + 1  

                                            if media <> "eksport" AND media <> "chart" then

                                                if media <> "print" then                           
                                                    call tilfojNyliniePaaJob(medarbKundeoplysX(x, 3), medarbKundeoplysX(x, 0))                      
                                                end if 'print

                                                if lastxmid <> medarbKundeoplysX(x, 3) AND ddcView = 1 then
                                                    %>
                                                    <tr id="tr_medarb_<%=medarbKundeoplysX(x,3)%>">
                                                        <%call medarbejderLinje(medarbKundeoplysX(x, 3),jobStartKri,datoInterval,rdimnb,1) %>

                                                        <%for y = 0 TO numoffdaysorweeksinperiode %>
                                                            <td style="text-align:right; width:60px;"><b><span id="monthTot_<%=medarbKundeoplysX(x,3)%>_<%=y %>">0,00</span></b></td>
                                                            <td style="text-align:right; width:60px;"><b><span id="monthTot_real_<%=medarbKundeoplysX(x,3)%>_<%=y %>">0,00</span></b></td>
                                                        <%next %>

                                                        <td style="text-align:right;"><b><span class="pertotal_fc" id="<%=medarbKundeoplysX(x,3) %>">0,00</span></b></td>
                                                        <td style="text-align:right;"><b><span class="pertotal_rl" id="<%=medarbKundeoplysX(x,3) %>">0,00</span></b></td>
                                                        <td style="text-align:right;"><b><span class="pertotal_saldo" id="<%=medarbKundeoplysX(x,3) %>">0,00</span></b></td>
                                                        <td style="width:15px; text-align:center;"><span style="color:#afafaf;" id="<%=medarbKundeoplysX(x, 3) %>" class="btn btn-sm btn-default fa fa-plus addField_job"></span></td>
                                                    </tr>
                                                    <%
                                                end if
                                        
                                                %>
                                                    <tr class="tr_medarb tr_medarb_<%=medarbKundeoplysX(x,3)%> tr_medarb_<%=medarbKundeoplysX(x,3)%>_<%=medarbKundeoplysX(x, 0)%>" style="visibility:hidden; display:none; background-color:white;">
                                                <%
                                                    if lastxmid <> medarbKundeoplysX(x, 3) AND ddcView <> 1 then
                                                        call medarbejderLinje(medarbKundeoplysX(x, 3),jobStartKri,datoInterval,rdimnb,1) 
                         
                                                        'lastxmid = medarbKundeoplysX(x, 3)
        
                                                    else
                                                    %>
                                                        <%if ddcView <> 1 then %>
                                                        <td style="vertical-align:top;">                        
                                                        <%
                                                        if vis_simpel = 2 then                                   
                                                            call meStamdata(medarbKundeoplysX(x, 3))
                                                        %>
                                                            <i><%=meNavn%></i>
                                                        <%end if %>  
                                                        </td>
                                                        <%end if %>
                                                    <%end if %>


                                                    <td style="vertical-align:top;">
                                                        <%            
                                                        '** SKAL Jobnavn og kundee navn skives igen (er det første linje på job pr. emdarb.)
                                                        if (lastJid <> medarbKundeoplysX(x, 0) OR lastxmid <> medarbKundeoplysX(x, 3) OR (fob = 0 AND len(medarbKundeoplysX(x, 0)) <> 0) OR ddcView = 1) then%>
                                                            <%fob = 1%>

                                                            <%if ddcView <> 1 then %>
                                                            <%=medarbKundeoplysX(x, 5)%> (<%=medarbKundeoplysX(x, 6)%>) <!--&nbsp;<span class="btpush" style="padding:1px; float:right; background-color:#CCCCCC; font-size:9px;" id="btpush_<%=medarbKundeoplysX(x, 3) %>_<%=medarbKundeoplysX(x, 0) %>_<%=x%>_<%=y%>">>>|</span>--><br />
                                                            <%end if %>

                                                            <%if ddcView <> 1 then %>
                                                            <b><%=medarbKundeoplysX(x, 2)%> (<%=medarbKundeoplysX(x, 1)%>)</b>  
                                                            <%else %>
                                                            <%=medarbKundeoplysX(x, 2)%> (<%=medarbKundeoplysX(x, 1)%>)
                                                            <%end if %>

                                                            <%
                                                                showJobst = 0   
                                                                select case lcase(medarbKundeoplysX(x, 25))
                                                                case "aktiv"
                                                                showJobst = 0
                                                                case "lukket"
                                                                jstaCol = "red"
                                                                showJobst = 1
                                                                case else
                                                                jstaCol = "#999999"
                                                                showJobst = 1
                                                                end select
                                
                                                                if cint(showJobst) = 1 then
                                                                %>
                                                                <span style="font-size:9px; color:<%=jstaCol%>;"><%=medarbKundeoplysX(x, 25) %></span>
                                                                <%end if%>
                                                                <br />
		
                                                        <%
                                                            if thisMedarbJoblist = "" then
                                                            thisMedarbJoblist = "<span id='sp_medarbjoblist_"&medarbKundeoplysX(x, 3)&"_"& medarbKundeoplysX(x, 0) &"' class='sp_medarbjoblist' style='color:#5582d2;'>" & medarbKundeoplysX(x, 2) & "</span>"
                                                            else
                                                            thisMedarbJoblist = thisMedarbJoblist &", <span id='sp_medarbjoblist_"&medarbKundeoplysX(x, 3)&"_"& medarbKundeoplysX(x, 0) &"' class='sp_medarbjoblist' style='color:#5582d2;'>"& medarbKundeoplysX(x, 2) & "</span>"
                                                            end if%>                    

                                                            <%if ddcView <> 1 then %>
		                                                    <span style="font-size:9px; color:#999999;">
		                                                    <%if len(trim(medarbKundeoplysX(x, 13))) <> 0 then %>
	                                                        <%=medarbKundeoplysX(x, 13) %>
		                                                    <%end if %>
		
		                                                    <%if len(trim(medarbKundeoplysX(x, 15))) <> 0 then %>
		                                                    , <%=medarbKundeoplysX(x, 15) %>
		                                                    <%end if %>
		                                                    </span>       
                                                            <%end if %>
                                                        <%
                                                        else
                    
              
                                                        end if %>

                                                        <%
                                                            '**** Aktivitet ***'
                                                            if cint(visAkt) = 0 AND ddcView <> 1 then 'viser ikke kolonne, men viser alligevel aktnavn hvis ikke linjen er uspec 

                                                            aktNavn = ""
                                                            if medarbKundeoplysX(x, 17) <> 0 then
                                                            strSQlAktnavn = "SELECT navn FROM aktiviteter WHERE id = "& medarbKundeoplysX(x, 17)
                            
                       
                                                            oRec6.open strSQlAktnavn, oConn, 3
                                                            if not oRec6.EOF then
                                                            aktNavn = oRec6("navn")
                                                            end if
                                                            oRec6.close


                            
                                                            %><br /><i><%=aktNavn %></i>
                                                            <%else %>
                                                            <i><%=resbelaeg_txt_050 %></i>
                                                            <%end if %>                

                                                        <%end if %>        
                                                    </td>     
                                                        

                                                    <!-- Ramme -->
                                                    <%if cint(visRamme) = 1 AND ddcView <> 1 then %>         
                                                   <td style="vertical-align:top;">
               
                                                      <%call arsRammeSub %>
               
                                                       <%if media <> "print" then %>
                                                       <input id="Text2" type="text" class="form-control input-small" name="FM_aarsramme_timer"  value="<%=formatnumber(arrsNormThis, 0) %>" />
                                                       <%else %>
                                                       <%=formatnumber(arrsNormThis, 0) %>
                                                       <%end if %>
                                                       (<%=resbelaeg_txt_074 %>:<%=aarsRamme %>)

                                                       <input id="Text6" type="hidden" name="FM_aarsramme_timer"  value="#" />
                                                       <input id="nl_medid_<%=medarbKundeoplysX(x, 0)%>" type="hidden" name="FM_aarsramme_medarb"  value="<%=medarbKundeoplysX(x, 3) %>" />
                                                       <input id="Hidden5" type="hidden" name="FM_aarsramme_rid"  value="<%=arrsIdThis %>" />
                                                       <input id="nl_jobid_<%=medarbKundeoplysX(x, 0)%>" type="hidden" name="FM_aarsramme_jobid"  value="<%=medarbKundeoplysX(x, 0) %>" />
                                                       <input id="Text5" type="hidden" name="FM_aarsramme_aar"  value="<%=aarsRamme %>" />

                                                   <%thisMTotRamme = thisMTotRamme + arrsNormThis  %>
                                                   </td>
                                                   <%end if %>


                                                   <%if len(trim(medarbKundeoplysX(x, 17))) <> 0 then
                                                   medarbKundeoplysX(x, 17) = medarbKundeoplysX(x, 17)
                                                   else
                                                   medarbKundeoplysX(x, 17) = 0
                                                   end if
                                                   %>


                                                    <!-- Aktiviteter --->
                                                  <%if cint(visAkt) = 1 then %>
                                                  <td style="vertical-align:top;">
                                                            <%
                                                            'WWF: aktive via timereg_usejob og faktuerbare
                                                            '*** PoSTIV aktivering af aktiviteter slået til

                                                            call hentaktiviteter(positiv_aktivering_akt_val, medarbKundeoplysX(x, 0), medarbKundeoplysX(x, 3), aty_sql_realhoursAkt)

                                                            'response.Write strSQLa & "<br>18:" & medarbKundeoplysX(x, 18)
                                                            'response.flush 
                                                            aSel = ""
                                                            aktText = ""
                    
                                                                if media <> "print" then%>
                                                            <select class="aaFM_jobid form-control input-small" id="aaFM_jobid_<%=x%>" name="sFM_aktid">
                                                            <option value="0">(<%=resbelaeg_txt_050 %>)</option>
                                                            <option value="0"><%=resbelaeg_txt_052 %>..? HER HEP</option>
                                                            <%
                                                            end if

                                                            oRec4.open strSQLa, oConn, 3
                                                            While not oRec4.EOF 

                                                             if ISNULL(oRec4("fase")) <> true AND len(trim(oRec4("fase"))) <> 0 then
                                                             fsNavn = " | fase: "& oRec4("fase")
                                                             else
                                                             fsNavn = ""
                                                             end if
                     
                                                             if cdbl(medarbKundeoplysX(x, 17)) = cdbl(oRec4("aid")) then
                                                             aSel = "SELECTED"
                                                             aktText = oRec4("aktnavn") & fsNavn
                                                             else
                                                             aSel = ""
                                                             end if
                     
                     
                                                              if media <> "print" then%>
                                                            <option value="<%=oRec4("aid")%>" <%=aSel %>><%=oRec4("aktnavn") &" "& fsNavn%></option>
                                                            <%end if

                                                            oRec4.movenext
                                                            wend
                                                            oRec4.close 
                    
                                                              if media <> "print" then%>
        
                                                            </select><br />
                                                            <%else %>
                                                            <%=aktText %>
                                                            <%end if %>
                    
                    
                    
                                                              <%if media <> "print" then 'AND cint(visAktiv) = 1 then %>
	                    	                                        <a href="#" id="anl_a_<%=medarbKundeoplysX(x, 0)%>_<%=medarbKundeoplysX(x, 3) %>" class="rodlille"><%=resbelaeg_txt_053 %> +</a> &nbsp; 
		                                                      <%end if%>
       
                                                    </td>
		                                            <%end if 'visAkt %>

                                            <%end if 'media %>



                                            <%
                                                '**************************************************************
		                                        '***** Udkskriver forecast timer pr. md. på joblinje **********
		                                        '**************************************************************

                                                if y = 0 then
	                                            forcastTimerJobtot = 0
                                                end if
                                                lastMonth = monththis

					                            y = 0
					                            for y = 0 TO numoffdaysorweeksinperiode

                                                'if lto = "mmmi" then
					                            'Response.Write periodeSel &","& y &","& startdato &","& monthUse & ""& numoffdaysorweeksinperiode  &"<br>"
                                                'end if
						
						                            call antaliperiode(periodeSel, y, startdato, monthUse)
                                                    timerThis = 0
                                                    procThis = 0
                                                    n = 0
                                                    'divider = 0
                                                    'foundone = 0	
                                                    For n = 0 to n_arr(medarbKundeoplysX(x, 3)) 'medarbKundeoplysX(x, 17) + 1 ' AND foundone = 0 'UBOUND(midjiddato)

                                                        'Response.Write "n: "& n &" "& medarbKundeoplysX(x, 3) & "<br>"
									                    '***** Uge visning, tjkker ikke MD ****'
									                    if periodeSel = 3 then
									
								                            if (midjiddato(medarbKundeoplysX(x, 3),n, 4) = newWeek) AND _
								                              midjiddato(medarbKundeoplysX(x, 3),n, 1) = medarbKundeoplysX(x, 0) AND _
                                                             midjiddato(medarbKundeoplysX(x, 3),n, 5) = medarbKundeoplysX(x, 17) AND _
                                                              midjiddato(medarbKundeoplysX(x, 3),n, 2) = medarbKundeoplysX(x, 3) AND _
								                               datepart("yyyy", midjiddato(medarbKundeoplysX(x, 3),n, 3)) = datepart("yyyy", nDay) then
    								        
								                                timerThis = midjiddato(medarbKundeoplysX(x, 3),n, 0)
                                                                procThis = midjiddato(medarbKundeoplysX(x, 3),n, 6)
                                           
    										
										                        'foundone = 1
    										
    										
									                        end if
								    
								    
								                        else
								                        '*** MD visning tjekker ikke om uge passer **'

                                  
								    
								                                if (datepart("m", midjiddato(medarbKundeoplysX(x, 3),n, 3), 2,2) = datepart("m", nDay, 2,2)) AND _
								                                    midjiddato(medarbKundeoplysX(x, 3),n, 1) = medarbKundeoplysX(x, 0) AND midjiddato(medarbKundeoplysX(x, 3),n, 5) = medarbKundeoplysX(x, 17) AND  _
                                                                    cdbl(midjiddato(medarbKundeoplysX(x, 3),n, 2)) = cdbl(medarbKundeoplysX(x, 3)) AND _
								                                    datepart("yyyy", midjiddato(medarbKundeoplysX(x, 3),n, 3)) = datepart("yyyy", nDay) then
								        
								                                    timerThis = midjiddato(medarbKundeoplysX(x, 3),n, 0)
                                                                    procThis = midjiddato(medarbKundeoplysX(x, 3),n, 6)
                                                                    'Response.Write "<br>timerThis" & timerThis
										      
										                            'foundone = 1
										                            'dividerThisMth( = divider + 1 
										      
									 
								                                end if

								                        end if

                                                    loops = loops + 1
									                next


                                                    if periodeSel <> 3 then 'AND cint(medarbKundeoplysX(x, 24)) = 1  'fra uge til gns pr. md medarbKundeoplysX(x, 24) = 1 = ANGIVET SOM UGE
                                                    divider = dividerThisMth(datepart("m", nDay, 2,2)) '+ 1
                                                    procThis = formatnumber((procThis/divider),0)
                                                    else
                                                    procThis = formatnumber((procThis),0)
                                                    end if
                                   

                                                    'response.write "procThis: "& procThis & " divider: "& divider & "<br>"

                                                    procThis = replace(procThis, ".", "")
									
									                if media <> "eksport" AND media <> "chart" then 
									                Response.flush
									                end if


                                                  
						                            xx = 0


                                                    if media <> "eksport" AND media <> "chart" then
                                                        if ddcView <> 1 then
					                                        if timerThis = 0 then
					                                        bgTimer = "#ffffff"
					                                        else
                                                            bgTimer = "#DCF5BD"
                                                            end if
                                                        end if
					                                end if


                                                    lastMonth = newMonth
					                                dato = nday
					                                stThisDato = dato
					                                txtwidth = 60

                                                   showTimerThis = 0

                                                    if cint(showasproc) = 1 then 'vis som procent
                                                        if procThis = 0 then
					                                    showTimerThis = ""
					                                    else
					                                    showTimerThis = procThis
					                                    end if                 
                                                    else					
				                                        if timerThis = 0 then
					                                    showTimerThis = ""
					                                    else
					                                    showTimerThis = timerThis
					                                    end if
                                                    end if

                                                    if media <> "eksport" AND media <> "chart" then
					                                %>
					                                <td style="vertical-align:top; background-color:<%=bgTimer%>; text-align:center;">
					                                <%
					                                end if


                                                    '*** Realiserede timer /medarbejder på job pr. md
					
					                                sumTimer = 0
					
					                                if periodeSel <> 3 then
					                                    mmwwSQLkri = " MONTH(tdato) = "& datepart("m", nday, 2, 2) 
					                                else
                                                        call thisWeekNo53_fn(nday)
                                                        mmwwSQLkri = " WEEK(tdato, 1) = "& thisWeekNo53 'datepart("ww", nday, 2, 2) 
					                                end if


                                                 
                                                     if medarbKundeoplysX(x, 1) <> fc_lastjobnrWrt then
                                                        aktSQLkriExcludeAkOnUspec = ""
                                                     end if




                                                     if medarbKundeoplysX(x, 17) <> 0 then 'specifik akt.
                                                        aktSQLkri = " AND taktivitetid = "& medarbKundeoplysX(x, 17)
                     
                                                        if medarbKundeoplysX(x, 1) <> fc_lastjobnrWrt then
                                                                aktSQLkriExcludeAkOnUspec = aktSQLkriExcludeAkOnUspec & " AND (taktivitetid <> "& medarbKundeoplysX(x, 17)
                                                        else
                                                                if medarbKundeoplysX(x, 17) <> fc_lastaktidWrt then
                                                                aktSQLkriExcludeAkOnUspec = aktSQLkriExcludeAkOnUspec & " AND taktivitetid <> "& medarbKundeoplysX(x, 17)
                                                                end if
                                                        end if
                        
                                                    else
                                                        aktSQLkri = ""
                                                        if medarbKundeoplysX(x, 17) = 0 AND aktSQLkriExcludeAkOnUspec <> "" AND instr(aktSQLkriExcludeAkOnUspec, ")") = 0 then
                                                        aktSQLkriExcludeAkOnUspec = aktSQLkriExcludeAkOnUspec & ")" 'trim til klar
                       
                                                        'aktSQLkriExcludeAkOnUspec = "" 'trim nulstillet
                                                        end if
                                                    end if



                                                    strSQLtimer = "SELECT sum(timer) AS sumtimerFaktisk FROM timer WHERE tmnr = "& medarbKundeoplysX(x, 3) &" AND tjobnr = '"& medarbKundeoplysX(x, 1) &"' "& aktSQLkri 

                                                    if medarbKundeoplysX(x, 17) = 0 then 'specifik akt.
                                                    strSQLtimer = strSQLtimer & aktSQLkriExcludeAkOnUspec 
                                                    end if  

                                                    strSQLtimer = strSQLtimer & " AND ("& aty_sql_realhours &") AND (YEAR(tdato) = "& datepart("yyyy", nday, 2, 2) &" AND " & mmwwSQLkri &")"
				                                    'Response.Write strSQLtimer
				                                    'Response.Flush
				                                    'Response.end
				    
				                                    oRec3.open strSQLtimer, oConn, 3
				                                    while not oRec3.EOF
    				
				                                    sumTimer = oRec3("sumtimerFaktisk")
    				
				                                    oRec3.movenext
				                                    wend
    				
				                                    oRec3.close


                                                    'Response.Write sumTimer & "<br>"
    				
				                                    if len(trim(sumTimer)) <> 0 then
				                                    sumTimer = sumTimer
				                                    else
				                                    sumTimer = 0
				                                    end if
				    
				                                    if len(trim(timerThis)) <> 0 then
				                                    timerForecast = timerThis 'showTimerThis
				                                    else
				                                    timerForecast = 0
				                                    end if

                                                    if len(trim(procThis)) <> 0 then
				                                    procForecast = procThis/1 
				                                    else
				                                    procForecast = 0
				                                    end if



                                                    '*** Afvigelse %				    
				                                    if cdbl(timerForecast) <= cdbl(sumTimer) then
						                                if cdbl(sumTimer) <> 0 then
                                                        afvThis = 100 - (timerForecast/sumTimer) * 100
                                                        else
                                                        afvThis = 0
                                                        end if
                                                    else
                                                        if cdbl(timerForecast) <> 0 then
                                                        afvThis = 100 - (sumTimer/timerForecast) * 100
                                                        'afvThis = replace(afvThis, "-", "")
                                                        else
                                                        afvThis = 0
                                                        end if
                                                    end if

                                                    'Response.Write "timerForecast:" & timerForecast & "<br>"
				                                    'Response.Write "sumTimer: " & sumTimer & "<br>"


                                                    if media <> "print" AND media <> "eksport" AND media <> "chart" then
				   	
				   	                                    '*** Er det admin eller jobansvarlig ****'
                                                        '*** Bør allerede være fundet i søgefilteret **
            				   	
				   	                                    'if (cint(session("mid")) = cint(medarbKundeoplysX(x, 3)) OR _
				   	                                    'session("mid") = medarbKundeoplysX(x, 11) OR _
				   	                                    'session("mid") = medarbKundeoplysX(x, 12)) AND _
                                                        redTimer = 0
				   	                                    if ((month(cdate(dagsdato)) <= month(cdate(nday)) AND year(cdate(dagsdato)) = year(cdate(nday))) OR (year(cdate(dagsdato)) < year(cdate(nday)))) OR level = 1 then
				   	                                    iptype = "text"
                                                        redTimer = 1
				   	                                    else 
				   	                                    iptype = "hidden"
				   	                                    end if
				   	                                    %> 
                				
                				
    				                                     <input type="<%=iptype%>" name="FM_timer" id="FM_timer_<%=medarbKundeoplysX(x, 3) %>_<%=medarbKundeoplysX(x, 0) %>_<%=x%>_<%=y%>" class="form-control input-small" value="<%=showTimerThis %>" style="width:60px; text-align:right;">

                                                         <%if redTimer = 0 AND ddcView = 1 then %>
                                                         <input style="width:60px; text-align:right;" type="text" class="form-control input-small" value="<%=showTimerThis %>" readonly />  <!-- Bliver kun brugt til at vise timerne, uden man kan redigere -->
                                                        <%end if %>

                                                        <%if ddcView <> 1 then %>
                                                            <%if iptype = "text" then%>
                                
                                                            <!--<input class="bt" id="bt_<%=medarbKundeoplysX(x, 3) %>_<%=medarbKundeoplysX(x, 0) %>_<%=x%>_<%=y%>" type="button" value=">>" style="font-size:8px;"  />-->

                                                            <span class="bt fa fa-share" style="font-size:9px;" id="bt_<%=medarbKundeoplysX(x, 3) %>_<%=medarbKundeoplysX(x, 0) %>_<%=x%>_<%=y%>"></span>

    				                                        <%else %>
    				                                        <%=showTimerThis %>
                                
    				                                        <%end if %>
                                                        <%end if %>
                                
    				                                    <!--'*** hidden array felt. *** Bruges så der kan skelneshvis man bruger komma.-->
					                                    <input type="hidden" name="FM_timer" id="FM_timer" value="#">
					                                <%else 
					
					                                if media <> "eksport" AND media <> "chart" then%>
					                                     <%=showTimerThis %> 
					                                <%end if %>
					        
					                                <%end if 'Print / media%>



                                                    <%if media <> "eksport" AND media <> "chart" then %>

                                                        <%if sumTimer <> 0 AND ddcView <> 1 then %>
                                                            <span style="font-size:9px; color:#5582d2;">
                            
                                                                <%if cint(showasproc) = 1 then %>
                                                                        f:  <%=formatnumber(timerThis, 0) %> t. <!-- (<=formatnumber(procThis, 0) %>%) --> / 
                                                                    <%end if%>

					                                        r: <%=formatnumber(sumTimer, 0)%> t.  <!--afv.: <%=formatnumber(afvThis, 0)%>%-->

                                                            </span>
					                                    <%end if %>
				
                    	                            </td>

                                                   <%if ddcView = 1 then %>
                                                        <td><input type="text" class="form-control input-small" readonly value="<%=formatnumber(sumTimer, 2) %>" style="width:60px; text-align:right;" /></td>
                                                   <%end if %>

					                                <%end if 'media %>

                                                    <%
                                                    '** Normtimer i periode til eksport
                                                    if media = "eksport" then
                                                        if lastxmid <> medarbKundeoplysX(x, 3) then
                                                            '** NormTimer pr medarb pr uge / md
                                                            'if periodesel = "6" OR periodesel = "12" then 
                                
                               

                                                            'call stKriInterval612(y, jobStartKri)

                                                            'response.write "lastxmid: "& lastxmid &" medarbKundeoplysX(x, 3)" & medarbKundeoplysX(x, 3) & " y: "& y &" jobStartKriThisM: "& jobStartKriThisM &" datoIntervalThisM: "& datoIntervalThisM &"<br>"
                                                            'response.flush
                            
	                                                        'call normtimerPer(medarbKundeoplysX(x, 3), jobStartKriThisM, 6, 0)
                                                            'normTimerMedarb = formatnumber(ntimPer*4,5, 0)        
                               
                                                            'response.write "normTimerMedarb: "& normTimerMedarb 
                                                            'response.flush
                                                            'else 
                            
                                                            'if y = 0 then
                                                            'jobStartKriNormT = day(jobStartKri) &"-"& month(jobStartKri) &"-"& year(jobStartKri)
                                                            'end if
                                                            'jobStartKri = dateAdd("d", y*7, jobStartKriNormT)
                                                            '** Altid pr. uge i udtræk, da det er det mest faste hodlepunkt mm. man henter en kolonne ud for hver md.
                                                            call normtimerPer(medarbKundeoplysX(x, 3), jobStartKri, 6, 0) 
                                                            normTimerMedarb = formatnumber(ntimPer, 1)               
                                                            'end if 

                                                        end if
                                                    end if


                                                    if (media = "eksport" AND vis <> 1)  then
                         
                                                         if lastxmid <> medarbKundeoplysX(x, 3) then

                                                            '*** Bregner pr. uge til at omregne mellem timer og %
                                                            '*** Er nødt til at øre det igen, da vi ovenfor berenger på både pr. md eller ugfe alt efter peridoevalg. herunder SKAL det altid være pr. uge
                                                             call normtimerPer(medarbKundeoplysX(x, 3), jobStartKri, 6, 0) 

                                                         end if

                                                        if lastxmid <> medarbKundeoplysX(x, 3) OR lastJid <> medarbKundeoplysX(x, 0) then
                                                            call arsRammeSub
                                                        end if
                                                    end if


                                                    if media = "eksport" then

                                                         if cint(vis) = 1 then 'Alm 

                                                            if (lastAktJidExpVis_1 <> medarbKundeoplysX(x, 0) &"_"& medarbKundeoplysX(x, 17) &"_"& medarbKundeoplysX(x, 3)) then
                                        
                                                                csvTxt = csvTxt & vbcrlf & medarbKundeoplysX(x, 5)&";"& medarbKundeoplysX(x, 6) &"" _
                                                                &";" & medarbKundeoplysX(x, 2) & ";"& medarbKundeoplysX(x, 1) & ";"
                                        
                                                                'if cint(visAktiv) = 1 then ALTID med da der kan væres flere linier på samme job ved uspec. Dette skal kunne ses i udtrækkket.
                                                                csvTxt = csvTxt & medarbKundeoplysX(x, 18) & ";" & medarbKundeoplysX(x, 17) & ";" 
                                                                'end if                

                                                                csvTxt = csvTxt & medarbKundeoplysX(x, 4) &" ("& medarbKundeoplysX(x, 7) &");" _
					                                            & medarbKundeoplysX(x, 8) &";" _ 
					                                            & medarbKundeoplysX(x, 9) &";" _ 
					                                            & normTimerMedarb

                                                                '0: Jobid, 3: Mid, 17: aktid
                                                            
                                                                lastAktJidExpVis_1 = medarbKundeoplysX(x, 0) &"_"& medarbKundeoplysX(x, 17) &"_"& medarbKundeoplysX(x, 3)
                                       
                                                            end if

                                                            csvTxt = csvTxt &";"& showTimerThis 

                                                            if cint(expvisreal) = 1 then
                                        
                                                            if len(trim(sumTimer)) <> 0 then
                                                            sumTimerExp = formatnumber(sumTimer,2)
                                                            else
                                                            sumTimerExp = ""
                                                            end if

                                                            csvTxt = csvTxt &";"& sumTimerExp
					                                        end if


                                                        else 'PIVOT 


                                                            csvTxt = csvTxt & vbcrlf & medarbKundeoplysX(x, 4) &" ("& medarbKundeoplysX(x, 7) &")"
					                                        csvTxt = csvTxt & ";"& medarbKundeoplysX(x, 8) 
					                                        csvTxt = csvTxt & ";"& medarbKundeoplysX(x, 9) &";" _ 
					                                        & normTimerMedarb 
					

                                                            'if (lastxmid <> medarbKundeoplysX(x, 3) OR lastJid <> medarbKundeoplysX(x, 0)) AND cint(isArsNormWrtEksp) = 0 then

                                                            if cint(visRamme) = 1 then
                                                            'csvTxt = csvTxt & ";"& formatnumber(ntimPer, 2) 
                                                            csvTxt = csvTxt & ";"& aarsRamme & ";" & arrsNormThis
                                                            isArsNormWrtEksp = 1
                                                            end if

                                                            'else
                                                            'csvTxt = csvTxt & ";;;" 
                                                            'end if

					                                        csvTxt = csvTxt & ";"& medarbKundeoplysX(x, 2) & ";" & medarbKundeoplysX(x, 1)

                                                            'if cint(visAkt) = 1 then ALTID med i udtræk der kan være flere linier på ammme job. Skal fremgå
                                                            csvTxt = csvTxt & ";"& medarbKundeoplysX(x, 18) & ";" & medarbKundeoplysX(x, 17) 
					                                        'end if                

					                                        csvTxt = csvTxt & ";"& medarbKundeoplysX(x, 13) & ";" & medarbKundeoplysX(x, 14)
					                                        csvTxt = csvTxt & ";"& medarbKundeoplysX(x, 15) & ";" & medarbKundeoplysX(x, 16)
					
					                                        csvTxt = csvTxt & ";"& medarbKundeoplysX(x, 5) & ";" & medarbKundeoplysX(x, 6)
					
					                                        'if lto <> "execon" then
					                                        csvTxt = csvTxt & ";"& csvDatoerUge(y) 
					                                        'end if
					
					                                        csvTxt = csvTxt & ";"& csvDatoerMD(y) &";"& csvDatoerAR(y) &";"& timerForecast 
                        
                                                            if cint(visStatus) = 1 then
                                                            csvTxt = csvTxt &";"& sumTimer &";"& formatnumber(afvThis, 0)
                                                            end if
                                       

					                                        'csvTxt = csvTxt & vbcrlf & ";;;;;;;"& csvDatoer(y) &";" & timerThis

                                                        end if 'vis

                                                    end if 'media
                                                    %>




                                                    <%if media <> "eksport" AND media <> "chart" then  %>

                                                        <input type="hidden" class="ahFM_jobid_<%=x%>" name="FM_aktid" id="Hidden4" value="<%=medarbKundeoplysX(x, 17)%>">
                                                        <input type="hidden" name="FM_aktid_old" id="Hidden6" value="<%=medarbKundeoplysX(x, 17)%>">
					                                    <input type="hidden" name="FM_jobid" id="Hidden3" value="<%=medarbKundeoplysX(x, 0)%>">
					                                    <input type="hidden" name="FM_medarbid" id="FM_medarbid" value="<%=medarbKundeoplysX(x, 3)%>">
					                                    <input type="hidden" name="FM_dato" id="FM_dato" value="<%=stThisDato%>">

                                                    <%
					                                end if

                      
                                                    fc_lastaktidWrt = medarbKundeoplysX(x, 17)
					                                fc_lastjobnrWrt = medarbKundeoplysX(x, 1)
					
					                                medarbTotalTimer(y) = medarbTotalTimer(y) + timerForecast
					                                timerTotalTotal(y) = timerTotalTotal(y) + timerForecast

                                                    medarbTotalProc(y) = medarbTotalProc(y) + procForecast
					                                procTotalTotal(y) = procTotalTotal(y) + procForecast

                                                    medarbTimerReal(y) = medarbTimerReal(y) + sumTimer
					                                realTotalTotal(y) = realTotalTotal(y) + sumTimer
					
					                                timerMJobTotal(y) = timerMJobTotal(y) + timerForecast
                                                    procMJobTotal(y) = procMJobTotal(y) + procForecast
                    
					                                realMJobTotal(y) = realMJobTotal(y) + sumTimer

                                                    '** Jobtotal
                                                    realTimerJobtot = realTimerJobtot + sumTimer 
					                                forcastTimerJobtot = forcastTimerJobtot + timerForecast
                                                    forcastProcJobtot = forcastProcJobtot + procForecast 

                                                    '** Medarbejder total
                                                    forecastProcMJobtotGrand = forecastProcMJobtotGrand + procForecast
                                                    forecastTimerMJobtotGrand = forecastTimerMJobtotGrand + timerForecast
				                                    realTimerMJobtotGrand = realTimerMJobtotGrand + sumTimer  
                    
                                                    lastY = y
                  

					                            next 'y periode 


                                                saldoJobTot = (forcastTimerJobtot - realTimerJobtot)
                                                saldoJobTotIgnper = (forcastTimerTotAktMedignPer - realTimerTotAktMedignPer)
                                                        
                                                        
                                                if media <> "eksport" AND media <> "chart" then

                                                    if cint(visStatus) = 1 then%>
                    
                                                        <%if ddcView <> 1 then %>
                                                        <td style="text-align:right;"><%=formatnumber(medarbKundeoplysX(x, 20), 0) %> <br />
                                                           <span style="font-size:10px; color:#999999;"><%=formatnumber(medarbKundeoplysX(x, 19), 0) %> DKK<br /></span>

                                                            <span style="font-size9px; color:#999999;">
                                                            <%if medarbKundeoplysX(x, 20) <> 1 then %>
                                                            rest.: <%=formatnumber(medarbKundeoplysX(x, 21), 0) %> t.
                                                            <%else %>
                                                            afsl.:<%=formatnumber(medarbKundeoplysX(x, 21), 0) %> %
                                                            <%end if %>

                                                            </span>
                                                        </td>
                                                        <%end if %>

                                                        <td style="vertical-align:top; text-align:right;"><%=formatnumber(forcastTimerJobtot, 0) %> <!-- (<%=formatnumber(forcastProcJobtot/dividerAll, 0) %>%) -->
                                                            <%if ddcView <> 1 then %>
                                                            <br /><span style="font-size:9px; color:#999999;"><%=formatnumber(forcastTimerTotAktMedignPer, 0) %></span>
                                                            <%end if %>
                                                        </td>

			                                            <td style="vertical-align:top; text-align:right;"><%=formatnumber(realTimerJobtot, 2)%>
                                                            <%if ddcView <> 1 then %>
                                                            <br /><%=realTimIngPerTxt %>
                                                            <%end if %>
                                                        </td>	

		                                                <td style="vertical-align:top; text-align:right;">

                                                        <%
                                                        if ddcView <> 1 then
                                                            if saldoJobTot > 0 then
                                                            bgColSaldoTot = "#6CAE1C" 
                                                            else
                                                            bgColSaldoTot = "darkred"
                                                            end if
                                                        else
                                                            bgColSaldoTot = ""
                                                        end if
                                                        %>


                                                        <%
                                                        if saldoJobTotIgnper > 0 then
                                                        bgColSaldoTotig = "#6CAE1C" 
                                                        else
                                                        bgColSaldoTotig = "darkred"
                                                        end if%>

                                                        <span style="color:<%=bgColSaldoTot%>;">
                                                        <%=formatnumber(saldoJobTot, 0) %></span>
                                                            <%if ddcView <> 1 then %>
                                                            <br /><span style="font-size:9px; color:<%=bgColSaldoTotig%>;"><%=formatnumber(saldoJobTotIgnper, 0) %></span>
                                                            <%end if %>
		                                                </td>
                                                    <%
                                                end if 'visStatus

                                             end if 'media






                                            if media = "eksport" AND vis = "1" then

                                                if forcastTimerJobtot <> 0 then
                                                forcastTimerJobtotExpTxt = forcastTimerJobtot
                                                else
                                                forcastTimerJobtotExpTxt = ""
                                                end if

                                                if forcastProcJobtot <> 0 then
                                                forcastProcJobtotExpTxt = forcastProcJobtot
                                                else
                                                forcastProcJobtotExpTxt = ""
                                                end if

                                                if forcastTimerTotAktMedignPer <> 0 then
                                                forcastTimerTotAktMedignPerExpTxt = forcastTimerTotAktMedignPer
                                                else
                                                forcastTimerTotAktMedignPerExpTxt = ""
                                                end if


                                                if realTimerJobtot <> 0 then
                                                realTimerJobtotExpTxt = realTimerJobtot
                                                else
                                                realTimerJobtotExpTxt = ""
                                                end if


                                                if realTimerTotAktMedignPer <> 0 then
                                                realTimerTotAktMedignPerExpTxt = realTimerTotAktMedignPer
                                                else
                                                realTimerTotAktMedignPerExpTxt = ""
                                                end if

                                                if cint(visStatus) = 1 then 

                                                    if lastAktJidExpVis_1Total <> lastAktJidExpVis_1 then

                                                        csvTxt = csvTxt  &";"& formatnumber(medarbKundeoplysX(x, 20), 0)
                                                        csvTxt = csvTxt  &";"& formatnumber(medarbKundeoplysX(x, 19), 0)

                                                        csvTxt = csvTxt & ";"& forcastTimerJobtotExpTxt
                                                        csvTxt = csvTxt & ";"& forcastTimerTotAktMedignPerExpTxt
                                                        csvTxt = csvTxt & ";"& realTimerJobtotExpTxt
                                                        csvTxt = csvTxt & ";"& realTimerTotAktMedignPerExpTxt
                                                        csvTxt = csvTxt & ";"& saldoJobTot
                                                        csvTxt = csvTxt & ";"& saldoJobTotIgnper

                                                        lastAktJidExpVis_1Total = lastAktJidExpVis_1

                                                    end if
                                    
                                                end if


                                            end if


                                           if media <> "eksport" AND media <> "chart" then%>
		
		                                    </tr>
                                            
                                          <%end if %>



                                    <%
                                    end if 'Forecast overskreddet


                                    'Response.write "<tr><td colspan=1000>herS<br></td></tr>"

        
                                    forcastTimerTotJobMedignPer = forcastTimerTotJobMedignPer + forcastTimerTotAktMedignPer
                    
                                    realTimerJobtotGrand = realTimerJobtotGrand + realTimerJobtot 

		                            forecastTimerJobtotGrand = forecastTimerJobtotGrand + forcastTimerJobtot
                                    forecastProcJobtotGrand = forecastProcJobtotGrand + forcastProcJobtot

                                    forcastTimerTotJobMedignPerGrand = forcastTimerTotJobMedignPerGrand + forcastTimerTotAktMedignPer 'forcastTimerTotJobMedignPer

                                    pageForecastGrandtot = pageForecastGrandtot + forcastTimerJobtot 
		                            pageRealGrandtot = pageRealGrandtot + realTimerJobtot


                                     '*** Grand Grand total i bunde ALLE medarbejdere / Alle job **'
                                    fcGrandGrand = fcGrandGrand + forcastTimerJobtot
                                    fpGrandGrand = fpGrandGrand + forcastProcJobtot
                                    fcignPerGrandGrand = fcignPerGrandGrand + forcastTimerTotAktMedignPer
                                    realGrandGrand = realGrandGrand + realTimerJobtot 
                                    realignPerGrandGrand = realignPerGrandGrand + realTimerTotAktMedignPer

                                   jobBgtGrandGrand = jobBgtGrandGrand +  medarbKundeoplysX(x, 19)
                                   jobTimerGrandGrand = jobTimerGrandGrand +  medarbKundeoplysX(x, 20)
        
        
                                   realTimerJobtot = 0 
		                           forcastTimerJobtot = 0
                                   forcastProcJobtot = 0


                                   'strjIdWrt = strjIdWrt & "#"&medarbKundeoplysX(x, 0)&"#," 
		                           lastJid = medarbKundeoplysX(x, 0)
	                               lastJid_Aid = medarbKundeoplysX(x, 0) &"_"& medarbKundeoplysX(x, 17)


                                   lastJobans1 = medarbKundeoplysX(x, 11)
		                           lastJobans2 = medarbKundeoplysX(x, 12)


                                end if '** Lastjobid


                                if lastxmid <> medarbKundeoplysX(x, 3) then
                                    lastxmid = medarbKundeoplysX(x, 3)
                                end if


                            next 'medarb loop
        
                            Response.flush

                            %>

                                <%
                                    '*************************************************************************************************
                                    '********** Grandtotaler i bunden ****************************************************************
                                    '*************************************************************************************************

                                    if media <> "eksport" then

                                        if cint(vis_job_fc_neg) <> 1 then

                                            '*****************************************
		                                    '*** Tildel timer  / Guiden aktive job****
		                                    '*****************************************
                                            
                                            if media <> "chart" then 
				           
                                                'if ddcView <> 1 then
                                                call jobtotalprmedarb
                                                'end if
                                                      				        
                                                'if cint(jobidsel) = 0 OR ddcView = 1 then
                                                if jobidsel = "0" OR ddcView = 1 then
                                                    '**** Medarb. total ***'
                                                    'if cint(vis_simpel) <> 1 then
                                                    mTotthisMid = lastxmid
                                                    'if ddcView <> 1 then
                                                    call medarbTotal
                                                    'end if
                                                    'end if
                                                end if
                            
                                                'if cint(vis_simpel) <> 1 then
                                                call grandtotal
                                                'end if

                                            end if


                                            'if cint(jobidsel) = 0 OR ddcView = 1 then
                                             if jobidsel = "0" OR ddcView = 1 then

                                                if media = "chart" then
                            
                                                        'Response.write "her"

                                                        call meStamdata(lastxmid)
                                                        csvTxtA = csvTxtA & meTxt 

                                                        for y = 0 TO numoffdaysorweeksinperiode
                                                        csvTxtA = csvTxtA & ";"& medarbTotalTimer(y)
                                                        medarbTotalTimer(y) = 0 
                                                        next

                                                        csvTxtA = csvTxtA & "xx99123sy#z"
                                                end if
                                            end if


                                            if media <> "print" AND media <> "chart" AND cint(hideEmptyEmplyLines) <> 1 then
				   
                                                '** Henter medarbejder linier, (til at oprette forecast) på valgte medarbejder der ikke er vist på siden ***'
                                                for m = 0 to UBOUND(intMids)
                     

                                                     'Response.Write "medarbwrt: " & medarbwrt & " intMids(m): "& intMids(m) & "<br>"

                                                     if instr("#"&medarbwrt&"#", intMids(m)) = 0 then
                                                     'call tilfojNylinie(intMids(m))
                                                      call tilfojNylinie(intMids(m), visRamme, visAktiv)
                                                      call medarbejderLinje(intMids(m),jobStartKri,datoInterval,rdimnb,0)
                                                     end if
                
                                                next

                                            end if 'Print


                                            if media <> "chart" then

                                                if medarbSel = "0" then
                                                %>
				                                     <!--- Grandtotal --->
				
				                                    <tr>
					                                    <td style="text-align:right;"><b><%=resbelaeg_txt_054 %>:</b><br />
					                                    <%=resbelaeg_txt_055 %></font></td>
                                                        <td bgcolor="#ffffff">&nbsp;</td>
					
					                                    <%
					
						                                    for y = 0 TO numoffdaysorweeksinperiode
						    
						                                    '*** Afvigelse ***
			                                                tottotAfv = 0
    			
    				    
				                                            if timerTotalTotal(y) >= realTotalTotal(y) then
                                                                if timerTotalTotal(y) <> 0 then
                                                                tottotAfv = 100 - (realTotalTotal(y)/timerTotalTotal(y)) * 100
                                                                else
                                                                tottotAfv = 0
                                                                end if
                                                            else
                                                                if realTotalTotal(y) <> 0 then
                                                                tottotAfv = 100 - (timerTotalTotal(y)/realTotalTotal(y)) * 100
                                                                else
                                                                tottotAfv = 0
                                                                end if
                                                            end if
						
						                                    %>
							
						                                    <td><b><%=formatnumber(timerTotalTotal(y),0)%></b> <br /> <%=formatnumber(realTotalTotal(y), 0)%> ~ <%=formatnumber(tottotAfv,0) %>%</td>
						                                    <%
						                                    next
						                                    %>
						                                    <td style="text-align:right;"><b><%=formatnumber(pageForecastGrandtot,0) %></b></td>
						                                    <td style="text-align:right;"><%=formatnumber(pageRealGrandtot,0) %></td>
				                                            <td style="text-align:right;">&nbsp;</td>
                                                    </tr>
				                                <%
				                                end if 'medarbSel

                                            end if 'chart

                                        end if 'cint(vis_job_fc_neg) <> 1


                                    end if 'export
                                %>
                            <!--</thead>-->
                        </table>

                        <%if media <> "eksport" then %>
                            <div class="row">
                                <div class="col-lg-12"><button type="submit" class="btn btn-success btn-sm pull-right"><b><%=resbelaeg_txt_049 %></b></button></div>
                            </div>
                        <%end if %>

                        <%end if %>
                        </form>

                        <!-- Job first tabel -->
                        <%if sortering = 0 then %>

                        <%
                            dagsdato = now
                            nday = startdato
                        %>

                        <form action="ressource_belaeg_jbpla.asp?func=redalle" method="post">
                        
                           <input type="hidden" id="FM_medarb_new" name="FM_medarb" value="<%=request("FM_medarb") %>" />
                           <input type="hidden" name="FM_vis_job_u_fc" value="<%=vis_job_u_fc %>" />
                           <!-- <div class="row">
                                <div class="col-lg-2"><span id="createnewbtn" class="btn btn-success btn-sm">+</span></div>
                            </div> -->

                        <table class="table table-bordered table-hover">                           

                            <%call datoeroverskrift(1) %>

                            <%

                            if medarbSQlKri <> "" then
                                medarbSQlKri = " AND " & medarbSQlKri
                            end if

                            if len(medarbSQlKri2) <> 0 then
                                medarbSQlKri2 = " AND r.medid IN ("&medarbSQlKri2&")"
                            else
                                medarbSQlKri2 = ""
                            end if

                            if cint(startdatoYY) = cint(slutdatoYY) then
                                orandsign = "AND"
                            else
                                orandsign = "OR"
                            end if

                            dim grandtotIperFC, grandtotIperRL
                            redim grandtotIperFC(12), grandtotIperRL(12)

                            forecastIalt = 0
                            realisertIalt = 0
                            if vis_job_u_fc = 1 then
                            strSQlJob = "SELECT id as jobid, jobnavn, jobnr FROM job j WHERE "& jobSQLkri
                            else
                            strSQlJob = "SELECT j.id as jobid, jobnavn, jobnr FROM job j LEFT JOIN ressourcer_md r ON (r.jobid = j.id) WHERE "& jobSQLkri &" AND timer <> 0 AND ((r.md >= "& startdatoMD &" AND r.aar = "& startdatoYY &") "& orandsign &" ( r.md <= "& slutdatoMD &" AND r.aar = "& slutdatoYY &"))  "& medarbSQlKri2 &" GROUP BY jobid"
                            end if
                            'response.Write strSQlJob 
                            'response.End
                            oRec.open strSQlJob, oConn, 3
                            while not oRec.EOF

                                

                                'total fc for jobbet i perioderne
                                dim totFCperiode
                                redim totFCperiode(12)

                                'total rl for jobbet i perioderne
                                dim totRLperiode
                                redim totRLperiode(12)


                                ' Forecast timer
                                dim timerMedarbIperiode
                                redim timerMedarbIperiode(500, 100, 12, 2) 'timerMedarbIperiode(medid, nummer, md, aar) - betydning 

                                dim medarbejderid, medarbejdernavn, aktivitetsid, aktivitetsnavn, aktivitetsidAlle
                                redim medarbejderid(250), medarbejdernavn(250), aktivitetsid(1000, 250), aktivitetsnavn(250, 250), aktivitetsidAlle(250) 'aktivitetsnavn(nummer, medid), aktivitetsid(nummer, medid)


                                'Henter først medarbejdere med fc på jobbet
                                if vis_job_u_fc = 0 then
                                    strSQLmedarb = "SELECT r.medid as medid, m.mnavn as mnavn FROM ressourcer_md r LEFT JOIN medarbejdere m ON (m.mid = r.medid) WHERE jobid = "& oRec("jobid") & " AND ((r.md >= "& startdatoMD &" AND r.aar = "& startdatoYY &") "& orandsign &" ( r.md <= "& slutdatoMD &" AND r.aar = "& slutdatoYY &"))" & medarbSQlKri
                                else
                                    strSQLmedarb = "SELECT mid as medid, mnavn FROM medarbejdere as m WHERE mid <> 0 "& medarbSQlKri
                                end if
                                'response.Write strSQLmedarb
                                oRec2.open strSQLmedarb, oConn, 3
                                m = 0
                                while not oRec2.EOF
                                    medarbejderid(m) = oRec2("medid")
                                    medarbejdernavn(m) = oRec2("mnavn")
                                m = m + 1
                                oRec2.movenext
                                wend
                                oRec2.close


                                strSQLfcTimer = "SELECT sum(timer) as timer, medid, md, aar, m.mnavn as mnavn, aktid, a.navn as aktnavn FROM ressourcer_md LEFT JOIN medarbejdere m ON (m.mid = medid)" 
                                strSQLfcTimer = strSQLfcTimer & " LEFT JOIN aktiviteter a ON (a.id = aktid)"
                                strSQLfcTimer = strSQLfcTimer & " WHERE jobid ="& oRec("jobid") & " AND ((md >= "& startdatoMD &" AND aar = "& startdatoYY &") "& orandsign &" ( md <= "& slutdatoMD &" AND aar = "& slutdatoYY &"))" & medarbSQlKri
                                strSQLfcTimer = strSQLfcTimer & " GROUP BY medid, jobid, aktid, md, aar ORDER BY aktid, md, aar"
                                'response.Write "<tr><td>herher" & strSQLfcTimer & "</td><tr>"
                                oRec2.open strSQLfcTimer, oConn, 3 
                                x = 0
                                antal = 0
                                timeraar = 1
                                lastaktid = -1
                                while not oRec2.EOF

                                    if oRec2("aktid") <> lastaktid AND antal <> 0 then
                                        x = x + 1
                                    end if
                                    
                                    if oRec2("aar") = startdatoYY then
                                        timeraar = 1
                                    else
                                        timeraar = 2                                       
                                    end if
                                    'response.Write "timeer " & timeraar
                                    timerMedarbIperiode(oRec2("medid"), x, oRec2("md"), timeraar) = oRec2("timer")
                                    'response.Write "<tr><td>"& timerMedarbIperiode(oRec2("medid"), x, oRec2("md"), timeraar) & " x "& x &"</td><tr>"
                                    'medarbejderid(x) = oRec2("medid")
                                    'medarbejdernavn(x) = oRec2("mnavn")
                                    aktivitetsid(x, oRec2("medid")) = oRec2("aktid")
                                    aktivitetsnavn(x, oRec2("medid")) = oRec2("aktnavn")
                                    aktivitetsidAlle(x) = oRec2("aktid")
                                     
                                
                                lastaktid = oRec2("aktid")
                                antal = antal + 1
                                oRec2.movenext
                                wend
                                oRec2.close

                                ' Realiseret timer for aktiviteter med forecast paa

                                dim realTimerMedarbIperiode, totFCiPeriode
                                redim realTimerMedarbIperiode(500, 100, 12, 2)
                                'realTimerMedarbIperiode(mid, nummer, month, year) - betydning

                                 for a = 0 to UBOUND(aktivitetsidAlle)
                                    if aktivitetsidAlle(a) <> "" then
                                        strSQLrlTimer = "SELECT sum(timer) as timer, tmnr, month(tdato), year(tdato), tjobnr, taktivitetid, tdato FROM timer " _
                                                        & "WHERE tjobnr = '"& oRec("jobnr") &"' AND taktivitetid = '"& aktivitetsidAlle(a) &"' AND ((month(tdato) >= "& startdatoMD &" AND year(tdato) = "& startdatoYY &") OR (month(tdato) <= "& slutdatoMD &" AND year(tdato) = "& slutdatoYY &")) "_
                                                        & "GROUP BY tmnr, tjobnr, taktivitetid, month(tdato), year(tdato)"

                                        oRec2.open strSQLrlTimer, oConn, 3
                                        while not oRec2.EOF
                                            if year(oRec2("tdato")) = startdatoYY then
                                                timeraar = 1
                                            else
                                                timeraar = 2                                       
                                            end if
                                            realTimerMedarbIperiode(oRec2("tmnr"), a, month(oRec2("tdato")), timeraar) = oRec2("timer")
                                        oRec2.movenext    
                                        wend
                                        oRec2.close
                                    end if
                                 next
                               
                            %>

                            <tr id="fieldfor_<%=oRec("jobnr") %>">
                                <th style="width:250px; cursor:pointer;" class="expandjob" id="<%=oRec("jobnr") %>"><span class="expandsign" id="expandsign_<%=oRec("jobnr") %>">+</span> <span style="font-size:14px; line-height:18px;"><b><%=oRec("jobnavn") & " ("&oRec("jobnr")&")" %></b></span></th>
                                
                                <%
                                loopmd = startdatoMD
                                looprealaar = startdatoYY
                                loopaar = 1
                                %>

                                <%for y = 0 TO numoffdaysorweeksinperiode %>

                                    <%
                                     if y <> 0 then
                                        if loopmd = 12 then
                                            loopmd = 1
                                            loopaar = loopaar + 1
                                            looprealaar = looprealaar + 1
                                        else
                                            loopmd = loopmd + 1
                                        end if
                                    end if
                                    %>



                                    <th style="text-align:right; width:60px;"><b><span id="totFC_<%=oRec("jobid") %>_<%=y %>"></span></b></th>
                                    <th style="text-align:right; width:60px;"><b><span id="totReal_<%=oRec("jobid") %>_<%=y %>"></span></b></th>

                                <%next %>

                                <th style="text-align:right;"><b><span id="jobtotFC_<%=oRec("jobid") %>"></span></b></th>
                                <th style="text-align:right;"><b><span id="jobtotRL_<%=oRec("jobid") %>"></span></b></th>
                                <th style="text-align:right;"><b><span id="jobtotSaldo_<%=oRec("jobid") %>"></span></b></th>
                                <th style="width:15px; text-align:center;"><span style="color:#afafaf;" class="btn btn-sm btn-default fa fa-plus addField" id="<%=oRec("jobnr") %>" data-jobid="<%=oRec("jobid") %>"></span></th>
                            </tr>
                            <%
                            lastmedid = -1
                            for m = 0 to UBOUND(medarbejderid)                             
                            %>
                                <%if medarbejderid(m) <> "" AND medarbejderid(m) <> lastmedid then %>  
                                    <%
                                        lastaktid = -1                                        
                                    %>
                                    <%for a = 0 to UBOUND(aktivitetsidAlle) %>
                                        <%
                                        thisaktid = aktivitetsid(a, medarbejderid(m))
                                        thisaktnavn = aktivitetsnavn(a, medarbejderid(m))
                                        medarbakttotFC = 0
                                        medarbakttotRL = 0
                                        %>
                                        <%
                                        if thisaktid <> "" AND aktivitetsidAlle(a) <> lastaktid then %>
                                                                                        
                                            <tr class="fieldfor_<%=oRec("jobnr") %> columns" style="display:none;">
                                                <td><%=medarbejdernavn(m) %></td>

                                                <%
                                                loopmd = startdatoMD
                                                looprealaar = startdatoYY
                                                loopaar = 1
                                                %>

                                                <%for y = 0 TO numoffdaysorweeksinperiode %>

                                                    <%
                                                    if y <> 0 then
                                                        if loopmd = 12 then
                                                            loopmd = 1
                                                            loopaar = loopaar + 1
                                                            looprealaar = looprealaar + 1
                                                        else
                                                            loopmd = loopmd + 1
                                                        end if
                                                    end if

                                                    timerForecast = timerMedarbIperiode(medarbejderid(m), a, loopmd, loopaar)
                                                    timerRealiseret = realTimerMedarbIperiode(medarbejderid(m), a, loopmd, loopaar)

                                                    medarbakttotFC = medarbakttotFC + timerForecast
                                                    medarbakttotRL = medarbakttotRL + timerRealiseret
                                                    medarbaktSaldo = medarbakttotFC - medarbakttotRL 
                                                    
                                                    totFCperiode(y) = totFCperiode(y) + timerForecast
                                                    totRLperiode(y) = totRLperiode(y) + timerRealiseret
                                                    
                                                    nday = looprealaar &"-"& loopmd &"-01"

                                                    redTimer = 0

                                                    if ((month(cdate(dagsdato)) <= month(cdate(nday)) AND year(cdate(dagsdato)) = year(cdate(nday))) OR (year(cdate(dagsdato)) < year(cdate(nday)))) OR level = 1 then
				   	                                    iptype = "text"
                                                        redTimer = 1
				   	                                else 
				   	                                    iptype = "hidden"
				   	                                end if

                                                    %>

                                                    <td>
                                                        <input type="hidden" name="FM_dato" value="<%="1-"& loopmd &"-"& looprealaar%>" />
                                                        <input type="hidden" name="FM_medarbid" value="<%=medarbejderid(m) %>" />
                                                        <input type="hidden" name="FM_jobid" value="<%=oRec("jobid") %>" />
                                                        <input type="hidden" name="FM_aktid" value="<%=thisaktid %>" />
                                                        <input type="hidden" name="FM_aktid_old" value="<%=thisaktid %>" />
                                                                                                                
                                                        <input style="text-align:right; width:60px;" type="<%=iptype %>" name="FM_timer" value="<%=timerForecast %>" class="form-control input-small" />

                                                        <%if redTimer = 0 then %>
                                                        <input style="text-align:right; width:60px;" type="text" class="form-control input-small" value="<%=timerForecast %>" readonly /> <!-- Bliver kun brugt til at vise timerne, uden man kan redigere -->
                                                        <%end if %> 

                                                        <input type="hidden" name="FM_timer" id="FM_timer" value="#" />
                                                    </td>
                                                    <td><input style="text-align:right; width:60px;" type="text" readonly value="<%=formatnumber(timerRealiseret, 2) %>" class="form-control input-small" /></td>
                                                    
                                                    
                                                <%next %>

                                                <td style="text-align:right;"><%=formatnumber(medarbakttotFC, 2) %></td>
                                                <td style="text-align:right;"><%=formatnumber(medarbakttotRL, 2) %></td>
                                                <td style="text-align:right;"><%=formatnumber(medarbaktSaldo, 2) %></td>
                                               <!-- <td style="width:15px; text-align:center;"><span style="color:#afafaf;" class="btn btn-sm btn-default fa fa-plus addField" id="<%=oRec("jobnr") %>" data-jobid="<%=oRec("jobid") %>" data-aktid="<%=thisaktid %>"></span><%=" aktid "& thisaktid %></td> -->
                                                                
                                                <%
                                                forecastIalt = forecastIalt + medarbakttotFC
                                                realisertIalt = realisertIalt + medarbakttotRL
                                                %>

                                            </tr>                                            
                                        <%end if %>
                                        <%lastaktid = aktivitetsidAlle(a) %>
                                    <%next %>

                                <%end if %>
                                <%lastmedid = medarbejderid(m) %>
                            <%next %>

                            <%
                            'Dem uden forecast kommer her, så dem med står føst
                            if vis_job_u_fc = 1 then
                                lastmedid = -1
                                for m = 0 to UBOUND(medarbejderid)
                                    if medarbejderid(m) <> "" AND medarbejderid(m) <> lastmedid then

                                        lastaktid = -1
                                        aktFindes = 0
                                        for a = 0 to UBOUND(aktivitetsidAlle)
                                            thisaktid = aktivitetsid(a, medarbejderid(m))
                                            thisaktnavn = aktivitetsnavn(a, medarbejderid(m))

                                            if thisaktid <> "" AND aktivitetsidAlle(a) <> lastaktid then
                                                aktFindes = 1
                                            end if
                                        next

                                        if aktFindes = 0 then                                   
                                        %>
                                    
                                            <%

                                            totalMedReal = 0

                                            'Finder den aktivitet der skal forecastes på
                                            aktidFC = 0
                                            strSQL = "SELECT id FROM aktiviteter WHERE job = "& oRec("jobid")
                                            oRec2.open strSQL, oConn, 3 
                                            if not oRec2.EOF then
                                                aktidFC = oRec2("id")
                                            end if
                                            oRec2.close
                                            %>
                                            
                                            <tr class="fieldfor_<%=oRec("jobnr") %> columns" style="display:none;">
                                                <td><%=medarbejdernavn(m) %></td>

                                                <%
                                                loopmd = startdatoMD
                                                looprealaar = startdatoYY
                                                loopaar = 1

                                                dim timerRealiseret
                                                redim timerRealiseret(12,2)
                                                strSQL = "SELECT sum(timer) as sumtimer, tdato FROM timer WHERE tmnr = "& medarbejderid(m) &" AND taktivitetid = "& aktidFC & " AND ((month(tdato) >= "& startdatoMD &" AND year(tdato) = "& startdatoYY &") OR (month(tdato) <= "& slutdatoMD &" AND year(tdato) = "& slutdatoYY &")) "_
                                                & "GROUP BY month(tdato), year(tdato)"
                                                oRec2.open strSQL, oConn, 3
                                                while not oRec2.EOF
                                                    if year(oRec2("tdato")) = startdatoYY then
                                                        timeraar = 1
                                                    else
                                                        timeraar = 2                                       
                                                    end if

                                                    timerRealiseret(month(oRec2("tdato")), timeraar) = oRec2("sumtimer")
                                                oRec2.movenext
                                                wend
                                                oRec2.close

                                                for y = 0 TO numoffdaysorweeksinperiode

                                                    if y <> 0 then
                                                        if loopmd = 12 then
                                                            loopmd = 1
                                                            loopaar = loopaar + 1
                                                            looprealaar = looprealaar + 1
                                                        else
                                                            loopmd = loopmd + 1
                                                        end if
                                                    end if
                                                   
                                                    if ((month(cdate(dagsdato)) <= month(cdate(nday)) AND year(cdate(dagsdato)) = year(cdate(nday))) OR (year(cdate(dagsdato)) < year(cdate(nday)))) OR level = 1 then
				   	                                    iptype = "text"
                                                        redTimer = 1
				   	                                else 
				   	                                    iptype = "hidden"
				   	                                end if

                                                    totalMedReal = totalMedReal + formatnumber(formatnumber(timerRealiseret(loopmd, loopaar), 2))

                                                    %>
                                                    <td>
                                                        <input type="hidden" name="FM_dato" value="<%="1-"& loopmd &"-"& looprealaar%>" />
                                                        <input type="hidden" name="FM_medarbid" value="<%=medarbejderid(m) %>" />
                                                        <input type="hidden" name="FM_jobid" value="<%=oRec("jobid") %>" />
                                                        <input type="hidden" name="FM_aktid" value="<%=aktidFC %>" />
                                                        <input type="hidden" name="FM_aktid_old" value="<%=aktidFC %>" />

                                                        <input style="text-align:right; width:60px;" type="<%=iptype %>" name="FM_timer" value="0" class="form-control input-small" />

                                                        <%if redTimer = 0 then %>
                                                        <input style="text-align:right; width:60px;" type="text" class="form-control input-small" value="0" readonly /> <!-- Bliver kun brugt til at vise timerne, uden man kan redigere -->
                                                        <%end if %> 

                                                        <input type="hidden" name="FM_timer" id="FM_timer" value="#" />
                                                    </td>

                                                    <td><input type="text" class="form-control input-small" value="<%=formatnumber(timerRealiseret(loopmd, loopaar), 2) %>" readonly style="text-align:right; width:60px;" /></td>
                                                    <%

                                                next
                                                %>

                                                <td style="text-align:right">0,00</td>
                                                <td style="text-align:right"><%=formatnumber(totalMedReal, 2) %></td>
                                                <td style="text-align:right">
                                                    <%
                                                    if totalMedReal > 0 then
                                                        response.Write formatnumber(-totalMedReal, 2)
                                                    else
                                                        response.Write formatnumber(totalMedReal, 2)
                                                    end if
                                                    %>
                                                </td>

                                            </tr>
                                        <%
                                        end if

                                    end if
                                next
                            end if
                            %>




                           <% 
                            jobFCtot = 0
                            jobRLtot = 0
                            for y = 0 TO numoffdaysorweeksinperiode
                                response.write "<input type='hidden' value='"&formatnumber(totRLperiode(y),2)&"' class='totRealPer' id='"&oRec("jobid")&"_"&y&"' />"
                                response.write "<input type='hidden' value='"&formatnumber(totFCperiode(y),2)&"' class='totFCPer' id='"&oRec("jobid")&"_"&y&"' />"

                               jobFCtot = jobFCtot + totFCperiode(y)
                               jobRLtot = jobRLtot + totRLperiode(y)
                               
                               grandtotIperFC(y) = grandtotIperFC(y) + totFCperiode(y)
                               grandtotIperRL(y) = grandtotIperRL(y) + totRLperiode(y)
                            next 

                            response.Write "<input type='hidden' class='jobFCtot' value='"& formatnumber(jobFCtot,2) &"' id='"&oRec("jobid")&"' />"
                            response.Write "<input type='hidden' class='jobRLtot' value='"& formatnumber(jobRLtot,2) &"' id='"&oRec("jobid")&"' />"

                            jobtotSaldo = jobFCtot - jobRLtot
                            response.Write "<input type='hidden' class='jobtotSaldo' value='"& formatnumber(jobtotSaldo,2) &"' id='"&oRec("jobid")&"' />"

                            %>

                            <%
                            oRec.movenext
                            wend
                            oRec.close
                            %>

                            <%
                            saldoIalt = forecastIalt - realisertIalt
                            %>

                            <tr>
                                <th style="text-align:right;">Grandtotal</th>
                                
                                <%for y = 0 TO numoffdaysorweeksinperiode %>
                                    <th style="text-align:right;"><b><%=formatnumber(grandtotIperFC(y),2) %></b></th>
                                    <th style="text-align:right;"><b><%=formatnumber(grandtotIperRL(y),2) %></b></th>
                                <%next %>

                                <th style="text-align:right;"><%=formatnumber(forecastIalt, 2) %></th>
                                <th style="text-align:right;"><%=formatnumber(realisertIalt, 2) %></th>
                                <th style="text-align:right;"><%=formatnumber(saldoIalt, 2) %></th>
                            </tr>

                        </table>

                            <div class="row">
                                <div class="col-lg-12"><button type="submit" class="btn btn-success btn-sm pull-right"><b><%=resbelaeg_txt_049 %></b></button></div>
                            </div>

                        </form>
                        <%end if %>

                        <%if media <> "eksport" then %>
                            <%if media <> "print" AND media <> "chart" AND ddcView <> 1 then%>
                                <div class="row">
                                    <div class="col-lg-5">
                                        *)<%=" " & resbelaeg_txt_056 %><br />
                                        f:<%=" " & resbelaeg_txt_057 %><br />
                                        r:<%=" " & resbelaeg_txt_058 %><br />
                                        s:<%=" " & resbelaeg_txt_059 %><br />
                                        n:<%=" " & resbelaeg_txt_060 %><br /><br />
                                    </div>
                                  <!--  <div class="col-lg-7">
                                        <button type="submit" class="btn btn-success btn-sm pull-right"><b><%=resbelaeg_txt_049 %></b></button>
                                    </div> -->
                                    <input id="ymax" name="ymax" value="<%=y-1 %>" type="hidden" />
                                </div>
                            <%end if %>
                        <%end if %>

                  <!--  </form> -->



                    <%
                        '***************** Eksport *************************'
                        if media = "eksport" OR media = "chart" then
                            if vis <> 1 then 'PIVOT
		            
                                csvTxtTop = resbelaeg_txt_090&";"&resbelaeg_txt_167&";"&resbelaeg_txt_167&";"
                                'csvTxtTop = "Medarbejder;Initialer;Sidst opd.;" 

                                select case periodesel 
                                    case 6,12
                                    csvTxtTop = csvTxtTop & resbelaeg_txt_169 &";" 
                                    case else
                                    csvTxtTop = csvTxtTop & resbelaeg_txt_170 &";"
                                end select

		
                                if cint(visRamme) = 1 then 
                                    csvTxtTop = csvTxtTop & resbelaeg_txt_171&";"&resbelaeg_txt_172&";"
                                end if

                                csvTxtTop = csvTxtTop & resbelaeg_txt_073 &";"&resbelaeg_txt_173&";"

                                'if cint(visAkt) = 1 then ALTID MED I UDTRÆK
                                csvTxtTop = csvTxtTop & resbelaeg_txt_174 &";"&resbelaeg_txt_175&";"
                                'end if

                                csvTxtTop = csvTxtTop & resbelaeg_txt_176 &";"&resbelaeg_txt_177&";"&resbelaeg_txt_178&";"&resbelaeg_txt_177&";"&resbelaeg_txt_179&";"&resbelaeg_txt_180&";"
                                csvTxtTop = csvTxtTop & resbelaeg_txt_181 &";"
                                csvTxtTop = csvTxtTop & resbelaeg_txt_182 &";"&resbelaeg_txt_183&";"&resbelaeg_txt_184&";"
                    
                                if cint(visStatus) = 1 then
                                csvTxtTop = csvTxtTop & resbelaeg_txt_101 &";"&resbelaeg_txt_185&" "&resbelaeg_txt_186
                                end if

                            else '** Alm eksport


                                if media = "chart" then
                                    csvTxtTop = resbelaeg_txt_090 & ";" 
                                    csvTxtTop = csvTxtTop & csvTxtTopMd & "xx99123sy#z"            
                                else
                                    csvTxtTop = resbelaeg_txt_179 &";"&resbelaeg_txt_180&";"&resbelaeg_txt_073&";"&resbelaeg_txt_173&";"           
                                    'if cint(visAkt) = 1 then ALTID MED I UDTRÆK
                                    csvTxtTop = csvTxtTop & resbelaeg_txt_174 &";"&resbelaeg_txt_175&";"
                                    'end if
                                    csvTxtTop = csvTxtTop & resbelaeg_txt_090 &";"&resbelaeg_txt_167&";"&resbelaeg_txt_168&";" 
                                    'select case periodesel 
                                    'case 6,12
                                    'csvTxtTop = csvTxtTop & "Norm. timer pr. md (gns.);" 
                                    'case else
                                    csvTxtTop = csvTxtTop & resbelaeg_txt_187 &";"
                                    'end select
                                    csvTxtTop = csvTxtTop & csvTxtTopMd 

                                    if cint(visStatus) = 1 then
                                        csvTxtTop = csvTxtTop & resbelaeg_txt_188 &";"&resbelaeg_txt_189&";"&resbelaeg_txt_190&";"&resbelaeg_txt_191&";"&resbelaeg_txt_101&";"&resbelaeg_txt_192&"; "&resbelaeg_txt_194&"; "&resbelaeg_txt_193&" "
                                    end if   
         
                                end if                
                            end if


                            if media = "chart" then               
                               ekspTxt = csvTxtTop & csvTxtA 
                               ekspTxt = replace(ekspTxt, "xx99123sy#z", vbcrlf)
                            else
                               ekspTxt = csvTxtTop & csvTxt & csvTxtTotal
                               ekspTxt = replace(ekspTxt, "xx99123sy#z", vbcrlf)
                            end if


                            call TimeOutVersion()

                            filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	                        filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)

                            Set objFSO = server.createobject("Scripting.FileSystemObject")

                            if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\ressource_belaeg_jbpla.asp" then
            							
		                        Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\"& toVer &"\inc\log\data\ressexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
		                        Set objNewFile = nothing
		                        Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\"& toVer &"\inc\log\data\ressexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
	                        else            		
		                        Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\ressexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, false)
		                        Set objNewFile = nothing
		                        Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\ressexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8, -1)            		
	                        end if

                            file = "ressexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"

                            objF.WriteLine(ekspTxt)

                            %>
                                <form>
                                    <input type="hidden" value="<%=file %>" name="FM_filename" id="FM_filename" />
                                </form>
                            <%
                            objF.close
                            objF.close


                            if media = "chart" then
                            %>
                                <script src="../timereg/inc/res_hc_jav.js"></script>
                            <%  
                               maddVal = 1150
                               for m = 0 to UBOUND(intMids) 
                               maddVal = maddVal + 50 
                               next
                            %>
                                <div id="container" style="width: <%=maddVal%>px; height: 400px; border:0px;"></div>
                            <%
                            else
                            %>
                                <table border=0 cellspacing=1 cellpadding=0 width="200">
	                            <tr><td valign=top bgcolor="#ffffff" style="padding:5px;">
	                            <img src="../ill/outzource_logo_200.gif" />
	                            </td>
	                            </tr>
	                            <tr>
	                            <td valign=top bgcolor="#ffffff" style="padding:5px 5px 5px 15px;">
	                            <a href="../inc/log/data/<%=file%>" class=vmenu target="_blank"><%=resbelaeg_txt_062 %> >></a> <!-- onClick="Javascript:window.close()"-->
	                            </td></tr>
	                            </table>
                            <%
                            end if

                            Response.end
	                        Response.redirect "../inc/log/data/"& file &""
                                
                        end if 'export






                        if media <> "print" AND media <> "eksport" AND media <> "chart" then
                        %>
                            <%if ddcView <> 1 then %>
                            <br /><br />
                            <div class="row">
                                <div class="col-lg-2"><b><%=resbelaeg_txt_066 %></b></div>
                            </div>
                        <%
                            lnk = "periodeSel="&periodeSel&"&FM_medarb="&thisMiduse&"&FM_medarb_hidden="&thisMiduse&"&jobselect="&jobnavn&"&mdselect="&monththis&"&aarselect="&yearThis&"&FM_vis_job_u_fc="&vis_job_u_fc&"&FM_vis_job_fc_neg="&vis_job_fc_neg&"&FM_hideEmptyEmplyLines="&hideEmptyEmplyLines&"&FM_vis_simpel="&vis_simpel&"&FM_vis_kolonne_simpel="&vis_kolonne_simpel&"&FM_vis_kolonne_simpel_akt="&vis_kolonne_simpel_akt&"&FM_sog="&sogTxt
                        %>
                            <form action="ressource_belaeg_jbpla.asp?media=eksport&<%=lnk%>" method="post" target="_blank"> 
                                <div class="row">
                                    <div class="col-lg-2"><input id="Submit5" type="submit" value="<%=resbelaeg_txt_063 %>" class="btn btn-default btn-sm" /></div>
                                </div>
                            </form>

                            <form action="ressource_belaeg_jbpla.asp?media=eksport&<%=lnk%>&vis=1" method="post" target="_blank">  
                                <div class="row">
                                    <div class="col-lg-2"><input id="Submit6" type="submit" value="<%=resbelaeg_txt_064 %>" class="btn btn-default btn-sm" /><br />
                                        <input type="checkbox" name="FM_expvisreal" value="1" <%=expvisrealCHK %> /><%=" " & resbelaeg_txt_065 %>
                                    </div>
                                </div>
                            </form>

                            <br />

                            <form action="ressource_belaeg_jbpla.asp?media=print&<%=lnk%>" method="post" target="_blank">
                                <div class="row">
                                    <div class="col-lg-2"><input id="Submit7" type="submit" value="<%=resbelaeg_txt_076 %>" class="btn btn-default btn-sm" /></div>
                                </div>
                            </form>
                        <%
                            if cint(level) = 1 then
                            %>
                                <br />
                                <form action="ressource_belaeg_jbpla.asp?func=overfor&<%=lnk%>" method="post"> 
                                    <div class="row">
                                        <div class="col-lg-12"><input type="checkbox" value="1" name="FM_overfor" /> <%=resbelaeg_txt_077 %> <span style="color:#999999; font-size:9px;">(<%=resbelaeg_txt_195 %>)</span></div>
                                    </div>
                                    <div class="row">
                                        <%if periodeSel = 3 then %>
                                            <div class="col-lg-1"><%=resbelaeg_txt_198 %>:</div>
                                            <div class="col-lg-2">
                                                <select name="FM_overforfra" class="form-control input-small">
                                                    <%for u = 1 to 53 
                                            

                                                        call thisWeekNo53_fn("1/"& monththis & "/" & yearthis)

                                                        if thisWeekNo53 = u - 1 then 'datepart("ww", now, 2,2)
                                                        uSEL = "SELECTED"
                                                        else
                                                        uSel = ""
                                                        end if%>
                                                        <option value="<%=u %>" <%=uSel %>>Uge: <%=u & " - "& datepart("yyyy", now, 2,2)%></option>
                                                    <%next %>
                                                </select>
                                            </div>
                                            <div class="col-lg-1"><%=resbelaeg_txt_197 %></div>
                                        <%else %>
                                            <div class="col-lg-1"><%=resbelaeg_txt_198 %>:</div>
                                            <div class="col-lg-1">
                                                 <select name="FM_overforfra" class="form-control input-small">
                                                    <%for u = 1 to 12 
                                        
                                                        if cint(monththis) = u then 'datepart("m", now, 2,2)
                                                        uSEL = "SELECTED"
                                                        else
                                                        uSel = ""
                                                        end if
                                                        
                                                        select case u
                                                            case 1
                                                            monthTxt = resbelaeg_txt_149
                                                            case 2
                                                            monthTxt = resbelaeg_txt_150
                                                            case 3
                                                            monthTxt = resbelaeg_txt_151
                                                            case 4
                                                            monthTxt = resbelaeg_txt_152
                                                            case 5
                                                            monthTxt = resbelaeg_txt_153
                                                            case 6
                                                            monthTxt = resbelaeg_txt_154
                                                            case 7
                                                            monthTxt = resbelaeg_txt_155
                                                            case 8
                                                            monthTxt = resbelaeg_txt_156
                                                            case 9
                                                            monthTxt = resbelaeg_txt_157
                                                            case 10
                                                            monthTxt = resbelaeg_txt_158
                                                            case 11
                                                            monthTxt = resbelaeg_txt_159
                                                            case 12
                                                            monthTxt = resbelaeg_txt_160
                                                        end select
                                                    %>
                                                    <option value="<%=u %>" <%=uSel %>><%=monthTxt &" "& yearthis %></option>
                                                    <%next %>
                                                </select>
                                            </div>
                                            <div class="col-lg-2"><%=resbelaeg_txt_199 %></div>
                                        <%end if %>
                                    </div>
                                    <div class="row">
                                        <div class="col-lg-2">
                                            <input type="hidden" name="FM_overforfraAar" value="<%=yearthis %>" />
                                            <input id="Submit8" type="submit" value="<%=resbelaeg_txt_077 %>" class="btn btn-default btn-sm" />
                                        </div>
                                    </div>
                                    
                                    
                                </form>
                    <%
                            end if 'cint level
                            end if ' <> ddcViwe

                        
                        else
                            Response.Write("<script language=""JavaScript"">window.print();</script>")
                        end if 'print
                    %>


                    <style>

                        /* The Modal (background) */
                        .modal {
                            display: none; /* Hidden by default */
                            position: fixed; /* Stay in place */
                            z-index: 1; /* Sit on top */
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
                            width: 1000px;
                            height: 150px;
                        }
   
                    </style>


                    <!-- Opret ny forecast pop up -->
                    <div id="createnew" class="modal">
                        <div class="modal-content">
                            <span style="padding-left:95%; font-size:150%; cursor:pointer; color:darkgray;" id="createnewclose" class="fa fa-times"></span>

                            <br /><br />
                            <table class="table table-striped table-bordered">
                                <thead>
                                    <tr>
                                        <th style="width:150px;">Job</th>
                                        <th style="width:150px;">Aktivitet</th>
                                        <th style="width:150px;">Medarbejder</th>
                                        <th style="text-align:center;">Jan</th>
                                        <th style="text-align:center;">Feb</th>
                                        <th style="text-align:center;">Mar</th>
                                        <th style="text-align:center;">Apr</th>
                                        <th style="text-align:center;">Maj</th>
                                        <th style="text-align:center;">Jun</th>                                       
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td><input type="text" class="form-control input-small" placeholder="Søg job" /></td>
                                        <td><input type="text" class="form-control input-small" placeholder="Søg aktivitet" /></td>
                                        <td><input type="text" class="form-control input-small" placeholder="Søg medarbejder" /></td>
                                        <th><input type="text" class="form-control input-small" /></th>
                                        <th><input type="text" class="form-control input-small" /></th>
                                        <th><input type="text" class="form-control input-small" /></th>
                                        <th><input type="text" class="form-control input-small" /></th>
                                        <th><input type="text" class="form-control input-small" /></th>
                                        <th><input type="text" class="form-control input-small" /></th>

                                        <th><span class="btn btn-default btn-sm"><b>Opret</b></span></th>

                                    </tr>
                                </tbody>
                            </table>
                        </div>
                    </div>





                </div>

            </div>
        </div>




    </div>
</div>



<!--#include file="../inc/regular/footer_inc.asp"-->