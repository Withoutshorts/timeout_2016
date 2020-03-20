
<!-- Version 01-06-2018 --->



<%

public simple_joblist_jobnavn, simple_joblist_jobnr, simple_joblist_kunde, simple_joblist_ans, simple_joblist_salgsans, simple_joblist_fomr, simple_joblist_prg, simple_joblist_stdato, simple_joblist_sldato, simple_joblist_status, simple_joblist_tidsforbrug, simple_joblist_budgettid
    function simplejobfelter()

            simple_joblist_jobnavn = 0
            simple_joblist_jobnr = 0
            simple_joblist_kunde = 0
            simple_joblist_ans = 0 
            simple_joblist_salgsans = 0
            simple_joblist_fomr = 0
            simple_joblist_prg = 0
            simple_joblist_stdato = 0
            simple_joblist_sldato = 0
            simple_joblist_status = 0
            simple_joblist_tidsforbrug = 0
            simple_joblist_budgettid = 0


            strSQL = "SELECT simple_joblist_jobnavn, simple_joblist_jobnr, simple_joblist_kunde, simple_joblist_ans, simple_joblist_salgsans, simple_joblist_fomr, simple_joblist_prg, simple_joblist_stdato, simple_joblist_sldato, simple_joblist_status, simple_joblist_tidsforbrug, simple_joblist_budgettid "_
            &" FROM licens WHERE id = 1"
            oRec.open strSQL, oConn, 3
                if not oRec.EOF then
                    simple_joblist_jobnavn = oRec("simple_joblist_jobnavn")
                    simple_joblist_jobnr = oRec("simple_joblist_jobnr")
                    simple_joblist_kunde = oRec("simple_joblist_kunde")
                    simple_joblist_ans = oRec("simple_joblist_ans")
                    simple_joblist_salgsans = oRec("simple_joblist_salgsans")
                    simple_joblist_fomr = oRec("simple_joblist_fomr")
                    simple_joblist_prg = oRec("simple_joblist_prg")
                    simple_joblist_stdato = oRec("simple_joblist_stdato")
                    simple_joblist_sldato = oRec("simple_joblist_sldato")
                    simple_joblist_status = oRec("simple_joblist_status")
                    simple_joblist_tidsforbrug = oRec("simple_joblist_tidsforbrug")
                    simple_joblist_budgettid = oRec("simple_joblist_budgettid")
                end if
            oRec.close

    end function



public strFomr_navn, strFomr_id, visJobFomr
function forretningsomrJobId(jobid)


    '*** Forretningsområder **' 
    strFomr_navn = ""
    strFomr_id = ""
    visJobFomr = 0


    '**** Job ****'
    strSQLfrel = "SELECT for_fomr, fomr.navn FROM fomr_rel "_
    &" LEFT JOIN fomr ON (fomr.id = for_fomr) WHERE for_jobid = "& jobid & " AND for_aktid = 0 GROUP BY for_fomr"

    'Response.Write strSQLfrel
    'Response.flush
    f = 0
    fj = 0
    oRec3.open strSQLfrel, oConn, 3
    while not oRec3.EOF

    if f = 0 then
    strFomr_navn = "<b>Job:</b> "
    end if

    strFomr_navn = strFomr_navn & oRec3("navn") & ", " 
    strFomr_id = strFomr_id &",#"& oRec3("for_fomr") & "#"

    if instr(strFomr_rel, "#"&oRec3("for_fomr")&"#") <> 0 then
    visJobFomr = 1
    else
    visJobFomr = visJobFomr
    end if

    f = f + 1
    fj = fj + 1
    oRec3.movenext
    wend
    oRec3.close




    '************ Aktiviteter *****'
        strSQLfrel = "SELECT for_fomr, fomr.navn FROM fomr_rel "_
    &" LEFT JOIN fomr ON (fomr.id = for_fomr) WHERE for_jobid = "& jobid & " AND for_aktid <> 0 GROUP BY for_fomr"

    'Response.Write strSQLfrel
    'Response.flush
    f = 0
    oRec3.open strSQLfrel, oConn, 3
    while not oRec3.EOF

    if f = 0 and fj > 0 then
        strbreak = "<br>"
    else
        strbreak = ""
    end if

    if f = 0 then
    strFomr_navn = strFomr_navn & strbreak &"<b>"&job_txt_554&": </b>"
    end if

    strFomr_navn = strFomr_navn & oRec3("navn") & ", " 
    strFomr_id = strFomr_id &",#"& oRec3("for_fomr") & "#"

    if instr(strFomr_rel, "#"&oRec3("for_fomr")&"#") <> 0 then
    visJobFomr = 1
    else
    visJobFomr = visJobFomr
    end if

    f = f + 1
    oRec3.movenext
    wend
    oRec3.close
                            



                    



    if f <> 0 then
    len_strFomr_navn = len(strFomr_navn)
    left_strFomr_navn = left(strFomr_navn, len_strFomr_navn - 2)
    strFomr_navn = left_strFomr_navn
end if		    


end function


%>