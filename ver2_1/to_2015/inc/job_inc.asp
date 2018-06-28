
<!-- Version 01-06-2018 --->



<%
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

    if f = 0 then
    strFomr_navn = strFomr_navn & "<br><b>"&job_txt_554&": </b>"
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





function OpretAktivitet







%>