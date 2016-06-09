

<%
'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
Session.LCID = 1030

'strConnect = "driver={MySQL ODBC 3.51 Driver};server=81.19.249.35; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_epi_catitest;"
'** DK
'strConnect = "driver={MySQL ODBC 3.51 Driver};server=195.189.130.210; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_epi;"
'** NO
'strConnect = "driver={MySQL ODBC 3.51 Driver};server=195.189.130.210; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_epi_no;"
'** UK
strConnect = "driver={MySQL ODBC 3.51 Driver};server=195.189.130.210; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_epi_uk;"
	
Response.write strConnect & "<br><br>"

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject ("ADODB.Recordset")
Set oRec2 = Server.CreateObject ("ADODB.Recordset")
Set oCmd = Server.CreateObject ("ADODB.Command")

oConn.open strConnect


strSQL = "SELECT pwc.dato, pwc.konto, pwc.postext, pwc.belob, pwc.jobnr, pwc.init, pwc.extsysid, pwc.id, "_
&" j.id AS jid, serviceaft, m.mid FROM mat_pwc_import_201604 AS pwc "_
&" LEFT JOIN job AS j ON (j.jobnr = pwc.jobnr) "_
&" LEFT JOIN medarbejdere AS m ON (m.init = pwc.init) WHERE pwc.postext <> '' ORDER BY extsysid" 
' AND 
'AND postext = 'Webinterviews'
'AND postext <> 'xWebinterviews'
'extsysid <> 0
Response.Write strSQL & "<hr>"
Response.flush

'Response.end
x = 0
oRec.open strSQL, oConn, 3
while not oRec.EOF 
				 
                 if isNull(oRec("postext")) = false then
                 strNavn = replace(oRec("postext"), "'", "")				
				 else
                 strNavn = "-"
                 end if				

                 belob = replace(oRec("belob"), ".", "")
                 belob = replace(belob, ",", ".")                
                 
                 'belob = oRec("belob")

                 dblKobsPris = belob
                 dblSalgsPris = belob            

                 if isNull(oRec("jid")) = false then
                 jobid = oRec("jid")
                 else
                 jobid = 3 '** Intern tid
                 end if

                 strEditor = "PWC Import"

                 strDato = year(now) & "/" & month(now) & "/"& day(now)

                 if isNull(oRec("mid")) = false then
                    
                    '*** Ekstern materiale køb **'
                    'if len(oRec("mid")) > 4 OR (_
                    'left(oRec("mid"), 1) = 1 OR _
                    'left(oRec("mid"), 1) = 2 OR _
                    'left(oRec("mid"), 1) = 3 OR _
                    'left(oRec("mid"), 1) = 4 OR _
                    'left(oRec("mid"), 1) = 5 OR _
                    'left(oRec("mid"), 1) = 6 OR _
                    'left(oRec("mid"), 1) = 7 OR _
                    'left(oRec("mid"), 1) = 8 OR _
                    'left(oRec("mid"), 1) = 9 OR _
                    'left(oRec("mid"), 1) = 0 _
                    ') AND (left(lcase(oRec("mid")), 4) <> "intw" AND left(lcase(oRec("mid")), 3) <> "itw") then

                    if len(trim(oRec("mid"))) > 4 AND (left(lcase(oRec("mid")), 3) <> "itw" AND left(lcase(oRec("mid")), 4) <> "intw") then
                    useMid = 620
                    else
                    useMid = oRec("mid")
                    end if
                 else
                 useMid = 621 'materiale forbrug ekstern / ikke fundet
                 end if

                 usemrn = useMid

                 'regDatoSQL = year(oRec("dato")) & "/" & month(oRec("dato")) & "/"& day(oRec("dato"))

                 regDatoSQLy = right(oRec("dato"), 4)
                 regDatoSQLd = left(oRec("dato"), 2) 
                 regDatoSQLm = mid(oRec("dato"), 4,2)

                 regDatoSQL = regDatoSQLy &"/"& regDatoSQLm &"/"& regDatoSQLd

                 if isNull(oRec("serviceaft")) = false then
                 aftid = oRec("serviceaft")
                 else
                 aftid = 0
                 end if

                 if useMid <> 620 then
                 intkode = 0
                 else
                 intkode = 2
                 end if

                 intValuta = 1
                 bilagsnr = 9900&oRec("extsysid")
                 dblKurs = 100
                 personlig = 0

                 extsysid = oRec("extsysid")

                 extsysid = 9900&extsysid 

			    strSQL = "INSERT INTO materiale_forbrug "_
				&" (matid, matantal, matnavn, matvarenr, matkobspris, matsalgspris, matenhed, jobid, "_
				&" editor, dato, usrid, matgrp, forbrugsdato, serviceaft, valuta, intkode, bilagsnr, kurs, personlig, extsysid) VALUES "_
				&" (0, 1, '"& strNavn &"', '0', "_
				&" "& dblKobsPris &", "& dblSalgsPris &", 'Stk.', "& jobid &", "_
				&" '"& strEditor &"', '"& strDato &"', "& usemrn &", "_
				&" 0, '"& regDatoSQL &"', "& aftid &", "& intValuta &", "_
				&" "& intkode &", '"& bilagsnr &"', "& dblKurs &", "& personlig &", "& extsysid &")"
				
				Response.Write strSQL & "<br>"
				Response.flush
				
                'oConn.execute(strSQL)

                             
x = x + 1
oRec.movenext
wend
oRec.close



%>
<br><br><br>
Opdatering gennemført!<br />
<%=x %> records.
