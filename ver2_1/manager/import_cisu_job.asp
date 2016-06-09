

<%
strConnect = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_cisu;"

	
Response.write strConnect & "<br><br>"

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject("ADODB.Recordset")
Set oRec2 = Server.CreateObject("ADODB.Recordset")
Set oRec3 = Server.CreateObject("ADODB.Recordset")
Set oCmd = Server.CreateObject ("ADODB.Command")

oConn.open strConnect



dim mid

timerIndlast = 0
tnavnUse = "Søren Asboe Jørgensen"

strSQL = "SELECT t.id As extsysid, t.editor, tdato, aktnavn, timer, m.mid, mnavn FROM timer_import_temp AS t "_
&" LEFT JOIN medarbejdere AS m ON (m.mnavn = t.editor)" _
&" WHERE t.id > 0 AND m.mnavn IS NOT NULL AND m.mnavn = '"& tnavnUse &"' ORDER BY t.id"

'AND aktnavn = '52. Fravær - Barnets første sygedag'
Response.Write strSQL & "<hr>"
Response.flush

t = 0

oRec.open strSQL, oConn, 3
while not oRec.EOF 
								

     jobkundefundet = 0

     meId = oRec("mid")
     meNavn = oRec("mnavn")

     dlbTimer = oRec("timer")
     dlbTimer = replace(dlbTimer, ",", ".")

     cdDatoSQL = year(oRec("tdato")) & "/" & month(oRec("tdato")) & "/" & day(oRec("tdato"))

     aktFakturerbar = 1

    timerkom = ""
    intTimepris = 400

    jobFastPris = 1

     kostpris = 0
     jobSeraft = 0
     intValuta = 1
     kurs = 100 
     intTempImpId = oRec("extsysid")
     importFrom = 50

							
                                '*** Henter konvertering OLD NEW
                                 strSQLjoboldnew = "SELECT newname, aktivitet FROM job_oldnew WHERE oldname = '"& oRec("aktnavn") & "'"
                                 oRec2.open strSQLjoboldnew, oConn, 3
                                 if not oRec2.EOF then 
                                            
                    
                                    aktNavnAcces = oRec2("aktivitet")
                                    jobkundefundet = 0


                                    strSQLjob = "SELECT j.id, j.jobnavn, j.jobnr, jobknr, k.kkundenavn, k.kkundenr, a.id as aid, "_
                                    &" a.navn as anavn, a.fakturerbar "_
                                    &" FROM job AS j "_
                                    &" LEFT JOIN kunder AS k ON (k.kid = jobknr)"_
                                    &" LEFT JOIN aktiviteter AS a ON (a.job = j.id AND a.navn = '"& oRec2("aktivitet") &"')"_
                                    &" WHERE left(jobnavn, 20) = '"& left(oRec2("newname"), 20) &"'"


                                    response.write "<br><br>"& strSQLjob & ":<br>"
                                    response.flush

                                    oRec3.open strSQLjob, oConn, 3
                                    if not oRec3.EOF then 


                                    kNavn = oRec3("kkundenavn")
                                    kNr = oRec3("kkundenr")
    

                                    jobNavn = oRec3("jobnavn")
                                    intJobNr = oRec3("jobnr")
        
                                    aktNavn = oRec3("anavn")
                                    aktId = oRec3("aid")
    
                                    aktFakturerbar = oRec3("fakturerbar")

                                    jobkundefundet = 1
            
    
                                    end if
                                    oRec3.close                        
        
    
                                end if
                                oRec2.close
    
    
    
    
                                if cint(jobkundefundet) = 1 then

    
                                if len(trim(aktId)) = 0 AND kNr = 1 then
                            
                                aktId = 96
                                aktNavn = aktNavnAcces

                                end if

                               
                                if len(trim(aktNavn)) <> 0 then
                        

                                 '*** Indlæser Timer ***'
                                strSQLupdTimer = "INSERT INTO timer " _
                                & "(" _
                                & " timer, tfaktim, tdato, tmnavn, tmnr, tjobnavn, tjobnr, tknavn, tknr, timerkom, " _
                                & " TAktivitetId, taktivitetnavn, Taar, TimePris, TasteDato, fastpris, tidspunkt, " _
                                & " editor, kostpris, seraft, " _
                                & " valuta, kurs, extSysId, sttid, sltid, origin" _
                                & ") " _
                                & " VALUES " _
                                & " (" _
                                & dlbTimer & ", " & aktFakturerbar & ", " _
                                & "'" & cdDatoSQL & "', " _
                                & "'" & meNavn & "', " _
                                & meID & ", " _
                                & "'" & left(jobNavn,50) & "', " _
                                & "'" & intJobNr & "', " _
                                & "'" & kNavn & "'," _
                                & kNr & ", " _
                                & "'" & timerkom & "', " _
                                & aktId & ", " _
                                & "'" & aktNavn & "', " _
                                & Year(Now) & ", " _
                                & intTimepris & ", " _
                                & "'" & Year(Now) & "/" & Month(Now) & "/" & Day(Now) & "', " _
                                & jobFastPris & ", " _
                                & "'00:00:01', " _
                                & "'Access Import', " _
                                & Replace(kostpris, ",", ".") & ", " _
                                & jobSeraft & ", " _
                                & intValuta & ", " _
                                & Replace(kurs, ",", ".") & ", '" & intTempImpId & "', '00:00:00', '00:00:00', " & importFrom & ")"

                                Response.Write strSQLupdTimer& "<br>"
                                
                                'oConn.execute(strSQLupdTimer)
                                Response.flush

                                timerIndlast = timerIndlast + oRec("timer")


                                else

                                f = f + 1
        
                                fejllog = fejllog & ", "& intTempImpId

                                fejllog = fejllog & "<br>tdato: "& tdato & ", aktnavn:"& aktnavn &", timer: "& timer & "<br><br>"

                                end if
                            

                                end if
                                
                               

    t = t + 1

oRec.movenext
wend
oRec.close



%>
<br><br><br>
<b><%=tnavnUse %></b><br />
Antal records: <b><%=t %></b>, timerindlæst: <b><%=timerIndlast %></b> timer<br />
Fejl: <%=f %>
fejllog: <br />
<%=fejllog %>


<br />
<br />
Opdatering gennemført!
