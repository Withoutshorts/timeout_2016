

<%
strConnect = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_oko;"

	
Response.write strConnect & "<br><br>"

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject("ADODB.Recordset")
Set oRec2 = Server.CreateObject("ADODB.Recordset")
Set oRec3 = Server.CreateObject("ADODB.Recordset")
Set oCmd = Server.CreateObject ("ADODB.Command")

oConn.open strConnect



dim mid

strSQL = "SELECt id AS jid, jobnr FROM job WHERE beskrivelse = 'Direktor. for FødevareErhverv GUDP' ORDER BY id LIMIT 100"
'NaturErhversstyrelsen
Response.Write strSQL & "<hr>"
Response.flush

t = 0

oRec.open strSQL, oConn, 3
while not oRec.EOF 
								

                                strSQLmtypTp2 = "SELECT mid, medarbejdertype, timepris_a3 FROM medarbejdere AS m "_
                                &" LEFT JOIN medarbejdertyper AS mt ON (mt.id = m.medarbejdertype) WHERE mid > 0 ORDER BY mid "
							
                                'timepris_a2
                                Response.Write strSQLmtypTp2 & "<hr>"
                                Response.flush

                              
                                oRec2.open strSQLmtypTp2, oConn, 3
                                while not oRec2.EOF

                        
                                'tp2 = replace(oRec2("timepris_a2"), ",", ".") NAER
                                tp2 = replace(oRec2("timepris_a3"), ",", ".") 'DIREKTOR
                                
                                 
                                strSQLdeltp = "DELETE FROM timepriser WHERE jobid = "& oRec("jid") &" AND medarbid = " & oRec2("mid")
                                response.write strSQLdeltp & "<br>"
                                'oConn.execute(strSQLdeltp)

                                strSQLupdJob = "INSERT INTO timepriser (jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta) VALUES "_
                                &" ("& oRec("jid") &", 0, "& oRec2("mid") &", 2, "& tp2 &", 1)" 
                                response.write strSQLupdJob & "<br>"
                                'oConn.execute(strSQLupdJob)                        

                                strSQLupdTimer = "UPDATE timer SET timepris = " & tp2 & " WHERE tjobnr = " & oRec("jobnr") &" AND tmnr = " & oRec2("mid")
                                Response.Write strSQLupdTimer & "<br>"
                                
                                'oConn.execute(strSQLupdTimer)
                                Response.flush


                            
                                oRec2.movenext
                                wend
                                oRec2.close
                                
                               

    t = t + 1

oRec.movenext
wend
oRec.close



%>
<br><br><br>
Antal records: <%=t %><br />

<br />
<br />
Opdatering gennemført!
