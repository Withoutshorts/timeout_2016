

<%
strConnect = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_epi2017;"

	
Response.write strConnect & "<br><br>"

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject ("ADODB.Recordset")
Set oRec2 = Server.CreateObject ("ADODB.Recordset")
Set oCmd = Server.CreateObject ("ADODB.Command")

oConn.open strConnect



dim mid
x = 0

'grpn = "Letbane"
strSQL = "SELECT mid FROM medarbejdere WHERE (mid > 0 AND mansat = 3 AND left(init, 3) = 'INO')" 'AND mland = '" & grpn & "'"

Response.Write strSQL & "<hr>"
Response.flush

oRec.open strSQL, oConn, 3
while not oRec.EOF 
								
								
                                mid = oRec("mid")
                                dontinsert = 0
                                pgrpid = 0
                
                                strSQLselect = "SELECT id FROM projektgrupper WHERE id = 10" 'navn LIKE '"& grpn &"%'" 
                                oRec2.open strSQLselect, oConn, 3
                                if not oRec2.EOF then

                                pgrpid = oRec2("id")                           

                                end if
                                oRec2.close

                              
                    

                                'strSQLselect = "SELECT medarbejderid FROM progrupperelationer WHERE medarbejderid = "& mid & " AND projektgruppeid <> 10" 
                                'oRec2.open strSQLselect, oConn, 3
                                'if not oRec2.EOF then

                                'dontinsert = 1
                                'end if
                                'oRec2.close

                                if dontinsert = 0 AND pgrpid <> 0 then
								strSQLm = "INSERT INTO progrupperelationer (projektgruppeid, medarbejderid, teamleder) VALUES ("& pgrpid &", "& mid &", 0)"
                                'strSQLm = "DELETE FROM progrupperelationer WHERE projektgruppeid = 15"
                                Response.Write strSQLm & "<br>"
                                Response.flush
                                'oConn.execute(strSQLm)
                                end if


x = x + 1
oRec.movenext
wend
oRec.close



%>
<br><br><br>
Antal records: <%=x %><br />
Opdatering gennemført!
