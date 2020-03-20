

<%
'strConnect = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_oko;"

 strConnect = "timeout_tia64"

Response.write strConnect & "<br><br>"

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject ("ADODB.Recordset")
Set oRec2 = Server.CreateObject ("ADODB.Recordset")
Set oCmd = Server.CreateObject ("ADODB.Command")

oConn.open strConnect



dim mid
x = 0


strSQL = "SELECT init FROM medarbejdere WHERE mansat = 1 ORDER BY mid LIMIT 0,10"

Response.Write strSQL & "<hr>"
Response.flush

oRec.open strSQL, oConn, 3
while not oRec.EOF 
				
    
    'response.redirect("https://timeout.cloud/timeout_xp/wwwroot/ver2_14/timereg_net/importer_med_tia_nav.aspx?no="& oRec("init"))
								
              

x = x + 1
oRec.movenext
wend
oRec.close



%>
<br><br><br>
Antal records: <%=x %><br />
Opdatering gennemført!
