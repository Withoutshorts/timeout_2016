

<%
'strConnect = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_oko;"
 'strConnect = "timeout_oko64"
 strConnect = "timeout_cflow64"

Response.write strConnect & "<br><br>"

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject ("ADODB.Recordset")
Set oRec2 = Server.CreateObject ("ADODB.Recordset")
Set oCmd = Server.CreateObject ("ADODB.Command")

oConn.open strConnect



dim mid




'dim x,y
'x = "monitor_job"
'for each x in Request.Cookies
'  response.write("<p>")
'  if Request.Cookies(x).HasKeys then
'    for each y in Request.Cookies(x)
'      response.write(x & ":" & y & "=" & Request.Cookies(x)(y))
'      response.write("<br>")
'    next
'  else
'    Response.Write(x & "=" & Request.Cookies(x) & "<br>")
'  end if
'  response.write "</p>"
'next


x = 0

strSQL = "SELECT mid, medarbejdertype, mnavn FROM medarbejdere WHERE mid > 0"

Response.Write strSQL & "<hr>"
Response.flush

oRec.open strSQL, oConn, 3
while not oRec.EOF 
								
								
                mid = oRec("mid")
                jobidC = ""

                'if request.Cookies("monitor_job")(mid).HasKeys = true then
                jobidC = request.Cookies("monitor_job")(""&mid&"")
                aktidC = request.Cookies("monitor_akt")(""&mid&"")
                'end if


                'if jobidC <> "" then
                strSQLins = "INSERT INTO login_historik_aktivejob_rel SET (lha_jobid,lha_mid) VALUES ("& jobidC &" ,"& mid &")"
                'oConn.execute(strSQLins)
                response.write strSQLins & "<br>"
                response.Flush
                'end if

               

x = x + 1
oRec.movenext
wend
oRec.close



%>
<br><br><br>
Antal records: <%=x %><br />
Opdatering gennemført!
