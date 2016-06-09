

<%
strConnect = "driver={MySQL ODBC 3.51 Driver};server=81.19.249.35; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_epi;"
'strConnect = "driver={MySQL ODBC 3.51 Driver};server=81.19.249.35; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_dencker;"
'81.19.249.35
	
Response.write strConnect & "<br><br>"

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject ("ADODB.Recordset")
Set oRec2 = Server.CreateObject ("ADODB.Recordset")
Set oRec3 = Server.CreateObject ("ADODB.Recordset")
Set oRec4 = Server.CreateObject ("ADODB.Recordset")
Set oCmd = Server.CreateObject ("ADODB.Command")

oConn.open strConnect


strSQL = "SELECT id, medarbid, jobid, aktid FROM timepriser WHERE id <> 0 LIMIT 950000, 100000 " 'AND aktid = 0" 'AND medarbid = 31" 'aktid = 0


Response.Write strSQL & "<br><br>"
Response.flush

t = 0
oRec.open strSQL, oConn, 3
while not oRec.EOF 
								
								
                              
								    
            strSQLdel = "DELETE FROM timepriser WHERE id <> "& oRec("id") &" AND jobid =  "& oRec("jobid") &" AND aktid = "& oRec("aktid") &" AND medarbid = "& oRec("medarbid") 
            Response.Write "."' strSQLdel & "<br>"

            select case right(t, 2)
            case "00"
            Response.Write "<br>"
            Response.flush
            end select

            select case right(t, 3)
            case "000"
            Response.Write "<br><br>"& t &"<br><br>"
            Response.flush
            end select

            oConn.execute(strSQLdel)
            


                              

t = t + 1
oRec.movenext
wend
oRec.close



%>
<br><br><br>
Opdatering gennemført!
