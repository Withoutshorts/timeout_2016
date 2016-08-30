

<%
'strConnect = "driver={MySQL ODBC 3.51 Driver};server=195.189.130.210; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_synergi1;"
strConnect = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_nt;"

	
Response.write strConnect & "<br><br>"

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject ("ADODB.Recordset")
Set oRec2 = Server.CreateObject ("ADODB.Recordset")
Set oRec3 = Server.CreateObject ("ADODB.Recordset")
Set oRec4 = Server.CreateObject ("ADODB.Recordset")
Set oRec5 = Server.CreateObject ("ADODB.Recordset")
Set oCmd = Server.CreateObject ("ADODB.Command")

oConn.open strConnect


%>
<!--#include file="../inc/regular/global_func.asp"-->
<%


        

	lastKid = 0
	strSQLj = "SELECT f.jobid, faknr "_
	&" FROM fakturaer AS f "_
    &" WHERE fid > 0 ORDER BY faknr"
	
    Response.write strSQLj & "<br><br>"	
	'Response.flush
	
    	
    oRec4.open strSQLj, oConn, 3
    while not oRec4.EOF 

                    if oRec4("jobid") <> 0 then

                    strSQLupd = "UPDATE job SET jobstatus = 0 WHERE id = "& oRec4("jobid")

                    Response.Write strSQLupd & " ("& oRec4("jobid") &" // Faknr = "& oRec4("faknr") &") <br>"
                    'oConn.execute(strSQLupd)


                     j = j + 1
                    end if

              

            Response.flush

  
    oRec4.movenext						
	wend 							
    oRec4.close
                             




%>
<br>
Antal job: <%=j %><br><br>
Opdatering gennemført!
