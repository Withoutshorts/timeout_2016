

<%
'strConnect = "driver={MySQL ODBC 3.51 Driver};server=195.189.130.210; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_synergi1;"
strConnect = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_synergi1;"

	
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
	strSQLj = "SELECT j.id AS jid, j.jobstatus, j.jobnavn, j.jobslutdato, j.jobnr "_
	&" FROM job AS j "_
    &" WHERE (jobstatus = 0 AND lukkedato = '2002-01-01') ORDER BY jobnr LIMIT 10"
	
    Response.write strSQLj & "<br><br>"	
	'Response.flush
	
    	
    oRec4.open strSQLj, oConn, 3
    while not oRec4.EOF 

    'Response.write oRec4("jobnavn") & " slutdato: "& oRec4("jobslutdato") &": "& oRec4("tdato") & "<br>"
            
            strSQLtimer = "SELECT tdato, tjobnr FROM timer WHERE tjobnr = "& oRec4("jobnr") & " ORDER BY tdato DESC LIMIT 1"

            oRec3.open strSQLtimer, oConn, 3
            if not oRec3.EOF then

                    if len(trim(oRec3("tdato"))) <> 0 AND len(trim(oRec4("jobslutdato"))) <> 0 then

                    if cdate(oRec4("jobslutdato")) < cDate(oRec3("tdato")) then
                    strSQLupd = "UPDATE job SET jobslutdato = "& oRec3("tdato") & " WHERE id = "& oRec4("jid")

                    Response.Write strSQLupd & " ("& oRec4("jobslutdato") &")<br><br>"
                    'oConn.execute(strSQLupd)

                    end if
                    end if 

            end if
            oRec3.close


            Response.flush

   j = j + 1
    oRec4.movenext						
	wend 							
    oRec4.close
                             




%>
<br>
Antal job: <%=j %><br><br>
Opdatering gennemført!
