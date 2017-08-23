

<%
'strConnect = "driver={MySQL ODBC 3.51 Driver};server=195.189.130.210; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_synergi1;"
'strConnect = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_tia;"
strConnect = "timeout_tia64"
	
Response.write strConnect & "<br><br>"

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject ("ADODB.Recordset")
Set oRec2 = Server.CreateObject ("ADODB.Recordset")
Set oRec3 = Server.CreateObject ("ADODB.Recordset")
Set oRec4 = Server.CreateObject ("ADODB.Recordset")
Set oRec5 = Server.CreateObject ("ADODB.Recordset")
Set oCmd = Server.CreateObject ("ADODB.Command")

oConn.open strConnect




        

	lastKid = 0
	strSQLj = "SELECT j.id AS jid, j.jobstatus, j.jobnavn, j.jobslutdato, j.jobnr, a.navn as aktnavn, a.avarenr "_
	&" FROM job AS j "_
    &" LEFT JOIN aktiviteter a ON (a.job = j.id AND avarenr LIKE '%210') "_
    &" WHERE (jobnr = 'PS1134'"_
    &" OR jobnr =  'PS1132'"_
    &" OR jobnr =  'PS1140'"_
    &" OR jobnr =  'PS1178'"_
    &" OR jobnr =  'PS1143'"_
    &" OR jobnr =  'PS1145'"_
    &" OR jobnr =  'PS1100'"_
    &" OR jobnr =  'PS1150'"_
    &" OR jobnr =  'PS1102'"_
    &" OR jobnr =  'PS1157'"_
    &" OR jobnr =  'ABS1000'"_
    &" OR jobnr =  'ADM1005'"_
    &" OR jobnr =  'ADM300'"_
    &" OR jobnr =  'ADM330'"_
    &" OR jobnr =  'ADM600'"_
    &" OR jobnr =  'DEV1006'"_
    &" OR jobnr =  'DEV1007'"_
    &" OR jobnr =  'DEV1011'"_
    &" OR jobnr =  'DEV1040'"_
    &" OR jobnr =  'DEV1042'"_
    &" OR jobnr =  'DEV1043'"_
    &" OR jobnr =  'DEV1045'"_
    &" OR jobnr =  'PS1162'"_
    &" OR jobnr =  'PS1990'"_
    &" OR jobnr =  'PS1993'"_
    &" OR jobnr =  'PS1137') ORDER BY jobnr"

    Response.write strSQLj & "<br><br>"	
	'Response.flush
	
    	
    oRec4.open strSQLj, oConn, 3
    while not oRec4.EOF 


    Response.write "Jobnr: "& oRec4("jobnr") & " aktnavn: "& oRec4("aktnavn") & " TASKno: " & oRec4("avarenr") & "<br>"

            Response.flush

   j = j + 1
    oRec4.movenext						
	wend 							
    oRec4.close
                             




%>
<br>
Antal job: <%=j %><br><br>
Opdatering gennemført!
