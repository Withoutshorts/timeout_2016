

<%
strConnect = "driver={MySQL ODBC 3.51 Driver};server=195.189.130.210; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_lyng;"

	
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

dim mid

strSQL = "SELECT mid, mnavn FROM medarbejdere WHERE mid = 1 AND mansat = 1 ORDER BY mnavn LIMIT 50"

Response.Write strSQL & "<hr>"
Response.flush

j = 0
oRec5.open strSQL, oConn, 3
while not oRec5.EOF 


Response.write "<u>"& oRec5("mnavn") & "</u><br>"
        

        call hentbgrppamedarb(oRec5("mid"))
	
    '*** Søger efter job først, da timereg_usejob kun indeholder de job der er aktive lige nu. 
    '*** Du kan godt være tilmeldt et job via dine projektgrupper uden det er med i timereg_usejob

	lastKid = 0
	strSQLj = "SELECT j.id AS jid, j.jobstatus, j.jobnavn, j.jobnr, j.jobknr, Kid, Kkundenavn, Kkundenr, "_
	&" useasfak FROM job j, kunder"_
    &" WHERE (jobstatus = 1 OR jobstatus = 3) AND kunder.Kid = j.jobknr  "& strPgrpSQLkri &""_
	&" GROUP BY j.id, kid ORDER BY Kkundenavn, jobnavn"
	
    Response.write strSQLj & "<br><br>"	
	Response.flush
	
    	
    oRec4.open strSQLj, oConn, 3
    while not oRec4.EOF 

    jobFandtes = 0
    
    Response.write oRec4("jobnavn") & "<br>"
            
            strSQLTU = "SELECT jobid FROM timereg_usejob WHERE jobid = "& oRec4("jid") & " AND medarb = "& oRec5("mid")
            oRec3.open strSQLTU, oConn, 3
            if not oRec3.EOF then

            jobFandtes = 1

            end if
            oRec3.close
    
            '**** Opretter job i TU **
            if jobFandtes = 0 then

              strSQLins = "INSERT INTO timereg_usejob (medarb, jobid) VALUES ("& oRec5("mid") &", "& oRec4("jid") &")"

              Response.write "<b>"& strSQLins &"</b><br>"
              'oConn.execute(strSQLins)
                                   

            end if
   j = j + 1
    oRec4.movenext						
	wend 							
    oRec4.close
                             

oRec5.movenext
wend
oRec5.close



%>
<br>
Antal job: <%=j %><br><br>
Opdatering gennemført!
