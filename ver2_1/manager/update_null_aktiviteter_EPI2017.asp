
<%
function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
	end function


'*** Opdaterer timer, tfaktim til 0, 1, 2 ***
'strConnect = "driver={MySQL};server=192.168.1.33; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_wps;"
strConnect = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_epi2017;"

'strConnect = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_epi_BAK;"
  
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=root;pwd=;database=timeout_intranet;"
'strConnect = "mySQL_timeOut_intranet"
Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject ("ADODB.Recordset")
Set oRec2 = Server.CreateObject ("ADODB.Recordset")
Set oCmd = Server.CreateObject ("ADODB.Command")
	
oConn.open strConnect


strSQLjob = "SELECT t.id, j.jobnr, aktnavn, aktid, j.id AS jobid FROM timerjan2017flytnull t "_
&" LEFT JOIN job j ON (j.jobnr = t.jobnr) WHERE t.jobnr <> '21544' AND aktid <> 190245 AND overfort = 0 ORDER BY id"

oRec2.open strSQLjob, oConn, 3
While not oRec2.EOF 


Response.write "<br><br><h4>"& oRec2("jobnr") &"</h4>"





       if oRec2("jobid") <> 0 then

        strSQLaktins = "INSERT INTO aktiviteter (id, navn, job, fakturerbar, "_
        & "projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7,"_
        & "projektgruppe8, projektgruppe9, projektgruppe10, aktstatus, budgettimer, aktbudget, aktbudgetsum, aktstartdato, aktslutdato, aktkonto, fase, avarenr, fomr, sortorder, antalstk "_
        & " ) VALUES "_
        & " ("& oRec2("aktid")&", '" & oRec2("aktnavn") & "', " & oRec2("jobid") & ", 1,"_
        & " 1,1,1,1,1,1,1,1,1,1,0,0,0,0,'2017-01-01', "_
        & "'2017-01-01', 0, '', '', 0, 100, 0)"

        response.write strSQLaktins  & "<br>"
        Response.flush

        'oConn.execute(strSQLaktins)


        strSQLjobo = "UPDATE timerjan2017flytnull SET overfort = 1 WHERE id  = " & oRec2("id")
        response.write strSQLjobo & "<br><br>"
        'oConn.execute(strSQLjobo) 
        Response.flush

        end if


x = x + 1



oRec2.movenext
wend
oRec2.close
		
%>
<br>
X: <%=x %>
Opdateringen er gennemført!
