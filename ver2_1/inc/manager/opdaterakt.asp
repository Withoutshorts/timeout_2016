
<%
function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
	end function



strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_bowe;"
Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject ("ADODB.Recordset")
Set oRec2 = Server.CreateObject ("ADODB.Recordset")
Set oCmd = Server.CreateObject ("ADODB.Command")
	
oConn.open strConnect

x = 0
a = 0
strSQL = "SELECT id, jobnavn, jobnr FROM job WHERE jobstatus <> 3 ORDER BY id"
oRec.open strSQL, oConn, 3
While not oRec.EOF 
		
		strSQL2 = "SELECT id, navn FROM aktiviteter WHERE fakturerbar = 1 AND job = " & oRec("id")
		oRec2.open strSQL2, oConn, 3
		while not oRec2.EOF
			
			nytNavn = oRec2("navn") &" Dag"
			strSQL3 = "UPDATE aktiviteter SET navn = '"& nytNavn &"', tidslaas = 1, tidslaas_st = '08:00:00', tidslaas_sl = '18:00:00' WHERE id = " & oRec2("id")
			'Response.write strSQL3 & "<br>"
			'oConn.execute(strSQL3)
		a = a + 1
		oRec2.movenext
		wend
		
		oRec2.close

x = x + 1		
oRec.movenext
wend
oRec.close


		
%>
Der er ialt <b><%=x%></b> job og <b><%=a%></b> aktiviteter.
<br>

