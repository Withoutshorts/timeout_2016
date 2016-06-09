
<%
function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
	end function



strConnect = "driver={MySQL};server=127.0.0.1; Port=3306; uid=root;pwd=;database=timeout_sthaus;"
Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject ("ADODB.Recordset")
Set oRec2 = Server.CreateObject ("ADODB.Recordset")
Set oCmd = Server.CreateObject ("ADODB.Command")
	
oConn.open strConnect

x = 0
strSQL = "SELECT id, jobnavn, jobnr FROM job WHERE fakturerbart <> 1 ORDER BY id"
oRec.open strSQL, oConn, 3
While not oRec.EOF 
		
		Response.write oRec("id") & " - " & oRec("jobnavn") & "&nbsp; ("& oRec("jobnr") &")<br>"

x = x + 1		
oRec.movenext
wend
oRec.close


		
%>
Der er ialt <b><%=x%></b> job.
<br>

