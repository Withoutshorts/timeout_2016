<%
lto = "abu"
strConnThis = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_abu;"
	

strConnect = strConnThis
	
	Set oConn = Server.CreateObject("ADODB.Connection")
	Set oRec = Server.CreateObject ("ADODB.Recordset")
	Set oCmd = Server.CreateObject ("ADODB.Command")
	
oConn.open strConnect

strSQL = "SELECT domaner, kundeid FROM kundedom WHERE kundeid <> 0"
'Response.write strSQL
'Response.flush
oRec.open strSQL, oConn, 3 
while not oRec.EOF 
        
        Response.Write oRec("kundeid") &"<br>"& oRec("domaner") & "<hr>"
		
		strSQL2 = "UPDATE kunder SET url = '"& oRec("domaner") &"' WHERE kkundenr ="& oRec("kundeid")  
		'Response.write strSQL2 & "<br>"
		'oConn.execute (strSQL2)
oRec.movenext
wend
oRec.close 
%>
