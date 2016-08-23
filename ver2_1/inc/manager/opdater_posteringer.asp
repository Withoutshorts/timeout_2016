<%
strConnThis = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_intranet;"
lto = "outz"

strConnect = strConnThis
	
	'Set oConn = Server.CreateObject("ADODB.Connection")
	'Set oRec = Server.CreateObject ("ADODB.Recordset")
	'Set oCmd = Server.CreateObject ("ADODB.Command")
	
oConn.open strConnect

strSQL = "SELECT kontonr, modkontonr, id, bilagsnr, posteringsdato FROM posteringer WHERE bilagsnr = 224"
'Response.write strSQL
'Response.flush
oRec.open strSQL, oConn, 3 
while not oRec.EOF 

Response.write oRec("posteringsdato") &" # "& oRec("id") & ": " & oRec("kontonr") &" - " & oRec("modkontonr") & " bilgsnr: ("& oRec("bilagsnr") &")<br>"
		
		strSQL2 = "UPDATE posteringer SET kontonr = "& oRec("modkontonr") &", modkontonr = "& oRec("kontonr") &" WHERE id ="& oRec("id")  
		Response.write strSQL2 & "<br>"
		'oConn.execute (strSQL2)
oRec.movenext
wend
oRec.close 
%>
