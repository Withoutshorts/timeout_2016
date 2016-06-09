

<%
strConn = "driver={MySQL};server=192.168.1.33; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_intranet;"'objDBConn(0).text  'Connection


strConnect = strConn
	
	Set oConn = Server.CreateObject("ADODB.Connection")
	Set oRec = Server.CreateObject ("ADODB.Recordset")
	Set oRec2 = Server.CreateObject ("ADODB.Recordset")
	Set oRec3 = Server.CreateObject ("ADODB.Recordset")
	Set oCmd = Server.CreateObject ("ADODB.Command")
	
	oConn.open strConnect



strSQL = "SELECT licens FROM licens WHERE id = 1"
	oRec.Open strSQL, oConn, 0, 1, 1
	if not oRec.EOF then
	licensto = oRec("licens")
	end if
	oRec.close
	
'strSQL = "INSERT INTO licens (licens) VALUES ('1111')"
'oConn.execute(strSQL)
	
Response.write  licensto & " databasen kører."
%>
