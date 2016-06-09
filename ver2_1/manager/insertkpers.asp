

<%
strConnect = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_essens;"
	
Response.write strConnect

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject ("ADODB.Recordset")
Set oRec2 = Server.CreateObject ("ADODB.Recordset")
Set oCmd = Server.CreateObject ("ADODB.Command")

oConn.open strConnect


strSQL = "SELECT kid, telefon FROM kunder WHERE kid <> 774 AND kid <> 1002"
oRec.open strSQL, oConn, 3
while not oRec.EOF 
		
		Response.Write oRec("kid") & " : "
		
		if trim(oRec("telefon")) <> "" then
		strSQL2 = "UPDATE kontaktpers SET kundeid = "& oRec("kid")&" WHERE kundeid = "& oRec("telefon")
		'oConn.execute strSQL2
		Response.Write strSQL2 & "<br>"
		Response.flush
		end if
		
oRec.movenext
wend
oRec.close

%>
<br><br><br>
Opdatering gennemført!

<%
strSQL = "SELECT navn, titel, id FROM kontaktpers WHERE id > 797"
oRec.open strSQL, oConn, 3
while not oRec.EOF 
		
		strSQL2 = "UPDATE kontaktpers SET navn = '"& oRec("navn") &" "& oRec("titel") &"' WHERE id = "& oRec("id")
		'oConn.execute strSQL2
		Response.Write strSQL2 & "<br>"
		
		
oRec.movenext
wend
oRec.close






%>
<br><br><br>
Opdatering gennemført!
