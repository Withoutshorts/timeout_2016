

<%
strConnect = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_epi2017;"

	
Response.write strConnect & "<br><br>"

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject ("ADODB.Recordset")
Set oRec2 = Server.CreateObject ("ADODB.Recordset")
Set oCmd = Server.CreateObject ("ADODB.Command")

oConn.open strConnect



dim mid
x = 0

'grpn = "Letbane"
strSQL = "SELECT kundenr, kundenavn FROM kunder_epi2017_import"


Response.Write strSQL & "<hr>"
Response.flush

kundenr2017keepSQL = " WHERE kid > 1 "
oRec.open strSQL, oConn, 3
while not oRec.EOF 
								
								
           kundenr2017keepSQL = kundenr2017keepSQL & " AND kkundenr <> '"& oRec("kundenr") &"' AND kkundenavn <> '"& oRec("kundenavn") &"'"   


x = x + 1
oRec.movenext
wend
oRec.close


Response.write kundenr2017keepSQL
Response.flush
strSQL = "SELECT kundenr, kundenavn FROM kunder_epi2017_import"

'Response.Write strSQL & "<hr>"
'Response.flush

'oRec.open strSQL, oConn, 3
'while not oRec.EOF 
								
								
                    


'x = x + 1
'oRec.movenext
'wend
'oRec.close


%>
<br><br><br>
Antal records: <%=x %><br />
Opdatering gennemført!
