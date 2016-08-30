
<%
'*** Opdaterer timer, tfaktim til 0, 1, 2 ***
strConnect = "driver={MySQL};server=127.0.01; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_kringit;"

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject ("ADODB.Recordset")
Set oRec2 = Server.CreateObject ("ADODB.Recordset")
Set oCmd = Server.CreateObject ("ADODB.Command")
	
oConn.open strConnect

    strSQL = "SELECT id, navn, aktstatus FROM aktiviteter WHERE navn LIKE 'k.%'"
    oRec.open strSQL, oConn, 3
    while not oRec.EOF
    
    Response.Write oRec("id") &", "& oRec("navn") &", "& oRec("aktstatus") &"<br>"
    
    'oConn.execute("UPDATE aktiviteter SET aktstatus = 0 WHERE id = "& oRec("id") &"")
    
    oRec.movenext
    wend
    oRec.close




		
%>
<br>
Opdateringen er gennemført!
