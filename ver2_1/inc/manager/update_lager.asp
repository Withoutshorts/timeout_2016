
<%
'*** Opdaterer lager ***
strConnect = "driver={MySQL};server=127.0.0.1; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_bowe;"

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject ("ADODB.Recordset")
Set oRec2 = Server.CreateObject ("ADODB.Recordset")
Set oCmd = Server.CreateObject ("ADODB.Command")
	
oConn.open strConnect

    strSQL = "SELECT id, varenr, saldo, indkobspris, navn FROM mat_import WHERE id > 601 ORDER BY id"
    oRec.open strSQL, oConn, 3
    x = 0
    while not oRec.EOF
    
    Response.Write oRec("id") &", "& oRec("varenr") &", "& oRec("saldo") &"<br>"
    
        strSQL2 = "SELECT id, varenr FROM materialer WHERE varenr = '"& oRec("varenr") &"' AND matgrp = 2"
        oRec2.Open strSQL2, oConn, 3
        
        if not oRec2.EOF then
        'oConn.execute("UPDATE materialer SET antal = "&oRec("saldo")&" WHERE varenr = '"& oRec("varenr") &"' AND matgrp = 2")
        else
        Response.Write "Findes ikke i forvejen. Indsætter...<br>"
        strSQLins = "INSERT INTO materialer (antal, navn, indkobspris, varenr, matgrp) VALUES ("& oRec("saldo") &", '"& replace(oRec("navn"), "'", "''") &"', "& replace(oRec("indkobspris"),",",".") &", '"& oRec("varenr") &"', 2)"
        
        Response.Write strSQLins & "<br>"
        Response.flush
        
        'oConn.execute(strSQLins)
        
        x = x + 1
        end if
        oRec2.close
        
    oRec.movenext
    wend
    oRec.close




		
%>
<br>
Opdateringen er gennemført!<br />
Der er indasat <%=x %> records.
