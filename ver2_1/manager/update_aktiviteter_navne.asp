
<%
function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
	end function


'*** Opdaterer timer, tfaktim til 0, 1, 2 ***
'strConnect = "driver={MySQL};server=192.168.1.33; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_wps;"
strConnect = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_symbion;"
  
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=root;pwd=;database=timeout_intranet;"
'strConnect = "mySQL_timeOut_intranet"
Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject ("ADODB.Recordset")
Set oRec2 = Server.CreateObject ("ADODB.Recordset")
Set oCmd = Server.CreateObject ("ADODB.Command")
	
oConn.open strConnect

strSQL = "SELECT id, navn, beskrivelse FROM aktiviteter ORDER BY id"
oRec.open strSQL, oConn, 3
While not oRec.EOF 
		
		'newName1 = Replace(oRec("taktivitetnavn"), "Dag", "(Dag)")
		'newName2 = Replace(newName1, "Aften", "(Aften)")
		'newName3 = Replace(newName2, "Nat", "(Nat)")
		
		strSQL = "UPDATE aktiviteter SET navn = '"& oRec("navn") &" "& oRec("beskrivelse") &"', beskrivelse = '' WHERE id = "& oRec("id") 
		
		Response.write strSQL & "<br>"
		
		'oConn.execute(strSQL)
		
oRec.movenext
wend
oRec.close


		
%>
<br>
Opdateringen er gennemført!
