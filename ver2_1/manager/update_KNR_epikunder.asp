

<%
strConnect = "driver={MySQL ODBC 3.51 Driver};server=195.189.130.210; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_epi_no;"

	
Response.write strConnect & "<br><br>"

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject ("ADODB.Recordset")
Set oRec2 = Server.CreateObject ("ADODB.Recordset")
Set oCmd = Server.CreateObject ("ADODB.Command")

oConn.open strConnect



dim mid

strSQL = "SELECT ki_knr, ki_knr_nav, ki_knavn FROM ki_import_kunder_knr"


Response.Write strSQL & "<hr>"
Response.flush

oRec.open strSQL, oConn, 3
while not oRec.EOF 
								

							
                                if len(trim(oRec("ki_knr_nav"))) <> 0 AND oRec("ki_knr") <> "0" then    

                                
                                strSQLi = "UPDATE kunder SET kkundenr =  '"& oRec("ki_knr_nav") &"' WHERE kkundenr = '"& cstr(oRec("ki_knr")) & "'"
                           
                                Response.Write oRec("ki_knavn") & ":"&  strSQLi & "<br>"
                                
                                'oConn.execute(strSQLi)
                                Response.flush
                                
                                end if
                                
                               

oRec.movenext
wend
oRec.close



%>
<br><br><br>
Opdatering gennemført!
