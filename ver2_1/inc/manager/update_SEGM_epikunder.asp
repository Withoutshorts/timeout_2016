

<%
strConnect = "driver={MySQL ODBC 3.51 Driver};server=195.189.130.210; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_epi;"

	
Response.write strConnect & "<br><br>"

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject ("ADODB.Recordset")
Set oRec2 = Server.CreateObject ("ADODB.Recordset")
Set oCmd = Server.CreateObject ("ADODB.Command")

oConn.open strConnect



dim mid

strSQL = "SELECT ki_knr, ki_fomr, kt.id AS segmentid, ki_cvr FROM ki_import_kunder_fomr AS ki "_
&" LEFT JOIN kundetyper AS kt ON (kt.navn = ki.ki_fomr) WHERE ki.ki_kid >= 0"

Response.Write strSQL & "<hr>"
Response.flush

oRec.open strSQL, oConn, 3
while not oRec.EOF 
								

							
                                if len(trim(oRec("segmentid"))) <> 0 AND oRec("ki_knr") <> "0" then    

                                
                                strSQLi = "UPDATE kunder SET ktype =  "& oRec("segmentid") &" WHERE (kkundenr = '"& cstr(oRec("ki_knr")) & "') OR (cvr = '"& oRec("ki_cvr") &"' AND cvr <> '0')"
                           
                                Response.Write  strSQLi & "<br>"
                                
                                'oConn.execute(strSQLi)
                                Response.flush
                                
                                end if
                                
                               

oRec.movenext
wend
oRec.close



%>
<br><br><br>
Opdatering gennemført!
