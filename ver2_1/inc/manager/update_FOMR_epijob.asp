

<%
strConnect = "driver={MySQL ODBC 3.51 Driver};server=195.189.130.210; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_epi;"

	
Response.write strConnect & "<br><br>"

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject ("ADODB.Recordset")
Set oRec2 = Server.CreateObject ("ADODB.Recordset")
Set oCmd = Server.CreateObject ("ADODB.Command")

oConn.open strConnect



dim mid

strSQL = "SELECT fi.jobnr, business_alb, f.id AS fomrid, j.id AS jobid FROM forretningsomr_import AS fi "_
&" LEFT JOIN job AS j ON (j.jobnr = fi.jobnr)" _
&" LEFT JOIN fomr AS f ON (business_area_label = business_alb) WHERE fi.id >= 10 AND j.id IS NOT NULL ORDER BY fi.id"

Response.Write strSQL & "<hr>"
Response.flush

oRec.open strSQL, oConn, 3
while not oRec.EOF 
								

							

                                
                                strSQLi = "INSERT INTO FOMR_REL (for_fomr, for_jobid, for_aktid, for_faktor) VALUES ("& oRec("fomrid") &", "& oRec("jobid") &", 0, 100) "
                           
                                'Response.Write  strSQLd & "<br>"
                                
                                'oConn.execute(strSQLi)
                                Response.flush
                                
                               

oRec.movenext
wend
oRec.close



%>
<br><br><br>
Opdatering gennemført!
