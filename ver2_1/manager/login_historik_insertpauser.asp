
<%
'*** Lukketid **'
strConnect = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_cst;"

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject("ADODB.Recordset")
Set oRec2 = Server.CreateObject("ADODB.Recordset")
Set oRec3 = Server.CreateObject("ADODB.Recordset")
Set oCmd = Server.CreateObject("ADODB.Command")

oConn.Open strConnect


               
				    
				    '*** Henter seneste login / logud ***
				    strSQL = "SELECT dato, minutter, id, mid FROM login_historik WHERE stempelurindstilling = 1 AND dato BETWEEN '2016-06-01' AND '2016-06-15' AND mid = 33 ORDER BY dato LIMIT 1000"
    				
				    fo_logud = 0
				    oRec.open strSQL, oConn, 3
                    while not oRec.EOF
    
                                    LoginDato = year(oRec("dato")) &"/"& month(oRec("dato")) &"/"& day(oRec("dato"))
                                    LoginDatoKl = year(oRec("dato")) &"/"& month(oRec("dato")) &"/"& day(oRec("dato")) & " 00:00:00"
                                    
                                    LoginDatoDelpauDagnr = datePart("w", oRec("dato"), 2,2)

                               
                                    psDato = year(psDato) &"/"& month(psDato) &"/"& day(psDato)

                                    strSQLp = "INSERT INTO login_historik SET dato = '"& LoginDato &"', "_
	                                &" login = '"& LoginDatoKl &"', "_
	                                &" logud = '"& LoginDatoKl &"', "_
					                &" stempelurindstilling = -1, minutter = 30, "_
					                &" manuelt_afsluttet = 0, kommentar = '', mid = 33" 
					        
					                Response.Write "pause: "& strSQLp & "<br>"
					                Response.flush
                                    'oConn.execute(strSQLp)

                                    
				          
                      oRec.movenext      
				      wend
				      oRec.close


Response.end
%>
<br>

