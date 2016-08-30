
<%
'*** Lukketid **'
strConnect = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_jttek;"

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject("ADODB.Recordset")
Set oRec2 = Server.CreateObject("ADODB.Recordset")
Set oRec3 = Server.CreateObject("ADODB.Recordset")
Set oCmd = Server.CreateObject("ADODB.Command")

oConn.Open strConnect


               
				    
				    '*** Henter seneste login / logud ***
				    strSQL = "SELECT dato, minutter, id, mid FROM login_historik WHERE stempelurindstilling = 1 AND manuelt_afsluttet = 2 AND dato > '2015-01-01' ORDER BY dato LIMIT 1000"
    				
				    fo_logud = 0
				    oRec.open strSQL, oConn, 3
                    while not oRec.EOF
    
                                    LoginDatoDelpau = year(oRec("dato")) &"/"& month(oRec("dato")) &"/"& day(oRec("dato"))
                                    
                                    LoginDatoDelpauDagnr = datePart("w", oRec("dato"), 2,2)

                                    Response.write "<br>LoginDatoDelpauDagnr: " & LoginDatoDelpauDagnr 

                                    if LoginDatoDelpauDagnr = 5 then
                                    lmt = 1
                                    else
                                    lmt = 2
                                    end if

                                    strSQLpauser = "SELECT dato, minutter, id, mid FROM login_historik WHERE stempelurindstilling = -1 AND dato = '"& LoginDatoDelpau &"' AND mid = " & oRec("mid") & " ORDER BY id LIMIT "& lmt
    	                            response.write "<br><br>"& strSQLpauser 
                                    response.flush			
        
				                    fo_logud = 0
				                    oRec2.open strSQLpauser, oConn, 3
                                    while not oRec2.EOF 
                        
                       
                                                Response.Write "<br>ID: " & oRec2("id") & " - Dato: "& oRec2("dato") & " minutter: " & oRec2("minutter") 
	                            
	                            
                                               'strSQLpauser2 = "SELECT dato, minutter, id FROM login_historik WHERE dato = '"& LoginDatoDelpau &"' AND minutter = "& oRec2("minutter") &" AND mid = " & oRec("mid") &" AND id <> "& oRec2("id")
    				
                                                'response.write "<br>"&strSQLpauser2 
                                                'response.flush 

				                                'fo_logud = 0
				                                'oRec3.open strSQLpauser2, oConn, 3
                                                'if not oRec3.EOF then

	                                            strSQLpDel = "DELETE FROM login_historik WHERE dato = '"& LoginDatoDelpau &"' AND minutter = "& oRec2("minutter") &" AND mid = " & oRec2("mid") &" AND id <> "& oRec2("id")
                                                'oConn.execute(strSQLpDel)
				                
				                                'Response.Write "<br><b>"& strSQLpDel & "</b>"
				                                'Response.end

                                                      
				                                'end if
				                                'oRec3.close

                                      oRec2.movenext      
				                      wend
				                      oRec2.close
				                
				          
                      oRec.movenext      
				      wend
				      oRec.close


Response.end
%>
<br>

