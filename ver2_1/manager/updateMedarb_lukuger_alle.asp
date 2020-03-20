

<%
strConnect = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_plan;"

	
Response.write strConnect & "<br><br>"

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject ("ADODB.Recordset")
Set oRec2 = Server.CreateObject ("ADODB.Recordset")
Set oCmd = Server.CreateObject ("ADODB.Command")

oConn.open strConnect



dim mid
x = 0

'grpn = "Letbane"
strSQL = "SELECT mid FROM medarbejdere WHERE (mid > 0) AND mansat = 1" 

Response.Write strSQL & "<hr>"
Response.flush

oRec.open strSQL, oConn, 3
while not oRec.EOF 
								
								
                           
                                    cDateUgeTilAfslutning = "30-09-2018"
                                    cDateAfs = year(now) & "-" & month(now) & "-"& day(now)
                                
                                    d = 0
                                    for d = 0 TO 365
    
                                    if d = 0 then
                                    cDateUgeTilAfslutning = cDateUgeTilAfslutning
                                    else
                                    cDateUgeTilAfslutning = dateAdd("d", -7, cDateUgeTilAfslutning)
                                    end if

                                    cDateUgeTilAfslutning = year(cDateUgeTilAfslutning) & "-" & month(cDateUgeTilAfslutning) & "-" & day(cDateUgeTilAfslutning)

                    
                                    if year(cDateUgeTilAfslutning) = "2018" then

                                    medarbejderid = oRec("mid") 
                                    intStatusAfs = 0
                                    intStatusAfs = 1
                                    splithr = 0
        
        
                                    findesden = 0
                                    findesdenSQL = "SELECT id FROM ugestatus WHERE uge = '"& cDateUgeTilAfslutning &"' AND mid = " & medarbejderid & " AND splithr = 0"
                                    
                                    Response.Write findesdenSQL & "<br>"
                                    Response.flush
                                    oRec2.open findesdenSQL, oConn, 3
                                    if not oRec2.EOF then
                                    findesden = 1
                                    end if
                                    oRec2.close


                                    if cint(findesden) = 0 then
								    strSQLafslut = "INSERT INTO ugestatus (uge, afsluttet, mid, status, splithr) VALUES ('"& cDateUgeTilAfslutning &"', '"& cDateAfs &"', "& medarbejderid &", "& intStatusAfs &", "& splithr &")" 
		                            'Response.Write strSQLafslut
		                            'Response.flush
		                            'oConn.execute(strSQLafslut)
                                    end if


                                    end if

                                    next
                                
                            
                                'Response.Write strSQLafslut & "<br>"
                                Response.flush
                            


x = x + 1
oRec.movenext
wend
oRec.close



%>
<br><br><br>
Antal records: <%=x %><br />
Opdatering gennemført!
