

<%
'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
Session.LCID = 1030

'strConnect = "driver={MySQL ODBC 3.51 Driver};server=81.19.249.35; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_epi_catitest;"
'** DK 195.189.130.210
'strConnect = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_epi;"
'** NO
'strConnect = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_epi_no;"
'** UK
'strConnect = "driver={MySQL ODBC 3.51 Driver};server=195.189.130.210; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_epi_uk;"
	
'**EPI2017
strConnect = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_epi2017;"

Response.write strConnect & "<br><br>"

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject ("ADODB.Recordset")
Set oRec2 = Server.CreateObject ("ADODB.Recordset")
Set oRec3 = Server.CreateObject ("ADODB.Recordset")
Set oCmd = Server.CreateObject ("ADODB.Command")

oConn.open strConnect

'* FØR LOGIN
'strSQL= "SELECT * FROM login_historik WHERE DATE_FORMAT(login,'%H:%i:%s') > '05:00:00' AND DATE_FORMAT(login,'%H:%i:%s') < '07:50:00' AND dato BETWEEN '2017-05-01' AND '2017-05-31'"

'* EFTER LOGUD
strSQL= "SELECT * FROM login_historik WHERE DATE_FORMAT(logud,'%H:%i:%s') > '20:00:00' AND DATE_FORMAT(login,'%H:%i:%s') < '24:00:00' AND dato BETWEEN '2017-05-01' AND '2017-05-31'"

Response.Write strSQL & "<hr>"
Response.flush

'Response.end
x = 0
oRec.open strSQL, oConn, 3
while not oRec.EOF 
				 

                
       


                tmnr = oRec("mid")
                meNavn = "-"

                strSQLmnavn = "SELECT mnavn, init, mnr, email, medarbejdertype, ansatdato, opsagtdato, forecaststamp, brugergruppe, visskiftversion, mansat, timer_ststop, create_newemployee, mcpr "_
                &" FROM medarbejdere WHERE mid = "& tmnr

	            oRec3.open strSQLmnavn, oConn, 3
	            if not oRec3.EOF then
    	
	            meNavn = oRec3("mnavn")
	   
    	
	            end if
	            oRec3.close

                tastedato = year(now) & "-" & month(now) & "-" & day(now)            
                tdato = year(oRec("dato")) & "-" & month(oRec("dato")) & "-" & day(oRec("dato"))

                'FØR login
                'tdatoSTtid = oRec("login")
                'tdatoSLtid = oRec("dato") &" 08:00:00"
                'EFTER logud

                tdatoSTtid = oRec("dato") &" 20:00:00"
                tdatoSLtid = oRec("logud")

                timerthis = dateDiff("n", tdatoSTtid, tdatoSLtid, 2,2)

			    strSQLins = "INSERT INTO timer "_
				&" (tastedato, editor, tmnr, tmnavn, tfaktim, tdato, timer, taktivitetnavn, taktivitetid, tjobnr) VALUES "_
				&" ('"& tastedato &"', 'TU Support indlæst', "& tmnr &", '"& meNavn &"', '53', '"& tdato &"', "& replace(formatnumber(timerthis/60, 2), ",",".")  &", 'AVI tillæg 15', 247416, 3)"
				
				Response.Write "OK1: Login: "& tdatoSTtid &": "& strSQLins & "<br>"
				Response.flush
				
                'oConn.execute(strSQLins)

                             
x = x + 1
oRec.movenext
wend
oRec.close



%>
<br><br><br>
Opdatering gennemført!<br />
<%=x %> records.
