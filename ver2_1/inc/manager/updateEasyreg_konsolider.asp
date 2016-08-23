

<%
'strConnect = "driver={MySQL ODBC 3.51 Driver};server=81.19.249.35; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_dencker;"
'strConnect = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_dencker_BAK;"

'strConnect = "driver={MySQL ODBC 3.51 Driver};server=81.19.249.35; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_epi;"
'strConnect = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_epi;"



lto = "xepi"
	
Response.write strConnect & "<br><br>"

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject ("ADODB.Recordset")
Set oRec2 = Server.CreateObject ("ADODB.Recordset")
Set oCmd = Server.CreateObject ("ADODB.Command")

oConn.open strConnect





public timerThis, pleaseUpdate, lastUpdatedRecordID

lastMid = 0
lastJobid = 0
lastAkt = 0
lastMnth = 0

timerThis = 0
pleaseUpdate = 0
lastUpdatedRecordID = 0


select case lto 
case "dencker"
dateNow = dateAdd("m",-1, now)
dateNowSQL = year(dateNow) &"/"& month(dateNow) &"/1" 
strSQL = "SELECT tid, tmnr, tjobnr, taktivitetid, tdato, timer FROM timer WHERE timer < 0.15 AND tdato <= '"& dateNowSQL &"' ORDER BY tmnr, tjobnr, taktivitetid, tdato" 
case "epi"
dateNow = dateAdd("m", 0, now)
dateNowSQL = year(dateNow) &"/"& month(dateNow) &"/1" 
strSQL = "SELECT tid, tmnr, tjobnr, taktivitetid, tdato, timer FROM timer WHERE timer < 4.5 AND tdato <= '"& dateNowSQL &"' AND tfaktim = 91 ORDER BY tmnr, tjobnr, taktivitetid, tdato"
case "xepi"
 
strSQL = "SELECT tk_id, tk_timerid FROM timer_konsolideret WHERE tk_timerid > 183030 ORDER BY tk_timerid LIMIT 1000"


end select

Response.Write strSQL & "<br><br>"
Response.flush

x = 0
oRec.open strSQL, oConn, 3
while not oRec.EOF 
					
                    			
							

                   
                        
                        if cdbl(x) <> 0 then

                        '*** skal timer konsolideres på sin egen akt eller på en fælles aktivitet ***'
                        'if lastAkt <> oRec("taktivitetid") then 
                        'pleaseUpdate = 1
                        'end if

                        'if lastJobid <> oRec("tjobnr") then
                        'pleaseUpdate = 1
                        'end if

                        'if lastMid <> oRec("tmnr") then
                        'pleaseUpdate = 1
                        'end if

                        '**** Skal de samles på månedsbasis ***'
                        'if lastMnth <> month(oRec("tdato")) then
                        'pleaseUpdate = 1
                        'end if

                        '**** Skal de samles på dagsbasis ***'
                        'if lastDay <> day(oRec("tdato")) then
                        'pleaseUpdate = 1
                        'end if

                        end if

                    
                     if cint(pleaseUpdate) = 1 then '**indlæsrecord

                        'call updateTimer

                    else

                         if cdbl(lastUpdatedRecordID) <> cdbl(lastRecordID) then
                         delrecords = delrecords & " OR tid = "& lastRecordID
                         end if

                    end if
                    
                    'timerThis = timerThis + formatnumber(oRec("timer"))           

                   
                    

                    'Response.Write oRec("tid") &", "& oRec("tmnr") &", "& oRec("tjobnr") &", "& oRec("taktivitetid") &", "& oRec("tdato") &", "& oRec("timer") & "<br>"	

                    '**** Alle records indlæses til konsolider, hvis nu der skal rulles tilbage ****'
                    'timerK = replace(oRec("timer"), ",", ".")
                    'datok = year(oRec("tdato")) & "/" & month(oRec("tdato")) & "/" & day(oRec("tdato"))

                    'strSQLkon = "INSERT INTO timer_konsolideret (tk_dato, tk_timerid, tk_mnr, tk_jobnr, tk_aid, tk_timer, dato_kons) "_
                    '&" VALUES ('"& datok &"', "& oRec("tid") &", "& oRec("tmnr") &", "& oRec("tjobnr") &", "& oRec("taktivitetid") &", "& timerK &", '"& dateNowSQL &"')"
                    'oConn.execute(strSQLkon)

                    'Response.write strSQLkon
                    'Response.end
                    'Response.Write timerK
                    'Response.end

                    'strSQLd = "DELETE FROM timer_konsolideret WHERE tk_timerid = " & oRec("tk_timerid") &" AND tk_id <> " & oRec("tk_id")
                    'oConn.execute(strSQLd)


'lastMid = oRec("tmnr")
'lastJobid = oRec("tjobnr")
'lastAkt = oRec("taktivitetid")
'lastMnth = month(oRec("tdato"))
'lastRecordID = oRec("tid")
'lastDay = day(oRec("tdato"))
 

x = x + 1
oRec.movenext
wend
oRec.close

 call updateTimer





'**** Func *****'
sub updateTimer
                    

                    '**** Opdaterer ***'
                    timerThis = replace(timerThis, ".", "")
                     timerThis = replace(timerThis, ",", ".")
                     strSQLupd = "UPDATE timer SET timer = " & timerThis & ", origin = 91 WHERE tid = " & lastRecordID
                     'Response.write strSQLupd & "<br>"

                     lastUpdatedRecordID = lastRecordID
                    
                     'oConn.execute(strSQLupd)
                     timerThis = 0
                     pleaseUpdate = 0


                     '***** Renser ****'
                    strSQLdel = "DELETE FROM timer WHERE tid = 0 "& delrecords

                    'Response.write "<br>"& strSQLdel & "<hr>"
                    'oConn.execute(strSQLdel)

                    delrecords = ""

                    Response.flush

end sub



%>
antal records: <%=x %>
<br><br><br>
Opdatering gennemført!
