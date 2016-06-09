
<%
'*** Lukketid **'
strConnect = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_dencker;"

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject ("ADODB.Recordset")
Set oRec2 = Server.CreateObject ("ADODB.Recordset")
Set oCmd = Server.CreateObject ("ADODB.Command")

oConn.Open strConnect


strUsrId = 12	
LoginDateTime = year(now)&"/"& month(now)&"/"&day(now)&" "& datepart("h", now) &":"& datepart("n", now) &":"& datepart("s", now) 
				LoginDato = year(now)&"/"& month(now)&"/"&day(now)
				
				    
				    '*** Henter seneste login / logud ***
				    strSQL = "SELECT id, dato, login, logud, mid, stempelurindstilling FROM login_historik WHERE "_
				    &" mid = "& strUsrId &" AND stempelurindstilling > -1 ORDER BY id DESC limit 0, 1"
    				
				    fo_logud = 0
				    oRec.open strSQL, oConn, 3
                    if not oRec.EOF then
    
                        
                       
                      
                                
                                logudDag = weekday(oRec("dato"), 2)
                                
                                Response.Write logudDag &" dato:"& oRec("dato") &": wdn:"& weekdayname(weekday(oRec("dato"),1)) &"<br>"
                                
                                select case logudDag
                                case 1
                                dagkri = "normtid_sl_man"
                                case 2
                                dagkri = "normtid_sl_tir"
                                case 3
                                dagkri = "normtid_sl_ons"
                                case 4
                                dagkri = "normtid_sl_tor"
                                case 5
                                dagkri = "normtid_sl_fre"
                                case 6
                                dagkri = "normtid_sl_lor"
                                case 7
                                dagkri = "normtid_sl_son"
                                end select
                                
                                '*** Henter firmalukketid ****
                                lukketid = "23:59:00"
                                strSQL = "SELECT "& dagkri &" AS lukketid FROM licens WHERE id = 1 "
                                
                                'Response.Write strSQL
                                'Response.flush
                                oRec2.open strSQL, oConn, 3
                                if not oRec2.EOF then
                                    if len(trim(oRec2("lukketid"))) <> 0 then
                                    lukketid = oRec2("lukketid")
                                    else
                                    lukketid = "23:59:00"
                                    end if
                                End if
                                oRec2.close
                                
                                Response.Write "<br>Lukketid: "& formatdatetime(lukketid, 3)
                                
                                
                                LogudDateTime = year(oRec("dato"))&"/"& month(oRec("dato"))&"/"&day(oRec("dato"))&" "& formatdatetime(lukketid, 3)
                                
                                '**** Minutter beregning ***
                                loginTidAfr = formatdatetime(oRec("login"), 3)
                                logudTidAfr = formatdatetime(LogudDateTime, 3)
                               
                               
                                minThisDIFF = datediff("s", loginTidAfr, logudTidAfr)/60
	                            minThisDIFF = replace(formatnumber(minThisDIFF, 0), ",", ".")
	                            
	                            
	                            
	                            strSQLupd = "UPDATE login_historik SET logud = '"& LogudDateTime &"', minutter = "& minThisDIFF &" WHERE id = "& oRec("id")
				                
				                Response.Write "<br>"& strSQLupd
				                'Response.end
				                
				                
				      end if
				      oRec.close


Response.end
%>
<br>

