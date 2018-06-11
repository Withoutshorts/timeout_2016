

<%
'strConnect = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_oko;"
 strConnect = "timeout_oko64"
 strConnect = "timeout_cflow64"

Response.write strConnect & "<br><br>"

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject ("ADODB.Recordset")
Set oRec2 = Server.CreateObject ("ADODB.Recordset")
Set oCmd = Server.CreateObject ("ADODB.Command")

oConn.open strConnect



dim mid
x = 0


strSQL = "SELECT mid, medarbejdertype, mnavn FROM medarbejdere WHERE mid > 1"

Response.Write strSQL & "<hr>"
Response.flush

oRec.open strSQL, oConn, 3
while not oRec.EOF 
								
								
                mid = oRec("mid")
                mtypenavnnavn = "Mtyp: "& oRec("mnavn")
                strDato = year(now) & "-"& month(now) & "-" & day(now)                          
                strTimepris = 400
                dubKostpris = 0
                normtimer_son = "7.4"
                normtimer_man = "7.4"
                normtimer_tir = "7.4"
                normtimer_ons = "7.4"
                normtimer_tor = "7.4"
                normtimer_fre = "7.4"
                normtimer_lor = "0"
                normtimer_son = "0"
                strEditor = "UpdateService"

                strTimepris1 = 300
                strTimepris2 = 0
                strTimepris3 = 0
                strTimepris4 = 0
                strTimepris5 = 0

                tp0_valuta = 1
                tp1_valuta = 1
                tp2_valuta = 1
                tp3_valuta = 1
                tp4_valuta = 1
                tp5_valuta = 1
                              
                sostergp = 0
                mtsortorder = 0
                mgruppe = 0
                afslutugekri = 0
                afslutugekri_proc = 0
                noflex = 0
            
                kostpristarif_A = 0
                kostpristarif_B = 0
                kostpristarif_C = 0
                kostpristarif_D = 0

                strSQlinsMtyp = "INSERT INTO medarbejdertyper (type, timepris, editor, dato, kostpris, normtimer_son, normtimer_man, normtimer_tir, normtimer_ons, normtimer_tor, normtimer_fre, normtimer_lor, "_
				&" timepris_a1, timepris_a2, timepris_a3, timepris_a4, timepris_a5, "_
				&" tp0_valuta, tp1_valuta, tp2_valuta, tp3_valuta, tp4_valuta, tp5_valuta, sostergp, mtsortorder, mgruppe, afslutugekri, afslutugekri_proc, noflex, "_
                &" kostpristarif_A, kostpristarif_B, kostpristarif_C, kostpristarif_D) VALUES"_
				&" ('"& mtypenavnnavn &"', "& strTimepris &", '"& strEditor &"', '"& strDato &"', "& dubKostpris &", "_
				&" "& normtimer_son &", "& normtimer_man &", "& normtimer_tir &", "& normtimer_ons &", "_
				&" "& normtimer_tor &", "& normtimer_fre &", "& normtimer_lor &", "& strTimepris1 &", "_
				&" "& strTimepris2 &", "& strTimepris3 &", "& strTimepris4 &", "& strTimepris5 &", "_
				&" "& tp0_valuta &","& tp1_valuta &","& tp2_valuta &","& tp3_valuta &","& tp4_valuta &","& tp5_valuta &", "& sostergp &", "_
                &" "& mtsortorder &", "& mgruppe &","& afslutugekri &","& afslutugekri_proc &", "& noflex &", "_
                &" "& kostpristarif_A &","& kostpristarif_B &","& kostpristarif_C &","& kostpristarif_D &""_
				&" )"

        
                oConn.execute(strSQlinsMtyp)

                lastType = 0
                strSQLlastmtyp = "SELECT id FROM medarbejdertyper WHERE id <> 0 ORDER BY id DESC"
                oRec2.open strSQLlastmtyp, oConn, 3
                if not oRec2.EOF then

                lastType = oRec2("id") 

                end if
                oRec2.close
            

                strSQLlastmtypupd = "UPDATE medarbejdere SET medarbejdertype = "& lastType &" WHERE mid = " & oRec("mid")
                oConn.execute(strSQLlastmtypupd)

x = x + 1
oRec.movenext
wend
oRec.close



%>
<br><br><br>
Antal records: <%=x %><br />
Opdatering gennemført!
