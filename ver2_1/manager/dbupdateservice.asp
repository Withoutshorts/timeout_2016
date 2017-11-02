
<% if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then %>
<!--include file="../inc/connection/aktivedb_inc.asp"-->
<%else %>
<!--include file="../inc/connection/aktivedb_r_inc.asp"-->
<%end if %>

<!--include file="../inc/connection/aktivedb_inc.asp"-->
<!--#include file="../inc/connection/aktivedb_r_inc.asp"-->

<%

if len(request("opdateralledb")) <> 0 then

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject("ADODB.Recordset")



if len(request("brug_sql")) <> 0 then
brugSQL = request("brug_sql")
else
brugSQL = 0
end if


dim strSQL
select case brugSQL
case "0"
strSQL = split(request("SQL"), ";")
case "1"
strSQL = split(request("SQL1"), ";")
case "2"
strSQL = split(request("SQL2"), ";")
case "3"
strSQL = split(request("SQL3"), ";")
case "4"
strSQL = split(request("SQL4"), ";")
case "5"
strSQL = split(request("SQL5"), ";")
case "6"
strSQL = split(request("SQL6"), ";")
case "7"
strSQL = split(request("SQL7"), ";")
case "8"
strSQL = split(request("SQL8"), ";")
end select


opdaterDB = 1

if opdaterDB = 1 then
a = 0
		
		
				for b = 0 to Ubound(strSQL) 
					Response.write "<hr>Ny SQL:<br>"
					Response.write strSQL(b) & "<br>"
					Response.flush
					x = 1
					numberoflicens = 171
					For x = 1 To numberoflicens  
						
						call aktivedb(x)
						
                        Response.write strConnect_aktiveDB &"<br>"
						
						if strConnect_aktiveDB <> "nogo" then
							
							
							lto = 0
							passive = 0
								
								
								'*** Antal medarbejdere pr. LTO (til TSA) *****
								select case brugSQL
								case "0"
								
								
								
								
								
								'*** HER ***
								Response.write x &"<br>"& strSQL(b) & "<br><br>"
								Response.flush

								if x > 0 then 'AND x <> 130 then
                                oConn.open strConnect_aktiveDB
							    
                                '*** DENNE LINJE INDLÆSER // UDKOMMENTER NÅR FILEN IKKE ER AKTIV
                                oConn.execute(strSQL(b)) 
                                
                                '* TJECK x NUMBER OF LICENS fra inc/connection/db_conn...asp filen.
								'* DISSE SKAL STEMME 
								
								oConn.close

							    end if
								
								
								a = a + 1
								
								
								case "1"
								'Response.write strSQL
								oConn.open strConnect_aktiveDB
								strSQLx = "SELECT l.key, licens, licenstype, erp, crm, sdsk, listatus, autogktimer FROM licens l WHERE id = 1" 
								
								'Response.Write strSQLx
								'Response.flush
								
								oRec.open strSQLx, oConn, 3
								while not oRec.EOF
								
								Response.Write "<b>Lincens indehaver: " & oRec("licens") & "</b><br>"
								Response.write ""& oRec("licenstype") & "<br>"
								Response.Write "erp: " & oRec("erp") & " crm: " & oRec("crm") & ", sdsk: " & oRec("sdsk") & "<br>"
                                Response.write "autogktimer: "&  oRec("autogktimer")
								Response.Write "<br>Listatus: " & oRec("listatus") & "<br>"
								
								listatus = oRec("listatus")
								
								key = oRec("key")
								
								    
								
	                                          
								                
								oRec.movenext
								wend
								oRec.close
								
								   
								    strSQL(b) = "SELECT * FROM medarbejdere WHERE mid = 1"
								    oRec.open strSQL(b), oConn, 3
								    while not oRec.EOF
								    Response.write oRec("mnavn") & "<br>"
								    a = a + 1
								    lto = lto + 1
    								
    								
								    if cint(listatus) = 1 then
								    
								    aktivelbrugereialt = aktivelbrugereialt + 1
								     
								     if oRec("mansat") = "3" then
								     passive = passive + 1
								     end if
								    
								    else
								    aktivelbrugereialt = aktivelbrugereialt 
								    end if
    								
								    oRec.movenext
								    wend
    								
								    Response.write "Antal medarbejdere denne licens: <b>" & lto &"</b>, heraf passive: <b>"& passive &"</b><br><br>"
								    oRec.close
								
								oConn.close
								
								
								case "2"
								'Response.write "<hr>Henter fakturaer: "& strConnect_aktiveDB &"<br>"
								oConn.open strConnect_aktiveDB
								
								oRec.open "SELECT fid, faknr, beloeb, kurs, valuta FROM Fakturaer WHERE fid <> 0 AND valuta <> 1", oConn, 3
								while not oRec.EOF
								
								'Response.write oRec("fid") &" - Faknr: <b>"& oRec("faknr") & "</b> | " & oRec("beloeb") & " | kurs:  " & oRec("kurs") & " | valuta: " & oRec("valuta") & "<br>"
								
								oRec.movenext
								wend
								oRec.close
								oConn.close
								
								case "3"
								
								a = 0
								'strEmails = ""
								
								oConn.open strConnect_aktiveDB
								oRec.open strSQL(b), oConn, 3
								while not oRec.EOF
								strEmails = strEmails & oRec("email") & ";"
								
								a = a + 1
								oRec.movenext
								wend
								oRec.close
								oConn.close
								
								
								Response.write "Antal brugere der skal modtage nyhedsbrev: <b>" & a &"</b><br><br>"
								
								case "4"
								
								a = 0
								oConn.open strConnect_aktiveDB
								oRec.open "SELECT licens FROM licens WHERE id = 1", oConn, 3
								while not oRec.EOF
								
								Response.Write "<b>Lincens indehaver: " & oRec("licens") & "</b><br>"
								
								oRec.movenext
								wend
								oRec.close
								
								
								
								oRec.open strSQL(b), oConn, 3
								while not oRec.EOF
								
								Response.Write oRec("tid") & " - "& oRec("tfaktim") & " - "& oRec("tdato") & " - "& oRec("tmnavn") & " - "& oRec("tjobnavn") & " - "& oRec("tjobnr") &" | "& oRec("taktivitetnavn") &" id: "& oRec("taktivitetid") & "<br>"
								        
								
								a = a + 1
								oRec.movenext
								wend
								oRec.close
								oConn.close
								
								case 5
								
								oConn.open strConnect_aktiveDB
								'strSQL = "SELECT md, aar, id FROM ressourcer_md WHERE id <> 0"
	                            oRec.open strSQL(b), oConn, 3
	                            While not oRec.EOF
                        	            
	                                    thisdate = "1/"& oRec("md")&"/"&oRec("aar")
	                                    thisWeek = datepart("ww", thisdate, 2,2)
                        	            
                        	            
	                                    strSQLu = "UPDATE ressourcer_md SET uge = "& thisWeek & " WHERE id = "& oRec("id")
	                                    Response.write "<b>"& thisdate &"</b>: "& strSQLu & "<br>"
	                                    'Response.flush
	                                    'oConn.execute strSQLu
                        	            
                        	    
	                            oRec.movenext
	                            wend 
	                            oRec.close
	                            
	                            oConn.close
	                            
	                            case 6
								
								oConn.open strConnect_aktiveDB
								Response.Write strSQL(b) & "<br>"
								oRec.open strSQL(b), oConn, 3
	                            While not oRec.EOF
                        	            
	                                   Response.Write oRec("id") & " | " & oRec("usrid") & " | " & oRec("matnavn") & "<br>"
                        	           strSQLu = "UPDATE materiale_forbrug SET intkode = 0, personlig = 1 WHERE id = "& oRec("id")
	                                    Response.write ""& strSQLu & "<br>"
	                                    'Response.flush
	                                    'oConn.execute strSQLu
                        	    
	                            oRec.movenext
	                            wend 
	                            oRec.close
	                            
	                            oConn.close
                        	
                        	
	                        'Response.Write "ok"
	                        'Response.end

                                case 7
								
								oConn.open strConnect_aktiveDB
								Response.Write strSQL(b) & "<br>"
								oRec.open strSQL(b), oConn, 3
	                            While not oRec.EOF
                        	            
	                                   Response.Write oRec("tmnavn") & " | " & oRec("tdato") & " | " & oRec("taktivitetnavn") &" | " & oRec("timer") & "<br>"
                        	          
                        	    
	                            oRec.movenext
	                            wend 
	                            oRec.close
	                            
	                            oConn.close


                               case 8
								
								oConn.open strConnect_aktiveDB
								Response.Write strSQL(b) & "<br>"
								oRec.open strSQL(b), oConn, 3
	                            While not oRec.EOF
                        	            
	                                   Response.Write oRec("id") &" "& oRec("navn") & " type: " & oRec("type") & "<br>"
                        	          
                        	    
	                            oRec.movenext
	                            wend 
	                            oRec.close
	                            
	                            oConn.close
                        	
                        	
	                        'Response.Write "ok"
	                        'Response.end
								
								end select
								
								
							
							
						end if
						
					next
		
			next
		
		
Response.Write "<table width=600 cellspacing=0 cellpadding=0 border=0><tr><td>"& strEmails & "</td></tr></table>"
								

		
Response.write "Alle db opdateret, uden fejl."
Response.write "<br>Records: <b>"& a &"</b><br>"
Response.Write "Aktive brugere ialt: <b>" & aktivelbrugereialt & "</b><br>"
end if

else

%>

<form action="dbupdateservice.asp?opdateralledb=1" method="post">
<b>SQL statement:</b> 
<br>Denne sætning vil opdaterer alle TimeOut databaser på produktions serveren.<br>
<b>Split med ";"</b><br><br>
<input type="radio" name="brug_sql" id="brug_sql" value="0" CHECKED>  Brug textarea til opdatering af <u>alle</u> TimeOut Databaser.<br>
<textarea cols="50" rows="7" name="SQL"></textarea>
<br>
<br><b>Eller brug en af disse faste SQL kald:</b><br>
<input type="radio" name="brug_sql" id="brug_sql" value="1"> Brug "SELECT mnavn, mansat, mid FROM medarbejdere WHERE mansat <> 2"
<input type="hidden" name="SQL1" id="SQL1" value="SELECT mnavn, mansat, mid FROM medarbejdere WHERE mansat <> 2"><!-- AND mansat <> 3 -->
<br />
<input type="radio" name="brug_sql" id="brug_sql" value="2"> Hent fakturaer og valuta
<input type="hidden" name="SQL2" id="SQL2" value="SELECT * FROM fakturaer">
<br />
<input type="radio" name="brug_sql" id="Radio1" value="3"> Brug "SELECT email FROM medarbejdere WHERE (nyhedsbrev = 1 AND email <> '') AND mansat <> 2 AND mansat <> 3" / Eller alle admin (brugergrp 3)
<input type="hidden" name="SQL3" id="sql3" value="SELECT email FROM medarbejdere WHERE (brugergruppe = 3 AND email <> '') AND mansat <> 2 AND mansat <> 3"> 'nyhedsbrev = 1 
<br />
<input type="radio" name="brug_sql" id="Radio2" value="4"> Brug "SELECT tid, tdato, tmnavn, tjobnavn, tjobnr, taktivitetnavn, taktivitetid, tfaktim FROM timer WHERE (tfaktim = 12 OR tfaktim = 13) "
<input type="hidden" name="SQL4" id="SQL4" value="SELECT tid, tdato, tmnavn, tjobnavn, tjobnr, taktivitetnavn, taktivitetid, tfaktim FROM timer WHERE (tfaktim = 12 OR tfaktim = 13)" />
<br />
<input type="radio" name="brug_sql" id="Radio3" value="5"> Brug "SELECT md, aar, id FROM ressourcer_md WHERE id <> 0"
<input type="hidden" name="SQL5" id="Hidden1" value="SELECT md, aar, id FROM ressourcer_md WHERE id <> 0" />
<br />
<br />
<input type="radio" name="brug_sql" id="Radio4" value="6"> Brug "SELECT id, matnavn, usrid FROM materiale_forbrug WHERE intkode = 2"
<input type="hidden" name="SQL6" id="Hidden2" value="SELECT id, matnavn, usrid FROM materiale_forbrug WHERE intkode = 2" />
<br /><br />
<input type="radio" name="brug_sql" id="Radio4" value="7"> Brug "SELECT tmnavn, taktivitetnavn, tdato, timer FROM timer WHERE (tfaktim = 19 OR tfaktim = 11) AND tdato BETWEEN '2015-05-01' AND '2015-05-20'"
<input type="hidden" name="SQL7" id="Hidden2" value="SELECT tmnavn, taktivitetnavn, tdato, timer FROM timer WHERE (tfaktim = 19 OR tfaktim = 11) AND tdato BETWEEN '2015-05-01' AND '2015-05-20'" />
<br /><br />
<input type="radio" name="brug_sql" id="Radio4" value="8"> Brug "SELECT id, navn, type FROM leverand WHERE id <> 0"
<input type="hidden" name="SQL8" id="Hidden2" value="SELECT id, navn, type FROM leverand WHERE id <> 0" />


<br />
<br><br>
<br>
<input type="submit" value="Kør SQL -->">
</form>
<br><br>
<u>20061017.1</u><br>
ALTER TABLE materialer ADD (minlager int default 0 NOT NULL);<br>
INSERT INTO dbversion (dbversion) VALUES (20061017.1)
<br><br>

<u>20061030.1</u><br>
ALTER TABLE fakturaer ADD (jobbesk text );
INSERT INTO dbversion (dbversion) VALUES (20061030.1)
<br><br>

<u>20061106.1</u><br>
ALTER TABLE licens ADD (lukafm int default 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20061106.1)
<br><br>

<u>20061106.2</u><br>
ALTER TABLE licens ADD (autogk int default 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20061106.2)
<br><br>

<u>20061110.1</u><br>
ALTER TABLE licens ADD (sdsk int default 0 NOT NULL);
ALTER TABLE kunder ADD (sdskpriogrp int default 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20061110.1)

<br><br>
<u>20061116.1</u><br>
ALTER TABLE sdsk ADD (editor2 varchar (50), dato2 date);
INSERT INTO dbversion (dbversion) VALUES (20061116.1)

<br><br>
<u>20061116.2</u><br>
ALTER TABLE sdsk_prio_grp ADD (gemtider int default 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20061116.2)

<br><br>
<b>20061126.1</b><br>
ALTER TABLE sdsk_status ADD (luk int default 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20061126.1)

<br><br>
<b>20061126.2</b><br>
ALTER TABLE aktiviteter ADD (incidentid int default 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20061126.2)

<br><br>
<b>20061207.1</b><br>
ALTER TABLE licens ADD (
normtid_st_man time DEFAULT '08:00:00' NOT NULL, 
normtid_sl_man time DEFAULT '17:00:00' NOT NULL, 
normtid_st_tir time DEFAULT '08:00:00' NOT NULL, 
normtid_sl_tir time DEFAULT '17:00:00' NOT NULL, 
normtid_st_ons time DEFAULT '08:00:00' NOT NULL, 
normtid_sl_ons time DEFAULT '17:00:00' NOT NULL, 
normtid_st_tor time DEFAULT '08:00:00' NOT NULL, 
normtid_sl_tor time DEFAULT '17:00:00' NOT NULL, 
normtid_st_fre time DEFAULT '08:00:00' NOT NULL, 
normtid_sl_fre time DEFAULT '17:00:00' NOT NULL, 
normtid_st_lor time DEFAULT '08:00:00' NOT NULL, 
normtid_sl_lor time DEFAULT '17:00:00' NOT NULL, 
normtid_st_son time DEFAULT '08:00:00' NOT NULL, 
normtid_sl_son time DEFAULT '17:00:00' NOT NULL
);
INSERT INTO dbversion (dbversion) VALUES (20061207.1)


<br><br>
<b>20061208.1</b><br>
ALTER TABLE licens ADD (brugabningstid int default 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20061208.1)


<br><br>
<b>20061215.1</b><br>
ALTER TABLE sdsk_rel ADD (editor2 varchar(50), dato2 date);
INSERT INTO dbversion (dbversion) VALUES (20061215.1)

<br /><br />
<b>20061221.1</b><br>
ALTER TABLE licens ADD (autolukvdato int default 0 NOT NULL, autolukvdatodato int default 28 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20061221.1)


<br /><br />
<b>20070210.1</b><br />
ALTER TABLE licens ADD (ignorertid_st time DEFAULT '08:00:00' NOT NULL, ignorertid_sl time DEFAULT '08:00:00' NOT NULL );
INSERT INTO dbversion (dbversion) VALUES (20070210.1)

<br /><br />
<b>20070210.2</b><br />
ALTER TABLE login_historik ADD (kommentar text );
INSERT INTO dbversion (dbversion) VALUES (20070210.2)

<br /><br />
<b>20070210.3</b><br />
ALTER TABLE licens ADD (stpause double default 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20070210.3)

<br /><br />
<b>20070210.4</b><br />
ALTER TABLE login_historik ADD (manuelt_afsluttet int default 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20070210.4)

<br /><br />
<b>20070212.1</b><br />
ALTER TABLE licens ADD (stpause2 double default 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20070212.1)



<br /><br />
<b>20070223.1</b><br />
ALTER TABLE licens ADD (p1_man int NOT NULL default 0,
p1_tir int NOT NULL default 0,
p1_ons int NOT NULL default 0,
p1_tor int NOT NULL default 0,
p1_fre int NOT NULL default 0,
p1_lor int NOT NULL default 0,
p1_son int NOT NULL default 0,
p2_man int NOT NULL default 0,
p2_tir int NOT NULL default 0,
p2_ons int NOT NULL default 0,
p2_tor int NOT NULL default 0,
p2_fre int NOT NULL default 0,
p2_lor int NOT NULL default 0,
p2_son int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20070223.1)

<br /><br />
<b>20070402.1</b> (20070222.1)<br />
ALTER TABLE todo_new ADD (tododato date, forvafsl int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20070402.1)


<br /><br />
----
<b>20070403.1</b><br />
ALTER TABLE licens ADD (fakturanr int NOT NULL default 0,
rykkernr int NOT NULL default 0,
kreditnr int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20070403.1)


<br /><br />
<b>20070409.1</b><br />
ALTER TABLE kundetyper ADD (rabat int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20070409.1)

<br /><br />
<b>20070416.1</b><br />
ALTER TABLE faktura_det ADD (rabat double (12,2) default 0 NOT NULL);
ALTER TABLE fak_med_spec ADD (medrabat double (12,2) default 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20070416.1)


<br /><br />
<b>20070417.1</b><br />
ALTER TABLE fakturaer ADD (timersubtotal double (12,2) default 0 NOT NULL ) ;
INSERT INTO dbversion (dbversion) VALUES (20070417.1)

<br /><br />
<b>20070419.1</b><br />
ALTER TABLE faktura_det ADD (enhedsang int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20070419.1)

<br /><br />
<b>20070424.1</b><br />
ALTER TABLE fakturaer ADD (visjoblog int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20070424.1)
 

<br /><br />
<b>20070503.1</b><br />
ALTER TABLE fakturaer ADD (visrabatkol int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20070503.1)

<br /><br />
<b>20070508.1</b><br />
ALTER TABLE fakturaer ADD (vismatlog int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20070508.1)
 

<br /><br />
<b>20070508.2</b><br />
ALTER TABLE fak_med_spec ADD (venter_brugt double (12,2) NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20070508.2)

<br /><br />
<b>20070521.1</b><br />
ALTER TABLE fak_med_spec ADD (enhedsang int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20070521.1)

<br /><br />
<b>20070618.1</b><br />
ALTER TABLE materiale_forbrug ADD (serviceaft int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20070618.1)

<br /><br />
<b>20070618.2</b><br />
ALTER TABLE fakturaer ADD (shadowcopy int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20070618.2)

<br /><br />
<b>20070619.1</b><br />
ALTER TABLE fakturaer ADD (rabat double(12,2) NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20070619.1)
 
<br /><br />
<b>20070702.1</b><br />
ALTER TABLE fakturaer ADD (visjoblog_timepris int NOT NULL default 0, visjoblog_enheder int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20070702.1)

<br /><br />
<b>20070829.1</b><br />
ALTER TABLE materialer ADD (sortorder double NOT NULL default 1000);
INSERT INTO dbversion (dbversion) VALUES (20070829.1)

<br /><br />
<b>20070829.2</b><br />
ALTER TABLE materialer ADD (betegnelse varchar(255));
INSERT INTO dbversion (dbversion) VALUES (20070829.2)


<br /><br />
<b>20070917.1</b><br />
ALTER TABLE licens ADD (crm int NOT NULL default 0, erp int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20070917.1)


<br /><br />
<b>20071001.1</b><br />
ALTER TABLE licens ADD (tilbudsnr int NOT NULL default 0,
jobnr int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20071001.1)

<br /><br />
<b>20071010.1</b><br />
ALTER TABLE licens ADD (fakprocent int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20071010.1)


<br /><br />
<b>20071010.2</b><br />
UPDATE licens SET fakprocent = 25 WHERE id = 1;
INSERT INTO dbversion (dbversion) VALUES (20071010.2)

<br /><br />
<b>20071024.1</b><br />
ALTER TABLE nogletalskoder ADD (ap_type int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20071024.1)

<br /><br />
<b>20071116.1</b><br />
ALTER TABLE job ADD (lukafmjob int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20071116.1)


<br /><br />
<b>20071120.1</b><br />
ALTER TABLE materialer ADD (lokation varchar(255));
INSERT INTO dbversion (dbversion) VALUES (20071120.1)

<br /><br />
<b>20071213.1</b><br />
UPDATE medarbejdere SET pw = MD5(pw) WHERE mid = mid;
INSERT INTO dbversion (dbversion) VALUES (20071213.1)

<br /><br />
<b>20071217.1</b><br />
ALTER table materiale_grp ADD (av double NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20071217.1)

<br /><br />
<b>20071219.1</b><br />
ALTER TABLE fakturaer ADD (visafsfax int(11) NOT NULL DEFAULT 0) ;
INSERT INTO dbversion (dbversion) VALUES (20071219.1)

<br /><br />
<b>20080226.1</b><br />
ALTER TABLE fakturaer ADD (erfakbetalt int(11) NOT NULL DEFAULT 0) ;
INSERT INTO dbversion (dbversion) VALUES (20080226.1)

<br /><br />
<b>20080304.1</b><br />
ALTER TABLE fak_mat_spec ADD (ikkemoms int(11) NOT NULL DEFAULT 0) ;
INSERT INTO dbversion (dbversion) VALUES (20080304.1)

<br /><br />
<b>20080304.2</b><br />
ALTER TABLE fakturaer ADD (subtotaltilmoms double(12,2) NOT NULL DEFAULT 0) ;
INSERT INTO dbversion (dbversion) VALUES (20080304.2)

<br /><br />
<b>20080327.1</b><br />
ALTER TABLE job ADD (valuta int(11) NOT NULL DEFAULT 1) ;
INSERT INTO dbversion (dbversion) VALUES (20080327.1)

<br /><br />
<b>20080327.2</b><br />
ALTER TABLE fakturaer ADD (valuta int(11) NOT NULL DEFAULT 1, kurs double(12,2) NOT NULL );
INSERT INTO dbversion (dbversion) VALUES (20080327.2)


<br /><br />
<b>20080331.1</b><br />
ALTER TABLE fakturaer ADD (sprog int(11) NOT NULL DEFAULT 1) ;
INSERT INTO dbversion (dbversion) VALUES (20080331.1)

<br /><br />
<b>20080425.1</b><br />
ALTER TABLE licens ADD (bgt int(11) NOT NULL DEFAULT 0) ;
INSERT INTO dbversion (dbversion) VALUES (20080425.1)

<br /><br />
<b>20080425.2</b><br />
ALTER TABLE medarbejdere ADD (ansatdato date NOT NULL DEFAULT '2002-01-01') ;
INSERT INTO dbversion (dbversion) VALUES (20080425.2)

<br /><br />
<b>20080507.1</b><br />
ALTER TABLE  budget_medarb_rel  ADD (timepris double(12,2) NOT NULL DEFAULT 0) ;
INSERT INTO dbversion (dbversion) VALUES (20080507.1)

<br /><br />
<b>20080508.1</b><br />
ALTER TABLE medarbejdere ADD (opsagtdato date NOT NULL DEFAULT '2044-01-01') ;
INSERT INTO dbversion (dbversion) VALUES (20080508.1)

<br /><br />
<b>20080508.2</b><br />
ALTER TABLE budget_medarb_rel ADD (ntimerpruge double (12,2) NOT NULL DEFAULT 0) ;
INSERT INTO dbversion (dbversion) VALUES (20080508.2)

<br /><br />
<b>20080510.1</b><br />
ALTER TABLE  stopur ADD (incident int(11) NOT NULL DEFAULT 0, incidentlog int(11) NOT NULL DEFAULT 0, aktid int(11) NOT NULL DEFAULT 0 ) ;
INSERT INTO dbversion (dbversion) VALUES (20080510.1)

<br /><br />
<b>20080513.1</b><br />
ALTER TABLE  stopur ADD (timereg_overfort int(11) NOT NULL DEFAULT 0, sid_godkendt int(11) NOT NULL DEFAULT 0) ;
INSERT INTO dbversion (dbversion) VALUES (20080513.1)

<br /><br />
<b>20080516.1</b><br />
ALTER TABLE  stopur ADD (dato date NOT NULL DEFAULT '2002-01-01', editor varchar(255) ) ;
INSERT INTO dbversion (dbversion) VALUES (20080516.1)

<br /><br />
<b>20080530.1</b><br />
ALTER TABLE  sdsk ADD (creator int(11) NOT NULL DEFAULT 0) ;
INSERT INTO dbversion (dbversion) VALUES (20080530.1)

<br /><br />
<b>20080603.1</b><br />
ALTER TABLE sdsk ADD (kpers int NOT NULL DEFAULT 0, kpers2 varchar(50));
INSERT INTO dbversion (dbversion) VALUES (20080603.1)

<br /><br />
<b>20080603.2</b><br />
ALTER TABLE sdsk ADD (kpers2email varchar(50));
INSERT INTO dbversion (dbversion) VALUES (20080603.2)

<br /><br />
<b>20080604.1</b><br />
ALTER TABLE kunder CHANGE regnr regnr varchar(20), CHANGE kontonr kontonr varchar(20);
INSERT INTO dbversion (dbversion) VALUES (20080604.1)

<br /><br />
<b>20080609.1</b><br />
ALTER TABLE stopur ADD (kommentar text);
INSERT INTO dbversion (dbversion) VALUES (20080609.1)

<br /><br />
<b>20080621.1</b><br />
ALTER TABLE  fakturaer ADD (istdato date NOT NULL DEFAULT '2002-01-01' ) ;
INSERT INTO dbversion (dbversion) VALUES (20080621.1)

<br /><br />
<b>20080708.1</b><br />
ALTER TABLE medarbejdere ADD (sprog int NOT NULL DEFAULT 1);
INSERT INTO dbversion (dbversion) VALUES (20080708.1)

<br /><br />
<b>20080721.1</b><br />
ALTER TABLE materiale_forbrug ADD (valuta int NOT NULL DEFAULT 1, bilagsnr varchar (50), 
intkode int NOT NULL DEFAULT 0 );
INSERT INTO dbversion (dbversion) VALUES (20080721.1)

<br /><br />
<b>20080812.1</b><br />
ALTER TABLE fakturaer ADD (momskonto int NOT NULL DEFAULT 1, visperiode int NOT NULL DEFAULT 0 );
INSERT INTO dbversion (dbversion) VALUES (20080812.1)


<br /><br />
<b>20080813.1</b><br />
ALTER TABLE materiale_forbrug ADD (kurs double(12,2) NOT NULL DEFAULT 100);
INSERT INTO dbversion (dbversion) VALUES (20080813.1)

<br /><br />
<b>20080821.1</b><br />
ALTER TABLE materiale_forbrug ADD (afregnet int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20080821.1)


<br /><br />
<b>200800820.99</b><br />
ALTER TABLE kunder ADD (ean varchar (50));
INSERT INTO dbversion (dbversion) VALUES (200800820.99)

<br /><br />
<b>20080901.1</b><br />
ALTER TABLE timepriser ADD (6valuta int NOT NULL DEFAULT 1);
INSERT INTO dbversion (dbversion) VALUES (20080901.1)

<br /><br />
<b>20080901.2</b><br />
ALTER TABLE medarbejdertyper ADD (tp0_valuta int NOT NULL DEFAULT 1, tp1_valuta int NOT NULL DEFAULT 1, tp2_valuta int NOT NULL DEFAULT 1, tp3_valuta int NOT NULL DEFAULT 1, tp4_valuta int NOT NULL DEFAULT 1, tp5_valuta int NOT NULL DEFAULT 1);
INSERT INTO dbversion (dbversion) VALUES (20080901.2)

<br /><br />
<b>20080903.1</b><br />
ALTER TABLE timer ADD (valuta int NOT NULL DEFAULT 1, kurs double(12,2) NOT NULL DEFAULT 100);
INSERT INTO dbversion (dbversion) VALUES (20080903.1)

<br /><br />
<b>20080912.1</b><br />
ALTER TABLE fak_med_spec ADD (valuta int NOT NULL DEFAULT 1, kurs double(12,2) NOT NULL DEFAULT 100);
INSERT INTO dbversion (dbversion) VALUES (20080912.1)

<br /><br />
<b>20080912.2</b><br />
ALTER TABLE fak_mat_spec ADD (valuta int NOT NULL DEFAULT 1, kurs double(12,2) NOT NULL DEFAULT 100);
INSERT INTO dbversion (dbversion) VALUES (20080912.2)

<br /><br />
<b>20080912.3</b><br />
ALTER TABLE faktura_det ADD (valuta int NOT NULL DEFAULT 1, kurs double(12,2) NOT NULL DEFAULT 100);
INSERT INTO dbversion (dbversion) VALUES (20080912.3)

<br /><br />
<b>20080930.1</b><br />
UPDATE fakturaer SET kurs = 100 WHERE valuta = 1;
ALTER TABLE fakturaer MODIFY kurs double(12,2) NOT NULL DEFAULT 100;
INSERT INTO dbversion (dbversion) VALUES (20080930.1)

<br /><br />
<b>20081002.1</b><br />
ALTER TABLE serviceaft ADD (valuta int NOT NULL DEFAULT 1);
INSERT INTO dbversion (dbversion) VALUES (20081002.1)

<br /><br />
<b>20081006.1</b><br />
ALTER TABLE faktura_rykker ADD (valuta int NOT NULL DEFAULT 1, kurs double(12,2) NOT NULL DEFAULT 100);
INSERT INTO dbversion (dbversion) VALUES (20081006.1)


<br /><br />
<b>20081020.1</b><br />
ALTER TABLE fakturaer ADD (visjoblog_mnavn int NOT NULL DEFAULT 1);
INSERT INTO dbversion (dbversion) VALUES (20081020.1)


<br /><br />
<b>20081103.1</b><br />
ALTER TABLE materialer MODIFY  betegnelse text;
INSERT INTO dbversion (dbversion) VALUES (20081103.1)


<br /><br />
<b>20081106.1</b><br />
ALTER TABLE aktiviteter MODIFY navn varchar(65);
INSERT INTO dbversion (dbversion) VALUES (20081106.1)

<br /><br />
<b>20081107.1</b><br />
ALTER TABLE sdsk ADD (duedate datetime NOT NULL default '2001-1-1 00:00:00', useduedate INT NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20081107.1)

<br /><br />
<b>20081112.1</b><br />
ALTER TABLE job ADD (faktype int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20081112.1)

<br /><br />
<b>20081112.2</b><br />
ALTER TABLE fakturaer ADD (jobfaktype int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20081112.2)


<br /><br />
<b>20081117.1</b><br />
ALTER TABLE job CHANGE faktype jobfaktype int  NOT NULL DEFAULT 0;
INSERT INTO dbversion (dbversion) VALUES (20081117.1)


<br /><br />
<b>20081204.1</b><br />
ALTER TABLE kunder ADD (betbetint int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20081204.1)


<br /><br />
<b>20081204.2</b><br />
ALTER TABLE fakturaer ADD (betbetint int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20081204.2)

<br /><br />
<b>20081212.1</b><br />
ALTER TABLE aktiviteter ADD (antalstk double NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20081212.1)

<br /><br />
<b>20081212.2</b><br />
ALTER TABLE aktiviteter MODIFY navn varchar(100);
INSERT INTO dbversion (dbversion) VALUES (20081212.2)

<br /><br />
<b>20090114.1</b><br />
ALTER TABLE kontoplan MODIFY kontonr double(12,0) NOT NULL;
INSERT INTO dbversion (dbversion) VALUES (20090114.1)

<br /><br />
<b>20090114.2</b><br />
ALTER TABLE medarbejdere ADD (nyhedsbrev int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20090114.2)


<br /><br />
<b>20090114.3</b><br />
ALTER TABLE job ADD kommentar varchar(200);
INSERT INTO dbversion (dbversion) VALUES (20090114.3)

<br /><br />
<b>20090123.1</b><br />
ALTER TABLE job ADD rekvnr varchar(50);
INSERT INTO dbversion (dbversion) VALUES (20090123.1)

<br /><br />
<b>20090212.1</b><br />
ALTER TABLE licens ADD (autogktimer int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20090212.1)

<br /><br />
<b>20090223.1</b><br />
ALTER TABLE fakturaer ADD (istdato2 date NOT NULL default '2002-01-01' );
INSERT INTO dbversion (dbversion) VALUES (20090223.1)

<br /><br />
<b>20090223.2</b><br />
ALTER TABLE fakturaer ADD (brugfakdatolabel int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20090223.2)

<br /><br />
<b>20090226.1</b><br />
ALTER TABLE licens ADD (listatus int NOT NULL DEFAULT 1);
INSERT INTO dbversion (dbversion) VALUES (20090226.1)


<br /><br />
<b>20090227.1</b><br />
CREATE TABLE job_ulev_ju (
ju_id INT NOT NULL AUTO_INCREMENT,
ju_navn VARCHAR(100) NOT NULL,
ju_ipris DOUBLE(12,2) NOT NULL DEFAULT 0,
ju_faktor DOUBLE(12,2) NOT NULL DEFAULT 0,
ju_belob DOUBLE(12,2) NOT NULL DEFAULT 0,
ju_jobid INT NOT NULL DEFAULT 0,
PRIMARY KEY (ju_id)
);
INSERT INTO dbversion (dbversion) VALUES (20090227.1)


<br /><br />
<b>20090228.1</b><br />
ALTER TABLE job ADD (
jo_gnstpris double (12,2) NOT NULL DEFAULT 0,
jo_gnsfaktor double (12,2) NOT NULL DEFAULT 1,
jo_gnsbelob double (12,2) NOT NULL DEFAULT 0,
jo_bruttofortj double (12,2) NOT NULL DEFAULT 0,
jo_dbproc double (12,2) NOT NULL DEFAULT 0
);
INSERT INTO dbversion (dbversion) VALUES (20090228.1)


<br /><br />
<b>20090314.1</b><br />
ALTER TABLE faktura_det ADD (fak_sortorder double(12,2) NOT NULL DEFAULT 1);
INSERT INTO dbversion (dbversion) VALUES (20090314.1)

<br /><br />
<b>20090409.1</b><br />
ALTER TABLE timer MODIFY taktivitetnavn varchar (100);
INSERT INTO dbversion (dbversion) VALUES (20090409.1)



<br /><br />
<B>20090410.1</B><br />
CREATE TABLE akt_typer (
aty_id  INT NOT NULL AUTO_INCREMENT,
aty_label VARCHAR(100),
aty_desc VARCHAR(100),
aty_on INT NOT NULL DEFAULT 0,
aty_on_realhours INT NOT NULL DEFAULT 0,
aty_on_invoiceble INT NOT NULL DEFAULT 0,
aty_on_invoice INT NOT NULL DEFAULT 0,
aty_on_invoice_chk INT NOT NULL DEFAULT 0,
aty_on_workhours INT NOT NULL DEFAULT 0,
aty_pre DOUBLE(12,2) NOT NULL DEFAULT 0,
aty_sort DOUBLE(4,2) NOT NULL DEFAULT 0,
aty_on_recon INT NOT NULL DEFAULT 0,
aty_enh INT NOT NULL DEFAULT 0,
aty_on_adhoc INT NOT NULL DEFAULT 0,
PRIMARY KEY (aty_id)
);



INSERT INTO akt_typer (aty_id,
 aty_label, aty_desc,
aty_on, aty_on_realhours, aty_on_invoiceble, aty_on_invoice, aty_on_invoice_chk 
, aty_on_workhours, aty_pre, aty_sort, aty_on_recon, aty_enh, aty_on_adhoc)
VALUES 
(1, 'global_txt_129','Fakturerbare',1,1,1,1,1,0,0,1.0,1,0,0),
(2, 'global_txt_131','Ikke fakturerbar',1,1,2,1,0,0,0,1.1,1,0,0),
(5, 'global_txt_130','Km',1,0,1,1,1,0,0,2.0,1,3,0),
(6, 'global_txt_132','Salg  | Newbizz',1,1,2,0,0,0,0,1.4,1,0,0),
(10, 'global_txt_133','Frokost',1,0,0,0,0,1,0.5,2.1,1,0,0),
(15, 'global_txt_143','Ferie optjent',1,1,0,0,0,0,0,3.0,1,0,0),
(11, 'global_txt_134','Ferie planlagt',1,0,0,0,0,0,0,3.1,1,0,0),
(14, 'global_txt_135','Ferie afholdt',1,1,0,0,0,0,0,3.2,1,0,0),
(16, 'global_txt_156','Ferie udbetalt',1,0,0,0,0,0,0,3.3,1,0,0),
(12, 'global_txt_136','Ferie fridage optjent',1,0,0,0,0,0,0,3.4,1,0,0),
(13, 'global_txt_137','Ferie fridage brugt',1,1,0,0,0,0,0,3.5,1,0,0),
(17, 'global_txt_149','Ferie fridage udbetalt',1,0,0,0,0,0,0,3.6,1,0,0),
(30, 'global_txt_144','Overarbejde',1,0,0,0,0,0,0,4.0,1,0,0),
(31, 'global_txt_145','Afspadsering',1,1,0,0,0,0,0,4.1,1,0,0),
(32, 'global_txt_146','Afspad. udbetalt',1,0,0,0,0,0,0,4.2,1,0,0),
(33, 'global_txt_159','Afspad. ønskes udb.',1,0,0,0,0,0,0,4.3,1,0,0),
(20, 'global_txt_138','Syg',1,1,0,0,0,0,0,5.0,1,0,0),
(21, 'global_txt_139','Barn syg',1,1,0,0,0,0,0,5.1,1,0,0),
(7, 'global_txt_147','Flex brugt',0,1,0,0,0,0,0,2.3,0,0,0),
(8, 'global_txt_148','Sundhed',0,1,0,0,0,0,0,5.2,0,0,0),
(9, 'global_txt_150','Pause',0,0,0,0,0,0,0.33,2.2,0,0,0),
(51, 'global_txt_151','Nat',0,1,1,0,0,0,0,1.70,0,0,0),
(52, 'global_txt_152','Weekend',0,1,1,0,0,0,0,1.71,0,0,0),
(53, 'global_txt_153','Weekend Nat',0,1,1,0,0,0,0,1.72,0,0,0),
(54, 'global_txt_154','Aften',0,1,1,0,0,0,0,1.73,0,0,0),
(55, 'global_txt_155','Weekend Aften',0,1,1,0,0,0,0,1.74,0,0,0),
(60, 'global_txt_157','Ad-hoc',0,1,1,1,1,0,0,1.8,1,0,1),
(61, 'global_txt_158','Stk. Antal',0,0,1,0,0,0,0,2.4,1,1,0),
(81, 'global_txt_160','Læge',0,1,0,0,0,0,0,5.3,0,0,0),
(90, 'global_txt_161','E1',0,1,1,1,1,0,0,1.90,1,0,0),
(91, 'global_txt_162','E2',0,1,1,1,1,0,0,1.91,1,0,0)
;

INSERT INTO dbversion (dbversion) VALUES (20090410.1);

<br /><br />
<B>20090430.1</B><br />
UPDATE aktiviteter SET fakturerbar = 5 WHERE fakturerbar = 2;
UPDATE aktiviteter SET fakturerbar = 2 WHERE fakturerbar = 0;
INSERT INTO dbversion (dbversion) VALUES (20090430.1);

<br /><br />
<b>20090501.1</b><br />
ALTER TABLE login_historik ADD (login_first datetime, logud_first datetime);
INSERT INTO dbversion (dbversion) VALUES (20090501.1)

<br /><br />
<b>20090505.2</b><br />
UPDATE akt_typer SET aty_pre = 0 WHERE aty_id = 10;
INSERT INTO dbversion (dbversion) VALUES (20090505.2)

<br /><br />
<b>20090527.1</b><br />
ALTER TABLE login_historik ADD (manuelt_oprettet int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20090527.1)

<br /><br />
<b>20090603.1</b><br />
ALTER TABLE fakturaer DROP index fnr;
ALTER TABLE fakturaer ADD index fnr (faknr);
INSERT INTO dbversion (dbversion) VALUES (20090603.1)


<br /><br />
<b>20090608.1</b><br />
ALTER TABLE sdsk ADD (priotype int NOT NULL default 0, sortorder int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20090608.1)

<br /><br />
<b>20090608.2</b><br />
ALTER TABLE aktiviteter ADD (tidslaas_man int NOT NULL default 0,
tidslaas_tir int NOT NULL default 0, tidslaas_ons int NOT NULL default 0,
tidslaas_tor int NOT NULL default 0,tidslaas_fre int NOT NULL default 0,
tidslaas_lor int NOT NULL default 0,tidslaas_son int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20090608.2)

<br /><br />
<b>20090609.1</b><br />
ALTER TABLE timer ADD (bopal int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20090609.1)


<br /><br />
<b>20090612.1</b><br />
ALTER TABLE medarbejdere ADD (madr varchar(100), mpostnr varchar(50), mcity varchar(100), mland varchar(50),
mtlf varchar(50));
INSERT INTO dbversion (dbversion) VALUES (20090612.1)


<br /><br />
<b>20090622.1</b><br />
ALTER TABLE fakturaer ADD (fakbetkom varchar(250));
INSERT INTO dbversion (dbversion) VALUES (20090622.1)


<br /><br />
<b>20090623.1</b><br />
ALTER TABLE projektgrupper ADD (opengp int NOT NULL default 1);
INSERT INTO dbversion (dbversion) VALUES (20090623.1)

<br /><br />
<b>20090623.2</b><br />
UPDATE projektgrupper SET opengp = 0 WHERE id =10;
INSERT INTO dbversion (dbversion) VALUES (20090623.2)


<br /><br />
<b>20090623.3</b><br />
ALTER TABLE medarbejdere ADD (mcpr varchar(50), mkoregnr varchar(50));
INSERT INTO dbversion (dbversion) VALUES (20090623.3)


<br /><br />
<b>20090625.1</b><br />
ALTER TABLE timer ADD (destination varchar (50));
INSERT INTO dbversion (dbversion) VALUES (20090625.1)


<br /><br />
<b>20090626.1</b><br />
ALTER TABLE licens ADD (kmdialog int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20090626.1)


<br /><br />
<b>20090701.1</b><br />
UPDATE akt_typer SET aty_on_invoiceble = 0 WHERE aty_id = 5;
INSERT INTO dbversion (dbversion) VALUES (20090701.1)

<br /><br />
<b>20090708.1</b><br />
ALTER TABLE licens ADD (licensstdato date NOT NULL default '2002-01-01');
INSERT INTO dbversion (dbversion) VALUES (20090708.1)

<br /><br />
<b>20090713.1</b><br />
ALTER TABLE ressourcer_md ADD (uge int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20090713.1)

<br /><br />
<b>20090812.1</b><br />
ALTER TABLE login_historik ADD (ipn varchar (50) NOT NULL default '0');
INSERT INTO dbversion (dbversion) VALUES (20090812.1)

<br /><br />
<b>20090828.1</b><br />
INSERT INTO akt_typer (aty_id,
 aty_label, aty_desc,
aty_on, aty_on_realhours, aty_on_invoiceble, aty_on_invoice, aty_on_invoice_chk 
, aty_on_workhours, aty_pre, aty_sort, aty_on_recon, aty_enh, aty_on_adhoc)
VALUES 
(18, 'global_txt_164','Ferie Fridage Planlagt',1,0,0,0,0,0,0,3.7,1,0,0),
(19, 'global_txt_165','Ferie afh. u. Løn',1,1,0,0,0,0,0,3.8,1,0,0),
(50, 'global_txt_166','Dag',0,1,1,0,0,0,0,1.75,0,0,0);
INSERT INTO dbversion (dbversion) VALUES (20090828.1)

<br /><br />
<b>20090925.1</b><br />
ALTER TABLE aktiviteter ADD (fase varchar(150));
INSERT INTO dbversion (dbversion) VALUES (20090925.1)


<br /><br />
<b>20090928.1</b><br />
ALTER TABLE materiale_forbrug ADD (personlig int NOT NULL default 0, 
godkendt int NOT NULL default 0, gkaf varchar(150), gkdato date);
INSERT INTO dbversion (dbversion) VALUES (20090928.1)


<br /><br />
<b>20091005.1</b><br />
CREATE TABLE tilbuds_skabeloner (
ts_id  INT NOT NULL AUTO_INCREMENT,
ts_navn VARCHAR(100) NOT NULL,
ts_txt longtext,
ts_dato date NOT NULL,
ts_editor VARCHAR(100) NOT NULL,
ts_kundeid INT NOT NULL DEFAULT 0,
PRIMARY KEY (ts_id)
);

<br /><br />
INSERT INTO dbversion (dbversion) VALUES (20091005.1)



<br /><br />
<b>20091012.1</b><br />
CREATE TABLE enheds_typer (
et_id  INT NOT NULL,
et_navn VARCHAR(100) NOT NULL,
PRIMARY KEY (et_id)
);

INSERT INTO enheds_typer (et_id, et_navn) VALUES 
(-1, 'ingen'),
(0, 'time'),
(1, 'stk.'),
(2, 'enhed'),
(3, 'km');

INSERT INTO dbversion (dbversion) VALUES (20091012.1)



<br /><br />
<b>20091029.1</b><br />

ALTER TABLE filer CHANGE dato dato date NOT NULL;
ALTER TABLE filer ADD (filertxt longtext);
INSERT INTO dbversion (dbversion) VALUES (20091029.1)

 
 
 
<br /><br />

<br /><br />
<b>20091030.1</b><br />
CREATE TABLE pdf_values (
id  INT NOT NULL AUTO_INCREMENT,
pdf_txt  longtext,
pdf_footer text,
PRIMARY KEY (id)
);


INSERT INTO dbversion (dbversion) VALUES (20091030.1)
 

<br /><br />

<br /><br />
<b>20091105.1</b><br />
INSERT INTO foldere (id, navn, kundeid, kundese, jobid, editor, dato) VALUES 
(1000, 'Print filer', 0, 0, 0, 'TimeOut Support', '2009-11-05');
		

INSERT INTO dbversion (dbversion) VALUES (20091105.1)


<br /><br />
<b>20091105.2</b><br />
ALTER TABLE kunder CHANGE adresse adresse varchar (200);
INSERT INTO dbversion (dbversion) VALUES (20091105.2)


<br /><br />
<b>20091127.1 - 20091208.2</b><br />
ALTER TABLE faktura_det ADD (momsfri int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20091127.1);

ALTER TABLE fakturaer ADD (momssats int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20091206.1);

ALTER TABLE fakturaer ADD (modtageradr varchar(250), usealtadr INT Not Null default 0, vorref varchar(150));
INSERT INTO dbversion (dbversion) VALUES (20091208.1);

ALTER TABLE aktiviteter ADD (sortorder INT Not Null default 1000);
INSERT INTO dbversion (dbversion) VALUES (20091208.2)


<br /><br />
<b>20091210.1 - 20091215.1</b><br />
ALTER TABLE aktiviteter ADD (bgr INT Not Null default 0, aktbudgetsum double(10,2) not NULL default 0);
ALTER TABLE job ADD (usejoborakt_tp INT Not Null default 0);
INSERT INTO dbversion (dbversion) VALUES (20091210.1);

ALTER TABLE job ADD (ski INT Not Null default 0, job_internbesk text);
INSERT INTO dbversion (dbversion) VALUES (20091215.1);


<br /><br />
<b>20091215.3 - 20091223.1</b><br />
ALTER TABLE fak_mat_spec ADD (matfrb_mid INT Not Null default 0, matfrb_id double(12,0) Not Null default 0);
ALTER TABLE materiale_forbrug ADD (erfak INT Not Null default 0);
INSERT INTO dbversion (dbversion) VALUES (20091215.3);

ALTER TABLE fakturaer ADD (showmatasgrp INT Not Null default 0);
ALTER TABLE fak_mat_spec ADD (matgrp INT Not Null default 0);
INSERT INTO dbversion (dbversion) VALUES (20091223.1)

<br /><br />
<b>20091215.2</b><br />
ALTER TABLE fakturaer ADD (fak_ski INT Not Null default 0);
INSERT INTO dbversion (dbversion) VALUES (20091215.2) 

<br /><br />
<b>20100121.1</b><br />
ALTER TABLE fakturaer ADD (hidesumaktlinier INT Not Null default 0);
INSERT INTO dbversion (dbversion) VALUES (20100121.1)

<br /><br />
<b>20100223.1</b><br />
ALTER TABLE fakturaer ADD (sideskiftlinier INT Not Null default 0, labeldato date);
INSERT INTO dbversion (dbversion) VALUES (20100223.1)


<br /><br />
<b>20100223.2</b><br />
UPDATE fakturaer SET labeldato = istdato2 WHERE fid <> 0;
INSERT INTO dbversion (dbversion) VALUES (20100223.2)


<br /><br />
<b>20100224.1</b><br />
ALTER TABLE akt_gruppe ADD (forvalgt INT Not Null default 0);
INSERT INTO dbversion (dbversion) VALUES (20100224.1)


<br /><br />
<b>20100224.2</b><br />
ALTER TABLE progrupperelationer ADD (teamleder INT Not Null default 0);
INSERT INTO dbversion (dbversion) VALUES (20100224.2)

<br /><br />
<b>20100318.1</b><br />
ALTER TABLE fakturaer CHANGE konto konto double (12,0), CHANGE modkonto modkonto double(12,2);
ALTER TABLE posteringer CHANGE kontonr kontonr double (12,0), CHANGE modkontonr modkontonr double(12,2);
INSERT INTO dbversion (dbversion) VALUES (20100318.1)

<br /><br />
<b>20100325.1</b><br />
ALTER TABLE job ADD (abo INT Not Null default 0, ubv INT Not Null default 0);
INSERT INTO dbversion (dbversion) VALUES (20100325.1)

<br /><br />
<b>20100325.2</b><br />
ALTER TABLE fakturaer ADD (fak_abo INT Not Null default 0, fak_ubv INT Not Null default 0);
INSERT INTO dbversion (dbversion) VALUES (20100325.2)


<br /><br />
<b>20100329.1</b><br />
ALTER TABLE timereg_usejob ADD (easyreg INT Not Null default 0);
INSERT INTO dbversion (dbversion) VALUES (20100329.1)
 

<br /><br />
<b>20100330.1</b><br />
ALTER TABLE aktiviteter ADD (easyreg INT Not Null default 0);
INSERT INTO dbversion (dbversion) VALUES (20100330.1)

<br /><br />
<b>20100330.2</b><br />
UPDATE akt_typer SET aty_on_adhoc = 0;
INSERT INTO dbversion (dbversion) VALUES (20100330.2)

<br /><br />
<b>20100416.1</b><br />
CREATE TABLE login_historik_terminal (
lht_id  INT NOT NULL AUTO_INCREMENT,
lht_dt timestamp,
lht_logind datetime,
lht_mid INT NOT NULL DEFAULT 0,
lht_type INT NOT NULL DEFAULT 0,
lht_status INT NOT NULL DEFAULT 0,
PRIMARY KEY (lht_id)
);
INSERT INTO dbversion (dbversion) VALUES (20100416.1);

<br /><br />
<b>20100416.2</b><br />
ALTER table job ADD (stade_tim_proc INT NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20100416.2)

<br /><br />
<b>20100424.1</b><br />
ALTER table job ADD (sandsynlighed DOUBLE (12,2) NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20100424.1)

<br /><br />
<b>20100504.1</b><br />
ALTER table aktiviteter ADD INDEX (easyreg);
INSERT INTO dbversion (dbversion) VALUES (20100504.1)


<br /><br />
<b>20100520.1</b><br />
CREATE TABLE  lon_korsel  (
lk_id  INT NOT NULL AUTO_INCREMENT,
lk_dato date,
lk_editor VARCHAR(100),
lk_stamp timestamp,
PRIMARY KEY (lk_id)
);
INSERT INTO dbversion (dbversion) VALUES (20100520.1)

<br /><br />
<b>20100602.1</b><br />
ALTER TABLE fakturaer ADD (visikkejobnavn int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20100602.1)

<br /><br />
<b>20100608.1</b><br />
ALTER TABLE job ADD (jobans3 int NOT NULL default 0, 
jobans4 int NOT NULL default 0, 
jobans5 int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20100608.1)  

<br /><br />
<b>20100616.1</b><br />
ALTER TABLE medarbejdere ADD (bonus INT DEFAULT 0 NOT  NULL);
INSERT INTO dbversion (dbversion) VALUES (20100616.1)

<br /><br />
<b>20100616.2</b><br />
ALTER table job ADD (diff_timer DOUBLE (12,2) NOT NULL DEFAULT 0, diff_sum DOUBLE (12,2) NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20100616.2)

<br /><br />
<b>20100617.1</b><br />
ALTER table job ADD (jo_udgifter_intern DOUBLE (12,2) NOT NULL DEFAULT 0,
jo_udgifter_ulev DOUBLE (12,2) NOT NULL DEFAULT 0,
jo_bruttooms DOUBLE (12,2) NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20100617.1)


<br /><br />
<b>20100619.1</b><br />
CREATE TABLE  job_ulev_jugrp  (
jugrp_id  INT NOT NULL AUTO_INCREMENT,
jugrp_dato date,
jugrp_editor VARCHAR(100),
jugrp_navn  VARCHAR (200),
jugrp_forvalgt INT NOT NULL DEFAULT 0,
PRIMARY KEY (jugrp_id)
);
INSERT INTO dbversion (dbversion) VALUES (20100619.1)

<br /><br />
<b>20100619.2</b><br />
ALTER table job_ulev_ju ADD (ju_favorit INT NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20100619.2)

<br /><br />
<b>20100619.3</b><br />
ALTER table job_ulev_ju ADD (ju_fase VARCHAR (200));
INSERT INTO dbversion (dbversion) VALUES (20100619.3)

<br /><br />
<b>20100701.1</b><br />
ALTER table kunder MODIFY bank VARCHAR (200);
INSERT INTO dbversion (dbversion) VALUES (20100701.1)

<br /><br />
<b>20100706.1</b><br />
ALTER table job ADD 
(
jobans_proc_1 DOUBLE (12,2) NOT NULL DEFAULT 0,
jobans_proc_2 DOUBLE (12,2) NOT NULL DEFAULT 0,
jobans_proc_3 DOUBLE (12,2) NOT NULL DEFAULT 0,
jobans_proc_4 DOUBLE (12,2) NOT NULL DEFAULT 0,
jobans_proc_5 DOUBLE (12,2) NOT NULL DEFAULT 0
)
;
INSERT INTO dbversion (dbversion) VALUES (20100706.1)

<br /><br />
<b>20100731.1</b><br />
ALTER table milepale_typer ADD (mpt_fak INT NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20100731.1)

<br /><br />
<b>20100731.2</b><br />
ALTER table fakturaer ADD (hidefasesum INT NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20100731.2)


<br /><br />
<b>20100731.3</b><br />
ALTER table faktura_det ADD (fase VARCHAR(255));
INSERT INTO dbversion (dbversion) VALUES (20100731.3)

<br /><br />
<b>20100810.1-..</b><br />
ALTER table timer ADD (extSysId double(12,0) NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20100810.1);
<br /><br />
CREATE TABLE  timer_imp_err  (
id  INT NOT NULL AUTO_INCREMENT,
dato date,
extsysid DOUBLE(12,0) NOT NULL DEFAULT 0,
errId INT NOT NULL DEFAULT 0,
PRIMARY KEY (id)
);
INSERT INTO dbversion (dbversion) VALUES (20100810.2);
<br /><br />
ALTER table fakturaer ADD (hideantenh INT NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20100816.1)

<br /><br />
<b>20100817.1</b><br />
ALTER table login_historik_terminal ADD (fo_logud_kode INT NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20100817.1)

<br /><br />
<b>20100913.1</b><br />
alter table materiale_forbrug CHANGE matantal matantal double(12,2) NOT NULL;
INSERT INTO dbversion (dbversion) VALUES (20100913.1)

<br /><br />
<b>20100913.2</b><br />
alter table fakturaer ADD medregnikkeioms INT NOT NULL default 0;
INSERT INTO dbversion (dbversion) VALUES (20100913.2)


<br /><br />
<b>20100916.1</b><br />
alter table milepale ADD aftaleid INT NOT NULL default 0;
INSERT INTO dbversion (dbversion) VALUES (20100916.1)


<br /><br />
<b>20100916.2</b><br />
alter table licens ADD timeout_version varchar (50) NOT NULL default 'ver2_1';
INSERT INTO dbversion (dbversion) VALUES (20100916.2)

<br /><br />
<b>20100921.1</b><br />
alter table todo_new ADD public INT NOT NULL default 0;
INSERT INTO dbversion (dbversion) VALUES (20100921.1)


<br /><br />
<b>20100930.1</b><br />
ALTER table licens ADD (fakturanr_kladde double (12,0) NOT NULL DEFAULT 3010001; 
INSERT INTO dbversion (dbversion) VALUES (20100930.1)

<br /><br />
<b>20100930.2</b><br />
ALTER table licens CHANGE fakturanr fakturanr double(12,0) NOT NULL;
ALTER table licens CHANGE kreditnr kreditnr double(12,0) NOT NULL;
ALTER table licens CHANGE jobnr jobnr double(12,0) NOT NULL;
ALTER table licens CHANGE tilbudsnr tilbudsnr double(12,0) NOT NULL;
INSERT INTO dbversion (dbversion) VALUES (20100930.2)

<br /><br />
<b>20100930.3</b><br />
ALTER table fakturaer ADD (fak_laast int NOT NULL DEFAULT 0); 
INSERT INTO dbversion (dbversion) VALUES (20100930.3)


<br /><br />
<b>20101027.1</b><br />
CREATE TABLE  fomr_rel  (
for_id  INT NOT NULL AUTO_INCREMENT,
for_fomr INT NOT NULL DEFAULT 0,
for_jobid DOUBLE(12,0) NOT NULL DEFAULT 0,
for_aktid DOUBLE(12,0) NOT NULL DEFAULT 0,
for_faktor DOUBLE (12,2) NOT NULL DEFAULT 1,
PRIMARY KEY (for_id)
);
INSERT INTO dbversion (dbversion) VALUES (20101027.1)

<br /><br />
<b>20101101.1</b><br />
ALTER TABLE fomr_rel ADD INDEX (for_aktid);
ALTER TABLE fomr_rel ADD INDEX (for_fomr);
INSERT INTO dbversion (dbversion) VALUES (20101101.1)

<br /><br />
<b>20101117.1</b><br />
ALTER table licens ADD (kmpris double (12,2) NOT NULL DEFAULT 3.56); 
INSERT INTO dbversion (dbversion) VALUES (20101117.1)

<br /><br />
<b>20101212.1</b><br />
UPDATE fakturaer SET fak_laast = 1 WHERE betalt = 1; 
INSERT INTO dbversion (dbversion) VALUES (20101212.1)

<br /><br />
<b>20110107.1</b><br />
ALTER table timer ADD (origin int NOT NULL DEFAULT 0);
ALTER table timer_imp_err ADD (origin int NOT NULL DEFAULT 0); 
INSERT INTO dbversion (dbversion) VALUES (20110107.1)

<br /><br />
<b>20110125.1</b><br />
ALTER table akt_typer ADD (aty_hide_on_treg int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20110125.1)


<br /><br />
<b>20110225.1</b><br />
ALTER table materiale_forbrug ADD (extsysid double NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20110225.1)

<br /><br />
<b>20110315.1</b><br />
ALTER table ressourcer_md ADD (aktid int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20110315.1)

<br /><br />
<b>20110321.1</b><br />
CREATE INDEX medarb_aktid_jobid
ON timepriser (medarbid,jobid, aktid);
INSERT INTO dbversion (dbversion) VALUES (20110321.1)

<br /><br />
<b>20110324.1</b><br />
INSERT INTO akt_typer (aty_id,
 aty_label, aty_desc,
aty_on, aty_on_realhours, aty_on_invoiceble, aty_on_invoice, aty_on_invoice_chk 
, aty_on_workhours, aty_pre, aty_sort, aty_on_recon, aty_enh, aty_on_adhoc, aty_hide_on_treg)
VALUES (22, 'global_txt_171','Barsel',1,0,0,0,1,0,0,5.41,1,0,0,0)
; INSERT INTO dbversion (dbversion) VALUES (20110324.1)


<br /><br />
<B>20110404.1</B><br />
CREATE TABLE ressourcer_ramme (
id  INT NOT NULL AUTO_INCREMENT,
jobid INT NOT NULL DEFAULT 0,
medid INT NOT NULL DEFAULT 0,
aar INT NOT NULL DEFAULT 0,
timer DOUBLE(12,2) NOT NULL DEFAULT 0,
PRIMARY KEY (id)
);
INSERT INTO dbversion (dbversion) VALUES (20110404.1)


<br /><br />
<B>20110407.1</B><br />
CREATE INDEX jam_inx2
ON ressourcer_md (jobid, aktid, medid);
INSERT INTO dbversion (dbversion) VALUES (20110407.1)

<br /><br />
<B>20110427.1</B><br />
ALTER table job_ulev_ju ADD (ju_stk double (12,2) NOT NULL DEFAULT 1, ju_stkpris double (12,2) NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20110427.1)

<br /><br />
<B>20110511.1</B><br />
CREATE INDEX jobid_tp_inx
ON timepriser (jobid);
INSERT INTO dbversion (dbversion) VALUES (20110511.1)

<br /><br />
<b>20110516.1</b><br />
ALTER table licens ADD (p1_grp varchar (50), p2_grp varchar (50));
INSERT INTO dbversion (dbversion) VALUES (20110516.1)

<br /><br />
<b>20110527.1</b><br />
CREATE TABLE medarbejdertyper_timebudget (
id  INT NOT NULL AUTO_INCREMENT,
jobid INT NOT NULL DEFAULT 0,
typeid INT NOT NULL DEFAULT 0,
timer DOUBLE(12,2) NOT NULL DEFAULT 0,
timepris DOUBLE(12,2) NOT NULL DEFAULT 0,
faktor DOUBLE(12,2) NOT NULL DEFAULT 0,
belob DOUBLE(12,2) NOT NULL DEFAULT 0,
PRIMARY KEY (id)
);
INSERT INTO dbversion (dbversion) VALUES (20110527.1)

<br /><br />
<b>20110623.1 - 20110706.1</b><br />
ALTER TABLE timereg_usejob ADD  (forvalgt int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20110623.1);
CREATE INDEX tu_inx
ON timereg_usejob (jobid, medarb);
INSERT INTO dbversion (dbversion) VALUES (20110627.2);
UPDATE projektgrupper SET navn = 'Alle-gruppen (alle medarbejdere)' WHERE id = 10;
INSERT INTO dbversion (dbversion) VALUES (20110705.1);
alter table job change kommentar kommentar text;
INSERT INTO dbversion (dbversion) VALUES (20110706.1);

<br /><br />
<b>20110710.1 - 20110710.4</b><br />
CREATE INDEX job_inx3
ON fakturaer (jobid);
INSERT INTO dbversion (dbversion) VALUES (20110710.1);
CREATE INDEX fakja_inx
ON fakturaer (jobid, aftaleid);
INSERT INTO dbversion (dbversion) VALUES (20110710.4);

<br /><br />
<b>20110818.1</b><br />
ALTER TABLE aktiviteter ADD  (fravalgt int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20110818.1)

<br /><br />
<b>20110819.1</b><br />
ALTER TABLE job_ulev_ju ADD  (ju_fravalgt int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20110819.1)

<br /><br />
<b>20110829.1</b><br />
ALTER TABLE medarbejdertyper_timebudget ADD  (belob_ff double(12,2) NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20110829.1)

<br /><br />
<b>20110830.1</b><br />
ALTER TABLE medarbejdertyper ADD  (sostergp int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20110830.1)


<br /><br />
<b>20110902.1</b><br />
ALTER TABLE medarbejdertyper ADD  (mtsortorder int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20110902.1)

<br /><br />
<b>20110906.1</b><br />
CREATE UNIQUE INDEX job_jobnr_inx 
ON job (jobnr);
INSERT INTO dbversion (dbversion) VALUES (20110906.1)

<br /><br />
<b>20110907.1</b><br />
CREATE INDEX mf_jobid_inx 
ON materiale_forbrug (jobid);
INSERT INTO dbversion (dbversion) VALUES (20110907.1)

<br /><br />
<b>20110909.1</b><br />
CREATE TABLE  timer_konsolideret  (
tk_id  INT NOT NULL AUTO_INCREMENT,
tk_dato date,
tk_timerid DOUBLE(12,0) NOT NULL DEFAULT 0,
tk_mnr DOUBLE(12,0) NOT NULL DEFAULT 0,
tk_jobnr DOUBLE(12,0) NOT NULL DEFAULT 0,
tk_aid DOUBLE(12,0) NOT NULL DEFAULT 0,
tk_timer DOUBLE(12,0) NOT NULL DEFAULT 0,
PRIMARY KEY (tk_id)
);
INSERT INTO dbversion (dbversion) VALUES (20110909.1)


<br /><br />
<b>20110909.2</b><br />
ALTER TABLE timer_konsolideret ADD  (dato_kons date);
INSERT INTO dbversion (dbversion) VALUES (20110909.2)

<br /><br>
<b>20110909.3</b><br />
alter table timer_konsolideret change tk_timer tk_timer double(12,2);
INSERT INTO dbversion (dbversion) VALUES (20110909.3)


<br /><br>
<b>20110909.4</b><br />
CREATE UNIQUE INDEX tk_timerid_inx 
ON timer_konsolideret (tk_timerid);
INSERT INTO dbversion (dbversion) VALUES (20110909.4)


<br /><br /><b>20110818.1</b><br />
ALTER TABLE timereg_usejob ADD  (forvalgt_sortorder int NOT NULL DEFAULT 0, forvalgt_af int NOT NULL DEFAULT 0, forvalgt_dt date);
INSERT INTO dbversion (dbversion) VALUES (20110818.1)


<br /><br /><b>20110916.1</b><br />
alter table job ADD virksomheds_proc double(12,2) NOT NULL default 0;
INSERT INTO dbversion (dbversion) VALUES (20110916.1)

<br /><br /><b>20110921.1</b><br />
alter table medarbejdertyper_timebudget ADD kostpris double(12,2) NOT NULL default 0;
INSERT INTO dbversion (dbversion) VALUES (20110921.1)


<br /><br /><b>20111003.1</b><br />
alter table milepale ADD belob double(12,2) NOT NULL default 0;
INSERT INTO dbversion (dbversion) VALUES (20111003.1)


<br /><br /><b>20111009.1</b><br />
UPDATE medarbejdere SET visguide = 11 WHERE mid <> 0;
INSERT INTO dbversion (dbversion) VALUES (20111009.1)



<br /><br /><b>20111027.1</b><br />
alter table materialer CHANGE indkobspris indkobspris double(12,3)  NOT NULL default 0;
alter table materialer CHANGE salgspris salgspris double(12,3)  NOT NULL default 0;

alter table materiale_forbrug CHANGE matkobspris matkobspris double(12,3)  NOT NULL default 0;
alter table materiale_forbrug CHANGE matsalgspris matsalgspris double(12,3)  NOT NULL default 0;

alter table fak_mat_spec CHANGE matenhedspris matenhedspris double(12,3)  NOT NULL default 0;
alter table fak_mat_spec CHANGE matbeloeb matbeloeb double(12,2)  NOT NULL default 0;


INSERT INTO dbversion (dbversion) VALUES (20111027.1)


<br /><br /><b>20111212.1</b><br />
alter table ugestatus ADD (ugegodkendt int NOT NULL default 0, ugegodkendtaf int NOT NULL default 0, 
ugegodkendtTxt varchar(50), ugegodkendtdt date);
INSERT INTO dbversion (dbversion) VALUES (20111212.1)


<br /><br /><b>20120102.1</b><br />
alter table job CHANGE jobnr jobnr varchar(50)  NOT NULL default 0;
alter table timer CHANGE tjobnr tjobnr varchar(50)  NOT NULL default 0;
alter table timer_konsolideret CHANGE tk_jobnr tk_jobnr varchar(50)  NOT NULL default 0;
INSERT INTO dbversion (dbversion) VALUES (20120102.1)


<br /><br /><b>20120416.1</b><br />
alter table timereg_usejob ADD (aktid double NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20120416.1)

<br /><br /><b>20120417.1</b><br />
alter table job ADD (syncslutdato int NOT NULL default 0, lukkedato date NOT NULL default '2002-01-01');
INSERT INTO dbversion (dbversion) VALUES (20120417.1)

<br /><br /><b>20120424.1</b><br />
alter table job ADD (altfakadr int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20120424.1)


<br /><br /><b>20120427.1</b><br />
INSERT INTO akt_typer (aty_id,
 aty_label, aty_desc,
aty_on, aty_on_realhours, aty_on_invoiceble, aty_on_invoice, aty_on_invoice_chk 
, aty_on_workhours, aty_pre, aty_sort, aty_on_recon, aty_enh, aty_on_adhoc, aty_hide_on_treg)
VALUES
 (23, 'global_txt_172', 'Omsorgsdage',1,0,0,0,1,0,0,5.42,1,0,0,0), 
 (24, 'global_txt_173','Seniortimer',1,0,0,0,1,0,0,5.43,1,0,0,0),
 (25, 'global_txt_174','1 maj timer',1,0,0,0,1,0,0,3.9,1,0,0,0)
; INSERT INTO dbversion (dbversion) VALUES (20120427.1)



<br /><br /><b>20120516.1</b><br />
alter table fakturaer ADD (afsender int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20120516.1)


<br /><br /><b>20120611.1</b><br />
CREATE TABLE  job_igv_status  (
id  INT NOT NULL AUTO_INCREMENT,
dato date,
editor VARCHAR(100),
medid INT NOT NULL DEFAULT 0,
aar INT NOT NULL DEFAULT 0,
maaned INT NOT NULL DEFAULT 0,
PRIMARY KEY (id)
);
INSERT INTO dbversion (dbversion) VALUES (20120611.1)


<br /><br /><b>20120615.1</b><br />
alter table licens ADD (jobasnvigv int NOT NULL default 1);
INSERT INTO dbversion (dbversion) VALUES (20120615.1)


<br /><br /><b>20120625.1</b><br />
UPDATE licens SET jobasnvigv = 0 WHERE id = 1;
INSERT INTO dbversion (dbversion) VALUES (20120625.1)


<br /><br /><b>20120625.2</b><br />
CREATE TABLE  timer_import_temp  (
id  INT NOT NULL AUTO_INCREMENT,
dato date,
editor VARCHAR(100),
origin INT NOT NULL DEFAULT 0,
medarbejderid  VARCHAR(100),
jobid  DOUBLE(12,2) NOT NULL DEFAULT 0,
aktnavn  VARCHAR(100),
timer  DOUBLE(12,2) NOT NULL DEFAULT 0,
tdato  date,
lto VARCHAR(100),
PRIMARY KEY (id)
);
INSERT INTO dbversion (dbversion) VALUES (20120625.2)


<br /><br /><b>20120629.1</b><br />
alter table timer_import_temp ADD (overfort int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20120629.1)

<br /><br /><b>20120705.1</b><br />
alter table timer_import_temp ADD (timerkom text);
INSERT INTO dbversion (dbversion) VALUES (20120705.1)

<br /><br /><b>20120808.1</b><br />
ALTER table filer CHANGE filnavn filnavn varchar(250);
INSERT INTO dbversion (dbversion) VALUES (20120808.1) 

<br /><br /><b>20120914.1</b><br />
alter table job ADD (preconditions_met int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20120914.1)


<br /><br /><b>20120919.1</b><br />
alter table aktiviteter ADD (brug_fasttp int NOT NULL default 0, brug_fastkp int NOT NULL default 0, fasttp double(12,2) NOT NULL default 0, fasttp_val int NOT NULL default 0, fastkp double(12,2) NOT NULL default 0, fastkp_val int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20120919.1)


<br /><br /><b>20120920.1</b><br />
alter table fakturaer ADD (overfort_erp int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20120920.1)

<br /><br /><b>20120928.1</b><br />
alter table aktiviteter ADD (avarenr varchar(250));
INSERT INTO dbversion (dbversion) VALUES (20120928.1)

<br /><br /><b>20121022.1</b><br />
alter table fakturaer ADD (vis_jobbesk int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20121022.1)


<br /><br /><b>20121027.1</b><br />
alter table kunder ADD (regnr_b varchar(10), kontonr_b varchar(50), regnr_c varchar(10), kontonr_c varchar(50));
INSERT INTO dbversion (dbversion) VALUES (20121027.1)


<br /><br /><b>20121027.2</b><br />
ALTER table kunder CHANGE regnr regnr varchar(10),
CHANGE kontonr kontonr varchar(50);
INSERT INTO dbversion (dbversion) VALUES (20121027.2) 

<br /><br /><b>20121027.3</b><br />
alter table fakturaer ADD (kontonr_sel int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20121027.3)


<br /><br /><b>20121029.1</b><br />
alter table kunder ADD (
bank_b varchar(250), swift_b varchar(50), iban_b varchar(50), 
bank_c varchar(250), swift_c varchar(50), iban_c varchar(50)
);
INSERT INTO dbversion (dbversion) VALUES (20121029.1)




<br /><br /><b>20121016.2</b><br />
CREATE TABLE  gantts  (
id  INT NOT NULL AUTO_INCREMENT,
dato date,
editor VARCHAR(100),
medid int NOT NULL DEFAULT 0,
jobids  text,
navn varchar(255),
PRIMARY KEY (id)
);
INSERT INTO dbversion (dbversion) VALUES (20121016.2)

<br /><br /><b>20121029.2</b><br />
ALTER table gantts CHANGE jobids jobnrs text;
INSERT INTO dbversion (dbversion) VALUES (20121029.2) 


<br /><br /><b>20121029.3</b><br />
alter table gantts ADD (
datomidtpkt date NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20121029.3)

<br /><br /><b>20121104.1</b><br />
alter table licens ADD (
positiv_aktivering_akt int DEFAULT 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20121104.1)

<br /><br /><b>20121105.1</b><br />
alter table licens ADD (
showeasyreg int DEFAULT 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20121105.1)


<br /><br /><b>20121105.2</b><br />
alter table medarbejdere ADD (
bankaccount varchar (255));
INSERT INTO dbversion (dbversion) VALUES (20121105.2)


<br /><br /><b>20121109.1</b><br />
alter table licens ADD (
showupload int DEFAULT 0 NOT NULL,
forcebudget_onakttreg int DEFAULT 0 NOT NULL
);
INSERT INTO dbversion (dbversion) VALUES (20121109.1)


<br /><br /><b>20121113.1</b><br />
ALTER table fakturaer CHANGE modtageradr modtageradr text;
INSERT INTO dbversion (dbversion) VALUES (20121113.1) 


<br /><br /><b>20121114.1</b><br />
alter table medarbejdere ADD (
visskiftversion int DEFAULT 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20121114.1)

<br /><br /><b>20121115.1</b><br />
alter table kontaktpers ADD (kpean varchar(250), kptype INT NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20121115.1)


<br /><br /><b>20121211.1</b><br />
CREATE TABLE  medarbtyper_grp  (
id  INT NOT NULL AUTO_INCREMENT,
dato date,
editor VARCHAR(100),
navn  VARCHAR(100),
PRIMARY KEY (id)
);
INSERT INTO dbversion (dbversion) VALUES (20121211.1)


<br /><br /><b>20121211.2</b><br />
alter table medarbejdertyper ADD (mgruppe INT NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20121211.2)


<br /><br /><b>20121222.1</b><br />
INSERT INTO akt_typer (aty_id,
 aty_label, aty_desc,
aty_on, aty_on_realhours, aty_on_invoiceble, aty_on_invoice, aty_on_invoice_chk 
, aty_on_workhours, aty_pre, aty_sort, aty_on_recon, aty_enh, aty_on_adhoc, aty_hide_on_treg)
VALUES (111, 'global_txt_175','Ferie overført',0,0,0,0,1,0,0,3.11,1,0,0,0);
INSERT INTO akt_typer (aty_id,
 aty_label, aty_desc,
aty_on, aty_on_realhours, aty_on_invoiceble, aty_on_invoice, aty_on_invoice_chk 
, aty_on_workhours, aty_pre, aty_sort, aty_on_recon, aty_enh, aty_on_adhoc, aty_hide_on_treg)
VALUES (112, 'global_txt_176','Ferie optjent u. løn',0,0,0,0,1,0,0,3.12,1,0,0,0);
 INSERT INTO dbversion (dbversion) VALUES (20121222.1)



 <br /><br /><b>20130123.1</b><br />
 CREATE INDEX inx_fakid ON faktura_det (fakid);
INSERT INTO dbversion (dbversion) VALUES (20130123.1) 


<br /><br /><b>20130110.1</b><br />
alter table medarbtyper_grp ADD (opencalc INT NOT NULL DEFAULT 1 );
INSERT INTO dbversion (dbversion) VALUES (20130110.1)

<br /><br />20130210.1<br />
alter table kontaktpers ADD (kpcvr varchar (255) );
INSERT INTO dbversion (dbversion) VALUES (20130210.1)

<br /><br />20130131.1<br />
CREATE INDEX inx_tu_mja ON timereg_usejob (medarb, jobid, aktid);
INSERT INTO dbversion (dbversion) VALUES (20130131.1) 

<br /><br />20130221.1<br />
CREATE INDEX inx_timer_tregsum ON timer (Tdato, Tmnr, Tknr);
INSERT INTO dbversion (dbversion) VALUES (20130221.1) 

<br /><br />20130221.2<br />
CREATE INDEX inx_timer_tregsum2 ON timer (Tdato, Tmnr, tjobnr);
INSERT INTO dbversion (dbversion) VALUES (20130221.2) 



<br /><br />20130228.1<br />
alter table licens ADD (globalfaktor DOUBLE(12,2) NOT NULL DEFAULT 1 );
INSERT INTO dbversion (dbversion) VALUES (20130228.1)

<br/><br />20130305.1<br />
INSERT INTO akt_typer (aty_id,
 aty_label, aty_desc,
aty_on, aty_on_realhours, aty_on_invoiceble, aty_on_invoice, aty_on_invoice_chk 
, aty_on_workhours, aty_pre, aty_sort, aty_on_recon, aty_enh, aty_on_adhoc, aty_hide_on_treg)
VALUES (113, 'global_txt_177', 'Korrektion til Komme / Gå',1,0,0,0,0,2,0,2.5,0,0,0,0);

INSERT INTO akt_typer (aty_id,
 aty_label, aty_desc,
aty_on, aty_on_realhours, aty_on_invoiceble, aty_on_invoice, aty_on_invoice_chk 
, aty_on_workhours, aty_pre, aty_sort, aty_on_recon, aty_enh, aty_on_adhoc, aty_hide_on_treg)
VALUES (114, 'global_txt_178', 'Korrektion til Realiseret',1,1,0,0,0,0,0,2.6,0,0,0,0);

INSERT INTO dbversion (dbversion) VALUES (20130305.1)


<br /><br />20130313.1<br />
UPDATE akt_typer SET aty_on_realhours = 0 WHERE aty_id =114;
INSERT INTO dbversion (dbversion) VALUES (20130313.1)


<br /><br />20130313.2<br />
CREATE INDEX inx_ti_tfak_tmnr_tdato ON timer (tmnr, tdato, tfaktim);
INSERT INTO dbversion (dbversion) VALUES (20130313.2) 

<br /><br />20130313.3<br />
DROP INDEX inx_ti_tfak_tmnr_tdato ON timer;

<br /><br />20130412.1<br />
alter table licens ADD (bdgmtypon int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20130412.1)

<br /><br />20130419.1<br />
alter table kunder ADD (regnr_d varchar(10), kontonr_d varchar(50),
bank_d varchar(250), swift_d varchar(50), iban_d varchar(50));
INSERT INTO dbversion (dbversion) VALUES (20130419.1)

<br /><br />20130422.1<br />
alter table kunder ADD (regnr_e varchar(10), kontonr_e varchar(50),
bank_e varchar(250), swift_e varchar(50), iban_e varchar(50));
INSERT INTO dbversion (dbversion) VALUES (20130422.1)

<br /><br />20130514.1<br />
alter table medarbejdere ADD  (medarbejdertype_grp int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20130514.1)

<br /><br />20130515.1<br />
UPDATE medarbejdere SET medarbejdertype_grp = 1 WHERE medarbejdertype_grp = 0 ;
INSERT INTO dbversion (dbversion) VALUES (20130515.1)


<br /><br />20130515.2<br />
UPDATE medarbejdertyper SET mgruppe = 1 WHERE mgruppe = 0 ;
INSERT INTO dbversion (dbversion) VALUES (20130515.2)

<br /><br />20130515.3<br />
INSERT INTO medarbtyper_grp SET dato = '2013-05-15', editor = 'TimeOut Support', navn = 'Medarbejdertype Hovedgruppe 1', opencalc = 1;
INSERT INTO dbversion (dbversion) VALUES (20130515.3)


<br /><br />20130523.1<br />
CREATE UNIQUE INDEX inx_mids ON medarbejdere (mid);
INSERT INTO dbversion (dbversion) VALUES (20130523.1);
CREATE INDEX inx_mansat ON medarbejdere (mansat);
INSERT INTO dbversion (dbversion) VALUES (20130523.2);
CREATE INDEX inx_minit ON medarbejdere (init);
INSERT INTO dbversion (dbversion) VALUES (20130523.3)

<br /><br />20130529.1<br />
CREATE INDEX inx2_mtb_jobid_typeid ON medarbejdertyper_timebudget  (jobid, typeid);
INSERT INTO dbversion (dbversion) VALUES (20130529.1) 


<br /><br />20130607.1<br />
CREATE INDEX inx_tim_aktid ON timer (taktivitetid);
INSERT INTO dbversion (dbversion) VALUES (20130607.1) 


<br /><br />20130607.2<br />
CREATE TABLE  timer_konsolideret_md  (
id  INT NOT NULL AUTO_INCREMENT,
dato date,
jobid int NOT NULL DEFAULT 0,
mtype int NOT NULL DEFAULT 0,
mtgid int NOT NULL DEFAULT 0,
timer double (12,2) NOT NULL DEFAULT 0,
belob double (12,2) NOT NULL DEFAULT 0,
kost double (12,2) NOT NULL DEFAULT 0,
PRIMARY KEY (id)
);
INSERT INTO dbversion (dbversion) VALUES (20130607.2)


<br /><br />20130607.3<br />
RENAME TABLE timer_konsolideret_md TO timer_konsolideret_tot;
CREATE INDEX inx_mj ON timer_konsolideret_tot (mtgid, jobid);
INSERT INTO dbversion (dbversion) VALUES (20130607.3)


<br /><br />20130607.4<br />
CREATE INDEX inx_tojob ON timer_konsolideret_tot (jobid);
INSERT INTO dbversion (dbversion) VALUES (20130607.4)

<br /><br />20130618.1<br />
alter table timer_konsolideret_tot ADD  (opd_dato timestamp);
INSERT INTO dbversion (dbversion) VALUES (20130618.1)


<br /><br />20130228.1<br />
alter table fakturaer ADD (fakglobalfaktor DOUBLE(12,2) NOT NULL DEFAULT 1 );
INSERT INTO dbversion (dbversion) VALUES (20130228.1)




<br /><br />20131008.1<br />
CREATE INDEX inx_lh_mid ON login_historik (mid);
INSERT INTO dbversion (dbversion) VALUES (20131008.1) 


<br /><br />20131008.1<br />
CREATE INDEX inx_lh_mid2 ON login_historik (mid);
INSERT INTO dbversion (dbversion) VALUES (20131008.1) 

<br /><br />20131008.2<br />
CREATE INDEX inx_lh_mid_dt2 ON login_historik (mid, dato);
INSERT INTO dbversion (dbversion) VALUES (20131008.2) 


<br /><br />20131105.1<br />
alter table licens ADD  (regnskabsaar_start date NOT NULL default '2001-01-01');
INSERT INTO dbversion (dbversion) VALUES (20131105.1)


<br /><br />20131105.2<br />
alter table licens ADD  (forcebudget_onakttreg_afgr Int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20131105.2)


<br /><br />20131107.1<br />
UPDATE akt_typer SET aty_desc ='Barsel/Orlov' WHERE aty_id = 22;
INSERT INTO dbversion (dbversion) VALUES (20131107.1)


<br /><br />20131108.1<br />
INSERT INTO akt_typer (aty_id,
 aty_label, aty_desc,
aty_on, aty_on_realhours, aty_on_invoiceble, aty_on_invoice, aty_on_invoice_chk 
, aty_on_workhours, aty_pre, aty_sort, aty_on_recon, aty_enh, aty_on_adhoc, aty_hide_on_treg)
VALUES (115, 'global_txt_179','Tjenestefri',0,0,0,0,1,0,0,5.44,1,0,0,0);
INSERT INTO dbversion (dbversion) VALUES (20131108.1)


<br /><br />20131126.1<br />
ALTER table kunder CHANGE postnr postnr varchar(25);
INSERT INTO dbversion (dbversion) VALUES (20131126.1) 

<br /><br />20131126.2<br />
ALTER table kunder CHANGE cvr cvr varchar(25);
INSERT INTO dbversion (dbversion) VALUES (20131126.2) 


<br /><br />20131211.1<br />
alter table job ADD (laasmedtpbudget int default 0);
INSERT INTO dbversion (dbversion) VALUES (20131211.1)

<br /><br />20131220.1<br />
UPDATE kunder SET land = 'Danmark' WHERE land = 'DK';
INSERT INTO dbversion (dbversion) VALUES (20131220.1)

<br /><br />20140110.1<br />
alter table kunder ADD (kinit varchar(50));
INSERT INTO dbversion (dbversion) VALUES (20140110.1)


<br /><br />20140120.1<br />
alter table timer_import_temp ADD (jobnavn varchar(250) NULL, aktid int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20140120.1)


<br /><br />20140120.2<br />
ALTER table timer_import_temp CHANGE jobid jobid double(12,0) NOT NULL default 0; 
ALTER table timer_import_temp CHANGE aktid aktid double(12,0) NOT NULL default 0;
INSERT INTO dbversion (dbversion) VALUES (20140120.2)


<br /><br />20140120.3<br />
alter table timer_import_temp ADD (errid int NOT NULL default 0, errmsg varchar(50) NULL);
INSERT INTO dbversion (dbversion) VALUES (20140120.3)


<br /><br />20140122.1<br />
alter table ressourcer_md  ADD (proc double(12,2) NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20140122.1)

<br /><br />20140209.1<br />
alter table timer ADD (timeralt double(12,2) NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20140209.1)

<br /><br />20140211.1<br />
alter table licens  ADD (lukaktvdato int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20140211.1)


<br /><br />20140211.2<br />
alter table job  ADD (salgsans1 int NOT NULL default 0, salgsans2 int NOT NULL default 0,
salgsans3 int NOT NULL default 0, salgsans4 int NOT NULL default 0, 
salgsans5 int NOT NULL default 0,
salgsans1_proc double(12,2) NOT NULL default 0,
salgsans2_proc double(12,2) NOT NULL default 0,
salgsans3_proc double(12,2) NOT NULL default 0,
salgsans4_proc double(12,2) NOT NULL default 0,
salgsans5_proc double(12,2) NOT NULL default 0
);
INSERT INTO dbversion (dbversion) VALUES (20140211.2)


<br /><br />20140211.3<br />
alter table licens ADD (salgsans int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20140211.3)


-- Er ikke Kørt: ----

<br /><br />20140212.1<br />
CREATE TABLE  wip_historik  (
id  INT NOT NULL AUTO_INCREMENT,
dato date,
medid  INT NOT NULL DEFAULT 0,
jobid  DOUBLE(12,2) NOT NULL DEFAULT 0,
restestimat DOUBLE(12,2) NOT NULL DEFAULT 0,
stade_tim_proc  INT NOT NULL DEFAULT 0,
PRIMARY KEY (id)
);
INSERT INTO dbversion (dbversion) VALUES (20140212.1);

<br /><br />20140218.1<br />
CREATE INDEX inx_a_fakturerbar ON aktiviteter (fakturerbar);
INSERT INTO dbversion (dbversion) VALUES (20140218.1) 


<br /><br />20140226.1<br />
alter table licens ADD (smileyaggressiv int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20140226.1)


<br /><br />20140314.1<br />
alter table medarbejdertyper ADD (afslutugekri int NOT NULL default 0,afslutugekri_proc double(12,2) NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20140314.1)


<br /><br />20140316.1<br />
alter table licens ADD (timerround int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20140316.1)

<br /><br />20140317.1<br />
CREATE INDEX inx_ugest ON ugestatus (mid);
INSERT INTO dbversion (dbversion) VALUES (20140317.1)

<br /><br />20140317.2<br />
CREATE INDEX inx_ugest_mid_uge ON ugestatus (mid, uge);
INSERT INTO dbversion (dbversion) VALUES (20140317.2)  

<br /><br />20140322.1<br />
alter table materiale_forbrug ADD (aktid int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20140322.1)

<br /><br />20140409.1<br />
alter table fakturaer ADD (totbel_afvige int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20140409.1)



<br /><br />20140521.1-6<br />
CREATE INDEX inx_fi_j3 ON filer (jobid);
INSERT INTO dbversion (dbversion) VALUES (20140521.1);
CREATE INDEX inx_fi_k2 ON filer (kundeid);
INSERT INTO dbversion (dbversion) VALUES (20140521.1);
CREATE INDEX inx_fo_k2 ON foldere (kundeid);
INSERT INTO dbversion (dbversion) VALUES (20140521.3);
CREATE INDEX inx_fo_j2 ON foldere (jobid);
INSERT INTO dbversion (dbversion) VALUES (20140521.4);
CREATE INDEX inx_fi_n2 ON filer (filnavn);
INSERT INTO dbversion (dbversion) VALUES (20140521.5);
CREATE INDEX inx_fo_n2 ON foldere (navn);
INSERT INTO dbversion (dbversion) VALUES (20140521.6);


<br /><br />20140527.1<br />
alter table akt_typer ADD (aty_pre_dg varchar(50) NOT NULL default '',  aty_pre_prg varchar(50) NOT NULL default '10' );
INSERT INTO dbversion (dbversion) VALUES (20140527.1)

<br /><br />20140528.2<br />
CREATE INDEX inx_progrprel_jm ON progrupperelationer (projektgruppeid, medarbejderid);
INSERT INTO dbversion (dbversion) VALUES (20140528.2) 

<br /><br />20140610.1<br />
UPDATE licens SET timeout_version = '2_14' WHERE id = 1



<br /><br />20140617.1-14 NT<br />
alter table job  ADD (department varchar(50));
INSERT INTO dbversion (dbversion) VALUES (20140617.1);
alter table job  ADD (dt_proto_dead date, dt_proto_sent date, dt_sms_dead date, dt_sms_sent date, dt_photo_dead date, dt_photo_sent date, dt_exp_order date);
INSERT INTO dbversion (dbversion) VALUES (20140617.10);
alter table job  ADD ( dt_confb_etd date, dt_confb_eta date, dt_confs_etd date, dt_confs_eta date, dt_actual_etd date, dt_actual_eta date);
INSERT INTO dbversion (dbversion) VALUES (20140617.12);
alter table job  ADD ( dt_firstorderc date, dt_ldapp date, dt_sizeexp date, dt_sizeapp date, dt_ppexp date, dt_ppapp date, dt_shsexp date, dt_shsapp date);
INSERT INTO dbversion (dbversion) VALUES (20140617.13);
alter table job  ADD (destination varchar(50), transport varchar(50), orderqty double(12,0) DEFAULT 0 NOT NULL, shippedqty double(12,0) DEFAULT 0 NOT NULL, supplier_invoiceno varchar(50));
INSERT INTO dbversion (dbversion) VALUES (20140617.13);
alter table job  ADD (origin varchar(50));
INSERT INTO dbversion (dbversion) VALUES (20140617.2);
alter table job  ADD (supplier int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20140617.3);
alter table job  ADD (composition varchar(200) );
INSERT INTO dbversion (dbversion) VALUES (20140617.5);
alter table job  ADD (product_group int DEFAULT 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20140617.7);
alter table job  ADD (scsq_note text, sample_note text);
INSERT INTO dbversion (dbversion) VALUES (20140617.8);
alter table job  ADD (dt_enq_st date, dt_enq_end date);
INSERT INTO dbversion (dbversion) VALUES (20140617.9);
alter table job  ADD (dt_sour_dead date);
INSERT INTO dbversion (dbversion) VALUES (20140617.18);
alter table job  ADD (collection  varchar(50));
INSERT INTO dbversion (dbversion) VALUES (20140617.19)


alter table job  ADD (dt_sup_photo_dead date, dt_sup_sms_dead date);
INSERT INTO dbversion (dbversion) VALUES (20140617.20);
alter table job  ADD (tax_pc double (12,2) NOT NULL DEFAULT 0, comm_pc double (12,2) NOT NULL DEFAULT 0, freight_pc double (12,2) NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20140617.21);
alter table job  ADD (cost_price_pc double (12,2) NOT NULL DEFAULT 0, sales_price_pc double (12,2) NOT NULL DEFAULT 0, tgt_price_pc double (12,2) NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20140617.22);
alter table job  ADD (sales_price_pc_valuta int NOT NULL DEFAULT 1, cost_price_pc_valuta int NOT NULL DEFAULT 1);
INSERT INTO dbversion (dbversion) VALUES (20140617.23);
alter table job  ADD (tgt_price_pc_valuta int NOT NULL DEFAULT 1);
INSERT INTO dbversion (dbversion) VALUES (20140617.24);
alter table job  ADD (cost_price_pc_base  Double (12,2) NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20140617.25)




<br /><br />20140625.1<br />
alter table licens ADD (teamleder_flad INT default 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20140625.1)

<br /><br />20140708.1<br />
alter table licens ADD (forcebudget_onakttreg_filt_viskunmbgt INT default 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20140708.1)


<br /><br />20140709.1<br />
alter table licens ADD (stempelur_hidelogin INT default 0 NOT NULL, stempelur_igno_komkrav INT default 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20140709.1)


<br /><br />20140627.27<br />
alter table job ADD (kunde_betbetint int NOT NULL DEFAULT 0, kunde_levbetint int NOT NULL DEFAULT 0, lev_betbetint int NOT NULL DEFAULT 0, lev_levbetint int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20140627.27)

<br /><br />20140821.1<br />
alter table licens ADD (SmiWeekOrMonth INT default 0 NOT NULL, SmiantaldageCount INT default 1 NOT NULL, SmiantaldageCountClock INT default 12 NOT NULL, SmiTeamlederCount INT default 1 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20140821.1)

<br /><br />20140901.1<br />
alter table licens ADD (hidesmileyicon INT default 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20140901.1)

<br /><br />20140617.3<br />
alter table job  ADD (supplier int NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20140617.3)


<br /><br />20140617.35<br />
alter table job  ADD (invoice_no varchar(50) NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20140617.35)


<br /><br />20141106.1<br />
alter table progrupperelationer ADD (notificer INT default 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20141106.1)


<br /><br />20141111.1<br />
UPDATE fak_sprog SET navn = 'Danish / DK' WHERE id = 1;
UPDATE fak_sprog SET navn = 'English / UK' WHERE id = 2;
UPDATE fak_sprog SET navn = 'Swedish / SE' WHERE id = 3;
INSERT INTO dbversion (dbversion) VALUES (20141111.1)

<br /><br />20141111.2<br />
alter table materiale_forbrug ADD (ava double(12,2) default 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20141111.2)

<br /><br />20141211.1<br />
alter table materiale_grp ADD (mgkundeid int default 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20141211.1)

<br /><br />20141216.1<br />
alter table materiale_forbrug MODIFY ava double(12,4) default 0 NOT NULL;
INSERT INTO dbversion (dbversion) VALUES (20141216.1)


<br /><br />20141216.2<br />
INSERT INTO akt_typer (aty_id,
 aty_label, aty_desc,
aty_on, aty_on_realhours, aty_on_invoiceble, aty_on_invoice, aty_on_invoice_chk 
, aty_on_workhours, aty_pre, aty_sort, aty_on_recon, aty_enh, aty_on_adhoc, aty_hide_on_treg)
VALUES (27, 'global_txt_185','Aldersreduktion optjent',0,0,0,0,0,0,0,5.27,1,0,0,0);
INSERT INTO akt_typer (aty_id,
 aty_label, aty_desc,
aty_on, aty_on_realhours, aty_on_invoiceble, aty_on_invoice, aty_on_invoice_chk 
, aty_on_workhours, aty_pre, aty_sort, aty_on_recon, aty_enh, aty_on_adhoc, aty_hide_on_treg)
VALUES (28, 'global_txt_186','Aldersreduktion brugt',0,0,0,0,0,0,0,5.28,1,0,0,0);
INSERT INTO akt_typer (aty_id,
 aty_label, aty_desc,
aty_on, aty_on_realhours, aty_on_invoiceble, aty_on_invoice, aty_on_invoice_chk 
, aty_on_workhours, aty_pre, aty_sort, aty_on_recon, aty_enh, aty_on_adhoc, aty_hide_on_treg)
VALUES (29, 'global_txt_187','Aldersreduktion udbetalt',0,0,0,0,0,0,0,5.29,1,0,0,0);
 INSERT INTO dbversion (dbversion) VALUES (20141216.2)

<br /><br />20141219.1<br />
alter table medarbejdertyper ADD (noflex int default 0 NOT NULL);
alter table medarbejdere ADD (showtimereg_start_stop int default 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20141219.1)


<br /><br />20141222.1<br />
alter table licens ADD (visAktlinjerSimpel int default 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20141222.1)

<br /><br />20141222.2<br />
alter table medarbejdere ADD (timer_ststop int default 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20141222.2)


<br /><br />20150108.1<br />
INSERT INTO akt_typer (aty_id,
 aty_label, aty_desc,
aty_on, aty_on_realhours, aty_on_invoiceble, aty_on_invoice, aty_on_invoice_chk 
, aty_on_workhours, aty_pre, aty_sort, aty_on_recon, aty_enh, aty_on_adhoc, aty_hide_on_treg)
VALUES (26, 'global_txt_188','Aldersreduktion planlagt',0,0,0,0,0,0,0,5.30,1,0,0,0);
INSERT INTO akt_typer (aty_id,
 aty_label, aty_desc,
aty_on, aty_on_realhours, aty_on_invoiceble, aty_on_invoice, aty_on_invoice_chk 
, aty_on_workhours, aty_pre, aty_sort, aty_on_recon, aty_enh, aty_on_adhoc, aty_hide_on_treg)
VALUES (120, 'global_txt_189','Omsorgsdag 2 planlagt',0,0,0,0,0,0,0,5.31,1,0,0,0);
INSERT INTO akt_typer (aty_id,
 aty_label, aty_desc,
aty_on, aty_on_realhours, aty_on_invoiceble, aty_on_invoice, aty_on_invoice_chk 
, aty_on_workhours, aty_pre, aty_sort, aty_on_recon, aty_enh, aty_on_adhoc, aty_hide_on_treg)
VALUES (121, 'global_txt_190','Omsorgsdag 10 planlagt',0,0,0,0,0,0,0,5.32,1,0,0,0);
INSERT INTO akt_typer (aty_id,
 aty_label, aty_desc,
aty_on, aty_on_realhours, aty_on_invoiceble, aty_on_invoice, aty_on_invoice_chk 
, aty_on_workhours, aty_pre, aty_sort, aty_on_recon, aty_enh, aty_on_adhoc, aty_hide_on_treg)
VALUES (122, 'global_txt_191','Omsorgsdag Koverteret planlagt',0,0,0,0,0,0,0,5.33,1,0,0,0);
 INSERT INTO dbversion (dbversion) VALUES (20141216.2)


<br /><br />20150115.1<br />
alter table akt_typer ADD (aty_on_calender int default 1 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20150115.1)

<br /><br />20150127.1<br />
alter table job  ADD (filepath1 varchar(255), filepath2 varchar(255), filepath3 varchar(255));
INSERT INTO dbversion (dbversion) VALUES (20150127.1)

<br /><br />20150203.1<br />
INSERT INTO akt_typer (aty_id,
 aty_label, aty_desc,
aty_on, aty_on_realhours, aty_on_invoiceble, aty_on_invoice, aty_on_invoice_chk 
, aty_on_workhours, aty_pre, aty_sort, aty_on_recon, aty_enh, aty_on_adhoc, aty_hide_on_treg)
VALUES (123, 'global_txt_192','Ulempe1706 udbetalt',0,0,0,0,0,0,0,5.30,1,0,0,0);
INSERT INTO akt_typer (aty_id,
 aty_label, aty_desc,
aty_on, aty_on_realhours, aty_on_invoiceble, aty_on_invoice, aty_on_invoice_chk 
, aty_on_workhours, aty_pre, aty_sort, aty_on_recon, aty_enh, aty_on_adhoc, aty_hide_on_treg)
VALUES (124, 'global_txt_193','UlempeWeekend udbetalt',0,0,0,0,0,0,0,5.31,1,0,0,0);
INSERT INTO dbversion (dbversion) VALUES (20150203.1)


<br /><br />20150127.1<br />
alter table fak_mat_spec ADD (fms_aktid INT NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20150127.1)

<br /><br />20150305.1<br />
alter table kunder ADD (regnr_f VARCHAR(50), kontonr_f VARCHAR(50), bank_f VARCHAR(50), swift_f VARCHAR(50), iban_f VARCHAR(50));
INSERT INTO dbversion (dbversion) VALUES (20150305.1)

<br /><br />20150312.1<br />
alter table fakturaer ADD (fak_fomr int DEFAULT 0 NOT NULL);
alter table fomr ADD (konto int DEFAULT 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20150310.1);

alter table fomr ADD (jobok int DEFAULT 1 NOT NULL, aktok int DEFAULT 1 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20150310.2);

alter table job ADD (fomr_konto int DEFAULT 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20150311.1);

alter table licens ADD (fomr_mandatory int DEFAULT 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20150312.1)


<br /><br />20150317.1-4<br />
alter table kunder ADD (kfak_sprog int DEFAULT 1 NOT NULL, kfak_moms int DEFAULT 1 NOT NULL);
alter table job ADD (jfak_sprog int DEFAULT 1 NOT NULL, jfak_moms int DEFAULT 1 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20150317.1);

CREATE TABLE fak_moms  (
id  INT NOT NULL AUTO_INCREMENT,
moms int NOT NULL DEFAULT 0,
PRIMARY KEY (id)
);
INSERT INTO dbversion (dbversion) VALUES (20150317.2);


INSERT INTO fak_moms (moms) VALUES (25);
INSERT INTO fak_moms (moms) VALUES (8);
INSERT INTO fak_moms (moms) VALUES (14);
INSERT INTO fak_moms (moms) VALUES (20);
INSERT INTO fak_moms (moms) VALUES (22);
INSERT INTO fak_moms (moms) VALUES (0);
INSERT INTO dbversion (dbversion) VALUES (20150317.3);

alter table kunder ADD (kfak_valuta int DEFAULT 1 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20150317.4)


<br /><br />20150324.1<br />
alter table medarbejdertyper ADD (kostpristarif_A double(12,2) DEFAULT 0 NOT NULL, 
kostpristarif_B double(12,2) DEFAULT 0 NOT NULL,
kostpristarif_C double(12,2) DEFAULT 0 NOT NULL,
kostpristarif_D double(12,2) DEFAULT 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20150324.1)


<br /><br />20150324.2<br />
alter table aktiviteter ADD (kostpristarif varchar(50) DEFAULT 0 NOT NULL); 
INSERT INTO dbversion (dbversion) VALUES (20150324.2)

<br /><br />20150407.1<br />
CREATE INDEX inx_progrp ON projektgrupper (id);
INSERT INTO dbversion (dbversion) VALUES (20150407.1) 

<br /><br />20150416.1<br />
alter table ressourcer_ramme ADD (aktid bigint DEFAULT 0 NOT NULL); 
INSERT INTO dbversion (dbversion) VALUES (20150416.1) 

<br /><br />20150417.1<br />
alter table ressourcer_ramme ADD (fctimepris double(12,2) DEFAULT 0 NOT NULL); 
INSERT INTO dbversion (dbversion) VALUES (20150417.1) 

<br /><br />20150421.1<br />
alter table ressourcer_ramme ADD (fctimeprish2 double(12,2) DEFAULT 0 NOT NULL); 
INSERT INTO dbversion (dbversion) VALUES (20150421.1) 

<br /><br />20150504.1<br />
alter table fomr ADD (business_unit VARCHAR(250), business_area_label VARCHAR(250)); 
INSERT INTO dbversion (dbversion) VALUES (20150504.1) 

<br /><br />20150529.1<br />
alter table job  ADD (alert int DEFAULT 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20150529.1)


<br /><br />20150513.1<br />
CREATE TABLE  budget_job (
id INT NOT NULL AUTO_INCREMENT,
dato date,
editor VARCHAR(100),
budgetnavn  VARCHAR(100),
jobid INTEGER NOT NULL,
refno VARCHAR(100),
prodno VARCHAR(100),
budget_tot double(12,2),
budget_extra_fo double(12,2),
valuta INTEGER,
kurs double(12,2),
budget_view INTEGER NOT NULL,
rapport_view INTEGER NOT NULL,
stdato date,
sldato date, 
budgetstatus INTEGER,
PRIMARY KEY (id)
);
INSERT INTO dbversion (dbversion) VALUES (20150513.1)

<br /><br />20150515.2<br />
CREATE TABLE  budget_job_per (
id INT NOT NULL AUTO_INCREMENT,
budgetid INTEGER,
beskrivelse varchar(250),
budget_app double(12,2),
valuta INTEGER,
kurs double,
stdato date,
sldato date,
currentperiod INTEGER NOT NULL default 0,  
PRIMARY KEY (id)
);
INSERT INTO dbversion (dbversion) VALUES (20150515.2)


<br /><br />20150515.3<br />
CREATE TABLE  budget_job_exp (
id INT NOT NULL AUTO_INCREMENT,
budgetid INTEGER,
periodeid INTEGER,
fase VARCHAR(250),
aktid INTEGER,
aktnavn VARCHAR(250),
konto VARCHAR(250),
budget_app double(12,2) NOT NULL default 0,
valuta INTEGER,
kurs double(12,2),
perdato date,
periode_no INTEGER,
PRIMARY KEY (id)
);
INSERT INTO dbversion (dbversion) VALUES (20150515.3)


<br /><br />20150602.1<br />
CREATE TABLE budget_job_exp_rel (
id INT NOT NULL AUTO_INCREMENT,
budgetid INTEGER,
periodeid INTEGER,
aktid INTEGER,
belob double(12,2),
dato date,
PRIMARY KEY (id)
);
INSERT INTO dbversion (dbversion) VALUES (20150602.1)

<br /><br />20150604.1<br />
alter table akt_gruppe ADD (skabelontype int NOT NULL default 0); 
INSERT INTO dbversion (dbversion) VALUES (20150604.1) 

<br /><br />20150612.1<br />
alter table licens ADD (matlageruminemailnot Date DEFAULT '2002-01-01' NOT NULL); 
INSERT INTO dbversion (dbversion) VALUES (20150612.1) 


<br /><br />20150612.2<br />
alter table akt_gruppe ADD (aktgrp_account int NOT NULL default 0); 
INSERT INTO dbversion (dbversion) VALUES (20150612.2) 


<br /><br />20150623.1<br />
alter table licens ADD (akt_maksbudget_treg int NOT NULL default 0); 
INSERT INTO dbversion (dbversion) VALUES (20150623.1) 

<br /><br />20150623.2<br />
alter table materialer ADD (fabrikat VARCHAR(255)); 
INSERT INTO dbversion (dbversion) VALUES (20150623.2) 

<br /><br />20150701.1<br />
alter table materiale_grp ADD (medarbansv int DEFAULT 0 NOT NULL); 
INSERT INTO dbversion (dbversion) VALUES (20150701.1) 


<br /><br />20150701.2<br />
alter table licens ADD (minimumslageremail int DEFAULT 0 NOT NULL); 
INSERT INTO dbversion (dbversion) VALUES (20150701.2) 

<br /><br />20150811.1<br />
alter table budget_job ADD (b_valuta int DEFAULT 0 NOT NULL, b_valuta_kurs double(12,2) DEFAULT 0 NOT NULL); 
INSERT INTO dbversion (dbversion) VALUES (20150811.1) 

<br /><br />20150819.1<br />
ALTER TABLE fomr ADD (fomr_segment VARCHAR(255));
INSERT INTO dbversion (dbversion) VALUES (20150819.1) 

<br /><br />20150917.1<br />
ALTER TABLE licens ADD (fomr_account INT DEFAULT 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20150917.1)

<br /><br />20150917.2<br />
ALTER TABLE aktiviteter ADD (aktkonto INT DEFAULT 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20150917.2)


<br /><br />20150921.1<br />
ALTER TABLE licens ADD (visAktlinjerSimpel_datoer INT DEFAULT 1 NOT NULL);
ALTER TABLE licens ADD (visAktlinjerSimpel_timebudget INT DEFAULT 1 NOT NULL);
ALTER TABLE licens ADD (visAktlinjerSimpel_realtimer INT DEFAULT 1 NOT NULL);
ALTER TABLE licens ADD (visAktlinjerSimpel_restimer INT DEFAULT 1 NOT NULL);
ALTER TABLE licens ADD (visAktlinjerSimpel_medarbtimepriser INT DEFAULT 1 NOT NULL);
ALTER TABLE licens ADD (visAktlinjerSimpel_medarbrealtimer INT DEFAULT 1 NOT NULL);
ALTER TABLE licens ADD (visAktlinjerSimpel_akttype INT DEFAULT 1 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20150921.1) 

<br /><br />20150921.2<br />
UPDATE licens SET visAktlinjerSimpel_restimer = 0 WHERE id = 1


<br /><br />20151006.1<br />
ALTER TABLE licens ADD (timesimon INT DEFAULT 0 NOT NULL);
ALTER TABLE licens ADD (timesimh1h2 INT DEFAULT 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20151006.1) 

<br /><br />20151027.1<br />
ALTER TABLE projektgrupper ADD (orgvir INT DEFAULT 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20151027.1) 

<br /><br />20151104.1<br />
ALTER TABLE licens ADD (timesimtp INT DEFAULT 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20151104.1)

<br /><br />20151209.1<br />
ALTER TABLE kundetyper ADD (ktlabel VARCHAR (50));
INSERT INTO dbversion (dbversion) VALUES (20151209.1)  

<br /><br />20151215.1<br />
ALTER TABLE licens ADD (budgetakt INT DEFAULT 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20151215.1) 


<br /><br />20151215.2<br />
CREATE TABLE traveldietexp (
diet_id INT NOT NULL AUTO_INCREMENT,
diet_mid INT NOT NULL DEFAULT 0,
diet_namedest VARCHAR(250) NOT NULL,
diet_stdato datetime,
diet_sldato datetime,
diet_jobid INT NOT NULL DEFAULT 0,
diet_konto INT NOT NULL DEFAULT 0,
diet_dayprice INT NOT NULL DEFAULT 0,
diet_maksamount DOUBLE(12,2) NOT NULL DEFAULT 0,
diet_morgen INT NOT NULL DEFAULT 0,
diet_morgenamount DOUBLE(12,2) NOT NULL DEFAULT 0,
diet_middag INT NOT NULL DEFAULT 0,
diet_middagamount DOUBLE(12,2) NOT NULL DEFAULT 0,
diet_aften INT NOT NULL DEFAULT 0,
diet_aftenamount DOUBLE(12,2) NOT NULL DEFAULT 0,
diet_rest DOUBLE(12,2) NOT NULL DEFAULT 0,
diet_min25proc DOUBLE(12,2) NOT NULL DEFAULT 0,
diet_total DOUBLE(12,2) NOT NULL DEFAULT 0,
diet_bilag INT NOT NULL DEFAULT 0,
diet_traveldays INT NOT NULL DEFAULT 0,
PRIMARY KEY (diet_id)
);
INSERT INTO dbversion (dbversion) VALUES (20151215.2)

<br /><br />20151215.3<br />
ALTER TABLE traveldietexp CHANGE diet_traveldays diet_traveldays DOUBLE(12,2);
INSERT INTO dbversion (dbversion) VALUES (20151215.3)

<br /><br />20151217.1<br />
ALTER TABLE materiale_grp ADD (konto INT DEFAULT 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20151217.1) 

<br /><br />20151217.2<br />
ALTER TABLE materiale_grp CHANGE konto matgrp_konto INT DEFAULT 0 NOT NULL;
INSERT INTO dbversion (dbversion) VALUES (20151217.2)

<br /><br />20160102.1<br />
ALTER TABLE licens ADD (akt_maksforecast_treg INT DEFAULT 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20160102.1) 

<br /><br />20160104.1<br />
ALTER TABLE licens ADD (traveldietexp_on INT DEFAULT 0 NOT NULL, traveldietexp_maxhours DOUBLE(12,2) DEFAULT 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20160104.1) 

<br /><br />20160107.1<br />
CREATE TABLE mtype_mal_fordel_fomr (
mmff_id INT NOT NULL AUTO_INCREMENT,
mmff_mal DOUBLE(12,2) NOT NULL DEFAULT 0,
mmff_mtype INT NOT NULL DEFAULT 0,
mmff_fomr INT NOT NULL DEFAULT 0,
PRIMARY KEY (mmff_id)
);
INSERT INTO dbversion (dbversion) VALUES (20160107.1)


<br /><br />20160113.1<br />
ALTER TABLE aktiviteter ADD (extsysid DOUBLE(12,0) NOT NULL default 0);
ALTER TABLE job_ulev_ju ADD (extsysid DOUBLE(12,0) NOT NULL default 0);
INSERT INTO dbversion (dbversion) VALUES (20160113.1) 

<br /><br />20160113.2<br />
ALTER TABLE licens ADD (medarbtypligmedarb INT DEFAULT 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20160113.2) 

<br /><br />20160114.3<br />
ALTER TABLE job_ulev_ju ADD (ju_konto INT DEFAULT 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20160114.3)


<br /><br />20160115.1<br />
ALTER TABLE materiale_forbrug ADD (mf_konto INT DEFAULT 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20160115.1)  


<br /><br />20160120.1<br />
CREATE TABLE traveldietexp_jobrel (
tj_id INT NOT NULL AUTO_INCREMENT,
tj_trvlid INT NOT NULL DEFAULT 0,
tj_jobid INT NOT NULL DEFAULT 0,
tj_percent DOUBLE(12,2) NOT NULL DEFAULT 0,
PRIMARY KEY (tj_id)
);
INSERT INTO dbversion (dbversion) VALUES (20160120.1)


<br /><br />20160205.1<br />
INSERT INTO akt_typer (aty_id,
 aty_label, aty_desc,
aty_on, aty_on_realhours, aty_on_invoiceble, aty_on_invoice, aty_on_invoice_chk 
, aty_on_workhours, aty_pre, aty_sort, aty_on_recon, aty_enh, aty_on_adhoc, aty_hide_on_treg)
VALUES (125, 'global_txt_194','Rejse',0,0,0,0,0,0,0,3.13,0,0,0,0);

INSERT INTO akt_typer (aty_id,
 aty_label, aty_desc,
aty_on, aty_on_realhours, aty_on_invoiceble, aty_on_invoice, aty_on_invoice_chk 
, aty_on_workhours, aty_pre, aty_sort, aty_on_recon, aty_enh, aty_on_adhoc, aty_hide_on_treg)
VALUES (92, 'global_txt_195','E3',0,0,0,0,0,0,0,1.92,0,0,0,0);
INSERT INTO dbversion (dbversion) VALUES (20160205.1)

<br /><br />20160610.1<br />
Alter table timer_import_temp MODIFY timerkom VARCHAR(250) 

<br /><br />20160623.1<br />
ALTER TABLE licens ADD pa_aktlist INT DEFAULT 0 NOT NULL;
INSERT INTO dbversion (dbversion) VALUES (20160623.1) 

<br /><br />20160823.1<br />
CREATE TABLE national_holidays (
nh_id INT NOT NULL AUTO_INCREMENT,
nh_name INT NOT NULL DEFAULT 0,
nh_duration INT NOT NULL DEFAULT 0,
nh_date date DATE NOT NULL DEFAULT '2010-01-01',
nh_editor VARCHAR(50),
nh_editor_date DATE NOT NULL DEFAULT '2010-01-01',
PRIMARY KEY (nh_id)
);
INSERT INTO dbversion (dbversion) VALUES (20160823.1)

<br /><br />20160623.2<br />
ALTER TABLE national_holidays ADD nh_date DATE NOT NULL DEFAULT '2010-01-01';
INSERT INTO dbversion (dbversion) VALUES (20160623.2) 


<br /><br />20160908.1<br />
ALTER TABLE job
MODIFY COLUMN jobnavn varchar(150) DEFAULT '' NOT NULL;
ALTER TABLE kontaktpers
 MODIFY COLUMN navn varchar(150);
ALTER TABLE kunder
 MODIFY COLUMN kkundenavn varchar(150) DEFAULT '' NOT NULL;
ALTER TABLE aktiviteter ADD sttid TIME DEFAULT '00:00:00' NOT NULL;
ALTER TABLE aktiviteter ADD sltid TIME DEFAULT '00:00:00' NOT NULL;
INSERT INTO dbversion (dbversion) VALUES (20160908.1) 


<br /><br />20160908.2<br />
ALTER TABLE aktiviteter
CHANGE sttid aktsttid TIME DEFAULT '00:00:00' NOT NULL;
ALTER TABLE aktiviteter
CHANGE sltid aktsltid TIME DEFAULT '00:00:00' NOT NULL;
INSERT INTO dbversion (dbversion) VALUES (20160908.2)

<br /><br />20160921.1<br />
ALTER TABLE timer
MODIFY COLUMN tjobnavn varchar(150) DEFAULT '' NOT NULL;
INSERT INTO dbversion (dbversion) VALUES (20160921.1) 

<br /><br />20160921.2<br />
ALTER TABLE timer
MODIFY COLUMN kurs double(12,4) DEFAULT 0 NOT NULL;
INSERT INTO dbversion (dbversion) VALUES (20160921.2) 


<br /><br />20161025.1<br />
ALTER TABLE licens ADD (smiley_agg_lukhard INT DEFAULT 0 NOT NULL, 
mobil_week_reg_job_dd INT DEFAULT 0 NOT NULL, 
mobil_week_reg_akt_dd INT DEFAULT 0 NOT NULL, 
week_showbase_norm_kommegaa INT DEFAULT 0 NOT NULL, 
mobil_week_reg_akt_dd_forvalgt INT DEFAULT 0 NOT NULL);
ALTER TABLE kontaktpers ADD (kp_interest_christmas INT DEFAULT 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20161025.1) 


<br /><br />20161102.1<br />
ALTER TABLE lon_korsel ADD (lk_close_projects INT DEFAULT 0 NOT NULL);
ALTER TABLE lon_korsel ADD (lk_correction_on INT DEFAULT 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20161102.1) 

<br /><br />20161102.1<br />
ALTER TABLE job_ulev_ju ADD (ju_konto_label VARCHAR(50));
INSERT INTO dbversion (dbversion) VALUES (20161102.1) 


<br /><br />20161116.2<br />
ALTER TABLE traveldietexp
MODIFY COLUMN diet_dayprice double(12,2) DEFAULT 0 NOT NULL;
INSERT INTO dbversion (dbversion) VALUES (20161116.2) 

<br /><br />20161124.2<br />
ALTER TABLE aktiviteter ADD (pl_important_act INT DEFAULT 0 NOT NULL, 
pl_header INT DEFAULT 0 NOT NULL, 
pl_reccour INT DEFAULT 0 NOT NULL, 
pl_employee INT DEFAULT 0 NOT NULL);
INSERT INTO dbversion (dbversion) VALUES (20161124.1) 


<br /><br />20161214.1<br />
ALTER TABLE medarbejdere ADD (create_newemployee Int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20161214.1) 

<br /><br />20161221.1<br />
ALTER TABLE licens ADD (budget_mandatory Int NOT NULL DEFAULT 0, tilbud_mandatory Int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20161221.1) 

<br /><br />20161221.2<br />
ALTER TABLE licens ADD (show_salgsomk_mandatory Int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20161221.2) 


<br /><br />20170117.1<br />
ALTER TABLE licens ADD (multible_licensindehavere Int NOT NULL DEFAULT 0, 
fakturanr_2 Int NOT NULL DEFAULT 0, kreditnr_2 Int NOT NULL DEFAULT 0,
fakturanr_3 Int NOT NULL DEFAULT 0, kreditnr_3 Int NOT NULL DEFAULT 0,
fakturanr_4 Int NOT NULL DEFAULT 0, kreditnr_4 Int NOT NULL DEFAULT 0,
fakturanr_5 Int NOT NULL DEFAULT 0, kreditnr_5 Int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20170117.1);

ALTER TABLE licens ADD (
fakturanr_kladde_2 Int NOT NULL DEFAULT 0,
fakturanr_kladde_3 Int NOT NULL DEFAULT 0,
fakturanr_kladde_4 Int NOT NULL DEFAULT 0,
fakturanr_kladde_5 Int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20170117.2);

ALTER TABLE kunder ADD (
lincensindehaver_faknr_prioritet Int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20170117.3);
 
ALTER TABLE job ADD (
lincensindehaver_faknr_prioritet_job Int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20170117.4);


<br /><br />20170213.1<br />
ALTER TABLE job ADD (
jo_valuta Int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20170213.1) 

<br /><br />20170216.1<br />
5.56 UPDATE ON SELECT<br />
UPDATE job AS U1, valutaer AS U2 
SET U1.jo_valuta = U2.id
WHERE U2.grundvaluta = 1 AND U1.jo_valuta = 0

<br /><br />
2.23 alternativ UPDATE ON SELECT<br />
update job a
left join valutaer b on
		b.grundvaluta = 1
set
    a.jo_valuta = b.id
Where a.jo_valuta = 0
<br /><br />
UPDATE job SET jo_valuta = 1 WHERE jo_valuta = 0 

<br /><br />20170222.1<br />
ALTER TABLE medarbejdertyper ADD (
kp1_valuta Int NOT NULL DEFAULT 1);

UPDATE medarbejdertyper AS U1, valutaer AS U2 
SET U1.kp1_valuta = U2.id
WHERE U2.grundvaluta = 1 AND U1.kp1_valuta = 1;
INSERT INTO dbversion (dbversion) VALUES (20170222.1) 

<br /><br />20170223.1<br />
ALTER TABLE timer ADD (
kpvaluta Int NOT NULL DEFAULT 1);

UPDATE timer AS U1, valutaer AS U2 
SET U1.kpvaluta = U2.id
WHERE U2.grundvaluta = 1 AND U1.kpvaluta = 1;
INSERT INTO dbversion (dbversion) VALUES (20170223.1) 

<br /><br />20170224.1<br />
ALTER TABLE timer ADD (
kpvaluta_kurs double(12,2) NOT NULL DEFAULT 100);

UPDATE timer AS U1, valutaer AS U2 
SET U1.kpvaluta_kurs = U2.kurs
WHERE U2.grundvaluta = 1 AND U1.kpvaluta = U2.id;
INSERT INTO dbversion (dbversion) VALUES (20170224.1) 

<br /><br />20170326.1<br />
ALTER TABLE licens ADD (
showeasyreg_per Int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20170326.1) 


<br /><br />20170329.1<br />
ALTER TABLE login_historik ADD (
lh_longitude double (25,20) NOT NULL DEFAULT 0,
lh_latitude double (25,20) NOT NULL DEFAULT 0,
lh_longitude_logud double (25,20) NOT NULL DEFAULT 0,
lh_latitude_logud double (25,20) NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20170329.1) 

<br /><br />20170331.1<br />
ALTER TABLE medarbejdertyper ADD (
mt_mobil_visstopur INT NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20170331.1) 

<br /><br />20170410.1<br />
ALTER TABLE job ADD (
jo_valuta_kurs double (12,2) NOT NULL DEFAULT 100);
INSERT INTO dbversion (dbversion) VALUES (20170410.1)

<br /><br />20170410.2<br />
UPDATE job AS U1, valutaer AS U2 
SET U1.jo_valuta_kurs = U2.kurs
WHERE U2.id = U1.jo_valuta
INSERT INTO dbversion (dbversion) VALUES (20170410.2) 

<br /><br />20170308.1<br />
CREATE TABLE akt_bookings (
ab_id INT NOT NULL AUTO_INCREMENT,
ab_name INT NOT NULL DEFAULT 0,
ab_date DATE NOT NULL DEFAULT '2010-01-01',
ab_startdate DATETIME DEFAULT NULL,
ab_enddate DATETIME DEFAULT NULL,
ab_medid INT NOT NULL DEFAULT 0,
ab_aktid INT NOT NULL DEFAULT 0,
ab_jobid INT NOT NULL DEFAULT 0,
ab_serie INT NOT NULL DEFAULT 0,
ab_editor VARCHAR(50),
ab_editor_date DATE NOT NULL DEFAULT '2010-01-01',
ab_important INT NOT NULL DEFAULT 0,
ab_end_after INT NOT NULL DEFAULT 0,
PRIMARY KEY (ab_id)
);
INSERT INTO dbversion (dbversion) VALUES (20170308.1);


<br /><br />
CREATE TABLE akt_bookings_rel (
abl_id INT NOT NULL AUTO_INCREMENT,
abl_bookid INT NOT NULL DEFAULT 0,
abl_medid INT NOT NULL DEFAULT 0,
PRIMARY KEY (abl_id)
);
INSERT INTO dbversion (dbversion) VALUES (20170320.1)


<br /><br />20170613.1<br />
ALTER TABLE timereg_usejob ADD (
favorit Int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20170613.2) 


<br /><br />20170615.1<br />
ALTER TABLE timer ADD (
godkendtdato date DEFAULT '2002-01-01');
INSERT INTO dbversion (dbversion) VALUES (20170615.1) 

<br /><br />20170503.1<br />
ALTER TABLE licens ADD (
smiweekormonth_hr Int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20170503.1) 

<br /><br />20170503.2<br />
ALTER TABLE ugestatus ADD (
splithr Int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20170503.2) 

<br /><br />20170704.1<br />
ALTER TABLE timer ADD (
overfort Int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20170704.1) 

<br /><br />20170811.1<br />
ALTER TABLE job ADD (
jo_usefybudgetingt Int NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20170811.1) 

<br /><br />20170811.2<br />
ALTER TABLE ressourcer_ramme ADD (
rr_budgetbelob double(12,2) NOT NULL DEFAULT 0);
INSERT INTO dbversion (dbversion) VALUES (20170811.2) 

<br /><br />20170907.1<br />
CREATE TABLE eval (
eval_id double NOT NULL AUTO_INCREMENT,
eval_jobid double NOT NULL DEFAULT 0,
eval_evalvalue INT NOT NULL DEFAULT 0,
eval_jobvaluesuggested double NOT NULL DEFAULT 0,
PRIMARY KEY (eval_id)
);
INSERT INTO dbversion (dbversion) VALUES (20170907.1)


<br /><br />20170914.2<br />
ALTER TABLE fakturaer ADD (fak_rekvinr VARCHAR(255) NOT NULL DEFAULT '');
INSERT INTO dbversion (dbversion) VALUES (20170914.2)

<br /><br />20170919.1<br />
ALTER TABLE eval ADD (
eval_comment varchar(255));
INSERT INTO dbversion (dbversion) VALUES (20170919.1);

<br /><br />20170919.2<br />
ALTER TABLE kunder ADD (
kstatus int NOT NULL default 1);
INSERT INTO dbversion (dbversion) VALUES (20170919.2);

<br /><br />20170922.1<br />
CREATE TABLE abonner_file_email (
afe_id INT NOT NULL AUTO_INCREMENT,
afe_file VARCHAR(250),
afe_email VARCHAR(250),
afe_sent INT NOT NULL DEFAULT 0,
afe_date DATE NOT NULL DEFAULT '2010-01-01',
PRIMARY KEY (afe_id)
);
INSERT INTO dbversion (dbversion) VALUES (20170922.1) 


<br /><br />20170927.1<br />
ALTER TABLE timer ADD (
overfortdt Date NOT NULL DEFAULT '2002-01-01');
INSERT INTO dbversion (dbversion) VALUES (20170927.1) 

<br /><br />20171005.1<br />
CREATE TABLE job_status (
js_id INT NOT NULL DEFAULT 0,
js_navn VARCHAR(255) NOT NULL DEFAULT '',
PRIMARY KEY (js_id));
INSERT INTO dbversion (dbversion) VALUES (20171005.1);

<br /><br />20171005.2<br />
INSERT INTO job_status (js_id, js_navn) VALUES (0,'Lukket');
INSERT INTO job_status (js_id, js_navn) VALUES (1,'Aktiv');
INSERT INTO job_status (js_id, js_navn) VALUES (2,'Passiv / Til fakturering');
INSERT INTO job_status (js_id, js_navn) VALUES (3,'Tilbud');
INSERT INTO job_status (js_id, js_navn) VALUES (4,'Gennemsyn');
INSERT INTO job_status (js_id, js_navn) VALUES (5,'Evaluering');
INSERT INTO dbversion (dbversion) VALUES (20171005.2);

<br /><br />20171009.2<br />

ALTER TABLE eval DROP COLUMN eval_diff;
ALTER TABLE eval DROP COLUMN eval_suggested_hours;
ALTER TABLE eval DROP COLUMN eval_suggested_hourly_rate;
ALTER TABLE eval DROP COLUMN eval_original_price;

ALTER TABLE eval ADD 
(
eval_diff double(12,2) NOT NULL DEFAULT 0,
eval_suggested_hours double(12,2) NOT NULL DEFAULT 0,
eval_suggested_hourly_rate double(12,2) NOT NULL DEFAULT 0,
eval_original_price double(12,2) NOT NULL DEFAULT 0
);
INSERT INTO dbversion (dbversion) VALUES (20171009.2) 


<br /><br />20171022.1<br />
Alter table licens add 
(
vis_resplanner INT default 0,
vis_favorit INT default 1,
vis_projektgodkend INT default 0,
pa_tilfojvmedopret INT default 0
);
INSERT INTO dbversion (dbversion) VALUES ('20171022.1');
ALTER TABLE job ADD (
extracost double(12,2) NOT NULL DEFAULT 0, extracost_txt varchar(255));
INSERT INTO dbversion (dbversion) VALUES ('20171022.2')


<br /><br />20171027.1<br />
ALTER TABLE job MODIFY COLUMN lincensindehaver_faknr_prioritet_job VARCHAR(255);
UPDATE dbversion SET dbversion = '20171027.1' WHERe id = 1;

ALTER TABLE medarbejdere Add (med_lincensindehaver INT NOT NULL default 0);
UPDATE dbversion SET dbversion = '20171027.2' WHERe id = 1

<br /><br />20171031.1<br />
ALTER TABLE eval Add (
eval_fakbartimer double(12,2) NOT NULL default 0,
eval_fakbartimepris double(12,2) NOT NULL default 0,
ubemandet_maskine_timer double(12,2) NOT NULL default 0,
ubemandet_maskine_timePris double(12,2) NOT NULL default 0,
laer_timer double(12,2) NOT NULL default 0,
laer_timepris double(12,2) NOT NULL default 0,
easy_reg_timepris double(12,2) NOT NULL default 0,
ikke_fakbar_tid_timer double(12,2) NOT NULL default 0,
ikke_fakbar_tid_timepris double(12,2) NOT NULL default 0
);
UPDATE dbversion SET dbversion = '20171031.1' WHERe id = 1


ALTER TABLE eval Add (easy_reg_timer double(12,2) NOT NULL default 0)

<br /><br />20171030.1<br />
CREATE table your_rapports (
rap_id INT(11) NOT NULL AUTO_INCREMENT,
rap_mid INT(11) NOT NULL, 
rap_navn VARCHAR(255) NOT NULL default 'your_rapport',
rap_url VARCHAR(255) NOT NULL default 'thisfile',
rap_criteria TEXT,
rap_dato DATE DEFAULT '2002-01-01' NOT NULL,
rap_editor VARCHAR(255),
PRIMARY KEY (rap_id)
); 
UPDATE dbversion SET dbversion = '20171030.1' WHERe id = 1

<%




end if



%>