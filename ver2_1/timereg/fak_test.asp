<%response.buffer = true

useTest = 1%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/convertDate.asp"--> 
<!--#include file="inc/functions_inc.asp"-->  
<%

a = 0
dim aval
redim aval(a)

if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	'*** S�tter lokal dato/kr format. Skal inds�ttes efter kalender.
	Session.LCID = 1030
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	select case func
	case "slet"
	'*** Her sp�rges om det er ok at der slettes en medarbejder ***
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:290; top:155; width:60%; height:600; visibility:visible;">
	<table cellspacing="0" cellpadding="0" border="0" width="400" bgcolor="#ffffe1">
	<tr>
	    <td colspan=2 style="padding-top:5; padding-left:30; border-left:1px darkred solid; border-right:1px darkred solid; border-top:1px darkred solid;"><img src="../ill/alert.gif" width="44" height="45" alt="" border="0">Du er ved at <b>slette</b> en faktura-skrivelse.<br>
		Er dette korrekt?<br><br>&nbsp;</td>
	</tr>
	<tr>
	   <td width=250 style="padding:5; border-left:1px darkred solid; border-bottom:1px darkred solid;"><a href="javascript:history.back()"><img src="../ill/pile_tilbage.gif" width="10" height="6" alt="" border="0">&nbsp;Nej</a></td>
	   <td align=right style="padding:5; border-right:1px darkred solid; border-bottom:1px darkred solid;"><a href="fak.asp?menu=stat_fak&func=sletok&id=<%=id%>&FM_job=<%=Request("FM_job")%>&FM_usedatointerval=<%=request("FM_usedatointerval")%>&FM_start_dag=<%=request("FM_start_dag")%>&FM_start_mrd=<%=request("FM_start_mrd")%>&FM_start_aar=<%=request("FM_start_aar")%>&FM_slut_dag=<%=request("FM_slut_dag")%>&FM_slut_mrd=<%=request("FM_slut_mrd")%>&FM_slut_aar=<%=request("FM_slut_aar")%>&FM_fakint=<%=trim(request("FM_fakint"))%>">Ja&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"></a></td>
	</tr>
	</table>
	<br><br>
		<br>
		<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
		<br>
		<br>
	</div>
	<%
	case "sletok"
	
	
	oConn.execute("DELETE FROM fakturaer WHERE Fid = "& id &"")
	oConn.execute("DELETE FROM faktura_det WHERE fakid = "& id &"")
	oConn.execute("DELETE FROM fak_med_spec WHERE fakid = "& id &"")
	
	if request("FM_usedatointerval") = "1" then
	usedval = 1
	else
	usedval = 0
	end if
	
	Response.redirect "fak_osigt.asp?menu=stat_fak&FM_job=" & Request("FM_job") &"&FM_usedatointerval="&usedval&"&FM_start_dag="&request("FM_start_dag")&"&FM_start_mrd="&request("FM_start_mrd")&"&FM_start_aar="&request("FM_start_aar")&"&FM_slut_dag="&request("FM_slut_dag")&"&FM_slut_mrd="&request("FM_slut_mrd")&"&FM_slut_aar="&request("FM_slut_aar")&"&FM_fakint="&request("FM_fakint")
	
	case "dbred", "dbopr" 
	
	function SQLBlessDOT(s)
	dim tmpdot
	tmpdot = s
	tmpdot = replace(tmpdot, ".", "")
	SQLBlessDOT = tmpdot
	end function
	
	function SQLBless(s)
	dim tmp
	tmp = s
	tmp = replace(tmp, ",", ".")
	SQLBless = tmp
	end function
	
	function SQLBlessPunkt(s)
	dim tmpPunkt
	tmpPunkt = s
	tmpPunkt = replace(tmpPunkt, ".", ",")
	SQLBlessPunkt = tmpPunkt
	end function
	
	intBeloeb = SQLBless(Request("FM_Beloeb"))
	intTimer = SQLBless(Request("FM_Timer"))
	intFaknum = SQLBless(Request("FM_faknr"))
	
	if len(request("FM_faktype")) <> 0 then
	intFaktype = request("FM_faktype")
	else
	intFaktype = 0
	end if
	
	
	'*** Her tjekkes om alle required felter er udfyldt. ***
	if len(Request("FM_faknr")) = 0 OR len(Request("FM_timer")) = 0 OR len(Request("FM_beloeb")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<%
	errortype = 26
	call showError(errortype)
	
	else%>
			<!--#include file="inc/isint_func.asp"-->
			<%
			call erDetInt(intBeloeb)
			if isInt > 0 then
			%>
			<!--#include file="../inc/regular/header_inc.asp"-->
			<!--#include file="../inc/regular/topmenu_inc.asp"-->
			<%
			errortype = 27
			call showError(errortype)
			
			isInt = 0
					else
					
					call erDetInt(intTimer)
					if isInt > 0 then
					%>
					<!--#include file="../inc/regular/header_inc.asp"-->
					<!--#include file="../inc/regular/topmenu_inc.asp"-->
					<%
					errortype = 28
					call showError(errortype)
					isInt = 0
							
							else
								
								
									'** Inds�tter cookie **
									Response.Cookies("clastfaknr") = "0"
									Response.Cookies("clastfaknr").Expires = date + 1
									
									strSQL2 = "SELECT faknr FROM fakturaer WHERE faknr = '"& intFaknum &"' AND Fid <> "& id &""
									oRec2.open strSQL2, oConn, 3
									
									if not oRec2.EOF then 'faknr findes allerede
									%>
									<!--#include file="../inc/regular/header_inc.asp"-->
									<!--#include file="../inc/regular/topmenu_inc.asp"-->
									<%
									errortype = 29
									call showError(errortype)
									oRec2.close
									
									else
							
							
							
							'*** Hvis alle required er udfyldt ***
							
									fakDato = Request("FM_start_aar") & "/" & Request("FM_start_mrd") & "/" & Request("FM_start_dag") '& time 
									showfakDato = Request("FM_start_dag") & "/" & Request("FM_start_mrd") & "/" & Request("FM_start_aar")
									strEditor = session("user")
									strDato = session("dato")
									strjobid = Request("jobid")
									intFakbetalt = request("FM_betalt")
									intfakadr = request("FM_Kid")
									
									if request("FM_att_vis") = "on" then
									strAtt = request("FM_att")
									else
									strAtt = "Ingen"
									end if
									
									if intFakbetalt <> 1 then
									intFakbetalt = 0
									else
									intFakbetalt = intFakbetalt 
									end if
									
									intEnhedsang = request("FM_enheds_ang")
																		
									'** Database format dato **
									'dtb_dato = Request("FM_b_start_aar") & "/" & Request("FM_b_start_mrd") & "/" & Request("FM_b_start_dag")
									dtb_dato = year(now) & "/"& month(now) & "/" & day(now)
									
									function SQLBlessK(s)
									dim tmp
									tmp = s
									tmp = replace(tmp, "'", "''")
									SQLBlessK = tmp
									end function
									
									'************************
									'*** findes allerede ****
									'************************
									function SQLBless2(s)
									dim tmp2
									tmp2 = s
									tmp2 = replace(tmp2, ",", ".")
									SQLBless2 = tmp2
									end function
									
									function SQLBless3(s)
									dim tmp3
									tmp3 = s
									tmp3 = replace(tmp3, ".", ",")
									SQLBless3 = tmp3
									end function
									
									
									varKonto = request("FM_kundekonto")
									varModkonto = request("FM_modkonto")
									strKomm = SQLBlessK(Request("FM_Komm"))
									antalAkt = Request("antal_x")
									
									subtotal = 0
									
									
									'*** Beregner moms ****
									intTotal = intBeloeb
									
									
									if request("FM_debkre") = "d" then
									intTotal_konto = intTotal
									intTotal_modkonto = 0 - SQLBless3(intTotal)
									intTotal_modkonto = SQLBless2(intTotal_modkonto)
									else
									intTotal_konto = 0 - SQLBless3(intTotal)
									intTotal_konto = SQLBless2(intTotal_konto)
									intTotal_modkonto = intTotal
									end if
									
									intMoms = 0
									intNetto = 0
									intModMoms = 0
									intModNetto = 0
									showmoms = 0
									
									
									'** Beregner moms og netto **
									'*** Konto ***
									strSQL = "SELECT momskode, kvotient FROM kontoplan LEFT JOIN momskoder ON (momskoder.id = momskode) WHERE kontonr = "& varKonto
									oRec.open strSQL, oConn, 3 
									if not oRec.EOF then
										if oRec("kvotient") <> 1 then
											
												if oRec("kvotient") = 5 then
												intMoms = (SQLBless3(intTotal_konto) * 0.25)  'Alm
												else
												intMoms = (SQLBless3(intTotal_konto) * 0.05)  'Rep
												end if
											
												if intMoms < 0 then
												showmoms = -(intMoms)
												else
												showmoms = intMoms
												end if
											
											intMoms = SQLBless2(intMoms) 
											
										else
										intMoms = 0
										end if
									intNetto = intTotal_konto 
									end if
									oRec.close
									
									
									'*** Komma til punktum
									showmoms = replace(showmoms, ",", ".")
									
									'** Beregner moms og netto **
									'*** Modkonto ***
									strSQL = "SELECT momskode, kvotient FROM kontoplan LEFT JOIN momskoder ON (momskoder.id = momskode) WHERE kontonr = "& varModkonto
									oRec.open strSQL, oConn, 3 
									if not oRec.EOF then
										if oRec("kvotient") <> 1 then
												if oRec("kvotient") = 5 then
												intModMoms = (SQLBless3(intTotal_modkonto) * 0.25)  'Alm
												else
												intModMoms = (SQLBless3(intTotal_modkonto) * 0.05)  'Rep
												end if
												
										intModMoms = SQLBless2(intModMoms) 
										
										else
										intModMoms = 0
										end if
									intModNetto = intTotal_modkonto 'SQLBless2((intTotal_modkonto - (formatnumber(intModMoms, 2))))
									end if
									oRec.close
									'***********************************************************************************
							
									
							
							if request("FM_land_vis") = "on" then
							intModland = 1
							else
							intModland = 0
							end if
							
							if request("FM_att_vis") = "on" then
							intModAtt = 1
							else
							intModAtt = 0
							end if
							
							if request("FM_tlf_vis") = "on" then
							intModTlf = 1
							else
							intModTlf = 0
							end if
							
							if request("FM_cvr_vis") = "on" then
							intModCvr = 1
							else
							intModCvr = 0
							end if
							
							if request("FM_yswift_vis") = "on" then
							intAfsSwift = 1
							else
							intAfsSwift = 0
							end if
							
							
							if request("FM_yiban_vis") = "on" then
							intAfsIban = 1
							else
							intAfsIban = 0
							end if
							
							if request("FM_ycvr_vis") = "on" then
							intAfsICVR = 1
							else
							intAfsICVR = 0
							end if
								
							if request("FM_yemail_vis") = "on" then
							intAfsEmail = 1
							else
							intAfsEmail = 0
							end if
							
							if request("FM_ytlf_vis") = "on" then
							intAfsTlf = 1
							else
							intAfsTlf = 0
							end if		
									
							'************************************************************************************
							'*** Opdaterer / Redigerer faktura 													*
							'************************************************************************************
							if func = "dbred" then
							
											strSQLupd = "UPDATE fakturaer SET"_
											&" faknr = '"& intFaknum &"', "_
											&" fakdato = '"& fakDato &"', "_
											&" timer = "& intTimer  &", "_
											&" beloeb = "& intBeloeb &", "_
											&" kommentar = '"& strKomm &"', "_
											&" editor = '"& strEditor &"', "_
											&" dato = '"& strDato &"', "_
											&" tidspunkt = '23:59:59', "_
											&" betalt = "& intFakbetalt &", "_
											&" b_dato = '"& dtb_dato &"', "_
											&" fakadr = "& intfakadr &", "_
											&" att = '"& strAtt &"', "_
											&" konto = "& varKonto &", "_
											&" modkonto = "& varModkonto &", "_
											&" faktype = "& intFaktype &", "_
											&" vismodland = "& intModland &", "_
											&" vismodatt = "& intModAtt &", "_
											&" vismodtlf = "& intModTlf &", "_
											&" vismodcvr = "& intModCvr &", "_
											&" visafstlf = "& intAfsTlf &", "_
											&" visafsemail = "& intAfsEmail &", "_
											&" visafsswift = "& intAfsSwift &", "_
											&" visafsiban = "& intAfsIban &", "_
											&" visafscvr = "& intAfsICVR &", "_
											&" moms = "& showmoms &", enhedsang = "& intEnhedsang &""_
											&" WHERE Fid = "& id 
											
											'Response.write (strSQLupd)
											oConn.execute(strSQLupd)
											
											strSQL = "SELECT oprid FROM fakturaer WHERE Fid="& id
											oRec.open strSQL, oConn, 3 
											if not oRec.EOF then
											oprid = oRec("oprid")
											end if
											oRec.close 
											
											
											if oprid <> 0 then '0 hvis fak er fra f�r 1/8-2004
											strSQL = "DELETE FROM posteringer WHERE oprid ="&oprid&""
											'Response.write strSQL
											oconn.execute(strSQL)
											end if
											
													if intFakbetalt <> 0 then '** Opretter tilh�rende posteringer hvis det ikke er kladde. **	
													'***************** Redigerer posteringer p� de valgte konti **********
													oConn.execute("INSERT INTO posteringer (editor, dato, modkontonr, kontonr, bilagsnr, beloeb, nettobeloeb, moms, tekst, posteringsdato, status, att, oprid)"_
													&" VALUES ('"& strEditor &"', '"& strDato &"', '"& varModkonto &"', '"& varKonto &"', '"& intFaknum &"', "_
													&" "& intTotal_konto &", "&intNetto&", "&intMoms&", '"& request("FM_jobnavn") &"', "_
													&" '"& fakDato &"', "_
													&" 1, "& session("Mid") &", "& oprid &")")
													
													
													'**** Modkonto ***
													oConn.execute("INSERT INTO posteringer (editor, dato, modkontonr, kontonr, bilagsnr, beloeb, nettobeloeb, moms, tekst, posteringsdato, status, att, oprid)"_
													&" VALUES ('"& strEditor &"', '"& strDato &"', '"& varKonto &"', '"& varModkonto &"', '"& intFaknum &"', "_
													&" "& intTotal_modkonto &", "&intModNetto&", "&intModMoms&", '"& request("FM_jobnavn") &"', "_
													&" '"& fakDato &"', "_
													&" 1, "& session("Mid") &", "& oprid &")")
													'*********************************************************************
													end if
							thisfakid = id
							end if
							'***
											
											
											
							'************************************************************************************
							'*** Opretter faktura *********														*
							'************************************************************************************
							if func = "dbopr" then
												
												if intFaktype <> 0 then '*** kreditnota eller rykker
												parentfak = request("id")
												else
												parentfak = 0
												end if
													
													'***tidspunkt = 23:23:59 pga luk for indtastning muligheden p� timeregsiden
													
													strSQL = ("INSERT INTO fakturaer"_
													&" (faknr, fakdato, jobid, timer, beloeb, kommentar, dato, editor, tidspunkt, "_
													&" betalt, b_dato, fakadr, att, faktype, konto, modkonto, parentfak, "_
													&" vismodland, vismodatt, vismodtlf, vismodcvr, visafstlf, visafsemail, "_
													&" visafsswift, visafsiban, visafscvr, moms, enhedsang) VALUES ("_
													&" '"& intFaknum &"',"_
													&" '"& fakDato &"',"_
													&" "& strjobid &","_
													&" "& intTimer &","_
													&" "& intBeloeb &","_
													&" '"& strKomm &"',"_
													&" '"& strDato &"',"_
													&" '"& strEditor &"', '23:59:59', "_
													&" "& intFakbetalt &", '"& dtb_dato &"', "& intfakadr &", "_
													&" '"& strAtt &"', "& intFaktype &", '"& varKonto &"', "_
													&" '"& varModkonto &"', "& parentfak &", "_
													&" "& intModland &", "& intModAtt &", "& intModTlf &", "_
													&" "& intModCvr &", "& intAfsTlf &", "& intAfsEmail &", "& intAfsSwift &", "_
													&" "& intAfsIban &", "& intAfsICVR &", "& showmoms &", "& intEnhedsang &")")
													
													'Response.write strSQL
													oConn.execute(strSQL)
													
													'**** Sletter faknr fra resevering ***
													'strSQLF = "DELETE FROM fak_opr_faknr WHERE sesid = "& session("mid")
													'oConn.execute(strSQLF)
													
													
														'** Henter fak id ***
														strSQL3 = "SELECT Fid FROM fakturaer"
														oRec3.open strSQL3, oConn, 3
														oRec3.movelast
														if not oRec3.EOF then
														thisfakid = cint(oRec3("Fid")) 
														end if 
														oRec3.close
													
													
											'** Opretter tilh�rende posteringer hvis det ikke er kladde. **		
											if intFakbetalt <> 0 then 
													
														'***************** Opretter posteringer p� de valgte konti **********
														strSQLK = "INSERT INTO posteringer (editor, dato, modkontonr, kontonr, bilagsnr, beloeb, nettobeloeb, moms, tekst, posteringsdato, status, att)"_
														&" VALUES ('"& strEditor &"', '"& strDato &"', '"& varModkonto &"', '"& varKonto &"', '"& intFaknum &"', "_
														&" "& intTotal_konto &", "&intNetto&", "&intMoms&", '"& request("FM_jobnavn") &"', "_
														&" '"& fakDato &"', "_
														&" 1, "& session("Mid") &")"
														
														oConn.execute(strSQLK)
														
														'*** Finder det netop opr. id ***
														strSQL = "SELECT id FROM posteringer ORDER BY id DESC"
														oRec.Open strSQL, oConn, 3 
														if not oRec.EOF then
														oprid = oRec("id")
														end if
														oRec.close
														
														oConn.execute("UPDATE posteringer SET oprid = "& oprid &" WHERE id = "& oprid &"")
														oConn.execute("UPDATE fakturaer SET oprid = "& oprid &" WHERE Fid = "& thisfakid &"")
														
														'**** Modkonto ***
														strSQLMK = "INSERT INTO posteringer (editor, dato, modkontonr, kontonr, bilagsnr, beloeb, nettobeloeb, moms, tekst, posteringsdato, status, att, oprid)"_
														&" VALUES ('"& strEditor &"', '"& strDato &"', '"& varKonto &"', '"& varModkonto &"', '"& intFaknum &"', "_
														&" "& intTotal_modkonto &", "&intModNetto&", "&intModMoms&", '"& request("FM_jobnavn") &"', "_
														&" '"& fakDato &"', "_
														&" 1, "& session("Mid") &", "& oprid &")"
														
														oConn.execute(strSQLMK)
														
														'*********************************************************************
														'*** Hvis Rykker oprettes krediteres tidligere faktura ***************
														'********************************************************************* 
														if intFaktype = 2 then
																strSQL = "SELECT oprid FROM fakturaer WHERE Fid="& request("id")
																'Response.write strSQL & "<br><br>"
																oRec.open strSQL, oConn, 3 
																if not oRec.EOF then
																oprid = oRec("oprid")
																end if
																oRec.close 
																
																
																	if oprid <> 0 then '** undg� tidligere fakturaer fra f�r �konomi del.
																	
																			strSQL = "SELECT editor, dato, modkontonr, "_
																			&" kontonr, bilagsnr, beloeb, nettobeloeb, moms, "_
																			&" tekst, posteringsdato, status, att "_
																			&" FROM posteringer WHERE oprid=" & oprid
																			
																			oRec.open strSQL, oConn, 3 
																			if not oRec.EOF then
																				
																				useModkonto = oRec("modkontonr")
																				useKonto = oRec("kontonr")
																				useTotal_konto = oRec("beloeb")
																				useNetto = oRec("nettobeloeb")
																				useMoms = oRec("moms")
																				
																			end if
																			oRec.close 
																				
																				if trim(len(request("bilagsnrkredit"))) <> 0 then
																				useFaknum = request("bilagsnrkredit")
																				else
																				useFaknum = 0
																				end if
																			
																			strSQLK = "INSERT INTO posteringer (editor, dato, modkontonr, kontonr, bilagsnr, beloeb, nettobeloeb, moms, tekst, posteringsdato, status, att, oprid)"_
																			&" VALUES ('"& strEditor &"', '"& strDato &"', '"& useModkonto &"', '"& useKonto &"', '"& useFaknum &"', "_
																			&" "& -(SQLBless2(useTotal_konto)) &", "&-(SQLBless2(useNetto))&", "&-(SQLBless2(useMoms))&", '"& request("FM_jobnavn") &"', "_
																			&" '"& fakDato &"', "_
																			&" 1, "& session("Mid") &", "& oprid &")"
																			
																			oConn.execute(strSQLK)
																			'Response.write strSQLK
																			
																			'**** Modkonto ***
																			strSQLMK = "INSERT INTO posteringer (editor, dato, modkontonr, kontonr, bilagsnr, beloeb, nettobeloeb, moms, tekst, posteringsdato, status, att, oprid)"_
																			&" VALUES ('"& strEditor &"', '"& strDato &"', '"& useKonto &"', '"& useModkonto &"', '"& useFaknum &"', "_
																			&" "& SQLBless2(useTotal_konto) &", "&SQLBless2(useNetto)&", "&SQLBless2(useMoms)&", '"& request("FM_jobnavn") &"', "_
																			&" '"& fakDato &"', "_
																			&" 1, "& session("Mid") &", "& oprid &")"
																			
																			oConn.execute(strSQLMK)
																	 	
																	end if 'Tidligere
														end if  'Rykker
											end if '** kladde / endelig
											'***
							end if	'** Opret
							'***
							
														
												
							'***********************************************************************************
							'*** Hvis faktura allerede er oprrettet en gang i denne session ***
							'***********************************************************************************
								
								 
								'************************************************************** 
								'*** Inds�tter akt i fak_det 
								'**************************************************************
								if func = "dbred" then
								oConn.execute("DELETE FROM faktura_det WHERE fakid = "& id &"")
								thisfakid = id
								end if
								
								'***** n starter p� 1 **************
								'***** Aktivitets udspecificering ************
								for intcounter = 0 to antalAkt - 1
								
								thisAktId = request("aktId_n_"&intcounter&"")
									
									'** antal aktiviter udskrevet pga. forskellige timepriser 
									'antalsumaktprakt = request("antal_subtotal_akt_"&intcounter&"") 
									antalsumaktprakt = request("antal_n_"&intcounter&"") 
									
									for intcounter3 = -(antalsumaktprakt) to antalsumaktprakt
									
												
												'*** Enhedsprosen p� denn akt.
												if len(request("FM_enhedspris_"& intcounter &"_"&intcounter3&"")) <> 0 then
												enhpris = SQLBless2(request("FM_enhedspris_"& intcounter &"_"&intcounter3&""))
												else
												enhpris = 0
												end if
									
										'*** Vis sum aktivitet p� print (og DB)
										if len(request("FM_show_akt_"& intcounter &"_"&intcounter3&"")) = 1 then
												
												timerThis = request("FM_timerthis_"& intcounter &"_"&intcounter3&"")
												if len(timerThis) <> 0 then
												timerThis = SQLBless2(timerThis)
												else 
												timerThis = 0
												end if
												
												if len(request("FM_beloebthis_"& intcounter &"_"&intcounter3&"")) <> 0 then
												beloebThis = SQLBless2(request("FM_beloebthis_"& intcounter &"_"&intcounter3&""))
												else
												beloebThis = 0
												end if
												
												
												strSQL_sumakt = ("INSERT INTO faktura_det "_
												&" (antal, beskrivelse, aktpris, fakid, enhedspris, aktid, showonfak) "_
												&" VALUES ("& timerThis &", "_
												&"'" & SQLBlessK(request("FM_aktbesk_"& intcounter &"_"&intcounter3&"")) &"', "_
												&""  & beloebThis &", "_
												&""  & thisfakid &", "& enhpris &", "& thisAktId &", 1)")
												
												'Response.write strSQL_sumakt  & "<br>"
												oConn.execute(strSQL_sumakt)
												
										end if 'show sumaktivitet
													
													
										'***** Medarbejder udspecificering i db ************	
										antalmedspec2 = request("antal_n_"&intcounter&"") 'medarbejdere
										for intcounter2 = 0 to antalmedspec2 - 1
												
												'*** Passer timeprisen p� denne akt og medarbejder **
												thismedarbtpris = request("FM_mtimepris_"& intcounter2&"_"&intcounter&"") 
												if SQLBless2(thismedarbtpris) = SQLBless2(enhpris) then
											
													'** registrerer timer (fak og vent) p� alle medarbejdere 
													if len(request("FM_mid_"&intcounter2&"_"&intcounter&"")) <> 0 then
													thisMid = request("FM_mid_"&intcounter2&"_"&intcounter&"")
													else
													thisMid = 0
													end if 
													
														'* Nulstiller evt. tidligere indtastninger p� denne faktura ***
														oConn.execute("DELETE FROM fak_med_spec WHERE fakid = "& thisfakid &" AND aktid = "&thisAktId&" AND mid = "&thisMid&"")
													 	
														'* Bel�b
														if len(request("FM_mbelob_"& intcounter2 &"_"&intcounter&"")) <> 0 then
														useBeloeb = SQLBless2(request("FM_mbelob_"& intcounter2 &"_"&intcounter&""))
														else
														useBeloeb = 0
														end if
														
														'* Venter
														if len(request("FM_m_vent_"&intcounter2&"_"&intcounter&"")) <> 0 then
														useVenter = SQLBless2(request("FM_m_vent_"&intcounter2&"_"&intcounter&""))
														else
														useVenter = 0
														end if
														
														
														'** Nulstiler altid vente timer inden der tildeles nye vente timer for denne medarbejder p� denne aktivitet 
														'** (Uanset hvilken fak akt. h�rer til)
														
														oConn.execute("UPDATE fak_med_spec SET venter = 0 WHERE aktid = "&thisAktId&" AND mid = "&thisMid&"")
													 	
														
														if len(request("FM_m_fak_"&intcounter2&"_"&intcounter&"")) <> 0 then
															'*** Hvis show sum-akt ikke er true skal faktimer altid v�re = 0
															if len(request("FM_show_akt_"& intcounter &"_"&intcounter3&"")) = 1 then 
															useFak = SQLBless2(request("FM_m_fak_"&intcounter2&"_"&intcounter&""))
															else
															useFak = 0
															end if
														else
															useFak = 0
														end if
														
														
														if len(request("FM_mtimepris_"& intcounter2&"_"&intcounter&"")) <> 0 then
														usemedTpris = SQLBless2(request("FM_mtimepris_"& intcounter2&"_"&intcounter&""))
														else
														usemedTpris = 0
														end if 
														
														if request("FM_show_medspec_"&intcounter2&"_"&intcounter&"") = "show" then
														showonfak = 1
														else
														showonfak = 0
														end if
														
														'*** Inds�tter i db ****
														strSQL = ("INSERT INTO fak_med_spec (fakid, aktid, mid, fak, venter, tekst, enhedspris, beloeb, showonfak) "_
														&" VALUES ("& thisfakid &", "&thisAktId&", "_
														&"" &request("FM_mid_"&intcounter2&"_"&intcounter&"")&", "_
														&"" &useFak&", "_
														&"" &useVenter&", "_
														&"'"&SQLBlessK(request("FM_m_tekst_"& intcounter2&"_"&intcounter&""))&"', "_
														&"" &usemedTpris&", "_
														&"" &useBeloeb&", "& showonfak &")")
														oConn.execute(strSQL)
														
														
													end if 'thismedarbtpris
											next 'Medarbejdere
										next 'Sumaktiviteter
								next 'Antal aktiviteter
								
							
							'**********************************************************************************************
							
							
												
												
											'*** Viser den oprettede faktura til print *****%>
											<!-------------------------------Sideindhold------------------------------------->
											<%
											Response.redirect "fak_godkendt.asp?id="&thisfakid&"&menu=stat_fak&FM_job="&Request("FM_job")&"&FM_usedatointerval="&request("FM_usedatointerval")&"&FM_start_dag_ival="&request("FM_start_dag_ival")&"&FM_start_mrd_ival="&request("FM_start_mrd_ival")&"&FM_start_aar_ival="&request("FM_start_aar_ival")&"&FM_slut_dag_ival="&request("FM_slut_dag_ival")&"&FM_slut_mrd_ival="&request("FM_slut_mrd_ival")&"&FM_slut_aar_ival="&request("FM_slut_aar_ival")&"&FM_fakint_ival="&request("FM_fakint_ival")
											
			
											
									end if
							end if
					end if
			end if
	case else
	'**************************** Opret/rediger Faktura Data ************************************
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<SCRIPT language=javascript src="inc/fak_func.js"></script>
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!--#include file="inc/fak_inc_subs.asp" -->
	<!-------------------------------Sideindhold------------------------------------->
	<%
	public strDag
	public strMrd
	public strMrdNavn
	public function dag() 
		if func <> "red" then
				
			Select case strMrd
			case "4", "6", "9", "11"
				if strDag > 30 then
				strDag = 30
				else
				strDag = strDag
				end if
			case 2
				if strDag > 28 then
				strDag = 28
				else
				strDag = strDag
				end if
			case else
			strDag = strDag
			end select
			
		end if
			
	end function
	
	
	sub top
	%>
	<div id="sidetop" style="position:absolute; border:2; left:<%=pleft+0%>; top:<%=ptop%>; visibility:visible; z-index:100;">
	<br><img src="../ill/header_fak.gif" alt="" border="0" width="279" height="34">
	<br>
	<%
	end sub
	%>
	
	
	<%
	jobnr = Request("jobnr")
	jobid = request("jobid")
	timeforbrug = 0
	faktureretBeloeb = 0
	
	pleft = 200
	ptop = 50
	
	Redim thisaktid(0)
	Redim thisaktnavn(0)
	Redim thisAktTimer(0)
	Redim thisAktBeloeb(0)
	Redim usefastpris(0)
	Redim thisTimePris(0)
	
	
	if func = "red" then
	'*********************************************************************
	' Rediger faktura 
	'*********************************************************************
	
	
						strSQL = "SELECT Fid, faknr, fakdato, jobid, timer, beloeb, kommentar, dato, editor, betalt, b_dato, fakadr, att, faktype, konto, modkonto FROM fakturaer WHERE Fid = "& id 
						
						oRec.open strSQL, oConn, 3
						
						if not oRec.EOF then
						strEditor = oRec("editor")
						strFaknr = oRec("faknr")
						strTdato = formatdatetime(oRec("fakdato"), 2)
								'*** Hvis faktype er <> for faktura skiftes dbfunc til dbopr ***
								if request("faktype") <> 0 AND request("rykkreopr") = "j" then
									dbfunc = "dbopr"
									strNote = "<font color=darkred>Husk at �ndre nr. i forhold til det oprindelige faktura nummer!</font>"
									strNote1 = "Nedenst�ende udspecificering er en kopi af den oprindelige faktura."
									rykelnotopr = 1
										Select case request("faktype")
										case 1
										strKom = "Kontakt os venligst vedr�rende overf�rsel / tilbagebetaling."
										case 2
										strKom = "Hvis jeres indbetaling har krydset dette brev, s� se venligst bort fra denne skrivelse."
										End select
								else
									rykelnotopr = 0
									strNote = ""
									dbfunc = "dbred"
									strNote1 = ""
									strKom = oRec("kommentar")
								end if
								
						strTimer = oRec("timer")
						strBeloeb = oRec("beloeb")
						StrUdato = "12/12/2007"
						strDato = oRec("dato")
						intBetalt = oRec("betalt")
						strB_dato = oRec("b_dato")
						intFakid = oRec("Fid")
						intfakadr = oRec("fakadr")
						strAtt = oRec("att")
						intFaktype = oRec("faktype")
						intKonto = oRec("konto")
						intModKonto = oRec("modkonto")
						end if
						oRec.close
						
						
						if len(intModKonto) <> 0 then
						intModKonto = intModKonto
						else
						intModKonto = 0
						end if
						
						if intBetalt = 1 then
						betaltch = "checked"
						else
						betaltch = ""
						end if 
						
						call faktureredetimerogbelob()
						
						strSQL = "SELECT jobTpris, budgettimer, fastpris, jobknr, jobnr, jobnavn, ikkebudgettimer, jobans1, jobans2 FROM job WHERE id = " & jobid
						oRec.open strSQL, oConn, 3
						if not oRec.EOF then
							intIkkeBtimer = oRec("ikkebudgettimer")
							intBudgettimer = oRec("budgettimer")
							intJobTpris = oRec("jobTpris")
							fastpris = oRec("fastpris")
							if intfakadr <> 0 then
							intjobknr = intfakadr
							else
							intjobknr = oRec("jobknr")
							end if
							intjobnr = oRec("jobnr")
							strJobnavn = oRec("jobnavn")
							
							jobans1 = oRec("jobans1")
							jobans2 = oRec("jobans2")
						end if
						oRec.close
		'*******************************************************************************************
	
	
		call top
		'** Mulighed for at slette og se sidst redigeret dato ****
		%><img src="../ill/blank.gif" width="10" height="318" alt="" border="0">
		<table cellspacing="0" cellpadding="5" border="0" width="370" bgcolor="#EFF3FF">
		<%
		if rykelnotopr <> 1 then%>
			<tr>
				<td valign="bottom" style="border:1px #003399 solid;" height="25">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b>
				<!--&nbsp;&nbsp;<a href="fak.asp?menu=stat_fak&func=slet&id=<=id%>&FM_job=<=Request("FM_job")%>"><img src="../ill/slet_eks.gif" width="20" height="20" alt="Slet denne faktura" border="0"></a>-->
				</td>
			</tr>
		<%end if%>
		</table>
		</div>
		<%
	else
	call top
	%>
	</div>
	<%
		'************************************************************************'
		' Opret faktura 
		'************************************************************************
			
			
			'****************************************************************
			'*** Finder faknr ***
			'****************************************************************
			if request.Cookies("clastfaknr") = "0" OR len(request.Cookies("clastfaknr")) = 0 then
				lastFaknr = 0 	
				
				strSQL = "SELECT Fid, faknr FROM fakturaer WHERE faktype = 0 ORDER BY fid DESC" 
				oRec.open strSQL, oConn, 3
				
				while not oRec.EOF 
				
				call erDetInt(oRec("faknr"))
					if isInt = 0 then
						if cint(oRec("faknr")) >= lastFaknr then
						lastFaknr = cint(oRec("faknr")) + 1
						else
						lastFaknr = lastFaknr
						end if
					end if
				isInt = 0
				oRec.movenext
				wend
				oRec.close
				
				
				
			
				'*** Er faknr allerede under opr? ****
				strSQL = "SELECT id, faknr FROM fak_opr_faknr"
				oRec.open strSQL, oConn, 3
				if not oRec.EOF then
			
					if cint(lastFaknr) <= cint(oRec("faknr")) then
						lastFaknr = oRec("faknr") + 1
					end if
				
				end if
				oRec.close
				
				
				
			if lastFaknr <> 0 then 
			'*** Opdaterer lastfaknr ***
			strSQLF = "DELETE FROM fak_opr_faknr WHERE faknr <> 0"
			oConn.execute(strSQLF)
			strSQLF = "INSERT INTO fak_opr_faknr (faknr) VALUES ('"& lastFaknr &"')"
			oConn.execute(strSQLF)
			end if
		
			
			'** Inds�tter cookie **
			Response.Cookies("clastfaknr") = lastFaknr
			Response.Cookies("clastfaknr").Expires = date + 65
			
			else
				'Response.Cookies("clastfaknr") = ""
				lastFaknr = request.Cookies("clastfaknr")
			end if
		
		
		call faktureredetimerogbelob()
		
		strSQL = "SELECT jobTpris, budgettimer, fastpris, jobknr, jobnr, jobnavn, ikkebudgettimer, jobans1, jobans2 FROM job WHERE id = " & jobid
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then
			intIkkeBtimer = oRec("ikkebudgettimer")
			intBudgettimer = oRec("budgettimer")
			intJobTpris = oRec("jobTpris")
			fastpris = oRec("fastpris")
			'*** Hvis der er skiftet kunde ****
			if len(request("newkundeid")) <> 0 then
			intjobknr = request("newkundeid")
			else
			intjobknr = oRec("jobknr")
			end if
			intjobnr = oRec("jobnr")
			strJobnavn = oRec("jobnavn")
			
			jobans1 = oRec("jobans1")
			jobans2 = oRec("jobans2")
		end if
		oRec.close
		
		intKonto = 0
		intModKonto = 0
		
		strFaknr = lastFaknr
		
		strEditor = ""
		strTdato = "" 'month(now)&"/"&day(now)&"/"& year(now)
		strKom = ""
		dbfunc = "dbopr"
		varSubVal = "Opretpil" 
		'strTimer = Request("ttf")
		'strBeloeb = Request("ktf")
		'*******************************************************************************************
	end if
	%>
	
	<%
	'** Skift kunde er fjernet pga. det nulstiller datointerval 
	'** Brug istedet forrretningsomr�der til at angive en attribut der s�ledes kan summere salg.
	'** Dette betyder at der skal oprettes et job, for hver enkelt salg af. f.eks timeOut, for at der kan udesendes en faktura.
	brugskiftkunde = 0
	if func <> "red" AND brugskiftkunde = 1 then%>
	<div id="vaelgmodtager" style="position:absolute; left:<%=pleft%>; top:<%=ptop+310%>; visibility:visible;">
	<table cellspacing="0" cellpadding="0" border="0" width="300">
	<tr><form action="fak.asp?menu=stat_fak" method="post" name="skiftkunde">
		<td><br>Hvis ovenst�ende firmaoplysninger ikke passer,<br>kan der skiftes kunde her: <br>
		
			<!-- hvis der skiftes kunde mister man det valgte datointerval i fakkorsel. -->
			<input type="hidden" name="FM_jobnavn" value="<%=strJobnavn%>">
			<input type="hidden" name="FM_job" value="<%=Request("FM_job")%>">
			<input type="hidden" name="id" value="<%=id%>">
			<input type="hidden" name="jobnr" value="<%=jobnr%>">
			<input type="hidden" name="jobid" value="<%=jobid%>">
			<input type="hidden" name="ttf" value="<%=strTimer%>">
			<input type="hidden" name="ktf" value="<%=strBeloeb%>">
			<input type="hidden" name="faktype" value="<%=request("faktype")%>">
			<input type="hidden" name="faknr" value="<%=request("faknr")%>">
			<input type="hidden" name="rykkreopr" value="<%=request("rykkreopr")%>">
		<select name="newkundeid" style="font-family: arial,helvetica,sans-serif; font-size: 10px;">
		<%
		strSQL = "SELECT Kid, kkundenavn, kkundenr FROM kunder WHERE ketype <> 'e'" 		
		oRec.open strSQL, oConn, 3
		While not oRec.EOF
		if cint(intjobknr) = cint(oRec("Kid")) then
		ksel = "SELECTED"
		else
		ksel = ""
		end if
		%><option value="<%=oRec("Kid")%>" <%=ksel%>>(<%=oRec("kkundenr")%>)&nbsp;<%=oRec("kkundenavn")%></option><% 
		oRec.movenext
		wend
		oRec.close
		%>
		</select>
		<input type="image" src="../ill/pillillexp_tp.gif" name="skiftkunde"><br>
		<font class=megetlillesort>Hvis der skiftes kunde nulstilles et evt. valgt datointerval, og datoen for sidste faktura bruges som interval startdato.</font></form>
		</td>
	</tr>
	</table>
	</div>
	<%end if%>
	
	
	
	<div style="position:absolute; left:16; top:472; height:150;">
	<table width=150 cellspacing=0 cellpadding=0 border=0><form name=beregn>
	<tr bgcolor="#5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td colspan=2 valign="top"><img src="../ill/tabel_top.gif" width="134" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td class=alt colspan=2><b>Beregn en timepris:</b></td>
	</tr>
	<tr bgcolor="#eff3ff">
		<td colspan=2 style="border-left:1px #003399 solid; padding-left:5px; padding-top:2px;">Bel�b: <input type="text" name="beregn_belob" id="beregn_belob" value="0" size="5"> : </td>
		<td colspan=2 style="border-right:1px #003399 solid; padding-right:5px; padding-top:2px;">Timer: <input type="text" name="beregn_timer" id="beregn_timer" value="0" size="5"></td>
	</tr>
	<tr bgcolor="#eff3ff">
		<td colspan=4 style="border-left:1px #003399 solid; border-right:1px #003399 solid;  border-bottom:1px #003399 solid; padding-left:5px; padding-right:5px; padding-bottom:3px">
			<input type="button" name="beregn" id="beregn" value=" = " onClick="beregntimepris()" style="font-size:9px;"> <input type="text" name="beregn_tp" id="beregntp" value="0" style="width:99;">
		</td>
	</tr></form></table>								
	</div>
	
	
	<%
	if func <> "red" then
	id = 0 'S�ttes s� dato2 virker!
	else
	id = id
	end if
	%>	
	<!--#include file="inc/dato2.asp"-->
	<form name=main id=main action="fak.asp?menu=stat_fak&func=<%=dbfunc%>" method="post">
			<%
			'if request("FM_usedatointerval") = "1" then
			'*** Bruger altid datointerval ****
			usedt_ival = 1
			'else
			'usedt_ival = 0
			'end if
			%>
			<input type="hidden" name="FM_usedatointerval" value="<%=usedt_ival%>">
			<input type="hidden" name="FM_start_dag_ival" value="<%=request("FM_start_dag")%>">
			<input type="hidden" name="FM_start_mrd_ival" value="<%=request("FM_start_mrd")%>">
			<input type="hidden" name="FM_start_aar_ival" value="<%=request("FM_start_aar")%>">
			<input type="hidden" name="FM_slut_dag_ival" value="<%=request("FM_slut_dag")%>">
			<input type="hidden" name="FM_slut_mrd_ival" value="<%=request("FM_slut_mrd")%>">
			<input type="hidden" name="FM_slut_aar_ival" value="<%=request("FM_slut_aar")%>">
			
			<input type="hidden" name="FM_fakint_ival" value="<%=trim(request("FM_fakint"))%>">
			
	
	<input type="hidden" name="FM_faktype" value="<%=request("faktype")%>">
	<input type="hidden" name="FM_jobnavn" value="<%=strJobnavn%>">
	<input type="hidden" name="FM_job" value="<%=Request("FM_job")%>">
	<input type="hidden" name="id" value="<%=id%>">
	<input type="hidden" name="jobid" value="<%=jobid%>">
	
	<div id="modtager" style="position:absolute; left:<%=pleft%>; top:<%=ptop+50%>; visibility:visible;">
	<table cellspacing="0" cellpadding="0" border="0" width="370" bgcolor="#EFF3FF">
	<tr bgcolor="#5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td valign="top"><img src="../ill/tabel_top.gif" width="354" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td class=alt><b>Modtager:</b></td>
	</tr>
	<tr>
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td valign=top>
		<br>
		<%
		strSQL = "SELECT Kid, kkundenavn, kkundenr, adresse, postnr, city, land, telefon, cvr FROM kunder WHERE Kid =" & intjobknr  		
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then 
			intKnr = oRec("kkundenr")
			strKnavn = oRec("kkundenavn")
			strKadr = oRec("adresse")
			strKpostnr = oRec("postnr")
			strBy = oRec("city")
			strLand = oRec("land")
			intKid = oRec("Kid")
			intCVR = oRec("cvr")
			intTlf = oRec("telefon")
			'attA = oRec("kontaktpers1")
			'attB = oRec("kontaktpers2")
			'attC = oRec("kontaktpers3")
			'attD = oRec("kontaktpers4")
			'attE = oRec("kontaktpers5")
		end if
		
		oRec.close
		%>
		<input type="hidden" name="FM_Kid" value="<%=intKid%>">
		<table cellpadding=0 cellspacing=0 border=0 width=100%>
		<tr>
			<td colspan=2><b>Adresse:</b></td>
		</tr>
		<tr>
			<td colspan=2 class=lille>Kundenr: <%=intKnr%></td>
		</tr>
		<tr>
			<td colspan=2><%=strKnavn%></td>
		</tr>
		<tr>	
			<td colspan=2><%=strKadr%></td>
		</tr>
		<tr>	
			<td colspan=2><%=strKpostnr%>&nbsp;&nbsp;<%=strBy%></td>
		</tr>
		<tr>	
			<td width=200><%=strLand%></td><td ><input type="checkbox" name="FM_land_vis" value="on">Vis land p� fak.
		</td>
		<tr>	
			<td>Att:&nbsp;<select name="FM_att" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid; width:200;">
					<%
				strSQLkpersCount = "SELECT id, navn FROM kontaktpers WHERE kundeid = "& intKid &" ORDER BY id"
				oRec2.open strSQLkpersCount, oConn, 3
					
					while not oRec2.EOF
						if strAtt = oRec2("id") then
						attSel = "SELECTED"
						else
						attSel = ""
						end if%>
						<option value="<%=oRec2("id")%>" <%=attSel%>><%=oRec2("navn")%></option>
						<%
					oRec2.movenext
					wend
					oRec2.close
					%>
				<%if strAtt = "991" then
				selth1 = "SELECTED"
				else
				selth1 = ""
				end if%>	
				<option value="991" <%=selth1%>>Den �konomi ansvarlige</option>
				<%if strAtt = "992" then
				selth2 = "SELECTED"
				else
				selth2 = ""
				end if%>	
				<option value="992" <%=selth2%>>Regnskabs afd.</option>
				<%if strAtt = "993" then
				selth3 = "SELECTED"
				else
				selth3 = ""
				end if%>	
				<option value="993" <%=selth3%>>Administrationen</option>
				</select></td>
				<td ><input type="checkbox" name="FM_att_vis" value="on" checked>Vis att. p� fak.</td>
		</tr>
		<tr><td><br>Tlf: <%=intTlf%></td><td><br><input type="checkbox" name="FM_tlf_vis" value="on">Vis tlf. p� fak.<br></td></tr>
		<tr><td><b>Cvr/SE nr:</b> <%=intCVR%></td><td ><input type="checkbox" name="FM_cvr_vis" value="on">Vis cvr p� fak.</td></tr>
		<tr><td colspan=2><b>Kundekonto:</b>
		<select name="FM_kundekonto" style="width:200;">
		<option value="0">(0)&nbsp;&nbsp;Ingen konto</option>
		<%
			strSQL = "SELECT kontonr, navn, id, kid FROM kontoplan ORDER BY kontonr, navn"
			oRec.open strSQL, oConn, 3 
			while not oRec.EOF
			
			if func = "red" then
				if intKonto = oRec("kontonr") then
				selkon = "SELECTED"
				else
				selkon = ""
				end if
			else
				if intKid = oRec("kid") then
				selkon = "SELECTED"
				else
				selkon = ""
				end if
			end if
			
			%>
			<option value="<%=oRec("kontonr")%>" <%=selkon%>>(<%=oRec("kontonr")%>)&nbsp;&nbsp;<%=oRec("navn")%></option>
			<%
			oRec.movenext
			Wend 
			oRec.close
		%>
		</select>&nbsp;&nbsp;<select name="FM_debkre">
				<%
					if request("faktype") <> 1 then
					selK = ""
					selD = "SELECTED"
					else
					selK = "SELECTED"
					selD = ""
					end if
				
				%>
		<option value="k" <%=selK%>>Krediter</option>
		<option value="d" <%=selD%>>Debiter</option>
		</select></td></tr>
		</table>
		</td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<tr>
		<td colspan=3 style="border-right:1px #003399 solid; border-bottom:1px #003399 solid; border-left:1px #003399 solid;">&nbsp;</td>
	</tr>
	</table>
	</div>
	
	<!--- afsender af faktura --->
	<div id="afsender" style="position:absolute; left:<%=pleft+400%>; top:<%=ptop+50%>;">
	<table cellspacing="0" cellpadding="0" border="0" width="300" bgcolor="#FFFFFF">
	<tr bgcolor="#5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td valign="top"><img src="../ill/tabel_top.gif" width="284" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td class=alt><b>Afsender:</b></td>
	</tr>
	<tr><td width="30" style="border-left:1px #003399 solid;">&nbsp;</td>
	<td>
		<%
		strSQL = "SELECT adresse, postnr, city, land, telefon, email, regnr, kkundenavn, kontonr, cvr, bank, swift, iban, kid FROM kunder WHERE useasfak = 1"
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then 
			yourbank = oRec("bank")
			yourRegnr = oRec("regnr")
			yourKontonr = oRec("kontonr")
			yourCVR = oRec("cvr")
			yourNavn = oRec("kkundenavn")
			yourAdr = oRec("adresse")
			yourPostnr = oRec("postnr")
			yourCity = oRec("city")
			yourLand = oRec("land")
			yourEmail = oRec("email")
			yourTlf = oRec("telefon")
			yourSwift = oRec("swift")
			yourIban = oRec("iban")
			afskid = oRec("kid")
		end if
		oRec.close
		%>
		
		<table cellpadding=0 cellspacing=0 border=0 width=100%>
		<tr>
			<td colspan=2><br><b>Adresse:</b></td>
		</tr>
		
		<tr>
			<td colspan=2><%=yourNavn%></td>
		</tr>
		<tr>
			<td colspan=2><%=yourAdr%></td>
		</tr>
		<tr>
			<td colspan=2><%=yourPostnr%>&nbsp;&nbsp;<%=yourCity%></td>
		</tr>
		<tr>
			<td width=120 colspan=2><%=yourLand%></td>
		</tr>
		<tr>
			<td><br>Tlf: <%=yourTlf%></td><td><br><input type="checkbox" name="FM_ytlf_vis" value="on" CHECKED>Vis tlf. p� fak.</td>
		</tr>
		<tr>
			<td>Email: <%=yourEmail%></td><td ><input type="checkbox" name="FM_yemail_vis" value="on">Vis email p� fak.</td>
		</tr>
		<tr>
			<td colspan=2><br><b>Bank information:</b></td>
		</tr>
		<tr>
			<td colspan=2>Bank: <%=yourBank%></td>
		</tr>
		<tr>
			<td colspan=2>Reg. og kontonr: <%=yourRegnr%> - <%=yourKontonr%></td>
		</tr>
		<tr>
			<td>Swift: <%=yourSwift%></td><td ><input type="checkbox" name="FM_yswift_vis" value="on">&nbsp;Vis p� fak.</td>
		</tr>
		<tr>
			<td>Iban: <%=yourIban%> </td><td ><input type="checkbox" name="FM_yiban_vis" value="on">&nbsp;Vis p� fak.</td>
		</tr>
		<tr>
			<td><b>Cvr/SE nr:</b> <%=yourCVR%></td><td><input type="checkbox" name="FM_ycvr_vis" value="on" CHECKED>&nbsp;Vis p� fak.</td>
		</tr>
		</table>
		
		</td><td width="30" style="border-right:1px #003399 solid;">&nbsp;</td>
		</tr>
		<tr><td width="30" style="border-left:1px #003399 solid;">&nbsp;</td>
		<td><b>Modkonto:</b>&nbsp;
		<select name="FM_modkonto" style="width:200px;">
		<option value="0">(0)&nbsp;&nbsp;Ingen konto</option>
		<%
				strSQL = "SELECT kontonr, navn, id, kid FROM kontoplan WHERE kid = "& afskid &" ORDER BY kontonr, navn"
				
				oRec.open strSQL, oConn, 3 
				while not oRec.EOF 
				if intModKonto = oRec("kontonr") then
				selkon = "SELECTED"
				else
				selkon = ""
				end if
				%>
				<option value="<%=oRec("kontonr")%>" <%=selkon%>>(<%=oRec("kontonr")%>)&nbsp;&nbsp;<%=oRec("navn")%></option>
				<%
				oRec.movenext
				Wend 
				oRec.close
		%>
		</select>
		</td><td width="30" style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<tr>
		<td colspan=3 style="border-right:1px #003399 solid; border-bottom:1px #003399 solid; border-left:1px #003399 solid;">&nbsp;</td>
	</tr>
	</table>
	</div>
	<!-- afsender slut -->
	
	
	<div id="aktiviteter" style="position:absolute; left:<%=pleft%>; top:<%=ptop+420%>; visibility:visible; z-index:100;">
	<table cellspacing="0" cellpadding="0" border="0" width="700">
	<tr bgcolor="#5582D2">
		<td width="8" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td colspan=4 style="border-top:1px #003399 solid;" class=alt><b>Faktura og job info.</b></td>
		<td align=right valign=top width="8"><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td style="border-left:1px #003399 solid;" rowspan="7">&nbsp;</td>
		<td colspan=4 height=5><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td style="border-right:1px #003399 solid;" rowspan="7">&nbsp;</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan=2 valign=top style="padding-left:10;" class=lille><br><u>Joboplysninger:</u><br>
		(<%=jobnr%>)&nbsp;&nbsp;<%=strJobnavn%><br>
		<u>Forkalkuleret timer:</u><br>
		Fakturerbare: <%=intBudgettimer%> <br>
		Ikke fakturerbare: <%=intIkkeBtimer%> <br>
		<%if fastpris = 1 then
		jtype = "Fastpris"
		else
		jtype = "L�bende timer"
		end if%>
		Jobtype: <font color=red><%=jtype%> </font><br>
		Budget/Pris p� job:	<%=formatcurrency(intJobTpris, 2)%><br>
		<u>Jobansvarlig(e):</u><br>
		<%
		
		strSQL2 = "SELECT mnavn, mnr, mid FROM medarbejdere WHERE  mid = "& jobans1
		
		'Response.write strSQL2
		'Response.flush
		oRec2.open strSQL2, oConn, 3 
		if not oRec2.EOF then
		jobans1txt = oRec2("mnavn") & " ("& oRec2("mnr") &")"
		end if
		oRec2.close
		
		strSQL2 = "SELECT mnavn, mnr, mid FROM medarbejdere WHERE  mid = "& jobans2
		oRec2.open strSQL2, oConn, 3 
		if not oRec2.EOF then
		jobans2txt = oRec2("mnavn") & " ("& oRec2("mnr") &")"
		end if
		oRec2.close
		%>
		
		<%=jobans1txt%><br>
		<%=jobans2txt%><br>&nbsp;
		</td><td colspan=2 valign=top style="padding-left:10;"><br>
		<b>Faktureret p� dette job indtil dagsdato: </b><br>
		
		<%
		if len(faktureretBeloeb) <> 0 then
		faktureretBeloeb = faktureretBeloeb
		else
		faktureretBeloeb = 0
		end if
		%>
		Bel�b: <%=formatcurrency(faktureretBeloeb, 2)%><br>
		Timer: <%=timeforbrug%>
		
		<%if fastpris = 1 then%>
		<br><br>
		<b>Restbel�b til fakturering:</b><br>
		<%
		rastfak = (intJobTpris - faktureretBeloeb)
		%><font color=darkred>
		<b><%=formatcurrency(rastfak, 2)%></b>
		</font>
		<input type="hidden" name="restbelob" id="restbelob" value="<%=rastfak%>">
		<%else%>
		<input type="hidden" name="restbelob" id="restbelob" value="100000000">
		<%end if%>
		<br>&nbsp;
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan=4 style="padding-left:10;">
		<b>Type:</b>&nbsp;&nbsp;
		<%
		'** Faktype ***
		select case request("faktype")
		case "0"
		strFaktypeNavn = "Faktura"
		case "1"
		strFaktypeNavn = "Kreditnota"
		case "2"
		strFaktypeNavn = "Rykker"
		case else
		strFaktypeNavn = "Faktura"
		end select
		%>
		
		<%=strFaktypeNavn%>
		
		&nbsp;&nbsp;
		
		<b><%=strFaktypeNavn%> nr:</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="text" name="FM_faknr" value="<%=strFaknr%>" size="14" maxlength="10">
		<input type="hidden" name="FM_faknr_opr" value="<%=strFaknr%>">
		
		<%if len(trim(strNote)) <> 0 then%>
		<br><%=strNote%><br>
		<%end if%>
		
		<%if request("faktype") = "2" then%>
		<br>
		<b>Bilags nr:</b><img src="ill/blank.gif" width="23" height="1" alt="" border="0">&nbsp;<input type="text" name="bilagsnrkredit" value="0" size="14" maxlength="10">
		<br>N�r denne rykker oprettes, krediteres samtidig den oprindelige faktura.<br> Krediteringen bliver gemt med det angivne bilagsnummer.<br>
		<%end if%>
		<font color="darkred"><%=strNote1%></font><br>&nbsp;
		</td>
	</tr>
	
	<tr bgcolor="#FFFFFF">
		<td colspan=2 style="padding-left:10;" valign=top width=400>
		<%if request("faktype") = 0 then%>
			<b>Kladde / Endelig</b><br>
			<%if intBetalt <> 0 then%>
			Godkendt <input type="radio" name="FM_betalt" value="1" checked>
			</td><td colspan=2 style="padding-left:10;" valign=top width=300>&nbsp;
			<%else%>
			<input type="radio" name="FM_betalt" value="0" checked>Opret som <b>kladde</b>.&nbsp;&nbsp; <input type="radio" name="FM_betalt" value="1"> eller som <font color="limegreen"><b>godkend</b></font> som endelig.
			</td><td colspan=2 style="padding-left:10;" valign=top width=300><b>Note 1:</b>
			N�r en faktura godkendes som endelig, bliver der oprettet en tilh�rende postering p� de ovenst�ende valgte konti.
			N�r fakturaen er godkendt kan den <u>ikke</u> l�ngere slettes eller opdateres.</font>
			<%end if%>
		<%else%>&nbsp; 
		</td><td colspan=2 style="padding-left:10;" valign=top width=300>&nbsp;
		<input type="hidden" name="FM_betalt" value="1">
		<%end if%>
		</td>
	</tr>
	
	
	
	<%
	if len(lastFakdato) <> 0 then
	lastFakdato = lastFakdato
	else
	lastFakdato = "2001/1/1"
	end if
	
	'***************************************************************************
	'*** S�tter datointerval / periode afgr�nsning eller bruger lastFakdato ****
	'*** Bruger ALTID datointerval ****
	
	
	'if cint(usedt_ival) = 1 then
	
		'** Periodeinterval brugt.
		'** Start dato
		call datofindes(request("FM_start_dag"),request("FM_start_mrd"),request("FM_start_aar"))
		stdato = request("FM_start_aar")&"/"&request("FM_start_mrd")&"/"&dagparset
		showStDato = dagparset&"/"&request("FM_start_mrd")&"/"&request("FM_start_aar")
		
		'** Slut dato
		call datofindes(request("FM_slut_dag"),request("FM_slut_mrd"),request("FM_slut_aar"))
		slutdato = request("FM_slut_aar")&"/"&request("FM_slut_mrd")&"/"&dagparset
		showSlutDato = dagparset&"/"&request("FM_slut_mrd")&"/"&request("FM_slut_aar") 
		useLastFakdato = 0
	
	'else
		
		'stdato_temp = dateadd("d", 1, lastFakdato)
		'stdato = year(stdato_temp)&"/"&month(stdato_temp)&"/"&day(stdato_temp)
		'slutdato = "2014/1/1"
		
		'showStDato = stdato 'lastFakdato
		'showSlutDato = slutdato 
		'useLastFakdato = 1
		
	'end if
	
	
	'**** Hvis startdato ligger f�r sidste fakdato s�ttes stdato = sidste fakdato ****
	if cdate(showStDato) > cdate(lastFakdato) then 'AND cint(usedt_ival) = 1 
		stdatoKri = stdato
		useLastFakdato = 0
	else
		stdatoKri_temp = dateadd("d", 1, lastFakdato)
		stdatoKri = year(stdatoKri_temp)&"/"&month(stdatoKri_temp)&"/"&day(stdatoKri_temp)
		showStDato = stdatoKri_temp 'lastFakdato
		useLastFakdato = 1
	end if
	'******************************************************
	
	if func <> "red" then%>
	<tr bgcolor="#FFFFFF">
		<td colspan=2 style="padding-left:10; padding-bottom:10; padding-top:10;" valign=top><b>Periode afgr�nsning:</b><br>
		<%=formatdatetime(showStDato, 1)%>
		<%if useLastFakdato = 1 then%>
		&nbsp;<font color="red">(Korrigeret til dagen efter sidste faktura dato)</font>
		<%end if%>
		
		<%if datediff("d", lastFakdato, showStDato) > 1 then%>
		&nbsp;<font color="red">(Se note 2)</font>
		<%end if%>
		
		<br>
		til <%=formatdatetime(showSlutDato, 1)%></td>
		<td colspan=2 style="padding-top:10; padding-left:10;" valign=top><b>Note 2:</b> Sidste faktura dato p� dette job er: <b><%=formatdatetime(lastFakdato, 1)%></b><br>
		Korrekt startdato p� periodeafgr�nsning vil v�re dagen efter sidste faktura dato.</td>
	</tr>
	<%else%>
	<tr bgcolor="#FFFFFF">
		<td colspan=4 style="padding-left:10;" valign=top>&nbsp;</td>
	</tr>
	<%end if%>
	
	<tr bgcolor="#FFFFFF">
		<td colspan=2 style="padding-left:10;" valign=top>
		<b><%=strFaktypeNavn%> dato:</b>&nbsp;&nbsp;
		
		<%'** Bruger afg. dato fra interval som faktura dato-
		if cint(usedt_ival) = 1 and func <> "red" then
		strDag = request("FM_slut_dag")
		strMrd = request("FM_slut_mrd")
		strAar = request("FM_slut_aar")
		end if
		
		'*** findes datoen?? ***
		call dag()
		
		
		Response.write formatdatetime(strDag &"/"& strMrd &"/"& strAar, 1)
		%>
		<input type="hidden" name="FM_start_dag" id="FM_start_dag" value="<%=strDag%>">
		<input type="hidden" name="FM_start_mrd" id="FM_start_mrd" value="<%=strMrd%>">
		<input type="hidden" name="FM_start_aar" id="FM_start_aar" value="<%=strAar%>">
		<br>&nbsp;</td>
		<td colspan=2 style="padding-left:10;" valign=top width=300><b>Note 3:</b> 
		Fakturadatoen f�lger altid slutdatoen p� det valgte datointerval.
		<br>&nbsp;
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan=4 style="padding-left:10;" valign=top><b>Enhed angives som:</b><br>
		<input type="radio" name="FM_enheds_ang" value="0" checked> Timepris&nbsp;&nbsp; 
		<input type="radio" name="FM_enheds_ang" value="1"> Stk. pris &nbsp;&nbsp;
		<input type="radio" name="FM_enheds_ang" value="2"> Enhedspris.
		</td>
	</tr>
	
	<tr bgcolor="#FFFFFF">
		<td valign="top" colspan=6  style="border-bottom:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	</table>
	
	<table cellspacing="0" cellpadding="0" border="0" width="700" bgcolor="#FFFFFF">
	<%
	'***********************************************************************************************
	'*** Henter aktiviteter ************************************************************************
	'***********************************************************************************************
	if func = "red" then
		
		isTPused = "#"
		
		'******************* Henter oprettede aktiviteter *******************************
		strSQL = "SELECT faktura_det.id, aktid, aktiviteter.navn FROM faktura_det LEFT JOIN aktiviteter ON (aktiviteter.id = aktid) WHERE fakid = "& intFakid &" GROUP BY aktid" 
		oRec.open strSQL, oConn, 3
		x = 0
		While not oRec.EOF 
		
			Redim preserve thisaktnavn(x)
			thisaktnavn(x) = oRec("navn") 'oRec("beskrivelse")
			Redim preserve thisaktid(x)
			thisaktid(x) = oRec("aktid")
			
			
		x = x + 1
		
		oRec.movenext
		wend 
		oRec.close
	
	else
			'**** Sidste faktura tidspunkt ****
			if len(strfaktidspkt) > 0 then
			LastFakTid = datepart("h", strfaktidspkt)&":"& datepart("n", strfaktidspkt)&":"& datepart("s", strfaktidspkt)
			else
			LastFakTid =  "00:00:02"
			end if
			
			
			'*** Henter aktiviter og finder registrede timer siden sidste faktura dato. (periodeafgr�nsning) *** 
			strSQL = "SELECT job.fastpris, job.jobtpris, job.budgettimer, aktstatus, aktiviteter.job, aktiviteter.id AS aid, aktiviteter.navn AS anavn FROM aktiviteter LEFT JOIN job ON (job.id = aktiviteter.job) WHERE aktiviteter.job = "& jobid &" AND aktiviteter.aktstatus = 1 AND aktiviteter.fakturerbar = 1 ORDER BY aktiviteter.id"		
			oRec.open strSQL, oConn, 3
			
			x = 0
			lastAktId = 0
			timerThisT = 0
			timerThisD = 0
			strKom = "Netto Kontant 8 dage."
			totaltimer = 0
			lastaktid = 0
			
			
			while not oRec.EOF
				if cint(lastaktid) <> cint(oRec("aid")) then
						
						Redim preserve thisaktid(x)
						Redim preserve usefastpris(x)
						thisaktid(x) = oRec("aid")
						usefastpris(x) = oRec("fastpris")
						
						'** timepris p� fastpris job ***
							if cint(oRec("fastpris")) = 1 then
								akttimepris = oRec("jobtpris")/oRec("budgettimer")
							else
								akttimepris = 0
							end if
							
						
						'** antal registerede timer pr. akt.
						strSQL2 = "SELECT sum(timer.timer) AS timerthis FROM timer WHERE taktivitetid =" & oRec("aid") &" AND tfaktim = 1 AND (Tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"')"   ' OR Tdato = '" & stdatoKri &"' AND timer.tidspunkt > '"& LastFakTid &"')"		
						oRec2.open strSQL2, oConn, 3
						
						if not oRec2.EOF then
							timerThisD = oRec2("timerthis")
						end if
						oRec2.close
						
							if len(timerThisD) <> 0 then
							timerThisD = timerThisD 
							else
							timerThisD = 0
							end if
							
						
						
					Redim preserve thisaktnavn(x)
					thisaktnavn(x) = oRec("anavn") 
					Redim preserve thisAktTimer(x)
					thisAktTimer(x) = timerThisD
					'Redim preserve thisTimePris(x)
					'thisTimePris(x) = akttimepris
					Redim preserve thisAktBeloeb(x)
					thisAktBeloeb(x) = timerThisD * akttimepris
					
					
					
					timerThisD = 0
					thisBeloebD = 0
					
					
				x = x + 1		
				end if
			
			lastaktid = oRec("aid")		
			oRec.movenext
			wend
			oRec.close
	
	end if
	'***********************************************************************************
	
	
	
	
	
	'********************* Udskriver aktiviterer *********************
	'aktiswritten = "#0#,"
	for x = 0 to x - 1 
	
	%>
	<input type="hidden" name="sumaktbesk_<%=x%>" id="sumaktbesk_<%=x%>" value="<%=thisaktnavn(x)%>">
	<%
		'if instr(aktiswritten, "#"&thisaktid(x)&"#") = 0 then
		n = 1
		a = 0
		sa = 0
	
					if func <> "red" then
							'**** projektgrupper p� job ******************************************
							'Tjekker at der den projgp der findes p� akt ogs� findes p� jobbet. 
							'G�lder kun job oprettet f�r 5/11-2004, da der her er lavet en 
							'funktion der tjekker det samme n�r et job redigeres. 
							'*********************************************************************
							
								jobPgstring = ""
							
								strSQL = "SELECT id, projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM job WHERE id = "& jobid
								oRec2.open strSQL, oConn, 3
									if not oRec2.EOF then
									
									strProjektgr1 = oRec2("projektgruppe1")
									strProjektgr2 = oRec2("projektgruppe2")
									strProjektgr3 = oRec2("projektgruppe3")
									strProjektgr4 = oRec2("projektgruppe4")
									strProjektgr5 = oRec2("projektgruppe5")
									strProjektgr6 = oRec2("projektgruppe6")
									strProjektgr7 = oRec2("projektgruppe7")
									strProjektgr8 = oRec2("projektgruppe8")
									strProjektgr9 = oRec2("projektgruppe9")
									strProjektgr10 = oRec2("projektgruppe10")
									
									jobPgstring = "#"&strProjektgr1&"#,#"&strProjektgr2&"#,#"&strProjektgr3&"#,#"&strProjektgr4&"#,#"&strProjektgr5&"#,#"&strProjektgr6&"#,#"&strProjektgr7&"#,#"&strProjektgr8&"#,#"&strProjektgr9&"#,#"&strProjektgr10&"#"
									
									end if
								oRec2.close
							
							
							
							'*** Medarbejder udspecificering ***
							strSQL3 = "SELECT projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM aktiviteter WHERE id = "& thisaktid(x)
							oRec3.open strSQL3, oConn, 3
							if not oRec3.EOF then
							
								gp1 = oRec3("projektgruppe1")
								gp2 = oRec3("projektgruppe2")
								gp3 = oRec3("projektgruppe3")
								gp4 = oRec3("projektgruppe4")
								gp5 = oRec3("projektgruppe5")
								gp6 = oRec3("projektgruppe6")
								gp7 = oRec3("projektgruppe7")
								gp8 = oRec3("projektgruppe8")
								gp9 = oRec3("projektgruppe9")
								gp10 = oRec3("projektgruppe10")
							
							end if
							oRec3.Close 
							
							'*** Findes gruppen ogs� p� jobbet?? ****
							if len(gp1) <> 0 AND instr(jobPgstring, "#"&gp1&"#") > 0 then
							gp1 = gp1
							else
							gp1 = 1
							end if
							
							if len(gp2) <> 0 AND instr(jobPgstring, "#"&gp2&"#") > 0 then
							gp2 = gp2
							else
							gp2 = 1
							end if
							
							if len(gp3) <> 0 AND instr(jobPgstring, "#"&gp3&"#") > 0 then
							gp3 = gp3
							else
							gp3 = 1
							end if
							
							if len(gp4) <> 0 AND instr(jobPgstring, "#"&gp4&"#") > 0 then
							gp4 = gp4
							else
							gp4 = 1
							end if
							
							if len(gp5) <> 0 AND instr(jobPgstring, "#"&gp5&"#") > 0 then
							gp5 = gp5
							else
							gp5 = 1
							end if
							
							if len(gp6) <> 0 AND instr(jobPgstring, "#"&gp6&"#") > 0 then
							gp6 = gp6
							else
							gp6 = 1
							end if
							
							if len(gp7) <> 0 AND instr(jobPgstring, "#"&gp7&"#") > 0 then
							gp7 = gp7
							else
							gp7 = 1
							end if
							
							if len(gp8) <> 0 AND instr(jobPgstring, "#"&gp8&"#") > 0 then
							gp8 = gp8
							else
							gp8 = 1
							end if
							
							if len(gp9) <> 0 AND instr(jobPgstring, "#"&gp9&"#") > 0 then
							gp9 = gp9
							else
							gp9 = 1
							end if
							
							if len(gp10) <> 0 AND instr(jobPgstring, "#"&gp10&"#") > 0 then
							gp10 = gp10
							else
							gp10 = 1
							end if	
							
							
					end if
					
					if thisaktid(x) <> 0 then%>
					<tr bgcolor="#d6dff5"><td colspan=6><img src="../ill/blank.gif" width="1" height="20" alt="" border="0"><br></td></tr>
					<tr bgcolor="#66CC33">
						<td style="border-left:1px #8caae6 solid; border-top:1px #8caae6 solid;">&nbsp;</td>
						<td width=175 style="border-top:1px #8caae6 solid;"><br><font class=megetlillesort>Vis p� fak. &nbsp;Fak | (v) Vent</td>
						<td width=335 style="border-top:1px #8caae6 solid;"><br>&nbsp;&nbsp;&nbsp;<b><%=thisaktnavn(x)%></b><font class=megetlillesort>&nbsp;&nbsp;&nbsp;Medarbejder(e)&nbsp;&nbsp;</td>
						<td width=105 style="border-top:1px #8caae6 solid;" align=center><br><font class=megetlillesort>Timepris</td>
						<td style="border-top:1px #8caae6 solid;" align=right><br><font class=megetlillesort>Bel�b&nbsp;&nbsp;&nbsp;</td>
						<td style="border-top:1px #8caae6 solid; border-right:1px #8caae6 solid;"><br><font class=megetlillesort>&nbsp;&nbsp;Afrund</td>
					</tr>
					<!--<tr bgcolor="#66CC33"><td colspan=6 style="border-bottom:1px #003399 solid; border-left:1px #003399 solid; border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="2" alt="" border="0"><br></td></tr>-->
			
					<%end if
					
					'***************************************************
					'Faktimer, Venter og Bel�b p� medarbejdere 
					'***************************************************
					
					faktot = 0
					medarbstrpris = ""
					
					
					if func = "red" then
						
						strSQL10 = "SELECT faktura_det.id, antal, faktura_det.beskrivelse, aktpris, faktura_det.showonfak AS akt_showonfak, "_
						&" faktura_det.enhedspris, "_
						&" faktura_det.aktid, "_
						&" fak, venter, tekst, fak_med_spec.enhedspris, beloeb, mid, fak_med_spec.showonfak AS showonfak "_
						&" FROM faktura_det "_
						& "LEFT JOIN fak_med_spec ON (fak_med_spec.fakid = faktura_det.fakid AND fak_med_spec.aktid = faktura_det.aktid)"_
						&" WHERE faktura_det.fakid = "& intFakid &" AND faktura_det.aktid = "& thisaktid(x) &" GROUP BY mid"
						
						
					else
						
						usedmedabId = " Tmnr = 0 AND "
						
						strSQL10 = "SELECT DISTINCT(medarbejderid), tmnr, mnavn, mid, timeprisalt, medarbejdertype FROM progrupperelationer, medarbejdere LEFT JOIN "_
						&" timer ON (Tmnr = medarbejderid AND taktivitetid ="& thisaktid(x) &" AND tfaktim = 1 AND Tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"')"_  
						&" LEFT JOIN timepriser ON (aktid = "& thisaktid(x) &" AND medarbid = medarbejderid) WHERE mid = medarbejderid AND (projektgruppeid = "& gp1 &" OR projektgruppeid = "& gp2 &" OR projektgruppeid = "& gp3 &" OR projektgruppeid = "& gp4 &" OR projektgruppeid = "& gp5 &" OR projektgruppeid = "& gp6 &" OR projektgruppeid = "& gp7 &" OR projektgruppeid = "& gp8 &" OR projektgruppeid = "& gp9 &" OR projektgruppeid = "& gp10 &") GROUP BY medarbejderid ORDER BY projektgruppeid"
						
						'Response.write strSQL10
						'Response.flush
						
					end if
					
					oRec3.open strSQL10, oConn, 3
					
					while not oRec3.EOF
								
								if func = "red" then
									
									fak = oRec3("fak")
									venter = oRec3("venter")
									'Response.write "fak "& fak  & "venter  " & venter & "<br>"
									medarbejderTimepris = oRec3("enhedspris")
									medarbejderBeloeb = oRec3("beloeb")
									txt = oRec3("tekst")
									medarbid = oRec3("mid")		
									showonfak = oRec3("showonfak")
									akt_showonfak = 0'oRec3("akt_showonfak")
									
									
									if len(medarbejderTimepris) <> 0 then
									medarbejderTimepris = SQLBlessDOT(formatnumber(medarbejderTimepris, 2))
									else
									medarbejderTimepris = SQLBlessDot(formatnumber(0, 2))
									end if
								
								else
										
										'***************************************************************
										'Timer medarbejder 
										'***************************************************************
										
										strSQL2 = "SELECT sum(timer) AS timer FROM timer WHERE Tmnr = "&oRec3("medarbejderid")&" AND taktivitetid ="& thisaktid(x) &" AND tfaktim = 1 AND (Tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"')" 'OR Tdato = '" & stdatoKri &"' AND timer.tidspunkt > '"& LastFakTid &"'
										oRec2.open strSQL2, oConn, 3 
										
										if not oRec2.EOF then 
											medarbejderTimer = oRec2("timer")
										end if
										oRec2.close
										
										if len(medarbejderTimer) <> 0 then
											medarbejderTimer = medarbejderTimer
										else
											medarbejderTimer = 0
										end if
										
										usedmedabId = usedmedabId & " Tmnr <> "& oRec3("medarbejderid") &" AND "
										
										'***************************************************************
										'Timepris pr. medarbejder ***
										'***************************************************************
										%>
										<!--#include file="inc/fak_inc_mtimepris_opret.asp" -->
										<%
										
										'***************************************************************
										'Fak, Venter og Bel�b
										'***************************************************************
										call fakventerbelob(thisaktid(x), oRec3("mid"), oRec3("mnavn"), medarbejderTimer, medarbejderTimepris)
										
											
								end if
								
							
							
							'**************************************************************************
							'Medarbejder timepris samme som forrige?? 
							'Kalder sum-aktivitet
							'**************************************************************************
							
								if instr(medarbstrpris, "#"&MedarbejderTimepris&"#") = 0 then
									a = a + 1
									Redim preserve aval(a)
									aval(a) = MedarbejderTimepris
									call subtotalakt(x, a)
									sa = sa + 1
								end if
							
							'***
							
							
							'** Henter tom sum-akt ***
							call subtotalakt(x, -(n))
							
							
							
							
							'***** Medarbejder timer p� akt. timer skt. pris og oms�tning  ************	
							%>	
							<tr>
								<td style="border-left:1px #86B5E4 solid;">&nbsp;</td>
								<td><input type="hidden" name="FM_mid_<%=n%>_<%=x%>" value="<%=medarbid%>">
								<%
								if showonfak = 1 then
								chkshow = "CHECKED"
								else
								chkshow = ""
								end if
								
								if instr(medarbstrpris, "#"&MedarbejderTimepris&"#") = 0 then
								a = a
								else
									for intcounter = 0 to a - 1
										
										if MedarbejderTimepris = aval(intcounter) then
										a = intcounter
										end if
										
									next
								end if
								
								
								 
								%>
								<input type="hidden" name="FM_m_aval_opr_<%=n%>_<%=x%>" id="FM_m_aval_opr_<%=n%>_<%=x%>" value="<%=a%>">
								<input type="hidden" name="FM_m_aval_<%=n%>_<%=x%>" id="FM_m_aval_<%=n%>_<%=x%>" value="<%=a%>">
								
								<input type="checkbox" name="FM_show_medspec_<%=n%>_<%=x%>" <%=chkshow%> id="FM_show_medspec_<%=n%>_<%=x%>" value="show">
								<!--<input type="text" name="FM_m_fak_<=n%>_<=x%>" id="FM_m_fak_<=n%>_<=x%>" onKeyup="offsetfak('<=x%>','<=n%>'), tjektimer2('<=x%>','<=n%>'), updantaltimerprakt('<=x%>','<=n%>', 1)" value="<=fak%>" size=3 style="!border: 1px; background-color: #FFFFFF; border-color: limegreen; border-style: solid; font-size:9px">&nbsp;-->
								<input type="text" maxlength="<%=maxl%>" name="FM_m_fak_<%=n%>_<%=x%>" id="FM_m_fak_<%=n%>_<%=x%>" onKeyup="offsetfak('<%=x%>','<%=n%>'), tjektimer2('<%=x%>','<%=n%>')" value="<%=fak%>" size=4 style="!border: 1px; background-color: #FFFFFF; border-color: limegreen; border-style: solid; font-size:9px">
								<input type="button" name="beregn" id="beregn" value="Calc" onClick="enhedsprismedarb('<%=x%>','<%=n%>')" style="font-size:8px">
								&nbsp;&nbsp;&nbsp;
								
								<input type="hidden" name="FM_m_fak_opr_<%=n%>_<%=x%>" id="FM_m_fak_opr_<%=n%>_<%=x%>" value="<%=fak%>">
								
								
								<%if func <> "red" then%>
								<font class=megetlillesort>(<%=showventer%>)</font>
								<input type="text" name="FM_m_vent_<%=n%>_<%=x%>" id="FM_m_vent_<%=n%>_<%=x%>" onKeyup="offsetventer('<%=x%>','<%=n%>'), tjektimer4('<%=x%>','<%=n%>')" value="<%=venter%>" size=1 style="!border: 1px; background-color: #FFFFFF; border-color: silver; border-style: solid; font-size:9px">
								<input type="hidden" name="FM_m_vent_opr_<%=n%>_<%=x%>" id="FM_m_vent_opr_<%=n%>_<%=x%>" value="<%=fak%>">
								<%else%>
							    <input type="text" name="FM_m_vent_<%=n%>_<%=x%>" id="FM_m_vent_<%=n%>_<%=x%>" onKeyup="offsetventer('<%=x%>','<%=n%>'), tjektimer4('<%=x%>','<%=n%>')" value="<%=venter%>" size=1 style="!border: 1px; background-color: #FFFFFF; border-color: silver; border-style: solid; font-size:9px">
								<input type="hidden" name="FM_m_vent_opr_<%=n%>_<%=x%>" id="FM_m_vent_opr_<%=n%>_<%=x%>" value="<%=venter%>">
								<%end if%>
								
								
								</td>
								<td><input type="text" name="FM_m_tekst_<%=n%>_<%=x%>" value="<%=txt%>" size=40 style="!border: 1px; background-color: #FFFFFF; border-color: silver; border-style: solid; font-size:9px">
								<%if func <> "red" then%>
								<%=tp%>
								<%end if%>
								</td>
								<td align=right style="padding-right:10;"><input type="hidden" name="FM_mtimepris_opr_<%=n%>_<%=x%>" id="FM_mtimepris_opr_<%=n%>_<%=x%>" value="<%=medarbejderTimepris%>">
								<input type="text" name="FM_mtimepris_<%=n%>_<%=x%>" id="FM_mtimepris_<%=n%>_<%=x%>" onKeyup="offsetmtp('<%=x%>','<%=n%>'), tjektimer3('<%=x%>','<%=n%>')" value="<%=medarbejderTimepris%>" size=5 style="!border: 1px; background-color: #FFFFFF; border-color: #999999; border-style: solid; font-size:9px">
								<input type="button" name="beregn" id="beregn" value="Calc" onClick="enhedsprismedarb('<%=x%>','<%=n%>')" style="font-size:8px">
								<!-- onFocus="higlightakt('<=x%>','<=n%>','1')" -->
								</td>
								<td style="padding-left:2;">
								<%
								if len(medarbejderBeloeb) <> 0 then
								mbelob = medarbejderBeloeb
								else
								mbelob = 0
								end if
								%>
								<div style="position:relative; left:5; top:0; width:65; border-bottom:1px #8caae6 solid; background-color:#ffffff; padding-right:3px;" align="right" name="medarbbelobdiv_<%=n%>_<%=x%>" id="medarbbelobdiv_<%=n%>_<%=x%>"><b><%=SQLBlessDot(formatnumber(mbelob, 2))%></b></div>
								<input type="hidden" name="FM_mbelob_<%=n%>_<%=x%>" id="FM_mbelob_<%=n%>_<%=x%>" value="<%=SQLBlessDot(formatnumber(mbelob, 2))%>">
								</td>
								<td style="border-right:1px #86B5E4 solid;" width="50">&nbsp;
								<input type="button" value="+" onClick="decimaler('plus', '<%=n%>', '<%=x%>')" style="font-size:9px; background-color:#8caae6; width:14px; height:17;">
							<input type="button" value="-" onClick="decimaler('minus', '<%=n%>', '<%=x%>')" style="font-size:9px; background-color:#8caae6; width:10px; height:17;">
							</td>
							</tr>
							<%
					
					medarbstrpris = medarbstrpris & ", #"&MedarbejderTimepris&"#"
					faktot = faktot + fak
					mbelobtot = mbelobtot + mbelob
					
					fak = 0 
					venter = 0
					medarbejderBeloeb = 0
					n = n + 1
					
					
					
					oRec3.movenext
					wend
					oRec3.close
					

					faktot = 0
					mbelobtot = 0
			
			
			call subtotalakt(x, 0)
			%>
			
			<tr><td colspan=6 bgcolor="#ffffff" style="border-left:1px #8caae6 solid; border-bottom:1px #8caae6 solid; border-right:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="10" alt="" border="0"><br></td></tr>
			<tr><td colspan=6 bgcolor="#ffffff"><img src="../ill/blank.gif" width="1" height="8" alt="" border="0"><br></td></tr>
			
			<%
			Response.write strAktsubtotal
			strAktsubtotal = ""
			%>
			<tr><input type="hidden" name="antal_subtotal_akt_<%=x%>" id="antal_subtotal_akt_<%=x%>" value="<%=sa%>">
			<td colspan=3 bgcolor="#8caae6" style="padding-left:5;"><%=thisaktnavn(x)%> ialt:
			<input type="text" name="timer_subtotal_akt_<%=x%>" id="timer_subtotal_akt_<%=x%>" value="<%=subtotaltimer%>" style='border:0px; font-size:9; width:45; background-color:#d6dff5;'> timer</td>
			<td colspan=3 bgcolor="#8caae6" align=right style="padding-right:5;">
			<input type="text" name="belob_subtotal_akt_<%=x%>" id="belob_subtotal_akt_<%=x%>" value="<%=subtotalbelob%>" style='border:0px; font-size:9; width:55; background-color:#d6dff5;'> kr.</td>
			</tr>
			
			<%if func <> "red" then%>
			<tr>
				<td colspan=6 bgcolor="#d6dff5" style="padding-left:5;">
				<%
				'** Gemte timer p� medarbejdere der ikke l�ngere har adgang til denne akt ***
				
				len_usedmedabId = len(usedmedabId)
				temp_usedmedabId = left(usedmedabId, len_usedmedabId - 4)
				usedmedabIdKri = temp_usedmedabId
				
				strSQL2 = "SELECT sum(timer) AS timer FROM timer WHERE "& usedmedabIdKri &" AND taktivitetid ="& thisaktid(x) &" AND tfaktim = 1 AND (Tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"')" 'OR Tdato = '" & stdatoKri &"' AND timer.tidspunkt > '"& LastFakTid &"'
				'Response.write strSQL2
				oRec2.open strSQL2, oConn, 3 
				
				if not oRec2.EOF then 
					medarbejderGTimer = oRec2("timer")
				end if
				oRec2.close
				
				if len(medarbejderGTimer) <> 0 then
					medarbejderGTimer = medarbejderGTimer
				else
					medarbejderGTimer = 0
				end if
				
				if medarbejderGTimer <> 0 then
				Response.write "Skjulte timer p� aktivitet:<b> " & formatnumber(medarbejderGTimer, 2) & "</b> timer"
				else
				Response.write "&nbsp;"
				end if
				%>
				</td>
			</tr>
			<%end if%>
			
			<input type="hidden" name="aktId_n_<%=x%>" id="aktId_n_<%=x%>" value="<%=thisaktid(x)%>">
			<input type="hidden" name="antal_n_<%=x%>" id="antal_n_<%=x%>" value="<%=n%>">
			<input type="hidden" name="highest_aval_<%=x%>" id="highest_aval_<%=x%>" value="<%=sa%>">
			<%
		subtotaltimer = 0 
		subtotalbelob = 0 
		'aktiswritten = aktiswritten & ",#"&thisaktid(x)&"#"		
		'end if	
	next
	%>
	<input type="hidden" name="lastactive_x" id="lastactive_x" value="0">
	<input type="hidden" name="lastactive_a" id="lastactive_a" value="0">
	<%
	
	
	'*** antal felter ***
	antalxx = x
	antalnn = n
	%>
	<input type="hidden" name="antal_x" value="<%=antalxx%>">
	<input type="hidden" name="antal_n" id="antal_n" value="<%=antalnn%>">
	
	<%
	if antalxx = 0 AND func = "red" then 
	%>
	<tr bgcolor="#d6dff5">
		<td colspan=6 style="padding:10px;"><br><br><font color=red><b>Denne faktura <u>kladde</u> er oprettet uden der er valgt nogen aktiviteter.</b> <br>
		Derer derfor ikke nogen aktiviteter eller medarbejder timer der kan redigeres.<br><br>
		Slet denne kladde fra faktura-oversigten og opret en ny faktura med dette faktura nummer.</font><br><br>&nbsp;</td>
	<tr>
	<%
	end if
	%>
	
	<tr bgcolor="#d6dff5">
		<td colspan=6 style="border-bottom:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="20" alt="" border="0"><br></td></tr>
	<tr>
	<td valign="top" style="border-left:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="30" alt="" border="0"></td>
	<td colspan="2" style="padding-left:40; padding-top:10;">
	Timer ialt:
	<%
	'*********************************************************
	'Total timer og bel�b
	'*********************************************************
	if intTimer <> 0 then
	intTimer = intTimer 
	else
	intTimer = 0
	end if
	
	%><input type="hidden" name="FM_Timer" value="<%=intTimer%>">
	<div style="position:relative; width:65; border-bottom:2px limegreen solid; background-color:#ffffff; padding-right:3px;" align="right" name="divtimertot" id="divtimertot"><b><%=intTimer%></b></div>
	
	<td colspan="2" style="padding-right:10; padding-top:10;" align="right" style="padding-right:16;">Nettobel�b (subtotal):
		<!-- strBeloeb -->
		<%
		if len(totalbelob) <> 0 then
		thistotbel = SQLBlessDot(formatnumber(totalbelob, 2))
		else
		thistotbel = formatnumber(0, 2)
		end if
		%>
		<input type="hidden" name="FM_beloeb" value="<%=thistotbel%>">
		<div style="position:relative; width:65; border-bottom:2px #86B5E4 solid; background-color:#ffffff; padding-right:3px;" align="right" name="divbelobtot" id="divbelobtot"><b><%=thistotbel%></b></div>
		</td>
		<td valign="top" align=right style="border-right:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="25" alt="" border="0"></td>
	</tr>
	<%
	'***********************************************************
	%>
	<tr>
		<td valign="top" style="border-left:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="190" alt="" border="0"></td>
		<td valign="top" colspan="4" style="padding-left:50;"><br><br>Betalings betingelser / Kommentar: 
		<br><textarea cols="70" rows="5" name="FM_komm"><%=strKom%></textarea><br>
		<br>&nbsp;</td>
		<td valign="top" align=right style="border-right:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="190" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top" colspan="6" style="border-bottom:1px #8caae6 solid;"><img src="../ill/blank.gif" width="700" height="1" alt="" border="0"></td>
	</tr>
	</table>
	<br><br>
	
	<table cellpadding=0 border=0 cellspacing=0>
	<tr>
		<td valign="top">&nbsp;</td>
		<td><br><br><img src="ill/blank.gif" width="260" height="1" alt="" border="0">
		<input type="image" src="../ill/opretpil_fak.gif">
		<td valign="top" align=right>&nbsp;</td>
	</tr>
	<!-- Nedenst�ende Bruges af javascript --->
	<input type="hidden" name="FM_showalert" id="FM_showalert" value="0">
	</form>
	</table>
	<br><br>
	
<br>
<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
<br>
<br>
</div>
									
									<%if func <> "red" then%>
									<!-- joblog -->
									<div style="position:relative; visibility:visible; left:17; top:525; width:150; height:200; background-color:#FFFFFF; border:1px #003399 solid; overflow:auto; padding:3;">
									<b>Joblog i valgt periode:</b><br><font class=megetlillesort>
									B�de fakturerbare og ikke-faktorerbare timer vises i jobloggen.<%
									strSQL = "SELECT tmnr, tmnavn, tdato, timer, timerkom, taktivitetnavn, taktivitetid, tfaktim FROM timer WHERE tjobnr = '"& jobnr &"' AND tfaktim <> 5 AND (Tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"') ORDER BY taktivitetid, taktivitetnavn, tdato DESC"
									'response.write strSQL
									oRec.open strSQL, oConn, 3 
									x = 0
									while not oRec.EOF 
										
										if oRec("taktivitetid") <> laktaktid then
											
											if oRec("tfaktim") = 1 then
											imgthis = "<img src='../ill/fakbarikon_14px.gif' width='14' height='14' alt='' border='0'>"
											else
											imgthis = "<img src='../ill/notfakbarikon_14px.gif' width='14' height='14' alt='' border='0'>"
											end if
											
										Response.write "<br><br>"& imgthis &"&nbsp;&nbsp;"& oRec("taktivitetnavn") &"<hr>"
										end if
										
										Response.write "<u>"& oRec("tdato") &":</u> "& oRec("timer") &" "
										Response.write ""& left(oRec("tmnavn"), 12) &"<br>"
										
										if len(oRec("timerkom")) <> 0 then
										Response.write "<font class=megetlillesilver>"& oRec("timerkom") & "</font><br><br>"
										end if
									
									x = x + 1
									laktaktid = oRec("taktivitetid")  
									oRec.movenext
									wend
									oRec.close 
									
									if x = 0 then
									Response.write "<br>Ingen registreringer!"
									end if
									%>
									</div>
									<%end if%>

<%end select%>
<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->


