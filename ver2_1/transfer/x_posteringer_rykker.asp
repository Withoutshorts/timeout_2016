'** Posteringer ***		
											'** Opretter tilhørende posteringer hvis faktua godkendes som endelig **		
											if intFakbetalt <> 0 then 
														
														call opretpos() '*** /inc/functions_inc.asp
														
														'*********************************************************************
														'*** Hvis Rykker oprettes krediteres tidligere faktura ***************
														'********************************************************************* 
														if intType = 2 then
																strSQL = "SELECT oprid FROM fakturaer WHERE Fid="& request("id")
																'Response.write strSQL & "<br><br>"
																oRec.open strSQL, oConn, 3 
																if not oRec.EOF then
																oprid = oRec("oprid")
																end if
																oRec.close 
																
																
																	if oprid <> 0 then '** undgå tidligere fakturaer fra før økonomi del.
																	
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
