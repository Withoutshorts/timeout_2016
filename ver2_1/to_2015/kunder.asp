<%response.buffer = true 
 Session.LCID = 1030%>

<!--#include file="../inc/connection/conn_db_inc.asp"-->


<%'** JQUERY START ************************* %>

<%'** JQUERY END ************************* %>


<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<%'GIT 20160811 - SK
 '** ER SESSION UDLØBET  ****
    if len(session("user")) = 0 then
    errortype = 5
	call showError(errortype)
    response.end
    end if
 %>


<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->






    <!--#include file="inc/kunder_inc.asp"-->
    <%

    if media <> "print" then    
        call menu_2014
    end if
    

    %>

<div id="wrapper">
<div class="content">


<%

    select case func
    case "addlevel", "sublevel"
	thiskid = request("thiskid")
	hotprev = cint(request("hotprev"))
	if func = "addlevel" then
		if hotprev < 2 then
		hotprev = hotprev + 1
		else
		hotprev = hotprev
		end if
	else
		if hotprev > -2 then
		hotprev = hotprev - 1
		else
		hotprev = hotprev
		end if
	end if
	
	oConn.execute("UPDATE kunder SET hot = "& hotprev &" WHERE kid = "& thiskid &"")
	Response.redirect "kunder.asp?menu=crm&shokselector=1&ketype=e&selpkt=osigt&redirectid="&thiskid&"&medarb="&medarb

	case "slet"

	    oskrift = "kunder" 
        slttxta = "Du er ved at <b>SLETTE</b> en kunde. Er dette korrekt?<br><br>Du vil samtidig <b><u>slette alle</u></b> tilhørende job, aktiviteter, ressourceforecast og fakturaer på denne kunde."
        slttxtb = "" 
        slturl = "kunder.asp?func=sletok&id="& id

        call sletcnf_2015(oskrift, slttxta, slttxtb, slturl)

	
	
	case "sletok"
	'*** Her slettes en kunde ***'
	
	'** Sletter tilhørende job, aktiviteter og timer  
	
	
	strSQL = "SELECT id FROM job WHERE jobknr = "& id
    'Response.write strSQL
    'Response.end
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
			
			
			strSQL2 = "SELECT id FROM aktiviteter WHERE job = "& oRec("id") &"" 
	
			oRec2.open strSQL2, oConn, 3
			while not oRec2.EOF 
		
		
				oConn.execute("DELETE FROM aktiviteter WHERE id = "& oRec2("id") &"")
		
				oConn.execute("DELETE FROM timer WHERE TAktivitetId = "& oRec2("id") &"")
		
			oRec2.movenext
			wend
			oRec2.close
	
	
	'*** Sletter ressource timer på job ****
	oConn.execute("DELETE FROM ressourcer WHERE jobid = "& oRec("id")  &"")
	
	'*** Sletter fra Guiden aktive job timer på job ****
	oConn.execute("DELETE FROM timereg_usejob WHERE jobid = "& oRec("id")  &"")
	
	'*** Sletter job ***
	oConn.execute("DELETE FROM job WHERE jobknr = "& id  &"") '**kunde ID **'
	
	'*** Sletter aftaler ***
	oConn.execute("DELETE FROM serviceaft WHERE kundeid = "& oRec("id") &"")
	
	       
	        '*** Sletter ikke fakturaer på kunder / job **'
	        '*** Sletter fakturaer på job ****
	        'strSQLfak = "SELECT fid FROM fakturaer WHERE jobid = "& oRec("id") &""
	        'oRec2.open strSQLfak, oConn, 3
        	
	        'while not oRec2.EOF 
        	        
	        '        oConn.execute("DELETE FROM faktura_det WHERE fakid = "& oRec2("fid") &"")
	        '        oConn.execute("DELETE FROM fak_med_spec WHERE fakid = "& oRec2("fid") &"")
	        '        oConn.execute("DELETE FROM fak_mat_spec WHERE matfakid = "& oRec2("fid") &"")
        	        
	        'oRec2.movenext
	        'wend
	        'oRec2.close
        	
        	
	        'oConn.execute("DELETE FROM fakturaer WHERE jobid = "& oRec("id") &"")
	        
	
	

	
	oRec.movenext
	wend
	
	oRec.close
	
	
	
	
	'**** Sletter CRM aktioner ***'        
	oConn.execute("DELETE FROM crmhistorik WHERE kundeid = "& id &"")
	
	'**** Sletter kontaktpersoner ***'        
	oConn.execute("DELETE FROM kontaktpers WHERE kundeid = "& id  &"")
	
   
	'*** Indsætter i delete historik ****'
	kkundenavn = "-"
	kknr = "0"
	
	strSQLk = "SELECT kkundenavn, kkundenr FROM kunder WHERE kid = " & id
	oRec.open strSQLk, oConn, 3
	if not oRec.EOF then
	
	kkundenavn = oRec("kkundenavn")
	kknr = oRec("kkundenr")
	
	end if
	oRec.close
	
	
	
	call insertDelhist("kun", id, kknr, kkundenavn, session("mid"), session("user"))
	
	'*** Sletter kunden ***'
	oConn.execute("DELETE FROM kunder WHERE Kid = "& id &"")
	
	
	
	Response.redirect "kunder.asp?menu="&menu&"&ketype="&ketype&"&shokselector=1&func=hist&selpkt=osigt&medarb="&medarb
	
	
	case "sletfil"
	
	'*** Her spørges om det er ok at der slettes en fil ***
	%>
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:190; top:50; visibility:visible;">
	<br><br><br>
	<table cellspacing="2" cellpadding="2" border="0">
	<tr>
	    <td><img src="../ill/alert.gif" width="44" height="45" alt="" border="0">Du er ved at <b>slette</b> en fil. Er dette korrekt?</td>
	</tr>
	<tr>
	   <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	   <a href="kunder.asp?menu=kund&func=sletfilok&id=<%=id%>&filnavn=<%=request("filnavn")%>&kundeid=<%=request("kundeid")%>&medarb=<%=medarb%>">Ja</a>&nbsp;&nbsp;&nbsp;<a href="javascript:history.back()">Nej</a></td>
	</tr>
	</table>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%
	case "sletfilok"
	'*** Her slettes en fil ***
	'ktv
	'strPath =  "E:\www\timeout_xp\wwwroot\ver2_1\upload\"&lto&"\" & Request("filnavn")
	'Qwert
	strPath =  "d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\upload\"&lto&"\" & Request("filnavn")
	'Response.write strPath
	
	on Error resume Next 

	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	Set fsoFile = FSO.GetFile(strPath)
	fsoFile.Delete
	
	
	oConn.execute("DELETE FROM filer WHERE id = "& id &"")
	Response.redirect "kunder.asp?menu=kund&func=red&id="& request("kundeid")&""



    case "dbopr", "dbred"
	'*** Her indsættes en ny kunde i db ****
		if len(request("FM_navn")) = 0 then
		
		errortype = 13
		call showError(errortype)
			
			else
			
			if len(trim(request("FM_knr"))) = 0 OR len(request("FM_knr")) > 15 then
			
			errortype = 21
			call showError(errortype)
			
			else
					
					strKnr = request("FM_knr")
                    strKnr = replace(strKnr, " ", "")
					strKnrOK = "y"	
					'*** Her tjekkes for dubletter i db ***
					strSQL = "SELECT Kkundenr, Kkundenavn FROM kunder WHERE Kkundenr LIKE '"& strKnr &"' AND Kid <> "& id
					
					'Response.write strSQL
					
					oRec.open strSQL, oConn, 3	
					if not oRec.EOF then
					'Response.write "<br>her!"
					strKnrOK = "n"
					
					errKundenavn = oRec("Kkundenavn")
					errKundenr = oRec("Kkundenr")
					
					end if
					oRec.close
				
					if strKnrOK = "n" then

					
					errortype = 12
					call showError(errortype)
					
					else
				
				
				'''if request("FM_oprsomkunde") = "on" then
				'dbketype = "k"
				'	if request("FM_oprsomfirma") = "on" then
					dbketype = "ke"
				'	end if
				'else
				'	if request("FM_oprsomfirma") = "on" then
				 '	dbketype = "e"
				'	else
				'	dbketype = "n"
				'	end if
				'end if
				
				function SQLBless2(s)
				dim tmp
				tmp = s
				tmp = replace(tmp, "'", "''")
				SQLBless2 = tmp
				end function
				
				strNavn = Replace(request("FM_navn"), Chr(34), "")
				strNavn = SQLBless2(strNavn)
				
				
				intKtype = request("FM_type")
				
				strEditor = session("user")
				strDato = session("dato")
				
				strAdr = SQLBless2(request("FM_adr")) 
				strPostnr = request("FM_postnr")
				strCity = request("FM_city") 
				strLand = request("FM_land")
		 		strTlf = request("FM_tlf")
				strMobil = request("FM_mobil")
				strAlttlf= request("FM_alttlf")
				strFax = request("FM_fax")
				strEmail = request("FM_email")
				strEan = trim(request("FM_ean"))
				strWWW = request("FM_www")
				
				if len(request("FM_nace")) <> 0 then
				strNACE = trim(request("FM_nace"))
				else
				strNACE = "0"
				end if
				
				'*** Kundeansv. ***
				if len(request("FM_kundeans_1")) <> 0 then
				intKundeans1 = request("FM_kundeans_1")
				else
				intKundeans1 = 0
				end if
				
				if len(request("FM_kundeans_2")) <> 0 then
				intKundeans2 = request("FM_kundeans_2")
				else
				intKundeans2 = 0
				end if
				
				if len(request("FM_logo")) <> 0 then
				logo = request("FM_logo")
				else
				logo = 0
				end if
				
				
				'* Accounts **'
				if len(request("FM_regnr")) <> 0 then
				intRegnr = request("FM_regnr")
				else
				intRegnr = ""
				end if
				
				if len(request("FM_kontonr")) <> 0 then
				intKontonr = request("FM_kontonr")
				else
				intKontonr = ""
				end if

                
				if len(request("FM_regnr_b")) <> 0 then
				intRegnr_b = request("FM_regnr_b")
				else
				intRegnr_b = ""
				end if
				
				if len(request("FM_kontonr_b")) <> 0 then
				intKontonr_b = request("FM_kontonr_b")
				else
				intKontonr_b = ""
				end if
				


                
				if len(request("FM_regnr_c")) <> 0 then
				intRegnr_c = request("FM_regnr_c")
				else
				intRegnr_c = ""
				end if
				
				if len(request("FM_kontonr_c")) <> 0 then
				intKontonr_c = request("FM_kontonr_c")
				else
				intKontonr_c = ""
				end if


                if len(request("FM_regnr_d")) <> 0 then
				intRegnr_d = request("FM_regnr_d")
				else
				intRegnr_d = ""
				end if
				
				if len(request("FM_kontonr_d")) <> 0 then
				intKontonr_d = request("FM_kontonr_d")
				else
				intKontonr_d = ""
				end if

				
                if len(request("FM_regnr_e")) <> 0 then
				intRegnr_e = request("FM_regnr_e")
				else
				intRegnr_e = ""
				end if
				
				if len(request("FM_kontonr_e")) <> 0 then
				intKontonr_e = request("FM_kontonr_e")
				else
				intKontonr_e = ""
				end if


                if len(request("FM_regnr_f")) <> 0 then
				intRegnr_f = request("FM_regnr_f")
				else
				intRegnr_f = ""
				end if
				
				if len(request("FM_kontonr_f")) <> 0 then
				intKontonr_f = request("FM_kontonr_f")
				else
				intKontonr_f = ""
				end if

				
				strBank = request("FM_bank")
				strIban = request("FM_iban")
				strSwift = request("FM_swift")

                strBank_b = request("FM_bank_b")
				strIban_b = request("FM_iban_b")
				strSwift_b = request("FM_swift_b")

                strBank_c = request("FM_bank_c")
				strIban_c = request("FM_iban_c")
				strSwift_c = request("FM_swift_c")
                

                strBank_d = request("FM_bank_d")
				strIban_d = request("FM_iban_d")
				strSwift_d = request("FM_swift_d")

                strBank_e = request("FM_bank_e")
				strIban_e = request("FM_iban_e")
				strSwift_e = request("FM_swift_e")

                strBank_f = request("FM_bank_f")
				strIban_f = request("FM_iban_f")
				strSwift_f = request("FM_swift_f")

                if len(request("FM_cvr")) <> 0 then
				intCvr = trim(request("FM_cvr"))
				else
				intCvr = 0
				end if
				
				if len(request("FM_useasfak")) <> "0" then
				intuseasfak = request("FM_useasfak")
				else
				intuseasfak = 0
				end if
				
                if len(trim(request("FM_lincensindehaver_faknr_prioritet"))) then
                lincensindehaver_faknr_prioritet = request("FM_lincensindehaver_faknr_prioritet")
                else
                lincensindehaver_faknr_prioritet = 0
                end if

				strKomm = SQLBless2(request("FM_Komm"))
				strLevbet = SQLBless2(request("FM_levbet"))
				strBetbet = SQLBless2(request("FM_betbet"))
				
				if len(request("FM_prio_grp")) <> 0 then
				priogrp = request("FM_prio_grp")
				else
				priogrp = 0
				end if
				
				betbetint = request("FM_betbetint")

                if len(trim(request("FM_kfak_moms"))) <> 0 then
                kfak_moms = request("FM_kfak_moms")
                else
                    select case lto
                    case "epi_uk", "intranet - local"
                    kfak_moms = 4 '20%
                    case else
                    kfak_moms = 1
                    end select
                end if

                if len(trim(request("FM_kfak_sprog"))) <> 0 then
                kfak_sprog = request("FM_kfak_sprog")
                else
                    select case lto
                    case "epi_uk", "nt", "intranet - local"
                     kfak_sprog = 2 'uk
                    case else
                     kfak_sprog = 1
                    end select
               
                end if

                if len(trim(request("FM_valuta_0"))) <> 0 then
                kfak_valuta = request("FM_valuta_0")
                else
                kfak_valuta = 1
                end if
                        
				
				
                '*** Hvis denne kunde vælges som primær licens indehaver nulstilles den eksisterende ****'
                '*** Ændret 20170117
                '*** Det er tilladt at have multible selskaber
            
                call multible_licensindehavereOn()

				if intuseasfak = 1 AND cint(multible_licensindehavere) = 0 then
				oConn.execute("UPDATE kunder SET useasfak = 0 WHERE useasfak = 1")
				end if
		        
		        if request("FM_opdater_txt_felter") = "1" then
				strSQLinsval = strSQLinsval & "'" & strKomm & "', "_
				&" '"& strBetbet &"', '"& strLevbet &"', "_
				&" "& betbetint
				
				strSQLinsflds = strSQLinsflds & ", beskrivelse, "_
				&" betbet, levbet, betbetint"
				
				else
				
				strSQLinsval = ""
				strSQLinsflds = ""
				
				end if
		        
				if func = "dbopr" then
				
				strSQLins = "INSERT INTO kunder (Kkundenavn, Kkundenr, editor, Kdato, "_
				& "adresse, "_
				& "postnr, "_
				& "city, "_
				& "land, "_
				& "telefon, "_
				& "telefonmobil, "_
				& "telefonalt, "_
				& "fax, "_
				& "email, "_
				& "ean, "_
				& "url, "_
				& "ketype, "_
				& "regnr, "_
				& "kontonr, "_
                & "regnr_b, "_
				& "kontonr_b, "_
                & "regnr_c, "_
				& "kontonr_c, "_
				& "regnr_d, "_
				& "kontonr_d, "_
                & "regnr_e, "_
                & "regnr_f, "_
                & "kontonr_e, "_
                & "kontonr_f, "_
                & "cvr, "_
				& "bank, "_
                & "bank_b, "_
                & "bank_c, "_
                & "bank_d, "_
                & "bank_e, "_
                & "bank_f, "_
				& "useasfak, "_
				& "logo, swift, swift_b, swift_c, swift_d, swift_e, swift_f, iban, iban_b, iban_c, iban_d, iban_e, iban_f, ktype, kundeans1, "_
				&" kundeans2, nace, kfak_moms, kfak_sprog, kfak_valuta, "_
				&" sdskpriogrp "& strSQLinsflds &", lincensindehaver_faknr_prioritet) VALUES "_
		 		& " ('"& strNavn &"', '"& strKnr &"', '"& strEditor &"', '"& strDato &"', "_
				& "'" & strAdr &"', "_ 
				& "'" & strPostnr &"', "_ 
				& "'" & strCity &"', "_ 
				& "'" & strLand &"', "_
				& "'" & strTlf &"', "_ 
				& "'" & strMobil &"', "_
				& "'" & strAlttlf &"', "_
				& "'" & strFax &"', "_
				& "'" & strEmail &"', "_ 
				& "'" & strEan &"', "_ 
				& "'" & strWWW &"', "_
				& "'" & dbketype &"', "_   
				& "'" & intRegnr &"', "_
				& "'" & intKontonr &"', "_
                & "'" & intRegnr_b &"', "_
				& "'" & intKontonr_b &"', "_
                & "'" & intRegnr_c &"', "_
				& "'" & intKontonr_c &"', "_
				& "'" & intRegnr_d &"', "_
				& "'" & intKontonr_d &"', "_
                & "'" & intRegnr_e &"', "_
                & "'" & intRegnr_f &"', "_
				& "'" & intKontonr_e &"', "_
                & "'" & intKontonr_f &"', "_
                & "'" & intCVR &"', "_
				& "'" & strBank &"', "_
                & "'" & strBank_b &"', "_
                & "'" & strBank_c &"', "_
				& "'" & strBank_d &"', "_
                & "'" & strBank_e &"', "_
                & "'" & strBank_f &"', "_
                & "" & intuseasfak &", "_
				& "" & logo &", '"& strSwift &"', '"& strSwift_b &"', '"& strSwift_c &"', '"& strSwift_d &"', '"& strSwift_e &"', '"& strSwift_f &"', "_
                &"'"& strIban &"', '"& strIban_b &"', '"& strIban_c &"', '"& strIban_d &"', '"& strIban_e &"', '"& strIban_f &"', "_
				& "" & intKtype &", "& intKundeans1 &", "& intKundeans2 &","_
				&" '"& strNACE &"', "& kfak_moms &", "& kfak_sprog &", "& kfak_valuta &", "_
				&" "& priogrp &", "& strSQLinsval &", "& lincensindehaver_faknr_prioritet &")"
				
				
				'Response.Write "strSQLins " & strSQLins
				'Response.flush
				
                oConn.execute(strSQLins)
				
				
				'***** Henter id på den netop oprettede kunde *****
				strSQL = "SELECT kid FROM kunder ORDER BY kid DESC"
				oRec.open strSQL, oConn, 3
				if not oRec.EOF then
				 	thisKid = oRec("kid")
				end if
				oRec.close
				
				else
				
			
				
				strSQLupdate = "UPDATE kunder SET Kkundenavn ='"& strNavn &"', Kkundenr = '"& strKnr &"', Kdato = '" & strDato &"', editor = '" &strEditor &"', "_
				& "adresse = '" & strAdr & "', "_
				& "postnr = '" & strPostnr & "', "_
				& "city = '" & strCity & "', "_
				& "land = '" & strLand & "', "_
				& "telefon = '" & strTlf & "', "_
				& "telefonmobil = '" & strMobil & "', "_
				& "telefonalt = '" & strAlttlf & "', "_
				& "fax = '" & strFax & "', "_
				& "email = '" & strEmail & "', "_
				& "ean = '" & strEan & "', "_
				& "url = '" & strWWW & "', "_
				& "ketype = '" & dbketype & "', "_
				& "regnr = '" & intRegnr & "', "_
				& "kontonr = '" & intKontonr & "', "_
                & "regnr_b = '" & intRegnr_b & "', "_
				& "kontonr_b = '" & intKontonr_b & "', "_
                & "regnr_c = '" & intRegnr_c & "', "_
				& "kontonr_c = '" & intKontonr_c & "', "_
                & "regnr_d = '" & intRegnr_d & "', "_
				& "kontonr_d = '" & intKontonr_d & "', "_
                & "regnr_e = '" & intRegnr_e & "', "_
                & "regnr_f = '" & intRegnr_f & "', "_
				& "kontonr_e = '" & intKontonr_e & "', "_
                & "kontonr_f = '" & intKontonr_f & "', "_
				& "cvr = '" & intCvr & "', "_
				& "bank = '" & strBank & "', "_
                & "bank_b = '" & strBank_b & "', "_
                & "bank_c = '" & strBank_c & "', "_
				& "bank_d = '" & strBank_d & "', "_
                & "bank_e = '" & strBank_e & "', "_
                & "bank_f = '" & strBank_f & "', "_
                & "useasfak = " & intuseasfak & ", "_
				& "logo = "& logo &", "_
				& "swift = '" & strSwift &"', "_
                & "swift_b = '" & strSwift_b &"', "_
                & "swift_c = '" & strSwift_c &"', "_
				& "swift_d = '" & strSwift_d &"', "_
                & "swift_e = '" & strSwift_e &"', "_
                & "swift_f = '" & strSwift_f &"', "_
                & "iban = '" & strIban &"', ktype = "& intKtype &", "_
                & "iban_b = '" & strIban_b &"', "_
                & "iban_c = '" & strIban_c &"', "_
				& "iban_d = '" & strIban_d &"', "_
                & "iban_e = '" & strIban_e &"', "_
                & "iban_f = '" & strIban_f &"', "_
                & "kundeans1 = "& intKundeans1 &", kundeans2 = "& intKundeans2 &", nace = '"& strNACE &"', kfak_moms = "& kfak_moms &", kfak_sprog = "& kfak_sprog &", kfak_valuta = "& kfak_valuta &", "_
				& "sdskpriogrp = "& priogrp &", lincensindehaver_faknr_prioritet = "& lincensindehaver_faknr_prioritet 
				
				if request("FM_opdater_txt_felter") = "1" then
				strSQLupdate = strSQLupdate & ", beskrivelse = '" & strKomm & "', "_
				&" betbet = '"& strBetbet &"', levbet = '"& strLevbet &"', "_
				&" betbetint = "& betbetint 
				end if
				
				strSQLupdate = strSQLupdate &" WHERE Kid = "& id 
				
				'Response.Write strSQLupdate
				'Response.flush
				oConn.execute(strSQLupdate)
				
				thiskid = id
				end if
				
				
                '*** Irriterende ved update, at den gemmer  ***'
				'if len(thiskri) <> 0 AND func <> "dbopr" then
				'thiskri = thiskri
				'else
				'thiskri = left(strNavn, 1)
				'end if

                'response.write "thiskri: "& thiskri
                'response.end
				
				
				'**** Tilknyt foldere til kunde ****
				if len(request("FM_stfoldergruppe")) <> 0 then
				folderGrpId = request("FM_stfoldergruppe")
				else
				folderGrpId = 0
				end if
				
				if folderGrpId <> 0 then
						
						datoSQL = year(now) & "/" & month(now) & "/" & day(now)
						
						strSQL = "SELECT navn, id, kundese FROM foldere WHERE stfoldergruppe = "& folderGrpId
						oRec.open strSQL, oConn, 3 
						while not oRec.EOF 
						
							strSQL2 = "INSERT INTO foldere (navn, kundese, kundeid, jobid, editor, dato)"_
							&" VALUES ('"& oRec("navn") &"', "& oRec("kundese") &", "& thiskid &", 0, '"& strEditor &"', '"& datoSQL &"')"
							
							oConn.execute(strSQL2)
							
						oRec.movenext
						wend
						oRec.close 
				
				end if	
				
				
				
				
				if menu = "crm" AND rdir <> "1" then
				Response.redirect "kunder.asp?menu=crm&shokselector=1&ketype=e&redirectid="&thisKid&"&medarb="&medarb
				else
						
						'**** Tilknytter Konto ****
						if request("FM_opretkonto") = "1" then
        						
        						
						            if rdir = "1" then
						            Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
	                                Response.redirect "kontoplan.asp?menu=kon&func=opret&showmenu=no&kundeid="&thisKid
        						    'Response.Write("<script language=""JavaScript"">window.close();</script>")
        						
						            else
						            %><!--#include file="../inc/regular/header_inc.asp"--><%
						            Response.Write("<script language=""JavaScript"">window.open('kontoplan.asp?menu=kon&func=opret&showmenu=no&kundeid="&thisKid&"','','width=700,height=500,left=190,top=280');</script>")
        						
        						
						            %>
						            <div style="position:absolute; left:190; top:20;">
						            <h3>Tilføj Konto.</h3>
						            <table><tr><td>
						            Når den ønskede konto er oprettet vender du tilbage til<br>
						            kunde oversigten ved at bruge nedenstående link.<br>
						            Kontoen er oprettet korrekt når der er klikket på "opret" og det viste popup vindue ikke længere er synligt.<br>
						            <br>
						            <%
						            Response.write "<a href='kunder.asp?menu=kund&shokselector=1&ketype=k&id="&thisKid&"&medarb="&medarb&"'>Tilbage til kontaktoversigten</a>"
						            %></td></tr></table>
						            </div>
						            <%
						            end if
						
						else
						
						    if rdir = "1" then 'on the fly fra job, ell. SDSK
						    Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
	                        Response.Write("<script language=""JavaScript"">window.close();</script>")
						    else
						    Response.redirect "kunder.asp?menu=kund&shokselector=1&ketype=k&id="&thisKid&"&medarb="&medarb&"&lastkid="&thisKid
						    end if
						
						end if
				end if
			
			
			
				
				
			end if
		end if
	end if

  

    '********************************************************************
    '**************** CREATE / EDIT *************************************                                 
    case "red", "opret"
    
    %><script src="js/kunder_jav.js" type="text/javascript"></script><%

    if func = "opret" then

        if menu = "crm" then
		kchecked = ""
		fchecked = "checked"
		else
		kchecked = "checked"
		fchecked = ""
		end if
	
	intKtype = 0
	strNavn = ""
	'varSubVal = "opretpil" 
	'varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	
	prio_grp = 1
	
	strLevbet = ""
	strBetbet = ""

    useasfakCHK0 = "SELECTED"

    kfak_moms = 1
    kfak_sprog = 1
    call basisValutaFN()
    valuta = basisValId '** Basis


    headerTxt = global_txt_180 '"opret"
    
                                    
   else


    
    strSQL = "SELECT Kid, Kkundenavn, Kkundenr, Kdato, editor, "_
	& "adresse, "_
	& "postnr, "_
	& "city, "_
	& "land, "_
	& "telefon, "_
	& "telefonmobil, "_
	& "telefonalt, "_
	& "fax, "_
	& "email, "_
	& "ean, "_
	& "url, "_
	& "ketype, beskrivelse, "_
	& "regnr, "_
	& "kontonr, "_
    & "regnr_b, "_
	& "kontonr_b, "_
    & "regnr_c, "_
	& "kontonr_c, "_
	& "regnr_d, "_
	& "kontonr_d, "_
    & "regnr_e, "_
	& "kontonr_e, "_
    & "regnr_f, "_
	& "kontonr_f, "_
    & "cvr, "_
	& "bank, swift, iban, "_
    & "bank_b, swift_b, iban_b, "_
    & "bank_c, swift_c, iban_c, "_
	& "bank_d, swift_d, iban_d, "_
    & "bank_e, swift_e, iban_e, "_
    & "bank_f, swift_f, iban_f, "_
    & "useasfak, "_
	& "logo, ktype, kundeans1, kundeans2, nace, betbet, levbet, sdskpriogrp, betbetint, kfak_moms, kfak_sprog, kfak_valuta, lincensindehaver_faknr_prioritet "_
	&" FROM kunder WHERE Kid=" & id
	
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	intKid = oRec("Kid") 
	strNavn = oRec("Kkundenavn")
	strKnr = oRec("Kkundenr")
	strDato = oRec("Kdato")
	strLastUptDato = oRec("Kdato")
	strEditor = oRec("editor")
	strAdr = oRec("adresse")
	strPostnr = oRec("postnr")
	strCity = oRec("city")
	strLand = oRec("land")
	strTlf = oRec("telefon")
	strMobil = oRec("telefonmobil")
	strAlttlf = oRec("telefonalt")
	strFax = oRec("fax")
	strEmail  = oRec("email")
	strEan  = oRec("ean")
	strWWW = oRec("url")
	ketype = oRec("ketype")
	strKomm = oRec("beskrivelse")
	intregnr = oRec("regnr")
	intkontonr = oRec("kontonr")
    intregnr_b = oRec("regnr_b")
	intkontonr_b = oRec("kontonr_b")
    intregnr_c = oRec("regnr_c")
	intkontonr_c = oRec("kontonr_c")
    intregnr_d = oRec("regnr_d")
	intkontonr_d = oRec("kontonr_d")
    intregnr_e = oRec("regnr_e")
	intkontonr_e = oRec("kontonr_e")

    intregnr_f = oRec("regnr_f")
	intkontonr_f = oRec("kontonr_f")
	intCVR = oRec("cvr")
	
    strBank = oRec("bank")
    strBank_b = oRec("bank_b")
    strBank_c = oRec("bank_c")
    strBank_d = oRec("bank_d")
    strBank_e = oRec("bank_e")
    strBank_f = oRec("bank_f")
    
    strIban_b = oRec("iban_b")
    strIban_c = oRec("iban_c")
    strIban_d = oRec("iban_d")
	strIban_e = oRec("iban_e")
    strIban_f = oRec("iban_f")
    strIban = oRec("iban")
	
    strSwift = oRec("swift")
    strSwift_b = oRec("swift_b")
    strSwift_c = oRec("swift_c")
	strSwift_d = oRec("swift_d")
    strSwift_e = oRec("swift_e")
    strSwift_f = oRec("swift_f")
    logo = oRec("logo")
	intKtype = oRec("ktype")
	strNACE = oRec("nace")
	
	intKundeans1 = oRec("kundeans1")
	intKundeans2 = oRec("kundeans2")
	
	prio_grp = oRec("sdskpriogrp")
	
	
	eruseasfak = oRec("useasfak")
    lincensindehaver_faknr_prioritet = oRec("lincensindehaver_faknr_prioritet")

    useasfakCHK1 = ""
    useasfakCHK2 = ""
    useasfakCHK3 = ""
    useasfakCHK0 = ""
    useasfakCHK5 = ""
    useasfakCHK6 = ""
    useasfakCHK7 = ""

    select case eruseasfak
    case 1
    useasfakCHK1 = "SELECTED"
    case 2
    useasfakCHK2 = "SELECTED"
    case 3
    useasfakCHK3 = "SELECTED"
    case 0
    useasfakCHK0 = "SELECTED"
    case 5
    useasfakCHK5 = "SELECTED"
    case 6
    useasfakCHK6 = "SELECTED"
    case 7
    useasfakCHK7 = "SELECTED"
    case else
    useasfakCHK0 = "SELECTED"
    end select

	
	
	select case ketype 
	case "ke" 
	kchecked = "checked"
	fchecked = "checked"
	case "k"
	kchecked = "checked"
	fchecked = ""
	case "e"
	kchecked = ""
	fchecked = "checked"
	case else
	kchecked = "checked"
	fchecked = ""
	end select
	
	strLevbet = oRec("levbet")
	strBetbet = oRec("betbet")
	
	betbetint = oRec("betbetint")

    kfak_moms = oRec("kfak_moms")
    kfak_sprog = oRec("kfak_sprog")

    valuta = oRec("kfak_valuta")

	
	end if
	oRec.close

    dbfunc = "dbred"
    
    kundeOskriftTrim = global_txt_124
    kundeOskriftTrim_len = len(kundeOskriftTrim)
    kundeOskriftTrim_left = left(kundeOskriftTrim, kundeOskriftTrim_len - 1)
    headerTxt = tsa_txt_251 &" "& kundeOskriftTrim_left '"rediger"

    end if
    
    %>


    <div class="container">
      <div class="portlet">
        <h3 class="portlet-title">
          <u><%=headerTxt%></u>
            
        </h3>
          
         
           <form action="kunder.asp?func=<%=dbfunc%>" method="post">
	    <input type="hidden" name="menu" value="<%=menu%>">
	    <input type="hidden" name="id" id="id" value="<%=id%>">
	    <input type="hidden" name="ketype" value="<%=ketype%>"> 
	    <!--<input type="hidden" name="FM_soeg" value="<%=thiskri%>">-->
	    <input type="hidden" name="medarb" value="<%=medarb%>">
	    <input type="hidden" name="rdir" value="<%=rdir%>">
        <input id="Hidden1" name="FM_opdater_txt_felter" value="1" type="hidden" />

        <!-- Opdater/Annuller knapper -->
        <div style="margin-top:-15px; margin-bottom:15px;">
        <%if func = "red" then %>
         <a class="btn btn-sm btn-default" href="kontaktpers.asp?kid=<%=id %>&issogsubmitted=1" role="button" target="_blank">Kontaktpersoner</a>
             <a class="btn btn-sm btn-info" href="kontaktpers.asp?func=opr&kid=<%=id %>" role="button" target="_blank">Opret Kontaktperson +</a>
            
            <a class="btn btn-sm btn-default" href="../timereg/filer.asp?kundeid=<%=id%>&jobid=0&nomenu=1" role="button" target="_blank">Filarkiv</a>
          <%end if %>  
          

           
          <%

          '*** Rettigheder til at opdatere kunde? 
          submitLevelOK = 1      
          select case lto 
          case "epi", "epi_no", "epi_uk", "epi_ab", "intranet - local", "epi2017"

                 
                if level <> 1 AND eruseasfak <> 5 AND func = "red" then
                submitLevelOK = 0
                end if

           end select %>

          
                <%if cint(submitLevelOK) = 1 then %>     
                <button type="submit" class="btn btn-success btn-sm pull-right" id="sbm_upd0"><b>Opdatér</b></button>
                <%else %>
                <button type="submit" class="btn btn-sm pull-right" id="sbm_upd0" disabled><b>Du har ikke rettigheder til at opdatere denne kunde.</b></button>
                <%end if %>

            <div class="clearfix"></div>
        </div>
               
        


        <div class="portlet-body">
            <!-- STAMDATA REDIGER NAVN / SORTERING / ID -->
         
          
            <section>
                <div class="well well-white">
                    <div class="row">
                        <div class="col-lg-12">
                            <h4 class="panel-title-well">Stamdata</h4>
                        </div>
                    </div>

                    <div class="row">
                        <div class="col-lg-1">&nbsp;</div>
                        <div class="col-lg-2">Navn:&nbsp<span style="color:red;">*</span></div>
                        <div class="col-lg-3">   
                              <input type="text" name="FM_navn" class="form-control input-small" value="<%=strNavn %>" placeholder="Navn"/>
                            </div>
                        <div class="col-lg-1">&nbsp</div>
                        <div class="col-lg-1">Kundetype:</div>
                        <div class="col-lg-4 pad-r30">   

                         

                             <%
             
                             name = "FM_useasfak"
                             wdt = 350    
                             sel = eruseasfak
                             autosubmit = 0
                             call ktyper(name, wdt, sel, autosubmit, lto, level) %>

      
       


                        <!--<select class="form-control input-small" name="ktype">
				        <option  value="0">Alle</option>
				        <%
						        strSQL = "SELECT navn, id FROM kundetyper ORDER BY navn"
						        oRec.open strSQL, oConn, 3
						        while not oRec.EOF
						
							        if cint(ktype) = cint(oRec("id")) then
							        isSelected = "SELECTED"
							        else
							        isSelected = ""
							        end if
						
						        %>
						        <option value="<%=oRec("id")%>" <%=isSelected%>><%=left(oRec("navn"), 15)%></option>
						        <%
						        oRec.movenext
						        wend
						        oRec.close
						        %>
				           </select>
                            -->

                            </div>
                       </div>
                                
                        <!-- KundeId -->
                  <div class="row">
                         <div class="col-lg-1">&nbsp;</div>
                        <div class="col-lg-2 pad-t5">Kundenummer:&nbsp<span style="color:red;">*</span>

                               
                        </div>
                      

                        <div class="col-lg-3">   

                            <%if func <> "red" then
		
		                     
		
		                        if lto = "execon" OR lto = "immenso" then
		                            if len(lastUsedKnr) > 4 then
		                                call erDetInt(right(lastUsedKnr, 4))
		                                    if isInt = 0 then
	                                        strKnr = "D0"& (right(lastUsedKnr, 4)/1) + 1
	                                        else
	                                        strKnr = 0
	                                        end if
	                                    isInt = 0
	                                 else
	                                    strKnr = 0
	                                 end if
		                        end if

                                if lto = "userminds" then
                                    call erDetInt(right(lastUsedKnr, 4))
		                            if isInt = 0 then
                                    strKnr = lastUsedKnr + 1
                                    else
                                    strKnr = ""
                                    end if
                                end if
		
		
		                        end if
		                       

		                        select case lto
		                        case "execon", "immenso", "userminds"
		                         if level = 1 then
		                         kidAaben = 1
		                         else
		                         kidAaben = 0
		                         end if
		                        case else
		                         kidAaben = 1
		                        end select
		
		                        if kidAaben = 1 then %>
		                        <input type="text" class="form-control input-small" name="FM_knr" value="<%=strKnr%>" placeholder="Kundenummer">
		                        <%else %>
		                        <input type="text" class="form-control input-small" name="FM_knrdis" value="<%=strKnr%>" DISABLED>
		                        <input type="hidden" name="FM_knr" value="<%=strKnr%>">
		                        <%end if %>
                                
                                <!--
		                        <br />Kan indeholde både tal og bogstaver, maks 15 kar. 
		                        <br />Ved brug af FI nummer må kunde Id maks være på 8 karakterer.
                            -->
                                <%if func = "opret" then %>
                                <span style="color:#999999; font-size:10px;">
		                        Sidste kundenr:
		                        <%strSQL = "SELECT kkundenr, kkundenavn FROM kunder WHERE kid <> 0 ORDER BY kid DESC"
		                        oRec.open strSQL, oConn, 3
		                        lastUsedKnr = "0"
		                        if not oRec.EOF then
		                        lastUsedKnr = oRec("kkundenr") 
                                lastKnavn = oRec("kkundenavn")
		                        end if
		                        oRec.close
		
		                        Response.Write "<b>"& lastUsedKnr & "</b> ("& lastKnavn &")"

                                %>

                                   </span>
                                <%end if %>

                           
                        </div>
                        <div class ="col-lg-1">&nbsp</div>
                        <div class="col-lg-1">Segment:</div>
                        <div class="col-lg-4 pad-r30">   
                              <select class="form-control input-small" name="FM_type" id="FM_type">
                                <%if intKtype = 0 then
                                ktype=SEL = "SELECTED"
                                else
                                ktype=SEL = ""
                                end if%>

                                <option value="0" <%=ktype=SEL %>>Vælg type (Ingen)</option>
		
		                        <%
		                        strSQLktype = "SELECT id, navn, rabat FROM kundetyper ORDER BY navn"
		                        oRec.open strSQLktype, oConn, 3 
		                        while not oRec.EOF 
			                        if oRec("id") = intKtype then
			                        selThis = "SELECTED"
			                        else
			                        selThis = ""
			                        end if
		                        %>
		                        <option value="<%=oRec("id")%>" <%=selThis%>><%=oRec("navn")%>  (<%=oRec("rabat") %> %)</option>
		                        <%
		                        oRec.movenext
		                        wend
		                        oRec.close 
		                        %>
		                        </select>
                            </div>
                       </div>
                    
                </div>
             </section>

            <!--Åben luk-->
            <section>
                <!-- Accordion -->
                <div class="panel-group accordion-panel" id="accordion-paneled">
                    <!-- PersonData -->
                    <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseOne">Kundedata</a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseOne" class="panel-collapse collapse">
                        <div class="panel-body">

                      <div class="row">
                        <div class="col-lg-1">&nbsp;</div>
                        <div class="col-lg-2">Adresse:</div>
                        <div class="col-lg-3">   
                              <input type="text" name="FM_adr" class="form-control input-small" value="<%=strAdr%>"/>
                            </div>
                        <div class="col-lg-1">&nbsp</div>
                        <div class="col-lg-2">Alt.tlf:</div>
                        <div class="col-lg-2">   
                              <input type="text" name="FM_alttlf" class="form-control input-small" value="<%=strAlttlf%>"  />
                            </div>
                        <div class="col-lg-1">&nbsp</div>
                       </div>

                       <div class="row">
                        <div class="col-lg-1">&nbsp;</div>
                        <div class="col-lg-2">Post Nr.:</div>
                        <div class="col-lg-3">   
                         <input type="text" name="FM_postnr" class="form-control input-small" value="<%=strPostnr%>" />
                        </div>
                        <div class="col-lg-1">&nbsp</div>
                        <div class="col-lg-2">Fax:</div>
                        <div class="col-lg-2">   
                              <input type="text" name="FM_fax" class="form-control input-small" value="<%=strfax%>"  />
                            </div>
                        <div class="col-lg-1">&nbsp</div>
                       </div>

                       <div class="row">
                        <div class="col-lg-1">&nbsp;</div>
                        <div class="col-lg-2">By:</div>
                        <div class="col-lg-3">   
                              <input type="text" name="FM_city" class="form-control input-small" value="<%=strcity%>"  />
                            </div>

                        <div class="col-lg-1">&nbsp</div>
                        <div class="col-lg-2">Email:</div>
                        <div class="col-lg-2">   
                              <input type="text" name="FM_email" class="form-control input-small" value="<%=stremail%>"  />
                            </div> 
                        </div>


                            <div class="row">
                                 <div class="col-lg-1">&nbsp;</div>
                        <div class="col-lg-2">Land:</div>
                        <div class="col-lg-3">   
                              <select class="form-control input-small" name="FM_land">
						            <!--#include file="../timereg/inc/inc_option_land.asp"-->
                                <%if func = "red" then%>
                        		<option SELECTED><%=strland%></option>
		                        <%else%>
		                        <option SELECTED>Danmark</option>
		                        <%end if%>
		                     </select>
                            </div>
                        <div class="col-lg-1">&nbsp</div>
                        <div class="col-lg-2">CVR nr:</div>
                        <div class="col-lg-2">   
                              <input type="text" name="FM_cvr" class="form-control input-small" value="<%=intcvr%>"  />
                        </div>
                        <div class="col-lg-1">&nbsp</div>
                      </div>

                       <div class="row">
                        <div class="col-lg-1">&nbsp;</div>
                        <div class="col-lg-2">Tel:</div>
                        <div class="col-lg-3">   
                              <input type="text" name="FM_tlf" class="form-control input-small" value="<%=strtlf%>"  />
                            </div>
                        <div class="col-lg-1">&nbsp</div>
                        <div class="col-lg-2">NCA kode:</div>
                        <div class="col-lg-2">   
                              <input type="text" name="FM_nace" class="form-control input-small" value="<%=strNACE%>"  />
                            </div>
                        <div class="col-lg-1">&nbsp</div> 
                        </div>

                            <div class="row">
                                <div class="col-lg-1">&nbsp;</div>
                                <div class="col-lg-2">Mobil:</div>
                                <div class="col-lg-3">   
                                      <input type="text" name="FM_mobil" class="form-control input-small" value="<%=strmobil%>"  />
                                    </div>

                                 <div class="col-lg-1">&nbsp;</div>
                                <div class="col-lg-2">EAN:</div>
                                <div class="col-lg-2">   
                                      <input type="text" name="FM_ean" class="form-control input-small" value="<%=strean%>"  />
                                    </div>
                                      
                                    </div>

                              <div class="row">
                        <div class="col-lg-1">&nbsp;</div>
                        <div class="col-lg-2">WWW: (Domæner)</div>
                        <div class="col-lg-8">   
                              <textarea id="FM_www" name="FM_www" style="height:80px;" class="form-control input-small"><%=strWWW%></textarea>
                            </div>

                       
                        </div>


                                </div>
                       </div>
                    </div>



                    <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseTwo">Konti</a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseTwo" class="panel-collapse collapse">
                        <div class="panel-body">
                        Følgende konti er knyttet til denne kunde via kontoplanen: (sættes i ERP)<br />
                            <div style="height:100px; padding:10px 10px 10px 10px; overflow:auto;">
                        <%  
                            
                            strSQLkp = "SELECT k.kontonr, k.navn, k.id, k.kid FROM kontoplan AS k "_
			                &" WHERE k.kid = "& id &" ORDER BY k.kontonr, k.navn"
                            
                            oRec6.open strSQLkp, oConn, 3
                            while not oRec6.EOF 
                            %>
                            <%=oRec6("navn") & " kontonr: " & oRec6("kontonr") %><br />
                            <%
                            oRec6.movenext
                            wend 
                            oRec6.close
                            %>
                        </div>

                        <br /><br />
                        <%if cint(eruseasfak) = 1 OR cint(eruseasfak) = 0 then 
                            
                        if cint(eruseasfak) = 1 then%>
                        Denne kunde står som licensindehaver og bruger fakturarækkefølge gruppe nr: (se kontrolpanel)<br />
                        <%else %>
                        Job/Projekter på denne kunde faktureres som udgangspunkt af følgende licensindehaver: <br /> 
                        <%end if%>


                        <%
                            
                            faknr_prioritet_0_SELECTED = ""
                            faknr_prioritet_2_SELECTED = ""
                            faknr_prioritet_3_SELECTED = ""
                            faknr_prioritet_4_SELECTED = ""
                            faknr_prioritet_5_SELECTED = ""

                            select case lincensindehaver_faknr_prioritet
                            case 0
                            faknr_prioritet_0_SELECTED = "SELECTED" 
                            case 2
                            faknr_prioritet_2_SELECTED = "SELECTED"
                            case 3
                            faknr_prioritet_3_SELECTED = "SELECTED"
                            case 4
                            faknr_prioritet_4_SELECTED = "SELECTED"
                            case 5
                            faknr_prioritet_5_SELECTED = "SELECTED"
                         end select  
                            
                        call multible_licensindehavereOn()%>

                            <select name="FM_lincensindehaver_faknr_prioritet" style="width:300px;">
                            <%if cint(eruseasfak) = 1 then %>
                        
                            <option value="0" <%=faknr_prioritet_0_SELECTED %>>1 (sidste fakturanr: <%=multi_fakturanr %>)</option>
                                
                                <%if cint(multible_licensindehavere) = 1 then %>
                                <option value="2" <%=faknr_prioritet_2_SELECTED %>>2 (sidste fakturanr: <%=multi_fakturanr_2 %>)</option>
                                <option value="3" <%=faknr_prioritet_3_SELECTED %>>3 (sidste fakturanr: <%=multi_fakturanr_3 %>)</option>
                                <option value="4" <%=faknr_prioritet_4_SELECTED %>>4 (sidste fakturanr: <%=multi_fakturanr_4 %>)</option>
                                <option value="5" <%=faknr_prioritet_5_SELECTED %>>5 (sidste fakturanr: <%=multi_fakturanr_5 %>)</option>
                                <%end if %>

                            <%else%>

                                 <%
                                  strSQLklicensow = "SELECT kkundenavn, kkundenr FROM kunder WHERE useasfak = 1 AND lincensindehaver_faknr_prioritet = 0 "
                                  oRec6.open strSQLklicensow, oConn, 3
                                  if not oRec6.EOF then  
                                    strLicensindehaver0 = oRec6("kkundenavn") &" "& oRec6("kkundenr")
                                  end if
                                  oRec6.close  

                                     if cint(multible_licensindehavere) = 1 then 

                                 strSQLklicensow = "SELECT kkundenavn, kkundenr FROM kunder WHERE useasfak = 1 AND lincensindehaver_faknr_prioritet = 2 "
                                  oRec6.open strSQLklicensow, oConn, 3
                                  if not oRec6.EOF then  
                                    strLicensindehaver2 = oRec6("kkundenavn") &" "& oRec6("kkundenr")
                                  end if
                                  oRec6.close  

                                     strSQLklicensow = "SELECT kkundenavn, kkundenr FROM kunder WHERE useasfak = 1 AND lincensindehaver_faknr_prioritet = 3 "
                                  oRec6.open strSQLklicensow, oConn, 3
                                  if not oRec6.EOF then  
                                    strLicensindehaver3 = oRec6("kkundenavn") &" "& oRec6("kkundenr")
                                  end if
                                  oRec6.close  

                                     strSQLklicensow = "SELECT kkundenavn, kkundenr FROM kunder WHERE useasfak = 1 AND lincensindehaver_faknr_prioritet = 4 "
                                  oRec6.open strSQLklicensow, oConn, 3
                                  if not oRec6.EOF then  
                                    strLicensindehaver4 = oRec6("kkundenavn") &" "& oRec6("kkundenr")
                                  end if
                                  oRec6.close  

                                     strSQLklicensow = "SELECT kkundenavn, kkundenr FROM kunder WHERE useasfak = 1 AND lincensindehaver_faknr_prioritet = 5 "
                                  oRec6.open strSQLklicensow, oConn, 3
                                  if not oRec6.EOF then  
                                    strLicensindehaver5 = oRec6("kkundenavn") &" "& oRec6("kkundenr")
                                  end if
                                  oRec6.close  

                                     end if

                                  %>
                                
                                  

                                <option value="0" <%=faknr_prioritet_1_SELECTED %>>1 (<%=strLicensindehaver0 %>)</option>
                                
                                <%if cint(multible_licensindehavere) = 1 then %>
                                <option value="2" <%=faknr_prioritet_2_SELECTED %>>2 (<%=strLicensindehaver2 %>)</option>
                                <option value="3" <%=faknr_prioritet_3_SELECTED %>>3 (<%=strLicensindehaver3 %>)</option>
                                <option value="4" <%=faknr_prioritet_4_SELECTED %>>4 (<%=strLicensindehaver4 %>)</option>
                                <option value="5" <%=faknr_prioritet_5_SELECTED %>>5 (<%=strLicensindehaver5 %>)</option>
                                <%end if %>

                             <%end if%>

                        </select>
                        <%end if %>


                            
                         <%if cint(eruseasfak) = 1 then %>

                            

                        <br /><br />
                        Nedenstående konti benyttes udelukkende som afsender konti på faktura layout.<br /><br />

                            
                       
                         <div class="row">
                        <div class="col-lg-1">&nbsp;</div>
                        <div class="col-lg-2">Konto:</div>
                        <div class="col-lg-1">regnr.:</div>
                        <div class="col-lg-1">   
                             <input type="text" name="FM_regnr" class="form-control input-small" value="<%=intregnr%>"  />
                            </div>
                        <div class="col-lg-1">Kontonr:</div>
                        <div class="col-lg-2">   
                              <input type="text" name="FM_kontonr" class="form-control input-small" value="<%=intkontonr%>"  />
                            </div>
                              <div class="col-lg-4">&nbsp;</div> 
                            </div>
                    
                            <div class="row">
                        <div class="col-lg-1">&nbsp;</div>
                        <div class="col-lg-2">Bank navn:</div>
                        <div class="col-lg-3">   
                              <input type="text" name="FM_bank" class="form-control input-small" value="<%=strBank %>"  />
                            </div>
                              <div class="col-lg-6">&nbsp;</div> 
                            </div>
                            <div class="row">
                        <div class="col-lg-1">&nbsp;</div>
                        <div class="col-lg-2">Swift kode:</div>
                        <div class="col-lg-3">   
                              <input type="text" name="FM_swift" class="form-control input-small" value="<%=strSwift %>"  />
                            </div>
                              <div class="col-lg-6">&nbsp;</div> 
                            </div>
                            <div class="row">
                        <div class="col-lg-1">&nbsp;</div>
                        <div class="col-lg-2">Iban kode:</div>
                        <div class="col-lg-3">   
                              <input type="text" name="FM_iban" class="form-control input-small" value="<%=strIban %>"  />
                            </div>
                              <div class="col-lg-6">&nbsp;</div> 
                            </div>
                            

                            <div class="row">
                               <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2 pad-t10"><a href="#" id="visflkonti" >Flere bankkonti</a></div>
                                <div class="col-lg-9">&nbsp</div>
                            </div>

                            


                            <%for k = 1 to 5 
                                
                                select case k
                                case 1
                                kontoalt = " (alt. 1)"
                                konto_ext = "_b"
                                intregnr = intregnr_b
                                intkontonr = intkontonr_b
                                strBank = strBank_b
                                strSwift = strSwift_b
                                strIban = strIban_b
                                case 2
                                kontoalt = " (alt. 2)"
                                konto_ext = "_c"
                                 intregnr = intregnr_c
                                intkontonr = intkontonr_c
                                strBank = strBank_c
                                strSwift = strSwift_c
                                strIban = strIban_c
                                case 3
                                kontoalt = " (alt. 3)"
                                konto_ext = "_d"
                                 intregnr = intregnr_d
                                intkontonr = intkontonr_d
                                strBank = strBank_d
                                strSwift = strSwift_d
                                strIban = strIban_d
                                case 4
                                kontoalt = " (alt. 4)"
                                konto_ext = "_e"
                                 intregnr = intregnr_e
                                intkontonr = intkontonr_e
                                strBank = strBank_e
                                strSwift = strSwift_e
                                strIban = strIban_e
                                case 5
                                kontoalt = " (alt. 5)"
                                konto_ext = "_f"
                                 intregnr = intregnr_f
                                intkontonr = intkontonr_f
                                strBank = strBank_f
                                strSwift = strSwift_f
                                strIban = strIban_f
                                end select%>

                            
                           <div class="row tr_konti" id="tr_konto<%=konto_ext %>">
                            <div class="col-lg-1">&nbsp;</div>
                            <div class="col-lg-2"><br />Konto <%=kontoalt %>:</div>
                            <div class="col-lg-1"><br />regnr.:</div>
                            <div class="col-lg-1"><br />   
                                 <input type="text" name="FM_regnr<%=konto_ext %>" class="form-control input-small" value="<%=intregnr%>"  />
                                </div>
                            <div class="col-lg-1"><br />Kontonr:</div>
                            <div class="col-lg-2"><br />   
                                  <input type="text" name="FM_kontonr<%=konto_ext %>" class="form-control input-small" value="<%=intkontonr%>"  />
                                </div>
                                  <div class="col-lg-4">&nbsp;</div> 
                            </div><!-- row -->
                    
                            <div class="row tr_konti">
                                <div class="col-lg-1">&nbsp;</div>
                                <div class="col-lg-2">Bank navn:</div>
                                <div class="col-lg-3">   
                                      <input type="text" name="FM_bank<%=konto_ext %>" class="form-control input-small" value="<%=strBank %>"  />
                                    </div>
                                      <div class="col-lg-6">&nbsp;</div> 
                            </div><!-- row -->

                            <div class="row tr_konti">
                                <div class="col-lg-1">&nbsp;</div>
                                <div class="col-lg-2">Swift kode:</div>
                                <div class="col-lg-3">   
                                      <input type="text" name="FM_swift<%=konto_ext %>" class="form-control input-small" value="<%=strSwift %>"  />
                                    </div>
                                      <div class="col-lg-6">&nbsp;</div> 
                            </div><!-- row -->

                            <div class="row tr_konti">
                                <div class="col-lg-1">&nbsp;</div>
                                <div class="col-lg-2">Iban kode:</div>
                                <div class="col-lg-3">   
                                      <input type="text" name="FM_iban<%=konto_ext %>" class="form-control input-small" value="<%=strIban %>"  />
                                    </div>
                                      <div class="col-lg-6">&nbsp;</div> 
                            </div><!-- row -->
                                
                                
                                
                                
                             <%next
                                 
                                 
                             end if 'eruseasfak = 1%>
                            

                 </div>
                        </div>
                    </div>

                    <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseThree">Kundeindstillinger</a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseThree" class="panel-collapse collapse">
                        <div class="panel-body">

                    <!--selectbox skal opdateres-->
                            
                            <div class="row">
                                <div class ="col-lg-1 pad-t20">&nbsp</div>
                                <div class ="col-lg-4 pad-t20">Kundeansvarlig/Keyaccount 1 & 2:</div>
                                <div class="col-lg-3 pad-t20">
                                  <select class="form-control input-small" name="FM_kundeans_1" id="FM_kundeans_1">
		                            <option value="0">Ingen</option>
			                            <%
			
			                            'if func <> "red" then
			                            strSQL = "SELECT mnavn, mnr, mid, init FROM medarbejdere WHERE mansat <> 2 ORDER BY mnavn"
			                            'else
			                            'strSQL = "SELECT mnavn, mnr, mid, kundeans1 FROM medarbejdere LEFT JOIN kunder ON (kunder.kid = "& id &")  WHERE mansat <> 2"
			                            'end if
			                            oRec.open strSQL, oConn, 3 
			                            while not oRec.EOF 
			
			                            if func <> "red" then
			                            usemed = session("mid")
			                            else
			                            usemed = intKundeans1
			                            end if
			
				                            if cint(usemed) = oRec("mid") then
				                            medsel = "SELECTED"
				                            else
				                            medsel = ""
				                            end if
			                            %>
			                            <option value="<%=oRec("mid")%>" <%=medsel%>><%=oRec("mnavn")%> 
                                            <%if len(trim(oRec("init"))) <> 0 then %>
                                            [<%=oRec("init")%>]
                                            <%end if %>

			                            </option>
			                            <%
			                            oRec.movenext
			                            wend
			                            oRec.close 
			                            %>
		                        </select>
                                </div>
                                
                                <div class="col-lg-3 pad-t20">
                                        <select class="form-control input-small" name="FM_kundeans_2" id="FM_kundeans_2">
		                                    <option value="0">Ingen</option>
			                                    <%
			
			                                    'if func <> "red" then
			                                    strSQL = "SELECT mnavn, mnr, mid, init FROM medarbejdere WHERE mansat <> 2 ORDER BY mnavn"
			                                    'else
			                                    'strSQL = "SELECT mnavn, mnr, mid, kundeans2 FROM medarbejdere LEFT JOIN kunder ON (kunder.kid = "& id &")  WHERE mansat <> 2"
			                                    'end if
			                                    oRec.open strSQL, oConn, 3 
			                                    while not oRec.EOF 
			
			                                    if func <> "red" then
			                                    usemed2 = 0
			                                    else
			                                    usemed2 = intKundeans2
			                                    end if
			
				                                    if cint(usemed2) = oRec("mid") then
				                                    medsel2 = "SELECTED"
				                                    else
				                                    medsel2 = ""
				                                    end if
			                                    %>
			                                    <option value="<%=oRec("mid")%>" <%=medsel2%>><%=oRec("mnavn")%> 

                                                    <%if len(trim(oRec("init"))) <> 0 then %>
                                            [<%=oRec("init")%>]
                                            <%end if %>
			                                    </option>
			                                    <%
			                                    oRec.movenext
			                                    wend
			                                    oRec.close 
			                                    %>
		                                </select>
                                </div>
                                
                            </div>

                          

                            <div class="row">
                            <div class="col-lg-1">&nbsp</div>
                            <div class="col-lg-10"><br />Beskrivelse:
                                <textarea id="TextArea2" name="FM_komm" cols="70" rows="7" class="form-control input-small"><%=strKomm %></textarea>
                            </div>
                           
                        </div>
                          
                            <!-- 20150927 GIVER IKKE MER MENING. AlLE oprettes auto til ke. Der skal ikke tages stilling til dette mere. ***'
                            <div class="row">
                                <div class="col-lg-1 pad-t20">&nbsp</div>
                                <div class="col-lg-2 pad-t20">TSA / CRM:</div>
                                <div class="-col-lg-9">&nbsp</div>
                            </div>
                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-3">Opret/tilføj som kunde i TSA systemet:</div>
                                <div class="col-lg-1"><input type="checkbox" name="FM_oprsomkunde" value="on" <%=kchecked%> /></div>
                                <div class="-col-lg-6">&nbsp</div>
                            </div>
                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-3">og/eller som firma i CRM systemet:</div>
                                <div class="col-lg-1"><input type="checkbox" name="FM_oprsomfirma" value="on" <%=fchecked%> /></div>
                                <div class="-col-lg-6 pad-b20">&nbsp</div>
                            </div>

                                -->


                              <%
	                        antalfoldere = 0
	                        strSQL = "SELECT count(id) AS antalfoldere FROM foldere WHERE kundeid = "& id &" AND kundeid <> 0 AND stfoldergruppe = 0"
	                        oRec.open strSQL, oConn, 3 
	                        if not oRec.EOF then
		                        antalfoldere  = oRec("antalfoldere")
	                        end if 
	                        oRec.Close					
	                        %>

                            <div class ="row">
                                <div class="col-lg-1 pad-t20">&nbsp</div>
                                <div class="col-lg-2 pad-t20">Tilknyt Standardfoldere?
                                    <span style="color:#999999">Der er findes allerede: <b><%=antalfoldere%> folder(e)</b> på denne kunde.
                                        <!--<br>
	                                Hvis du ønsker at tilknytte flere foldere kan det enten gøres ved at vælge en ny Standardfolder gruppe her,
	                               <br> eller ved manuelt at tilføje foldere under fanebladet "Filer".<br>--></span>
                                </div>
                                <div class="col-lg-2 pad-t20">
                                <select class="form-control input-small" name="FM_stfoldergruppe" id="FM_stfoldergruppe">
		                        <option value="0">Nej</option>
			                        <%
			
			                        strSQL = "SELECT navn, id FROM folder_grupper WHERE id <> 0 ORDER BY navn"
			                        oRec.open strSQL, oConn, 3 
			                        while not oRec.EOF 
			
			
			
				                        'if cint(usemed2) = oRec("mid") then
				                        'medsel2 = "SELECTED"
				                        'else
				                        'medsel2 = ""
				                        'end if
			                        %>
			                        <option value="<%=oRec("id")%>"><%=oRec("navn")%></option>
			                        <%
			                        oRec.movenext
			                        wend
			                        oRec.close 
			                        %>
		                        </select>
                                </div>
                                <div class="col-lg-7 pad-b20">&nbsp</div>
                            </div>

                            <div class="row">
                                <div class="col-lg-1 pad-t20">&nbsp</div>
                                <div class="col-lg-2 pad-t20">ServiceDesk Aftalegruppe: 
                                    <span style="color:#999999;">Når der er valgt en gruppe kan der oprettes incindets på denne kunde.</span></div>
                                <div class="col-lg-2 pad-t20">
                                    <select class="form-control input-small" name="FM_prio_grp" id="FM_prio_grp">
		                            <option value="0">Ingen</option>
	                            <%
	                            strSQL = "SELECT id, navn FROM sdsk_prio_grp WHERE id <> 0"
	                            oRec.open strSQL, oConn, 3
	                            while not oRec.EOF
		
		                            if cint(prio_grp) = oRec("id") then
		                            pgSEL = "SELECTED"
		                            else
		                            pgSel = ""
		                            end if%>
		
	                            <option value="<%=oRec("id")%>" <%=pgSel%>><%=oRec("navn")%></option>
	                            <%
	                            oRec.movenext
	                            wend
	
	                            oRec.close
	                            %>
	                        </select>
                                </div>
                                <div class="col-lg-8">&nbsp</div>
                            </div>

                         </div>
                      </div>
                   </div>

                    <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseFour">Faktura indstilllinger (debitor)</a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseFour" class="panel-collapse collapse">
                        <div class="panel-body">
                    
                        
                        

                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2">Moms:<br />
                                    <%strSQLmoms = "SELECT id, moms FROM fak_moms WHERE id <> 0 ORDER BY id " %>
                                    <select class="form-control input-small" name="FM_kfak_moms" style="width:160px" >
                                    <%oRec6.open strSQLmoms, oConn, 3
                                    while not oRec6.EOF 

                                        if cint(kfak_moms) = cint(oRec6("id")) then
                                        fakmomsSeL = "SELECTED"
                                        else
                                        fakmomsSeL = ""
                                        end if

                                    %><option value="<%=oRec6("id") %>" <%=fakmomsSeL %>><%=oRec6("moms") %>%</option><%
                  
                                    oRec6.movenext
                                    wend 
                                    oRec6.close%>
                                </select>
                                </div>
                                
                            </div>

                            <div class="row">
                                <div class="col-lg-1 pad-t10">&nbsp</div>
                                <div class="col-lg-2 pad-t10">Sprog: <br />
                                    <%strSQLsprog = "SELECT id, navn FROM fak_sprog WHERE id <> 0 ORDER BY id " %>
                                    <select class="form-control input-small" name="FM_kfak_sprog" style="width:160px" >

                                        <%oRec6.open strSQLsprog, oConn, 3
                                        while not oRec6.EOF 

                                            if cint(kfak_sprog) = cint(oRec6("id")) then
                                            faksprogSeL = "SELECTED"
                                            else
                                            faksprogSeL = ""
                                            end if

                                        %><option value="<%=oRec6("id") %>" <%=faksprogSeL %>><%=oRec6("navn") %></option><%
                  
                                        oRec6.movenext
                                        wend 
                                        oRec6.close%>
                                    </select>
                                </div>
                              
                            </div>

                               <div class="row">
                                <div class="col-lg-1 pad-t20">&nbsp</div>
                                <div class="col-lg-2 pad-t20">Valuta:

                                <%call basisValutaFN()
            
                                 call valutakoder(0, valuta, 1) %>
                                </div>
                                   </div>

                            <div class="row">
                                <div class="col-lg-1 pad-t20">&nbsp</div>
                                <div class="col-lg-3 pad-t20">Forfaldsdato: (kreditperiode)
                                    <% 
	                                select case lto
	                                case "execon", "immenso"
	                                betbetint = 8
	                                disa = 0'1
                                    lang = 0
	                                case else
	                                betbetint = betbetint
	                                disa = 0
                                    lang = 0
	                                end select

                                    nameid = "FM_betbetint"
	
	                            call betalingsbetDage(betbetint, disa, lang, nameid)
	                            %>
                                </div>
                                <div class ="col-lg-8 pad-b20">&nbsp</div>
                            </div>

                            <div class="row">
                                <div class="col-lg-1 pad-t20">&nbsp</div>
                                <div class="col-lg-10 pad-t20">Betallingsbetingelser:
                                    <textarea id="TextArea1" name="FM_betbet" cols="70" rows="7" class="form-control input-small"><%=strBetbet %></textarea>
                                </div>
                            </div>

                            <div class="row">
                                <div class="col-lg-1 pad-t20">&nbsp</div>
                                <div class="col-lg-10 pad-t20">Leveringsbetingelser:
                                    <textarea id="TextArea3" name="FM_levbet" cols="70" rows="7" class="form-control input-small"><%=strLevbet %></textarea>
                                </div>
                            </div>

                        </div>
                       </div>
                    </div>

                    <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseFive">Logoer</a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseFive" class="panel-collapse collapse">
                        <div class="panel-body">
                        

                      
		                Maks:  bredde: 150px * højde 100px<br />
		             
		                <a href="../timereg/upload.asp?type=kundelogo&kundeid=<%=id%>&jobid=0" target="_blank">Tilknyt nyt logo</a>
                            <br />&nbsp;
		                
                             
                                
                               

                                

	                <%
	                strSQL = "SELECT id, filnavn, editor, dato FROM filer WHERE kundeid = "& id &" AND type = 1 ORDER BY filnavn"
	
	                oRec.open strSQL, oConn, 3
	                j = 0
	                while not oRec.EOF
	                if len(trim(oRec("filnavn"))) <> 0 then
                         
	                        if logo = oRec("id") then
	                        logosel = "CHECKED"
	                        erfundet = 1
	                        else
	                        logosel = ""
	                        end if
	                %>
                     <div class="row">
                                <div class="col-lg-1">&nbsp</div>
	                <div class="col-md-2"><input type="radio" name="FM_logo" value="<%=oRec("id")%>" <%=logosel%>>&nbsp;&nbsp;<%=oRec("filnavn") %> </div>
                    <div class="col-md-2"><img src="../inc/upload/<%=lto%>/<%=oRec("filnavn")%>" width=150 height=100 alt='' border='0'></div>
                    <div class="col-md-2"><a href="kunder.asp?func=sletfil&id=<%=oRec("id")%>&kundeid=<%=id%>&filnavn=<%=oRec("filnavn")%>"><i class="fa fa-times" style="color:red;"></i></a></div>
                    </div><%
	          
	               
	                j = j + 1
	                end if
	                oRec.movenext
	                wend
	                oRec.close
	
	                if j = 0 then
	            
	               %>
                   <div class="row">
                    <div class="col-lg-1">&nbsp</div>
                   <div class="col-md-2">Ingen logoer er tilknyttet denne kunde.</div>
                   </div>
                   <%
	              
	
	                else
	
	                if erfundet <> 1 then
	                logoselingen = "CHECKED"
	                else
	                logoselingen = ""
	                end if
	
	                
	                 %>
                     <div class="row">
                     <div class="col-lg-1">&nbsp</div>
                     <div class="col-md-2"><input type="radio" name="FM_logo" value="0" <%=logoselingen%>>&nbsp;Intet logo valgt.</div>
                            </div>
                       <%
	               
	
	                end if
	                %>
                   

                                    <!--
                                    <div class="row">
                                        <div class="col-lg-1">&nbsp</div>
                                
                               

                                        <div class="col-md-4">

                                  

                                          <div class="fileinput fileinput-new" data-provides="fileinput">
                                            <div class="fileinput-preview thumbnail" data-trigger="fileinput" style="width: 200px; height: 150px;"></div>
                                                <div>
                                                  <span class="btn btn-default btn-file"><span class="fileinput-new">Select image</span><span class="fileinput-exists">Change</span><input type="file" name="..."></span>
                                                  <a href="#" class="btn btn-default fileinput-exists" data-dismiss="fileinput">Remove</a>
                                                </div>
                                            </div>

                                        </div> <!-- /.col -->



                                   <!-- </div>
                          

                          
                                </div>
                                </div>
                                         -->
                    </div>


        </div>
     </section>
    <br /><br /><br />

                     <%if func = "red" then %>
                    <div style="font-weight: lighter;">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></div>       
                    <%end if %>  

        &nbsp;
        </div><!-- well well-white -->

     <!-- Opdater/Annuller knapper -->
        

        <div style="margin-top:15px; margin-bottom:15px;">


            <%if func = "red" then %>
          
                    
            <%if cint(submitLevelOK) = 1 then %>     
                 <%call antalFakturaerKid(intKid) %>
		        <%if cint(antalFak) = 0 then %>
		        <a class="btn btn-primary btn-sm" role="button" href="kunder.asp?menu=<%=menu %>&func=slet&id=<%=intKid%>&ketype=e&medarb=<%=medarb%>"><b>Slet</b></a>&nbsp;
		        <%end if %>
            <%end if %>
            
            
                  
           
        <%end if %>
     
                <%if cint(submitLevelOK) = 1 then %>     
                <button type="submit" class="btn btn-success btn-sm pull-right" id="sbm_upd2"><b>Opdatér</b></button>
                <%else %>
                <button type="submit" class="btn btn-sm pull-right" id="sbm_upd0" disabled><b>Du har ikke rettigheder til at opdatere denne kunde.</b></button>
                <%end if %>

            
                     
         
            <div class="clearfix"></div>
        </div>

    </form>

    </div>



<%
case else    

    %><script src="js/kunder_list_jav.js" type="text/javascript"></script><%

    if len(trim(request("lastkid"))) <> 0 then
    lastkid = request("lastkid")
    else
    lastkid = 0
    end if

    if len(trim(request("land"))) <> 0 then
    landSEL = request("land")
    else
    landSEL = ""
    end if

%>

<!--Kunde liste-->
<div class="container">
     <div class="portlet">
        <h3 class="portlet-title">
          <u><%=global_txt_124 %></u>
        </h3>
        

         <form action="kunder.asp?menu=<%=menu%>&func=opret&ketype=<%=ketype%>&medarb=<%=medarb%>" method="post"> 
              
         <section>
                <div class="row">
                    <div class="col-lg-10">&nbsp;</div>
                    <div class="col-lg-2">
                <button class="btn btn-sm btn-success pull-right"><b>Opret ny +</b></button><br />&nbsp;
                </div>
            </div>
        </section>
         </form>
        

          <form action="kunder.asp?menu=<%=menu%>&viskunder=1&sogsubmitted=1" method="POST">
          
          <section>
                <div class="well">
                         <h4 class="panel-title-well">Søgefilter</h4>
                         <div class="row">
                             <div class="col-lg-4">
                                 Søg på Navn: <br />
			                    <input type="text" name="FM_soeg" id="FM_soeg" class="form-control input-small" placeholder="% Wildcard" value="<%=thiskri%>">

                               </div>
                              
                         


                         
                             <div class="col-lg-2">
                               Kundeans.:<select name="medarb" class="form-control input-small" onchange="submit()">
				                    <option value="0">Alle</option>
				                    <%
						                    strSQL = "SELECT mnavn, mid FROM medarbejdere WHERE mansat = 1 ORDER BY mnavn"
						                    oRec.open strSQL, oConn, 3
						                    while not oRec.EOF
						
							                    if cint(medarb) = cint(oRec("mid")) then
							                    isSelected = "SELECTED"
							                    else
							                    isSelected = ""
							                    end if
						
						                    %>
						                    <option value="<%=oRec("mid")%>" <%=isSelected%>><%=oRec("mnavn")%></option>
						                    <%
						                    oRec.movenext
						                    wend
						                    oRec.close
						                    %>
				                    </select>

                               </div>

                               <div class="col-lg-2">Type:

                    <%
                         'name = "FM_useasfak"
                         'wdt = 350    
                         'sel = eruseasfak
                        
                        autosubmit = 1
                        call ktyper(name, 150, intuseasfak, autosubmit, level, lto) %>


                               </div>

                              <div class="col-lg-2">

                                  Segment: (rabat %)<br /> <select name="ktype" class="form-control input-small" onchange="submit();">
				                    <option value="0">Alle</option>
                                      
				                    <%
						                    strSQL = "SELECT navn, id FROM kundetyper ORDER BY navn"
						                    oRec.open strSQL, oConn, 3
						                    while not oRec.EOF
						
							                    if cint(ktype) = cint(oRec("id")) then
							                    isSelected = "SELECTED"
							                    else
							                    isSelected = ""
							                    end if
						
						                    %>
						                    <option value="<%=oRec("id")%>" <%=isSelected%>><%=left(oRec("navn"), 15)%></option>
						                    <%
						                    oRec.movenext
						                    wend
						                    oRec.close


                                            if cint(ktype) = -1 then
                                            usegmentSEL = "SELECTED"
                                            else
                                            usegmentSEL = ""
                                            end if
						                    %>

                                      <option value="-1" <%=usegmentSEL %>>Uden segment</option>
				                    </select>

                              </div>

                               <div class="col-lg-2">Land: 
                                    <select name="land" class="form-control input-small" onchange="submit();">

                                        <%if landSEL = "" then
                                         isSelected = "SELECTED"
                                         else
                                         isSelected = ""
                                        end if %>

                                        <option value="" <%=isSelected%>>Alle</option>
                                   
                                        <%strSQL = "SELECT land FROM kunder WHERE land <> '' AND land IS NOT NULL GROUP BY land ORDER BY land"
                                         oRec.open strSQL, oConn, 3
						                    while not oRec.EOF

                                            if landSEL = oRec("land") then
                                            isSelected = "SELECTED"
                                            else
                                            isSelected = ""
                                            end if

                                         %>
						                    <option value="<%=oRec("land")%>" <%=isSelected%>><%=oRec("land")%></option>
						                    <%
						                    oRec.movenext
						                    wend
						                    oRec.close
                                       
                                       %>
                                         
                                        </select>
                                </div>
                                </div><!-- ROW -->

                    <div class="row">
                               <div class="col-lg-12"><br />
                                    <button type="submit" class="btn btn-secondary btn-sm pull-right"><b>Vis kunder >></b></button>
                                   </div>
                                   <!--<input type="submit" class="btn btn-sm btn-secondary pull-right" value=" Søg " /></div>--> 

                         </div><!-- ROW -->

                              <%if menu = "crm" then%>

                                           <div class="row">
                                        <div class="col-lg-4">Interesse:<br /><select name="FM_hotnot" style="width:60px;" onchange="submit()">
		                                <option value="99" <%=hotnotSEL0  %>>Alle</option>
		                                <option value="-2" <%=hotnotSEL1  %>>-2 (Not)</option>
		                                <option value="-1" <%=hotnotSEL2  %>>-1</option>
		                                <option value="0" <%=hotnotSEL3  %>>0</option>
		                                <option value="1" <%=hotnotSEL4  %>>+1</option>
		                                <option value="2" <%=hotnotSEL5  %>>+2 (Hot)</option>
            				
		                                </select>

                                                </div>
                                               <div class="col-lg-8">&nbsp;</div> 
                                        </div><!-- ROW -->
                           <%end if %>

                       
                          
                        


                    </div><!-- Well Well White -->
                    </section>
              </form>
         
      
			
			

          
          <!--

                <td><br /><br />Sortér efter:<br />
                    <select name="FM_sorter" style="width:60px;" onchange="submit()">
		       
		            <option value="0" <%=sortSel0  %>>Navn</option>
		            <option value="1" <%=sortSel1  %>>Id</option>
            				
		            </select>
                </td>
         -->
            
			
		

         

         
        <div class="portlet-body">

          <table id="kundeliste" class="table dataTable table-striped table-bordered table-hover ui-datatable"> 
             
            <thead>
              <tr>
                <th style="width: 12%">Kundenr.</th>
                <th style="width: 15%">Navn</th>
                <th style="width: 5%">Postnr.</th>
                <th style="width: 15%">By</th>
                <th style="width: 15%">Land</th>
                <th style="width: 5%">Telefon</th>
                <th style="width: 2%">Type</th>
                <th style="width: 3%">Segment</th>
                <th style="width: 10%">CVR nr.</th>
                <th style="width: 10%">Job</th>
                <th style="width: 10%">Faktura</th>
              </tr>
            </thead>
              <tbody>
                   <%
	
	
		
		            '** Søge kriterier **
		            if cint(visikkekunder) <> 1 OR cint(usekri) = 1 then
		
			            if cint(usekri) = 1 then '**SøgeKri
			            sqlsearchKri = " (Kkundenavn LIKE '"& thiskri &"%' OR (Kkundenr LIKE '"& thiskri &"%' AND kkundenr <> '0') OR telefon = '"& thiskri &"' OR url LIKE '"& thiskri &"%')" 
			            else
			            sqlsearchKri = " (kid <> 0)"
			            end if
			
		            else
		            sqlsearchKri = " (kid = 0)"
		            end if
		
		
		            if ktype > 0 then
		            typeKri = " AND ktype = "&ktype&" "
		            else
                       if cint(ktype) = -1 then 'ALLE uden segment
                       typeKri = " AND ktype = 0 "
                       else
		                typeKri = " AND ktype > -2 "
                       end if
		            end if
		
		
		            if cint(hotnot) = 99 then
	                hotnotKri = " AND hot <> 99 "
	                else
	                hotnotKri = " AND hot = "& hotnot
	                end if
		
		            kansSQLKri = useopraf_editor

                    if len(trim(landSEL)) <> 0 then
                    landSELSQL = " AND land = '"& landSEL &"'"
                    else
                    landSELSQL = ""
                    end if
	
	            if menu = "crm" then 'NOT IN USE 2050905

		            fontcolor = "fir_g"
		            strSQL = "SELECT kid, kkundenavn, kkundenr, useasfak, hot, "_
		            &" kunder.editor, ktype, kunder.kundeans1, kunder.kundeans2, adresse, postnr, city, land, telefon, "_
		            &" m1.mnavn AS kans1, m1.init AS kans1_init, m2.mnavn AS kans2, m2.init AS kans2_init, m1.mnr AS mnr1, m2.mnr AS mnr2, cvr, ean, url "_
		            &" FROM kunder "_
		            &" LEFT JOIN medarbejdere m1 ON m1.mid = kundeans1"_
		            &" LEFT JOIN medarbejdere m2 ON m2.mid = kundeans2"_
		            &" WHERE "& sqlsearchKri &" "& kansSQLKri &" AND kunder.ketype <> 'k' "& typeKri &""& hotnotKri & sortBy
		
	
	            else
		            strSQL = "SELECT Kid, Kkundenavn, Kkundenr, useasfak, ktype, kunder.kundeans1, kunder.kundeans2,"_
		            &" m1.mnavn AS kans1, m1.mnr AS mnr1, m1.init AS kans1_init, m2.mnavn AS kans2, m2.mnr AS mnr2, "_
		            &" m2.init AS kans2_init, adresse, postnr, city, land, telefon, cvr, ean, ktype, kt.navn AS segment, url FROM kunder "_
		            &" LEFT JOIN medarbejdere m1 ON m1.mid = kundeans1"_
		            &" LEFT JOIN medarbejdere m2 ON m2.mid = kundeans2"_
		            &" LEFT JOIN kundetyper AS kt ON (kt.id = kunder.ktype)"_
		            &" WHERE "& sqlsearchKri &" "& kansSQLKri &" "& useasfakSQL &" AND ketype <> 'xx' "& typeKri & landSELSQL & sortBy & " LIMIT 4000" 'ketype <> e
	            end if

    	        kids = "0"


	
	            '** Udskriv query til skærm **
	            'if session("mid") = 1 then
                'Response.write strSQL
	            'Response.flush
	            'end if
                '*****************************
	            cnt = 0
        
                    oRec.open strSQL, oConn, 3
        
                    while not oRec.EOF 
                       
                       '** Til Export fil ***
		                kids = kids & "," & oRec("kid")
                       
                      
                        if cint(lastKid) = cint(oRec("kid")) then
                       trBgCol = "#FFFFE1" 
                       else
                       trBgCol = ""
                       end if%>

                        <tr style="background-color:<%=trBgCol%>;">
                    <td><%=oRec("Kkundenr") %></td>   
                    <td> <a href="kunder.asp?func=red&id=<%=oRec("kid") %>"><%=oRec("kkundenavn") %></a> </td>
                    <td><%=oRec("postnr") %></td>
                    <td><%=oRec("city") %></td>
                    <td><%=oRec("land") %></td>
                    <td><%=oRec("telefon") %></td>
                    <td><%
                    
                    select case oRec("useasfak")
                    case 1
                    useasfakTxt = "Licensindehaver"
                    case 2
                    useasfakTxt = "Datterselskab"
                    case 3
                     useasfakTxt = "Andet"
                    case 0
                     useasfakTxt = "Kunde"
                    case 5
                     useasfakTxt = "CRM-relation"
                    case 6
                     useasfakTxt = "Leverandør"
                    case 7
                    useasfakTxt = "Underleverandør"
                    
                    case else
                    useasfakTxt = ""
                    end select

                     response.Write useasfakTxt
                    
                     %></td>
                <td><%=left(oRec("segment"),10) %></td>
                <td style="white-space:nowrap;"><%=oRec("cvr") %></td>
                <td style="white-space:nowrap; text-align:right;">
                    <% 
                    Antalx = 0 
                    Antaly = 0
                    strSQL2 = "SELECT id, jobstatus FROM job WHERE jobknr = "& oRec("Kid")
			        oRec2.open strSQL2, oConn, 3
			        while not oRec2.EOF 
			        
			        Antalx = Antalx + 1
			        if oRec2("jobstatus") = 1 then
			        Antaly = Antaly + 1
			        end if
			
			        oRec2.movenext
			        wend
			        oRec2.close
			        Antaljob = Antalx
			        Aktive = Antaly
                    %>
                    <%if Antalx <> 0 then  %> 
                    <a href="../timereg/jobs.asp?menu=job&FM_kunde=<%=oRec("Kid") %>"><span style="color:#999999"><%=Antaljob%></span> 
                        <%if Aktive <> 0 then  %>
                        (<%=Aktive %>)
                        <%end if %>

                    
                    </a>
                    <%if cint(oRec("useasfak")) <= 1 then 'DER KAN KUN OPRETTES JOB på kunder %>
                    <a href="../timereg/jobs.asp?menu=job&func=opret&FM_kunde=<%=oRec("Kid") %>&step=2&int=1&rdir=kon">+</a>
                    <% 
                        
                        end if %>

                    <%if Antalx = 0 AND cint(oRec("useasfak")) <= 1 then 'DER KAN KUN OPRETTES JOB på kunder  %> 
                    <a href="../timereg/jobs.asp?menu=job&func=opret&FM_kunde=<%=oRec("Kid") %>&step=2&int=1&rdir=kon">+</a>
                    <% 
                        end if
                    end if 'antal x %>
                   


                </td>
                <td style="text-align:right;">
                    <%
                    antalfak = 0
                    call antalFakturaerKid(oRec("kid")) %>

                    <%if cint(antalfak) <> 0 then %>
			            <%if (level <= 2 OR level = 6) then %>
			            <a href="erp_fakhist.asp?FM_kunde=<%=oRec("kid") %>" class=vmenu target="_blank"><%=antalfak %></a> stk.
			            <%else %>
			            <%=antalfak %> stk.
			            <%end if %>
                    <%end if %>
                </td>
              </tr>
              <%
                  
                  cnt = cnt + 1
                  oRec.movenext
                  wend 
                  oRec.close 
                  %>
            </tbody>

            
            <tfoot>
               <tr>
                <th>Kundenr.</th>
                <th>Navn</th>
                <th>Postnr.</th>
                <th>By</th>
                <th>Land</th>
                <th>Telefon</th>
                <th>Type</th>
                <th>Segment</th>
                <th>CVR nr.</th>
                <th>Job</th>
                <th>Faktura</th>
                </tr>
            </tfoot>
             
          </table>


          
            <br /><br />
                
            <section>
                <div class="row">
                     <div class="col-lg-12">
                        <b>Funktioner</b>
                        </div>
                    </div>
                  <form action="../timereg/kunder_eksport.asp?kpers=1" method="post" target="_blank">
                      <input id="Hidden2" name="kids" value="<%=kids%>" type="hidden" />
                     <div class="row">
                     <div class="col-lg-12 pad-r30">
                         
                    <input id="Submit5" type="submit" value="A) Eksport til csv. detail (incl. kontaktpers.)" class="btn btn-sm" /><br />
                    <!--Eksporter viste kunder og kontaktpersoner som .csv fil-->
                         
                         </div>


                </div>
                </form>
                <form action="../timereg/kunder_eksport.asp?kpers=0" method="post" target="_blank">
                    <input id="Hidden3" name="kids" value="<%=kids%>" type="hidden" />
                 <div class="row">
                     <div class="col-lg-12 pad-r30">
                         <input id="Submit6" type="submit" value="B) Eksport til .csv simpel" class="btn btn-sm" />
                        <!--Eksporter viste kunder som .csv fil-->
                         
                         </div>


                </div>
                    </form>
            </section>








        </div> <!-- /.portlet-body -->
      </div> <!-- /.portlet -->
        


<%end select  %>



</div><!-- cointainer -->
</div><!--content -->
</div><!-- wrapper -->



<!--#include file="../inc/regular/footer_inc.asp"-->