<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/kontakter_func.asp"-->
<!--#include file="../inc/regular/crm_func.asp"-->
<!-- #include file = "CuteEditor_Files/include_CuteEditor.asp" --> 
<!--#include file="inc/isint_func.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	'*** Faneblade, valgt **********
	'***** Hvilken div skal vises ****
	if len(request("showdiv")) <> 0 then
	showdiv = request("showdiv")
	else
	showdiv = "stamdata"
	end if
	
	if len(request("hidemenu")) <> 0 then
	hidemenu = request("hidemenu")
	else
	hidemenu = 0
	end if
	
	if len(request("rdir")) <> 0 then
	rdir = request("rdir")
	else
	rdir = 0
	end if
	
	if len(request("visikkekunder")) <> 0 then
	visikkekunder = 1
	else
	visikkekunder = 0
	end if
	
	func = request("func")
	menu = request("menu")
	ketype = request("ketype")
	
	if func = "opret" OR func = "dbopr" then
		if menu = "crm" then
		redirectID = request("redirectid")
		id = 0
		else
		redirectID = request("id")
		id = 0
		end if
	else
		if menu = "crm" then
		redirectID = request("redirectid")
		id = request("id")
		else
		id = request("id")
		redirectID = id
		end if
	end if
	
	dim content
	
	if ketype = "e" then
	ketype = ketype
	else
	ketype = "k"
	end if
	
	if len(request("FM_soeg")) <> 0 OR len(request("ktype")) <> 0 then 
	'** request("ktype") tjekker at form e submitted **'
	thiskri = request("FM_soeg")
	useKri = 1
	    
	    	response.cookies("tsa")("kundesog") = thiskri
	        response.cookies("tsa").expires = date + 32
	    
	else
	        
	        if request.cookies("tsa")("kundesog") <> "" then
	        thiskri = request.cookies("tsa")("kundesog")
	        useKri = 1
	        visikkekunder = 0
	        else
	        thiskri = ""
	        useKri = 0
	        end if
	        
	end if
	
	selmedarb = request("medarb")
	medarb = selmedarb
	
	if len(request("ktype")) <> 0 then
	ktype = request("ktype")
	else
	ktype = -1
	end if
	
	
	
	
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
	Response.redirect "kunder.asp?menu=crm&shokselector=1&ketype=e&selpkt=osigt&redirectid="&thiskid&"&FM_soeg="&thiskri&"&medarb="&medarb
	
	
	
	case "slet"
	'*** Her spørges om det er ok at der slettes en medarbejder ***
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	slttxtalt = ""
	slturlalt = ""
	
	slttxt = "Du er ved at <b>slette</b> en <b>kontakt</b>. Er dette korrekt?<br>"_
	&"Du vil samtidig slette alle job, og dertil tilhørende <br><b>timeregistreringer, aktiviteter, ressourcetimer, kontaktpersoner, aktions historik mm.<b/><br>Data kan ikke genskabes!"
    slturl = "kunder.asp?menu="&menu&"&ketype="&ketype&"&func=sletok&id="&id&"&FM_soeg="&thiskri&"&medarb="&medarb
	
	call sltque(slturl,slttxt,slturlalt,slttxtalt,210,100)
	
	
	
	case "sletok"
	'*** Her slettes en kunde ***'
	
	'** Sletter tilhørende job, aktiviteter og timer  
	
	
	strSQL = "SELECT id FROM job WHERE jobknr = "& id
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
	oConn.execute("DELETE FROM job WHERE jobknr = "& oRec("id")  &"")
	
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
	
	
	
	Response.redirect "kunder.asp?menu="&menu&"&ketype="&ketype&"&shokselector=1&func=hist&selpkt=osigt&FM_soeg="&thiskri&"&medarb="&medarb
	
	
	case "sletfil"
	
	'*** Her spørges om det er ok at der slettes en fil ***
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:190; top:50; visibility:visible;">
	<br><br><br>
	<table cellspacing="2" cellpadding="2" border="0">
	<tr>
	    <td><img src="../ill/alert.gif" width="44" height="45" alt="" border="0">Du er ved at <b>slette</b> en fil. Er dette korrekt?</td>
	</tr>
	<tr>
	   <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	   <a href="kunder.asp?menu=kund&func=sletfilok&id=<%=id%>&filnavn=<%=request("filnavn")%>&kundeid=<%=request("kundeid")%>&FM_soeg=<%=thiskri%>&medarb=<%=medarb%>">Ja</a>&nbsp;&nbsp;&nbsp;<a href="javascript:history.back()">Nej</a></td>
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
		%>
		<!--#include file="../inc/regular/header_inc.asp"-->
		<%
		errortype = 13
		call showError(errortype)
			
			else
			
			if len(trim(request("FM_knr"))) = 0 OR len(request("FM_knr")) > 15 then
			%>
			<!--#include file="../inc/regular/header_inc.asp"-->
			<%
			errortype = 21
			call showError(errortype)
			
			else
					
					strKnr = request("FM_knr")
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

					%>
					<!--#include file="../inc/regular/header_inc.asp"-->
					<%
				
					errortype = 12
					call showError(errortype)
					
					else
				
				
				if request("FM_oprsomkunde") = "on" then
				dbketype = "k"
					if request("FM_oprsomfirma") = "on" then
					dbketype = "ke"
					end if
				else
					if request("FM_oprsomfirma") = "on" then
				 	dbketype = "e"
					else
					dbketype = "n"
					end if
				end if
				
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
				strEan = request("FM_ean")
				strWWW = request("FM_www")
				
				if len(request("FM_nace")) <> 0 then
				strNACE = request("FM_nace")
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
				
				
				'* undersøg om det er int
				if len(request("FM_regnr")) <> 0 then
				intRegnr = request("FM_regnr")
				else
				intRegnr = 0
				end if
				
				if len(request("FM_kontonr")) <> 0 then
				intKontonr = request("FM_kontonr")
				else
				intKontonr = 0
				end if
				
				if len( request("FM_cvr")) <> 0 then
				intCvr = request("FM_cvr")
				else
				intCvr = 0
				end if
				
				strBank = request("FM_bank")
				strIban = request("FM_iban")
				strSwift = request("FM_swift")
				
				if len(request("FM_useasfak")) <> "0" then
				intuseasfak = request("FM_useasfak")
				else
				intuseasfak = 0
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
				
				
				if intuseasfak = 1 then
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
				& "cvr, "_
				& "bank, "_
				& "useasfak, "_
				& "logo, swift, iban, ktype, kundeans1, "_
				&" kundeans2, nace,"_
				&" sdskpriogrp "& strSQLinsflds &") VALUES "_
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
				& "'" & intCVR &"', "_
				& "'" & strBank &"', "_
				& "" & intuseasfak &", "_
				& "" & logo &", '"& strSwift &"', '"& strIban &"', "_
				& "" & intKtype &", "& intKundeans1 &", "& intKundeans2 &","_
				&" '"& strNACE &"', "_
				&" "& priogrp &" "& strSQLinsval &")"
				
				
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
				& "cvr = '" & intCvr & "', "_
				& "bank = '" & strBank & "', "_
				& "useasfak = " & intuseasfak & ", "_
				& "logo = "& logo &", "_
				& "swift = '" & strSwift &"', "_
				& "iban = '" & strIban &"', ktype = "& intKtype &", "_
				& "kundeans1 = "& intKundeans1 &", kundeans2 = "& intKundeans2 &", nace = '"& strNACE &"', "_
				& "sdskpriogrp = "& priogrp 
				
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
				
				
				if len(thiskri) <> 0 AND func <> "dbopr" then
				thiskri = thiskri
				else
				thiskri = left(strNavn, 1)
				end if
				
				
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
				
				
				
				
				if menu = "crm" then
				Response.redirect "kunder.asp?menu=crm&shokselector=1&ketype=e&redirectid="&thisKid&"&FM_soeg="&thiskri&"&medarb="&medarb
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
						        Kontakt oversigten ved at bruge nedenstående link.<br>
						        Kontoen er oprettet korrekt når der er klikket på "opret" og det viste popup vindue ikke længere er synligt.<br>
						        <br>
						        <%
						        Response.write "<a href='kunder.asp?menu=kund&shokselector=1&ketype=k&id="&thisKid&"&FM_soeg="&thiskri&"&medarb="&medarb&"'>Tilbage til kontaktoversigten</a>"
						        %></td></tr></table>
						        </div>
						        <%
						        end if
						
						else
						
						if rdir = "1" then 'on the fly fra job, ell. SDSK
						Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
	                    Response.Write("<script language=""JavaScript"">window.close();</script>")
						else
						Response.redirect "kunder.asp?menu=kund&shokselector=1&ketype=k&id="&thisKid&"&FM_soeg="&thiskri&"&medarb="&medarb
						end if
						
						end if
				end if
			
			
			
				
				
			end if
		end if
	end if
		
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af kunde ***
	
	
	
	
	
	
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
	varSubVal = "opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	
	prio_grp = 1
	
	strLevbet = ""
	strBetbet = ""
	
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
	& "cvr, "_
	& "bank, swift, iban, "_
	& "useasfak, "_
	& "logo, ktype, kundeans1, kundeans2, nace, betbet, levbet, sdskpriogrp, betbetint "_
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
	intCVR = oRec("cvr")
	strBank = oRec("bank")
	strIban = oRec("iban")
	strSwift = oRec("swift")
	logo = oRec("logo")
	intKtype = oRec("ktype")
	strNACE = oRec("nace")
	
	intKundeans1 = oRec("kundeans1")
	intKundeans2 = oRec("kundeans2")
	
	prio_grp = oRec("sdskpriogrp")
	
	if oRec("useasfak") = 1 then
	eruseasfak = "y"
	useasfakch = "CHECKED"
	else
	useasfakch = ""
	end if
	
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
	
	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "opdaterpil" 
	end if
	%>
	
	<script>
	function expand(div) {
		if (document.all(div).style.display == "none"){
			document.all(div).style.display = "";
		}else{
			document.all(div).style.display = "none";
		}
	}
	
	</script>
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	<%if hidemenu <> 1 then %>
	
	        
	    
	       
	        <!--#include file="../inc/regular/topmenu_inc.asp"-->
	        <div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	        
	       
	        <%if menu <> "crm" then%>
	        <%call tsamainmenu(5)%>
	        <%else%>
	        <%call crmmainmenu(2)%>
	        <%end if%>
	        </div>
	        
	        
	        <div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	        <%
	        call kundtopmenu()
	        %>
	        </div>
	        
	        <%
	        dleft = "20"
	        dtop = "80"
	        
	      
    else 
	
	
	dleft = "20"
	dtop = "60"
	
	end if %>
	
	<%call kundemenu(showdiv)%>
	
	
	
	
	<div id="stamdata" name="stamdata" style="position:relative; left:<%=dleft%>px; top:<%=dtop%>px; visibility:<%=stamdiv_vzb%>; display:<%=stamdiv_dsp%>; z-index:10000; width:800px;  background-color:#ffffff; padding:3px 3px 3px 3px; border:1px #8caae6 solid;">
	
	
	<table cellspacing="0" cellpadding="0" border="0" width="100%" bgcolor="#EFF3FF">
	<form action="kunder.asp?func=<%=dbfunc%>" method="post">
	<input type="hidden" name="menu" value="<%=menu%>">
	<input type="hidden" name="id" id="id" value="<%=id%>">
	<input type="hidden" name="ketype" value="<%=ketype%>">
	<input type="hidden" name="FM_soeg" value="<%=thiskri%>">
	<input type="hidden" name="medarb" value="<%=medarb%>">
	<input type="hidden" name="rdir" value="<%=rdir%>">
	
	<tr bgcolor="#5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/blank.gif" width="8" height="32" alt="" border="0"></td>
		<td colspan=3 valign="top"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/blank.gif" width="8" height="32" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
	<%
	call erSDSKaktiv()
	
	if dbfunc = "dbred" then
	    
	    if dsksOnOff = 1 then
	    rowspanThis = 26
	    else
		rowspanThis = 25
		end if
		%>
		<td colspan="5" class='alt'>Sidst opdateret den <b><%=strLastUptDato%></b> af <b><%=strEditor%></b>
		</td>
		</tr>
		<tr>
		<td colspan="5" align=right style="padding:5px 50px 5px 5px;">
		<%call antalFakturaerKid(intKid) %>
		<%if antalFak = 0 then %>
		<a href="kunder.asp?menu=<%=menu %>&func=slet&id=<%=intKid%>&ketype=e&FM_soeg=<%=thiskri%>&medarb=<%=medarb%>"><img src="../ill/slet_16.gif" alt="Slet" border="0"></a>
		<%end if %>
		</td>
		
		
	<%else
		if dsksOnOff = 1 then
	    rowspanThis = 27
	    else
		rowspanThis = 26
		end if
		%>
		<td colspan=5 class=alt><b>Stamdata</b>&nbsp;&nbsp;</td>
	<%end if%>
	</tr>
	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td rowspan="<%=rowspanThis%>" style="width:10px;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td colspan=2><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top" rowspan="<%=rowspanThis-1%>"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td style="padding-top:12px;" valign=top width=175><b><font class=roed>*</font>&nbsp;Navn:</b>&nbsp;&nbsp;</td>
		<td style="padding-top:8px;" class=lille><input type="text" name="FM_navn" value="<%=strNavn%>" size="55" style="border:1px #86B5E4 solid;"><br />
		" (situationstegn) er ikke tilladt i et kontaktnavn.</td>
		<td valign="top" rowspan="<%=rowspanThis-1%>"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    </tr>
	<tr>
		<td>Type / Segment og Rabat %:</td>
		<td>
		<select name="FM_type" id="FM_type" style="width:260px;">
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
		</td></tr>
	<tr>
		<td valign=top style="padding-top:5px;"><b><font class=roed>*</font>&nbsp;Kontakt id:</b><br />
		CVR, Tel. etc.
		
		<%if func <> "red" then %>
		<br />
		<font class="lillesort">
		<br />Sidst oprettet kontakt havde id:<br />
		<%strSQL = "SELECT kkundenr FROM kunder WHERE kid <> 0 ORDER BY kid DESC"
		oRec.open strSQL, oConn, 3
		lastUsedKnr = "0"
		if not oRec.EOF then
		lastUsedKnr = oRec("kkundenr") 
		end if
		oRec.close
		
		Response.Write "<b>"& lastUsedKnr & "</b></font>"
		
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
		
		
		end if
		%>
		</td>
		<td class=lille>
		<%
		select case lto
		case "execon", "immenso"
		 if level = 1 then
		 kidAaben = 1
		 else
		 kidAaben = 0
		 end if
		case else
		 kidAaben = 1
		end select
		
		if kidAaben = 1 then %>
		<input type="text" name="FM_knr" value="<%=strKnr%>" size="15">
		<%else %>
		<input type="text" name="FM_knrdis" value="<%=strKnr%>" DISABLED size="15">
		<input type="hidden" name="FM_knr" value="<%=strKnr%>">
		<%end if %>
		<br />Kan indeholde både tal og bogstaver, maks 15 kar. 
		<br />Ved brug af FI nummer må kunde Id maks være på 8 karakterer.
		</td>
	</tr>
	<tr>
		<td valign=top><br />Adresse:<br />
		c/o</td>
		<td class=lille valign=top><br />
            <textarea id="FM_adr" name="FM_adr" cols="40" rows="3"><%=strAdr%></textarea><br />
            NB: linjeskift angives som: &lt;br&gt;
            <br />&nbsp;</td>
	</tr>
	<tr>
		<td>Postnr:</td>
		<td><input type="text" name="FM_postnr" value="<%=strPostnr%>" size="10" style="border:1px #86B5E4 solid;"></td>
	</tr>
	<tr>
		<td>By:</td>
		<td><input type="text" name="FM_city" value="<%=strCity%>" size="40" style="border:1px #86B5E4 solid;"></td>
	</tr>
	<tr>
		<td>Land:</td>
		<td><select name="FM_land" style="border:1px #86B5E4 solid; width:260px;">
		<%if func = "red" then%>
		<option SELECTED><%=strLand%></option>
		<%else%>
		<option SELECTED>Danmark</option>
		<%end if%>
		<!--#include file="inc/inc_option_land.asp"-->
		</select>
		<!--<input type="text" name="FM_land" value="<=strLand%>" size="20" style="border:1px #86B5E4 solid;">--></td>
	</tr>
	<tr>
		<td>Tlf:</td>
		<td><input type="text" name="FM_tlf" value="<%=strTlf%>" size="12" style="border:1px #86B5E4 solid;"></td>
	</tr>
	<tr>
		<td>Mobil:</td>
		<td><input type="text" name="FM_mobil" value="<%=strMobil%>" size="12" style="border:1px #86B5E4 solid;"></td>
	</tr>
	<tr>
		<td>Alt. tlf:</td>
		<td><input type="text" name="FM_alttlf" value="<%=strAlttlf%>" size="12" style="border:1px #86B5E4 solid;"></td>
	</tr>
	<tr>
		<td>Fax:</td>
		<td><input type="text" name="FM_fax" value="<%=strFax%>" size="12" style="border:1px #86B5E4 solid;"></td>
	</tr>
	<tr>
		<td>Email:</td>
		<td><input type="text" name="FM_email" value="<%=strEmail%>" size="25" style="border:1px #86B5E4 solid;"></td>
	</tr>
		<tr>
		<td>EAN:</td>
		<td><input type="text" name="FM_ean" value="<%=strEan%>" size="25" style="border:1px #86B5E4 solid;"></td>
	</tr>
	<tr>
		<td valign=top>WWW adresse(r) / Domæner:</td>
		<td><textarea id="FM_www" name="FM_www" style="height:80px; width:350px;"><%=strWWW%></textarea></td>
	</tr>
    <tr>
		<td>CVR nr:</td>
		<td><input type="text" name="FM_cvr" value="<%=intCVR%>" size="25"></td>
	</tr>
	<tr>
		<td>NCA kode:</td>
		<td><input type="text" name="FM_nace" value="<%=strNACE%>" size="25"></td>
	</tr>
    <tr>
		<td>Bank:</td>
		<td>
        <input type="text" name="FM_bank" value="<%=strBank%>" size="25" style="border:1px yellowgreen solid;">
    </td>
	</tr>
	<tr>
		<td>Swift kode:</td>
		<td><input type="text" name="FM_swift" value="<%=strSwift%>" size="25" style="border:1px yellowgreen solid;"></td>
	</tr>
	<tr>
		<td>Iban kode:</td>
		<td><input type="text" name="FM_iban" value="<%=strIban%>" size="25" style="border:1px yellowgreen solid;"></td>
	</tr>
	<tr>
		<td>Konto:</td>
		<td>Regnr:&nbsp;<input type="text" name="FM_regnr" value="<%=intRegnr%>" size="4" style="border: 1px yellowgreen solid;"><br />
		Kontonr:&nbsp;<input type="text" name="FM_kontonr" value="<%=intKontonr%>" size="10" style="border: 1px yellowgreen solid;"></td>
	</tr>
	
	<%if menu <> "crm" then%>
	<tr>
		<td colspan="2">Denne kontakt skal stå som afsender på fakturaer: 
		<input type="checkbox" name="FM_useasfak" value="1" <%=useasfakch%> style="border:1px #86B5E4 solid;">&nbsp;(Licens indehaver)</td>
	</tr>
	<%else%>
	<tr>
		<td colspan=2>&nbsp;</td>
	</tr>
		<%'*** Husker om firma står som afsender på fak nå det redigeres fra CRM. ***
		if eruseasfak = "y" then%>
		<input type="hidden" name="FM_useasfak" value="1">
		<%end if%>
	<%end if%>
	
	<tr>
		<td colspan="2" style="padding-top:20px;" valign=top><b>Kontaktansvarlige / Keyaccount</b><br>
		<font class=megetlillesort>Angiv op til 2 kontaktansvarlige. Alle kan redigere i kundeoplysninger uanset ansvarlig.</font><br><br>
		<img src="../ill/ac0019-16.gif" width="16" height="16" alt="" border="0">&nbsp;<b>Kontaktansvarlig 1:</b>
		&nbsp;&nbsp;<select name="FM_kundeans_1" id="FM_kundeans_1">
		<option value="0">Ingen</option>
			<%
			
			'if func <> "red" then
			strSQL = "SELECT mnavn, mnr, mid FROM medarbejdere WHERE mansat <> 2"
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
			<option value="<%=oRec("mid")%>" <%=medsel%>><%=oRec("mnavn")%> (<%=oRec("mnr")%>)</option>
			<%
			oRec.movenext
			wend
			oRec.close 
			%>
		</select>
		<br><br>
		<img src="../ill/ac0012-16.gif" width="16" height="16" alt="" border="0">&nbsp;<b>Kontaktansvarlig 2:</b>
		&nbsp;&nbsp;<select name="FM_kundeans_2" id="FM_kundeans_2">
		<option value="0">Ingen</option>
			<%
			
			'if func <> "red" then
			strSQL = "SELECT mnavn, mnr, mid FROM medarbejdere WHERE mansat <> 2"
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
			<option value="<%=oRec("mid")%>" <%=medsel2%>><%=oRec("mnavn")%> (<%=oRec("mnr")%>)</option>
			<%
			oRec.movenext
			wend
			oRec.close 
			%>
		</select><br>&nbsp;</td>
	</tr>
	
	<%
	antalfoldere = 0
	strSQL = "SELECT count(id) AS antalfoldere FROM foldere WHERE kundeid = "& id &" AND stfoldergruppe = 0"
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
		antalfoldere  = oRec("antalfoldere")
	end if 
	oRec.Close					
	%>
	<tr>
	<td colspan=2><img src="../ill/ac0056-16.gif" width="16" height="16" alt="" border="0">&nbsp;<b>Tilknyt Standardfoldere?</b><br>
	Der er findes allerede: <b><%=antalfoldere%> folder(e)</b> på denne kunde.<br>
	<font class=megetlillesort> Hvis du ønsker at tilknytte flere foldere kan det enten gøres ved at vælge en ny Standardfolder gruppe her,
	<br> eller ved manuelt at tilføje foldere under fanebladet "Filer".<br></font>
	Tilknyt gruppe:&nbsp;<select name="FM_stfoldergruppe" id="FM_stfoldergruppe" style="width:200px;">
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
		</select><br>&nbsp;
	
	</td>
	</tr>
	<%if func = "opret" then%>
	<tr>
	<td colspan=2><img src="../ill/ac0044-16.gif" width="16" height="16" alt="" border="0">&nbsp;<b>Opret konto?</b>
	<input type="checkbox" name="FM_opretkonto" id="FM_opretkonto" value="1"> Ja, jeg ønsker at tilknytte en konto til denne kontakt.
	</td>
	</tr>
	<%end if%>
	
	
	<tr>
		<td colspan=2><br>
		<%
		'*** tjekke brugergruppe rettigheder ***
		'*** Hvis man er i CRM delen og ikke har rettigheds level <= 3 or = 6 
		'*** Så bliver man automatisk kun oprettet i CRM delen.
		'*** Brugergruppe 4 og 5 har ikke adgang til kunder i TSA. Derfor ingen probs her ***
		if level <= 3 OR level = 6 then%>
		<b>TSA / CRM:</b><br>
		Opret/tilføj som kontakt i TSA systemet: &nbsp; <input type="checkbox" name="FM_oprsomkunde" value="on" <%=kchecked%> style="border:1px #86B5E4 solid;"> <br>og/eller som firma i CRM systemet &nbsp;<input type="checkbox" name="FM_oprsomfirma" value="on" <%=fchecked%> style="border:1px #86B5E4 solid;"></td>
		<%else%>
		<input type="hidden" name="FM_oprsomfirma" value="on">
		<%end if%>
	</tr>
	
	
	<%
	if dsksOnOff = 1 then
	%>
	<tr>
		<td valign=top colspan=2><br><b>ServiceDesk Aftalegruppe:</b><br>
		<font class=megetlillesort>Når der er valgt en gruppe kan der oprettes incindets på denne kontakt.</font><br>
		<select name="FM_prio_grp" id="FM_prio_grp">
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
	</select></td>
	</tr>
	<%end if%>
	
	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/blank.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan=3 valign="bottom"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/blank.gif" width="8" height="10" alt="" border="0"></td>
	</tr>
	</table>
	
	
	<table width=100% cellpadding=20 cellspacing=0 border=0 bgcolor="#EFF3FF">
	    <tr>
	        <td align=right><br /><br />&nbsp;
                <input id="Submit4" type="submit" value=" Opdater >> " /></td>
	    </tr>
	</table>
	</div>
	

	
	
	
	
	<%
	if showdiv = "klog" then
	
	'**** Log ***********************************************************************
	%>
	<div id="log" name="log" style="position:relative; left:20px; top:80px; visibility:<%=klogdiv_vzb%>; display:<%=klogdiv_dsp%>; width:800px; background-color:#ffffff; padding:3px 3px 3px 3px; border:1px #8caae6 solid;">
	
	
	<table cellspacing="0" cellpadding="0" border="0" width="100%" bgcolor="#EFF3FF">
	<tr bgcolor="#5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/blank.gif" width="8" height="32" alt="" border="0"></td>
		<td valign="top"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/blank.gif" width="8" height="32" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td class='alt'>Beskrivelse, Leveringsbetingelser og Betalingsbetingelser.</td>
	</tr>
	<tr>
		<td valign="top" rowspan=3><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td valign=top><br><br><b>Beskrivelse:</b><br>
		<textarea id="TextArea2" name="FM_komm" cols="90" rows="10"><%=strKomm %></textarea>
		
		
		                <br />
            &nbsp;
		
		
		
		</td>
		<td valign="top" align="right" rowspan=3><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr>
	
	<td valign=top valign=top><b>Forfaldsdato: (kredit)</b><br />
		
	<% 
	 select case lto
	    case "execon", "immenso"
	    betbetint = 8
	    disa = 1
	    case else
	    betbetint = betbetint
	    disa = 0
	    end select
	
	call betalingsbetDage(betbetint, disa)
	%>
            
            
            <br /><br /><b>Betalingsbetingelser:</b><br />
        
            
             <%
	                    content = strBetbet
            			
			            
			            Set editorB = New CuteEditor
            					
			            editorB.ID = "FM_betbet"
			            editorB.Text = content
			            editorB.FilesPath = "CuteEditor_Files"
			            editorB.AutoConfigure = "Minimal"
            			
			            editorB.Width = 740
			            editorB.Height = 280
			            editorB.Draw()
		                %>
		                
    </td>
	</tr>
	<tr>
		<td valign=top><br /><br /><b>Leveringsbetingelser:</b><br>
            <textarea id="TextArea1" name="FM_levbet" cols="90" rows="10"><%=strLevbet %></textarea>
		
		                <br />
            &nbsp;
		
		
		</td>
	</tr>
	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/blank.gif" width="8" height="10" alt="" border="0"></td>
		<td valign="bottom"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/blank.gif" width="8" height="10" alt="" border="0"></td>
	</tr>
	</table>
	
	<br /><br />
	
	<table width=100% cellpadding=20 cellspacing=0 border=0 bgcolor="#EFF3FF">
	    <tr>
	        <td align=right>
                <input id="Submit3" type="submit" value=" Opdater >> " /></td>
	    </tr>
	</table>
	
	</div>
	
	
	<input id="Hidden1" name="FM_opdater_txt_felter" value="1" type="hidden" />
	<%
	else
	%>
    <input id="FM_opdater_txt_felter" name="FM_opdater_txt_felter" value="0" type="hidden" />
	<%
	end if
	
	
	'****************************** Logoer ***********************************'
	
	%>
	<div id="logoer" name="logoer" Style="position:relative; left:20px; top:80px; visibility:<%=logodiv_vzb%>; display:<%=logodiv_dsp%>; width:800px; background-color:#ffffff; padding:3px 3px 3px 3px; border:1px #8caae6 solid;">
	<table cellspacing="0" cellpadding="0" border="0" width="100%" bgcolor="#EFF3FF">
	<tr bgcolor="#5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/blank.gif" width="8" height="32" alt="" border="0"></td>
		<td colspan=3 valign="top"><img src="../ill/blank.gif" width="584" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/blank.gif" width="8" height="32" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td colspan=2 class='alt'>Logoer:<br />
		Maks:  bredde: 150 px * højde 100 px
		</td>
		<td align=right>
		<a href="javascript:popUp('upload.asp?type=kundelogo&kundeid=<%=id%>&jobid=0', '600', '600','200', '50')" target="_self" class='alt'>Tilknyt nyt logo&nbsp;<img src="../ill/pillillexp.gif" width="16" height="18" alt="" border="0"></a>
		</td>
	</tr>
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
	Response.write "<tr><td valign=top><img src='../ill/blank.gif' width='1' height='45' alt='' border='0'></td>"
	Response.write "<td><input type='radio' name='FM_logo' value='"& oRec("id") &"' "& logosel &">&nbsp;&nbsp;"& oRec("filnavn") & "</td>"
	Response.write "<td>&nbsp;<img src='../inc/upload/"&lto&"/"&oRec("filnavn")&"' width=80 height=30 alt='' border='0'></td>"
	Response.write "<td><a href='kunder.asp?func=sletfil&id="&oRec("id")&"&kundeid="& id &"&filnavn="&oRec("filnavn")&"'><img src='../ill/slet_16.gif' alt='Slet' border='0'></a></td>"
	Response.write "<td valign=top align=right><img src='../ill/blank.gif' width='1' height='45' alt='' border='0'></td></tr>"
	
	j = j + 1
	end if
	oRec.movenext
	wend
	oRec.close
	
	if j = 0 then
	Response.write "<tr><td valign=top><img src='../ill/blank.gif' width='1' height='45' alt='' border='0'></td>"
	Response.write "<td colspan=3>&nbsp;&nbsp;-- Ingen logoer er tilknyttet denne kunde. --</td>"
	Response.write "<td valign=top align=right><img src='../ill/blank.gif' width='1' height='45' alt='' border='0'></td></tr>"
	
	else
	
	if erfundet <> 1 then
	logoselingen = "CHECKED"
	else
	logoselingen = ""
	end if
	
	Response.write "<tr><td valign=top><img src='../ill/blank.gif' width='1' height='45' alt='' border='0'></td>"
	Response.write "<td colspan=3><input type='radio' name='FM_logo' value='0' "& logoselingen &">&nbsp;Intet logo valgt.</td>"
	Response.write "<td valign=top align=right><img src='../ill/blank.gif' width='1' height='45' alt='' border='0'></td></tr>"
	
	end if
	%>
	
	<tr bgcolor="#EFF3FF">
		<td valign="top" style="border-left:1px;"><img src="../ill/blank.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan=3 valign="bottom"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/blank.gif" width="8" height="10" alt="" border="0"></td>
	</tr>
	</table>
	
	
	<br /><br />
	<table width=100% cellpadding=10 cellspacing=0 border=0 bgcolor="#EFF3FF">
	    <tr>
	        <td align=right>
                <input id="Submit2" type="submit" value=" Opdater >> " /></td>
	    </tr>
	</table>
	
	</div>
	
	<br /><br /><br />&nbsp;
	
	
	
	
	
	
	
	
	<%case else
	thisfile = "kunder"
	%>
	
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	
	
	
	
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%if menu <> "crm" then%>
	<%call tsamainmenu(5)%>
	<%else%>
	<%call crmmainmenu(2)%>
	<%end if%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	call kundtopmenu()
	%>
	</div>
	
	
	
	
	
	
	
	
	<%if menu = "crm" then
	'*********************   Funktioner  CRM   ***************************************
	
	public y
	public x
	public showkunde
	function hentfirmaer()
		
		if instr(showkunde, "#"& oRec("Kid") &"#") = 0 then%>
		<tr>
			<td bgcolor="#CCCCCC" colspan="9"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
		</tr>
		<%
		'**** Markerer den sidst redigerede kunde **************
						
						if cint(oRec("kid")) = cint(redirectID) then
						bgthis = "#ffff99"
						else
			            bgthis = bgthis
						end if
					
		
		%>
		<tr bgcolor="<%=bgthis%>">
		<td valign="top"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		
		<td valign="top" style="padding:5px 5px 5px 5px;"><%=oRec("Kkundenr")%></td>
		
		<td valign="top" height="20" style="padding:5px 0px 3px 0px;"><a href="kunder.asp?menu=<%=menu%>&func=red&id=<%=oRec("Kid")%>&FM_soeg=<%=thiskri%>&medarb=<%=medarb%>">
		<%=oRec("Kkundenavn")%>
		</a>
		
		<%
		if oRec("useasfak") = 1 then
		%>
		<br><img src="../ill/pil_kalender.gif" width="10" height="12" alt="" border="0">&nbsp;(Licens indehaver)
		<%
		end if
		%>
		
		
		<br />
	
		<b><%=oRec("adresse") %><br />
		<%=oRec("postnr") %>, <%=oRec("city") %></b><br />
		<%=oRec("land") %>
		
		<%if len(trim(oRec("telefon"))) <> 0 then %>
		<br />Tel: <%=oRec("telefon")%>
		<%
		end if%>
		
		<%if len(trim(oRec("cvr"))) <> 0 then %>
		<br />CVR: <%=oRec("cvr")%>
		<%end if%>
		
		<%if len(trim(oRec("ean"))) <> 0 then %>
		<br />EAN: <%=oRec("ean")%>
		<%
		end if%>
		
		
		
		
		<font class=medlillesilver>
		<%if len(trim(oRec("kans1"))) <> 0 then%>
		<br>
		<%=oRec("kans1")%> (<%=oRec("mnr1")%>)
		        
		        <%if len(oRec("kans1_init")) <> 0 then %>
		        - <%=oRec("kans1_init")%>
		        <%end if %>
		        
		<%end if%>
		
		<%if len(trim(oRec("kans2"))) <> 0 then%>
		&nbsp;/&nbsp;
		<%=oRec("kans2")%> (<%=oRec("mnr2")%>)
		    
		    <%if len(oRec("kans2_init")) <> 0 then %>
		     - <%=oRec("kans2_init")%>
		    <%end if %>
		    
		<%end if%>
		</font>
		
		
		</td>
		<%if menu = "crm" then%>
			<%select case oRec("hot")
			case 0
			hotimg = "blank"
			case -1
			hotimg = "crm_kold"
			case -2
			hotimg = "crm_nej"
			case 1
			hotimg = "crm_varm"
			case 2
			hotimg = "crm_hot"
			case else
			hotimg = "blank"
			end select
			%> 
			<td valign=top style="padding-top:3px;"><b><%=oRec("hot") %></b> <img src="../ill/<%=hotimg%>.gif" width="20" height="10" alt="" border="0">&nbsp;
			<a href="kunder.asp?menu=crm&func=addlevel&thiskid=<%=oRec("Kid")%>&hotprev=<%=oRec("hot")%>&status=<%=status%>&id=<%=id%>&emner=<%=emner%>&medarb=<%=medarb%>&selpkt=osigt&FM_soeg=<%=thiskri%>" class=vmenu><b>+ Op</b></a>&nbsp;&nbsp;
			<a href="kunder.asp?menu=crm&func=sublevel&thiskid=<%=oRec("Kid")%>&hotprev=<%=oRec("hot")%>&status=<%=status%>&id=<%=id%>&emner=<%=emner%>&medarb=<%=medarb%>&selpkt=osigt&FM_soeg=<%=thiskri%>" class=vmenu><b>- Ned</b></a></td>
			<td width=220 valign="top" style="padding:5px 5px 5px 5px;">
			
			<%
			strSQL3 = "SELECT crmdato, id, navn FROM crmhistorik WHERE kundeid = "& oRec("Kid") & " ORDER BY crmdato DESC LIMIT 5"
			
			'Response.Write strSQL3 
			'Response.flush
			
			oRec3.open strSQL3, oConn, 3
			while not oRec3.EOF%>
		
			<a href="javascript:NewWin_popupaktion('crmhistorik.asp?menu=crm&func=red&id=<%=oRec("Kid")%>&crmaktion=<%=oRec3("id")%>&showinwin=j')" class=vmenualt>
			<%
			Response.write formatdatetime(oRec3("crmdato"), 2) 
			if len(oRec3("navn")) > 20 then
			Response.Write " | " & left(oRec3("navn"), 20) & ".."
			else
			Response.Write " | " & oRec3("navn")
			end if
			%>
			</a><br />
			<%
			oRec3.movenext
			wend
			oRec3.close %>
			
                &nbsp;</td>
			<td valign="top" style="padding:5px 5px 5px 5px;"><a href="javascript:NewWin_popupaktion('crmhistorik.asp?menu=crm&func=opret&id=<%=oRec("Kid")%>&showinwin=j')" class='vmenu'>Ny-aktion >></a></td>
			<td valign="top" style="padding:5px 5px 5px 5px;" align=right><a href="javascript:NewWin_large('../inc/regular/kundelogview.asp?useKid=<%=oRec("Kid")%>')" target="_self"><img src="../ill/joblog_small.gif" width="12" height="12" alt="Se aktions- og kunde log" border="0"></a>&nbsp;
			<a href="crmhistorik.asp?menu=crm&func=hist&id=<%=oRec("Kid")%>&selpkt=hist"><img src="../ill/hist_lille.gif" width="12" height="12" alt="Se historik" border="0"></a></td>
		<%else 'Else?? Hvordan kan den det når denne function kune hentes hvis menu = crm ????%>
			
			<td valign=top style="padding:5px 5px 5px 5px;">
			<%if oRec("ktype") <> 0 then %>
			<%=oRec("ktnavn") %>
			<%end if %>
			</td>
			<td valign=top style="padding:5px 5px 5px 5px;">
			<%
			strSQL2 = "SELECT navn, email, password, kundeid FROM kontaktpers WHERE kundeid =" &oRec("kid")
			oRec2.open strSQL2, oConn, 3
			kpes = 0
			while not oRec2.EOF
			if len(oRec2("password")) > 0 then 
			kpes = kpes + 1
			end if
			oRec2.movenext
			wend
			oRec2.close
			
			    if kpes > 0 AND (level < 3 OR level = 6) then
    			
			            select case lto
					    case "spritelab", "kringit", "abu", "kits", "outz"
					    %>
					    <a href="javascript:popUp('sdsk.asp?usekview=j&FM_kontakt=<%=oRec("Kid")%>','900','500','20','100')" target="_parent" class="vmenu">Se kontakts egen side >></a>
			            <%
					    case else
				        %>
				        <a href="javascript:popUp('joblog_k.asp?useKid=<%=oRec("Kid")%>','900','500','20','100')" target="_parent" class="vmenu">Se kontakts egen side >></a>
			            <%
				        end select
    			
			    %>
			    <%else%>
			    &nbsp;
			    <%end if%>
			
			</td>
			<td valign=top style="padding:5px 5px 5px 5px; width:150px;">
			Job: <a href="jobs.asp?menu=job&FM_kunde=<%=oRec("Kid") %>" class="vmenu"><%=Antal%></a>
			(<b><%=Aktive%></b> aktive)<br />
			<a href="jobs.asp?menu=job&func=opret&FM_kunde=<%=oRec("Kid") %>&step=2&int=1&rdir=kon" class="vmenu">Opret ny job >></a>
			
			<%call antalFakturaerKid(oRec("kid")) %>
			
			<br /><br />
			Fakturaer: 
			<%if level <= 2 OR levle = 6 then %>
			<a href="erp_fakhist.asp?FM_kunde=<%=oRec("kid") %>" class=vmenu target="_blank"><%=antalfak %></a> stk.
			<%else %>
			<b><%=antalfak %></b> stk.
			<%end if %>
			</td>
			
			
			
		<%end if%>
		
		    <%if oRec("Kid") = 1 then ' x > 0 OR %>
			<td>&nbsp;</td>
			<%else%>
			<td valign=top style="padding:5px 5px 5px 5px;">
			<a href="filer.asp?kundeid=<%=oRec("Kid")%>&jobid=0&nomenu=1" target="_blank" class=vmenu><img src="../ill/folder_document.png" border="0" alt="Filarkiv" /></a>
			&nbsp;
			
			
			<%
			
			
			if cint(antalfak) = 0 AND level = 1 then %>
			<a href="kunder.asp?menu=<%=memu %>&func=slet&id=<%=oRec("Kid")%>&FM_soeg=<%=thiskri%>&medarb=<%=medarb%>"><img src="../ill/slet_16.gif" alt="Slet" border="0"></a></td>
			<%else %>
			&nbsp;
			<%end if %>
			
			<%end if%>
		
		<%if menu <> "crm" then %>
		<td>&nbsp;</td>
		<%end if %>
		
		<td valign="top" style="padding:5px 5px 5px 5px;" align=right><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		</tr>
		<%
		x = 0
		y = 0
		showkunde = showkunde & "#" & oRec("kid") & "#"
		end if
	end function
	'***************************************************************************************************
	
	
	if len(medarb) = 0 then
	usemedarbKri = " AND aktionsid = crmhistorik.id AND medarbid = '" & session("Mid") & "' "
	else
		select case level 
		case "1", "2"
				if medarb = "0" then
				usemedarbKri = " AND editorid <> 0 "
				else
				usemedarbKri = " AND aktionsid = crmhistorik.id AND medarbid = '" & medarb & "' "
				end if
		
		end select
	end if
	
	
	'********* Slut CRM ***************
	useemneKri = " "
	usestatKri = " "
	end if
	
	
	if len(medarb) <> 0 then
	useMedarb = medarb
	else
	useMedarb = 0 'session("Mid")
	medarb = useMedarb
	end if
	
	
	if useMedarb = "0" then
		useopraf_editor = " "
	else
		useopraf_editor = " AND (kunder.kundeans1 = "& useMedarb &" OR kunder.kundeans2 = "& useMedarb &")"
	end if
	%>
	
	<!-------------------------------Sideindhold------------------------------------->
	
	<% 
	oskrift = "Kontakter"
	oskrift2 = "Opret ny kontakt"
	
	%>
	<div id="sindhold" style="position:absolute; left:20; top:127;">
	<%
	
	
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	
	
	
	oimg = "ikon_kunder_48.png"
	oleft = 0
	otop = 0
	owdt = 300
	oskrift = "Kontakter"
	
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	
	
	
	call filterheader(0,0,710,pTxt)
	
	%>
	
		
			
			<table cellspacing="0" cellpadding="0" border="0" width=100%>
			<tr><form action="kunder.asp?menu=<%=menu%>&viskunder=1" method="POST">
			<td style="padding-bottom:3px; padding-left:10px;">
			Søg på Navn, Id eller Tel.<br />
			<input type="text" name="FM_soeg" id="FM_soeg" value="<%=thiskri%>" size="30" style="font-family: arial,helvetica,sans-serif; font-size: 10px;"></td>
			
			<td style="padding-bottom:3px; padding-left:10px;">Kundeans.<br /> <select name="medarb" size="1" style="font-family: arial,helvetica,sans-serif; font-size: 10px; width:180;">
				<option value="0">Alle</option>
				<%
						strSQL = "SELECT mnavn, mid FROM medarbejdere ORDER BY mnavn"
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
				</td><td>
				Type / Segment:<br /> <select name="ktype" size="1" style="font-family: arial,helvetica,sans-serif; font-size: 10px; width:170px;">
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
						%>
				</select></td>
			

            <%if menu = "crm" then%>
            
            
            <%
            if len(request("FM_hotnot")) <> 0 then
            hotnot = request("FM_hotnot")
            response.cookies("tsa")("hotnot") = hotnot
            response.cookies("tsa").expires = date + 5
            else
                
                if request.cookies("tsa")("hotnot") <> "" then
                hotnot = request.cookies("tsa")("hotnot") 
                else
                hotnot = 99
                end if
            
            end if 

            select case cint(hotnot) 
            case -2
            hotnotSEL1 = "SELECTED"
            hotnotSEL2 = ""
            hotnotSEL3 = ""
            hotnotSEL4 = ""
            hotnotSEL5 = ""
            hotnotSEL0 = ""
            case -1
            hotnotSEL2 = "SELECTED"
            hotnotSEL1 = ""
            hotnotSEL3 = ""
            hotnotSEL4 = ""
            hotnotSEL5 = ""
            hotnotSEL0 = ""
            case 0
            hotnotSEL3 = "SELECTED"
            hotnotSEL2 = ""
            hotnotSEL1 = ""
            hotnotSEL4 = ""
            hotnotSEL5 = ""
            hotnotSEL0 = ""
            case 1
            hotnotSEL4 = "SELECTED"
            hotnotSEL2 = ""
            hotnotSEL3 = ""
            hotnotSEL1 = ""
            hotnotSEL5 = ""
            hotnotSEL0 = ""
            case 2
            hotnotSEL5 = "SELECTED"
            hotnotSEL2 = ""
            hotnotSEL3 = ""
            hotnotSEL4 = ""
            hotnotSEL1 = ""
            hotnotSEL0 = ""
            case else
            hotnotSEL5 = ""
            hotnotSEL2 = ""
            hotnotSEL3 = ""
            hotnotSEL4 = ""
            hotnotSEL1 = ""
            hotnotSEL0 = "SELECTED"
            end select

            %>


            <td>Interesse:<br /><select name="FM_hotnot" style="font-family: arial,helvetica,sans-serif; font-size: 10px; width:60;">
		            <option value="99" <%=hotnotSEL0  %>>Alle</option>
		            <option value="-2" <%=hotnotSEL1  %>>-2 (Not)</option>
		            <option value="-1" <%=hotnotSEL2  %>>-1</option>
		            <option value="0" <%=hotnotSEL3  %>>0</option>
		            <option value="1" <%=hotnotSEL4  %>>+1</option>
		            <option value="2" <%=hotnotSEL5  %>>+2 (Hot)</option>
            				
		            </select>
            	
            	
            	
            </td>
            <%end if %>
            
			<td align=right><br />
                <input id="Submit1" type="submit" value="Vis kontakter >>" /></td>
			</tr></form>
			</table>
		<!-- filter div -->
		</td></tr></table>
	</div>
	
	<%
	
	oleftpx = 0
	otoppx = 20
	owdtpx = 140
	
	call opretNy("kunder.asp?menu="&menu&"&func=opret&ketype="&ketype&"&FM_soeg="&thiskri&"&medarb="&medarb&"", "Opret ny kontakt", otoppx, oleftpx, owdtpx) 
	%>
	
	<br /><br />
    
    
    <%
    
    tTop = 20
	tLeft = 0
	tWdth = 904
	
	
	call tableDiv(tTop,tLeft,tWdth)
	%>
	
	
	<table cellspacing="0" cellpadding="0" border="0" width="100%" bgcolor="#EFF3FF">
 	<tr bgcolor="#5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/blank.gif" width="8" height="30" alt="" border="0"></td>
		<td colspan=7 valign="top"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/blank.gif" width="8" height="30" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
	<td height="30" class='alt'><b>Kontakt id</b></td>
	<%if menu = "crm" then%>
	<td class='alt'><b>Kontakt</b></td>
	<td class='alt'><b>Interesse niveau</b></td>
	<td class='alt'><b>Seneste 5 aktioner</b></td>
	<td class='alt' width="100">Ny aktion</td>
	<td align=right>&nbsp;</td>
	<td align=right>&nbsp;</td>
	<%else%>
	<td class='alt' style="width:400px;"><b>Navn</b></td>
	<td class='alt'><b>Type/Segment</b></td>
	<td align=right>&nbsp;</td>
	<td class='alt'><b>Jobs (Aktive)</b><br />
	Antal fakturaer</td>
	<td class=alt colspan=2>Filarkiv / Slet<br />
	(hvis der ikke findes fakturaer)</td>
	
	<%end if%>
	</tr>
	<%
	
	sort = Request("sort")
	
	
		
		'** Søge kriterier **
		if visikkekunder <> 1 then
		
			if usekri = 1 then '**SøgeKri
			sqlsearchKri = " (Kkundenavn LIKE '"& thiskri &"%' OR (Kkundenr LIKE '"& thiskri &"%' AND kkundenr <> '0') OR telefon = '"& thiskri &"')" 
			else
			sqlsearchKri = " (kid <> 0)"
			end if
			
		else
		sqlsearchKri = " (kid = 0)"
		end if
		
		
		if ktype > 0 then
		typeKri = " AND ktype = "&ktype&" "
		else
		typeKri = " AND ktype > -2 "
		end if
		
		
		if cint(hotnot) = 99 then
	    hotnotKri = " AND hot <> 99 "
	    else
	    hotnotKri = " AND hot = "& hotnot
	    end if
		
		kansSQLKri = useopraf_editor
	
	if menu = "crm" then
		fontcolor = "fir_g"
		strSQL = "SELECT kid, kkundenavn, kkundenr, useasfak, hot, "_
		&" kunder.editor, ktype, kunder.kundeans1, kunder.kundeans2, adresse, postnr, city, land, telefon, "_
		&" m1.mnavn AS kans1, m1.init AS kans1_init, m2.mnavn AS kans2, m2.init AS kans2_init, m1.mnr AS mnr1, m2.mnr AS mnr2, cvr, ean "_
		&" FROM kunder "_
		&" LEFT JOIN medarbejdere m1 ON m1.mid = kundeans1"_
		&" LEFT JOIN medarbejdere m2 ON m2.mid = kundeans2"_
		&" WHERE "& sqlsearchKri &" "& kansSQLKri &" AND kunder.ketype <> 'k' "& typeKri &""& hotnotKri &""_
		&" ORDER BY Kkundenavn"
		
	
	else
		strSQL = "SELECT Kid, Kkundenavn, Kkundenr, useasfak, ktype, kunder.kundeans1, kunder.kundeans2,"_
		&" m1.mnavn AS kans1, m1.mnr AS mnr1, m1.init AS kans1_init, m2.mnavn AS kans2, m2.mnr AS mnr2, "_
		&" m2.init AS kans2_init, adresse, postnr, city, land, telefon, cvr, ean, ktype, kt.navn AS ktnavn FROM kunder "_
		&" LEFT JOIN medarbejdere m1 ON m1.mid = kundeans1"_
		&" LEFT JOIN medarbejdere m2 ON m2.mid = kundeans2"_
		&" LEFT JOIN kundetyper AS kt ON (kt.id = kunder.ktype)"_
		&" WHERE "& sqlsearchKri &" "& kansSQLKri &" AND ketype <> 'e' "& typeKri &" ORDER BY Kkundenavn"
	end if

    	
	'Response.write strSQL
	'Response.Flush
	
	kids = "0"


	
	'** Udskriv query til skærm **
	'Response.write strSQL
	'Response.flush
	'*****************************
	cnt = 0
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	
	select case right(cnt, 1)
	case 1,3,5,7,9
	bgthis = "#ffffff"
	case else
	bgthis = "#eff3ff"
	end select
		
		'** Til Export fil ***
		kids = kids & "," & oRec("kid")
		
		
		'*** Hvis ikke crm vises job og aktive job ***
		if menu <> "crm" then
			x = 0
			y = 0
			
			strSQL2 = "SELECT id, jobstatus FROM job WHERE jobknr = "& oRec("Kid")
			oRec2.open strSQL2, oConn, 3
			while not oRec2.EOF 
			
			x = x + 1
			if oRec2("jobstatus") = 1 then
			y = y + 1
			end if
			
			oRec2.movenext
			wend
			oRec2.close
			Antal = x
			Aktive = y
		end if
		'** slut *** 
		
		'**** Henter firmaer med aktioner der er nyere end dags dato ***	
		call hentfirmaer()
		
		cnt = cnt + 1
		
		oRec.movenext
		wend
		oRec.close
		if cnt = 0 then
		%>
		<tr>
			
			<td colspan="9" height=50 style="padding-left:20;">&nbsp;Der er ingen kontakter der matcher de valgte søgekritetier.</td>
		</tr>
		<%
		end if
		%>
		<tr bgcolor="#5582D2">
			<td valign="top"><img src="../ill/blank.gif" width="8" height="20" alt="" border="0"></td>
			<td colspan=7 class=alt>Ialt <b><%=cnt%></b> kontakter på listen</td>
			<td valign="top" align="right"><img src="../ill/blank.gif" width="8" height="10" alt="" border="0"></td>
		</tr>
		</table>
		
		<!-- table div slut -->
		</div>
        
     
<%  



ptop = 33
pleft = 725
pwdt = 180

call eksportogprint(ptop,pleft, pwdt)
%>


<tr>
    <td align=center>
    <a href="kunder_eksport.asp?kids=<%=kids%>" target="_blank" class=vmenu>
    <img src="../ill/export1.png" border=0></a>
    </td>
    </td><td><a href="kunder_eksport.asp?kids=<%=kids%>" target="_blank" class=rmenu>A) Eksporter viste kontakter og kontaktpersoner som .csv fil</a></td>
</tr>

<tr>
    <td align=center>
    <a href="kunder_eksport.asp?kpers=-1&kids=<%=kids%>" target="_blank" class=vmenu>
    <img src="../ill/export1.png" border=0></a>
    </td>
    </td><td><a href="kunder_eksport.asp?kpers=-1&kids=<%=kids%>" target="_blank" class=rmenu>B) Eksporter viste kontakter som .csv fil</a></td>
</tr>
    
   
	
   </table>
</div>


        
        
		
		
		
	<br><br>
	<br>
	
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br><br>&nbsp;</div>
	<%end select%>


<%end if
%>
<!--#include file="../inc/regular/footer_inc.asp"-->
