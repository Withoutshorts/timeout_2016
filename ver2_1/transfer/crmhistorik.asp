<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	kl = request("kl")
	crmaktion = request("crmaktion")
	thisfile = "crmhistorik"
	
	select case func
	case "slet"
	
	'*** Her spørges om det er ok at der slettes en aktion ***
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:190; top:120; visibility:visible;">
	<table cellspacing="2" cellpadding="2" border="0">
	<tr>
	    <td><img src="../ill/alert.gif" width="44" height="45" alt="" border="0">&nbsp;Du er ved at <b>slette</b> en aktion. Er dette korrekt?</td>
	</tr>
	<tr>
	   <td><img src="../ill/blank.gif" width="44" height="1" alt="" border="0">&nbsp;<a href="crmhistorik.asp?menu=crm&func=sletok&id=<%=id%>&crmaktion=<%=crmaktion%>&ketype=e&emner=<%=request("emner")%>&status=<%=request("status")%>&medarb=<%=request("medarb")%>&selpkt=<%=request("selpkt")%>">Ja</a>&nbsp;&nbsp;&nbsp;<a href="javascript:history.back()">Nej</a></td>
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
	'*** Her slettes en crmaktion ***
	oConn.execute("DELETE FROM crmhistorik WHERE id = "& crmaktion &"")
	
	if request("selpkt") <> "kal" then
	Response.redirect "crmhistorik.asp?menu=crm&shokselector=1&ketype=e&id="&id&"&func=hist&selpkt=hist&emner="&request("emner")&"&status="&request("status")&"&medarb="&request("medarb")&""
	else
	Response.redirect "crmkalender.asp?menu=crm&shokselector=1&id="&id&"&medarb="&request("medarb")&""
	end if
	
	'**********************************************************************************************
	'** Her indsættes en ny aktion i db 													   ****
	'**********************************************************************************************
	case "dbopr", "dbred"
	personal = request("personal") 
	opretopf = request("FM_opf")
	
	'**********************************************************************
	'*** Validering starter 											***
	'**********************************************************************
	
	
	'** Validering af start og slut dato. ***
	if personal <> "y" then
	valstrCrmdato = request("FM_start_dag")&"/"&request("FM_start_mrd")&"/"&request("FM_start_aar")
	else
	valstrCrmdato = "1/1/2001"
	end if
	
	if personal <> "y" then
	valstrCrmdato_2 = request("FM_start_dag_2")&"/"&request("FM_start_mrd_2")&"/"&request("FM_start_aar_2")
	else
	valstrCrmdato_2 = "12/12/2014"
	end if 
	
	'***** Validering *******
	if cint(opretopf) = 1 AND cdate(valstrCrmdato) > cdate(valstrCrmdato_2) then%>	
		<!--#include file="../inc/regular/header_inc.asp"-->
		<!--#include file="../inc/regular/topmenu_inc.asp"-->
		<%
		errortype = 37
		call showError(errortype)	
				
	else
		'****************************************
		'**** sætter CRM-klokkeslet på aktion ***
		'****************************************
		if len(request("klok_timer")) = 0 then
		strKloktimer = "12"
		else
		strKloktimer = request("klok_timer")
		end if
		
		if len(request("klok_min")) = 0 then
		strKlokmin = "00"
		else
		strKlokmin = request("klok_min")
		end if
		
		if len(request("klok_timer_slut")) = 0 then
		strKloktimer_slut = "12"
		else
		strKloktimer_slut = request("klok_timer_slut")
		end if
		
		if len(request("klok_min_slut")) = 0 then
		strKlokmin_slut = "00"
		else
		strKlokmin_slut = request("klok_min_slut")
		end if
		
		'**** Sætter CRM-klokkeslet på aktion *****
		strKlokkeslet = strKloktimer &":"& strKlokmin &":"&"01"
		strKlokkeslet_slut = strKloktimer_slut &":"& strKlokmin_slut &":"&"01"
		
		
		'**** Sætter CRM-klokkeslet på opfølgende aktion *****
		strKlokkeslet_2 = request("klok_timer_2") &":"& request("klok_min_2") &":"&"01"
		strKlokkeslet_slut_2 = request("klok_timer_slut_2") &":"& request("klok_min_slut_2") &":"&"01"
		' *** '
		
		
				'***** Validering *******
				if cint(opretopf) = 1 AND (cdate(valstrCrmdato) = cdate(valstrCrmdato_2)) AND (strKlokkeslet > strKlokkeslet_2 OR strKlokkeslet_slut > strKlokkeslet_2 OR strKlokkeslet_slut > strKlokkeslet_slut_2) then%>	
				<!--#include file="../inc/regular/header_inc.asp"-->
				<!--#include file="../inc/regular/topmenu_inc.asp"-->
				<%
				errortype = 38
				call showError(errortype)	
							
				else
	
					'***** Validering *******
					if (cint(opretopf) = 2) AND (cdate(valstrCrmdato) >= cdate(valstrCrmdato_2)) then%>
						<!--#include file="../inc/regular/header_inc.asp"-->
						<!--#include file="../inc/regular/topmenu_inc.asp"-->
						<%
						errortype = 36
						call showError(errortype)
					else
		
					
					'*** Validering ***
					if (cint(opretopf) = 3 AND (request("FM_dageldato") = "dato" AND len(trim(request("FM_gentag_dato"))) = 0)) OR (cint(opretopf) = 3 AND (request("FM_dageldato") = "dag" AND len(trim(request("FM_gentag_ugedag"))) = 0)) OR (cint(opretopf) = 3 AND len(trim(request("FM_gentag_md"))) = 0) then
						%>
						<!--#include file="../inc/regular/header_inc.asp"-->
						<%
						errortype = 41
						call showError(errortype)
					else
						
		'*****************************
		'* Validering godkendt    ****
		'* Starter oprettelse 	 *****
		'*****************************
		%>
		<!--#include file="inc/convertDate.asp"-->
		<%
		
		
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function
		
		
		'*****************************
		'** Henter medarbejder id ****
		'*****************************
		strSQL = "SELECT mid FROM medarbejdere WHERE mnavn = '"& session("user") &"'"
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then
		intMid = oRec("mid")
		End if
		oRec.close
		''**''
		
		'****************************************************************************
		'*** Sætter værdier alt efter om det er en personlig note eller aktion. 	*
		'****************************************************************************
		if personal <> "y" then
			strNavn = SQLBless(request("FM_navn"))
			strEmne = request("FM_emne")
			intKundeId = request("id")
			strBesk = SQLBless(request("FM_besk"))
			strEditor = session("user")
			strDato = year(now)&"/"&month(now)&"/"&day(now)
			strCrmdato = request("FM_start_aar")&"/"&request("FM_start_mrd")&"/"&request("FM_start_dag")
			strStatus = request("FM_status")
			strKontaktform = request("FM_kform")
			strKpers = request("FM_kpers")
			'strKlokkeslet sættes under validering
			'strKlokkeslet_slut sættes under validering
			if opretopf = 2 then
			strCrmdato_slut = request("FM_start_aar_2")&"/"&request("FM_start_mrd_2")&"/"&request("FM_start_dag_2")
			else
			strCrmdato_slut = strCrmdato
			end if
			
			'** Hvis slut klokkeslet er før starttid klokkeslet
			if strKlokkeslet > strKlokkeslet_slut then
			strKlokkeslet_slut = strKlokkeslet 
			end if
		
		else '*** personal note ****
			strNavn = ""
			strEmne = 0
			intKundeId = -1
			strBesk = SQLBless(request("FM_note_personal"))
			strEditor = session("user")
			strDato = year(now)&"/"&month(now)&"/"&day(now)
			strCrmdato = request("FM_start_aar_per")&"/"&request("FM_start_mrd_per")&"/"&request("FM_start_dag_per")
			strStatus = 0
			strKontaktform = 0
			strKpers = ""
			strKlokkeslet = request("FM_klslet_personal")&":"&"01"
			strKlokkeslet_slut = dateadd("h", 1, strKlokkeslet)
			strCrmdato_slut = strCrmdato
		end if
		' *** '
		
		'***************************************************************
		'** Sætter værdier på Getagelser, hvis dette punkt er valgt. ***
		'***************************************************************
		if cint(opretopf) = 3 then
		
		'*** finder næste serienummer ***
		strSQL = "SELECT serialnb FROM crmhistorik ORDER BY serialnb DESC"
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then 
		nextSerialnb = oRec("serialnb") + 1
		else
		nextSerialnb = 0
		end if
		oRec.close
		
		Call updinsertAktion(strCrmdato, strCrmdato_slut, nextSerialnb)
		
		g = 0
		i = 0
		a = 0
		num = 0
		dim useThis_g
		Redim useThis_g(i)
		dim intDatoer
		Redim intDatoer(g)
		dageldato = request("FM_dageldato")
		slutgentag = request("FM_gentag_dag")&"/"&request("FM_gentag_mrd")&"/"&request("FM_gentag_aar")
		
		if len(request("FM_gentag_ugedag")) <> 0 then
		g_ugedag = request("FM_gentag_ugedag")
		else
		g_ugedag = "1"
		end if
		
		if len(request("FM_gentag_dato")) <> 0 then
		g_dato = request("FM_gentag_dato")
		else
		g_dato = "1"
		end if
		
		if len(request("FM_gentag_md")) <> 0 then
		g_md = request("FM_gentag_md")
		else
		g_md = "1"
		end if
		
		diff = (datediff("yyyy", strCrmdato, slutgentag) + 1)
		weekdiff = datediff("ww", strCrmdato, slutgentag)
				
			if dageldato = "dato" then
					
				For a = 0 to diff - 1
					
					if a = 0 then
					ya = year(strCrmdato)
					else
					ya = ya + 1
					end if
					
					intDatoer = Split(g_dato, ", ")
					For g = 0 to Ubound(intDatoer)
						
						intMD = Split(g_md, ", ")
						For i = 0 to Ubound(intMD)
							Redim preserve useThis_g(i)
							useThis_g(i) = intMD(i)&"/"&intDatoer(g)&"/"&ya 
							
							if (cdate(useThis_g(i)) <= cdate(slutgentag)) AND cdate(useThis_g(i)) >= cdate(strCrmdato) then
							Call updinsertAktion(convertDateYMD(useThis_g(i)), convertDateYMD(useThis_g(i)), nextSerialnb)
							end if
						Next
					Next
				Next
				
			else 'Uge dage er valgt
				
				crmDatoweekday = datepart("w", strCrmdato)
				crmDatoday = day(strCrmdato)
				'*** Tjekker hvor mange uger der er imellem crmdato og slutdato ***
				For a = 0 to weekdiff - 1
					
					intDatoer = Split(g_ugedag, ", ")
				 	For g = 0 to Ubound(intDatoer)
					'** Mangler at tilføje datoer for alle de valgte dage.
					
					
					'** tjekker om det er første gennemløb og finder dato på de dage der er valgt. **
					if a = 0 then
						dayadd = crmDatoday + (intDatoer(g) - crmDatoweekday)  
						usethisDato = dayadd &"/"& month(strCrmdato) &"/"& year(strCrmdato)
						if cint(intDatoer(g)) > cint(crmDatoweekday) then
						usethisDato = usethisDato
						else
						usethisDato = dateadd("ww", 1, usethisDato)
						end if
					else
						usethisDato = dateadd("ww", 1, usethisDato)
					end if 
						
						'** tjekker om der er valgt måneder **
						intMD = Split(g_md, ", ")
						For i = 0 to Ubound(intMD)
							if cint(month(usethisDato)) = cint(intMD(i)) then 
								if (cdate(usethisDato) <= cdate(slutgentag)) AND cdate(usethisDato) >= cdate(strCrmdato) then
									Response.write usethisDato &"<br>"
									Redim preserve useThis_g(i)
									useThis_g(i) = usethisDato
									'Call updinsertAktion(convertDateYMD(useThis_g(i)), convertDateYMD(useThis_g(i)), nextSerialnb)
								end if
							end if
						Next
					Next
				Next
			
			
			end if
			
			
		
		else
		intserialnb = 0
		Call updinsertAktion(strCrmdato, strCrmdato_slut, intserialnb)
		end if
		' *** '
	
		
		
		'**********************************************************
		'**** Opretter / Redigerer aktion(er) 				*******
		'**********************************************************
		Public aktionsid
		function updinsertAktion(usefuncdato, usefuncslutdato, useserial)
		
		thisfuncdato = usefuncdato
		thisfuncslutdato = usefuncslutdato
		thisserial = useserial
		if func = "dbopr" then
		oConn.execute("INSERT INTO crmhistorik (editor, dato, crmdato, kontaktemne, kontaktpers, status, komm, navn, kundeid, kontaktform, editorid, crmklokkeslet, crmKlokkeslet_slut, crmdato_slut, serialnb) VALUES ('"& strEditor &"', '"& strDato &"', '"& thisfuncdato &"', "& strEmne &", '"& strKpers &"', '"& strStatus &"', '"& strBesk &"', '"& strNavn &"', "& intKundeId &", "& strKontaktform &", "& intMid &", '"& strKlokkeslet &"', '"& strKlokkeslet_slut &"', '"& thisfuncslutdato &"', "& thisserial &")")
			
			strSQL = "SELECT id FROM crmhistorik ORDER BY id DESC"
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
				aktionsid = oRec("id")
			end if
			oRec.close
		else
			useSQL = "UPDATE crmhistorik SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"', crmdato = '"& thisfuncdato &"', kontaktemne = "& strEmne &", kontaktpers = '"& strKpers &"', status = '"& strStatus &"', komm = '"& strBesk &"', navn = '"& strNavn &"', kundeid = "& intKundeId &", kontaktform = "& strKontaktform &", editorid = "& intMid &", crmklokkeslet = '"& strKlokkeslet &"', crmklokkeslet_slut = '"& strKlokkeslet_slut &"', crmdato_slut = '"& thisfuncslutdato &"' WHERE id = "& crmaktion &""
			
			oConn.execute(useSQL)
			aktionsid = crmaktion
		end if
		
		call relationer()
		
		end function
		' *** '
		
		
		
		'*******************************************************
		'**** Opdaterer aktions relationer til medarbejdere ****
		'*******************************************************
		function relationer()
		oConn.execute("DELETE FROM aktionsrelationer WHERE aktionsid = "& aktionsid &"")
			
		if len(request("FM_medarbrel")) <> 0 then
			Dim medarbRelationer
			Dim b
			b = 0
			medarbRelationer = Split(request("FM_medarbrel"), ", ")
				For b = 0 to Ubound(medarbRelationer)
				oConn.execute("INSERT INTO aktionsrelationer (aktionsid, medarbid) VALUES ("& aktionsid &", "& medarbRelationer(b) &") ")
				next
		else
		oConn.execute("INSERT INTO aktionsrelationer (aktionsid, medarbid) VALUES ("& aktionsid &", "& intMid &") ")
		end if
		end function
		' *** '
		
		
		'*************************************************
		'*** Opretter opfølgning on then fly ... 		 *
		'*************************************************
		if cint(opretopf) = 1 AND func = "dbopr" then 
			
			strNavn_2 = SQLBless(request("FM_navn_2"))
			strEmne_2 = request("FM_emne")
			intKundeId_2 = request("id")
			strBesk_2 = SQLBless(request("FM_besk_2"))
			strEditor_2 = session("user")
			strDato_2 = year(now)&"/"&month(now)&"/"&day(now)
			strCrmdato_2 = request("FM_start_aar_2")&"/"&request("FM_start_mrd_2")&"/"&request("FM_start_dag_2")
			strStatus_2 = request("FM_status_2")
			strKontaktform_2 = request("FM_kform_2")
			strKpers_2 = request("FM_kpers_2")
			strCrmdato_slut = strCrmdato_2
			'strKlokkeslet_2 sættes under validering
			'strKlokkeslet_slut_2 sættes under validering
			
			'** Hvis slut klokkeslet er før starttid
			if strKlokkeslet_2 > strKlokkeslet_slut_2 then
			strKlokkeslet_slut_2 = strKlokkeslet_2 
			end if
			
			oConn.execute("INSERT INTO crmhistorik (editor, dato, crmdato, kontaktemne, kontaktpers, status, komm, navn, kundeid, kontaktform, editorid, crmklokkeslet, crmKlokkeslet_slut, crmdato_slut, serialnb) VALUES ('"& strEditor_2 &"', '"& strDato_2 &"', '"& strCrmdato_2 &"', "& strEmne_2 &", '"& strKpers_2 &"', '"& strStatus_2 &"', '"& strBesk_2 &"', '"& strNavn_2 &"', "& intKundeId_2 &", "& strKontaktform_2 &", "& intMid &", '"& strKlokkeslet_2 &"', '"& strKlokkeslet_slut_2 &"', '"& strCrmdato_slut &"', 0)")
			
			strSQL = "SELECT id FROM crmhistorik ORDER BY id DESC"
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
				aktionsid = oRec("id")
			end if
			oRec.close
		
		call relationer()
		
		end if
		' *** '
		
		
		
		
		'********************************************************************
		'** Oprettelse /redigering slut. Viser ok side til bruge eller		*
		'** Retunerer til crmkalender hvis det eren personlig note. 		*
		'********************************************************************
		
			if personal <> "y" then
			%>
			<html>
			<head>
			<title>timeOut 2.1</title>
			<LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print.css">
			</head>
			<body topmargin="0" leftmargin="10" class="regular">
			<table cellspacing="0" cellpadding="0" border="0">
			<tr>
		    <td valign="top"><br>
			<img src="../ill/header_crmosigt.gif" alt="" border="0"><hr align="left" width="360" size="1" color="#000000" noshade>
			</td></tr></table>
			<br>
			<table><tr><td>
			<img src="../ill/info.gif" width="42" height="38" alt="" border="0"><b>Aktionen er oprettet!</b><br><br>
			<%if cint(opretopf) = 3 then
				Response.write "<br>Følgende andre aktioner er indsat i denne serie:<br>"
				'For i = 0 to Ubound(intMD)
				'Response.write useThis_g(i) &"<br>"
				'next
			Response.write "<br><br>"	
			end if
			%>
			
			<a onClick="alt()" href="Javascript:window.close()">Luk dette vindue</a><br><br>
			NB) Husk at opdatere [F5] den side du kom fra. <br>
			så du kan se den nye aktion.
			</td></tr></table> 
			<%
			else
				if len(request("kalviskunde")) <> 0 then
				kalviskunde = request("kalviskunde")
				else
				kalviskunde = 0
				end if
			Response.redirect "crmkalender.asp?menu=crm&shokselector=1&strdag="&request("FM_start_dag_per")&"&strmrd="&request("FM_start_mrd_per")&"&straar="&request("FM_start_aar_per")&"&medarb="&request("kalmedarb")&"&id="&kalviskunde&"&ketype=e"
			end if
		' *** '
		
		end if '*validering
		end if '*validering
	end if '*validering
	end if '*validering
	
	case "opret", "red"
	
	'************************************************************************************************
	'*** Her indlæses form til rediger/oprettelse af ny aktion.										*
	'************************************************************************************************
	'** local kr/dato/nummer format
	session.LCID = 1030
	
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	trearsize = "40"
	'varSubVal = "opretpil" 
	'varbroedkrumme = "Opret ny aktion"
		
		'** Hvis der først skal vælges kunde ***
		if func = "opret" AND id = 0 then
		dbfunc = "opret"
		varSubVal = "step2pil" 
		varbroedkrumme = "Opret ny aktion step 1/2"
		else
		dbfunc = "dbopr"
		varSubVal = "opretpil"
		varbroedkrumme = "Opret ny aktion"
		end if
		'*** ****
		
	else
	
	strSQL = "SELECT editor, dato, crmdato, kontaktemne, kontaktpers, status, komm, navn, kundeid, kontaktform, crmklokkeslet, crmklokkeslet_slut, crmDato_slut FROM crmhistorik WHERE id=" & crmaktion 
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	crmDato = oRec("crmDato")
	StrTdato = crmDato
	intEmne = oRec("kontaktemne")
	intStatus = oRec("status")
	strBesk = oRec("komm")
	intKontaktform = oRec("kontaktform")
	strKlokkeslet = oRec("crmklokkeslet")
	strKlokkeslet_slut = oRec("crmklokkeslet_slut")
	crmDato_slut = oRec("crmDato_slut")
	end if
	oRec.close
	
	
	trearsize = "40"
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "opdaterpil" 
	end if
	
	if request("showinwin") <> "j" then%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<%else%>
	<html>
		<head>
		<title>timeOut 2.1</title>
		<LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print.css">
		<script LANGUAGE="javascript">
		function NewWin_popupcal(url)    {
		window.open(url, 'Calpick', 'width=250,height=250,scrollbars=no,toolbar=no');    
		}
		</script>
		</head>
		<body topmargin="0" leftmargin="0" class="regular">
	<%end if%>
	
	<script LANGUAGE="javascript">
	function insvaloskrift(){
	document.all["FM_navn_2"].value = document.all["FM_navn"].value;
	document.all["FM_navn_2"].focus()
	}
	function insvalkpers(){
	document.all["FM_kpers_2"].value = document.all["FM_kpers"].value;
	document.all["FM_kpers_2"].focus()
	}
	function insvalstat(){
	document.all["FM_status_2"].value = document.all["FM_status"].value;
	document.all["FM_status_2"].focus()
	}
	function insvalkform(){
	document.all["FM_kform_2"].value = document.all["FM_kform"].value;
	document.all["FM_kform_2"].focus()
	}
	
	function showakt1(){
	document.all["dato2"].style.visibility = "hidden";
	
	
	if (document.all["FM_opf2"].checked == true) {
	document.all["sd"].style.visibility = "hidden";
	document.all["div1B"].style.display = "none";
	document.all["div1B"].style.visibility = "hidden";
	}
	
	document.all["divak1"].style.top = "214";
	document.all["divak2"].style.top = "213";
	document.all["divak3"].style.top = "213";
	
	//document.all["divak1"].style.border = "1px red solid"
	
	document.all["div0A"].style.display = "";
	document.all["div0A"].style.visibility = "visible";
	document.all["div1A"].style.display = "";
	document.all["div1A"].style.visibility = "visible";
	document.all["div2A"].style.display = "";
	document.all["div2A"].style.visibility = "visible";
	
	document.all["div0B"].style.display = "none";
	document.all["div0B"].style.visibility = "hidden";
	document.all["div2B"].style.display = "none";
	document.all["div2B"].style.visibility = "hidden";
	
	document.all["div1C"].style.display = "none";
	document.all["div1C"].style.visibility = "hidden";
	}
	
	function showakt2(){
	document.all["dato2"].style.visibility = "visible";
	document.all["sd"].style.visibility = "hidden";
	
	document.all["divak1"].style.top = "213";
	document.all["divak2"].style.top = "214";
	
	document.all["divak2"].style.visibility = "visible";
	document.all["divak1"].style.top = "213";
	document.all["divak2"].style.top = "214";
	
	document.all["div0A"].style.display = "none";
	document.all["div0A"].style.visibility = "hidden";
	document.all["div1A"].style.display = "none";
	document.all["div1A"].style.visibility = "hidden";
	document.all["div2A"].style.display = "none";
	document.all["div2A"].style.visibility = "hidden";
	
	document.all["div0B"].style.display = "";
	document.all["div0B"].style.visibility = "visible";
	document.all["div1B"].style.display = "";
	document.all["div1B"].style.visibility = "visible";
	document.all["div2B"].style.display = "";
	document.all["div2B"].style.visibility = "visible";
	
	document.all["divak3"].style.visibility = "hidden";
	document.all["div1C"].style.display = "none";
	document.all["div1C"].style.visibility = "hidden";
	}
	
	function showopf(){
	document.all["dato2"].style.visibility = "visible";
	document.all["sd"].style.visibility = "hidden";
	
	document.all["divak2"].style.visibility = "visible";
	document.all["divak1"].style.top = "213";
	document.all["divak2"].style.top = "214";
	
	document.all["div0A"].style.display = "none";
	document.all["div0A"].style.visibility = "hidden";
	document.all["div1A"].style.display = "none";
	document.all["div1A"].style.visibility = "hidden";
	document.all["div2A"].style.display = "none";
	document.all["div2A"].style.visibility = "hidden";
	
	document.all["div0B"].style.display = "";
	document.all["div0B"].style.visibility = "visible";
	document.all["div1B"].style.display = "";
	document.all["div1B"].style.visibility = "visible";
	document.all["div2B"].style.display = "";
	document.all["div2B"].style.visibility = "visible";
	
	document.all["divak3"].style.visibility = "hidden";
	document.all["div1C"].style.display = "none";
	document.all["div1C"].style.visibility = "hidden";
	}
	
	function hideopf(){
	document.all["dato2"].style.visibility = "hidden";
	document.all["sd"].style.visibility = "hidden";
	
	document.all["divak2"].style.visibility = "hidden";
	document.all["divak1"].style.top = "214";
	
	document.all["div0A"].style.display = "";
	document.all["div0A"].style.visibility = "visible";
	document.all["div1A"].style.display = "";
	document.all["div1A"].style.visibility = "visible";
	document.all["div2A"].style.display = "";
	document.all["div2A"].style.visibility = "visible";
	
	document.all["div0B"].style.display = "none";
	document.all["div0B"].style.visibility = "hidden";
	document.all["div1B"].style.display = "none";
	document.all["div1B"].style.visibility = "hidden";
	document.all["div2B"].style.display = "none";
	document.all["div2B"].style.visibility = "hidden";
	
	document.all["divak3"].style.visibility = "hidden";
	document.all["div1C"].style.display = "none";
	document.all["div1C"].style.visibility = "hidden";
	}
	
	
	function showendd(){
	document.all["dato2"].style.visibility = "hidden";
	document.all["sd"].style.visibility = "visible";
	
	document.all["divak2"].style.visibility = "hidden";
	document.all["divak1"].style.top = "214";
	
	
	document.all["div0A"].style.display = "";
	document.all["div0A"].style.visibility = "visible";
	document.all["div1A"].style.display = "";
	document.all["div1A"].style.visibility = "visible";
	document.all["div2A"].style.display = "";
	document.all["div2A"].style.visibility = "visible";
	
	document.all["div0B"].style.display = "none";
	document.all["div0B"].style.visibility = "hidden";
	document.all["div1B"].style.display = "";
	document.all["div1B"].style.visibility = "visible";
	document.all["div2B"].style.display = "none";
	document.all["div2B"].style.visibility = "hidden";
	
	document.all["divak3"].style.visibility = "hidden";
	document.all["div1C"].style.display = "none";
	document.all["div1C"].style.visibility = "hidden";
	}
	
	function showreocc(){
	document.all["dato2"].style.visibility = "hidden";
	document.all["sd"].style.visibility = "hidden";
	
	document.all["divak2"].style.visibility = "hidden";
	document.all["divak3"].style.visibility = "visible";
	document.all["divak1"].style.top = "213";
	document.all["divak3"].style.top = "214";
	
	document.all["div0A"].style.display = "none";
	document.all["div0A"].style.visibility = "hidden";
	document.all["div1A"].style.display = "none";
	document.all["div1A"].style.visibility = "hidden";
	document.all["div2A"].style.display = "none";
	document.all["div2A"].style.visibility = "hidden";
	
	document.all["div0B"].style.display = "none";
	document.all["div0B"].style.visibility = "hidden";
	document.all["div1B"].style.display = "none";
	document.all["div1B"].style.visibility = "hidden";
	document.all["div2B"].style.display = "none";
	document.all["div2B"].style.visibility = "hidden";
	
	document.all["div1C"].style.display = "";
	document.all["div1C"].style.visibility = "visible";
	}
	
	function NewWin_popupcal(url)    {
		window.open(url, 'Calpick', 'width=250,height=250,scrollbars=no,toolbar=no');    
		}
	</script>
	<%
	showinwin = request("showinwin")
	
	if showinwin <> "j" then%>
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!--#include file="../inc/regular/rmenu.asp"-->
	<%end if
	
	if func = "red" then
	id = id
	'Response.write strKlokkeslet
	
	formatstrKlokkeslet = FormatDateTime(strKlokkeslet,4)
	klok_timer = left(formatstrKlokkeslet,2)
	klok_min = right(formatstrKlokkeslet,2)
	
	formatstrKlokkeslet_slut = FormatDateTime(strKlokkeslet_slut,4)
	klok_timer_slut = left(formatstrKlokkeslet_slut,2)
	klok_min_slut = right(formatstrKlokkeslet_slut,2)
	%>
	<!--#include file="inc/dato2.asp"-->
	<% 
	else
	
	if len(request("strdag")) <> 0 then
	strDag = request("strdag")
	strMrd = request("strmrd")
	strAar = request("straar")
	else
	strDag = day(now)
	strMrd = month(now)
	strAar = year(now)
	end if
	%>
	<!--#include file="inc/dato2_b.asp"-->
	<%
	
	
	klok_timer = left(kl, 2)
	if len(klok_timer) <> 0 then
	klok_timer = klok_timer
	else
	klok_timer = left(FormatDateTime(time,4),2)
	end if
	klok_timer_slut = klok_timer
	end if
	
	if request("selpkt") = "kal" then
	selpkt = "kal"
	kalvalues = "&kaldato="&request("kaldato")&"&kalstatus="&request("kalstatus")&"&kalemner="&request("kalemner")&"&kalmedarb="&request("kalmedarb")&"&kalviskunde="&request("kalviskunde")&""
	else
	selpkt = "hist"
	kalvalues = " "
	end if
	
	if showinwin <> "j" then
	divleft = 190
	divtop = 50
	else
	divleft = 10
	divtop = 20
	end if
	%>
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:<%=divleft%>; top:<%=divtop%>; visibility:visible;">
	<table cellspacing="0" cellpadding="0" border="0">
	<tr>
    <td valign="top"><br>
	<img src="../ill/header_crmosigt.gif" alt="" border="0"><hr align="left" width="360" size="1" color="#000000" noshade>
	</td></tr></table>
	<br>
	<table cellspacing="0" cellpadding="0" border="0" width="360" bgcolor="#EFF3FF">
	<form action="crmhistorik.asp?menu=crm&func=<%=dbfunc%>&ketype=e&selpkt=<%=selpkt&kalvalues%>" method="post">
	<input type="hidden" name="crmaktion" value="<%=crmaktion%>">
	<input type="hidden" name="showinwin" value="<%=showinwin%>">
	<tr bgcolor="#5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td colspan=3 valign="top"><img src="../ill/tabel_top.gif" width="343" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td colspan="3" style="height:20;" class=alt>
		<b>CRM aktion -- &nbsp;<%=varbroedkrumme%></b></td>
	</tr>
	<%if dbfunc = "dbred" then%>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td colspan="3"><br>Sidst opdateret den <b><%=formatdatetime(strDato, 1)%></b> af <b><%=strEditor%></b></td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<%end if
	
	table_topgif = "343"
	
	if func = "opret" AND id = 0 then
	%>
	<tr>
	<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="70" alt="" border="0"></td>
	<td style="padding-top:5;" colspan="3"><br>Kunde:&nbsp;&nbsp;<select name="id">
		<%
		strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE ketype <> 'k' ORDER BY Kkundenavn"
		oRec.open strSQL, oConn, 3
		while not oRec.EOF
		%>
		<option value="<%=oRec("Kid")%>"><%=left(oRec("Kkundenavn"), 9)%></option>
		<%
		oRec.movenext
		wend
		oRec.close
		%>
	</select>
	<br><br>&nbsp;</td>
	<input type="hidden" name="kl" value="<%=kl%>">
	<input type="hidden" name="strdag" value="<%=strDag%>">
	<input type="hidden" name="strmrd" value="<%=strMrd%>">
	<input type="hidden" name="straar" value="<%=strAar%>">
	<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="70" alt="" border="0"></td>
	</tr>
	</table>
	<%
	else
	
			'** Finder kunde ****
			strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE kid = " & id
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
			 strKundenavn = oRec("Kkundenavn")
			end if
			oRec.close
			
			if func = "opret" then%>
			<tr>
			<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="90" alt="" border="0"></td>
			<td colspan="3" style="padding-top:5;">
			<b>A)</b> Opret enkelt aktion: (1 Aktion) <input type="radio" checked name="FM_opf" id="FM_opf1" value="0" onclick="hideopf();">
			<br><b>B) </b>Tilføj <b>Opfølgning</b> til denne aktion?: (2 Aktion) <input type="radio" name="FM_opf" id="FM_opf2" value="1" onclick="showopf();">
			<br><b>C)</b> 1 Aktion løber <b>over flere dage</b>:<input type="radio" name="FM_opf" id="FM_opf3" value="2" onclick="showendd();">
			<br><b>D)</b> Gentag 1 Aktion i et <b>bestemt interval</b>:<input type="radio" name="FM_opf" id="FM_opf4" value="3" onclick="showreocc();">
			</td>
			<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="90" alt="" border="0"></td>
			</tr>
			<%end if%>
			<input type="hidden" name="id" value="<%=id%>">
			<tr bgcolor="#EFF3FF">
				<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
				<td colspan=3 valign="bottom"><img src="../ill/tabel_top.gif" width="346" height="1" alt="" border="0"></td>
				<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
			</tr>
	</table>
	<br>
	
	<!-- dato /slut dato -->
	<div id=dato2 name=dato2 style="position:absolute; top:270; left:10; visibility:hidden; z-index:20000;">
	Dato:</div>
	<div id=sd name=sd style="position:absolute; top:310; left:10; width:100; height:19; visibility:hidden; z-index:20000;">
	Slut dato:</div>
	
	<%if func = "red" then
	vizi = "hidden" 
	else
	vizi = "visible"
	end if%>
	
	<!-- 1 og 2 aktion knapper -->
	<div id=divak1 name=divak1 style="position:absolute; top:214; left:15; visibility:<%=vizi%>; z-index:800;">
	<table cellspacing="0" cellpadding="0" border="0">
	<tr bgcolor="5582D2">
			<td width="8" rowspan="2" valign=top><img src="../ill/tt_tabel_top_left.gif" width="8" height="15" alt="" border="0"></td>
			<td valign="top"><img src="../ill/tabel_top.gif" width="64" height="1" alt="" border="0"></td>
			<td align=right rowspan="2" valign=top><img src="../ill/tt_tabel_top_right.gif" width="8" height="15" alt="" border="0"></td>
	</tr>
	<tr bgcolor="5582D2">
			<td><a href="#" onClick="showakt1()" class='alt'>1. Aktion</a></td>
	</tr>
	</table>
	</div>
	<div id=divak2 name=divak2 style="position:absolute; top:214; left:100; visibility:hidden; z-index:800;">
	<table cellspacing="0" cellpadding="0" border="0">
	<tr bgcolor="5582D2">
			<td width="8" rowspan="2" valign=top><img src="../ill/tt_tabel_top_left.gif" width="8" height="15" alt="" border="0"></td>
			<td valign="top"><img src="../ill/tabel_top.gif" width="64" height="1" alt="" border="0"></td>
			<td align=right rowspan="2" valign=top><img src="../ill/tt_tabel_top_right.gif" width="8" height="15" alt="" border="0"></td>
	</tr>
	<tr bgcolor="5582D2">
			<td><a href="#" onClick="showakt2()" class='alt'>2. Aktion</a></td>
	</tr>
	</table>
	</div>
	<div id=divak3 name=divak3 style="position:absolute; top:214; left:100; visibility:hidden; z-index:200;">
	<table cellspacing="0" cellpadding="0" border="0">
	<tr bgcolor="5582D2">
			<td width="8" rowspan="2" valign=top><img src="../ill/tt_tabel_top_left.gif" width="8" height="15" alt="" border="0"></td>
			<td valign="top"><img src="../ill/tabel_top.gif" width="84" height="1" alt="" border="0"></td>
			<td align=right rowspan="2" valign=top><img src="../ill/tt_tabel_top_right.gif" width="8" height="15" alt="" border="0"></td>
	</tr>
	<tr bgcolor="5582D2">
			<td><a href="#" onClick="showreocc()" class='alt'>Gentag aktion</a></td>
	</tr>
	</table>
	</div>
	<!--slut -->
	
	<br>
	<div id=div0A name=div0A style="position:relative;  visibility:visible; display:; z-index:20;">
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF" width="360">
	<tr bgcolor="#5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td colspan=3 valign="top"><img src="../ill/tabel_top.gif" width="<%=table_topgif%>" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td colspan="3" style="height:20; padding-left:10;" class=alt><b><%=strKundenavn%>&nbsp;aktion:</b>&nbsp;</td>
	</tr>
	</table>
	</div>
	
	
	<div id=div0B name=div0B style="position:relative; visibility:hidden; display:none; z-index:20;">
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF" width="360">
	<tr bgcolor="#5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td colspan=3 valign="top"><img src="../ill/tabel_top.gif" width="<%=table_topgif%>" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td colspan="3" style="height:20; padding-left:10;" class=alt><b><%=strKundenavn%>&nbsp;2. aktion: (opfølgende aktion)</b>&nbsp;</td>
	</tr>
	</table>
	</div>
	
	<div id=div1A name=div1A style="position:relative; visibility:visible; display:; z-index:20;">
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF" width="360">
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="40" alt="" border="0"></td>
		<td style="padding-left:10;"><br>Dato:&nbsp;</td>
		<td><br>
		<select name="FM_start_dag">
		<option value="<%=strDag%>"><%=strDag%></option> 
		<option value="1">1</option>
	   	<option value="2">2</option>
	   	<option value="3">3</option>
	   	<option value="4">4</option>
	   	<option value="5">5</option>
	   	<option value="6">6</option>
	   	<option value="7">7</option>
	   	<option value="8">8</option>
	   	<option value="9">9</option>
	   	<option value="10">10</option>
	   	<option value="11">11</option>
	   	<option value="12">12</option>
	   	<option value="13">13</option>
	   	<option value="14">14</option>
	   	<option value="15">15</option>
	   	<option value="16">16</option>
	   	<option value="17">17</option>
	   	<option value="18">18</option>
	   	<option value="19">19</option>
	   	<option value="20">20</option>
	   	<option value="21">21</option>
	   	<option value="22">22</option>
	   	<option value="23">23</option>
	   	<option value="24">24</option>
	   	<option value="25">25</option>
	   	<option value="26">26</option>
	   	<option value="27">27</option>
	   	<option value="28">28</option>
	   	<option value="29">29</option>
	   	<option value="30">30</option>
		<option value="31">31</option></select>
		
		<select name="FM_start_mrd">
		<option value="<%=strMrd%>"><%=strMrdNavn%></option>
		<option value="1">jan</option>
	   	<option value="2">feb</option>
	   	<option value="3">mar</option>
	   	<option value="4">apr</option>
	   	<option value="5">maj</option>
	   	<option value="6">jun</option>
	   	<option value="7">jul</option>
	   	<option value="8">aug</option>
	   	<option value="9">sep</option>
	   	<option value="10">okt</option>
	   	<option value="11">nov</option>
	   	<option value="12">dec</option></select>
		
		
		<select name="FM_start_aar">
		<option value="<%=strAar%>">
		<%if id <> 0 AND func <> "opret" then%>
		20<%=strAar%>
		<%else%>
		<%=strAar%>
		<%end if%></option>
		<option value="2002">2002</option>
		<option value="2003">2003</option>
	   	<option value="2004">2004</option>
	   	<option value="2005">2005</option>
		<option value="2006">2006</option>
		<option value="2007">2007</option></select>&nbsp;&nbsp;<a href="javascript:NewWin_popupcal('../inc/regular/popupcalender_inc.asp?use=1')"><img src="../ill/popupcal.gif" alt="" border="0" width="16" height="15"></a>
		</td>
		<td>&nbsp;
		</td>
			<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="40" alt="" border="0"></td>
	</tr>
	</table>
	</div>
	
	<%if func = "red" then
	vizi1B = "visible"
	show1B = "" 
	else
	vizi1B = "hidden"
	show1B = "none"
	end if%>
	
	<div id=div1B name=div1B style="position:relative; visibility:<%=vizi1B%>; display:<%=show1B%>; z-index:20;">
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF" width="360">
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="60" alt="" border="0"></td>
		<td style="padding-left:10;" colspan="2" valign=top><br>
		<%
		'** Opfølgning / slut dato ***
		if func = "opret" OR func = "red" then
			
			if func = "opret" then 
			viz = "hidden"
			else
			viz = "visible"
			end if
			
			if func = "red" then%>
			Aktionen løber over flere dage:
				<%
				if len(crmDato_slut) <> 0 then
					if cdate(crmDato_slut) <> cdate(crmDato) then 
					chk = "checked"
					else
					chk = ""
					end if
				else
				chk = ""
				end if
			%><b>Ja:</b><input type="checkbox" <%=chk%> name="FM_opf" value="2"><br>
			Slut dato:
			<%end if
			
			if func = "opret" then
			opfdato = strDag&"/"&strMrd&"/"&strAar
			useopfdato = dateadd("d", 1, opfdato)
			else
				if len(crmDato_slut) <> 0 then
				useopfdato = crmDato_slut
				else
				useopfdato = crmDato
				end if
			end if
			
			useopfdag = day(useopfdato)
			useopfmrd = month(useopfdato)
			useopfmrdNavn = left(monthname(useopfmrd), 3)
			useopfaar = right(year(useopfdato), 2) 
			%>
			<img src="../ill/blank.gif" width="50" height="1" alt="" border="0">
			<select name="FM_start_dag_2">
			<option value="<%=useopfdag%>"><%=useopfdag%></option> 
			<option value="1">1</option>
		   	<option value="2">2</option>
		   	<option value="3">3</option>
		   	<option value="4">4</option>
		   	<option value="5">5</option>
		   	<option value="6">6</option>
		   	<option value="7">7</option>
		   	<option value="8">8</option>
		   	<option value="9">9</option>
		   	<option value="10">10</option>
		   	<option value="11">11</option>
		   	<option value="12">12</option>
		   	<option value="13">13</option>
		   	<option value="14">14</option>
		   	<option value="15">15</option>
		   	<option value="16">16</option>
		   	<option value="17">17</option>
		   	<option value="18">18</option>
		   	<option value="19">19</option>
		   	<option value="20">20</option>
		   	<option value="21">21</option>
		   	<option value="22">22</option>
		   	<option value="23">23</option>
		   	<option value="24">24</option>
		   	<option value="25">25</option>
		   	<option value="26">26</option>
		   	<option value="27">27</option>
		   	<option value="28">28</option>
		   	<option value="29">29</option>
		   	<option value="30">30</option>
			<option value="31">31</option></select>
			
			<select name="FM_start_mrd_2">
			<option value="<%=useopfmrd%>"><%=useopfmrdNavn%></option>
			<option value="1">jan</option>
		   	<option value="2">feb</option>
		   	<option value="3">mar</option>
		   	<option value="4">apr</option>
		   	<option value="5">maj</option>
		   	<option value="6">jun</option>
		   	<option value="7">jul</option>
		   	<option value="8">aug</option>
		   	<option value="9">sep</option>
		   	<option value="10">okt</option>
		   	<option value="11">nov</option>
		   	<option value="12">dec</option></select>
			
			<select name="FM_start_aar_2">
			<option value="<%=useopfaar%>">20<%=useopfaar%></option>
			<option value="2002">2002</option>
			<option value="2003">2003</option>
		   	<option value="2004">2004</option>
		   	<option value="2005">2005</option>
			<option value="2006">2006</option>
			<option value="2007">2007</option></select>&nbsp;&nbsp;
			<a href="javascript:NewWin_popupcal('../inc/regular/popupcalender_inc.asp?use=2')"><img src="../ill/popupcal.gif" alt="" border="0" width="16" height="15"></a>
		<%end if%>
		</td>
		<td>&nbsp;
		</td>
			<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="60" alt="" border="0"></td>
		</tr>
		</table>
		</div>
		
		
		
		
	
	<div id=div2A name=div2A style="position:relative; visibility:visible; display:; z-index:20;">
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF" width="360">
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
		<td style="padding-left:10;">Start tidspunkt:
		</td><td><select name="klok_timer">
		<option value="<%=klok_timer%>" SELECTED><%=klok_timer%></option>
		<option value="07">7</option>
	   	<option value="08">8</option>
	   	<option value="09">9</option>
	   	<option value="10">10</option>
	   	<option value="11">11</option>
	   	<option value="12">12</option>
	   	<option value="13">13</option>
	   	<option value="14">14</option>
	   	<option value="15">15</option>
	   	<option value="16">16</option>
	   	<option value="17">17</option>
	   	<option value="18">18</option>
	   	<option value="19">19</option>
	   	<option value="20">20</option>
	   	<option value="21">21</option>
	  	</select>&nbsp;:
		
		<select name="klok_min">
		<%if func <> "opret" then%>
		<option value="<%=klok_min%>" selected><%=klok_min%></option>
		<%end if%>
		<option value="00">00</option>
		<option value="15">15</option>
	   	<option value="30">30</option>
	   	<option value="45">45</option>
		</select>
		</td>
		<td>&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
		</tr>
		
		
		<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
		<td style="padding-left:10;">Slut tidspunkt:
		</td><td>
		<select name="klok_timer_slut">
		<option value="<%=klok_timer_slut%>" SELECTED><%=klok_timer_slut%></option>
		<option value="07">7</option>
	   	<option value="08">8</option>
	   	<option value="09">9</option>
	   	<option value="10">10</option>
	   	<option value="11">11</option>
	   	<option value="12">12</option>
	   	<option value="13">13</option>
	   	<option value="14">14</option>
	   	<option value="15">15</option>
	   	<option value="16">16</option>
	   	<option value="17">17</option>
	   	<option value="18">18</option>
	   	<option value="19">19</option>
	   	<option value="20">20</option>
	   	<option value="21">21</option>
	  	</select>&nbsp;:
		
		<select name="klok_min_slut">
		<%if func <> "opret" then%>
		<option value="<%=klok_min_slut%>" selected><%=klok_min_slut%></option>
		<%else%>
		<option value="30" SELECTED>30</option>
		<%end if%>
		<option value="00">00</option>
		<option value="15">15</option>
	   	<option value="30">30</option>
	   	<option value="45">45</option>
		</select>
		</td>
		<td>&nbsp;
		</td>
			<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
		</tr>
		
		<tr>
			<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
			<td style="padding-left:10;">Emne:</td>
			<td><select name="FM_emne">
			<%strSQL = "SELECT id, navn FROM crmemne ORDER BY navn"
			oRec.open strSQL, oConn, 3
			While not oRec.EOF %>
				<%if intEmne = oRec("id") then
				selected = "selected"
				else
				selected = ""
				end if%> 
			<option value="<%=oRec("id")%>" <%=selected%>><%=oRec("navn")%></option>
			<%oRec.movenext
			wend
			oRec.close%>
		</select></td>
		<td>&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
		</tr>
		<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
		<td style="padding-left:10;">Overskrift:</td>
		<td><input type="text" name="FM_navn" value="<%=strNavn%>" size="25" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		</td>
		<td>&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
		<td style="padding-left:10;">Kontakt person:</td>
		<td><select name="FM_kpers">
		<%strSQL = "SELECT kid, kontaktpers1, kontaktpers2, kontaktpers3, kontaktpers4, kontaktpers5 FROM kunder WHERE Kid = "& id
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then%>
			<%if len(trim(oRec("kontaktpers1"))) <> 0 then %>
			<option value="<%=oRec("kontaktpers1")%>"><%=oRec("kontaktpers1")%></option>
			<%end if%>
			<%if len(trim(oRec("kontaktpers2"))) <> 0 then %>
			<option value="<%=oRec("kontaktpers2")%>"><%=oRec("kontaktpers2")%></option>
			<%end if%>
			<%if len(trim(oRec("kontaktpers3"))) <> 0 then %>
			<option value="<%=oRec("kontaktpers3")%>"><%=oRec("kontaktpers3")%></option>
			<%end if%>
			<%if len(trim(oRec("kontaktpers4"))) <> 0 then %>
			<option value="<%=oRec("kontaktpers4")%>"><%=oRec("kontaktpers4")%></option>
			<%end if%>
			<%if len(trim(oRec("kontaktpers5"))) <> 0 then %>
			<option value="<%=oRec("kontaktpers5")%>"><%=oRec("kontaktpers5")%></option>
			<%end if%>
		<%end if
		oRec.close%>
		<option value="Ingen">Ingen</option>
	</select></td>
	<td>&nbsp;</td>
	<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="165" alt="" border="0"></td>
		<td valign=top colspan="3" style="padding-left:10;">Beskrivelse/ Log:<br>
		<textarea cols="<%=trearsize%>" rows="8" name="FM_besk" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"><%=strBesk%></textarea>
		&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="165" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
		<td style="padding-left:10;">Status:</td>
		<td><select name="FM_status">
		<%strSQL = "SELECT id, navn FROM crmstatus ORDER BY navn"
		oRec.open strSQL, oConn, 3
		While not oRec.EOF %>
		<%if intStatus = oRec("id") then
		selected = "selected"
		else
		selected = ""
		end if%> 
		<option value="<%=oRec("id")%>" <%=selected%>><%=oRec("navn")%></option>
		<%oRec.movenext
		wend
		oRec.close%>
		</select></td>
		<td>&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
		</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
		<td style="padding-left:10;">Kontaktform:</td>
		<td><select name="FM_kform">
			<%strSQL = "SELECT id, navn FROM crmkontaktform ORDER BY navn"
			oRec.open strSQL, oConn, 3
			While not oRec.EOF %>
			<%if intKontaktform = oRec("id") then
			selected = "selected"
			else
			selected = ""
			end if%> 
			<option value="<%=oRec("id")%>" <%=selected%>><%=oRec("navn")%></option>
			<%oRec.movenext
			wend
			oRec.close%>
		</select></td>
		<td>&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td colspan="3" style="padding-left:10;"><br>Tilføj medarbejdere:</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<%
		
		strSQL = "SELECT mid, mnavn FROM medarbejdere WHERE Brugergruppe = 1 OR Brugergruppe = 3 OR Brugergruppe = 6 ORDER BY mnavn"
		oRec.open strSQL, oConn, 0, 1
		while not oRec.EOF
		
			if func = "red" then
				strSQL2 = "SELECT medarbid FROM aktionsrelationer WHERE aktionsid = "& crmaktion &" AND medarbid = "& oRec("mid")
				oRec2.open strSQL2, oConn, 0, 1
				if not oRec2.EOF then
				strcheckmedarb = "CHECKED"
				else
				strcheckmedarb = " "
				end if
				oRec2.close
			else
				if oRec("mnavn") = session("user") then
				strcheckmedarb = "CHECKED"
				else
				strcheckmedarb = " "
				end if
			end if
		
		%> 
		<tr>
			<td colspan="2" style="padding-left:10; border-left: 1px solid #003399;"><%=oRec("mnavn")%></td><td colspan="3" style="padding-left:10; border-right: 1px solid #003399;"><input type="checkbox" name="FM_medarbrel" value="<%=oRec("mid")%>" <%=strcheckmedarb%> style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
			</td>
		</tr>
		<%
		x = x + 1
		oRec.movenext
		wend
		oRec.close
		%>
		
		</table>
		</div>
		
	
	<div id=div2B name=div2B style="position:relative; visibility:hidden; display:none; z-index:20;">
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF" width="360">
	<!--#include file="inc/crmhistorik_opf_inc.asp"-->
	</table>
	</div>
	
	
	<div id=div1C name=div1C style="position:relative; visibility:hidden; display:none; z-index:20;">
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF" width="360">
	<tr bgcolor="#5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
		<td colspan=3 valign="top"><img src="../ill/tabel_top.gif" width="<%=table_topgif%>" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td colspan="3" style="height:20; padding-left:10;" class=alt><b>Gentag aktion:</b>&nbsp;</td>
	</tr>
	<tr>
		<td valign="top" rowspan="3"><img src="../ill/tabel_top.gif" width="1" height="360" alt="" border="0"></td>
		<td style="padding-left:10;" valign=top colspan="3"><br>
		<b>Hvor ofte skal aktionen gentages:</b>
		</td>
		<td valign="top" align="right" rowspan="3"><img src="../ill/tabel_top.gif" width="1" height="360" alt="" border="0"></td>
	</tr>
	<tr>
		<td style="padding-left:10;" valign=top>Vælg <br><input type="radio" name="FM_dageldato" value="dag" checked><b>Ugedage</b><br>
		<input type="checkbox" name="FM_gentag_ugedag" value="1">&nbsp;Søndag<br>
		<input type="checkbox" name="FM_gentag_ugedag" value="2">&nbsp;Mandag<br>
		<input type="checkbox" name="FM_gentag_ugedag" value="3">&nbsp;Tirsdag<br>
		<input type="checkbox" name="FM_gentag_ugedag" value="4">&nbsp;Onsdag<br>
		<input type="checkbox" name="FM_gentag_ugedag" value="5">&nbsp;Torsdag<br>
		<input type="checkbox" name="FM_gentag_ugedag" value="6">&nbsp;Fredag<br>
		<input type="checkbox" name="FM_gentag_ugedag" value="7">&nbsp;Lørdag<br>
		<br><br>
		
		</td>
		<td valign=top width="190">eller <br><input type="radio" name="FM_dageldato" value="dato"><b>Datoer</b><br>
		<input type="checkbox" name="FM_gentag_dato" value="1">&nbsp;1
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="2">&nbsp;2
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="3">&nbsp;3<br>
		<input type="checkbox" name="FM_gentag_dato" value="4">&nbsp;4
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="5">&nbsp;5
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="6">&nbsp;6<br>
		<input type="checkbox" name="FM_gentag_dato" value="7">&nbsp;7
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="8">&nbsp;8
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="9">&nbsp;9<br>
		<input type="checkbox" name="FM_gentag_dato" value="10">&nbsp;10
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="11">&nbsp;11
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="12">&nbsp;12<br>
		<input type="checkbox" name="FM_gentag_dato" value="13">&nbsp;13
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="14">&nbsp;14
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="15">&nbsp;15<br>
		
		<input type="checkbox" name="FM_gentag_dato" value="16">&nbsp;16
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="17">&nbsp;17
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="18">&nbsp;18<br>
		
		<input type="checkbox" name="FM_gentag_dato" value="19">&nbsp;19
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="20">&nbsp;20
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="21">&nbsp;21<br>
		
		<input type="checkbox" name="FM_gentag_dato" value="22">&nbsp;22
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="23">&nbsp;23
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="24">&nbsp;24<br>
		
		<input type="checkbox" name="FM_gentag_dato" value="25">&nbsp;25
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="26">&nbsp;26
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="27">&nbsp;27<br>
		
		<input type="checkbox" name="FM_gentag_dato" value="28">&nbsp;28
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="29">&nbsp;29
		&nbsp;&nbsp;<input type="checkbox" name="FM_gentag_dato" value="30">&nbsp;30<br>
		
		<input type="checkbox" name="FM_gentag_dato" value="31">&nbsp;31<br></td>
		<td valign=top>og <br><img src="ill/blank.gif" width="1" height="4" alt="" border="0"><br><b>Måneder:</b><br>
		<input type="checkbox" name="FM_gentag_md" value="1">&nbsp;Jan<br>
		<input type="checkbox" name="FM_gentag_md" value="2">&nbsp;Feb<br>
		<input type="checkbox" name="FM_gentag_md" value="3">&nbsp;Mar<br>
		<input type="checkbox" name="FM_gentag_md" value="4">&nbsp;Apr<br>
		<input type="checkbox" name="FM_gentag_md" value="5">&nbsp;Maj<br>
		<input type="checkbox" name="FM_gentag_md" value="6">&nbsp;Jun<br>
		<input type="checkbox" name="FM_gentag_md" value="7">&nbsp;Jul<br>
		<input type="checkbox" name="FM_gentag_md" value="8">&nbsp;Aug<br>
		<input type="checkbox" name="FM_gentag_md" value="9">&nbsp;Sep<br>
		<input type="checkbox" name="FM_gentag_md" value="10">&nbsp;Okt<br>
		<input type="checkbox" name="FM_gentag_md" value="11">&nbsp;Nov<br>
		<input type="checkbox" name="FM_gentag_md" value="12">&nbsp;Dec<br></td>
	</tr>
	<tr>
		<td colspan="3" align=center><br><b>Slut gentagelser den:</b><br>
		<select name="FM_gentag_dag">
		<option value="<%=strDag%>"><%=strDag%></option> 
		<option value="1">1</option>
	   	<option value="2">2</option>
	   	<option value="3">3</option>
	   	<option value="4">4</option>
	   	<option value="5">5</option>
	   	<option value="6">6</option>
	   	<option value="7">7</option>
	   	<option value="8">8</option>
	   	<option value="9">9</option>
	   	<option value="10">10</option>
	   	<option value="11">11</option>
	   	<option value="12">12</option>
	   	<option value="13">13</option>
	   	<option value="14">14</option>
	   	<option value="15">15</option>
	   	<option value="16">16</option>
	   	<option value="17">17</option>
	   	<option value="18">18</option>
	   	<option value="19">19</option>
	   	<option value="20">20</option>
	   	<option value="21">21</option>
	   	<option value="22">22</option>
	   	<option value="23">23</option>
	   	<option value="24">24</option>
	   	<option value="25">25</option>
	   	<option value="26">26</option>
	   	<option value="27">27</option>
	   	<option value="28">28</option>
	   	<option value="29">29</option>
	   	<option value="30">30</option>
		<option value="31">31</option></select>
		
		<select name="FM_gentag_mrd">
		<option value="<%=strMrd%>"><%=strMrdNavn%></option>
		<option value="1">jan</option>
	   	<option value="2">feb</option>
	   	<option value="3">mar</option>
	   	<option value="4">apr</option>
	   	<option value="5">maj</option>
	   	<option value="6">jun</option>
	   	<option value="7">jul</option>
	   	<option value="8">aug</option>
	   	<option value="9">sep</option>
	   	<option value="10">okt</option>
	   	<option value="11">nov</option>
	   	<option value="12">dec</option></select>
		
		
		<select name="FM_gentag_aar">
		<option value="<%=strAar%>">
		<%if id <> 0 AND func <> "opret" then%>
		20<%=strAar%>
		<%else%>
		<%=strAar%>
		<%end if%></option>
		<option value="2002">2002</option>
		<option value="2003">2003</option>
	   	<option value="2004">2004</option>
	   	<option value="2005">2005</option>
		<option value="2006">2006</option>
		<option value="2007">2007</option></select>&nbsp;&nbsp;<a href="javascript:NewWin_popupcal('../inc/regular/popupcalender_inc.asp?use=3')"><img src="../ill/popupcal.gif" alt="" border="0" width="16" height="15"></a>
		</td>
	</tr>
	</table>
	</div>
	
	
	
	<%end if 'vælg kunde%>
	<table cellspacing="0" cellpadding="0" border="0" width="360" bgcolor="#EFF3FF">
	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan=3 valign="bottom"><img src="../ill/tabel_top.gif" width="<%=table_topgif%>" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#d6dff5">
		<td colspan="5"><br><br><img src="ill/blank.gif" width="120" height="1" alt="" border="0"><input type="image" name="subm" src="../ill/<%=varSubVal%>.gif"></td>
	</tr>
	</form>
	</table>
	<br><br>
	<br>
	<br>
	<br>
	</div>
	<%case else
	emner = request("emner")
	status = request("status")
	medarb = request("medarb")
	
	function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
	end function
	
	'** Hvis der søges på telefon nummer **
	if len(request("sogekri")) <> 0 then
	sogekri = trim(SQLBless(request("sogekri")))
	kundeKritlf = " Kkundenavn LIKE '%"& sogekri &"%' OR telefon LIKE '" & sogekri &"%'"_
	&" OR telefonmobil LIKE '%"& sogekri &"%'"_
	&" OR telefonalt LIKE '"& sogekri & "%' OR email LIKE '"& sogekri &"%'"_
	&" OR kontaktpers1 LIKE '%"& sogekri &"%'"_
	&" OR kontaktpers2 LIKE '%"& sogekri &"%'"_
	&" OR kontaktpers3 LIKE '%"& sogekri &"%'"_
	&" OR kontaktpers4 LIKE '%"& sogekri &"%'"_
	&" OR kontaktpers5 LIKE '%"& sogekri &"%'"_
	&" OR kpersemail1 LIKE '"& sogekri &"%'"_
	&" OR kpersemail2 LIKE '"& sogekri &"%' OR kpersemail3 LIKE '"& sogekri &"%'"_
	&" OR kpersemail4 LIKE '"& sogekri &"%' OR kpersemail5 LIKE '"& sogekri &"%'"_
	&" OR kperstlf1 LIKE '"& sogekri &"%' OR kperstlf2 LIKE '"& sogekri &"%'"_
	&" OR kperstlf3 LIKE '"& sogekri &"%'"_
	&" OR kperstlf4 LIKE '"& sogekri &"%' OR kperstlf5 LIKE '"& sogekri &"%'"_
	&" OR kpersmobil1 LIKE '"& sogekri &"%'"_
	&" OR kpersmobil2 LIKE '"& sogekri &"%' OR kpersmobil3 LIKE '"& sogekri &"%'"_
	&" OR kpersmobil4 LIKE '"& sogekri &"%' OR kpersmobil5 LIKE '"& sogekri &"%' "
	else
	sogekri = ""
	kundeKritlf = " Kid =" & id &""
	end if
	%>
	
	<!--#include file="../inc/regular/header_inc.asp"-->
	<script language="javascript">
	<!--
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	//-->
	</script>
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!--#include file="../inc/regular/rmenu.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:190; top:50; width:60%; height:600; visibility:visible;">
	<br>
	<%
	'**************************** Søgefunktion *********************************************'
	%>
	<span style="position:absolute; top:5; left:538;">
	<table cellspacing="0" cellpadding="0" border="0" width="120">
	<tr>
		<td><img src="../ill/v-menu_soeg_t.gif" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#EFF3FF">
		<td height="30" valign="middle" align="center">	
		<table>
		<tr><form action="crmhistorik.asp?menu=crm&func=hist&selpkt=hist" method="post">
		<td><input type="text" name="sogekri" size="28" maxlength="28" style="font-size : 11px;"></td>
		<td><input type="image" src="../ill/pil_groen_lille.gif" border="0"></td>
		</tr></form>
		</table></td>
		</tr>
	</table>
	</span>
	<%
	'**************************************************************************************
	%>

	
	
	<table cellspacing="0" cellpadding="0" border="0" width="740">
	<tr>
    <td valign="top"><br>
	<img src="../ill/header_crmhist.gif" alt="" border="0"><hr align="left" width="740" size="1" color="#000000" noshade>
	</td></tr></table>
	<br>
	
	<%
	'****************************************************************************************** 
	'SEARCH RESULTS.... Finder kunde udfra telefon, email navn etc. 
	'*******************************************************************************************
	if len(sogekri) <> 0 then
	Response.write "Der blev søgt på: <b>" & sogekri &"</b><br>" 
	Response.write "Følgende firmaer matchede søgningen:<br><br>"
	end if
	
	strSQL = "SELECT kkundenavn, Kid FROM kunder WHERE "& kundeKritlf& " ORDER BY kkundenavn"
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	valgtKunde = oRec("kkundenavn")
	id = oRec("Kid")
	
	if len(sogekri) <> 0 then
	Response.write "<a href='crmhistorik.asp?menu=crm&id="&oRec("Kid")&"&func=hist&selpkt=hist'>"& oRec("kkundenavn") & "&nbsp;</a><br>" 
	x = 1
	end if
	
	oRec.movenext
	wend
	oRec.close

	if x <> 1 AND len(sogekri) <> 0 then
	Response.write "<img src='../ill/alert.gif' width='44' height='45' alt='' border='0'>Der blev <b>ikke</b> fundet nogen emner der matchede søgekriteriet!"
	end if
	
	x = 0
	
	'********************************************************************************************
	' Udskriver aktions historik liste
	'********************************************************************************************
	if len(sogekri) = 0 then%>
	<div id="opretny" style="position:absolute; left:390; top:40; visibility:visible;">
	<a href="javascript:NewWin_popupaktion('crmhistorik.asp?menu=crm&func=opret&id=<%=id%>&ketype=e&showinwin=j')">Opret ny aktion <img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a>
	</div>
	<table cellspacing="0" cellpadding="0" border="0" width="740">
	<tr bgcolor="#5582D2">
	<td width="3"><img src="../ill/venstre_hjorne.gif" alt="" border="0"></td>
	<td width="730" style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 0px; border-color : #003399; border-style : solid;" class='alt'>&nbsp;Aktions historik&nbsp;&nbsp;&nbsp;<b><%=valgtKunde%></b>
	<img src="../ill/blank.gif" width="375" height="1" alt="" border="0"></td>
	<td width="3" align="right"><img src="../ill/hojre_hjorne.gif" alt="" border="0"></td>
	</tr>
	</table>
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr>
    <td valign="top">
	<table cellspacing="0" cellpadding="0" border="0" width="740" bgcolor="#EFF3FF">
	<% 
	'************************************************************************************ 
	'Table header 
	'************************************************************************************
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030%>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td width="160" height=20>&nbsp;&nbsp;<a href="crmhistorik.asp?menu=crm&sort=firm&func=hist&id=<%=id%>&selpkt=hist&emner=<%=emner%>&status=<%=status%>&medarb=<%=medarb%>"><b>Firma</b></a></td>
		<td width="120"><a href="crmhistorik.asp?menu=crm&sort=dato&func=hist&id=<%=id%>&selpkt=hist&emner=<%=emner%>&status=<%=status%>&medarb=<%=medarb%>"><b>Dato</b></a></td>
		<td width="110"><a href="crmhistorik.asp?menu=crm&sort=emne&func=hist&id=<%=id%>&selpkt=hist&emner=<%=emner%>&status=<%=status%>&medarb=<%=medarb%>"><b>Emne</b></a></td>
		<td width="100"><b>Kommentar</b></td>
		<td><a href="crmhistorik.asp?menu=crm&sort=status&func=hist&id=<%=id%>&selpkt=hist&emner=<%=emner%>&status=<%=status%>&medarb=<%=medarb%>"><b>Status</b></a></td>
		<td align=center>&nbsp;</td>
		<td width="90"><a href="crmhistorik.asp?menu=crm&sort=editor&func=hist&id=<%=id%>&selpkt=hist&emner=<%=emner%>&status=<%=status%>&medarb=<%=medarb%>"><b>Oprettet af</b></a></td>
		<td>&nbsp;</td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<%
	select case level 
	case "1", "2"
			if medarb = 0 OR len(medarb) = 0 then
			usemedarbKri = " AND editorid <> 0 "
			else
			usemedarbKri = " AND aktionsid = crmhistorik.id AND medarbid = '" & medarb & "' "
			end if
	case else
	usemedarbKri = " AND aktionsid = crmhistorik.id AND medarbid = '" & session("Mid") & "' "
	end select
	
	if id = 0 OR len(id) = 0 then
	useidKri = " kundeid <> " & id & ""
	else
	useidKri = " kundeid = " & id & ""
	end if
	
	if emner = 0 OR len(emner) = 0 then
	useemneKri = " "
	else
	useemneKri = " AND kontaktemne = " & emner & ""
	end if
	
	if status = 0 OR len(status) = 0 then
	usestatKri = " "
	else
	usestatKri = " AND status = " & status & ""
	end if
	
	
	sort = Request("sort")
	select case sort
	case "emne" 
	strSQL = "SELECT DISTINCT(crmhistorik.id), kid, kkundenavn, crmhistorik.editor, crmdato, kontaktemne, crmhistorik.navn, crmemne.navn AS enmenavn, crmstatus.navn AS statusnavn, crmkontaktform.navn AS kontaktform, ikon, crmhistorik.kontaktpers FROM crmhistorik, crmemne, crmstatus, crmkontaktform, kunder, aktionsrelationer WHERE "& useidKri &" "& useemneKri &""& usestatKri &""& usemedarbKri &" AND crmemne.id = kontaktemne AND crmstatus.id = status AND crmkontaktform.id = kontaktform AND kid = kundeid ORDER BY kontaktemne, crmhistorik.crmdato DESC"
	case "editor" 
	strSQL = "SELECT DISTINCT(crmhistorik.id), kid, kkundenavn, crmhistorik.editor, crmdato, kontaktemne, crmhistorik.navn, crmemne.navn AS enmenavn, crmstatus.navn AS statusnavn, crmkontaktform.navn AS kontaktform, ikon, crmhistorik.kontaktpers FROM crmhistorik, crmemne, crmstatus, crmkontaktform, kunder, aktionsrelationer WHERE "& useidKri &" "& useemneKri &""& usestatKri &""& usemedarbKri &" AND crmemne.id = kontaktemne AND crmstatus.id = status AND crmkontaktform.id = kontaktform AND kid = kundeid ORDER BY crmhistorik.editor, crmhistorik.crmdato DESC"
	case "status"
	strSQL = "SELECT DISTINCT(crmhistorik.id), kid, kkundenavn, crmhistorik.editor, crmdato, kontaktemne, crmhistorik.navn, crmemne.navn AS enmenavn, crmstatus.navn AS statusnavn, crmkontaktform.navn AS kontaktform, ikon, crmhistorik.kontaktpers FROM crmhistorik, crmemne, crmstatus, crmkontaktform, kunder, aktionsrelationer WHERE "& useidKri &" "& useemneKri &""& usestatKri &""& usemedarbKri &" AND crmemne.id = kontaktemne AND crmstatus.id = status AND crmkontaktform.id = kontaktform AND kid = kundeid ORDER BY statusnavn, crmhistorik.crmdato DESC"
	case "firm"
	strSQL = "SELECT DISTINCT(crmhistorik.id), kid, kkundenavn, crmhistorik.editor, crmdato, kontaktemne, crmhistorik.navn, crmemne.navn AS enmenavn, crmstatus.navn AS statusnavn, crmkontaktform.navn AS kontaktform, ikon, crmhistorik.kontaktpers FROM crmhistorik, crmemne, crmstatus, crmkontaktform, kunder, aktionsrelationer WHERE "& useidKri &" "& useemneKri &""& usestatKri &""& usemedarbKri &" AND crmemne.id = kontaktemne AND crmstatus.id = status AND crmkontaktform.id = kontaktform AND kid = kundeid ORDER BY kkundenavn, crmhistorik.crmdato"
	case else
	strSQL = "SELECT DISTINCT(crmhistorik.id), kid, kkundenavn, crmhistorik.editor, crmdato, kontaktemne, crmhistorik.navn, crmemne.navn AS enmenavn, crmstatus.navn AS statusnavn, crmkontaktform.navn AS kontaktform, ikon, crmhistorik.kontaktpers FROM crmhistorik, crmemne, crmstatus, crmkontaktform, kunder, aktionsrelationer WHERE "& useidKri &" "& useemneKri &""& usestatKri &""& usemedarbKri &" AND crmemne.id = kontaktemne AND crmstatus.id = status AND crmkontaktform.id = kontaktform AND kid = kundeid ORDER BY crmhistorik.crmdato DESC"
	end select
	
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	'************************************************************************************ 
	'Table Rows indhold 
	'************************************************************************************
	%>
	<tr>
		<td bgcolor="#5582D2" colspan="10"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr onmouseover="mOvr('gift',this,'#B4C7EF');" onmouseout="mOut(this,'');">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
		<td height=20>&nbsp;&nbsp;<a href="kunder.asp?menu=crm&func=red&id=<%=oRec("Kid")%>"><%=left(oRec("kkundenavn"), 18) &"</a><br>&nbsp;&nbsp;<font size='1' color='#999999'>" &oRec("kontaktpers")&"</font>"%></td>
		<td><font size="1"><%=formatdatetime(oRec("crmdato"), 1)%></font></td>
		<td><a href="crmhistorik.asp?selpkt=hist&menu=crm&func=red&id=<%=oRec("Kid")%>&crmaktion=<%=oRec("id")%>" class='vmenuglobal'><u><%=left(oRec("enmenavn"), 12)%></u></a>&nbsp;</td>
		<td><a href="crmhistorik.asp?selpkt=hist&menu=crm&func=red&id=<%=oRec("Kid")%>&crmaktion=<%=oRec("id")%>" class=rmenu>
		<%if len(oRec("navn")) > 16 then
		Response.write left(oRec("navn"), 16) &"..."
		else
		Response.write oRec("navn")
		end if
		%>
		</a>&nbsp;&nbsp;</td>
		<td><font size="1"><%=left(oRec("statusnavn"), 16)%></td>
		<td>&nbsp;&nbsp;<img src="../ill/<%=oRec("ikon")%>" border="0" alt="<%=oRec("kontaktform")%>"></td>
		<td><font size="1" color="#999999"><%=left(oRec("editor"), 16)%></td>
		<td><a href="crmhistorik.asp?selpkt=hist&menu=crm&func=slet&id=<%=oRec("Kid")%>&crmaktion=<%=oRec("id")%>&ketype=e&emner=<%=emner%>&status=<%=status%>&medarb=<%=medarb%>"><img src="../ill/slet_eks.gif" width="20" height="20" alt="" border="0"></a></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
	</tr>
	<%
	x = 0
	oRec.movenext
	wend
	%>
	<tr bgcolor="#5582D2">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan="8" valign="bottom"><img src="../ill/tabel_top.gif" width="721" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>	
	</table>
	</TD>
	</TR></TABLE>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%end if%>
<%end select%>
<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
