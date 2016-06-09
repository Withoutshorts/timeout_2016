	<%
	'*** Altid demo mode 
	session("lto") = "2.052-xxxx-B000"
	%>
	<!--#include file="../inc/connection/conn_db_inc.asp"-->
	<!--#include file="../inc/errors/error_inc.asp"-->
	
	<SCRIPT language=javascript src="inc/timereg_func.js"></script>
	<%
	'*********** Sætter sidens global variable ***********************************
	
	'* Hvis der vælges en anden bruger
	'* den den der er logget på.
	 
	 '*** Mid !***
		if len(Request("FM_use_me")) <> 0 then
		usemrn = Request("FM_use_me")
		else
		usemrn = session("Mid")
		end if
	
	'*** Midlertidig **
	usemrn = 1
	
	thisfile = "timereg"
	
	func = request("func")
	
	'**************** opdaterer guiden og viser timeregsiden ************************* 
	if func = "updguide" then
		oConn.execute("DELETE FROM timereg_usejob WHERE medarb = "& usemrn &"")
		
		useJob = request("FM_use_job")
		j = 0
		intuseJob = Split(useJob, ", ")
	   	For j = 0 to Ubound(intuseJob)
		oConn.execute("INSERT INTO timereg_usejob (medarb, jobid) VALUES ("& usemrn &", "& intuseJob(j) &")")
		next
		
		'if request("FM_visguide") = "j" then
			oConn.execute("UPDATE medarbejdere SET visguide = 1 WHERE Mid = "& usemrn &"")
			'else
			'oConn.execute("UPDATE medarbejdere SET visguide = 0 WHERE Mid = "& usemrn &"")
		'end if
	end if
	
	'Response.write "searchstring:" & request("searchstring")
	
	if len(request("searchstring")) = "0" then
		if len(request.cookies("kundetim")) <> 0 then
		searchstring = request.cookies("kundetim")
		else
		searchstring = 0
		end if
	else
	searchstring = cint(request("searchstring"))
	response.cookies("kundetim") = searchstring 
	response.cookies("kundetim").Expires = date + 20
	end if
	
	if len(request("filterA_B")) <> 0 then
	filterABsel = request("filterA_B")
	response.cookies("filterA_B") = filterABsel
	response.cookies("filterA_B").Expires = date + 20
	else
		if len(request.cookies("filterA_B")) <> 0 then
		filterABsel = request.cookies("filterA_B")
		else
		filterABsel = "a"
		end if
	end if
	
	
	'*** Kun hvis filter a er valgt benyttes kundeselect
	if filterABsel = "a" then
	searchstring = searchstring
	else
	searchstring = 0
	end if
	
	 if searchstring = 0 then
	 varSelectedJob = ""
	 else
	 varSelectedJob = " jobknr = "& searchstring &" AND " 
 	 end if	
	
	'*********** Globale Functions ********************************************
	'Bruges både af guiden og af timereg.
	
	public strMnavn
	public strMnr
	public strMType
	public brugergruppeId()
	public f
	public strSQLkri
	
	public function hentbrugergrupper(medarbid)
	f = 0
		Redim brugergruppeId(f)
		strSQL = "SELECT Mnavn, Mid, type, ProjektgruppeId, MedarbejderId FROM medarbejdere, medarbejdertyper, progrupperelationer WHERE medarbejdere.Mid="& medarbid &" AND medarbejdertyper.id = medarbejdere.Medarbejdertype AND MedarbejderId = Mid" 
			oRec.Open strSQL, oConn, 0, 1 
			
			While Not oRec.EOF
			Redim preserve brugergruppeId(f) 
				strMnavn = oRec("Mnavn")
				strMnr = oRec("Mid")
				strMType = oRec("type")
				brugergruppeId(f) = oRec("ProjektgruppeId")
				f = f + 1
				oRec.MoveNext
			Wend
		
		oRec.Close
		
		
		'Rettigheds tjeck på job
		'***********************************************************************
	  	strSQLkri =  varSelectedJob &" jobstatus = 1 AND fakturerbart = 1 AND kunder.Kid = job.jobknr AND ("
		
		for intcounter = 0 to f - 1  
	  
	  	strSQLkri = strSQLkri &" job.projektgruppe1 = "&brugergruppeId(intcounter)&""_
		&" OR "_
		&" job.projektgruppe2 = "&brugergruppeId(intcounter)&""_
		&" OR "_
		&" job.projektgruppe3 = "&brugergruppeId(intcounter)&""_
		&" OR "_
		&" job.projektgruppe4 = "&brugergruppeId(intcounter)&""_
		&" OR "_
	  	&" job.projektgruppe5 = "&brugergruppeId(intcounter)&""_
		&" OR "_
	  	&" job.projektgruppe6 = "&brugergruppeId(intcounter)&""_
		&" OR "_
	  	&" job.projektgruppe7 = "&brugergruppeId(intcounter)&""_
		&" OR "_
	  	&" job.projektgruppe8 = "&brugergruppeId(intcounter)&""_
		&" OR "_
	  	&" job.projektgruppe9 = "&brugergruppeId(intcounter)&""_
		&" OR "_
	  	&" job.projektgruppe10 = "&brugergruppeId(intcounter)&""_
		&" OR "
		
		next
	  	
		'** Trimmer de 2 sql states ***
		strSQLkri_len = len(strSQLkri)
		strSQLkri_left = strSQLkri_len - 3
		strSQLkri_use = left(strSQLkri, strSQLkri_left) 
	  	strSQLkri = strSQLkri_use & ") "
		
		
	end function
	'**************************************************************
	
	
	  '*** Vis Guiden ell. timereg
	  if func <> "updguide" then
		strSQL = "SELECT visguide FROM medarbejdere WHERE Mid = "& usemrn &""
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then
			if oRec("visguide") = 0 then
			visGuide = 0
			else
			visGuide = 1
			end if
		end if
		oRec.close
	  else
	  visGuide = 1
	  end if
	
	if func = "showguide" then
	visGuide = 0
	end if
	
	if cint(visGuide) = 0 then 
	
	
	'************************************************************************************************
	else
	'************************************************************************************************
	'*** Viser Timereg *******
	'************************************************************************************************

strdag = day(now)
strmrd = month(now)
straar = year(now)

useDate = strDag &"/"& strMrd &"/"& strAar
%>
	
	
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<!--#include file="inc/dato.asp"-->
	
	<!--#include file="inc/convertDate.asp"-->
	<!--#include file="inc/timereg_func_inc.asp"-->
	
<%



function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
end function

function SQLBless2(s)
		dim tmp2
		tmp2 = s
		tmp2 = replace(tmp2, ".", ",")
		SQLBless2 = tmp2
end function


%>
	<div id="sindhold" style="position:absolute; left:20; top:10; visibility:visible; width:530;">
	<table cellspacing="0" cellpadding="0" border="0" width="550">
	<tr>
    <td valign="top">
	<!-------------------------------Sideindhold------------------------------------->
	<img src="../ill/header_reg.gif" alt="" border="0">
	<hr align="left" width="330" size="1" color="#003399" noshade>
	
		
	
		
		
		<%
	
	'Henter de brugergrupper medarb. er medlem af 
	call hentbrugergrupper(usemrn)

%>
	</td>
</tr>
</table>

<img src="../ill/blank.gif" alt="" height="4" width="1" border="0"><br>
<table cellspacing="0" cellpadding="0" border="0" width="530">
<tr bgcolor="#5582D2">
	<td width="3"><img src="../ill/venstre_hjorne.gif" alt="" border="0"></td>
	<td width="74" style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 0px; border-color : #003399; border-style : solid;"><font class="stor-hvid">&nbsp;Eksterne job</font></td><td style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 0px; border-color : #003399; border-style : solid; padding-left:6; padding-top:1; padding-right:6;" class=alt><img src="../ill/eksternt_job_trans.gif" width="15" height="15" alt="" border="0"></td>
	<td  width="3" align="right"><img src="../ill/hojre_hjorne.gif" alt="" border="0"></td>
</tr>
</table>

<%
'***** Henter uge dage i den valgte uge ****
%>
<table cellspacing="0" cellpadding="0" border="0" width="530">
<form action="timereg_db.asp?FM_use_me=<%=usemrn%>" method="POST" name="timereg">
<input type="Hidden" name="FM_start_dag" value="<%=strdag%>">
<input type="Hidden" name="FM_start_mrd" value="<%=strmrd%>">
<input type="Hidden" name="FM_start_aar" value="<%=straar%>">
<!--<input type="Hidden" name="searchstring" value="<=searchstring%>">-->
<input type="Hidden" name="FM_mnr" value="<%=strMnr%>">
<!--#include file="inc/timereg_dage_inc.asp"-->
<%
call dageDatoer()
%>
<input type="hidden" name="year" value="<%=strRegAar%>">
<input type="hidden" name="datoSon" value="<%=tjekdag(1)%>">
<input type="hidden" name="datoMan" value="<%=tjekdag(2)%>">
<input type="hidden" name="datoTir" value="<%=tjekdag(3)%>">
<input type="hidden" name="datoOns" value="<%=tjekdag(4)%>">
<input type="hidden" name="datoTor" value="<%=tjekdag(5)%>">
<input type="hidden" name="datoFre" value="<%=tjekdag(6)%>">
<input type="hidden" name="datoLor" value="<%=tjekdag(7)%>">
</table>

<table cellspacing="0" cellpadding="0" border="0" width="530">
<tr>
	<td colspan=2 bgcolor="#5582d2" style="border-left:1px #003399 solid; padding-left:2; padding-top:2;">
		<table border=0 cellspacing=1 cellpadding=0><tr>
		<td align=center width=14 bgcolor="#FFFFe1"><font class='megetlillesort'>S</td>
		<td align=center width=13 bgcolor="#FFFFe1"><font class='megetlillesort'>M</td>
		<td align=center width=13 bgcolor="#FFFFe1"><font class='megetlillesort'>T</td>
		<td align=center width=13 bgcolor="#FFFFe1"><font class='megetlillesort'>O</td>
		<td align=center width=13 bgcolor="#FFFFe1"><font class='megetlillesort'>T</td>
		<td align=center width=13 bgcolor="#FFFFe1"><font class='megetlillesort'>F</td>
		<td align=center width=13 bgcolor="#FFFFe1"><font class='megetlillesort'>L</td>
		</tr></table>
	</td>
	<td width=200 bgcolor="#5582d2" style="border-right:1px #003399 solid; padding-left:2; padding-top:2;">
		<font class='megetlillehvid'>Timer Forkalkuleret / <font class='megetlilleblaa'>Realiseret
	</td>
</tr>
<!--
<tr style="border-left: 1px solid #003399; border-right: 1px solid #003399">
	<td colspan="3" bgcolor="#d6dff5" height=1><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
</tr>-->

<%
		
 
 '** Sætter jobnr på det sidst udskrevne job. ***
  'areJobPrinted = 0
  'strLastjobnr = 0
  'jobPrinted = "0#"
  'aktPrinted = "0#"
  intMnr = strMnr
  
  '** Array nr. ***
  de = 1

 '************************************************************************************************** 
  'Opbygger timereg SQL states
  '**************************************************************************************************
 	'** Midlertidigt
	filterABsel = "c"
	'*** Henter job fra guide ****
	Select case filterABsel
	case "a" 
			'*** Filter A ***
			if searchstring = 0 then 'Hvis Vis kunde ikke er brugt ell. der er valgt Alle!
				varUseJob = "("
				
				strSQL3 = "SELECT id, medarb, jobid FROM timereg_usejob WHERE medarb = "& usemrn &""
				oRec3.open strSQL3, oConn, 3
				
				while not oRec3.EOF 
				varUseJob = varUseJob & " job.id = "& oRec3("jobid") & " OR "
				oRec3.movenext
				wend 
				
				oRec3.close 
				
				if varUseJob = "(" then
				varUseJob = " job.id = 0 AND "
				else
				varUseJob_len = len(varUseJob)
				varUseJob_left = varUseJob_len - 3
				varUseJob_use = left(varUseJob, varUseJob_left) 
				varUseJob = varUseJob_use & ") AND "
				end if
				
			else
				varUseJob = " "
			end if
	
	case "b"
		
		'*** Filter B ***
		varUseJob = "("
		
		strSQL3 = "SELECT tmnr, tjobnr, tdato, timer, j.id AS jid FROM timer"_
		&" LEFT JOIN  job j ON (jobnr = tjobnr)"_
		&" RIGHT JOIN timereg_usejob ON (medarb = "& usemrn &" AND jobid = j.id)"_
		&" WHERE tmnr = medarb GROUP BY jid ORDER BY tid DESC LIMIT 0, 5" 
		
		'Response.write strSQL3
		'Response.flush
		oRec3.open strSQL3, oConn, 3 
		x = 0
		while not oRec3.EOF AND x < 5 
		
		if len(oRec3("jid")) <> 0 then
		varUseJob = varUseJob & " job.id = "& oRec3("jid") & " OR "
		end if
		
		x = x + 1
		oRec3.movenext
		wend
		oRec3.close 
		
		varUseJob = varUseJob & " job.id = 0) AND "
	
	case "c"
	
		'*** Filter C ***
		varUseJob = "("
		
		strSQL3 = "SELECT jobnavn, j.id AS jid FROM timereg_usejob "_
		&" RIGHT JOIN job j ON (j.id = jobid) "_
		&" WHERE medarb = "& usemrn &" GROUP BY jobnr ORDER BY j.id DESC LIMIT 0, 8"
		oRec3.open strSQL3, oConn, 3 
		x = 0
		while not oRec3.EOF 'AND x < 5 
		
		if len(oRec3("jid")) <> 0 then
		varUseJob = varUseJob & " job.id = "& oRec3("jid") & " OR "
		end if
		
		x = x + 1
		oRec3.movenext
		wend
		oRec3.close 
		
		varUseJob = varUseJob & " job.id = 0) AND "	
		
	end select
	
	
	
		'**** Finder kunder til dd ****
		strKundeKri = " AND ("
		strSQL3 = "SELECT j.id, medarb, jobid, kid, jobknr FROM timereg_usejob"_
		&" RIGHT JOIN job j ON (j.id = jobid)"_
		&" LEFT JOIN kunder ON (kid = j.jobknr) WHERE medarb = "& usemrn &" GROUP BY kid"
		'response.write strSQL3
		'Response.flush
		oRec3.open strSQL3, oConn, 3
		
		while not oRec3.EOF 
		strKundeKri = strKundeKri & " kid = "& oRec3("kid") & " OR "
		oRec3.movenext
		wend 
		
		oRec3.close 
	
	
 	'**************************** Rettigheds tjeck aktiviteter **************************
  	strSQLkri3 = " aktiviteter.job = job.id AND aktstatus = 1 AND ( "
	for intcounter = 0 to f - 1  
  
	strSQLkri3 = strSQLkri3 &" aktiviteter.projektgruppe1 = "& brugergruppeId(intcounter) &""_
	&" OR aktiviteter.projektgruppe2 = "& brugergruppeId(intcounter) &" "_
	&" OR aktiviteter.projektgruppe3 = "& brugergruppeId(intcounter) &" "_
	&" OR aktiviteter.projektgruppe4 = "& brugergruppeId(intcounter) &" "_
	&" OR aktiviteter.projektgruppe5 = "& brugergruppeId(intcounter) &" "_
	&" OR aktiviteter.projektgruppe6 = "& brugergruppeId(intcounter) &" "_
	&" OR aktiviteter.projektgruppe7 = "& brugergruppeId(intcounter) &" "_
	&" OR aktiviteter.projektgruppe8 = "& brugergruppeId(intcounter) &" "_
	&" OR aktiviteter.projektgruppe9 = "& brugergruppeId(intcounter) &" "_
	&" OR aktiviteter.projektgruppe10 = "& brugergruppeId(intcounter) &" OR "
	
	next
  	
	'** Trimmer sql states ***
	strSQLkri3_len = len(strSQLkri3)
	strSQLkri3_left = strSQLkri3_len - 3
	strSQLkri3_use = left(strSQLkri3, strSQLkri3_left)  
  	strSQLkri3 = strSQLkri3_use &") "
	'*************************************************************************************
	
  	'*******************************************************************
	'Finder startdato på den viste uge for at tjekke
	'om der er en faktura af nyere dato så 
	'indtastningen af timer kan lukkes 
	
	if datepart("ww", tjekdag(7)) = 1 then
		'if datepart("y", tjekdag(1), 0) > 7 then
		'varTjDatoUS_start = dateadd("yyyy", -1, tjekdag(1))
		'else
		'varTjDatoUS_start = varTjDatoUS_start
		'end if	
		
		'*** rettet 26/12-2004 *** 
		varTjDatoUS_start = cdate(convertDate(tjekdag(1)))
	else
	varTjDatoUS_start = cdate(convertDate(tjekdag(1)))
	end if	
	
	use_varTjDatoUS_start = convertDateYMD(varTjDatoUS_start)
	
	'Response.write varTjDatoUS_start & "<br>"
	
	'********************************************************************
	
	%>
	<!--#include file="inc/timereg_setvalues_a_inc.asp"-->
	<%
	
	lastfakdato = ""
	dtimeTtidspkt = ""
	antalguidenjobs = 0
	
	'******************************************************************** 
	'Gennemløber de job som brugeren har adgang til. 
	'Udskrivning af side starter her
	'0 jobid
	'1 jobnr
	'2 jobnavn
	'4 kundenavn
	'5 aktid
	'6 aktnavn
	'10 timer
	'12 dato
	'15 kommentar
	'16 fakturerbar
	'17 budgettimer
	'18 ikkebudgettimer
	'19 aktbudgettimer
	'20	offentlig
	'21 Kundeid (kid)
	
	'************** Main SQL Call ******************************************
	strSQL = "SELECT job.id AS id, jobnr, jobnavn, jobknr, kkundenavn, aktiviteter.id AS akt,"_
	&" aktiviteter.navn AS navn, aktiviteter.job, fakdato, f.tidspunkt AS ftidspkt, timer.timer, "_
	&" timer.TAktivitetId, timer.tdato, timer.tjobnr, timer.tidspunkt AS ttidspkt, timer.Timerkom, "_
	&" aktiviteter.fakturerbar, job.budgettimer, job.ikkebudgettimer,"_
	&" aktiviteter.budgettimer AS aktbtimer, timer.offentlig, kid FROM job, kunder"_
	&" LEFT JOIN aktiviteter ON ("& strSQLkri3 &")"_
	&" LEFT JOIN fakturaer f ON (f.jobid = job.id AND fakdato >= '"& use_varTjDatoUS_start &"')"_
	&" LEFT JOIN timer ON"_
	&" (TAktivitetId = aktiviteter.id AND Tmnr = "& intMnr &" AND (Tdato = '"& varTjDatoUS_son &"'"_
	&" OR Tdato = '"& varTjDatoUS_man &"'"_
	&" OR Tdato = '"& varTjDatoUS_tir &"'"_
	&" OR Tdato = '"& varTjDatoUS_ons &"'"_
	&" OR Tdato = '"& varTjDatoUS_tor &"'"_
	&" OR Tdato = '"& varTjDatoUS_fre &"'"_
	&" OR Tdato = '"& varTjDatoUS_lor &"')"_
	&") WHERE "& varUseJob & strSQLkri & " ORDER BY kkundenavn, jobnavn, id, navn" 
	
	'Response.write strSQL
	Response.flush
	'oRec.cachesize = 150
	oRec.Open strSQL, oConn, 3
	
	Dim aTable1Values
	lastknavn = ""
	
	
	if not oRec.EOF Then
	aTable1Values = oRec.GetRows()
	oRec.close
	
	ctime = 0
	'*** Udskriver records ****
	Dim iRowLoop, iColLoop
	For iRowLoop = 0 to UBound(aTable1Values, 2)
	
		
		'*** Hvis det er første gennemløb ****
		'*** Finder om der er flere akt. på det aktuelle job ****
		if iRowLoop > 0 then
		prevrecord = iRowLoop-1
		else
		prevrecord = iRowLoop
		end if
		
		'***** Hvis det er sidste gennemløb *******
		'*** Finder om der er flere job ****
		if iRowLoop < UBound(aTable1Values, 2) then
		nextrecord = iRowLoop + 1
		else
		nextrecord = "0"	
		end if
		
		if aTable1Values(1,iRowLoop) <> aTable1Values(1,(prevrecord)) OR iRowLoop = 0 then
		
		if len(aTable1Values(8,iRowLoop)) <> 0 then
		lastfakdato = aTable1Values(8,iRowLoop)
		lastFaktime = formatdatetime(aTable1Values(9,iRowLoop), 3)
		else
		lastfakdato = "1/13/2000"
		lastFaktime = "10:00:01 AM"
		end if
		
					if (aTable1Values(1,iRowLoop) <> aTable1Values(1,(prevrecord)) OR iRowLoop > 0) then 'AND len(aTable1Values(5,prevrecord)) <> 0 then				
					'Response.write "<tr><td colspan=9><b>Slut akt div -- "& aTable1Values(2,prevrecord) &" --</b></td></tr>"
					Response.write "</table></div></td></tr>"
					antalguidenjobs = antalguidenjobs + 1
					end if	
					
					if lastknavn <> aTable1Values(4,iRowLoop) then
					%>
					<tr>
						<td colspan=9 bgcolor="#8caae6" style="border-left:1px #003399 solid; solid; border-right:1px #003399 solid; padding-left:10; padding-top:3; border-top:1px #d6dff5 solid;" class=lysblaastor><b><%=aTable1Values(4,iRowLoop)%></b></td>
					</tr>
					<tr>
						<td bgcolor="#8caae6" height=4 colspan="9" style="border-left: 1px solid #003399; border-right: 1px solid #003399; border-bottom:1px #d6dff5 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
					</tr>
					<%end if%>
					<tr>
						<td valign=top width=30 height=20 style="background-color: #5582D2; padding-top:2px; padding-left : 2px; border-left: 1px solid #003399">
						<%
						'*** henter timer på job denne uge *** 
						call timerdage(aTable1Values(1,iRowLoop), intMnr)
						%></td>
						<td width=380 valign=top style="background-color: #5582D2; padding-left : 4px;padding-right : 4px;">
						<%
						'** Er der nogen aktiviteter **
						strSQL2 = "SELECT count(id) AS antalAkt FROM aktiviteter WHERE job = "& aTable1Values(0,iRowLoop) &""
						oRec2.Open strSQL2, oConn, 3
						if not oRec2.EOF then
						antalAkt = oRec2("antalAkt")
						else
						antalAkt = 0
						end if
						oRec2.close
							
						if antalAkt <> 0 then%>
						<a href="javascript:expand('<%=iRowLoop%>');"><img src="ill/plus.gif" width="9" height="9" border="0" name="Menub<%=iRowLoop%>" id="Menub<%=iRowLoop%>">&nbsp;</a>&nbsp;<font class='lilleblaa'>(<%=antalAkt%>)</font>&nbsp;&nbsp;
						<%else%>
						&nbsp;
						<%end if%>
						
						
						<%
						if len(aTable1Values(2,iRowLoop)) > 30 then
						jobnavnthis = left(aTable1Values(2,iRowLoop), 30) &".."
						else
						jobnavnthis = aTable1Values(2,iRowLoop)
						end if
						
						if level <= 3 OR level = 6 then%>
						<a href="jobs.asp?menu=job&func=red&id=<%=aTable1Values(0,iRowLoop)%>&int=1&rdir=treg" class="alt"><%=jobnavnthis%></a>
						<%else%>
						<font class='storhvid'><%=jobnavnthis%></font>
						<%end if%>
						
						&nbsp;<font class='lillehvid'><b>(<%=aTable1Values(1,iRowLoop)%>)</b>
						
						<!--&nbsp;
						<font class='lilleblaa'>(
						<if len(aTable1Values(4,iRowLoop)) > 20 then
						Response.write left(aTable1Values(4,iRowLoop), 20) &".."
						else
						Response.write aTable1Values(4,iRowLoop)
						end if%>
						)</font>-->
							
						<%
						'* finder antal brugte timer på job ***
						strSQL2 = "SELECT sum(timer) AS sumtimer FROM timer WHERE tjobnr = "& aTable1Values(1,iRowLoop) &" AND tfaktim <> 5 ORDER BY timer"
						oRec2.open strSQL2, oConn, 3
						if not oRec2.EOF then 
						timerbrugtthis = oRec2("sumtimer")
						end if
						oRec2.close
						
						if len(timerbrugtthis ) > 0 then
						timerbrugtthis = timerbrugtthis 
						else
						timerbrugtthis = 0
						end if
						
						%>
						</td>
						<td valign=top width=150 align=right style='padding-right:3px; background-color: #5582D2; border-right: 1px solid #003399;'><font class='megetlillehvid'><%=formatnumber(aTable1Values(17,iRowLoop) + aTable1Values(18,iRowLoop), 2)%> / <font class='megetlilleblaa'><%= formatnumber(timerbrugtthis, 2) %>
						<input type="checkbox" name="FM_flyttilguide" value="<%=aTable1Values(0,iRowLoop)%>"></td>
					</tr>
						
							
		<%
		'**** Starter div uanset om der findes aktiviteter ***
		%>
		<tr style="border-left: 1px solid #003399; border-right: 1px solid #003399">
			<td colspan="3" bgcolor="#8caae6" height=1><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		</tr>
		<tr style="border-left: 1px solid #003399; border-right: 1px solid #003399">
		<td colspan="3" width="530" bgcolor="#5582d2"><div ID="Menu<%=iRowLoop%>" Style="position: relative; display: none; width:530;">
		<table cellspacing="1" cellpadding="0" border="0" width=100% bgcolor="#ffffff">
		
		<%				
		end if
		
		'*** Hvis aktid <> 0 ****
		if len(aTable1Values(5,iRowLoop)) <> 0 then
		
			'**** Finder ugens timer på aktid ***%>
			<!--#include file="inc/timereg_setvalues_b_inc.asp"-->
			<%	
				
				'************** Aktivitetsnavn og timeforbrug i alt *********************''	
				if aTable1Values(5, iRowLoop) <> aTable1Values(5, nextrecord) OR aTable1Values(2, iRowLoop) <> aTable1Values(2, nextrecord) OR nextrecord = "0" then
					%>
					<input type="hidden" name="FM_jobnr_<%=iRowLoop%>" value="<%=aTable1Values(1,iRowLoop)%>">
					<input type="hidden" name="FM_Aid_<%=iRowLoop%>" value="<%=aTable1Values(5,iRowLoop)%>">
					
					<tr bgcolor=#D6DFF5>
						<td width=120 style='padding-left:2px; padding-top:2px;'>
						<%
						if len(aTable1Values(6, iRowLoop)) > 22 then
							Response.write "<font class=lille-kalender>"& left(aTable1Values(6, iRowLoop), 22) &"..</font>"
							else
							Response.write "<font class=lille-kalender>"& aTable1Values(6, iRowLoop) &"</font>"
						end if
						
						if aTable1Values(16,iRowLoop) = 2 then
						%>
						<img src="../ill/blank.gif" width="60" height="1" alt="" border="0">(Km)
						<%
						end if
						
						'* finder antal brugte timer på akt ***
							strSQL3 = "SELECT sum(timer) AS sumakttimer FROM timer WHERE taktivitetid = "& aTable1Values(5,iRowLoop) &" ORDER BY timer"
							oRec3.open strSQL3, oConn, 3
							if not oRec3.EOF then 
							timerbrugtAktthis = oRec3("sumakttimer")
							end if
							oRec3.close
							
							if len(timerbrugtAktthis) > 0 then
							timerbrugtAktthis = timerbrugtAktthis
							else
							timerbrugtAktthis = 0
							end if
						
						Response.write "<br><font class='megetlillesilver'>" & formatnumber(aTable1Values(19,iRowLoop), 2) &" / "& formatnumber(timerbrugtAktthis, 2)&"</font>"
						%></td>	
					<%
				end if
			
			
			
			
			if aTable1Values(5, iRowLoop) <> aTable1Values(5, nextrecord) OR aTable1Values(2, iRowLoop) <> aTable1Values(2, nextrecord) OR nextrecord = "0" then
				'**** Tjekker om der er fundet timeindtastninger på de enkelte dage ***
				%>
				<!--#include file="inc/timereg_setvalues_c_inc.asp"-->
				<%				
			 	'**** Henter feltfarver alt efter om der findes fakturaer på jobbet.
				call fakfarver()
				
				
					
				'*** Udskriver timefelter til siden ***********************	
				%>
				<td align="center">
				<input type="hidden" name="FM_son_opr_<%=iRowLoop%>" id="FM_son_opr_<%=iRowLoop%>" value="<%=sonTimerVal(sonRLoop)%>">
				<%if aTable1Values(16,iRowLoop) <> 2 then ' 2 = Kørsel (2 i akt.fakbar 5 i timer.tfaktim)/ alm aktivitet %>
				<input type="Text" name="Timer_son_<%=iRowLoop%>" id="Timer_son_<%=iRowLoop%>" maxlength="<%=maxl_son%>" value="<%=SQLBless2(sonTimerVal(sonRLoop))%>" onkeyup="setTimerTot(<%=iRowLoop%>, 'son'), tjektimer('son',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_son%>">
				<%else%>
				<input type="Text" name="Timer_son_<%=iRowLoop%>" id="Timer_son_<%=iRowLoop%>" maxlength="<%=maxl_son%>" value="<%=SQLBless2(sonTimerVal(sonRLoop))%>" onkeyup="tjekkm('son',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_son%>">
				<%end if%>
				<%call showcommentfunc(1)%>
				<a href="#" onclick="expandkomm('son',<%=iRowLoop%>);"><%=kommtrue%></a>
				</td>
				<td align="center">
				<input type="hidden" name="FM_man_opr_<%=iRowLoop%>" id="FM_man_opr_<%=iRowLoop%>" value="<%=manTimerVal(manRLoop)%>">
				<%if aTable1Values(16,iRowLoop) <> 2 then%>
				<input type="Text" name="Timer_man_<%=iRowLoop%>" id="Timer_man_<%=iRowLoop%>" maxlength="<%=maxl_man%>" value="<%=SQLBless2(manTimerVal(manRLoop))%>" onkeyup="setTimerTot(<%=iRowLoop%>, 'man'), tjektimer('man',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_man%>">
				<%else%>
				<input type="Text" name="Timer_man_<%=iRowLoop%>" id="Timer_man_<%=iRowLoop%>" maxlength="<%=maxl_man%>" value="<%=SQLBless2(manTimerVal(manRLoop))%>" onkeyup="tjekkm('man',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_man%>">
				<%end if%>
				<%call showcommentfunc(2)%>
				<a href="#" onclick="expandkomm('man',<%=iRowLoop%>);"><%=kommtrue%></a>
				</td>
				<td align="center">
				<input type="hidden" name="FM_tir_opr_<%=iRowLoop%>" id="FM_tir_opr_<%=iRowLoop%>" value="<%=tirTimerVal(tirRLoop)%>">
				<%if aTable1Values(16,iRowLoop) <> 2 then%>
				<input type="Text" name="Timer_tir_<%=iRowLoop%>" id="Timer_tir_<%=iRowLoop%>" maxlength="<%=maxl_tir%>" value="<%=SQLBless2(tirTimerVal(tirRLoop))%>" onkeyup="setTimerTot(<%=iRowLoop%>, 'tir'), tjektimer('tir',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_tir%>">
				<%else%>
				<input type="Text" name="Timer_tir_<%=iRowLoop%>" id="Timer_tir_<%=iRowLoop%>" maxlength="<%=maxl_tir%>" value="<%=SQLBless2(tirTimerVal(tirRLoop))%>" onkeyup="tjekkm('tir',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_tir%>">
				<%end if%>
				<%call showcommentfunc(3)%>
				<a href="#" onclick="expandkomm('tir',<%=iRowLoop%>);"><%=kommtrue%></a>
				</td>
				<td align="center">
				<input type="hidden" name="FM_ons_opr_<%=iRowLoop%>" id="FM_ons_opr_<%=iRowLoop%>" value="<%=onsTimerVal(onsRLoop)%>">
				<%if aTable1Values(16,iRowLoop) <> 2 then%>
				<input type="Text" name="Timer_ons_<%=iRowLoop%>" id="Timer_ons_<%=iRowLoop%>" maxlength="<%=maxl_ons%>" value="<%=SQLBless2(onsTimerVal(onsRLoop))%>" onkeyup="setTimerTot(<%=iRowLoop%>, 'ons'), tjektimer('ons',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_ons%>">
				<%else%>
				<input type="Text" name="Timer_ons_<%=iRowLoop%>" id="Timer_ons_<%=iRowLoop%>" maxlength="<%=maxl_ons%>" value="<%=SQLBless2(onsTimerVal(onsRLoop))%>" onkeyup="tjekkm('ons',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_ons%>">
				<%end if%>
				<%call showcommentfunc(4)%>
				<a href="#" onclick="expandkomm('ons',<%=iRowLoop%>);"><%=kommtrue%></a>
				</td>
				<td align="center">
				<input type="hidden" name="FM_tor_opr_<%=iRowLoop%>" id="FM_tor_opr_<%=iRowLoop%>" value="<%=torTimerVal(torRLoop)%>">
				<%if aTable1Values(16,iRowLoop) <> 2 then%>
				<input type="Text" name="Timer_tor_<%=iRowLoop%>" id="Timer_tor_<%=iRowLoop%>" maxlength="<%=maxl_tor%>" value="<%=SQLBless2(torTimerVal(torRLoop))%>" onkeyup="setTimerTot(<%=iRowLoop%>, 'tor'), tjektimer('tor',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_tor%>">
				<%else%>
				<input type="Text" name="Timer_tor_<%=iRowLoop%>" id="Timer_tor_<%=iRowLoop%>" maxlength="<%=maxl_tor%>" value="<%=SQLBless2(torTimerVal(torRLoop))%>" onkeyup="tjekkm('tor',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_tor%>">
				<%end if%>
				<%call showcommentfunc(5)%>
				<a href="#" onclick="expandkomm('tor',<%=iRowLoop%>);"><%=kommtrue%></a>
				</td>
				<td align="center">
				<input type="hidden" name="FM_fre_opr_<%=iRowLoop%>" id="FM_fre_opr_<%=iRowLoop%>" value="<%=freTimerVal(freRLoop)%>">
				<%if aTable1Values(16,iRowLoop) <> 2 then%>
				<input type="Text" name="Timer_fre_<%=iRowLoop%>" id="Timer_fre_<%=iRowLoop%>" maxlength="<%=maxl_fre%>" value="<%=SQLBless2(freTimerVal(freRLoop))%>" onkeyup="setTimerTot(<%=iRowLoop%>, 'fre'), tjektimer('fre',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_fre%>">
				<%else%>
				<input type="Text" name="Timer_fre_<%=iRowLoop%>" id="Timer_fre_<%=iRowLoop%>" maxlength="<%=maxl_fre%>" value="<%=SQLBless2(freTimerVal(freRLoop))%>" onkeyup="tjekkm('fre',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_fre%>">
				<%end if%>
				<%call showcommentfunc(6)%>
				<a href="#" onclick="expandkomm('fre',<%=iRowLoop%>);"><%=kommtrue%></a>
				</td>
				<td align="center">
				<input type="hidden" name="FM_lor_opr_<%=iRowLoop%>" id="FM_lor_opr_<%=iRowLoop%>" value="<%=lorTimerVal(lorRLoop)%>">
				<%if aTable1Values(16,iRowLoop) <> 2 then%>
				<input type="Text" name="Timer_lor_<%=iRowLoop%>" id="Timer_lor_<%=iRowLoop%>" maxlength="<%=maxl_lor%>" value="<%=SQLBless2(lorTimerVal(lorRLoop))%>" onkeyup="setTimerTot(<%=iRowLoop%>, 'lor'), tjektimer('lor',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_lor%>">
				<%else%>
				<input type="Text" name="Timer_lor_<%=iRowLoop%>" id="Timer_lor_<%=iRowLoop%>" maxlength="<%=maxl_lor%>" value="<%=SQLBless2(lorTimerVal(lorRLoop))%>" onkeyup="tjekkm('lor',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_lor%>">
				<%end if%>
				<%call showcommentfunc(7)%>
				<a href="#" onclick="expandkomm('lor',<%=iRowLoop%>);"><%=kommtrue%></a>
				</td>
			</tr>
			
			
				<%'kommentarer%>
				<input type="hidden" name="FM_kom_son_<%=iRowLoop%>" id="FM_kom_son_<%=iRowLoop%>" value="<%=sonKomm%>">
				<input type="hidden" name="FM_kom_man_<%=iRowLoop%>" id="FM_kom_man_<%=iRowLoop%>" value="<%=manKomm%>">
				<input type="hidden" name="FM_kom_tir_<%=iRowLoop%>" id="FM_kom_tir_<%=iRowLoop%>" value="<%=tirKomm%>">
				<input type="hidden" name="FM_kom_ons_<%=iRowLoop%>" id="FM_kom_ons_<%=iRowLoop%>" value="<%=onsKomm%>">
				<input type="hidden" name="FM_kom_tor_<%=iRowLoop%>" id="FM_kom_tor_<%=iRowLoop%>" value="<%=torKomm%>">
				<input type="hidden" name="FM_kom_fre_<%=iRowLoop%>" id="FM_kom_fre_<%=iRowLoop%>" value="<%=freKomm%>">
				<input type="hidden" name="FM_kom_lor_<%=iRowLoop%>" id="FM_kom_lor_<%=iRowLoop%>" value="<%=lorKomm%>">
				<%'Komm. offentlig tilgængelig for kunde%>
				<%
				if len(sonKomm_off) <> 0 then
				sonKomm_off = sonKomm_off
				else
				sonKomm_off = 0
				end if
				%>
				<input type="hidden" name="FM_off_son_<%=iRowLoop%>" id="FM_off_son_<%=iRowLoop%>" value="<%=sonKomm_off%>">
				<%
				if len(manKomm_off) <> 0 then
				manKomm_off = manKomm_off
				else
				manKomm_off = 0
				end if
				%>
				<input type="hidden" name="FM_off_man_<%=iRowLoop%>" id="FM_off_man_<%=iRowLoop%>" value="<%=manKomm_off%>">
				<%
				if len(tirKomm_off) <> 0 then
				tirKomm_off = tirKomm_off
				else
				tirKomm_off = 0
				end if
				%>
				<input type="hidden" name="FM_off_tir_<%=iRowLoop%>" id="FM_off_tir_<%=iRowLoop%>" value="<%=tirKomm_off%>">
				<%
				if len(onsKomm_off) <> 0 then
				onsKomm_off = onsKomm_off
				else
				onsKomm_off = 0
				end if
				%>
				<input type="hidden" name="FM_off_ons_<%=iRowLoop%>" id="FM_off_ons_<%=iRowLoop%>" value="<%=onsKomm_off%>">
				<%
				if len(torKomm_off) <> 0 then
				torKomm_off = torKomm_off
				else
				torKomm_off = 0
				end if
				%>
				<input type="hidden" name="FM_off_tor_<%=iRowLoop%>" id="FM_off_tor_<%=iRowLoop%>" value="<%=torKomm_off%>">
				
				<%
				if len(freKomm_off) <> 0 then
				freKomm_off = freKomm_off
				else
				freKomm_off = 0
				end if
				%>
				<input type="hidden" name="FM_off_fre_<%=iRowLoop%>" id="FM_off_fre_<%=iRowLoop%>" value="<%=freKomm_off%>">
				<%
				if len(lorKomm_off) <> 0 then
				lorKomm_off = lorKomm_off
				else
				lorKomm_off = 0
				end if
				%>
				<input type="hidden" name="FM_off_lor_<%=iRowLoop%>" id="FM_off_lor_<%=iRowLoop%>" value="<%=lorKomm_off%>">
				
					
			
			<%
			sonKomm = ""
			manKomm = ""
			tirKomm = ""
			onsKomm = ""
			torKomm = ""
			freKomm = ""
			lorKomm = ""
			
			sonKomm_off = 0
			manKomm_off = 0
			tirKomm_off = 0
			onsKomm_off = 0
			torKomm_off = 0
			freKomm_off = 0
			lorKomm_off = 0
			
			sonRLoop = ""
			manRLoop = ""
			tirRLoop = ""
			onsRLoop = ""
			torRLoop = ""
			freRLoop = ""
			lorRLoop = ""
			
			Response.flush
			
			end if
		de = iRowLoop
		end if 
		lastknavn = aTable1Values(4,iRowLoop)
	Next 
	
	'*** afslutter den sidtste div 
	if iRowLoop > 0 then 
	'Response.write "<tr><td><b>final akt div</b></td></tr>"%>
	</table></div></td></tr>
	<tr>
		<td bgcolor="#5582d2" height=5 colspan="9" style="border-left: 1px solid #003399; border-right: 1px solid #003399; border-bottom: 1px solid #003399;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr>
		<td colspan=9 align=right style="padding-right:2; padding-top:2;"><font class="megetlillesort">Marker job for at fjerne fra time-registreringen og overføre til "Guiden dine aktive job!".</td>
	</tr></table>
	<%end if%>
	
	<%
	'********************** slut gennemløb af eksterne job *****************************
	antal_de = de 
	%>
	<input type="hidden" name="FM_de" value="<%=antal_de%>">
	<%
	
	'*** Hvis der oRec.EOF true => Ikke adgang til nogen job ***************************
	else
	%>
	<table cellspacing="0" cellpadding="0" border="0" width="530">
	<tr>
		<td colspan=3 style='border-left: 1px solid #003399; border-right: 1px solid #003399; border-bottom: 1px solid #003399; background-color: #5582D2; padding-left : 38px;padding-right : 4px;' class='alt'>		
		<br>Du er <b>ikke</b> tildelt adgang til nogen eksterne-job.<br>
		Eller der er ikke nogen aktive job på den valgte kunde.<br><br>
		Kontakt din projektleder/ system administrator.<br>
		Eller brug <a href='timereg.asp?func=showguide' class='alt'><u>Guiden dine aktive job&nbsp;<font color='darkred'>!</font></u></a> for at få vist de job der er aktive på din profil.<br>&nbsp;</td>
	</tr>
	</table><%
	oRec.close
	end if
	
	
	
	
	antalinterneJobs = 0
	'**** Interne Job ****
	
	'*** antal interne job ***
	strSQL = "SELECT count(job.id) AS antalinternejob FROM job WHERE jobstatus = 1 AND fakturerbart = 0"
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	antalinterneJobs = oRec("antalinternejob")
	end if
	oRec.close 
	'Response.write antalinterneJobs
	
	'*** Midlertidig ***
	antalinterneJobs = 0
	
	if antalinterneJobs <> 0 then
	%>
	<br><br>
	<table cellspacing="0" cellpadding="0" border="0" width="530">
	<tr bgcolor="#5582D2">
	<td width="3"><img src="../ill/venstre_hjorne.gif" alt="" border="0"></td>
	<td width="74" style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 0px; border-color : #003399; border-style : solid;"><font class="stor-hvid">&nbsp;Interne job</font></td><td style="border-top : 1px; border-bottom : 0px; border-left : 0px; border-right : 0px; border-color : #003399; border-style : solid; padding-left:2; padding-top:1;"><img src="../ill/internt_job_trans.gif" width="15" height="15" alt="" border="0"></td>
	<td width="3" align="right"><img src="../ill/hojre_hjorne.gif" alt="" border="0"></td>
	</tr>
	</table>
	<table width='530' border='0' cellspacing='0' cellpadding='0'>
	<%
	call dageDatoer()
	%>
	</table>
	<table width='530' border='0' cellspacing='1' cellpadding='0' bgcolor="#FFFFFF">
	<%
	'** Udskriver Interne Jobnr **
	strLastjobnr = ""
	strSQL = "SELECT job.id, jobnr, jobnavn, kkundenavn FROM job, kunder WHERE jobstatus = 1 AND fakturerbart = 0 AND kunder.Kid = jobknr ORDER BY jobnavn"
	oRec.cachesize = 10
	oRec.Open strSQL, oConn, 0, 1
		
		di = antal_de + 1
		
		While Not oRec.EOF 
		
		
					'*** Finder timer på interne job. ****
					sonTimerValint = ""
					manTimerValint = ""
					tirTimerValint = ""
					onsTimerValint = ""
					torTimerValint = ""
					freTimerValint = ""
					lorTimerValint = ""
		
					
					strSQLtimer = "SELECT timer, tdato FROM timer WHERE Tjobnr = "&oRec("Jobnr")&" AND Tmnr = "& strMnr &" AND Tdato = '"& varTjDatoUS_son &"'"_
					&" OR "_
					&" Tjobnr = "&oRec("Jobnr")&" AND Tmnr = "&intMnr&" AND Tdato = '"& varTjDatoUS_man &"'"_
					&" OR "_
					&" Tjobnr = "&oRec("Jobnr")&" AND Tmnr = "&intMnr&" AND Tdato = '"& varTjDatoUS_tir &"'"_
					&" OR "_
					&" Tjobnr = "&oRec("Jobnr")&" AND Tmnr = "&intMnr&" AND Tdato = '"& varTjDatoUS_ons &"'"_
					&" OR "_
					&" Tjobnr = "&oRec("Jobnr")&" AND Tmnr = "&intMnr&" AND Tdato = '"& varTjDatoUS_tor &"'"_
					&" OR "_
					&" Tjobnr = "&oRec("Jobnr")&" AND Tmnr = "&intMnr&" AND Tdato = '"& varTjDatoUS_fre &"'"_
					&" OR "_
					&" Tjobnr = "&oRec("Jobnr")&" AND Tmnr = "&intMnr&" AND Tdato = '"& varTjDatoUS_lor &"'"
					
					
					oRec2.Open strSQLtimer, oConn, 0, 1  
					while not oRec2.EOF 
					
						select case DatePart("w", oRec2("tdato")) 
						case 1	
						sonTimerValint = SQLBless(oRec2("timer"))
						case 2
						manTimerValint = SQLBless(oRec2("timer"))
						case 3
						tirTimerValint = SQLBless(oRec2("timer"))
						case 4
						onsTimerValint = SQLBless(oRec2("timer"))
						case 5
						torTimerValint = SQLBless(oRec2("timer"))
						case 6
						freTimerValint = SQLBless(oRec2("timer"))
						case 7
						lorTimerValint = SQLBless(oRec2("timer"))
						end select
						
					oRec2.movenext
					wend
					oRec2.close
					'***************************************************
		
		
		'** Bgcolor på felter ***
		fmbg_son = "#FFDFDF" 
		fmbg_man = "#FFFFFF" 
		fmbg_tir = "#FFFFFF" 
		fmbg_ons = "#FFFFFF" 
		fmbg_tor = "#FFFFFF" 
		fmbg_fre = "#FFFFFF" 
		fmbg_lor = "#FFDFDF"  
		
		
		'*** Markerer dagsdato *** 
		select case weekday(useDate)
		case 1
		fmbg_son = "#eff3ff" 
		case 2
		fmbg_man = "#eff3ff" 
		case 3
		fmbg_tir = "#eff3ff" 
		case 4
		fmbg_ons = "#eff3ff"
		case 5
		fmbg_tor = "#eff3ff" 
		case 6
		fmbg_fre = "#eff3ff" 
		case 7
		fmbg_lor = "#eff3ff"
		end select
					
					
		Response.write "<tr bgcolor=#D6DFF5><td style='width : 116px; padding-left : 4px; padding-right : 4px;'><font class='lillesort'>"_
		& oRec("Jobnr") &"&nbsp;"& left(oRec("Jobnavn"), 15)&"</td>"
		%><input type="hidden" name="FM_jobnr_<%=di%>" id="FM_jobnr_<%=di%>" value="<%=oRec("jobnr")%>">
		<td align="center"><input type="hidden" name="FM_son_opr_<%=di%>" id="FM_son_opr_<%=di%>" value="<%=sonTimerValint%>"><input type="Text" name="Timer_son_<%=di%>" id="Timer_son_<%=di%>" maxlength="5" value="<%=SQLBless2(sonTimerValint)%>" onkeyup="setTimerTot(<%=di%>, 'son'), tjektimer('son',<%=di%>)"; style="background-color: <%=fmbg_son%>;" size=3></td>
		<td align="center"><input type="hidden" name="FM_man_opr_<%=di%>" id="FM_man_opr_<%=di%>" value="<%=manTimerValint%>"><input type="Text" name="Timer_man_<%=di%>" id="Timer_man_<%=di%>" maxlength="5" value="<%=SQLBless2(manTimerValint)%>" onkeyup="setTimerTot(<%=di%>, 'man'), tjektimer('man',<%=di%>)"; size=3 style="background-color: <%=fmbg_man%>;"></td>
		<td align="center"><input type="hidden" name="FM_tir_opr_<%=di%>" id="FM_tir_opr_<%=di%>" value="<%=tirTimerValint%>"><input type="Text" name="Timer_tir_<%=di%>" id="Timer_tir_<%=di%>" maxlength="5" value="<%=SQLBless2(tirTimerValint)%>" onkeyup="setTimerTot(<%=di%>, 'tir'), tjektimer('tir',<%=di%>)"; size=3 style="background-color: <%=fmbg_tir%>;"></td>
		<td align="center"><input type="hidden" name="FM_ons_opr_<%=di%>" id="FM_ons_opr_<%=di%>" value="<%=onsTimerValint%>"><input type="Text" name="Timer_ons_<%=di%>" id="Timer_ons_<%=di%>" maxlength="5" value="<%=SQLBless2(onsTimerValint)%>" onkeyup="setTimerTot(<%=di%>, 'ons'), tjektimer('ons',<%=di%>)"; size=3 style="background-color: <%=fmbg_ons%>;"></td>
		<td align="center"><input type="hidden" name="FM_tor_opr_<%=di%>" id="FM_tor_opr_<%=di%>" value="<%=torTimerValint%>"><input type="Text" name="Timer_tor_<%=di%>" id="Timer_tor_<%=di%>" maxlength="5" value="<%=SQLBless2(torTimerValint)%>" onkeyup="setTimerTot(<%=di%>, 'tor'), tjektimer('tor',<%=di%>)"; size=3 style="background-color: <%=fmbg_tor%>;"></td>
		<td align="center"><input type="hidden" name="FM_fre_opr_<%=di%>" id="FM_fre_opr_<%=di%>" value="<%=freTimerValint%>"><input type="Text" name="Timer_fre_<%=di%>" id="Timer_fre_<%=di%>" maxlength="5" value="<%=SQLBless2(freTimerValint)%>" onkeyup="setTimerTot(<%=di%>, 'fre'), tjektimer('fre',<%=di%>)"; size=3 style="background-color: <%=fmbg_fre%>;"></td>
		<td align="center"><input type="hidden" name="FM_lor_opr_<%=di%>" id="FM_lor_opr_<%=di%>" value="<%=lorTimerValint%>"><input type="Text" name="Timer_lor_<%=di%>" id="Timer_lor_<%=di%>" maxlength="5" value="<%=SQLBless2(lorTimerValint)%>" onkeyup="setTimerTot(<%=di%>, 'lor'), tjektimer('lor',<%=di%>)"; style="background-color: <%=fmbg_lor%>;" size=3></td>
		</tr>
		<%
		showWeekInt = 1
		di = di + 1
		strLastjobnr = oRec("Jobnr")
	oRec.MoveNext
	Wend 
	oRec.Close
	
%>
</table>
<%
else
di = antal_de + 1
end if


'*** Antal interne job til timreg_db.asp
antal_di = di - 1
%>
<input type="hidden" name="FM_di" value="<%=antal_di%>">
<input type="hidden" name="FM_strweek" value="<%=strWeek%>">


<table width='530' border='0' cellspacing='1' cellpadding='0'>
<tr>
    <td align="right"><br><input type="image" src="../ill/frem.gif"></td>
</tr>
</table>


<!--
</td>
</tr>
</table>-->




	<%'***** kommentar layer *** %>
	<div id="kom" name="kom" Style="position: absolute; display: none; left:20; top:170; z-index:2994; width:500; height:140; background-color:#D6DFF5; padding-top:3; padding-left:35; !border: 1px; border-color: #000000; border-style: solid;">
	Kommentar:
	<input type="hidden" name="daytype" id="daytype" value="son">
	<input type="hidden" name="rowcounter" id="rowcounter" value="0">
	<img src="../ill/blank.gif" width="190" height="1" alt="" border="0">
	(maks. 255 karakterer)&nbsp;
	<input type="text" name="antch" id="antch" size="3" maxlength="4"><br>
	<textarea cols="50" rows="6" id="FM_kom" name="FM_kom" onKeyup="antalchar();"></textarea><br>
	Skjul kommentar på kundeside:
	<select name="FM_off" id="FM_off" style="font-size:9px;">
	<option value="0" SELECTED>Nej</option>
	<option value="1">Ja</option>
	</select><img src="../ill/blank.gif" width="150" height="1" alt="" border="0">
	<a href="#" onclick="closekomm();" class=vmenu>Gem og Luk&nbsp;</a><br>&nbsp;
	</div>

<br>
<br>
<br><br>
</div>
















<%
end if 'Viskunder Wiz eller Timereg
%>
<!--#include file="../inc/regular/footer_inc.asp"-->





