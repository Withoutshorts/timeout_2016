<%
Response.buffer = True 
%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/functions_inc.asp"-->  
<%
if session("user") = "" then
%>
<!--#include file="../inc/regular/header_inc.asp"-->
<%
	errortype = 5
	call showError(errortype)
	'else
	%>
	<!--include file="inc/isint_func.asp"-->
	<%
	'call erDetInt(request("FM_fakint"))
	'if isInt > 0 AND len(trim(request("FM_fakint"))) > 0 then
	%>
	<!--include file="../inc/regular/header_inc.asp"-->
	<!--include file="../inc/regular/topmenu_inc.asp"-->
	<%
	'errortype = 42
	'call showError(errortype)
	
	'isInt = 0
	else
	
	session("erOprettet") = "n"
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%if request("print") <> "j" then%>			
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!--#include file="../inc/regular/rmenu.asp"-->
	<%end if%>
	<!--#include file="inc/convertDate.asp"-->
	
	<script language="javascript">
	<!--
	function NewWin_popupcal(url)    {
		window.open(url, 'Calpick', 'width=250,height=250,scrollbars=no,toolbar=no');    
	}
	
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	
	 if (document.images){
			plus = new Image(200, 200);
			plus.src = "ill/plus.gif";
			minus = new Image(200, 200);
			minus.src = "ill/minus2.gif";
			}
	
			function expand (de){
				if (document.all("D" + de)){
					if (document.all("D" + de).style.display == "none"){
					document.all("D" + de).style.display = "";
					document.images["plus" + de].src = minus.src;
				}else{
					document.all("D" + de).style.display = "none";
					document.images["plus" + de].src = plus.src;
				}
			}
		}
	
	
	function hidefaknr(){
	document.all["FM_fakint"].value = ""
	}
	
	function hidedatointerval(){
	//alert("hej")
	document.all["FM_usedatointerval"].checked = ""
	}
	//-->
	</script>
	<%
	'** Function til at fjerne decimaler
	Public left_intX
	Public function kommaFunc(x)
	if len(x) <> 0 then
	instr_komma = SQLBless(instr(x, "."))
		
		if instr_komma > 0 then
		left_intX = left(x, instr_komma + 2)
		else
		left_intX = x
		end if
	else
	left_intX = 0
	end if
	
	Response.write left_intX 
	end function
	
	public notfakbeloeb
	public intjobTpris
	public fastpris
	
	public function ikkafaktimerPaaJob()
	'right(year(now()), 2)
	'response.write startDato
	if cdate(startDato) > cdate(lastFakdato) then
	stdatoKri = startDato
	else
	stdatoKri_temp = dateadd("d", 1, lastFakdato)
	stdatoKri = year(stdatoKri_temp)&"/"&month(stdatoKri_temp)&"/"&day(stdatoKri_temp)
	end if
	
	'* Finder ikke fakturerede timer på det pågældende job **
				strSQL3 = "SELECT timer.Tid, timer.timer, timer.tdato, timer.tjobnr, timer.timepris, job.jobnr, job.id, tfaktim, timer.fastpris, jobTpris, budgettimer, timer.TAktivitetId, timer.tidspunkt FROM timer, job WHERE job.id = "& intArrJobnr(a) &" AND timer.tjobnr = job.jobnr AND tfaktim = 1 AND (timer.tdato BETWEEN '"& stdatoKri &"' AND '"& slutdato &"')" 'OR timer.tdato = '" & stdatoKri &"' AND timer.tidspunkt > '"& strfaktidspkt &"'   *****  'BETWEEN '"& lastFakdato &"' AND '"& dagsDatomySQLformat &"' AND tfaktim = 1"
				'timer.tdato >= '"& stdatoKri &"' AND timer.tdato <= '"& slutdato &"'
				'response.write strSQL3
				oRec3.open strSQL3, oConn, 3
				
				while not oRec3.EOF
				diff = DateDiff("d", oRec3("tdato"), stdatoKri)
				if cint(diff) <= 0 then
				'if formatdatetime(oRec3("tdato"), 0) > formatdatetime(lastFakdato, 0) then
							notfaktim = notfaktim + oRec3("timer")
							if oRec3("fastpris") = "1" then
							fastpris = "ya"
							notfakbeloeb = notfakbeloeb + (oRec3("timer") * (oRec3("jobTpris")/oRec3("budgettimer")))
							intjobTpris = oRec3("jobTpris")
							else
							intjobTpris = 0
							fastpris = "no"
							notfakbeloeb = notfakbeloeb + (oRec3("timer") * oRec3("timepris"))
							fakbar = "y"
							end if
				end if
					
					'if cint(diff) = 0 then
					'''if formatdatetime(oRec3("tdato"), 0) = formatdatetime(lastFakdato, 0) then
						 'if formatdatetime(strfaktidspkt, 3) > formatdatetime(oRec3("tidspunkt"), 3) then
							'notfaktim = notfaktim + oRec3("timer")
							'if oRec3("fastpris") = "1" then
							'fastpris = "ya"
							'notfakbeloeb = notfakbeloeb + (oRec3("timer") * (oRec3("jobTpris")/oRec3("budgettimer")))
							'intjobTpris = oRec3("jobTpris")
							'else
							'intjobTpris = 0
							'fastpris = "no"
							'notfakbeloeb = notfakbeloeb + (oRec3("timer") * oRec3("timepris"))
							'fakbar = "y"
							'end if
						 'end if
					'end if
				
				oRec3.movenext
				wend
				oRec3.close
				'****
	end function
	
	function SQLBless(s)
	dim tmp
	tmp = s
	tmp = replace(tmp, ",", ".")
	SQLBless = tmp
	end function
	
	fakbart = 0
	a = 0
	Redim intArrJobnr(a)
	strJobnrPrinted = "0#"
	'linket = request("linket")
	lastJobnr = 0
	fakTotalprJob = 0
	totalnotfakbeloeb = 0
	
	'strYear = request("year")
	'strReqAar = strYear 
	'strReqMrd = request("mrd")
	
	useFaknr = trim(request("FM_fakint"))
	lenuseFaknr = len(trim(request("FM_fakint")))
	barwhere = instr(request("FM_fakint"), "-")
	
	if cint(barwhere) > 0 then
	usebarwhere = (barwhere-1) 
	faknrfirstkri = left(useFaknr, usebarwhere)
	faknrlastkri = right(useFaknr, (lenuseFaknr-barwhere))
	end if
	
	if cint(barwhere) = 0 then
	usefakinterval = "fakturaer.faknr = '"& useFaknr &"'"
	else
	usefakinterval = "fakturaer.faknr >= '"& faknrfirstkri &"' AND fakturaer.faknr <= '"& faknrlastkri &"'"
	end if 
	
	'** Hvis der ikke er valgt et faknummer/interval, bruges dato interval ***
	if len(useFaknr) <> 0 then
	startDato = "2001/1/1"
	slutDato = "2014/1/1"
	usefakinterval = usefakinterval
	usefaknr = "j"
	else
		if cint(request("FM_usedatointerval")) = 1 then
			
		call datofindes(request("FM_start_dag"),request("FM_start_mrd"),request("FM_start_aar"))
		startDato = "20"&right(request("FM_start_aar"),2)&"/"& request("FM_start_mrd") &"/"& dagparset
		
		call datofindes(request("FM_slut_dag"),request("FM_slut_mrd"),request("FM_slut_aar"))
		slutDato = "20"&right(request("FM_slut_aar"),2)&"/"& request("FM_slut_mrd") &"/"& dagparset
		
		else
		startDato = "2001/1/1"
		slutDato = "2014/1/1"
		end if
	usefakinterval = ""
	usefaknr = "n"
	end if
	
	selmedarb = request("selmedarb")
	selaktid = request("selaktid")
	
	FM_job = request("FM_job")
	intJobnr = FM_job 'request("jobnr")
	
	
	'************************************************************
	'**** Valgte jobnr kriterie 							*****
	'************************************************************
	if usefaknr = "j" then
		strSQL = "SELECT id, faknr FROM job, fakturaer WHERE "& usefakinterval &" AND job.id = fakturaer.jobid ORDER BY job.id"
	else
		if intJobnr = "0" OR len(intJobnr) = "0" then
		intJobnrKri = "jobnr = 0"
		else
			Dim intJobnrKriValues 
			Dim i
			intJobnrKriValues = Split(intJobnr, ", ")
			For i = 0 to Ubound(intJobnrKriValues)
			intJobnrKri = intJobnrKri & "jobnr = " & intJobnrKriValues(i) &" OR "
			next
			intJobnrKriCount = len(intJobnrKri)
			trimintJobnrKri = left(intJobnrKri, (intJobnrKriCount-3))
			intJobnrKri = trimintJobnrKri 
		end if
		
		strSQL = "SELECT id FROM job WHERE "& intJobnrKri &" ORDER BY jobnavn"
	end if
	
	'Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"& strSQL
			
			oRec.open strSQL, oConn, 3
			while not oRec.EOF
			
			'*** Opbygning af jobnr til fakturaer *****
			areJobnrP = instr(strJobnrPrinted, ", " &oRec("id")&"#")
			if areJobnrP = "0" then 
			Redim preserve intArrJobnr(a)
			intArrJobnr(a) = oRec("id")
			a = a + 1
			end if
			
			strJobnrPrinted = strJobnrPrinted & ", " & oRec("id") &"#" 
			
			oRec.movenext
			wend
		oRec.close
	' *** '	
		
	
	'*** vis betalte / ikke betalte ***
	intbetalt = request("FM_vis_betalte")
	Select case cint(intbetalt)
	case 0
	kriBetalt = " "
	case 1
	kriBetalt = " AND betalt = 1 "
	case 2
	kriBetalt = " AND betalt <> 1 "
	case 3
	kriBetalt = " "
	end select
	
	
	thisfile = "fak_osigt"
	%>
	<!-------------------------------Sideindhold------------------------------------->
	<%
	if request("print") <> "j" then
	pleft = 190
	ptop = 55
	else
	pleft = 20
	ptop = 10
	end if
	%>	
	<div id="sindhold" style="position:absolute; left:<%=pleft%>; top:<%=ptop%>; visibility:visible; z-index:100;">
	<br><img src="../ill/header_fak.gif" alt="" border="0" width="279" height="34">
	<hr align="left" width="760" size="1" color="#000000" noshade>
	Oversigt over udsendte fakturaer, rykkere og kreditnotaer på de valgte <b>fakturerbare job</b>.<br>
	Der kan indsættes nye fakturer og de eksisterende kan redigeres eller slettes.<br>
	Fakturaer og timer omfatter time-registreringer for <b>samtlige medarbejdere</b>.
	<br><br><br>
<%
	lastFakID = 0
	notfaktim = 0
	notfakbeloeb = 0
	fakTotal = 0
	
	yesterdayDate = date 
	dagsDatomySQLformat = convertDateYMD(yesterdayDate)




'*********************************************************
'** Henter fakturaer på de valgte job 					**
'*********************************************************
foundone = 0
di = 0
for a = 0 to a - 1
f = "no"
divwritten = "n"
antalFak = 0
	
	'*** Vælger hvilket kriterie der skal bruges ***
	if usefaknr = "j" then
	usefakkri = "fakturaer.jobid=" & intArrJobnr(a) &" AND job.id = fakturaer.jobid AND Kid = jobknr AND " & usefakinterval &""
	else
	usefakkri = "fakturaer.jobid=" & intArrJobnr(a) &" AND job.id = fakturaer.jobid AND Kid = jobknr "& kriBetalt &" AND fakturaer.fakdato BETWEEN '"& startDato &"' AND '"& slutDato &"'"
	end if
	
	strSQL = "SELECT fakturaer.Fid, fakturaer.faknr, fakturaer.fakdato, fakturaer.jobid, fakturaer.timer, fakturaer.beloeb, fakturaer.kommentar, fakturaer.faktype, fakturaer.parentfak, jobnavn, jobnr, id, jobknr, fakturerbart, Kkundenavn, Kkundenr, tidspunkt AS faktidspkt, betalt FROM fakturaer, job, kunder WHERE "& usefakkri &" ORDER BY fakDato DESC, tidspunkt DESC"
	
		
		'*** Finder sidste fak uafhængig af valgt datointerval. ***
		strSQL3 = "SELECT fakturaer.fakdato AS lastFakEver FROM fakturaer WHERE fakturaer.jobid=" & intArrJobnr(a) &" ORDER BY fakDato DESC, tidspunkt DESC"
		oRec3.open strSQL3, oConn, 3 
		if not oRec3.EOF then
		lastFakdato = oRec3("lastFakEver")
		end if
		oRec3.close
		
		strSQL3 = "SELECT count(fakturaer.fid) AS antalFak FROM fakturaer WHERE fakturaer.jobid=" & intArrJobnr(a) &" AND fakturaer.fakdato BETWEEN '"& startDato &"' AND '"& slutDato &"'"
		oRec3.open strSQL3, oConn, 3 
		if not oRec3.EOF then
		antalFak = oRec3("antalFak")
		end if
		oRec3.close
		
		if len(lastFakdato) <> 0 then
		lastFakdato = lastFakdato
		else
		lastFakdato = "1/1/2002"
		end if
		
		oRec.open strSQL, oConn, 0, 1
		x = 0
		while not oRec.EOF
		foundone = 1 
		
		'** Skriver Kunde og jobnavn *** 
		if intArrJobnr(a) <> lastJobnr then
		
		
			if x = 0 then
				lastFakdato = convertDateYMD(oRec("fakdato")) 
				strfaktidspkt = formatdatetime(oRec("faktidspkt"), 3)
				x = x + 1
			end if	
			
			call ikkafaktimerPaaJob()
		
		
				'*** Hvis vis alle med registrerede timer siden sidste faktura er valgt *****	
				if (intbetalt = 3 AND SQLBless(notfaktim) > 0) OR intbetalt <> 3 then
				
					'*** Skal div være slået ud?? ***
					'if cint(di) = 0 then
					'if request("FM_usedatointerval") = 1 then
					'disp = ""
					'useplusmnusgif = "minus2"
					'else
					disp = "none"
					useplusmnusgif = "plus"
					'end if
					
					di = 1
					%>
					<table border="0" cellpadding="0" cellspacing="0" width="500">
					<tr>
						<td height=1 bgcolor="#003399" colspan="3"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					</tr>
					<tr bgcolor="#5582D2">
						<td><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
						<td height="30" style="padding-left : 4px; padding-right : 1px;">
						<a href="javascript:expand('<%=oRec("jobnr")%>');"><img src="ill/<%=useplusmnusgif%>.gif" width="9" height="9" border="0" name="plus<%=oRec("jobnr")%>"></a>&nbsp;
						<a href="javascript:expand('<%=oRec("jobnr")%>');" class='alt'>
						<b><%=oRec("jobnr")%>&nbsp;<%=oRec("jobnavn")%></b></a>&nbsp;&nbsp;&nbsp;<font class='lilleblaa'>(<%=oRec("Kkundenavn")%>)</font>&nbsp;&nbsp;&nbsp;
						<b><%=antalFak%></b> faktura(er)</td>
						<td align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
					</tr>
					<tr>
						<td height=1 bgcolor="#003399" colspan="3"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					</tr>
					</table>
					<%
					divwritten = "y"
					%>
					<DIV ID="D<%=oRec("jobnr")%>" Style="position: relative; display: <%=disp%>;">
					<table cellspacing="0" cellpadding="0" border="0" width="500" bgcolor="#ffffff">
					<tr bgcolor="#66CC33">
						<td width=10><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
						<td width="200" colspan=2 valign="bottom" style="padding-left:5px; " class='alt'><b>Job</b></td>
						<td width="70"  valign="bottom" align="right" style="padding-right:3px;" class='alt'><b>Fak. dato</b></td>
						<td width="70" valign="bottom" align="right" style="padding-right:3px;" class='alt'><b>Fakturanr.</b></td>
						<td width="70" valign="bottom" class='alt' style="padding-left:10px;"><b>Type</b>&nbsp;(Opr.)</td>
						<td width="70" valign="bottom" align="right" style="padding-right:10px;" class='alt'><b>Beløb</b></td>
						<td width=10 align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
					</tr>
					<%
					recJobnr_ya = oRec("jobnr")
					fakbart = oRec("fakturerbart") 'Er jobbet fakturerbart?
					f = "ya" 'Er der udskrevet fakturaer på jobbet?
				end if
		end if%>
		
		
		<%
		'************************************************************************************
		'*** Skal fakturaerne på det valgte job vises? lever de op til filter kriterier?? ****
		'*** Hvis vis alle med registrerede timer siden sidste faktura er valgt. 		*****
		'************************************************************************************
		
		if (intbetalt = 3 AND SQLBless(notfaktim) > 0) OR intbetalt <> 3 then
		
			'*** Skriver Fakturaer på det pågældende job ***
			if oRec("Fid") <> lastFakID then
		
				%>
				<tr bgcolor="#003399">
					<td height=1 colspan="8"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
				</tr>
				<%
				Select case oRec("faktype")
					case 0
					bgcol = "#ffffff"
					case 1
					bgcol = "#eff3ff"
					case 2
					bgcol = "#d6dff5"
					end select
				%>
				<tr bgcolor="<%=bgcol%>">
					<td rowspan="2"><img src="../ill/tabel_top.gif" width="1" height="40" alt="" border="0"></td>
					<td colspan=2 valign=top style="padding-left:5;">
					<%if request("print") <> "j" then
					
					'*** Rediger via gammel fak fil. (gamle fakturaer!)
					if cdate(oRec("fakdato")) < cdate("18/10/2004") then
					fakfil = "fak_org.asp"
					else
					fakfil = "fak.asp"
					end if
					
					%>		
						<a href="<%=fakfil%>?menu=stat_fak&func=red&jobid=<%=oRec("jobid")%>&id=<%=oRec("Fid")%>&jobnr=<%=oRec("jobnr")%>&mrd=<%=strReqMrd%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>&faktype=<%=oRec("faktype")%>&FM_usedatointerval=<%=request("FM_usedatointerval")%>&FM_start_dag=<%=request("FM_start_dag")%>&FM_start_mrd=<%=request("FM_start_mrd")%>&FM_start_aar=<%=request("FM_start_aar")%>&FM_slut_dag=<%=request("FM_slut_dag")%>&FM_slut_mrd=<%=request("FM_slut_mrd")%>&FM_slut_aar=<%=request("FM_slut_aar")%>&FM_fakint=<%=trim(request("FM_fakint"))%>" class=vmenu>
						(<%=oRec("jobnr")%>)&nbsp;
						<%=oRec("jobnavn")%></a>
					<%else%>
					<%=oRec("jobnavn")%>
					<%end if%>&nbsp;&nbsp;</td>
					<td valign=top align=right style="padding-right:3;"><b><%=formatdatetime(oRec("fakdato"), 0)%></b>&nbsp;</td>
					<td valign=top align=right style="padding-right:3;" class=blaa>&nbsp;&nbsp;<%=oRec("faknr")%></td>
					<td valign=top style="padding-right:3; padding-left:10;"><%
					if oRec("faktype") <> 0 then
							strSQL3 = "SELECT faknr FROM fakturaer WHERE Fid = " & oRec("parentfak")
							oRec3.open strSQL3, oConn, 3 
							if not oRec3.EOF then
							oprFaknr = oRec3("faknr")
							end if
							oRec3.close
					end if
					
					Select case oRec("faktype")
					case 0
					strFaktypeShow = "<b>Faktura</b>"
					case 1
					strFaktypeShow = "<b>Kredit</b>&nbsp;("&oprFaknr&")"
					case 2
					strFaktypeShow = "<b>Rykker</b>&nbsp;("&oprFaknr&")"
					end select
					Response.write strFaktypeShow 
					%>
					</td>
					<td align="right" valign=top style="padding-right:3; padding-left:3;"><%=formatnumber(oRec("beloeb"), 2)%> dkr</td>
					<td align="right" rowspan=2><img src="../ill/tabel_top.gif" width="1" height="40" alt="" border="0"></td>
				</tr>
				<tr bgcolor="<%=bgcol%>">
					<td colspan="6" height="10" valign="top" style="padding-right:3; padding-left:20;">
					<%
					strSQL2 = "SELECT Fid, parentfak FROM fakturaer WHERE parentfak = " &oRec("Fid")&" AND faktype = 2"
					oRec2.open strSQL2, oConn, 3 
					hasfak = 0
					if not oRec2.EOF then
					hasfak = 1
					end if
					oRec2.close
					
					if oRec("betalt") <> 0 then
					
						if hasfak = 0 AND oRec("faktype") <> 1 then%>
						Opret <a href="<%=fakfil%>?menu=stat_fak&func=red&jobid=<%=oRec("jobid")%>&id=<%=oRec("Fid")%>&jobnr=<%=oRec("jobnr")%>&mrd=<%=strReqMrd%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>&faktype=2&faknr=<%=oRec("faknr")%>&rykkreopr=j&FM_usedatointerval=<%=request("FM_usedatointerval")%>&FM_start_dag=<%=request("FM_start_dag")%>&FM_start_mrd=<%=request("FM_start_mrd")%>&FM_start_aar=<%=request("FM_start_aar")%>&FM_slut_dag=<%=request("FM_slut_dag")%>&FM_slut_mrd=<%=request("FM_slut_mrd")%>&FM_slut_aar=<%=request("FM_slut_aar")%>&FM_fakint=<%=trim(request("FM_fakint"))%>" class="kalblue">Rykker</a>&nbsp;&nbsp;&nbsp;&nbsp;
						Opret <a href="<%=fakfil%>?menu=stat_fak&func=red&jobid=<%=oRec("jobid")%>&id=<%=oRec("Fid")%>&jobnr=<%=oRec("jobnr")%>&mrd=<%=strReqMrd%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>&faktype=1&faknr=<%=oRec("faknr")%>&rykkreopr=j&FM_usedatointerval=<%=request("FM_usedatointerval")%>&FM_start_dag=<%=request("FM_start_dag")%>&FM_start_mrd=<%=request("FM_start_mrd")%>&FM_start_aar=<%=request("FM_start_aar")%>&FM_slut_dag=<%=request("FM_slut_dag")%>&FM_slut_mrd=<%=request("FM_slut_mrd")%>&FM_slut_aar=<%=request("FM_slut_aar")%>&FM_fakint=<%=trim(request("FM_fakint"))%>" class="kalblue">Kreditnota</a>
						<%else%>
						&nbsp;&nbsp;
						<%end if
						
					else
					Response.write "Status: <font class=lillegray>Kladde</font>&nbsp;&nbsp;"
					%>
					<a href="fak.asp?menu=stat_fak&func=red&jobid=<%=oRec("jobid")%>&id=<%=oRec("Fid")%>&jobnr=<%=oRec("jobnr")%>&mrd=<%=strReqMrd%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>&faktype=<%=oRec("faktype")%>&FM_usedatointerval=<%=request("FM_usedatointerval")%>&FM_start_dag=<%=request("FM_start_dag")%>&FM_start_mrd=<%=request("FM_start_mrd")%>&FM_start_aar=<%=request("FM_start_aar")%>&FM_slut_dag=<%=request("FM_slut_dag")%>&FM_slut_mrd=<%=request("FM_slut_mrd")%>&FM_slut_aar=<%=request("FM_slut_aar")%>&FM_fakint=<%=trim(request("FM_fakint"))%>" class=fir_gr>Gennemse og godkend</a>
					&nbsp;|&nbsp;<a href="<%=fakfil%>?menu=stat_fak&func=slet&id=<%=oRec("Fid")%>&FM_job=<%=oRec("jobnr")%>&FM_usedatointerval=<%=request("FM_usedatointerval")%>&FM_start_dag=<%=request("FM_start_dag")%>&FM_start_mrd=<%=request("FM_start_mrd")%>&FM_start_aar=<%=request("FM_start_aar")%>&FM_slut_dag=<%=request("FM_slut_dag")%>&FM_slut_mrd=<%=request("FM_slut_mrd")%>&FM_slut_aar=<%=request("FM_slut_aar")%>&FM_fakint=<%=trim(request("FM_fakint"))%>" class=fir_r>Slet denne kladde</a>
					<%end if%>
					</td>
				</tr>
				<%
				Select case oRec("faktype") 
				case 0
					if hasfak = 0 then
					fakTotalprJob = fakTotalprJob + oRec("beloeb")
					else
					fakTotalprJob = fakTotalprJob
					end if
				case 1
				fakTotalprJob = fakTotalprJob - oRec("beloeb")
				case 2
				fakTotalprJob = fakTotalprJob + oRec("beloeb")
				end select
				
				
				lastJobnr = intArrJobnr(a)
				lastFakID = oRec("Fid")
			end if
		end if
		oRec.movenext
		wend
		oRec.close
		%>
		<!--
		<tr bgcolor="#003399">
			<td height=1 colspan="8"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
		</tr>-->
		<%
		
		'fakTotal = fakTotal + fakTotalprJob 
		
		
		
		
		
		
		'***************************************************************************************
		'**** Udkskriver total omsætning på job og viser knap til oprettelse af ny faktura. ****
		'***************************************************************************************
		
		if (intbetalt = 3 AND SQLBless(notfaktim) > 0) OR intbetalt <> 3 then
		
				if fakbart = "1" then
				%>
				<tr bgcolor="#EFF3FF">
					<td><img src="../ill/tabel_top.gif" width="1" height="35" alt="" border="0"></td>
					<td colspan="2"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
					<td colspan="4" valign="bottom" colspan="2" align="right" style="padding-right : 5px;"><u>Balance:&nbsp;&nbsp;<b><%=formatnumber(fakTotalprJob,2)%> dkr.</b></u></td>
					<td align="right"><img src="../ill/tabel_top.gif" width="1" height="35" alt="" border="0"></td>
				</tr>
				<tr bgcolor="#EFF3FF">
					<td><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
					<td colspan="2" valign="bottom">&nbsp;</td>
					<td valign="bottom" align="right"><font size=1>timer
					<%if fastpris = "ya" then%>
					<font color=slategray> (stk.) </font>
					<%end if%>
					</font></td>
					<td valign="bottom" align="right" colspan="3"><font size=1>omsætning 
					<%if fastpris = "ya" then%>
					/ fastpris
					<%end if%></font>&nbsp;</td>
					<td align="right"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
				</tr>
				<tr bgcolor="#EFF3FF">
					<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
					<td align=right colspan="2" valign="bottom"><font size=1>Til fakturering:<br> (kun  fakturerbare)&nbsp;&nbsp;</td>
					<td valign="bottom" align="right" style="!border: 1 px; background-color: #FFFFFF; border-color: Gray; border-style: solid; border-right:0; padding-left : 4px; padding-right : 1px;"><%=SQLBless(notfaktim)%></td>
					<td valign="bottom" colspan="3" align="right" style="!border: 1 px; background-color: #FFFFFF; border-color: Gray; border-style: solid; padding-left : 4px; padding-right : 3px;">
					<%if fastpris = "ya" then
					Response.write formatnumber(notfakbeloeb, 2) &"&nbsp;/&nbsp;"
					Response.write formatnumber(intjobTpris, 2) 
					else 
					Response.write formatnumber(notfakbeloeb, 2)
					end if
					totalnotfakbeloeb = totalnotfakbeloeb + notfakbeloeb
					%> dkr</td>
					<td align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
				</tr>
				<tr bgcolor="#EFF3FF">
					<td><img src="../ill/tabel_top.gif" width="1" height="45" alt="" border="0"></td>
					<td colspan="6" height="30" align="right"><font size=1 color="#708090">
						<%if request("print") <> "j" then%>
						<br><a href="fak.asp?menu=stat_fak&ttf=<%=notfaktim%>&ktf=<%=kommaFunc(notfakbeloeb)%>&jobid=<%=intArrJobnr(a)%>&jobnr=<%=recJobnr_ya%>&mrd=<%=strReqMrd%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>&faktype=0&FM_usedatointerval=<%=request("FM_usedatointerval")%>&FM_start_dag=<%=request("FM_start_dag")%>&FM_start_mrd=<%=request("FM_start_mrd")%>&FM_start_aar=<%=request("FM_start_aar")%>&FM_slut_dag=<%=request("FM_slut_dag")%>&FM_slut_mrd=<%=request("FM_slut_mrd")%>&FM_slut_aar=<%=request("FM_slut_aar")%>&FM_fakint=<%=trim(request("FM_fakint"))%>">&nbsp;<img src="../ill/nyfak.gif" width="114" height="28" alt="" border="0"></a>&nbsp;&nbsp;
						<%end if%>
					</td>
					<td align="right"><img src="../ill/tabel_top.gif" width="1" height="45" alt="" border="0"></td>
				</tr>
				<tr><td colspan="8" height="10" bgcolor="#5582D2"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td></tr>
				<%end if
				
				notfaktim = 0
				notfakbeloeb = 0
		end if
		
		
		'**********************************************************************************
		'** Hvis der er ikke oprettet fakturaer på det pågældende job.					  *
		'** Og der er valgt en kørsel hvor der ikke er brugt en af følgende kritertier:   *
		'** 1) Ikke faktura interval 													  *
		'** 2) Ikke dato interval 														  *
		'** 3) Status skal være = 'vis alle'											  *
		'**********************************************************************************
		
		'AND cint(request("FM_usedatointerval")) <> 1 
		if f = "no" AND (len(useFaknr) <> 0 AND cint(intbetalt) = 0) then
			strSQL2 = "SELECT id, jobnavn, jobnr, fakturerbart, Kkundenavn FROM job, kunder WHERE id =" & intArrJobnr(a) &" AND Kid = jobknr" 
			oRec2.open strSQL2, oConn, 0, 1
				if not oRec2.EOF then
					if oRec2("fakturerbart") = "1" then
					%>
					<table border="0" cellpadding="0" cellspacing="0" width="500">
					<tr>
						<td height=1 bgcolor="#003399" colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					</tr>
					<tr bgcolor="#5582D2">
						<td><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
						<td colspan="4" height="30" style="padding-left : 4px; padding-right : 1px;">
						<a href="javascript:expand('<%=oRec2("jobnr")%>');"><img src="ill/plus.gif" width="9" height="9" border="0" name="plus<%=oRec2("jobnr")%>"></a>&nbsp;
						<a href="javascript:expand('<%=oRec2("jobnr")%>');" class='alt'>
						<b><%=oRec2("jobnr")%>&nbsp;<%=oRec2("jobnavn")%></b></a>&nbsp;&nbsp;&nbsp;<font class='lilleblaa'>(<%=oRec2("Kkundenavn")%>)</font>&nbsp;&nbsp;&nbsp;
						<b><%=antalFak%></b> faktura(er)</td>
						<td align="right"><img src="../ill/tabel_top.gif" width="1" height="30" alt="" border="0"></td>
					</tr>
					<tr>
						<td height=1 bgcolor="#003399" colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					</tr>
					</table>
					<%
					divwritten = "y"
					foundone = 1
					%>
					<DIV ID="D<%=oRec2("jobnr")%>" Style="position: relative; display: none;">
					<table cellspacing="0" cellpadding="0" border="0" width="500" bgcolor="#ffffff">
					<%
					recJobnr_no = oRec2("jobnr")
					fakbart = 1
					end if
				end if
			strfaktidspkt = "00:00:01 AM"
			oRec2.close
			
			call ikkafaktimerPaaJob()
			
			if fakbart = "1" then%>
			<tr bgcolor="#EFF3FF">
				<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="45" alt="" border="0"></td>
				<td colspan="2">&nbsp;Der er ikke oprettet fakturaer på dette job i den angivne periode.</td>
				<td valign="bottom" align="right"><font size=1>timer
				<%if fastpris = "ya" then%>
				<font color=slategray> (stk.) </font>
				<%end if%></font></td>
				<td valign="bottom" align="right"><font size=1>omsætning
				<%if fastpris = "ya" then%>
				/ fastpris
				<%end if%></font></td>
				<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="45" alt="" border="0"></td>
			</tr>
			<tr bgcolor="#EFF3FF">
				<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
				<td colspan="2" valign="bottom" align="right">Endnu ikke fakturede, fakturerbare timer:&nbsp;&nbsp;</td>
				<td valign="bottom" align="right" style="!border: 1 px; background-color: #FFFFFF; border-color: Gray; border-style: solid; border-right:0; padding-left : 4px; padding-right : 1px;"><%=SQLBless(notfaktim)%></td>
				<td valign="bottom" align="right" style="!border: 1 px; background-color: #FFFFFF; border-color: Gray; border-style: solid; padding-left : 4px; padding-right : 1px;">
				<%
				if fastpris = "ya" then
				Response.write formatnumber(notfakbeloeb,2) &"&nbsp;/&nbsp;"
				Response.write formatnumber(intjobTpris,2) 
				else 
				Response.write formatnumber(notfakbeloeb,2) 
				end if
				totalnotfakbeloeb = totalnotfakbeloeb + notfakbeloeb
				%> dkr</td>
				<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
			</tr>
			<tr bgcolor="#EFF3FF">
				<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="45" alt="" border="0"></td>
				<td colspan="4" height="30" align="right">
				<%if request("print") <> "j" then%>		
				<br><a href="fak.asp?menu=stat_fak&ttf=<%=notfaktim%>&ktf=<%=kommaFunc(notfakbeloeb)%>&jobid=<%=intArrJobnr(a)%>&jobnr=<%=recJobnr_no%>&mrd=<%=strReqMrd%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>&FM_usedatointerval=<%=request("FM_usedatointerval")%>&FM_start_dag=<%=request("FM_start_dag")%>&FM_start_mrd=<%=request("FM_start_mrd")%>&FM_start_aar=<%=request("FM_start_aar")%>&FM_slut_dag=<%=request("FM_slut_dag")%>&FM_slut_mrd=<%=request("FM_slut_mrd")%>&FM_slut_aar=<%=request("FM_slut_aar")%>&FM_fakint=<%=trim(request("FM_fakint"))%>">&nbsp;<img src="../ill/nyfak.gif" width="114" height="28" alt="" border="0"></a>&nbsp;&nbsp;
				<%end if%></td>
				<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="45" alt="" border="0"></td>
			</tr>
			<tr><td colspan="6" height="10" bgcolor="#5582D2"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td></tr>
			<%end if
			
			notfaktim = 0
			notfakbeloeb = 0
			lastJobnr = intArrJobnr(a)
		
		end if

			fakbart = 0
			fakTotalprJob = 0
			
			'** afsluttter inter tabel og div hvis de er påbegyndt ***
			if divwritten = "y" then ' OR cint(intbetalt) = 0 then%>
			</table>
			</div><img src="../ill/blank.gif" width="1" height="4" alt="" border="0"><br>
			<%
			end if

next
%><br><br><br>&nbsp;
</div>


<%if cint(foundone) = 0 then%>
<div style="position:absolute; left:250; top:230; z-index:10000; background-color:#ffffe1; border:1px darkred solid; padding:5;">
<table cellspacing="0" cellpadding="0" border="0" width=330>
<tr><td valign="top"><img src="../ill/alert.gif" width="44" height="45" alt="" border="0"></td>
<td style="padding-left:5;"><br>Der er ikke fundet nogen <b>Fakturaer</b> eller<b> job</b> der matcher de valgte kriterier.<br>
<br>Dette skyldes en af to ting:<br>
<b>1)</b>&nbsp;Der er ikke valgt nogen job.<br>
<b>2)</b>&nbsp;Der er angivet et faktura nummer der ikke findes.<br><br>&nbsp;</td></tr>
</table>
</div>
<%end if%>




<div style="position:absolute; left:720; top:130; z-index:10000;">
<table cellspacing="0" cellpadding="0" border="0" width=230>
<form action="fak_osigt.asp?menu=stat_fak" method="post" name="statselector" id="statselector">
<input type="hidden" name="FM_job" value="<%=FM_job%>">
	<tr bgcolor="#5582D2">
			<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
			<td colspan=2 valign="top"><img src="../ill/tabel_top.gif" width="218" height="1" alt="" border="0"></td>
			<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
			<td colspan=2 valign="top" class="alt">Kør ny oversigt</td>
	</tr>
	<tr bgcolor="#EFF3FF">
	<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="212" alt="" border="0"></td>
	<td colspan="2" valign="top">
	<%
	if cint(request("FM_usedatointerval")) = 1 AND len(trim(request("FM_fakint"))) = 0 then
	dvalch = "checked"
	else
	dvalch = ""
	end if
	%>
	<b>A)</b> <input type="checkbox" name="FM_usedatointerval" value="1" <%=dvalch%> onClick="hidefaknr()">
	Brug dato interval.<br>
	<!--#include file="inc/weekselector_s.asp"--><br><br>
	Brug dato interval, hvis du ønsker at finde fakturaer eller nye fakturerbare timer inden for et bestemt datointerval.<br>
	<%
	if len(trim(request("FM_fakint"))) = 0 then
		Select case cint(intbetalt)
		case 0
		stch0 = "checked"
		stch1 = ""
		stch2 = ""
		stch3 = ""
		case 1
		stch0 = ""
		stch1 = "checked"
		stch2 = ""
		stch3 = ""
		case 2
		stch0 = ""
		stch1 = ""
		stch2 = "checked"
		stch3 = ""
		case 3
		stch0 = ""
		stch1 = ""
		stch2 = ""
		stch3 = "checked"
		end select
	else
	stch0 = "checked"
	stch1 = ""
	stch2 = ""
	stch3 = ""
	end if
	%>
	<!--Marker "vis job" filter kriterier her. Vis:<br>
	<input type="radio" name="FM_vis_betalte" value="0" checked>&nbsp;Alle.<br>
	<input type="radio" name="FM_vis_betalte" value="1">&nbsp;Kun betalte.&nbsp;<br>
	<input type="radio" name="FM_vis_betalte" value="2">&nbsp;Kun <b>ikke</b> betalte.<br>
	<input type="radio" name="FM_vis_betalte" value="3">&nbsp;Kun for job med registreringer i det valgte dato interval.<br>
	<br>-->
	<input type="hidden" name="FM_vis_betalte" value="0">
	<br>
	<b>B)</b> Eller vælg et <b>faktura nr</b> / <b>interval</b> her, hvis du ønsker at finde bestemte fakturaer.<br>
	<font size=1 color=#999999>(ex. 1045 ell. 1045-1067)</font><br>
	
	<input type="text" name="FM_fakint" size="20" value="<%=trim(request("FM_fakint"))%>" onFocus="hidedatointerval()"><br>&nbsp;</td>
	<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="212" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="55" alt="" border="0"></td>
		<td colspan="2" align=center><br>
		<input type="image" src="../ill/statpil.gif" alt="Kør Statistik!"></td>
		<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="55" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan="2" valign="bottom"><img src="../ill/tabel_top.gif" width="218" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#d6dff5">
		<td valign="top"><img src="../ill/blank.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan="2" valign="bottom"><br>
		<font class=megetlillesort><b>Note(r):</b><br> Hvis der benyttes et datointerval hvor den valgte startdato ligger før datoen for den sidst oprettede faktura på det pågældende job, bruges faktura datoen som startdato når nye fakturerbare timer skal findes. <br><br>
		Begge de valgte dage er inklusive.<br><br>
		Faktura datoen oprettes således, at alle timer til og med faktura datoen er indbefattet. Når en faktura oprettes fjernes muligheden for at indtaste timer, frem til og med faktura datoen.
		</td>
		<td valign="top" align="right"><img src="../ill/blank.gif" width="8" height="10" alt="" border="0"></td>
	</tr>
	</table>
</div>

<!--
<div style="position:absolute; left:700; top:570;">
	<table celpadding=0 celspacing=0 width=250 border=0>
	<tr>
		<td colspan="5" align="right"><br>Samlet beløb  på de valgte job:</td>
		<td><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr>
		<td colspan="4" align="right">Udfaktureret beløb:</td>
		<td valign="bottom" align="right" style="!border: 2 px; background-color: #FFFFFF; border-color: LimeGreen; border-style: solid; padding-left : 4px; padding-right : 1px;"><b><=formatnumber(fakTotal, 2)%> dkr</b></td>
		<td><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<
	if len(intRestance) = 0 then
	intRestance = 0
	else
	intRestance = intRestance
	end if
	%>
	<tr>
		<td colspan="4" align="right">Restance:</td>
		<td valign="bottom" align="right" style="!border: 2 px; background-color: #FFFFFF; border-color: Green; border-style: solid; padding-left : 4px; padding-right : 1px;"><b><=formatnumber(intRestance,2)%> dkr</b></td>
		<td><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr>
		<td colspan="4" align="right">Beløb til fakturering:</td>
		<td valign="bottom" align="right" style="!border: 2 px; background-color: #FFFFFF; border-color: darkRed; border-style: solid; padding-left : 4px; padding-right : 1px;"><b><%
		'totalnotfakbeloeb = formatnumber(totalnotfakbeloeb, 2)
		'Response.write formatnumber(totalnotfakbeloeb, 2)%> dkr</b></td>
		<td><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr><td colspan="6" height="20" style="padding-left:15px;">
	<br><br>
	nb) Kun oprindelige fakturaer er med i de sumerede omsætningstal. Rykker skrivelser er ikke indregnet.<br>Kreditnotaer er fratrukket.<br><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"><br><br>
	<br><br>&nbsp;</td></tr>
	<%
	'fakTotal = 0
	'Set oRec = Nothing
	%>	
	</table>
</div>
-->

<%'end if '*** validering
end if %>
<!--#include file="../inc/regular/footer_inc.asp"-->
