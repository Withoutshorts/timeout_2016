<%response.buffer = true%>
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
	%>
	
	<%
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	'*** Hvis plus eller minus knap er brugt ***
	if len(request("selectplus")) <> 0 then
	selectplus = request("selectplus")
	else
	selectplus = 0
	end if
	
	if len(request("selectminus")) <> 0 then
	selectminus = request("selectminus")
	else
	selectminus = 0
	end if
	
	
	if len(request("mdselect")) <> 0 then
		monththis = request("mdselect")
		
		if selectminus <> 0 then
		monththis = monththis - 1
		end if
		
		if selectplus <> 0 then
		monththis = monththis + 1
		end if
		
		
		select case monththis 
		case 0
		monththis = 12
		yearthis = request("aarselect")
		yearthis = yearthis - 1
		case 13
		monththis = 1
		yearthis = request("aarselect")
		yearthis = yearthis + 1
		case else
		yearthis = request("aarselect")
		end select
		
	else
	monththis = month(now)
	yearthis = year(now)
	end if
	
	
			'** Antal dage i md ***
			Select case monththis
			case "4", "6", "9", "11"
				mthDays = 30
			case 2
					select case yearthis
					case 2004, 2008, 2012, 2016, 2020, 2024, 2028
					mthDays = 29
					case else 
					mthDays = 28
					end select
			case else
				mthDays = 31
			end select
			
			
	
	
	if len(request("jobselect")) <> 0 AND request("jobselect") <> "Jobnavn.. (Alle)" then
	jobnavn = request("jobselect")
	jobNavnKri = " AND jobnavn LIKE '"& jobnavn &"%' "
	else
	jobnavn = "Jobnavn.. (Alle)"
	jobNavnKri = " AND jobnavn <> '' "
	end if
	
	if len(request("visalle")) <> 0 then
	visalle = 1
	else
	visalle = 0
	end if
	%>
	<script language=javascript>
	function clearJnavn() {
	document.getElementById("jobselect").value = ""
	}
	
	function seladd(){
	document.getElementById("selectplus").value = 1 
	}
	
	function selminus(){
	document.getElementById("selectminus").value = 1 
	}
	
	function show(jobid) {
	lastdivval = document.getElementById("FM_lastdiv").value 
	document.getElementById("d_"+lastdivval+"").style.visibility = "hidden"
	document.getElementById("d_"+lastdivval+"").style.display = "none"
	
	document.getElementById("d_"+jobid+"").style.visibility = "visible"
	document.getElementById("d_"+jobid+"").style.display = ""
	
	document.getElementById("FM_lastdiv").value = jobid 
	}
	
	function hide(jobid) {
	document.getElementById("d_"+jobid+"").style.visibility = "hidden"
	document.getElementById("d_"+jobid+"").style.display = "none"
	}
	
	//function opretny(){
	//document.getElementById("opretny").style.visibility = "visible"
	//document.getElementById("opretny").style.display = ""
	//}
	
	</script>
		<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
		<!--#include file="../inc/regular/topmenu_hvd_inc.asp"-->
	
		
	
	<%leftPos = 20%>
	<%topPos = 50%>
	<!-------------------------------Sideindhold------------------------------------->
	
	<!-- Oversigts kalender -->
	<div id="overblik" style="position:absolute; left:<%=leftPos%>; top:<%=topPos%>; visibility:visible;">
	<a href="pg_allokering.asp?menu=job" class=vmenu target=_top><img src="../ill/soeg-knap_tilbage.gif" width="16" height="16" alt="" border="0"> Tilbage til joboversigten.</a>
	<table cellspacing="0" cellpadding="0" border="0">
	<form action="jbpla_k.asp" method="post" name="mdselect" id="mdselect">
	<tr>
	<td colspan=2 width=350>
	<%
	if visalle = 1 then
	chkvalle = "CHECKED"
	else
	chkvalle = ""
	end if 
	%>
	<input type="checkbox" name="visalle" id="visalle" value="1" <%=chkvalle%>>
	<font class="megetlillesort">Vis alle job, uanset uanset om jobperioden ligger uden for det valgte interval.
	</td>
	</tr>
	<tr>
	<td>
	<input type="text" name="jobselect" id="jobselect" size="15" value="<%=jobnavn%>" onFocus="clearJnavn()">
	&nbsp;<select name="mdselect">
	<option value="<%=monththis%>" SELECTED><%=left(monthname(monththis),3)%></option>
	<option value="1">Jan</option>
	<option value="2">Feb</option>
	<option value="3">Mar</option>
	<option value="4">Apr</option>
	<option value="5">Maj</option>
	<option value="6">Jun</option>
	<option value="7">Jul</option>
	<option value="8">Aug</option>
	<option value="9">Sep</option>
	<option value="10">Okt</option>
	<option value="11">Nov</option>
	<option value="12">Dec</option>
	</select><select name="aarselect">
	<option value="<%=yearthis%>" SELECTED><%=yearthis%></option>
	<option value="2002">2002</option>
	<option value="2003">2003</option>
	<option value="2004">2004</option>
	<option value="2005">2005</option>
	<option value="2005">2006</option>
	<option value="2005">2007</option>
	<option value="2005">2008</option>
	</select></td><td valign="bottom"><input type="image" src="../ill/pilstorxp.gif"></td></tr>
	
	<input type="hidden" name="selectminus" id="selectminus" value="0">
	<input type="hidden" name="selectplus" id="selectplus" value="0">
	
	<div style="position:absolute; left:0; top:65;">
	<input type="submit" name="submitplus" id="submitplus" value="-" style="font-size:9px; border:1px #66CC33 solid; background-color:#ffffff; width:15;" onClick="selminus()">
	</div>
	<div style="position:absolute; left:560; top:65;">
	<input type="submit" name="submitplus" id="submitplus" value="+" style="font-size:9px; border:1px #66CC33 solid; background-color:#ffffff; width:15;" onClick="seladd()">
	</div>
	</form>
	</table>
	<br>
	<a href="#" class=vmenu>Opret nyt job</a><br>
	</div>
	
	<div id="job" style="position:absolute; left:<%=leftPos%>; top:<%=topPos+85%>; visibility:visible; overflow:auto; height:350; width:575;">
	
	
	
	<%
	'*************************************************************************************
	'Joboversigten
	'Viser alle job det matcher filter i den valgte måned.
	'*************************************************************************************
	
	mthDaysUse = mthDays
	monthUse = monththis
	yearUse = yearthis
	
	jobStartKri = yearUse&"/"&monthUse&"/"&"1"
	jobSlutKri = yearUse&"/"&monthUse&"/"&mthDaysUse
	
	daynow = day(now)&"/"&month(now)&"/"&year(now)
	%>
	
	<table cellspacing="0" cellpadding="0" border="0">
	<tr bgcolor="#d6dff5"><td bgcolor="#8caae6" colspan=2 align=right valign="top" style="padding-right:2; padding-left:2;"><font class="megetlillesort"><%=left(monthname(monthUse), 3)%>&nbsp;<%=yearUse%></td>
		<td valign="top">
				<table cellspacing="0" cellpadding="0" border="0">
				<tr>
							<%for d = 0 to mthDaysUse - 1
							getnyDato = dateadd("d", d, jobStartKri)
							showday = d + 1
							
							if cdate(daynow) = cdate(getnyDato) then
							bgcolordatoer = "#FFCC00"
							else
							
								select case weekday(getnyDato)
								case 1,7
								bgcolordatoer = "#cccccc"
								case else
								bgcolordatoer = "#8caae6"
								end select
							
							end if
							
							%>
							<td bgcolor="<%=bgcolordatoer%>" valign="top" align=center width=14 style="border-left:1px #ffffff solid;"><font class="megetlillesort"><%=showday%></td>
							<%next%>
				</tr>
				</table>
			
		</td></tr>
	
	<%
	
	
	'** vis alle uanset datoKri
	if visalle = 1 then
	datoKri = ""
	else
	datoKri = "AND jobstartdato <= '"&jobSlutKri&"' AND jobslutdato >= '"&jobStartKri&"'"
	end if
	
	x = 0
	dim arrjobid, arrjobnavn, arrjobnr, arrjstdato, arrslutdato
	Redim arrjobid(x)
	Redim arrjobnavn(x)
	Redim arrjobnr(x)
	Redim arrjstdato(x)
	Redim arrslutdato(x)
		
	strSQL = "SELECT job.id AS jobid, jobnavn, jobnr, jobstartdato, jobslutdato, jobstatus FROM job WHERE fakturerbart = 1 AND jobstatus = 1 "& datoKri &" "& jobNavnKri &" ORDER BY jobnavn"
	'Response.write strSQL
	'Response.flush
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
	Redim preserve arrjobid(x)
	Redim preserve arrjobnavn(x)
	Redim preserve arrjobnr(x)
	Redim preserve arrjstdato(x)
	Redim preserve arrslutdato(x)
	
	arrjobid(x) = oRec("jobid")
	arrjobnavn(x) = oRec("jobnavn")
	arrjobnr(x) = oRec("jobnr")
	arrjstdato(x) = oRec("jobstartdato")
	arrslutdato(x) = oRec("jobslutdato")
		%>
		<tr>
			
			<td width="130" bgcolor="#eff3ff" style="padding-right:2; padding-left:2; border-bottom:1px #ffffff solid;"><font class="megetlillesort"><%=left(oRec("jobnavn"), 20)%> (<%=oRec("jobnr")%>)&nbsp;</td>
			<td width="30" bgcolor="#ffffff" style="border-bottom:1px #ffffff solid;"><a href="javascript:NewWin_large('jobplanner.asp?menu=job&id=<%=oRec("jobid")%>')"><img src="../ill/popupcal_small.gif" width="11" height="10" alt="Rediger jobperiode og aktiviteter" border="0"></a>
			<a href="javascript:NewWin_large('ressourcer.asp?menu=job&id=<%=oRec("jobid")%>')" class=vmenu><b>R</b></a></td>
	
		<td height=8 valign="top">
				<table cellspacing="0" cellpadding="0" border="0">
				<tr>
						<%for d = 0 to mthDaysUse - 1
							
							getnyDato = dateadd("d", d, jobStartKri)
							ressourcetimerDato = 0
							
							if cdate(oRec("jobstartdato")) <= cdate(getnyDato) AND cdate(oRec("jobslutdato")) >= cdate(getnyDato) then
							imgthis = "g-prik.gif"
									
									sqlday = d + 1
									sqlresdato = yearthis&"/"&monththis&"/"&sqlday 
									
									strSQL3 = "SELECT sum(timer) AS restimer FROM ressourcer WHERE ressourcer.jobid = "& oRec("jobid") &" AND dato = '"&sqlresdato&"'"
									'Response.write strSQL3
									oRec3.open strSQL3, oConn, 3 
									if not oRec3.EOF then
									ressourcetimerDato = oRec3("restimer")
									end if
									oRec3.close 
									
									
							else
							imgthis = "blank.gif"
							end if
							
							if ressourcetimerDato <> 0 then
							ressourcetimerDato = ressourcetimerDato
							else
							ressourcetimerDato = ""
							end if
							
							select case weekday(getnyDato)
							case 1,7
							bgcolordatoer = "#cccccc"
							case else
							bgcolordatoer = "#ffffff"
							end select
							
							%>	
							<td valign="top" align="center" width=14 height=20 bgcolor="<%=bgcolordatoer%>" style="border-left:1px #d6dff5 solid; border-top:1px #ffffff solid; padding-top:4px;"><img src="../ill/<%=imgthis%>" width="6" height="6" alt="" border="0"><br><font class=megetmegetlillesort><%=ressourcetimerDato%></td>
							<%
						next%>
				</tr>
				</table>
		</td></tr>
		<%
		x = x + 1
	oRec.movenext
	wend
	oRec.close
	%>
	</tr>
	</table>
	</div>
	
	
	
	
	
	<!--- restimer -->
	
	
	<!--#include file="ressource_belaeg_jbpla.asp"-->
	
	

<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
