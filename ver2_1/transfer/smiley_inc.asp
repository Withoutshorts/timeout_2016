
<%





function afslutuge()%>

<table border='0' cellspacing='0' cellpadding='0' width=130>
<form action="timereg_akt_2006.asp?func=opdatersmiley&fromsdsk=<%=fromsdsk%>" method="POST">
<tr>
   
	<td width=130 valign=top bgcolor="#ffffff" style="border-left:1px #8caae6 solid; border-right:1px #8caae6 solid; border-bottom:1px #8caae6 solid; padding:5px;">
	<%
	
	
	
	
	    ugeNrAfsluttet = "1-1-2044"
		detteaar = datepart("yyyy", tjekdag(7))
		sidstedagiuge = datepart("ww", tjekdag(7), 2, 2) 
		
		call erugeAfslutte(detteaar, sidstedagiuge, usemrn)
		
		
		
		
		'Response.Write "ugeNrAfsluttet: " & ugeNrAfsluttet &"#"& datepart("ww",ugeNrAfsluttet, 2,2)& "<br>"
		
		%>
		<b><%=tsa_txt_087 %>?</b>
		<br><a href="#" onClick="visSmileystatus();" class=vmenu><%=tsa_txt_088 %> >></a><br>
		<%
		if showAfsuge = 1 then%>
				<%
				'*** Man må gerne lukke uger frem i tiden 27012009 ***'
				'if (datepart("ww", tjekdag(7), 2, 2) <= datepart("ww", now, 2, 2) AND year(tjekdag(7)) <= year(now)) OR  year(tjekdag(7)) < year(now) then%>
				<input type="hidden" name="FM_mid" id="FM_mid" value="<%=usemrn%>">
				<input type="hidden" name="FM_afslutuge_sidstedag" id="FM_afslutuge_sidstedag" value="<%=tjekdag(7)%>">
				<input type="checkbox" name="FM_afslutuge" id="FM_afslutuge" value="1" onClick="showlukalleuger()"><%=tsa_txt_089 %> <b><%=datepart("ww", tjekdag(7), 2, 2)%></b>.
				<!--<font class=megetlillesort>Ved markering inden søndag kl. 24:00 vil denne uge betragtes som rettidigt afsluttet.</font><br>-->
				
				<div id="lukalleuger" name="lukalleuger" style="position:relative; visibility:hidden; display:none;">
				<input type="checkbox" name="FM_alleuger" id="FM_alleuger" value="1"><%=tsa_txt_090 %>: <%=datepart("ww", dateadd("ww", -1, tjekdag(7)), 2, 2)%>
				</div>
				<input type="submit" value="<%=tsa_txt_087 %>"> 
				
				<!--
				<else>
				<br />
				<font class=megetlillesort><=tsa_txt_005 %> <b><=datepart("ww", tjekdag(7), 2, 2)%></b> <=tsa_txt_091 %> <=weekdayname(weekday(tjekdag(1)))%>  <=tsa_txt_092 %> <b><=formatdatetime(tjekdag(1), 2) %></b>
				<end if>
				-->
		<%else%><br>
		<b><%=tsa_txt_005 %>&nbsp;<%=datepart("ww", tjekdag(7), 2, 2)%>&nbsp;<%=tsa_txt_093 %> <font color=limegreen><i>V</i></font></b>
		<br><font class=megetlillesort><%=tsa_txt_093 %>&nbsp;<%=weekdayname(weekday(cdAfs))%>&nbsp;<%=tsa_txt_092 %> <br /> 
		<%=formatdatetime(cdAfs, 2)%>&nbsp;<%=formatdatetime(cdAfs, 3)%><br /> 
		(<%=tsa_txt_095 %>&nbsp;<%=datepart("ww", cdAfs, 2, 2)%>)
		<%end if%>
	
	
	</td>
	</tr></form>
</table> 
<%end function%>




<%
function showsmiley()

 '***************************** Smiley ******************************
	%>
	<table border='0' cellspacing='0' cellpadding='0' width=130>
	<tr><td width=130 bgcolor="#ffffff" style="border:1px #8caae6 solid; border-right:1px #8caae6 solid; border-bottom:0px #ffffff solid; padding:5px;">
	<%
	
	'**** Sur Smiley overruler ******'
	antalAfsDato = 0
	
	'*** startdato: 1 DEC 2005 ***
	if year(useDate) <= 2005 then
	surDatoSQLSTART = "2005/11/1" '12
		if lto = "execon" then
		antalAfsDato = 47 '47 '47 Sidste hele uge inden man går ind i december.
		else
		antalAfsDato = 43 '43 Sidste hele uge inden man går ind i december.
		end if
	else
	surDatoSQLSTART = year(useDate)&"/1/1"
	antalAfsDato = 0
	end if
	
	surDatoSQLEND = year(useDate)&"/"&month(useDate)&"/"&day(useDate)
	strSQL3 = "SELECT uge, count(afsluttet) AS antalafs FROM ugestatus "_
	&" WHERE mid = "& usemrn &" AND uge BETWEEN '"& surDatoSQLSTART &"' AND '"& surDatoSQLEND &"'"_
	&" GROUP BY mid ORDER BY afsluttet DESC"
	
	'Response.write  strSQL3
	'Response.flush
	oRec3.open strSQL3, oConn, 3 
	if not oRec3.EOF then
		antalAfsDato = antalAfsDato + oRec3("antalafs")
	end if
	oRec3.close  
	
	useDateWeek = datepart("ww", useDate, 2, 2)
	'Response.write "lastAfsDatoWeek: " & antalAfsDato & " = "&  useDateWeek
	
	
	'** Glade ***'
	if (antalAfsDato + 1) >= useDateWeek then
	surSmil = 0
	
				'*** Viser Smiley ***
				if weekday(useDate, 2) = 1 then
				periodeDage = 6
				minusenuge = 1
				else
				periodeDage = datediff("d", mandag, useDate)
				minusenuge = 0
				end if
				
				
				smileyok = 1
				antaldage = 1
				
				'*** Tjekker om det skal være en glad eller en mellem fornøjet smiley, **'
				'***  baseret på medarbejder normerede uge i den valgte uge. **'
				 
				if weekday(useDate, 2) = 1 then '** Mandag altid glad Smiley
					
					antaldage = antaldage
					smileyok = smileyok
				
				else
				
							for i = 1 to periodeDage + 1
								
							thisDato = dateadd("ww", -(minusenuge), tjekdag(i))
							'Response.write thisDato & "# "
							thisDatoSQL = year(thisDato)&"/"&month(thisDato)&"/"&day(thisDato)
								
							select case weekday(thisDato, 2)
							case 7
							normdag = "normtimer_son"
							case 1
							normdag = "normtimer_man" 
							case 2
							normdag = "normtimer_tir"
							case 3
							normdag = "normtimer_ons"
							case 4
							normdag = "normtimer_tor"
							case 5
							normdag = "normtimer_fre"
							case 6
							normdag = "normtimer_lor"
							case else
							normdag = "normtimer_man"
							end select
				
							
							'*** Finder nuværende type **'
							timerThis = 0
							strSQL = "SELECT medarbejdertype, "& normdag &" AS timer FROM medarbejdere m"_
							&" LEFT JOIN medarbejdertyper t ON (t.id = m. medarbejdertype)"_
							&" WHERE m.mid = " & usemrn
							
							'Response.write strSQL & "<br><br>"
							'Response.flush
							
							oRec.open strSQL, oConn, 3 
							if not oRec.EOF then
							timerThis = oRec("timer")
									
									'*** Er der normeret timer på medarbejder denne dag i ugen uge **'
									if timerThis <> 0 then
										
										strSQL2 = "SELECT sum(timer) AS stimer FROM timer WHERE tdato = '"& thisDatoSQL &"' AND tmnr = "& usemrn &" GROUP BY tdato"
										'Response.write strSQL2 & "<br>"
										oRec2.open strSQL2, oConn, 3 
										while not oRec2.EOF 
										
										if oRec2("stimer") > 0 then
										smileyok = smileyok + 1
										end if
										
										oRec2.movenext
										wend
										oRec2.close 
									else
										smileyok = smileyok + 1
									end if
							oRec.movenext
							end if
							oRec.close 
							
							antaldage = antaldage + 1
							next
				end if
	else
	surSmil = 1
	end if
	
	smileyRnd = right(cint(Second(now)), 1)
	
	select case smileyRnd
	case 0, 1 
	smilVal = 1
	case 2, 3
	smilVal = 2
	case 4, 5
	smilVal = 3
	case 6, 7
	smilVal = 4
	case else
	smilVal = 5
	end select
		
		%>
		
		<b><%=tsa_txt_005 %>: <%=datepart("ww", useDate, 2, 2)%></b>&nbsp;
		
		<%	
		
		if surSmil = 1 then
				%>
				<b><%=tsa_txt_096 %></b> <br>
				<img src="../ill/sur_<%=smilVal%>.gif" alt="" border="0"><br>
				<font class=megetlillesort><%=tsa_txt_097 %><br></font>
				<%
		else
		
				if smileyok >= (antaldage-1) then
				%>
				<b><%=tsa_txt_098 %></b> <br>
				<img src="../ill/gladsmil_<%=smilVal%>.gif" alt="" border="0"><br>
				<!--<font class=megetlillesort>Du mangler ikke nogen timeregistreringer i denne uge.<br>-->
				<%
				else
				%>
				<b><%=tsa_txt_099 %></b> <br>
				<img src="../ill/mellemsmil_<%=smilVal%>.gif" alt="" border="0"><br>
				<font class=megetlillesort><%=tsa_txt_100 %><br></font>
				<%end if
		end if%>
		
		
		<!--<font class=megetlillesort><br>Nb: Smileyordning kan slås til og fra i medarbejder-profilen.</font>-->
		
	
	
	
	</td></tr>
		</table>
<%end function%>



<%
function smileystatus(medarbid)





if len(medarbid) <> 0 then
medarbid = cint(medarbid)
else
medarbid = 0
end if
    
    if thisfile <> "smileystatus.asp" then
    if cint(year(now)) = cint(useYear) then
	denneuge = datepart("ww", year(now)&"/"&month(now)&"/"&day(now), 2, 2)
	else
		if cint(year(now)) < cint(useYear) then
		denneuge = 0
		else
		denneuge = 52
		end if
	end if
	
	startdato = useYear&"/1/1"
	slutdato = useYear&"/12/31"
	
	else
	
    startdato = strAar&"/"&strMrd&"/"&strDag
	slutdato = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut
	
	startdatoBeregn = strDag&"/"&strMrd&"/"&strAar
	slutdatoBeregn = strDag_slut&"/"&strMrd_slut&"/"&strAar_slut
	denneuge = datediff("ww", startdatoBeregn, slutdatoBeregn, 2,2)
	
	end if
	
	'Response.write denneuge
	
	'antaluger = datediff("ww", cdate(startdato), cdate(slutdato), 2, 3)
	'Response.flush
	
	dim glade, sure, mnavn, mnr, ugerafsluttet, sletafsl
	Redim glade(0), sure(0), mnavn(0), mnr(0)
	
	m = 100
	v = 5200
	redim ugerafsluttet(m,v) 
	redim sletafsl(m,v) 
	redim yearaf(m,v)
	
	'Response.write medarbid
	
	if cint(medarbid) <> 0 then
	medarbKri = " m.mid = "& medarbid 
	else
	medarbKri = " mansat <> 2 " 
	end if
	
	strSQL = "SELECT m.mid AS mid, m.mnavn, m.mnr FROM medarbejdere m"_
	&" WHERE "& medarbKri &" ORDER BY m.mid"
	oRec.open strSQL, oConn, 3 
	
	'Response.write strSQL & "<hr>"
	
	x = 0
	lastmid = 0
	while not oRec.EOF 
	
	'Response.write lastmid &"<>"& oRec("mid") &"<br>"
	
		x = x + 1
		lastmid = oRec("mid")
		Redim preserve mnavn(x)
		Redim preserve mnr(x)
		Redim preserve glade(x)
		Redim preserve sure(x)
		
		
		mnavn(x) = oRec("mnavn")
		mnr(x) = oRec("mnr")
		
		'Response.write mnavn(x) & "<br>"
		
					strSQL2 = "SELECT u.status, u.afsluttet, u.uge, u.id FROM ugestatus u WHERE u.mid = "& oRec("mid") &" AND uge BETWEEN '"& startdato &"' AND '"& slutdato &"' ORDER BY uge" 
					'AND (WEEK(uge,1) BETWEEN 1 AND "& denneuge &" 
					
					'Response.write strSQL2 &"<br>"
					
					v = 0
					oRec2.open strSQL2, oConn, 3 
					while not oRec2.EOF 
					
					if v = 0 AND datepart("ww", oRec2("uge"), 2, 2) = 52 then
					
					else
					
					if oRec2("status") = 2 then
					glade(x) = glade(x) + 1
					        
					        if level = 1 AND thisfile <> "timereg_akt_2006" then
						    ugerafsluttet(x,v) = "<a href=smileystatus.asp?func=slet&id="&oRec2("id")&" class=vmenuglobal>"&datepart("ww", oRec2("uge"), 2, 2)&"</a>"
					        else
					        ugerafsluttet(x,v) = "<b>"&datepart("ww", oRec2("uge"), 2, 2)&"</b>"
					        end if
					        
					else
					sure(x) = sure(x) + 1
					        if level = 1 AND thisfile <> "timereg_akt_2006" then
						    ugerafsluttet(x,v) = "<a href=smileystatus.asp?func=slet&id="&oRec2("id")&" class=vmenualt>"&datepart("ww", oRec2("uge"), 2, 2)&"</a>"
					        else
					        ugerafsluttet(x,v) = datepart("ww", oRec2("uge"), 2, 2)
					        end if
					end if
					
					sletafsl(x,v) = oRec2("id")
					yearaf(x,v) = datepart("yyyy", oRec2("uge"), 2, 2)
					v = v + 1
					
					end if
					
					
					oRec2.movenext
					wend
					oRec2.close 
	
	oRec.movenext
	wend
	oRec.close 
	
	antalx = x
	'Response.write x
	
	%>
	
	
<table border=0 cellspacing=0 cellpadding=0 width=600>
<tr bgcolor="#5582D2">
	<td rowspan=2 height=30 style="border-top:1px #003399 solid; border-left:1px #003399 solid;">&nbsp;</td>
	<td colspan=4 class='alt' valign=middle style="border-top:1px #003399 solid;">&nbsp;</td>
	<td rowspan=2 style="border-top:1px #003399 solid; border-right:1px #003399 solid;">&nbsp;</td>
</tr>
<tr bgcolor="#5582D2">
	<td width=200 class=alt><b><%=tsa_txt_101 %></b></td>
	<td class=alt align=right><b><%=tsa_txt_102 %></b></td>
	<td class=alt align=right><b><%=tsa_txt_103 %></b></td>
	<td class=alt align=right><b><%=tsa_txt_104 %></b></td>
</tr>
	
	<%
	lastyear = 2000
	'Response.flush
	if x <> 0 then
			for x = 1 to antalx ' - 1
			%>
			<tr bgcolor="#ffffff">
				<td height=20 style="border-left:1px #003399 solid;">&nbsp;</td>
				<td><b><%=mnavn(x)%> (<%=mnr(x)%>)</b></td>
				<td align=right><b><%=glade(x)%></b></td>
				<td align=right><%=sure(x)%></td>
				<td align=right><%=glade(x)+sure(x)%> / <%=denneuge%></td>
				<td style="border-right:1px #003399 solid;">&nbsp;</td>
			</tr>
			<tr bgcolor="#ffffff"><td colspan=6 style="border-left:1px #003399 solid; border-right:1px #003399 solid; padding:10px;">
				<%=tsa_txt_105 %>:<br>
				<%
				for v = 0 to UBOUND(ugerafsluttet, 1)
				if len(ugerafsluttet(x,v)) <> 0 then%>
					
					<%if right(v, 1) = 0 AND v > 0 then%>
					<br>
					<%end if%>
					
					<%if lastyear <> yearaf(x,v) then %>
					<br /><b><%=yearaf(x,v)%></b><br />
					<%end if %>
					
				 		- <%=ugerafsluttet(x,v)%>
				 		
				<%	
				
				lastyear = yearaf(x,v)
				end if
				next
				%>
			</td></tr>
			<tr bgcolor="#ffffff">
				<td style="border-bottom:1px #003399 solid; border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
				<td colspan=4 class='alt' style="border-bottom:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
				<td style="border-bottom:1px #003399 solid; border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
			</tr>
			<%next
	else%>
	<tr bgcolor="#ffffff">
		<td height=100 style="border-left:1px #003399 solid;">&nbsp;</td>
		<td colspan=4 style=padding-left:20px;><%=tsa_txt_106 %></td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<%end if%>
	<tr bgcolor="#5582D2">
		<td height=10 style="border-bottom:1px #003399 solid; border-left:1px #003399 solid;">&nbsp;</td>
		<td colspan=4 class='alt' valign=middle style="border-bottom:1px #003399 solid;">&nbsp;</td>
		<td style="border-bottom:1px #003399 solid; border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	</table>
	<%end function%>