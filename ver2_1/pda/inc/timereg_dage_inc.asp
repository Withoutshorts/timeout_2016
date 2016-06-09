<%
sonTimerTot = 0
manTimerTot = 0
tirTimerTot = 0
onsTimerTot = 0
torTimerTot = 0
freTimerTot = 0
lorTimerTot = 0
areJobPrinted = 0

'*******************************************************************************************
'** Sætter værdier på de variable der skal bruges for at finde de datoer der skal vises
'*** Hvis der er valgt en uge 
'***/else Hvis der skal bruges dags dato. F.eks når man kommer ind
'*** på siden uden at have forhåndsvalgt parametre
'*** Bruges også af kalenderen
'*******************************************************************************************

	'*** Henter dato og dag mv. fra calender.asp
	'*** Finder den valgte uge **** 
	strWeek =  datepart("ww", useDate)
	strRegAar = strAar 'skal slettes og sættes til strAar
	'******************************** 
	
	'**** Finder datoer på ugens 7 dage **************
	strDtypeToday = weekday(useDate)
	
	select case strDtypeToday
	case 1
	sondag = dateadd("d", 0 , useDate)
	mandag = dateadd("d", 1 , useDate)
	tirsdag = dateadd("d", 2 , useDate)
	onsdag = dateadd("d", 3 , useDate)
	torsdag = dateadd("d", 4 , useDate)
	fredag = dateadd("d", 5 , useDate)
	lordag = dateadd("d", 6 , useDate)
	case 2
	mandag = dateadd("d", 0 , useDate)
	sondag = dateadd("d", -1 , useDate)
	tirsdag = dateadd("d", 1 , useDate)
	onsdag = dateadd("d", 2 , useDate)
	torsdag = dateadd("d", 3 , useDate)
	fredag = dateadd("d", 4 , useDate)
	lordag = dateadd("d", 5 , useDate)
	case 3
	tirsdag = dateadd("d", 0 , useDate)
	sondag = dateadd("d", -2 , useDate)
	mandag = dateadd("d", -1 , useDate)
	onsdag = dateadd("d", 1 , useDate)
	torsdag = dateadd("d", 2 , useDate)
	fredag = dateadd("d", 3 , useDate)
	lordag = dateadd("d", 4 , useDate)
	case 4
	onsdag = dateadd("d", 0 , useDate)
	sondag = dateadd("d", -3 , useDate)
	mandag = dateadd("d", -2 , useDate)
	tirsdag = dateadd("d", -1 , useDate)
	torsdag = dateadd("d", 1 , useDate)
	fredag = dateadd("d", 2 , useDate)
	lordag = dateadd("d", 3 , useDate)
	case 5
	torsdag = dateadd("d", 0 , useDate)
	sondag = dateadd("d", -4 , useDate)
	mandag = dateadd("d", -3 , useDate)
	tirsdag = dateadd("d", -2 , useDate)
	onsdag = dateadd("d", -1 , useDate)
	fredag = dateadd("d", 1 , useDate)
	lordag = dateadd("d", 2 , useDate)
	case 6
	fredag = dateadd("d", 0 , useDate)
	sondag = dateadd("d", -5 , useDate)
	mandag = dateadd("d", -4 , useDate)
	tirsdag = dateadd("d", -3 , useDate)
	onsdag = dateadd("d", -2 , useDate)
	torsdag = dateadd("d", -1 , useDate)
	lordag = dateadd("d", 1 , useDate)
	case 7
	lordag = dateadd("d", 0 , useDate)
	sondag = dateadd("d", -6 , useDate)
	mandag = dateadd("d", -5 , useDate)  
	tirsdag = dateadd("d", -4 , useDate)
	onsdag = dateadd("d", -3 , useDate)
	torsdag = dateadd("d", -2 , useDate)
	fredag = dateadd("d", -1 , useDate)
	end select
	
	
	Redim tjekdag(7)
	tjekdag(1) = sondag
	tjekdag(2) = mandag
	tjekdag(3) = tirsdag
	tjekdag(4) = onsdag
	tjekdag(5) = torsdag
	tjekdag(6) = fredag
	tjekdag(7) = lordag
	
	Redim writedagnavn(7)
	writedagnavn(1) = "Søn"
	writedagnavn(2) = "Man"
	writedagnavn(3) = "Tir"
	writedagnavn(4) = "Ons"
	writedagnavn(5) = "Tor"
	writedagnavn(6) = "Fre"
	writedagnavn(7) = "Lør"
	
	
	
public function dageDatoer()%>
<tr>
<td align="center" style="!background-color : #8DA7E5; border-left: 1px solid #003399; border-right: 1px solid #003399; width : 138px;">
<font style="color:#ffffff;font-family:arial,sans-serif;font-size:11px;"><b>Uge: 
<%
Response.write strWeek & "&nbsp;&nbsp;&nbsp;" & datepart("yyyy", useDate)
%>
</b></td>
<%
t = 1
for t = 1 to 7
	'*** Gennemløber de 7 dage ***
	if formatDatetime(useDate, 2) = formatDatetime(tjekdag(t), 2) then
	borderColor = "#D6DFF5"
	border = 2
	else
	borderColor = "#FFFFE1"
	border = 1
	end if%>
	<td height="25" valign="top" style="background-color: <%=borderColor%>; border-right: 1px solid #003399; width : 56px; padding-left:2; padding-right:2;">
	<font style="color:#000000; font-family:arial,sans-serif; font-size:9px;"><%=writedagnavn(t)%><br>
	<%
	'Len_datoThis = len(formatdatetime(tjekdag(t), 1))
	'writedatoThis = left(tjekdag(t), Len_datoThis - 2)
	'Response.write writedatoThis
	Response.write formatdatetime(tjekdag(t), 0) &"<br>"
	call helligdage(tjekdag(t))%></td>
	<%
next
%>
</tr>
<%
end function


'***************************************************************************************************
'*** CRM Kalenderen
'***

public function dageDatoer_crmKal(FiveOrOneView)%>
<tr bgcolor="#5582D2">
<td width="32" align="center" style="border-left: 1px solid #003399;">&nbsp;</td>
<%
if FiveOrOneView = 2 then
	%>
	<td height="20" class='alt' width="99" style="border-right: 1px solid #003399; padding-left:5;">
		<b><%=writedagnavn(weekday(useDate))%>&nbsp;<font class="lillesort">
		<%
		call helligdage(useDate)
		%>
		</font>
		<%
		Len_datoThis = len(formatdatetime(tjekdag(weekday(useDate)), 1))
		writedatoThis = left(tjekdag(weekday(useDate)), Len_datoThis - 3)
		Response.write "<br>"& writedatoThis%></b></td>
<%
else
t = 1
for t = 1 to 7
	if Weekday(tjekdag(t)) > 1 AND Weekday(tjekdag(t)) < 7 then
		'*** Gennemløber de 7 dage ***
		if t = 6 then
		bcol = "#003399"
		else
		bcol = "#FFFFFF"
		end if%>
		<td height="20" class='alt' width="99" style="border-right: 1px solid <%=bcol%>; padding-left:5;">
		<b><%=writedagnavn(t)%>&nbsp;<font class="lillesort">
		<%
		call helligdage(tjekdag(t))
		%>
		</font>
		<%
		
		Len_datoThis = len(formatdatetime(tjekdag(t), 1))
		writedatoThis = left(tjekdag(t), Len_datoThis - 3)
		Response.write "<br>" & writedatoThis%></b></td>
		<%
 	end if
next
end if
%>
</tr>
<%
end function

%>
<!--#include file="helligdage_func.asp"-->




