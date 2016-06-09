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
	strWeek = datepart("ww", useDate, 2, 2)
	'Response.Write useDate & " -- "
	'Response.Write strWeek
	strRegAar = strAar 'skal slettes og sættes til strAar
	'******************************** 
	
	'**** Finder datoer på ugens 7 dage **************
	strDtypeToday = weekday(useDate, 2)
	'Response.write strDtypeToday
	
	select case strDtypeToday
	case 1
	mandag = dateadd("d", 0 , useDate)
	sondag = dateadd("d", 6 , useDate)
	tirsdag = dateadd("d", 1 , useDate)
	onsdag = dateadd("d", 2 , useDate)
	torsdag = dateadd("d", 3 , useDate)
	fredag = dateadd("d", 4 , useDate)
	lordag = dateadd("d", 5 , useDate)
	case 2
	tirsdag = dateadd("d", 0, useDate)
	sondag = dateadd("d", 5, useDate)
	mandag = dateadd("d", -1, useDate)
	onsdag = dateadd("d", 1, useDate)
	torsdag = dateadd("d", 2, useDate)
	fredag = dateadd("d", 3, useDate)
	lordag = dateadd("d", 4, useDate)
	case 3
	onsdag = dateadd("d", 0 , useDate)
	sondag = dateadd("d", 4 , useDate)
	mandag = dateadd("d", -2 , useDate)
	tirsdag = dateadd("d", -1 , useDate)
	torsdag = dateadd("d", 1 , useDate)
	fredag = dateadd("d", 2 , useDate)
	lordag = dateadd("d", 3 , useDate)
	case 4
	torsdag = dateadd("d", 0 , useDate)
	sondag = dateadd("d", 3, useDate)
	mandag = dateadd("d", -3 , useDate)
	tirsdag = dateadd("d", -2 , useDate)
	onsdag = dateadd("d", -1 , useDate)
	fredag = dateadd("d", 1 , useDate)
	lordag = dateadd("d", 2 , useDate)
	case 5
	fredag = dateadd("d", 0 , useDate)
	sondag = dateadd("d", 2 , useDate)
	mandag = dateadd("d", -4 , useDate)
	tirsdag = dateadd("d", -3 , useDate)
	onsdag = dateadd("d", -2 , useDate)
	torsdag = dateadd("d", -1 , useDate)
	lordag = dateadd("d", 1 , useDate)
	case 6
	lordag = dateadd("d", 0 , useDate)
	sondag = dateadd("d", 1 , useDate)
	mandag = dateadd("d", -5 , useDate)  
	tirsdag = dateadd("d", -4 , useDate)
	onsdag = dateadd("d", -3 , useDate)
	torsdag = dateadd("d", -2 , useDate)
	fredag = dateadd("d", -1 , useDate)
	case 7
	sondag = dateadd("d", 0 , useDate)
	mandag = dateadd("d", -6 , useDate)
	tirsdag = dateadd("d", -5 , useDate)
	onsdag = dateadd("d", -4 , useDate)
	torsdag = dateadd("d", -3 , useDate)
	fredag = dateadd("d", -2 , useDate)
	lordag = dateadd("d", -1 , useDate)
	end select
	
	
	Redim tjekdag(7)
	tjekdag(7) = sondag
	tjekdag(1) = mandag
	tjekdag(2) = tirsdag
	tjekdag(3) = onsdag
	tjekdag(4) = torsdag
	tjekdag(5) = fredag
	tjekdag(6) = lordag
	
	Redim writedagnavn(7)
	writedagnavn(7) = tsa_txt_127 '"Sø"
	writedagnavn(1) = tsa_txt_128 '"Ma"
	writedagnavn(2) = tsa_txt_129 '"Ti"
	writedagnavn(3) = tsa_txt_130 '"On"
	writedagnavn(4) = tsa_txt_131 '"To"
	writedagnavn(5) = tsa_txt_132 '"Fr"
	writedagnavn(6) = tsa_txt_133 '"Lø"
	
	dim helligdageCol
	redim helligdageCol(7)
	
	
	
	
	
public function dageDatoer()%>
<tr>
<td colspan=9 bgcolor="#FFFFFF" style="padding:5px 5px 4px 5px; border-bottom:1px #cccccc solid;">
<b><%=tsa_txt_005 %>: 
<%
Response.write strWeek & "&nbsp;&nbsp;&nbsp;" & datepart("yyyy", useDate)
%>
</b></td>
</tr>
<tr>
<td bgcolor="#ffffff" style="border-bottom:1px #8caae6 solid; white-space:nowrap;">
   <img src="../ill/nut_and_bolt.png" alt="" border="0">
  <b><%=tsa_txt_125 %></b> <br /><span style="font-size:9px;">Type // <%=tsa_txt_222 & "ell. "& tsa_txt_186 %> // Forkalk. // Real. ialt (medarb.) <!--Aktivitet--></span> </td>
  <td bgcolor="#ffffff" style="border-bottom:1px #8caae6 solid; border-right:1px #8caae6 solid;">&nbsp;</td>

<%
t = 1
for t = 1 to 7
	'*** Gennemløber de 7 dage ***
	if formatDatetime(now, 2) = formatDatetime(tjekdag(t), 2) then
	borderColor = "red" '"#ffff99"
	border = 2
	bgcoldg = "#FFFFFF"
	else
	
	call helligdage(tjekdag(t), 0)
	helligdageCol(t) = erHellig
	
		if (t = 6  OR t = 7) OR helligdageCol(t) = 1 then
		borderColor = "#999999"
		border = 1
		bgcoldg = "gainsboro"
		else
		borderColor = "#8caae6"
		border = 1
		bgcoldg = "#FFFFFF"
		end if
	end if%>
	<td valign="top" style="width:65px; background-color: <%=bgcoldg%>; border-right:1px #8caae6 solid;  border-bottom:<%=border%>px <%=borderColor%> solid;" class=lille>
	<%=writedagnavn(t)%> <br>
	
	<%
	Response.write formatdatetime(tjekdag(t), 2) &"<br>"
	call helligdage(tjekdag(t), 1)%></td>
	<%
next
%>
</tr>
<%
end function


%>





