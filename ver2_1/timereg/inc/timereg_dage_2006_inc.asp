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
	
  
	
	
	dim helligdageCol
	redim helligdageCol(7)
	
	
	
	
public dageDatoerTxt	
public function dageDatoer(vis)

Redim writedagnavn(7)
writedagnavn(7) = tsa_txt_127 '"Sø"
writedagnavn(1) = tsa_txt_128 '"Ma"
writedagnavn(2) = tsa_txt_129 '"Ti"
writedagnavn(3) = tsa_txt_130 '"On"
writedagnavn(4) = tsa_txt_131 '"To"
writedagnavn(5) = tsa_txt_132 '"Fr"
writedagnavn(6) = tsa_txt_133 '"Lø"


'*** Medarb ***'
call meStamdata(usemrn)
    
    if len(trim(meInit)) <> 0 then
    meTxt = meNavn & " ["& meInit &"]"
    else
    meTxt = meTxt
    end if

call jq_format(meTxt)
meTxt = jq_formatTxt


if vis = 2 then
cspsThis = 5
bdsThis = 1
else
cspsThis = 4
bdsThis = 1
end if

dageDatoerTxt = "<tr><td colspan=2 bgcolor=""#FFFFFF"" style=""padding:5px 5px 4px 5px; border-bottom:"&bdsThis&"px #cccccc solid;""><b>"&tsa_txt_005&": "& datepart("ww", tjekdag(1), 2,2) & "&nbsp;&nbsp;&nbsp;" & datepart("yyyy", tjekdag(1), 2,2) &"</b></td>"
dageDatoerTxt = dageDatoerTxt & "<td colspan="&cspsThis+3&" bgcolor=""#FFFFFF"" align=right style=""padding:5px 5px 4px 5px; border-bottom:"&bdsThis&"px #cccccc solid; font-size:10px;""><i> "& meTxt &" - "& normTimerprUge &" t. pr. uge</i></td>"

dageDatoerTxt = dageDatoerTxt & "</tr>"


dageDatoerTxt = dageDatoerTxt & "<tr><td colspan=2 bgcolor=""#ffffff"" valign=top style=""width:290px; border-bottom:"&bdsThis&"px #8caae6 solid; border-right:"&bdsThis&"px #8caae6 solid; white-space:nowrap;"">"

if vis = 1 then
    
    'dageDatoerTxt = dageDatoerTxt & "<b>"& tsa_txt_125 &"</b>" 
    '<img src=""../ill/nut_and_bolt.png"" border=""0"">

    
    'select case lto
    'case "unik", "mmmi"
    'case else

    'if visSimpelAktLinje <> "1" AND visHRliste <> "1" then
    'dageDatoerTxt = dageDatoerTxt & "<br /><span style=""font-size:9px;"">Periode | "& tsa_txt_222 & " ell. "& tsa_txt_186 &" | Forkalk. timer | Real. timer" 
    
    'if (cint(aktBudgettjkOn) = 1) then
    'dageDatoerTxt = dageDatoerTxt & "<br>| Forecast"
    'end if

     'dageDatoerTxt = dageDatoerTxt & "</span>"

    'end if

    'end select


    dageDatoerTxt = dageDatoerTxt & "</td>"

else
dageDatoerTxt = dageDatoerTxt & "&nbsp;</td>"
end if

'dageDatoerTxt = dageDatoerTxt & "<td bgcolor=""#ffffff"" style=""border-bottom:"&bdsThis&"px #8caae6 solid; border-right:"&bdsThis&"px #8caae6 solid;"">&nbsp;</td>"




t = 1
for t = 1 to 7
	'*** Gennemløber de 7 dage ***
	
	
	call helligdage(tjekdag(t), 0, lto, usemrn)
	
	
		if (t = 6  OR t = 7) OR erHellig = 1 then
		borderColor = "#999999"
        border = bdsThis
	    bgcoldg = "gainsboro"
		else
		borderColor = "#8caae6"
		border = bdsThis
		bgcoldg = "#FFFFFF"
		end if
    
    if formatDatetime(now, 2) = formatDatetime(tjekdag(t), 2) then
	borderColor = "red" '"#ffff99"
	border = 2
    end if

    
    '*** ÆØÅ **'
    call jq_format(writedagnavn(t))
    writedagnavn(t) = jq_formatTxt

    if t <> 7 then
	dageDatoerTxt = dageDatoerTxt & "<td valign=top style='width:65px; background-color: "&bgcoldg&"; border-right:"&bdsThis&"px #8caae6 solid;  border-bottom:"&border&"px "&borderColor&" solid;' class='lille'>"& writedagnavn(t) &"<br>"
	else
    dageDatoerTxt = dageDatoerTxt & "<td valign=top style='width:65px; background-color: "&bgcoldg&"; border-bottom:"&border&"px "&borderColor&" solid;' class='lille'>"& writedagnavn(t) &"<br>"
	end if

    dageDatoerTxt = dageDatoerTxt & formatdatetime(tjekdag(t), 2) &"<br>"
	
    'dageDatoerTxt = dageDatoerTxt & "<td>" & formatdatetime(tjekdag(t), 2) 
    call jq_format(helligdagnavnTxt)
    helligdagnavnTxt = jq_formatTxt

    'call helligdage(tjekdag(t), 0)
    dageDatoerTxt = dageDatoerTxt & "<span style=""font-size:9px; color:#5c75AA; line-height:9px;"">"& helligdagnavnTxt & "</span></td>"

next


if vis = 2 then
dageDatoerTxt = dageDatoerTxt & "<td bgcolor=#ffffff class=lille valign=top align=right style=""border-bottom:"&bdsThis&"px #8caae6 solid; border-right:"&bdsThis&"px #8caae6 solid;""><br>Total</td>"
end if

dageDatoerTxt = dageDatoerTxt & "</tr>"

end function


%>






