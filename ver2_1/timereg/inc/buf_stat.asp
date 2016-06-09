<%
dim mrdNotFakTimer(12)
mrdNotFakTimer(1) = 0
mrdNotFakTimer(2) = 0
mrdNotFakTimer(3) = 0
mrdNotFakTimer(4) = 0
mrdNotFakTimer(5) = 0
mrdNotFakTimer(6) = 0
mrdNotFakTimer(7) = 0
mrdNotFakTimer(8) = 0
mrdNotFakTimer(9) = 0
mrdNotFakTimer(10) = 0
mrdNotFakTimer(11) = 0
mrdNotFakTimer(12) = 0
dim mrdFakTimer(12)
mrdFakTimer(1) = 0
mrdFakTimer(2) = 0
mrdFakTimer(3) = 0
mrdFakTimer(4) = 0
mrdFakTimer(5) = 0
mrdFakTimer(6) = 0
mrdFakTimer(7) = 0
mrdFakTimer(8) = 0
mrdFakTimer(9) = 0
mrdFakTimer(10) = 0
mrdFakTimer(11) = 0
mrdFakTimer(12) = 0
dim mrdOms(12)
mrdOms(1) = 0
mrdOms(2) = 0
mrdOms(3) = 0
mrdOms(4) = 0
mrdOms(5) = 0
mrdOms(6) = 0
mrdOms(7) = 0
mrdOms(8) = 0
mrdOms(9) = 0
mrdOms(10) = 0
mrdOms(11) = 0
mrdOms(12) = 0

dim mrdTotTimer(12)

	strSQL = "SELECT Tid, Tdato, Timer, Tfaktim, Tjobnr, Tjobnavn, Taar, TAktivitetId, TAktivitetNavn, Tmnr, Tmnavn, timepris FROM timer WHERE Taar = "&year(now)&" ORDER BY Tid"
	oRec.Open strSQL, oConn, 3
	
	while not oRec.EOF
	
	select case datepart("m",oRec("Tdato"))
	case 1
	if oRec("Tfaktim") = 1 then
	mrdFakTimer(1) = mrdFakTimer(1) + oRec("timer")
	mrdOms(1) = mrdOms(1) + (oRec("Timer") * oRec("timepris"))
	else
	mrdNotFakTimer(1) = mrdNotFakTimer(1) + oRec("timer")
	mrdOms(1) = mrdOms(1)
	end if
	mrdTotTimer(1) = mrdTotTimer(1) + oRec("timer")
	
	case 2
	if oRec("Tfaktim") = 1 then
	mrdFakTimer(2) = mrdFakTimer(2) + oRec("timer")
	mrdOms(2) = mrdOms(2) + (oRec("Timer") * oRec("timepris"))
	else
	mrdNotFakTimer(2) = mrdNotFakTimer(2) + oRec("timer")
	mrdOms(2) = mrdOms(2)
	end if
	mrdTotTimer(2) = mrdTotTimer(2) + oRec("timer")
	
	
	case 3
	if oRec("Tfaktim") = 1 then
	mrdFakTimer(3) = mrdFakTimer(3) + oRec("timer")
	mrdOms(3) = mrdOms(3) + (oRec("Timer") * oRec("timepris"))
	else
	mrdNotFakTimer(3) = mrdNotFakTimer(3) + oRec("timer")
	mrdOms(3) = mrdOms(3)
	end if
	mrdTotTimer(3) = mrdTotTimer(3) + oRec("timer")
	
	case 4
	if oRec("Tfaktim") = 1 then
	mrdFakTimer(4) = mrdFakTimer(4) + oRec("timer")
	mrdOms(4) = mrdOms(4) + (oRec("Timer") * oRec("timepris"))
	else
	mrdNotFakTimer(4) = mrdNotFakTimer(4) + oRec("timer")
	mrdOms(4) = mrdOms(4)
	end if
	mrdTotTimer(4) = mrdTotTimer(4) + oRec("timer")
	
	case 5
	if oRec("Tfaktim") = 1 then
	mrdFakTimer(5) = mrdFakTimer(5) + oRec("timer")
	mrdOms(5) = mrdOms(5) + (oRec("Timer") * oRec("timepris"))
	else
	mrdNotFakTimer(5) = mrdNotFakTimer(5) + oRec("timer")
	mrdOms(5) = mrdOms(5)
	end if
	mrdTotTimer(5) = mrdTotTimer(5) + oRec("timer")
	
	
	case 6
	if oRec("Tfaktim") = 1 then
	mrdFakTimer(6) = mrdFakTimer(6) + oRec("timer")
	mrdOms(6) = mrdOms(6) + (oRec("Timer") * oRec("timepris"))
	else
	mrdNotFakTimer(6) = mrdNotFakTimer(6) + oRec("timer")
	mrdOms(6) = mrdOms(6)
	end if
	mrdTotTimer(6) = mrdTotTimer(6) + oRec("timer")
	
	
	case 7
	if oRec("Tfaktim") = 1 then
	mrdFakTimer(7) = mrdFakTimer(7) + oRec("timer")
	mrdOms(7) = mrdOms(7) + (oRec("Timer") * oRec("timepris"))
	else
	mrdNotFakTimer(7) = mrdNotFakTimer(7) + oRec("timer")
	mrdOms(7) = mrdOms(7)
	end if
	mrdTotTimer(7) = mrdTotTimer(7) + oRec("timer")
	
	
	case 8
	if oRec("Tfaktim") = 1 then
	mrdFakTimer(8) = mrdFakTimer(8) + oRec("timer")
	mrdOms(8) = mrdOms(8) + (oRec("Timer") * oRec("timepris"))
	else
	mrdNotFakTimer(8) = mrdNotFakTimer(8) + oRec("timer")
	mrdOms(8) = mrdOms(8)
	end if
	mrdTotTimer(8) = mrdTotTimer(8) + oRec("timer")
	
	
	case 9
	if oRec("Tfaktim") = 1 then
	mrdFakTimer(9) = mrdFakTimer(9) + oRec("timer")
	mrdOms(9) = mrdOms(9) + (oRec("Timer") * oRec("timepris"))
	else
	mrdNotFakTimer(9) = mrdNotFakTimer(9) + oRec("timer")
	mrdOms(9) = mrdOms(9)
	end if
	mrdTotTimer(9) = mrdTotTimer(9) + oRec("timer")
	
	case 10
	if oRec("Tfaktim") = 1 then
	mrdFakTimer(10) = mrdFakTimer(10) + oRec("timer")
	mrdOms(10) = mrdOms(10) + (oRec("Timer") * oRec("timepris"))
	else
	mrdNotFakTimer(10) = mrdNotFakTimer(10) + oRec("timer")
	mrdOms(10) = mrdOms(10)
	end if
	mrdTotTimer(10) = mrdTotTimer(10) + oRec("timer")
	
	case 11
	if oRec("Tfaktim") = 1 then
	mrdFakTimer(11) = mrdFakTimer(11) + oRec("timer")
	mrdOms(11) = mrdOms(11) + (oRec("Timer") * oRec("timepris"))
	else
	mrdNotFakTimer(11) = mrdNotFakTimer(11) + oRec("timer")
	mrdOms(11) = mrdOms(11)
	end if
	mrdTotTimer(11) = mrdTotTimer(11) + oRec("timer")
	
	case 12
	if oRec("Tfaktim") = 12then
	mrdFakTimer(12) = mrdFakTimer(12) + oRec("timer")
	mrdOms(12) = mrdOms(12) + (oRec("Timer") * oRec("timepris"))
	else
	mrdNotFakTimer(12) = mrdNotFakTimer(12) + oRec("timer")
	mrdOms(12) = mrdOms(12)
	end if
	mrdTotTimer(12) = mrdTotTimer(12) + oRec("timer")
	end select
	
	totTimer = totTimer + oRec("timer")
	oRec.movenext
	wend
	
	timerTotFakbar = mrdFakTimer(1) + mrdFakTimer(2) + mrdFakTimer(3) + mrdFakTimer(4) + mrdFakTimer(5) + mrdFakTimer(6) + mrdFakTimer(7) + mrdFakTimer(8) + mrdFakTimer(9) + mrdFakTimer(10) + mrdFakTimer(11) + mrdFakTimer(12)
	timerTotNotFakbar = mrdNotFakTimer(1) + mrdNotFakTimer(2) + mrdNotFakTimer(3) + mrdNotFakTimer(4) + mrdNotFakTimer(5) + mrdNotFakTimer(6) + mrdNotFakTimer(7) + mrdNotFakTimer(8) + mrdNotFakTimer(9) + mrdNotFakTimer(10) + mrdNotFakTimer(11) + mrdNotFakTimer(12)
	mrdOmsTotal = mrdOms(1) + mrdOms(2) + mrdOms(3) + mrdOms(4) + mrdOms(5) + mrdOms(6) + mrdOms(7) + mrdOms(8) + mrdOms(9) + mrdOms(10) + mrdOms(11) + mrdOms(12)
	
	oRec.close
	
	
	
	'****************************************************************
	intMrdOms_1 = (mrdOms(1)/1000)/4
	if intMrdOms_1 > 1200 then
	intMrdOms_1 = 1200
	end if
	
	intMrdOms_2 = (mrdOms(2)/1000)/4
	if intMrdOms_2 > 1200 then
	intMrdOms_2 = 1200
	end if
	
	intMrdOms_3 = (mrdOms(3)/1000)/4
	if intMrdOms_3 > 1200 then
	intMrdOms_3 = 1200
	end if
	
	intMrdOms_4 = (mrdOms(4)/1000)/4
	if intMrdOms_4 > 1200 then
	intMrdOms_4 = 1200
	end if
	
	intMrdOms_5 = (mrdOms(5)/1000)/4
	if intMrdOms_5 > 1200 then
	intMrdOms_5 = 1200
	end if
	
	intMrdOms_6 = (mrdOms(6)/1000)/4
	if intMrdOms_6 > 1200 then
	intMrdOms_6 = 1200
	end if
	
	intMrdOms_7 = (mrdOms(7)/1000)/4
	if intMrdOms_7 > 1200 then
	intMrdOms_7 = 1200
	end if
	
	intMrdOms_8 = (mrdOms(8)/1000)/4
	if intMrdOms_8 > 1200 then
	intMrdOms_8 = 1200
	end if
	
	intMrdOms_9 = (mrdOms(9)/1000)/4
	if intMrdOms_9 > 1200 then
	intMrdOms_9 = 1200
	end if
	
	intMrdOms_10 = (mrdOms(10)/1000)/4
	if intMrdOms_10 > 1200 then
	intMrdOms_10 = 1200
	end if
	
	intMrdOms_11 = (mrdOms(11)/1000)/4
	if intMrdOms_11 > 1200 then
	intMrdOms_11 = 1200
	end if
	
	intMrdOms_12 = (mrdOms(12)/1000)/4
	if intMrdOms_12 > 1200 then
	intMrdOms_12 = 1200
	end if
	
	
	'************************************************************
	
	intFakTimer_1 = left(mrdFakTimer(1),4)/10
	if intFakTimer_1 > 3000 then
	intFakTimer_1 = 3000
	end if
	
	intFakTimer_2 = left(mrdFakTimer(2),4)/10
	if intFakTimer_2 > 3000 then
	intFakTimer_2 = 3000
	end if
	
	intFakTimer_3 = left(mrdFakTimer(3),4)/10
	if intFakTimer_3 > 3000 then
	intFakTimer_3 = 3000
	end if
	
	intFakTimer_4 = left(mrdFakTimer(4),4)/10
	if intFakTimer_4 > 3000 then
	intFakTimer_4 = 3000
	end if
	
	intFakTimer_5 = left(mrdFakTimer(5),4)/10
	if intFakTimer_5 > 3000 then
	intFakTimer_5 = 3000
	end if
	
	intFakTimer_6 = left(mrdFakTimer(6),4)/10
	if intFakTimer_6 > 3000 then
	intFakTimer_6 = 3000
	end if
	
	intFakTimer_7 = left(mrdFakTimer(7),4)/10
	if intFakTimer_7 > 3000 then
	intFakTimer_7 = 3000
	end if
	
	
	intFakTimer_8 = left(mrdFakTimer(8),4)/10
	if intFakTimer_8 > 3000 then
	intFakTimer_8 = 3000
	end if
	
	intFakTimer_9 = left(mrdFakTimer(9),4)/10
	if intFakTimer_9 > 3000 then
	intFakTimer_9 = 3000
	end if
	
	intFakTimer_10 = left(mrdFakTimer(10),4)/10
	if intFakTimer_10 > 3000 then
	intFakTimer_10 = 3000
	end if
	
	intFakTimer_11 = left(mrdFakTimer(11),4)/10
	if intFakTimer_11 > 3000 then
	intFakTimer_11 = 3000
	end if
	
	intFakTimer_12 = left(mrdFakTimer(12),4)/10
	if intFakTimer_12 > 3000 then
	intFakTimer_12 = 3000
	end if
	
	'*******************************************************
	intNotFakTimer_1 = left(mrdNotFakTimer(1),4)/10
	if intNotFakTimer_1 > 3000 then
	intNotFakTimer_1 = 3000
	end if
	
	intNotFakTimer_2 = left(mrdNotFakTimer(2),4)/10
	if intNotFakTimer_2 > 3000 then
	intNotFakTimer_2 = 3000
	end if
	
	intNotFakTimer_3 = left(mrdNotFakTimer(3),4)/10
	if intNotFakTimer_3 > 3000 then
	intNotFakTimer_3 = 3000
	end if
	
	intNotFakTimer_4 = left(mrdNotFakTimer(4),4)/10
	if intNotFakTimer_4 > 3000 then
	intNotFakTimer_4 = 3000
	end if
	
	intNotFakTimer_5 = left(mrdNotFakTimer(5),4)/10
	if intNotFakTimer_5 > 3000 then
	intNotFakTimer_5 = 3000
	end if
	
	intNotFakTimer_6 = left(mrdNotFakTimer(6),4)/10
	if intNotFakTimer_6 > 3000 then
	intNotFakTimer_6 = 3000
	end if
	
	intNotFakTimer_7 = left(mrdNotFakTimer(7),4)/10
	if intNotFakTimer_7 > 3000 then
	intNotFakTimer_7 = 3000
	end if
	
	intNotFakTimer_8 = left(mrdNotFakTimer(8),4)/10
	if intNotFakTimer_8 > 3000 then
	intNotFakTimer_8 = 3000
	end if
	
	intNotFakTimer_9 = left(mrdNotFakTimer(9),4)/10
	if intNotFakTimer_9 > 3000 then
	intNotFakTimer_9 = 3000
	end if
	
	intNotFakTimer_10 = left(mrdNotFakTimer(10),4)/10
	if intNotFakTimer_10 > 3000 then
	intNotFakTimer_10 = 3000
	end if
	
	intNotFakTimer_11 = left(mrdNotFakTimer(11),4)/10
	if intNotFakTimer_11 > 3000 then
	intNotFakTimer_11 = 3000
	end if
	
	intNotFakTimer_12 = left(mrdNotFakTimer(12),4)/10
	if intNotFakTimer_12 > 3000 then
	intNotFakTimer_12 = 3000
	end if
	%>

<br>

</div>
<DIV ID="Menu<%=de%>" Style="position:absolute; left:228; top:88; height:1000; display: none; background-color:#D6DFF5;">
<div style="position:absolute; top:0; left:5; width:75; height:15;"><table><tr><td>Omsætning:</td><td style='!border: 1px; background-color: DarkRed; border-color: #000000; border-style: solid; padding-left : 4px;padding-right : 4px;'><img src="ill/blank.gif" width="10" height="1" alt="" border="0"></td></tr></table></div>
<div style="position:absolute; top:0; left:103; width:25; height:15;"><table><tr><td>Eksterne:</td><td style='!border: 1px; background-color: #d2691e; border-color: #000000; border-style: solid; padding-left : 4px;padding-right : 4px;'><img src="ill/blank.gif" width="10" height="1" alt="" border="0"></td></tr></table></div>
<div style="position:absolute; top:0; left:185; width:25; height:15;"><table><tr><td>Interne:</td><td style='!border: 1px; background-color: #add8e6; border-color: #000000; border-style: solid; padding-left : 4px;padding-right : 4px;'><img src="ill/blank.gif" width="10" height="1" alt="" border="0"></td></tr></table></div>
<div style="position:absolute; top:98; left:63; height:1px; z-index:100;"><img src="ill/scala_streg.gif" width="370" height="1" alt="" border="0"></div>
<div style="position:absolute; top:198; left:63; height:1px; z-index:100;"><img src="ill/scala_streg.gif" width="370" height="1" alt="" border="0"></div>
<div style="position:absolute; top:299; left:63; height:1px; z-index:100;"><img src="ill/scala_streg.gif" width="370" height="1" alt="" border="0"></div>
<br><br>
<a href="stat.asp?menu=stat">2002</a><br>
<table cellspacing="1" cellpadding="0" border="0" width="390" bgcolor="LightSlateGray">
<tr bgcolor="#ffffff">
	<td align="center" width="60"><img src="ill/blank.gif" width="1" height="1" alt="" border="0"></td>
    <td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Jan</td>
   	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Feb</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Mar</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Apr</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Maj</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Jun</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Jul</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Aug</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Sep</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Okt</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Nov</td>
	<td align="center"><img src="ill/blank.gif" width="1" height="1" alt="" border="0">Dec</td>
</tr>
<tr bgcolor="#ffffff">
    <td align="right" bgcolor="#ffffff"><img src="ill/scala_cur1200.gif" width="60" height="300" alt="" border="0"></td>
    <td valign="bottom" width="30"><img src="ill/scalag1.gif" width="10" height="<%=intMrdOms_1%>" alt="" border="0"><img src="ill/scalag2.gif" width="10" height="<%=intFakTimer_1%>" alt="" border="0"><img src="ill/scalag3.gif" width="10" height="<%=intNotFakTimer_1%>" alt="" border="0"></td>
	<td valign="bottom" width="30" bgcolor="WhiteSmoke"><img src="ill/scalag1.gif" width="10" height="<%=intMrdOms_2%>" alt="" border="0"><img src="ill/scalag2.gif" width="10" height="<%=intFakTimer_2%>" alt="" border="0"><img src="ill/scalag3.gif" width="10" height="<%=intNotFakTimer_2%>" alt="" border="0"></td>
	<td valign="bottom" width="30"><img src="ill/scalag1.gif" width="10" height="<%=intMrdOms_3%>" alt="" border="0"><img src="ill/scalag2.gif" width="10" height="<%=intFakTimer_3%>" alt="" border="0"><img src="ill/scalag3.gif" width="10" height="<%=intNotFakTimer_3%>" alt="" border="0"></td>
	<td valign="bottom" width="30" bgcolor="WhiteSmoke"><img src="ill/scalag1.gif" width="10" height="<%=intMrdOms_4%>" alt="" border="0"><img src="ill/scalag2.gif" width="10" height="<%=intFakTimer_4%>" alt="" border="0"><img src="ill/scalag3.gif" width="10" height="<%=intNotFakTimer_4%>" alt="" border="0"></td>
	<td valign="bottom" width="30"><img src="ill/scalag1.gif" width="10" height="<%=intMrdOms_5%>" alt="" border="0"><img src="ill/scalag2.gif" width="10" height="<%=intFakTimer_5%>" alt="" border="0"><img src="ill/scalag3.gif" width="10" height="<%=intNotFakTimer_5%>" alt="" border="0"></td>
	<td valign="bottom" width="30" bgcolor="WhiteSmoke"><img src="ill/scalag1.gif" width="10" height="<%=intMrdOms_6%>" alt="" border="0"><img src="ill/scalag2.gif" width="10" height="<%=intFakTimer_6%>" alt="" border="0"><img src="ill/scalag3.gif" width="10" height="<%=intNotFakTimer_6%>" alt="" border="0"></td>
	<td valign="bottom" width="30"><img src="ill/scalag1.gif" width="10" height="<%=intMrdOms_7%>" alt="" border="0"><img src="ill/scalag2.gif" width="10" height="<%=intFakTimer_7%>" alt="" border="0"><img src="ill/scalag3.gif" width="10" height="<%=intNotFakTimer_7%>" alt="" border="0"></td>
	<td valign="bottom" width="30" bgcolor="WhiteSmoke"><img src="ill/scalag1.gif" width="10" height="<%=intMrdOms_8%>" alt="" border="0"><img src="ill/scalag2.gif" width="10" height="<%=intFakTimer_8%>" alt="" border="0"><img src="ill/scalag3.gif" width="10" height="<%=intNotFakTimer_8%>" alt="" border="0"></td>
	<td valign="bottom" width="30"><img src="ill/scalag1.gif" width="10" height="<%=intMrdOms_9%>" alt="" border="0"><img src="ill/scalag2.gif" width="10" height="<%=intFakTimer_9%>" alt="" border="0"><img src="ill/scalag3.gif" width="10" height="<%=intNotFakTimer_9%>" alt="" border="0"></td>
	<td valign="bottom" width="30" bgcolor="WhiteSmoke"><img src="ill/scalag1.gif" width="10" height="<%=intMrdOms_10%>" alt="" border="0"><img src="ill/scalag2.gif" width="10" height="<%=intFakTimer_10%>" alt="" border="0"><img src="ill/scalag3.gif" width="10" height="<%=intNotFakTimer_10%>" alt="" border="0"></td>
	<td valign="bottom" width="30"><img src="ill/scalag1.gif" width="10" height="<%=intMrdOms_11%>" alt="" border="0"><img src="ill/scalag2.gif" width="10" height="<%=intFakTimer_11%>" alt="" border="0"><img src="ill/scalag3.gif" width="10" height="<%=intNotFakTimer_11%>" alt="" border="0"></td>
	<td valign="bottom" width="30" bgcolor="WhiteSmoke"><img src="ill/scalag1.gif" width="10" height="<%=intMrdOms_12%>" alt="" border="0"><img src="ill/scalag2.gif" width="10" height="<%=intFakTimer_12%>" alt="" border="0"><img src="ill/scalag3.gif" width="10" height="<%=intNotFakTimer_12%>" alt="" border="0"></td> 
</tr>
<tr bgcolor="#ffffff">
    <td bgcolor="DarkRed" align="center"><font size="1" color="#ffffff"><%=mrdOmsTotal%></td>
    <td valign="bottom" width="30" align="center"><font size="1"><%=mrdOms(1)%></font></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdOms(2)%></font></td> 
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdOms(3)%></font></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdOms(4)%></font></td> 
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdOms(5)%></font></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdOms(6)%></font></td> 
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdOms(7)%></font></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdOms(8)%></font></td> 
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdOms(9)%></font></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdOms(10)%></font></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdOms(11)%></font></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdOms(12)%></font></td>  
</tr>
<img src="../ill/blank.gif" width="500" height="1" alt="" border="0"><a href="javascript:expand('<%=de%>');" class='vmenu'>>> Se årsoversigt af timeforbrug <img src="ill/plus.gif" width="9" height="9" border="0" name="Menub<%=de%>"></a>

<tr bgcolor="#ffffff">
    <td bgcolor="#d2691e" align="center"><font size="1" color="#ffffff"><%=timerTotFakbar%></td>
    <td valign="bottom" width="30" align="center"><font size="1"><%=mrdFakTimer(1)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdFakTimer(2)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdFakTimer(3)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdFakTimer(4)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdFakTimer(5)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdFakTimer(6)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdFakTimer(7)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdFakTimer(8)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdFakTimer(9)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdFakTimer(10)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdFakTimer(11)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdFakTimer(12)%></td>
</tr>
<tr bgcolor="#ffffff">
    <td bgcolor="#add8e6" align="center"><font size="1"><%=timerTotNotFakbar%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdNotFakTimer(1)%></td>
    <td valign="bottom" width="30" align="center"><font size="1"><%=mrdNotFakTimer(2)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdNotFakTimer(3)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdNotFakTimer(4)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdNotFakTimer(5)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdNotFakTimer(6)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdNotFakTimer(7)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdNotFakTimer(8)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdNotFakTimer(9)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdNotFakTimer(10)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdNotFakTimer(11)%></td>
	<td valign="bottom" width="30" align="center"><font size="1"><%=mrdNotFakTimer(12)%></td>
</tr>
</table>
<br>
</div>