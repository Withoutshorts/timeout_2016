<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<%
if session("user") = "" then
%>
<!--#include file="../inc/regular/header_inc.asp"-->
<%
	errortype = 5
	call showError(errortype)
	else
	if request("FM_medarb") = "" then%>
		<!--#include file="../inc/regular/header_inc.asp"-->
		<!--#include file="../inc/regular/topmenu_inc.asp"-->
		<%
		errortype = 40
		call showError(errortype)
		else
	%>
	<%if request("print") <> "j" then%>
	<!--#include file="../inc/regular/header_inc.asp"-->	
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!--#include file="../inc/regular/rmenu.asp"-->
	<%else%>
	<html>
	<head>
		<title>timeOut 2.1</title>
		<LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print_fak.css">
	</head>
	<body topmargin="0" leftmargin="0" class="regular">
	<%end if%>
	<!--#include file="inc/convertDate.asp"-->
	
	<script language="javascript">
	<!--
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	//-->
	</script>
	<%
	
	
	function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
		end function
		
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
	
	linket = request("linket")
	FM_medarb = request("FM_medarb")
	
	selmedarb = FM_medarb 'request("selmedarb")
	if len(selmedarb) = "0" then
	selmedarb = 0
	else
	selmedarb = selmedarb
	end if
	
	selaktid = request("selaktid")
	FM_job = request("FM_job")
	intJobnr = FM_job 'request("jobnr")
	
	if len(Request("mrd")) <> "0" then
		strReqMrd = Request("mrd")
	else
	strReqMrd = month(now)
	end if
	
	'**** finder ud af om der er valgt et år ***
	if len(Request("year")) <> "0" then
		if Request("year") = "-1" then
		strReqAar = "0"
		else
		strReqAar = Request("year")
		end if	
	else
	strReqAar = right(year(now), 2)
	end if
	
	
	'thisjoblog = "j"
	thisfile = "stat_pies"
	if request("print") <> "j" then%>
	<!--#include file="inc/joblog_z_mrdk.asp"-->
	<%
	end if
	%>
	<!--include file="inc/stat_submenu.asp"-->
	<%
	if request("print") <> "j" then
	pleft = 190
	ptop = 55
	else
	pleft = 0
	ptop = 0
	end if
	%>	
	<div id="sindhold" style="position:absolute; left:<%=pleft%>; top:<%=ptop%>; width:60%; height:600; visibility:visible; z-index:100;">
	<%if request("print") = "j" then%>
	<table cellspacing="0" cellpadding="0" border="0" width="880">
	<tr>
		<td bgcolor="#003399" width="650"><img src="../ill/logo_topbar_print.gif" alt="" border="0"></td>
		<td bgcolor="#FFFFFF" align=right><a href="javascript:window.close()"><img src="../ill/luk_xp.gif" width="30" height="28" alt="" border="0">&nbsp;Luk</a><img src="../ill/blank.gif" width="30" height="1" alt="" border="0"><a href="javascript:window.print()"><img src="../ill/print_xp.gif" width="28" height="30" alt="" border="0">&nbsp;Print</a><img src="../ill/blank.gif" width="30" height="1" alt="" border="0"></td>
	</tr>
	</table>
	<%
	else%>
	<table border=0 cellpadding=0 cellspacing=0 width="600">
	<tr>
	<td valign="top" width="163"><img src="../ill/logo_bg.gif" width="163" height="53" alt="" border="0"></td>
	<td valign="top"><b>Top 5 job og % fordeling. </b><br>
	Fordelt på <b>de valgte medarbejdere</b> i den aktuelle kørsel. Summeret på samtlige <b>job</b> uanset hvilket der er valgt i denne kørsel.</td>
	</tr>
	</table>
	<%end if%>
	<table cellspacing="0" cellpadding="0" border="0" width="650">
	<tr>
    <td valign="top">
	<!-------------------------------Sideindhold------------------------------------->
<%
'*** Finder de medarbejdere og job der er valgt ***
Dim intMedArbVal 
Dim b
Dim intJobnrKriValues 
Dim i
'Redim intTotalthis(0)
'Redim jobnavn(0)
Redim intTotalall(0)
Redim Mnavn(0)
Redim Mnr(0)
'public jobMedarbKri

QUOT = Chr(34)
CRLF = Chr(13) & Chr(10)

'*** create the pie chart ***
iLineCnt = 3
iDegPos = 0
iRectTop = -110

select case strReqMrd
case 0
datekri_start = "1/1/"&strReqAar
dateKri_slut = "1/1/"&(strReqAar+1)
case 12
datekri_start = "1/"&strReqMrd&"/"&strReqAar
dateKri_slut = "1/1/"&(strReqAar+1)
case else
datekri_start = "1/"&strReqMrd&"/"&strReqAar
dateKri_slut = "1/"&(strReqMrd+1)&"/"&strReqAar
end select




'**** Skift medarbejder *****
if request("print") <> "j" then%>
<div style="position:absolute; left:647; top:93; width:300;">
<a href="javascript:NewWin('<%=thisfile%>.asp?menu=stat&print=j&mrd=<%=strReqMrd%>&jobnr=<%=intJobnr%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>')" target="_self" class='rmenu'>&nbsp;Printer venlig version&nbsp;<img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a>
</div>
<div style="position:absolute; left:647; top:120;">
<table cellspacing="0" cellpadding="0" border="0" width="150">
<tr bgcolor="#5582d2">
	<td height="20">&nbsp;&nbsp;<font class="stor-hvid">Skift medarbejder:</td>
</tr>
<tr bgcolor="#EFF3FF">
	<td valign=top style="border-right: 1px solid #003399; border-left: 1px solid #003399; padding-right:5; padding-left:5; padding-top:2; padding-bottom:10;">
<%
end if
datekri_start = convertDateYMD(datekri_start)
datekri_slut = convertDateYMD(datekri_slut)


			
			nomore = 0 
			onefound = 0
			jobMedarbKri = ""
			intMedArbVal = Split(selmedarb, ", ")
			For b = 0 to Ubound(intMedArbVal)
				strSQL = "SELECT sum(timer) AS antTimer, Tmnavn, Tmnr FROM timer WHERE Tmnr = " & intMedArbVal(b) &" AND tdato BETWEEN '"& datekri_start &"' AND '"& datekri_slut &"' AND tfaktim <> 5 GROUP BY TMnr ORDER BY antTimer DESC"
				oRec.open strSQL, oConn, 3
				Redim preserve intTotalall(b)
				Redim preserve Mnavn(b)
				Redim preserve Mnr(b)
				
				if not oRec.EOF then
				intTotalall(b) = oRec("antTimer")
				Mnr(b) = oRec("Tmnr")
				Mnavn(b) = oRec("Tmnavn")
				if nomore = 0 then
				useb = b
				nomore = 1
				end if
				
				if request("print") <> "j" then%>
				<a href="stat_pies.asp?b=<%=b%>&menu=stat&weekselector=<%=weekselector%>&strWSelStartDato=<%=strWSelStartDato%>&strWSelEndDato=<%=strWSelEndDato%>&mrd=<%=strReqMrd%>&jobnr=<%=intJobnr%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>" class='rmenu'><%=Mnavn(b)%>&nbsp;&nbsp;<%=kommaFunc(intTotalall(b))%></a><br>
				<%end if
				
				
				onefound = 1
				else
				intTotalall(b) = 0
				Mnr(b) = 0
				Mnavn(b) = "-"
				end if
				oRec.close
			next
			
			
			if request("print") <> "j" then%>
			</td></tr>
			<tr bgcolor="#003399">
				<td height="1"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
			</tr>
			</table>
			</div><%end if

			
			'** Finder den valgte medarbejder
			if len(request("b")) <> 0 then
			b = request("b")
			else
			b = useb
			end if

			
			'*** Udskriver Pies  ****
			public x
			public varPieces
			public varPiecesText 
			public intTotalthis 
			public jobnavn
			public jobnr
			public medarbnavn
			public iRectTop
			
			
			topdiv = 250
			toppos = 170
			
			public function pieces(medarbnr, b)
			varPieces = ""
			varPiecesText = ""
			use_medarbnr = medarbnr
			count = b
		
			i = 0
			x = 10
			
				strSQL = "SELECT Tjobnr, Tjobnavn, sum(timer) AS antTimer, Tmnavn FROM timer WHERE Tmnr = " & Mnr(b) &" AND tdato BETWEEN '"& datekri_start &"' AND '"& datekri_slut &"' AND tfaktim <> 5 GROUP BY Tjobnavn ORDER BY antTimer DESC"
				
				oRec.open strSQL, oConn, 3
					
					while not oRec.EOF AND i < 5
					intTotalthis = oRec("antTimer")
					jobnavn = oRec("Tjobnavn")
					jobnr = oRec("Tjobnr")
					medarbnavn = oRec("Tmnavn")
					
					
					select case i
					case 0
					fillcolor_1 = 255
					fillcolor_2 = 55
					fillcolor_3 = 0
					case 1
					fillcolor_1 = 55
					fillcolor_2 = 142
					fillcolor_3 = 25
					case 2
					fillcolor_1 = 48
					fillcolor_2 = 162
					fillcolor_3 = 204
					case 3
					fillcolor_1 = 0
					fillcolor_2 = 153
					fillcolor_3 = 255
					case 4
					fillcolor_1 = 204
					fillcolor_2 = 255
					fillcolor_3 = 204 
					end select
					 
						
						
						varPieces = varPieces & "<PARAM NAME='Line00"&x&"' VALUE='SetFillColor(192, 192, 192'>"
						varPieces = varPieces & "<PARAM NAME='Line00"&x+1&"' VALUE='SetLineStyle(0)'>"
						varPieces = varPieces & "<PARAM NAME='Line00"&x+2&"' VALUE='Rect(100, "& CStr(iRectTop) + 5 &", 20, 15, 0)'>"
						varPieces = varPieces & "<PARAM NAME='Line00"&x+3&"' VALUE='SetLineStyle(1)'>"
						varPieces = varPieces & "<PARAM NAME='Line00"&x+4&"' VALUE='SetFillColor("&fillcolor_1&", "&fillcolor_2&", "&fillcolor_3&")'>"
						varPieces = varPieces & "<PARAM NAME='Line00"&x+5&"' VALUE='Pie(-140, -100, 200, 200, "& CStr(iDegPos) &", "& CStr(Round(intTotalthis/intTotalall(b)*360)) &", 0)'>"
						varPieces = varPieces & "<PARAM NAME='Line00"&x+6&"' VALUE='Rect(95, "& CStr(iRectTop) &", 20, 15, 0)'>"
						
						Response.write varPieces
						
						varPiecesText = varPiecesText & "<DIV STYLE='font-family:Tahoma,Verdana,Arial; position:absolute; top:"& topdiv + iRectTop + 35 &"; left:395'>"
						varPiecesText = varPiecesText & "<B>"& left(jobnavn, 14) &"</B>&nbsp;Count:&nbsp;&nbsp;"& intTotalthis &"&nbsp;timer&nbsp;&nbsp;(<B>"& FormatPercent(intTotalthis / intTotalall(b), 1) &"</B>)</DIV>"
					
						
					iDegPos = iDegPos + CInt(intTotalthis / intTotalall(b) * 360)
					iRectTop = iRectTop + 25
					x = x + 7 
			
					
					i = i + 1
					oRec.movenext
					wend
					
				oRec.close
				'*** % tal ***
				topdiv = topdiv + 351
			end function
			
			
			if onefound <> 0 then%>
				<DIV STYLE="font-family:Tahoma,Verdana,Arial; position:absolute; top:<%=toppos - 30%>; left:165">
				<b><%=Mnavn(b)%>:</b>&nbsp;&nbsp;<%=kommaFunc(intTotalall(b))%> timer i 
				<%if strReqMrd > 0 then
				wrtmonthnm = strReqMrd %>
				<%=monthname(wrtmonthnm)%> måned
				<%end if%> 
				<%=strReqAar%>.
				</div>
				<OBJECT ID="sgPie_<%=b%>" STYLE="position:absolute; height:230; width:300; top:<%=toppos%>; left:120" CLASSID="CLSID:369303C2-D7AC-11D0-89D5-00A0C90833E6">
				<PARAM NAME="Line0001" VALUE="SetLineColor(0, 0, 0)"> 
				<PARAM NAME="Line0002" VALUE="SetLineStyle(0)">
				<PARAM NAME="Line0003" VALUE="SetFillColor(192, 192, 192">
				<PARAM NAME="Line0004" VALUE="Oval(-130, -90, 200, 200, 0)">
				
				<param name='Line0005' value='Rect(100, <%= CStr(iRectTop) + 5 %>, 20, 15, 0)'>
				<param name='Line0006' value='SetLineStyle(1)'>
				<PARAM NAME='Line0007' VALUE='SetFillColor(253, 153, 0)'>
				<PARAM NAME='Line0008' VALUE='Pie(-140, -100, 200, 200, <%= CStr(iDegPos) %>, <%= CStr(Round(intTotalthis / intTotalall(b) * 360)) %> , 0)'>
				<PARAM NAME='Line0009' VALUE='Rect(95, <%= CStr(iRectTop) %>, 20, 15, 0)'>
				
				<%call pieces(Mnr(b), b)%> 
				</OBJECT>	
				<%
				Response.write varPiecesText
				'** Pies **
				toppos = toppos + 350
			end if
			
		if len(varPieces) = 0 then
		Response.write "<br><br><br><br><img src='../ill/alert.gif' width='44' height='45' alt='' border='0'>Der er ikke fundet nogen registreringer i den valgte periode!"
		End if
		%>		
<br><br>&nbsp;</td></tr></table>
<%if onefound <> 0 then%>
<table cellspacing="0" cellpadding="0" border="0" width="100%">
<tr>
<td>
<%
intTotalInt = 0
intTotal = 0
intTotalEks = 0

strSQL = "SELECT timer AS antTimerEks FROM timer WHERE Tmnr = " & Mnr(b) &" AND tdato BETWEEN '"& datekri_start &"' AND '"& datekri_slut &"' AND tfaktim = 1"
oRec.open strSQL, oConn, 3

while not oRec.EOF 
intTotalEks = intTotalEks + Cint(oRec("antTimerEks"))

oRec.movenext
wend

oRec.close

intTotalEksNotfakbar = 0

strSQL = "SELECT timer AS antTimerEks FROM timer WHERE Tmnr = " & Mnr(b) &" AND tdato BETWEEN '"& datekri_start &"' AND '"& datekri_slut &"' AND tfaktim = 2"
oRec.open strSQL, oConn, 3

while not oRec.EOF 
intTotalEksNotfakbar = intTotalEksNotfakbar + Cint(oRec("antTimerEks"))

oRec.movenext
wend

oRec.close

intTotalInt = 0

strSQL = "SELECT timer AS antTimerInt FROM timer WHERE Tmnr = " & Mnr(b) &" AND tdato BETWEEN '"& datekri_start &"' AND '"& datekri_slut &"' AND tfaktim = 0"
oRec.open strSQL, oConn, 3

while not oRec.EOF 
intTotalInt = intTotalInt + Cint(oRec("antTimerInt"))

oRec.movenext
wend

oRec.close
Set oConn = Nothing

intTotal = intTotalEks + intTotalInt + intTotalEksNotfakbar

'Response.write "intTotalEks: " & intTotalEks & " Timer<br>"
'Response.write "intTotalInt: " & intTotalInt & " Timer<br>"
%>

<%

blnIE4 = False
strUA = Request.ServerVariables("HTTP_USER_AGENT")
intMSIE = Instr(strUA, "MSIE ") + 5
If intMSIE > 5 Then
  If CInt(Mid(strUA, intMSIE, Instr(intMSIE, strUA, ".") - intMSIE)) >= 4 Then 
    blnIE4 = True
  End If 
End If


If blnIE4 Then

'*** create the pie chart ***
iLineCnt = 3
iDegPosAll = 0
iRectTopAll = -100
%>

<OBJECT ID=sgPie_all STYLE="position:absolute; height:230; width:300; top:420; left:119" CLASSID="CLSID:369303C2-D7AC-11D0-89D5-00A0C90833E6">
<PARAM NAME="Line0001" VALUE="SetLineColor(0, 0, 0)"> 
<PARAM NAME="Line0002" VALUE="SetLineStyle(0)">
<PARAM NAME="Line0003" VALUE="SetFillColor(192, 192, 192">
<PARAM NAME="Line0004" VALUE="Oval(-130, -90, 200, 200, 0)">
<!-- Intern pie -->
<PARAM NAME="Line0005" VALUE="Rect(100, <%= CStr(iRectTopAll) + 5 %>, 20, 15, 0)">
<PARAM NAME="Line0006" VALUE="SetLineStyle(1)">
<PARAM NAME="Line0007" VALUE="SetFillColor(153, 204, 48)">
<PARAM NAME="Line0008" VALUE="Pie(-140, -100, 200, 200, <%= CStr(iDegPosAll) %>, <%= CStr(Round(intTotalInt / intTotal * 360)) %>, 0)">
<PARAM NAME="Line0009" VALUE="Rect(95, <%= CStr(iRectTopAll) %>, 20, 15, 0)">
<%
iDegPosAll = iDegPosAll + CInt(intTotalInt / intTotal * 360)
iRectTopAll = iRectTopAll + 25 
%>
<PARAM NAME="Line0010" VALUE="SetFillColor(192, 192, 192">
<PARAM NAME="Line0011" VALUE="SetLineStyle(0)">
<PARAM NAME="Line0012" VALUE="Rect(100, <%= CStr(iRectTopAll) + 5 %>, 20, 15, 0)">
<PARAM NAME="Line0013" VALUE="SetLineStyle(1)">
<PARAM NAME="Line0014" VALUE="SetFillColor(253, 153, 0)">
<PARAM NAME="Line0015" VALUE="Pie(-140, -100, 200, 200, <%= CStr(iDegPosAll) %>, <%= CStr(Round(intTotalEks/intTotal*360)) %>, 0)">
<PARAM NAME="Line0016" VALUE="Rect(95, <%= CStr(iRectTopAll) %>, 20, 15, 0)">
<%

iDegPosAll = iDegPosAll + CInt(intTotalEks/ intTotal * 360)
iRectTopAll = iRectTopAll + 25 
%>
<!--Procent indeks forklaring -->
<PARAM NAME="Line0017" VALUE="SetFillColor(192, 192, 192">
<PARAM NAME="Line0018" VALUE="SetLineStyle(0)">
<PARAM NAME="Line0019" VALUE="Rect(100, <%= CStr(iRectTopAll) + 5 %>, 20, 15, 0)">
<PARAM NAME="Line0023" VALUE="Rect(95, <%= CStr(iRectTopAll) %>, 20, 15, 0)">
<!--Pie -->
<PARAM NAME="Line0020" VALUE="SetLineStyle(1)">
<PARAM NAME="Line0021" VALUE="SetFillColor(6, 76, 186)">
<PARAM NAME="Line0022" VALUE="Pie(-140, -100, 200, 200, <%= CStr(iDegPosAll) %>, <%= CStr(Round(intTotalEksNotfakbar/intTotal*360)) %>, 0)">
</OBJECT>

<DIV STYLE="font-family:Tahoma,Verdana,Arial,sans-serif; position:absolute; top:436; left:395; width:400;">
<B>Interne Job</B>  &nbsp; Count:  &nbsp;(<B><%= FormatPercent(intTotalInt / intTotal, 1) %></B>)
</DIV>
<DIV STYLE="font-family:Tahoma,Verdana,Arial,sans-serif; position:absolute; top:461; left:395; width:400;">
<B>Fakturerbare Job</B>  &nbsp; Count:  &nbsp;(<B><%= FormatPercent(intTotalEks / intTotal, 1) %></B>)
</DIV>
<DIV STYLE="font-family:Tahoma,Verdana,Arial,sans-serif; position:absolute; top:486; left:395; width:400;">
<B>Ikke fakturerbare Job</B>  &nbsp; Count: &nbsp;(<B><%= FormatPercent(intTotalEksNotfakbar / intTotal, 1) %></B>)
</DIV>
<%end if%>
</td>
</tr>
</table>
<%end if%>
<br><br><br><br>
</div>

<%end if 'validering
end if %>
<!--#include file="../inc/regular/footer_inc.asp"-->