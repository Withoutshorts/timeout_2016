<%
'********************************
'*** faktura layout opbygning ***
'********************************

'***** Fak header ******
public strTypeThis 
public strTypeThis2

function fakheader() 
%>
<table cellspacing="0" cellpadding="0" border="0" width="600">
		<tr><td>
	
		<%select case intFaktype
		case 0
		strtopimg = "faktura_top"
		strTypeThis = "faktura"
		strTypeThis2 = "Faktura"  
		case 1
		strtopimg = "kreditnota_top"
		strTypeThis = "kreditnota"
		strTypeThis2 = "Kreditnota"   
		case 2
		strtopimg = "rykker_top"
		strTypeThis = "rykker"
		strTypeThis2 = "Rykker"  
		case else
		strtopimg = "blank"
		end select%>
		
		<img src="../ill/<%=strtopimg%>.gif" width="271" height="25" alt="" border="0">
		</td></tr>
</table>
<%
end function

'***** Modtager boks ********
function modtager_layout()
%>
<table cellspacing="0" cellpadding="0" border="0" bgcolor="#F5f5f5">
<tr>
	<td valign="top"><img src="../ill/tabel_top_left_wsmoke.gif" width="8" height="23" alt="" border="0"></td>
	<td valign="top"><img src="../ill/tabel_top.gif" width="300" height="1" alt="" border="0"></td>
	<td valign="top"><img src="../ill/tabel_right_left_wsmoke.gif" width="8" height="23" alt="" border="0"></td>
</tr>
<tr bgcolor="#f5f5f5">
	<td width="1"><img src="../ill/tabel_top.gif" width="1" height="180" alt="" border="0"></td>
	<td valign="top"><font size=1 color="#999999">Afsender:&nbsp;<%=yourNavn%>&nbsp;&nbsp;<%=yourAdr%>&nbsp;&nbsp;<%=yourPostnr%>&nbsp;<%=yourCity%><br></font>
	<%=strKnavn%><br>
	<%=strKadr%><br>
	<%=strKpostnr%>&nbsp;&nbsp;<%=strBy%><br>
	
	<%if intVismodland = 1 then
	Response.write strLand &"<br>"
	end if
	
	if intVismodatt = 1 then
	Response.write "<b>Att: "& showAtt &"</b><br>"
	end if 
	%>
	
	<%if intVismodtlf = 1 then
	Response.write "<br>Tlf: "& intTlf
	end if%>
	
	<%if intVismodcvr = 1 then
	Response.write "<br>Cvr/SE nr: " & intCvr 
	end if%>
	
	<br><br>
	<font class="megetlillesort">Kundenr:&nbsp;<%=intKnr%></font>
	<br>
	<%=formatdatetime(fakdato, 1)%><br>
	<%=strTypeThis2%> nr:&nbsp;<b><%=varFaknr%></b>
	<td width="8" align="right"><img src="../ill/tabel_top.gif" width="1" height="180" alt="" border="0"></td>
</tr>
<tr>
	<td valign="bottom"><img src="../ill/tabel_bund_left_wsmoke.gif" width="8" height="23" alt="" border="0"></td>
	<td valign="bottom"><img src="../ill/tabel_top.gif" width="300" height="1" alt="" border="0"></td>
	<td valign="bottom"><img src="../ill/tabel_bund_right_wsmoke.gif" width="8" height="23" alt="" border="0"></td>
</tr>
</table>		
<%		
end function


'***** maintable **********
sub maintable
%>
<table cellspacing="0" cellpadding="0" border="0" width="600" bgcolor="#FFFFFF">
	<tr>
		<td valign="top" rowspan=2 width="10">&nbsp;</td>
		<td colspan="5" height="10">&nbsp;</td>
		<td valign="top" align=right rowspan=2>&nbsp;</td>
	</tr>
	<tr>
		<td style="padding-left : 4px;" valign="top" colspan="2">
<%
end sub

'**** maintable 2 *********
sub maintable_2
%>
</td>
<td width="150"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
<td style="padding-right : 25px;" valign="top" colspan="2">
<%
end sub


'**** maintable 3 *********
sub maintable_3
%>
</td>
</tr>
</table>
<%
end sub


'****** Afsender boks *******
function afsender_layout()
%>
<table cellspacing="0" cellpadding="0" border="0" bgcolor="#F5f5f5">
<tr>
	<td valign="top"><img src="../ill/tabel_top_left_wsmoke.gif" width="8" height="23" alt="" border="0"></td>
	<td valign="top"><img src="../ill/tabel_top.gif" width="200" height="1" alt="" border="0"></td>
	<td valign="top"><img src="../ill/tabel_right_left_wsmoke.gif" width="8" height="23" alt="" border="0"></td>
</tr>
<tr bgcolor="#f5f5f5">
	<td width="1"><img src="../ill/tabel_top.gif" width="1" height="180" alt="" border="0"></td>
	<td valign="top"><%=yourNavn%><br>
	<%=yourAdr%><br>
	<%=yourPostnr%>&nbsp;<%=yourCity%><br>
	<%=yourLand & "<br>"%>
	
	<%if intVisafstlf = 1 then
	Response.write "<br>Tlf: "& yourTlf
	end if%>
	
	<%if intVisafsemail = 1 then
	Response.write "<br>Email: " & yourEmail
	end if%>
	
	<br><br>
	Bank:&nbsp;<%=yourBank%><br>
	Reg. og kontonr: <%=yourRegnr%> - <%=yourKontonr%><br>
	
	<%if intVisafsswift = 1 then%>
	Swift:&nbsp;<%=yourSwift%><br>
	<%end if%>
	
	<%if intVisafsiban = 1 then%>
	Iban:&nbsp;<%=yourIban%><br>
	<%end if%>
	
	<%if intVisafscvr = 1 then%>
	Cvr/SE nr: <%=yourCVR%>
	<%end if%></td>
	<td width="8" align="right"><img src="../ill/tabel_top.gif" width="1" height="180" alt="" border="0"></td>
</tr>
<tr>
	<td valign="bottom"><img src="../ill/tabel_bund_left_wsmoke.gif" width="8" height="23" alt="" border="0"></td>
	<td valign="bottom"><img src="../ill/tabel_top.gif" width="200" height="1" alt="" border="0"></td>
	<td valign="bottom"><img src="../ill/tabel_bund_right_wsmoke.gif" width="8" height="23" alt="" border="0"></td>
</tr>
</table>
<%
end function

'****** Vedr. ***************
sub vedr
%>
<table cellspacing="0" cellpadding="0" border="0" width="600" bgcolor="#FFFFFF">
<tr>
	<td valign='top' height='40'>&nbsp;</td>
	<td colspan=7 style="padding:15px;">Vedrørende: 
	<%if jobid <> 0 then%>
	<b><%=strJobnavn%>&nbsp;(<%=intJobnr%>)</b><br>
		<%if len(trim(strJobBesk)) <> 0 then%>
		<%=strJobBesk%><br><br>
		<%end if%>
	<%
	else
	%>
	<b><%=strAftNavn%>&nbsp;(<%=intAftnr%>)</b><br>
	Varenr: <%=strAftVarenr%>
	<%
	end if
	%>
	</td>
	<td valign='top' align=right>&nbsp;</td>
</tr>
</table>
<%
end sub



'****** Udspecificering *******
function udspecificering(strFakdet)
%>
<table cellspacing="0" cellpadding="0" border="0" width="600" bgcolor="#FFFFFF">
<tr>
	<td valign='top'>&nbsp;</td>
	<td width="20"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td width=40 align=right style="padding-right:2;"><b>Antal</b></td>
	<td colspan="2" width="530" style="padding-left:5;"><b>Beskrivelse</b></td>
	<td align=right width="100" style='padding-right:30;'><b><%=enhed%></b></td>
	<td align=right colspan=2 style='padding-right:30;'><b>Pris</b></td>
	<td valign='top' align=right>&nbsp;</td>
</tr>
<%=strFakdet%>
<tr>
</table>
<%
end function


'***** Totaler og moms ***
function totogmoms()
%>
<table cellspacing="0" cellpadding="0" border="0" width="600" bgcolor="#FFFFFF">
<tr>
	<td valign='top'>&nbsp;</td>
	<td width="20"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td valign='top' width=300 style="padding-top:30; padding-right:8px;">Total antal:&nbsp;<b><%=formatnumber(intFaktureretTimer, 2)%></b></td>
	<td valign='top' align=right style="padding-top:30;">Subtotal:&nbsp;&nbsp;</td>
	<td valign='top' align=right style="padding-top:30; padding-right:8px;">&nbsp;&nbsp;<b><%=formatcurrency(intFaktureretBelob)%></b></td>
	<td width="20"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td valign='top' align=right>&nbsp;</td>
</tr>
<tr>
	<td valign='top'>&nbsp;</td>
	<td width="20"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0">
	<td valign='top'>&nbsp;</td>
	<td valign='top' align=right>Moms:&nbsp;&nbsp;</td>
	<td valign='top' align=right style="padding-right:8px;">&nbsp;&nbsp;<b>
	<%
	Response.write formatcurrency(intMoms,2)%></b></td>
	<td width="20"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td valign='top' align=right>&nbsp;</td>
</tr>
<tr>
	<td valign='top'>&nbsp;</td>
	<td width="20"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0">
	<td valign='top'>&nbsp;</td>
	<td valign='top' align=right><br>Total:&nbsp;&nbsp;</td>
	<td valign='top' align=right style="padding-right:8px;">&nbsp;&nbsp;____________<br>
	<%
	totalinclmoms = (intMoms + intFaktureretBelob)
	%>
	&nbsp;&nbsp;<b><%=formatcurrency(totalinclmoms,2)%></b><br>
	&nbsp;&nbsp;____________</td>
	<td width="20"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	<td valign='top' align=right>&nbsp;</td>
</tr>
</table>
<%
end function


'*** Komm. og bet. betingelser ***
sub komogbetbetingelser	
if h < 720 then
toph = 820
else
toph = h + 100
end if
%>
<div id="fakfooter" style="position:absolute; left:10; top:<%=toph%>; visibility:visible;">
<table cellspacing="0" cellpadding="0" border="0" width="600" bgcolor="#FFFFFF">
<tr>
	<td valign='top'>&nbsp;</td>
	<td width="20" valign="top"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0">
	<td colspan="4"><b>Kommentar og betalings betingelser vedr. denne <%=strTypeThis%>.</b><br><%=strKom%><br>&nbsp;</td>
	<td valign='top' align=right>&nbsp;</td>
</tr>
</table>
<!--
<table cellspacing="0" cellpadding="0" border="0" width="600">
<tr>
	<td colspan="7" valign="bottom"><img src="../ill/fakfooter.gif" width="600" height="42" alt="" border="0"></td>
</tr>
</table>-->
<br><br>&nbsp;
</div>
<%
end sub



'**** Brevhoved *************

'**** Brevfod ***************



'**** Link (højreside), ikke printbar ******
sub ikkeprintbar
%>
<div id="link" style="position:absolute; left:720; top:60; visibility:visible;">
<table cellspacing="0" cellpadding="0" border="0" width="240">
<tr>
	<td colspan="2" valign="bottom">
	<a href="Javascript:window.print()">Print&nbsp;<img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a><br><br>
	<%
	if request("histback") = 1 then%>
	<a href="Javascript:history.back()">Tilbage til faktura-oversigt&nbsp;<img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a>
	<%
	else
		if jobid <> 0 then%>
		<a href="fak_osigt.asp?menu=stat_fak&FM_job=<%=Request("FM_job")%>&FM_usedatointerval=<%=request("FM_usedatointerval")%>&FM_start_dag=<%=request("FM_start_dag_ival")%>&FM_start_mrd=<%=request("FM_start_mrd_ival")%>&FM_start_aar=<%=request("FM_start_aar_ival")%>&FM_slut_dag=<%=request("FM_slut_dag_ival")%>&FM_slut_mrd=<%=request("FM_slut_mrd_ival")%>&FM_slut_aar=<%=request("FM_slut_aar_ival")%>&FM_fakint=<%=request("FM_fakint_ival")%>">Tilbage til faktura-oversigt&nbsp;<img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a>
		<%else%>
		<a href="fak_serviceaft_osigt.asp?menu=stat_fak&id=<%=thiskid%>&aftaleid=<%=aftid%>&FM_aftnr=<%=sogaftnr%>">Tilbage til faktura-oversigt&nbsp;<img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a>
		<%end if%>
	<%
	end if%>
<br><br><br>
<b>Printeropsætning:</b><br>
Vælg portræt (Stående A4) for at printe fakturaen ud uden denne spalte.<br><br>
Indstil din browser til ikke at printe side- hoved og fod ud. Det gøres ved at klikke på menu punktet [Filer] (øverst til ventre) --> [Side opsætning].<br><br> Fjern teksten i sidehoved og sidefod.<br><br>
For at spare på farven i din printer patron, bør du også sørge for at printeren ikke er indstillet til at udskrive baggrundsfarver.
<br><br>
</td></tr>
<tr><td valign=top>
<b>Øvrige faktura oplysninger:</b><br>
Faktura er oprettet af: <br><%=strEditor%><br>
<%=formatdatetime(fakdato, 1)%> 
<br><br>

<b>Faktura type:</b><br>
<%=strTypeThis2%>
<br><br>

<b>Konti:</b><br>
Konto: <%=strKontoNavn%> (<%=intKontonr%>)<br>
Moms procent: <%=intKvotient%>

<br>
<br>Modkonto: <%=strModKontoNavn%> (<%=intModKontonr%>)<br><br>



<b>Følgende timer er lagt i databasen på hver enkelt medarbejder:</b><br>
<br>
<table cellspacing=0 cellpadding=0 border=0 width=240>
<%=strMedarbFakdet%>
</table>
<br><br>&nbsp;
</td></tr>
</table>
</div>
<%
end sub

%>
