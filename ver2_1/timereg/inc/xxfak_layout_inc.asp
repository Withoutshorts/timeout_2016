<%
'********************************
'*** faktura layout opbygning ***
'********************************

'***** Fak header ******
public strTypeThis 
public strTypeThis2
public pdfHTML

function fakheader() 
%>
<table cellspacing="0" cellpadding="0" border="0" width="600">
		<tr><td valign="top" style="padding:10px 10px 0px 10px;">
	
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
		
		
		<h3><%=strTypeThis2%> - <%=varFaknr%></h3>
		<!--<img src="../ill/<=strtopimg%>.gif" width="271" height="25" alt="" border="0">-->
		</td></tr>
</table>
<%
end function

'***** Modtager boks ********
function modtager_layout()
%>
<table cellspacing="0" cellpadding="0" border="0" width=330 bgcolor="#F5f5f5">

<tr bgcolor="#ffffe1">
	
	<td valign="top" style="padding:10px 10px 10px 10px; border:1px #5582d2 solid;"><font size=1 color="#999999">Afsender:&nbsp;<%=yourNavn%>&nbsp;&nbsp;<%=yourAdr%>&nbsp;&nbsp;<%=yourPostnr%>&nbsp;<%=yourCity%><br></font>
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
	</td>
	</tr>

</table>		
<%		
end function


'***** maintable **********
sub maintable
%>
<table cellspacing="0" cellpadding="0" border="0">
	<tr>
		
		<td colspan="5" height="10">&nbsp;</td>
		
	</tr>
	<tr>
		<td style="padding-left : 0px;" valign="top" colspan="2">
<%
end sub

'**** maintable 2 *********
sub maintable_2
%>
</td>
<td width="40"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
<td style="padding-right:0px;" valign="top" colspan="2">
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
<table cellspacing="0" cellpadding="0" border="0" width=250>

<tr>
	<td valign="top" bgcolor="#ffffff" style="padding:10px 10px 10px 10px; border:1px silver solid;"><%=yourNavn%><br>
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
</tr>
</table>
<%
end function

'****** Vedr. ***************
sub vedr
%>
<table cellspacing="0" cellpadding="0" border="0" width="620">
<tr>
	<td valign='top' height='40'>&nbsp;</td>
	<td colspan=7 style="padding:15px 15px 15px 0px;">Vedrørende: 
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
function udspecificering(strFakdet,aktmat)

'if aktmat = 1 then
'thisaktmat = "Aktiviteter og timer"
'else
'Response.Write "<br><br>"
'thisaktmat = "Materialer"
'end if%>

<!--<img src="../ill/blank.gif" width="10" height="2" border="0" /><=thisaktmat %>-->
<table cellspacing="0" cellpadding="0" border="0" width="620">
<tr>
	
	<td width=40 align=right style="padding-right:2px;"><b>Antal</b></td>
	<td colspan="2" width="230" style="padding-left:5px;"><b>Beskrivelse</b></td>
	<td align=right width="130" style='padding-right:10px;'><b>Enhedspris</b></td>
	<%if cint(visrabatkol) = 1 then %>
	<td align=right width="100" style='padding-right:10px;'><b>Rabat</b></td>
	<%end if %>
	<td align=right width="120"><b>Pris</b></td>
	
</tr>
<%=strFakdet%>
<tr>
</table>
<%
end function


'***** Sub-Total aktiviteter***

function subtotakt()
%>
<table cellspacing="0" cellpadding="0" border="0" width="620">
<tr>
	<td valign='top' width=300 style="padding:10px 50px 0px 14px;">Antal:&nbsp;<b><%=formatnumber(intFaktureretTimer, 2)%></b></td>
	<td valign='top' align=right style="padding:10px 0px 0px 0px;">Subtotal:&nbsp;&nbsp;<b><%=formatcurrency(aktsubtotal)%></b></td>
</tr>
</table>
<br />
<%
end function

'***** Sub-Total materialer***

function subtotmat()
%>
<table cellspacing="0" cellpadding="0" border="0" width="620">
<tr>
    <td valign='top' align=right style="padding:10px 0px 0px 0px;">Subtotal:&nbsp;&nbsp;<b><%=formatcurrency(matsubtotal)%></b></td>
</tr>
</table>
<%
end function

'***** Totaler og moms ***
function totogmoms()
%>
<table cellspacing="0" cellpadding="0" border="0" width="620">
<tr>
    <td valign='top' align=right style="padding:30px 0px 0px 0px;">Total:&nbsp;&nbsp;<b><%=formatcurrency(intFaktureretBelob)%></b></td>
</tr>
<tr>
	<td valign='top' align=right style="padding:0px 0px 0px 0px;">Moms:&nbsp;&nbsp;
	<%=formatcurrency(intMoms,2)%>
	<br>-------------------------------------------------</td>
	
</tr>
<tr>
	
	<td valign='top' align=right style="padding:0px 0px 0px 0px;">Total incl. moms:&nbsp;&nbsp;
	<%
	totalinclmoms = (intMoms + intFaktureretBelob)
	%>
	<b><%=formatcurrency(totalinclmoms,2)%></b><br />
	=================================</td>
	
</tr>
</table>
<%
end function


'*** Komm. og bet. betingelser ***
public toph, h

sub komogbetbetingelser	
if h < 720 then
toph = 850
else
toph = h + 100
end if
%>
<!--<div id="fakfooter" style="position:absolute; left:10px; top:<%=toph%>px; visibility:visible;">-->
<br /><br /><br /><br />
<table cellspacing="0" cellpadding="0" border="0" width="620">
<tr>
<td><b>Kommentar og betalings betingelser vedr. denne <%=strTypeThis%>.</b><br><%=strKomm%><br>&nbsp;</td>
</tr>
</table>
<!--
<table cellspacing="0" cellpadding="0" border="0" width="600">
<tr>
	<td colspan="7" valign="bottom"><img src="../ill/fakfooter.gif" width="600" height="42" alt="" border="0"></td>
</tr>
</table>-->
<br><br>&nbsp;
<!--</div>-->
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


function rykkergebyrer()

               %>
               <br /><br />
               <table cellspacing=0 cellpadding=0 border=0 width="620">
               <!--<tr><td style="padding:0px 0px 0px 30px;" colspan=4>Rykker gebyr</td></tr>-->
             
               
               <%
               strSQLryk = "SELECT rykkertxt, rykkerantal, rykkerbelob, rykkerdato FROM faktura_rykker WHERE fakid = "&  id
               oRec.open strSQLryk, oConn, 3
               r = 0
               while not oRec.EOF 
               
               if r = 0 then%>
                <tr>
               <td width=40 align=right style="padding-right:2px;"><b>Antal</b></td>
               <td width="250" style="padding-left:5px;"><b>Beskrivelse</b></td>
               <td><b>Dato</b></td>
               <td align=right><b>Pris</b></td>
               </tr>
               <%end if
              
               
               %><tr>
               <td align=right style="padding-right:2px;"><%=oRec("rykkerantal") %></td>
               <td style="padding-left:5px;"><%=oRec("rykkertxt") %></td>
               <td><%=formatdatetime(oRec("rykkerdato"), 1) %></td>
               <td align=right><%=formatcurrency(oRec("rykkerbelob")) %></td>
               </tr>
               <%
             
               intFaktureretBelob = intFaktureretBelob + oRec("rykkerbelob")
               
               r = r + 1 
               oRec.movenext
               wend
               oRec.close
               
               %></table><%

end function

%>
