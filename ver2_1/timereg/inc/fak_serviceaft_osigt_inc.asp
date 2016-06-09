<!--
Bruges af:
fak_serviceaft_osigt.asp
-->


		<%
		function saldo(enhThis, prisThis, tFakEnh, tFakBel, aftid, kid)
		%>
		<tr bgcolor="<%=bgt%>">
			<td valign="top" style="border-left:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
			<td colspan=4 align=right style="padding-right:5px; border-top:1px #999999 solid;"><b>Faktureret ialt:</b></td>
			<td align=right style="border-right:1px #999999 solid; border-top:1px #999999 solid; padding-right:5px;"><b><%=formatnumber(tFakEnh, 2)%></b></td>
			<td align=right style="border-right:1px #999999 solid; border-top:1px #999999 solid; padding-right:5px;"><b><%=formatnumber(tFakBel, 2)%></b></td>
			<%
			saldoEnheder = (tFakEnh - enhThis)
			saldoBelob = (tFakBel - prisThis)
			
			tFakEnh = 0
			tFakBel = 0
			%>
			<td align=right style="border-top:1px #999999 solid;">&nbsp;</td>
			<td>&nbsp;</td>
			<td valign="top" style="border-right:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		</tr>
		<tr bgcolor="<%=bgt%>">
			<td valign="top" style="border-left:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
			<td colspan=4 align=right style="padding-right:5px; padding-bottom:10px;"><b>Saldo:</b></td>
			<td align=right style="border-right:1px #999999 solid; padding-right:5px; padding-bottom:10px;"><b><%=formatnumber(saldoEnheder, 2)%></b></td>
			<td align=right style="border-right:1px #999999 solid; padding-right:5px; padding-bottom:10px;"><b><%=formatnumber(saldoBelob, 2)%></b></td>
			<%
			saldoEnheder = 0
			saldoBelob = 0
			%>
			<td colspan=2 align=right style="padding-bottom:10px;"><a href="fak_serviceaft_saldo.asp?menu=kon&aftid=<%=aftid%>&kundeid=<%=kid%>" class=vmenu>Udspecificering&nbsp;<img src="../ill/soeg-knap.gif" width="16" height="16" alt="" border="0"></a></td>
			<td valign="top" style="border-right:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		</tr>
		<%
		end function
		

		if len(request("filter_per")) <> 0 then
		filterKri = request("filter_per")
		response.cookies("fso_filterKri") = filterKri
		else
			if len(request.cookies("fso_filterKri")) <> 0 then
			filterKri = request.cookies("fso_filterKri")
			else
 			filterKri = 0
			end if 
		end if
		
		select case filterKri
		case "0"
		chkfilt0 = "CHECKED"
		chkfilt1 = ""
		chkfilt2 = ""
		useFilterKri = ""
		case "1"
		chkfilt0 = ""
		chkfilt1 = ""
		chkfilt2 = "CHECKED"
		useFilterKri = " AND erfaktureret < 1 "
		case "2"
		chkfilt0 = ""
		chkfilt1 = "CHECKED"
		chkfilt2 = ""
		useFilterKri = " AND (erfaktureret = 1 OR erfaktureret = 2) "
		end select
		
		
		'**** valgt kunde / sog på aftnr ***
		if len(trim(request("FM_aftnr"))) <> 0 AND request("FM_aftnr") <> 0 then
			sogaftnr = request("FM_aftnr")
			showsogaftnr = sogaftnr
			thiskid = 0 
		else
			sogaftnr = 0
			showsogaftnr = ""
			thiskid = id
		end if
		
		%>
		<table border=0 cellspacing=0 cellpadding=0 bgcolor="#ffffff" width="400">
		<form id=filteraftnr1 name=filteraftnr1 method=post action="<%=thisfile%>?menu=stat_fak&FM_soeg=<%=FM_soeg%>&func=<%=func%>">
		<tr><td colspan=4 style="border-left:1px #8caae6 solid; border-top:1px #8caae6 solid; border-right:1px #8caae6 solid; padding-top:8px;">
		&nbsp;&nbsp;<b>Søg Aftale nr:</b> <input type="text" name="FM_aftnr" id="FM_aftnr" value="<%=showsogaftnr%>" style="width:100px;">&nbsp;
		&nbsp;&nbsp;<input type="submit" value="Søg"><br>
		&nbsp;&nbsp;<font class=megetlillesort>(ignorerer andre kriterier)</font>
		
		</form>
		<form id=filteraftnr2 name=filteraftnr2 method=post action="<%=thisfile%>?menu=stat_fak&FM_soeg=<%=FM_soeg%>&func=<%=func%>">
		
		<%btop = 0%>
		
		&nbsp;&nbsp;<b>Eller vælg Kontakt:</b>&nbsp;<select name="id" size="1" style="font-size : 9px; width:165 px;" onChange="rensaftnr();">
		<option value="0">Alle</option>
		<%
				strSQL = "SELECT s.kundeid, k.kkundenavn, k.kkundenr, k.kid FROM serviceaft s LEFT JOIN  kunder k ON (k.kid = kundeid AND k.ketype <> 'e') WHERE s.kundeid <> 0 GROUP BY s.kundeid ORDER BY k.kkundenavn"
				'Response.write strSQL
				'Response.flush
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(thiskid) = cint(oRec("Kid")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if
				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%></option>
				<%
				oRec.movenext
				wend
				oRec.close
				%>
				
		</select><br>&nbsp;
		</td></tr>
		<tr>
			<td colspan=4 style="border-left:1px #8caae6 solid; border-top:<%=btop%>px #8caae6 solid; border-right:1px #8caae6 solid;">
			<%if fmudato = 1 then
			usedkri = "CHECKED"
			usedKri2 = ""
			else
			usedkri = ""
			usedKri2 = "CHECKED"
			end if%>
			<input type="radio" name="FM_usedatokri" value="0" <%=usedkri2%>> Ignorér datointerval.&nbsp;&nbsp;
			<input type="radio" name="FM_usedatokri" value="1" <%=usedkri%>> Vis kun aftaler med <u>startdato</u> i det valgte interval.<br>&nbsp;</td>
		<tr>
			<td style="border-left:1px #8caae6 solid; border-right:1px #8caae6 solid; padding-left:20px;" colspan=4>&nbsp;
			<!--#include file="weekselector_s.asp"--> <!-- b -->
			</td>
		</tr>
		<tr>
			<td colspan=4 style="border-left:1px #8caae6 solid; border-bottom:1px #8caae6 solid; border-right:1px #8caae6 solid;"><br>&nbsp;&nbsp;<b>Eksport status:</b><br>
			<input type="radio" name="filter_per" id="filter_per" value="0" <%=chkfilt0%>>Vis Alle<br>
			<input type="radio" name="filter_per" id="filter_per" value="1" <%=chkfilt2%>>Vis kun <b>endnu ikke</b> eksporterede aftaler.<br>
			<input type="radio" name="filter_per" id="filter_per" value="2" <%=chkfilt1%>>Vis kun <b>allerede</b> eksporterede aftaler.
			<br><br>&nbsp;&nbsp;<input type="submit" value="Brug dette filter"><br>&nbsp;
			</td>
		</tr>
		</form>
		</table>
		<br><br>

		
		
		

<table cellspacing="0" cellpadding="0" border="0" bgcolor="#5582d2" width=800>
<form action="fak_serviceaft_kvik.asp?FM_aftnr=<%=sogaftnr%>&thiskid=<%=thiskid%>" method="post" name="kvik" id="kvik">
	<tr bgcolor="#5582D2">
		<td width="8" rowspan="2" style="border-top:1px #003399 solid; border-left:1px #003399 solid;" valign=top><img src="../ill/blank.gif" width="8" height="32" alt="" border="0"></td>
		<td colspan=8 valign="top" style="border-top:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" style="border-top:1px #003399 solid; border-right:1px #003399 solid;" valign=top><img src="../ill/blank.gif" width="8" height="32" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td colspan=8 class='alt'><b>Fakturering af Aftaler.</b></td>
	</tr>
	<tr bgcolor="#ffffff">
		<td valign="top" height=35 style="border-left:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td colspan=8 valign="top"><b>Til Kvik fakturering vælg:</b>&nbsp;&nbsp; <a href="#" name="CheckAll" onClick="checkAll(document.kvik.FM_aft)" class=vmenu>Alle</a>
		&nbsp;&nbsp;|&nbsp;&nbsp;<a href="#" name="UnCheckAll" onClick="uncheckAll(document.kvik.FM_aft)" class=vmenu>Ingen</a>
		<br>
		<font class=roed><b>(!)</b></font> = Denne aftale har allerede været eksporteret.
		</td>
		<td valign="top" style="border-right:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
</table>
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#999999" width=800>
	
	<tr bgcolor="ffffff">
		<td valign="top" height=20 style="border-left:1px #8caae6 solid;">&nbsp;</td>
		<td><b>Kunde</b> (id)</td>
		<td>&nbsp;</td>
		<td><b>Aftale navn</b> (nr)</td>
		<td><b>Varenr.</b></td>
		<td align=right><b>Enh. tildelt</b></td>
		<td align=right><b>Aft. pris</b></td>
		<td align=center>Periode</td>
		<td align=right><b>Opret faktura</b></td>
		<td valign="top" style="border-right:1px #8caae6 solid;">&nbsp;</td>
	</tr>
	<%
	'*** SQL datoer ***
	strStartDato = strAar&"/"&strMrd&"/"&strDag
	strSlutDato = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut
	
	totFakEnheder = 0
	totFakBel = 0
	enhederDenne = 0 
	prisDenne = 0
	
	if sogaftnr = 0 then
	
		'** Kundekri **
		if cint(thiskid) <> 0 then
		kidKri = " s.kundeid = "& thiskid & " "
		else
		kidKri = " s.kundeid <> 0 "
		end if
		
		'** Dato kri ***
		if fmudato = 1 OR thisfile = "joblog_k" then
		strDatoKri = " AND s.stdato BETWEEN '"& strStartDato &"' AND '"& strSlutDato &"'"
		else
		strDatoKri = ""
		end if
	
	else
		
		strDatoKri = ""
		kidKri = " s.kundeid <> 0 "
		useFilterKri = " AND s.aftalenr LIKE "& sogaftnr &""
	
	end if
	
	lastKid = 0
	
	strSQL = "SELECT s.id, s.enheder, s.stdato, s.sldato, s.status, s.navn, s.pris, s.perafg, "_
	&" s.advitype, s.advihvor, s.erfornyet, s.erfaktureret, s.varenr, "_
	&" k.kkundenavn, k.kkundenr, kid, s.aftalenr, f.fakdato, f.fid, f.faknr, s.besk AS note, "_
	&" f.betalt, f.timer AS fakenheder, f.beloeb AS fakbelob "_
	&" FROM serviceaft s"_
	&" LEFT JOIN kunder k ON (kid = kundeid) "_
	&" LEFT JOIN fakturaer f ON (f.aftaleid = s.id) "_
	&" WHERE "& kidKri &" "& useFilterKri &" "& strDatoKri &" GROUP BY s.id, f.fid ORDER BY k.kkundenavn, s.id DESC" 
	
	
	'Response.write strSQL
	'Response.flush
	
	aftalIds = "0"
	antalBrugteEnhederTot = 0
	antalTildelteEnhederTot = 0
	
	oRec.open strSQL, oConn, 3 
	s = 0
	while not oRec.EOF 
	
	kundeid = oRec("kid")
	
	'select case right(s, 1)
	'case 0, 2, 4, 6, 8
	bgt = "#ffffff"
	'case else
	'bgt = ""
	'end select
	
	
	if lastaftid <> oRec("id") then%>
	<%
	'*** Saldo ****
	if s <> 0 then
		call saldo(enhederDenne, prisDenne, totFakEnheder, totFakBel, lastaftid, kundeid)
	end if%>
	
	<tr bgcolor="#d6dff5">
		<td colspan=10 height=10 style="border-left:1px #8caae6 solid; border-right:1px #8caae6 solid; border-top:3px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	
	<tr bgcolor="#d6dff5">
		<td valign="top" style="border-left:1px #8caae6 solid;">&nbsp;</td>
		<td width=180><b><%=oRec("kkundenavn")%></b>&nbsp;(<%=oRec("kkundenr")%>)</td>
		<td><input type="checkbox" name="FM_aft" id="FM_aft" value="<%=oRec("id")%>"></td>
		
		<td>
		
		<%if oRec("erfaktureret") = 0 then '*** Aftale er endnu ikke eksporteret!! 
		aftalIds = aftalIds & ", " & oRec("id")
		else
		%>
		<font class=roed><b>(!)</b></font>
		<%
		end if%>
		
		<b><%=oRec("navn")%></b> (<%=oRec("aftalenr")%>)
		
		
		
		</td>
		<td><%=oRec("varenr")%>&nbsp;</td>
		<td align=right style="border-right:1px #999999 solid; padding-right:5px;"><b><%=formatnumber(oRec("enheder"))%></b></td>
		<td align=right style="border-right:1px #999999 solid; padding-right:5px;"><b><%=formatnumber(oRec("pris"))%></b></td>
		<td class=lille align=center>
		<%if oRec("perafg") = 0 then%>
		
				<%if len(oRec("stdato")) <> 0 then%>
						
						<%=formatdatetime(oRec("stdato"), 2)%>
					
						<%if len(oRec("sldato")) <> 0 then%>
					 	til <%=formatdatetime(oRec("sldato"), 2)%>
						<%end if%>
				<%
				end if
				%>
		<%else%>
				<%if len(oRec("stdato")) <> 0 then%>
				<%=formatdatetime(oRec("stdato"), 2)%>
				<%end if%> 
		<%end if
		
		'**** tilsaldo ****
		enhederDenne = oRec("enheder")
		prisDenne = oRec("pris")
		%></td>
		
		
		
		
		<td style="padding-top:3px;" align=right>
		<a href="fak.asp?menu=stat_fak&ttf=0&ktf=0&aftaleid=<%=oRec("id")%>&FM_aftnr=<%=sogaftnr%>&thiskid=<%=thiskid%>&FM_medarb=0&FM_usedatointerval=1&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd%>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>">
			<img src="../ill/ac0010-16.gif" width="16" height="16" alt="Opret faktura" border="0"></a>
		</td>
		<td valign="top" style="border-right:1px #8caae6 solid;">&nbsp;</td>
	</tr>
	<tr bgcolor="<%=bgt%>">
		<td height=1 valign="top" style="border-left:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td colspan=4 style="border-bottom:1px #999999 solid;"><br><b>Faktura nr.</b></td>
		<td align=right style="border-right:1px #999999 solid; border-bottom:1px #999999 solid; padding-right:5px;"><br><b>Faktureret enheder</b></td>
		<td align=right style="border-right:1px #999999 solid; border-bottom:1px #999999 solid; padding-right:5px;"><br><b>Faktureret beløb</b></td>
		<td align=right style="border-bottom:1px #999999 solid; padding-right:5px;"><br><b>Status/Fakdato</b></td>
		<td>&nbsp;</td>
		<td valign="top" style="border-right:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
  
   <%
   end if
   '*********** Fakturaer på aftale ******
   
   if len(oRec("faknr")) <> 0 then%>
	<tr bgcolor="<%=bgt%>">
		<td height=1 valign="top" style="border-left:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td colspan=4"><%=oRec("faknr")%></td>
		<td align=right style="border-right:1px #999999 solid; padding-right:5px;"><%=formatnumber(oRec("fakenheder"), 2)%></td>
		<td align=right style="border-right:1px #999999 solid; padding-right:5px;"><%=formatnumber(oRec("fakbelob"), 2)%></td>
		
		<%
		'*** Tilsaldo ****
		totFakEnheder = totFakEnheder + oRec("fakenheder")
		totFakBel = totFakBel + oRec("fakbelob") 
		%>
		
		<td align=right style="padding-right:5px;">
		<%Select case oRec("betalt") 
		case 1%>
		<a href="fak_godkendt.asp?aftid=<%=oRec("id")%>&id=<%=oRec("fid")%>&FM_aftnr=<%=sogaftnr%>&thiskid=<%=thiskid%>" class=vmenuglobal><%=oRec("fakdato")%></a>
		<%case else%>
		<a href="fak.asp?menu=stat_fak&FM_aftnr=<%=sogaftnr%>&thiskid=<%=thiskid%>&aftaleid=<%=oRec("id")%>&func=red&id=<%=oRec("fid")%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd%>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>" class=vmenu>Rediger</a>&nbsp;&nbsp;|&nbsp;&nbsp;
		<a href="fak.asp?menu=stat_fak&aftaleid=<%=oRec("id")%>&func=slet&id=<%=oRec("fid")%>&FM_aftnr=<%=sogaftnr%>&thiskid=<%=thiskid%>" class=vmenuslet>Slet</a>
		<%
		end select%>
		</td>
		<td>&nbsp;</td>
		<td valign="top" style="border-right:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<%end if%>
	
	
	<%
	if oRec("erfaktureret") = 0 then
	
	if lastKid <> oRec("kid") then
	ekspTxt = ekspTxt &"H;"& oRec("kkundenr") & chr(013)
	end if
	
	ekspTxt = ekspTxt & "L;" & oRec("varenr") & ";" & oRec("enheder") &";" & formatnumber(oRec("pris"), 2) &";Aftalenr:"& oRec("varenr") &";" & oRec("navn") & ";Periode: " & formatdatetime(oRec("stdato"), 2) & " til "& formatdatetime(oRec("sldato"), 2) &";" & oRec("note")&";;" & chr(013)
	
	end if
	
	lastKid = oRec("kid")
	lastaftid = oRec("id")
	
		'*** Kun enheds/klip afhængige aftaler bregnes med her ***
		select case oRec("advitype")
		case 1
		antalTildelteEnhederTot = antalTildelteEnhederTot + oRec("enheder")
		case else
		antalTildelteEnhederTot = antalTildelteEnhederTot
		end select
	
	
	s = s + 1
	oRec.movenext
	wend
	oRec.close
	
	ekspTxt = ekspTxt & "=="
	
	if s = 0 then%>
	<tr bgcolor=#ffffff>
		<td valign="top" style="border-left:1 #8caae6 solid;">&nbsp;</td>
		<td colspan="8" height=30><br><font color=red><b>Der blev ikke fundet nogen aftaler.</b></font>
		<br>Tjeck at følgende er korrekt valgt:
		<ul>
		<li>Den ønskede kunde.
		<li>Datointerval.
		<li>Øvrige filter kriterier.</ul>
		</td>
		<td valign="top" style="border-right:1 #8caae6 solid;">&nbsp;</td>
	</tr>
	<%else
	call saldo(enhThis, prisThis, tFakEnh, tFakBel, lastaftid, kundeid)
	%>
	<%end if%>
	
	
	<tr bgcolor="#5582d2">
		<td valign="top" style="border-left:1px #8caae6 solid; border-bottom:1px #8caae6 solid;"><img src="../ill/blank.gif" width="8" height="30" alt="" border="0"></td>
		<td colspan="8" style="border-bottom:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td valign="top" align="right" style="border-right:1px #8caae6 solid; border-bottom:1px #8caae6 solid;"><img src="../ill/blank.gif" width="8" height="30" alt="" border="0"></td>
	</tr>
	</table>
	<br>
	
	<div style="visibility:hidden;">
	<input type="checkbox" id="FM_aft" value="0">
	</div>
	
	
	<div id=kvik name=kvik style="position:absolute; visibility:visible; display:; left:430px; top:0px; width:360px; height:313px; background-color:#ffff99; border:1px #003399 solid; padding:5px;">
	<input type="checkbox" name="FM_fak_alle" id="FM_fak_alle" value="1">&nbsp;<b>Kvik fakturering.</b>&nbsp;
	Brug denne funktion hvis du skal fakturere mange aftaler på en gang.
	
	<br><br><b>Faktura dato:</b> (Slut dato i det valgte datointerval)&nbsp;
	<b><%=formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 2)%></b>
	<input type="hidden" name="FM_fak_dag" id="FM_fak_dag" value="<%=strDag_slut%>">
	<input type="hidden" name="FM_fak_mrd" id="FM_fak_mrd" value="<%=strMrd_slut%>">
	<input type="hidden" name="FM_fak_aar" id="FM_fak_aar" value="<%=strAar_slut%>">
	
	<br><br>
	<input type="radio" name="FM_betalt" value="2" checked>Opret som faktura <b>kladde(r)</b>.
	<input type="radio" name="FM_betalt" value="1" ><b>Godkend</b> fakturae(r). 
	
	<br><br>
	<b>Kundekonto:</b><br>
	Hvis fakturaer oprettes som "Godkendte" kan der oprettes<br>
	en postering for hver enkelt faktura på nedenstående valgte konti:<br>
	
		<select name="FM_kundekonto" style="width:200; font-size:9px;">
		<option value="0">(0)&nbsp;&nbsp;Ingen konto</option>
		<%
			strSQL = "SELECT kontonr, navn, id, kid FROM kontoplan ORDER BY kontonr, navn"
			oRec.open strSQL, oConn, 3 
			while not oRec.EOF
			
			if func = "red" then
				if intKonto = oRec("kontonr") then
				selkon = "SELECTED"
				else
				selkon = ""
				end if
			else
				if intKid = oRec("kid") then
				selkon = "SELECTED"
				else
				selkon = ""
				end if
			end if
			
			%>
			<option value="<%=oRec("kontonr")%>" <%=selkon%>>(<%=oRec("kontonr")%>)&nbsp;&nbsp;<%=oRec("navn")%></option>
			<%
			oRec.movenext
			Wend 
			oRec.close
		%>
		</select>&nbsp;&nbsp;<select name="FM_debkre" style="font-size:9px;">
				<%
					if request("faktype") <> 1 then
					selK = ""
					selD = "SELECTED"
					else
					selK = "SELECTED"
					selD = ""
					end if
				
				%>
		<option value="k" <%=selK%>>Krediter</option>
		<option value="d" <%=selD%>>Debiter</option>
		</select>
	<br>
	<b>Modkonto:</b>&nbsp;
		<select name="FM_modkonto" style="width:200px; font-size:9px;">
		<option value="0">(0)&nbsp;&nbsp;Ingen konto</option>
		<%
				strSQL = "SELECT kontonr, navn, id, kid FROM kontoplan ORDER BY kontonr, navn"
				
				oRec.open strSQL, oConn, 3 
				while not oRec.EOF 
				if intModKonto = oRec("kontonr") then
				selkon = "SELECTED"
				else
				selkon = ""
				end if
				%>
				<option value="<%=oRec("kontonr")%>" <%=selkon%>>(<%=oRec("kontonr")%>)&nbsp;&nbsp;<%=oRec("navn")%></option>
				<%
				oRec.movenext
				Wend 
				oRec.close
		%>
		</select>
	<br><br>
	<b>Enhedsangivelse:</b><br>
	<input type="radio" name="FM_enheds_ang" value="0"> Timepris&nbsp;&nbsp; 
	<input type="radio" name="FM_enheds_ang" value="1"> Stk. pris &nbsp;&nbsp;
	<input type="radio" name="FM_enheds_ang" value="2" checked> Enhedspris.
	
	<br><br>
	&nbsp;&nbsp;<input type="submit" name="t" id="Fakturér" value="Fakturér">
		
	</div>
	</form>