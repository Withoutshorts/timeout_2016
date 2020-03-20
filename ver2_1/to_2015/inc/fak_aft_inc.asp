<br><br>

	<table cellspacing="0" cellpadding="0" border="0" width="700" bgcolor="#FFFFFF">
	<tr bgcolor="#d6dff5">
		<td colspan=5><img src="../ill/blank.gif" width="1" height="20" alt="" border="0"><br></td>
	</tr>
	<tr bgcolor="#66CC33">
		<td width=60 height=20 style="padding-left:10px; border-left:1px #8caae6 solid; border-top:1px #8caae6 solid;"><font class=megetlillesort>Varenr</td>
		<td width=80 style="border-top:1px #8caae6 solid;"><font class=megetlillesort>Enheder</td>
		<td style="border-top:1px #8caae6 solid;"><font class=megetlillesort>Navn og beskrivelse</td>
		<td width=80 style="border-top:1px #8caae6 solid;"><font class=megetlillesort>Enh. pris.</td>
		<td width=80 style="border-right:1px #8caae6 solid; border-top:1px #8caae6 solid;"><font class=megetlillesort>Pris</td>
	</tr>
	<tr>
		<%
		if len(intEnheder) <> 0 then
		intEnheder = formatnumber(intEnheder, 2)
		else
		intEnheder = formatnumber(0, 2)
		end if
		intEnheder = replace(intEnheder, ".", "")
		%>
		<td valign=top style="padding-left:10px; padding-top:3px; border-left:1px #8caae6 solid; border-bottom:1px #8caae6 solid;"><input type="text" name="FM_varenr" value="<%=strVarenr%>" size=4 style="border: 1px #000000 solid; background-color: #FFFFFF;"></td>
		<td valign=top style="padding-top:3px; border-bottom:1px #8caae6 solid;"><input type="text" name="FM_Timer" value="<%=intEnheder%>" size=6 style="border: 1px limegreen solid; background-color: #FFFFFF;" onKeyup="aft_stk_pris()"></td>
		<td valign=top style="padding-top:3px; border-bottom:1px #8caae6 solid;"><textarea cols="41" rows="8" name="FM_komm"><%=strBesk%></textarea><br>&nbsp;</td>
		<%'** Enhedspris ***
		'** Hvis perafg. = 0 then enhpris = pris
		'** Hvis klip afh. then enhpris = pris/enheder uanset om aftalen er pariodeafh.
		
		if intEnheder > 0 then
				stkpris = (intPris/intEnheder)
		else
		stkpris = 0
		end if
		
		'intPerafg 
		'intAdvitype 
		'intAdvihvor
		
		if len(stkpris) <> 0 then
		stkpris = formatnumber(stkpris, 2)
		else
		stkpris = formatnumber(0, 2)
		end if
		stkpris = replace(stkpris, ".", "")
		%>
		<td valign=top style="padding-top:3px; border-bottom:1px #8caae6 solid;"><input type="text" name="FM_stkpris" id="FM_stkpris" value="<%=stkpris%>" size=8 style="border-bottom: 1px #999999 solid; border-right: 0px;  border-top: 0px;  border-left: 0px; background-color: #FFFFFF;"></td>
		<%
		if len(intPris) <> 0 then
		intPris = formatnumber(intPris, 2)
		else
		intPris = formatnumber(0, 2)
		end if
		intPris = replace(intPris, ".", "")
		%>
		<td valign=top style="border-right:1px #8caae6 solid; border-bottom:1px #8caae6 solid; padding-top:3px;"><input type="text" name="FM_beloeb" value="<%=intPris%>" size=8 style="border: 1px #86B5E4 solid; background-color: #FFFFFF;" onKeyup="aft_stk_pris()"></td>
	</tr>
	
	<tr bgcolor="#d6dff5">
		<td colspan=5 align=center><br><input name="subm_on" id="subm_on"  type="image" src="../ill/opretpil_fak.gif"></td>
	</tr>
	</form>
	</table>
	<br><br>