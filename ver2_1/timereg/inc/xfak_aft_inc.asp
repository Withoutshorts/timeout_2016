    
    <div id="aktdiv" style="position:absolute; visibility:hidden; display:none; left:5px; top:105px; width:600px;">
    <input id="FM_enheds_ang" name="FM_enheds_ang" type="hidden" value="2" />
     
	<table cellspacing="0" cellpadding="0" border="0" width="640">

	<tr bgcolor="Yellowgreen">
		<td width=60 height=20 style="padding-left:10px;"><b>Varenr</b></td>
		<td width=100><b>Enheder</b></td>
		<td width=280 style="padding:3px 3px 3px 3px;"><b>Navn og beskrivelse</b>
		<br />Marker og formatér tekst:&nbsp;<input id="Button1" type="button" value="b" onclick="insertTag('b','FM_jobbesk')" style="font-size:9px; width:20px;">&nbsp;
		<input id="Button2" type="button" value="i" onclick="insertTag('i','FM_jobbesk')" style="font-size:9px; width:20px;">&nbsp;
		<input id="Button4" type="button" value="u" onclick="insertTag('u','FM_jobbesk')" style="font-size:9px; width:20px;">&nbsp;
		<input id="Button3" type="button" value="br" onclick="insertTag('br','FM_jobbesk')" style="font-size:9px; width:20px;">&nbsp;
		</td>
		<td width=60 style="padding:0px 0px 0px 10px;"><b>Enh. pris.</b></td>
		<td width=60><b>Rabat</b></td>
		<td width=80><b>Pris</b></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<%
		if len(intEnheder) <> 0 then
		intEnheder = formatnumber(intEnheder, 2)
		else
		intEnheder = formatnumber(0, 2)
		end if
		intEnheder = replace(intEnheder, ".", "")
		%>
		<td valign=top style="padding-left:10px; padding-top:3px; border-bottom:1px yellowgreen solid; border-left:1px yellowgreen solid;"><input type="text" name="FM_varenr" value="<%=strVarenr%>" style="font-size:9px; width:50px;"></td>
		<td valign=top style="padding-top:3px; border-bottom:1px yellowgreen solid;"><input type="text" name="FM_Timer" value="<%=intEnheder%>" style="font-size:9px; width:50px;" onKeyup="aft_stk_pris()"><br />
		<br />
		
		  <% 
	            '** Enheder i periode **
                                    
                sumEnhVedFak = 0
                sumTimerEfterFak = 0
                strSQLenh = "SELECT sum(timer * faktor) AS sumEnhEfterFak, sum(timer) AS sumTimerEfterFak FROM timer "_
                &" LEFT JOIN aktiviteter a ON (a.id = taktivitetid) WHERE "_
                &" tfaktim <> 5 AND seraft = "& aftid &""_ 
                &" AND tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"' GROUP BY seraft, taktivitetid"
	            
                'Response.Write strSQLenh &"<br>"
	            
                oRec3.open strSQLenh, oConn, 3
                while not oRec3.EOF
                
                sumEnhVedFak = sumEnhVedFak + oRec3("sumEnhEfterFak")
                sumTimerEfterFak = sumTimerEfterFak + oRec3("sumTimerEfterFak")
                
                oRec3.movenext
                wend
                oRec3.close
                 
                 Response.Write "<b>I Periode:</b><br>"
                 Response.Write "Timer: <b>"& sumTimerEfterFak & "</b><br>"
                 Response.Write "Enheder: <b>" & sumEnhVedFak & "</b><br>&nbsp;"

	            %>
		
		</td>
		<td valign=top style="padding-top:3px; border-bottom:1px yellowgreen solid;"><textarea name="FM_jobbesk" style="font-size:11px; width:270px; height:100px;"><%=strJobBesk%></textarea><br>&nbsp;</td>
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
		<td valign=top style="padding:3px 10px 0px 10px; border-bottom:1px yellowgreen solid;"><input type="text" name="FM_stkpris" id="FM_stkpris" value="<%=stkpris%>" style="font-size:9px; width:60px;"></td>
		<td valign=top style="padding:3px 0px 0px  0px ; border-bottom:1px yellowgreen solid;">
		
		
		<%
		                        select case intRabat
		                        case 0
		                        rSel0 = "SELECTED"
		                        rSel10 = ""
		                        rSel25 = ""
		                        rSel40 = ""
		                        rSel50 = ""
		                        rSel60 = ""
		                        rSel75 = ""
		                        case 10
		                        rSel0 = ""
		                        rSel10 = "SELECTED"
		                        rSel25 = ""
		                        rSel40 = ""
		                        rSel50 = ""
		                        rSel60 = ""
		                        rSel75 = ""
		                        case 25
		                        rSel0 = ""
		                        rSel10 = ""
		                        rSel25 = "SELECTED"
		                        rSel40 = ""
		                        rSel50 = ""
		                        rSel60 = ""
		                        rSel75 = ""
	                            case 40
		                        rSel0 = ""
		                        rSel10 = ""
		                        rSel25 = ""
		                        rSel40 = "SELECTED"
		                        rSel50 = ""
		                        rSel60 = ""
		                        rSel75 = ""
		                        case 50
		                        rSel0 = ""
		                        rSel10 = ""
		                        rSel25 = ""
		                        rSel40 = ""
		                        rSel50 = "SELECTED"
		                        rSel60 = ""
		                        rSel75 = ""
		                        case 60
		                        rSel0 = ""
		                        rSel10 = ""
		                        rSel25 = ""
		                        rSel40 = ""
		                        rSel50 = ""
		                        rSel60 = "SELECTED"
		                        rSel75 = ""
		                        case 75
		                        rSel0 = ""
		                        rSel10 = ""
		                        rSel25 = ""
		                        rSel40 = ""
		                        rSel50 = ""
		                        rSel60 = ""
		                        rSel75 = "SELECTED"
		                        end select
		
		%>
		
		
		<select id="FM_rabat" name="FM_rabat" style="width:50px; font-size:9px;" onChange="aft_stk_pris()">
                                        <option value="0"  <%=rSel0%>>0%</option>
                                        <option value="0.10" <%=rSel10%>>10%</option>
                                        <option value="0.25" <%=rSel25%>>25%</option>
                                        <option value="0.40" <%=rSel40%>>40%</option>
                                        <option value="0.50" <%=rSel50%>>50%</option>
                                        <option value="0.60" <%=rSel60%>>60%</option>
                                        <option value="0.75" <%=rSel75%>>75%</option>
                                        </select>
        </td>
		<td valign=top style="padding-top:3px; border-bottom:1px yellowgreen solid; border-right:1px yellowgreen solid;">
		<%
		if len(intPris) <> 0 then
		    
		    if intRabat <> 0 then
		    intPris = intPris - (intPris * intRabat/100)
		    else
		    intPris = intPris
		    end if
		    
		    intPris = formatnumber(intPris, 2)
		else
		intPris = formatnumber(0, 2)
		end if
		intPris = replace(intPris, ".", "")
		%>
		<input type="text" name="FM_beloeb" value="<%=intPris%>" style="font-size:9px; width:60px;" onKeyup="aft_stk_pris()"></td>
	</tr>
	
	<tr>
		<td colspan=6 align=right style="padding:10px 10px 10px 10px;"><br>
		<input name="subm_on" id="subm_on" type="submit" value="Se faktura" /></td>
	</tr>
	</table>
	<br /><br />
	          
	
&nbsp;
</div>

<div id=aktdiv_2 style="position:absolute; visibility:hidden; display:none; top:336px; width:640px; left:5px; border:0px #8cAAe6 solid;">
<table width=640 cellspacing=0 cellpadding=5 border=0><tr><td><a href="#" onclick="showdiv('modtagdiv')" class=vmenu><< Forrige</a></td><td align=right><a href="#" onclick="showdiv('betdiv')" class=vmenu>Næste >></a></td></tr></table>
 </div>