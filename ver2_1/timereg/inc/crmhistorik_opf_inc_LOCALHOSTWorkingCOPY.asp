
<tr>
<td colspan="5" style="padding-left:10px;"><br /><b>Planlæg opfølgende aktion:</b><br /></td>
</tr><tr>
	<td valign="top"></td><td style="padding-left:10; width:100px;">Startdato:</td>
    <td colspan="2">

				<select name="FM_start_dag_2">
				<%Call DropDownDay("FM_start_dag_2","crmdato") %></select>
				
				<select name="FM_start_mrd_2">
				<%Call DropDownMonth("FM_start_mrd_2","crmdato") %>
				</select>
				
				<select name="FM_start_aar_2">
                <%Call DropDownYear("FM_start_aar_2","crmdato") %>
                </select>&nbsp;&nbsp;
				<a href="javascript:NewWin_popupcal('../inc/regular/popupcalender_inc.asp?use=2')"><img src="../ill/popupcal.gif" alt="" border="0" width="16" height="15"></a>
    </td>
	<td></td></tr>
	<tr>
		<td valign="top"></td>
		<td style="padding-left:10;">Start tidspunkt:
		</td><td><select name="klok_timer_2">
		<%Call DropDownHour("klok_timer_2","crmklokkeslet") %>
	  	</select>&nbsp;:
		
		<select name="klok_min_2">
        <%Call DropDownMinute("klok_min_2","crmklokkeslet") %>
		</select>
		</td>
		<td>&nbsp;</td>
		<td valign="top" align="right"></td>
		</tr>
		<tr>
	<td valign="top"></td><td style="padding-left:10;">Slutdato:</td></td>
    <td >

				<select name="FM_slut_dag_2">
				<%Call DropDownDay("FM_slut_dag_2","crmdato") %></select>
				
				<select name="FM_slut_mrd_2">
				<%Call DropDownMonth("FM_slut_mrd_2","crmdato") %>
				</select>
				
				<select name="FM_slut_aar_2">
                <%Call DropDownYear("FM_slut_aar_2","crmdato") %>
                </select>&nbsp;&nbsp;
				<a href="javascript:NewWin_popupcal('../inc/regular/popupcalender_inc.asp?use=99')"><img src="../ill/popupcal.gif" alt="" border="0" width="16" height="15"></a>
    </td>
	<td></td></tr>
<tr>
	<tr>
		<td valign="top"></td>
		<td style="padding-left:10;">Slut tidspunkt:
		</td><td>
		<select name="klok_timer_slut_2">
		<%Call DropDownHour("klok_timer_slut_2","crmklokkeslet") %>
	  	</select>
		
		<select name="klok_min_slut_2">
        <%Call DropDownMinute("klok_min_slut_2","crmklokkeslet") %>
		</select>
				</td>
		<td>&nbsp;</td>
		<td valign="top" align="right"></td>
		</tr>
		<tr>
		<td></td><td colspan="2" style="padding-left:10;">
		<br><br>
		Overskrift:&nbsp;&nbsp;
		<a onClick="insvaloskrift();"><img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="Overfør fra Aktion 1" border="0"></a>&nbsp;<input type="text" name="FM_navn_2" value="<%=strNavn%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		
		
		
	<br>Kontaktperson:&nbsp;&nbsp;
	<a onClick="insvalkpers();"><img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="Overfør fra Aktion 1" border="0"></a>
	<select name="FM_kpers_2">
		<%strSQL = "SELECT id, navn FROM kontaktpers WHERE kundeid = "& id
		oRec.open strSQL, oConn, 3
		while not oRec.EOF
		if (cint(intKontaktpers) = oRec("id") and Request.Form("FM_kpers_2") = "") or (cint(request.Form("FM_kpers_2")) = oRec("id")) then
		kpersSel = "SELECTED"
		else
		kpersSel = ""
		end if
		%>
		<option value="<%=oRec("id")%>" <%=kpersSel%>><%=oRec("navn")%></option>
		<%oRec.movenext 
		wend
		oRec.close%>
		<option value="0">Ingen</option>
	</select>
	
	
	<br><br>Beskrivelse:<br>
	<% strBesk2 = Request.Form("FM_besk_2")%>
	<textarea cols="40" rows="8" name="FM_besk_2" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"><%=strBesk2 %></textarea>
	
	
		<br><br>Status:&nbsp;&nbsp;
		<a onClick="insvalstat();"><img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="Overfør fra Aktion 1" border="0"></a>&nbsp;
		<select name="FM_status_2">
		<%strSQL = "SELECT id, navn FROM crmstatus ORDER BY navn"
		oRec.open strSQL, oConn, 3
		While not oRec.EOF
        if (intStatus = oRec("id") AND request.Form("FM_status_2") = "") or cint(oRec("id")) = cint(request.Form("FM_status_2")) then
		selected = "selected"
		else
		selected = ""
		end if%>
		<option value="<%=oRec("id")%>" <%=selected%>><%=oRec("navn")%></option>
		<%oRec.movenext
		wend
		oRec.close%>
		</select>
		
		
	
		<br>Kontaktform:&nbsp;&nbsp;
		<a onClick="insvalkform();"><img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="Overfør fra Aktion 1" border="0"></a>&nbsp;
		<select name="FM_kform_2">
			<%strSQL = "SELECT id, navn FROM crmkontaktform ORDER BY navn"
			oRec.open strSQL, oConn, 3
			While not oRec.EOF 
			if (intKontaktform = oRec("id") AND Request.Form("FM_kform_2") = "") or cint(oRec("id")) = cint(Request.Form("FM_kform_2")) then
			selected = "selected"
			else
			selected = ""
			end if%> 
			<option value="<%=oRec("id")%>" <%=selected%>><%=oRec("navn")%></option>
			<%oRec.movenext
			wend
			oRec.close%>
		</select>
			</td>
			<td valign="top" align="right"></td>
		</tr>
		<tr>
	<tr>
		<td valign="top"></td>
		<td colspan="3" style="padding-left:10;"><br>Tilføj medarbejdere:</td>
		<td valign="top" align="right"></td>
	</tr>
				<%
				'** Opfølgning (tilknyt medarbejdere)**
				
				strSQL = "SELECT mid, mnavn FROM medarbejdere WHERE Brugergruppe = 1 OR Brugergruppe = 3 OR Brugergruppe = 6 OR Brugergruppe = 8 ORDER BY mnavn"
				oRec.open strSQL, oConn, 0, 1
				while not oRec.EOF
				
					if func = "red" then
						strSQL2 = "SELECT medarbid FROM aktionsrelationer WHERE aktionsid = "& crmaktion &" AND medarbid = "& oRec("mid")
						oRec2.open strSQL2, oConn, 0, 1
						if not oRec2.EOF then
						strcheckmedarb = "CHECKED"
						else
						strcheckmedarb = " "
						end if
						oRec2.close
					else
			if instr(request.Form("FM_medarbrel_2"),",") then
			CheckMid = split(Request.Form("FM_medarbrel_2"),",")
			for i = 0 to Ubound(CheckMid)
			if cint(CheckMid(i)) = cint(oRec("mid")) then
			strCheckit = "y"
			end if
			Next
			elseif cint(oRec("mid")) = cint(Request.Form("FM_medarbrel_2")) then
			strCheckit = "y"
			end if
				if (oRec("mnavn") = session("user") AND request.Form("FM_medarbrel_2") = "") or strCheckit = "y" then
				strcheckmedarb = "CHECKED"
				else
				strcheckmedarb = " "
				end if
				strCheckit = Null
				CheckMid = Null
		    end if
				%> 
				<tr>
			<td colspan="2" style="padding-left:10;"><%=oRec("mnavn")%></td><td colspan="3" style="padding-left:10;"><input type="checkbox" name="FM_medarbrel_2" value="<%=oRec("mid")%>" <%=strcheckmedarb%> style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
			</td>
		</tr><%
				x = x + 1
				oRec.movenext
				wend
				oRec.close
				%>
		 