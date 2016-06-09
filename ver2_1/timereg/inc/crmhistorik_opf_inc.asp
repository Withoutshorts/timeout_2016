<tr>
	<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="380" alt="" border="0"></td>
	<td colspan="3" style="padding-left:10;" valign="top">Start tidspunkt:&nbsp;&nbsp;
		<select name="klok_timer_2">
		<option value="<%=klok_timer%>" SELECTED><%=klok_timer%></option>
		<option value="07">7</option>
	   	<option value="08">8</option>
	   	<option value="09">9</option>
	   	<option value="10">10</option>
	   	<option value="11">11</option>
	   	<option value="12">12</option>
	   	<option value="13">13</option>
	   	<option value="14">14</option>
	   	<option value="15">15</option>
	   	<option value="16">16</option>
	   	<option value="17">17</option>
	   	<option value="18">18</option>
	   	<option value="19">19</option>
	   	<option value="20">20</option>
	   	<option value="21">21</option>
	  	</select>&nbsp;:
		
		<select name="klok_min_2">
		<option value="00">00</option>
		<option value="15">15</option>
	   	<option value="30">30</option>
	   	<option value="45">45</option>
		</select>
	
		<br>
		Slut tidspunkt:&nbsp;&nbsp;&nbsp;&nbsp;
		<select name="klok_timer_slut_2">
		<option value="<%=klok_timer_slut%>" SELECTED><%=klok_timer_slut%></option>
		<option value="07">7</option>
	   	<option value="08">8</option>
	   	<option value="09">9</option>
	   	<option value="10">10</option>
	   	<option value="11">11</option>
	   	<option value="12">12</option>
	   	<option value="13">13</option>
	   	<option value="14">14</option>
	   	<option value="15">15</option>
	   	<option value="16">16</option>
	   	<option value="17">17</option>
	   	<option value="18">18</option>
	   	<option value="19">19</option>
	   	<option value="20">20</option>
	   	<option value="21">21</option>
	  	</select>&nbsp;:
		
		<select name="klok_min_slut_2">
		<option value="00">00</option>
		<option value="15">15</option>
	   	<option value="30" selected>30</option>
	   	<option value="45">45</option>
		</select>
		
		
		<br><br>
		Overskrift:&nbsp;&nbsp;
		<a href="#" onClick="insvaloskrift();"><img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="Overfør fra Aktion 1" border="0"></a>&nbsp;<input type="text" name="FM_navn_2" value="<%=strNavn%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">
		
		
		
	<br>Kontaktperson:&nbsp;&nbsp;
	<a href="#" onClick="insvalkpers();"><img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="Overfør fra Aktion 1" border="0"></a>
	<select name="FM_kpers_2">
		<%strSQL = "SELECT id, navn FROM kontaktpers WHERE kundeid = "& id
		oRec.open strSQL, oConn, 3
		while not oRec.EOF
		if cint(intKontaktpers) = oRec("id") then
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
	<textarea cols="40" rows="8" name="FM_besk_2" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></textarea>
	
	
		<br><br>Status:&nbsp;&nbsp;
		<a href="#" onClick="insvalstat();"><img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="Overfør fra Aktion 1" border="0"></a>&nbsp;
		<select name="FM_status_2">
		<%strSQL = "SELECT id, navn FROM crmstatus ORDER BY navn"
		oRec.open strSQL, oConn, 3
		While not oRec.EOF %>
		<%if intStatus = oRec("id") then
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
		<a href="#" onClick="insvalkform();"><img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="Overfør fra Aktion 1" border="0"></a>&nbsp;
		<select name="FM_kform_2">
			<%strSQL = "SELECT id, navn FROM crmkontaktform ORDER BY navn"
			oRec.open strSQL, oConn, 3
			While not oRec.EOF %>
			<%if intKontaktform = oRec("id") then
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
			<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="380" alt="" border="0"></td>
		</tr>
		<tr>
			<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="200" alt="" border="0"></td>
			<td colspan="3" style="padding-left:10;">
			Tilføj medarbejdere til denne aktion:<br>
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
						if oRec("mnavn") = session("user") then
						strcheckmedarb = "CHECKED"
						else
						strcheckmedarb = " "
						end if
					end if
				
				%> 
				<input type="checkbox" name="FM_medarbrel_2" value="<%=oRec("mid")%>" <%=strcheckmedarb%> style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">&nbsp;<%=oRec("mnavn")%>
				<br><%
				x = x + 1
				oRec.movenext
				wend
				oRec.close
				%>&nbsp;
			
				
			</td>
			<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="200" alt="" border="0"></td>
		</tr>
		 