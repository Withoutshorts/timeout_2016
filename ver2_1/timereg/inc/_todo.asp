
	<table cellspacing=20 cellpadding=0 border=0 width=950>
	<tr>
		<td width=400 valign=top style="border:2px #cccccc solid; padding:10px;" bgcolor="#ffffff"><h4>ToDo listen</h4>
	
	<a href="Javascript:history.back();" class=todo_lille><img src="../ill/soeg-knap_tilbage.gif" width="16" height="16" alt="" border="0"> Tilbage</a>&nbsp;&nbsp;
	|&nbsp;&nbsp;<a href="webblik.asp?FM_parent=0" class=todo_lille>ToDo oversigt</a><br><br>
	<table cellspacing=0 cellpadding=0 border=0 widht=400>
	<tr><form method=post action="webblik.asp?FM_parent=<%=parent%>">
		<td colspan=4 style="border-bottom:1px #cccccc solid; padding:2px;">
		<%
		todolevel = request("todolevel")
		if todolevel < 3 then
		%>
		<a href="#" onClick="showtodo('<%=parent%>','<%=session("mid")%>')" class=todo_stor>Opret ny ToDo her&nbsp;<img src="../ill/soeg-knap.gif" width="16" height="16" alt="" border="0"></a>
		<%else%>
		<br><b>Der kan maks oprettes 3 niveauer ToDo's!</b>
		<%end if
		
		
		if len(request("FM_afsluttede")) <> 0 then
			visafl = request("FM_afsluttede")
			if visafl = "1" then
			visaflCHK0 = "CHECKED"
			visaflCHK1 = ""
			else
			visafl = 0
			visaflCHK0 = ""
			visaflCHK1 = "CHECKED"
			end if
			Response.cookies("webblik_visafl") = visafl
			Response.cookies("webblik_visafl").expires = date + 40
		else
			if request.cookies("webblik_visafl") = "1" then
			visafl = 1
			visaflCHK0 = "CHECKED"
			visaflCHK1 = ""
			else
			visafl = 0
			visaflCHK0 = ""
			visaflCHK1 = "CHECKED"
			end if
		end if
		
		%>
		
		
		<br><b>Vis afsluttede ToDo's?</b>
		<input type="radio" name="FM_afsluttede" id="FM_afsluttede" value="1" <%=visaflCHK0%>> Ja &nbsp;&nbsp;
		<input type="radio" name="FM_afsluttede" id="FM_afsluttede" value="0" <%=visaflCHK1%>> Nej &nbsp;&nbsp;
		<input type="submit" value="Vis" style="width:30; font-size:10px;"><br><br>
		
		
	</td>
	</form>
	</tr>
	<%
	'**** Todo ****************'
	if oldones <> "1" then
	sqlDatostart = year(now)&"/"& month(now)&"/"&day(now) 
	datointervalslut = dateadd("d", -7, sqlDatostart)  
	sqlDatoslut = year(datointervalslut)&"/"& month(datointervalslut) &"/"& day(datointervalslut)
	else
	sqlDatostart = year(now)&"/"& month(now)&"/"&day(now) 
	sqlDatoslut = "2005/1/1"
	end if
	
	
	if cint(parent) = 0 then
	
	datoSQL = " AND dato BETWEEN '"& sqlDatoslut &"' AND '"& sqlDatostart &"'"
	
	else
		
		
		'**** klikket Todo SQL **************
		strSQL = "SELECT t.navn, t.id, t.parent, t.dato, t.level, t.delt, t.afsluttet, t.sortorder"_
		&" FROM todo_new t WHERE t.id = "& parent 
		
		oRec.open strSQL, oConn, 3 
		if not oRec.EOF then
					
					'*** Opdaterer dato på redigering på den viste todo ***
					sqlToDoVsitDato = year(now) & "/" & month(now) & "/" & day(now)
					strSQL = "UPDATE todo_new SET dato = '"& sqlToDoVsitDato &"' WHERE id = "& oRec("id")
					oConn.execute(strSQL)
					
					antalsubs = 0
			
					strSQL2 = "SELECT count(parent) AS antalsubs FROM todo_new WHERE parent = " & oRec("id") &" GROUP BY parent"
					oRec2.open strSQL2, oConn, 3 
					if not oRec2.EOF then 
					
					antalsubs = oRec2("antalsubs")
					
					end if
					oRec2.close 
					
				acls = "todo_stor"
				
				call todolist(0)
		
		
		end if
		oRec.close 
		
		
	datoSQL = ""
	end if
	
	
	if visafl = 1 then
	visaflSQL = " AND t.afsluttet <> 99 "
	else
	visaflSQL = " AND t.afsluttet <> 1 "
	end if
	
	'**** Main Todo SQL **************
	strSQL = "SELECT t.navn, t.id, t.parent, t.dato, t.level, t.delt, t.afsluttet, t.sortorder, trn.medarbid, trn.todoid "_
	&" FROM todo_rel_new trn "_
	&" RIGHT JOIN todo_new t ON (t.id = trn.todoid "& visaflSQL &" AND (parent = "& parent &" "& datoSQL &"))"_
	&" WHERE medarbid = "& session("mid") &" GROUP BY t.id ORDER BY t.afsluttet, t.sortorder"
	
	
	'Response.write strSQL
	
	oRec.open strSQL, oConn, 3 
	while not oRec.EOF 
			
			antalsubs = 0
			
			strSQL2 = "SELECT count(parent) AS antalsubs FROM todo_new WHERE parent = " & oRec("id") &" GROUP BY parent"
			oRec2.open strSQL2, oConn, 3 
			if not oRec2.EOF then 
			
			antalsubs = oRec2("antalsubs")
			
			end if
			oRec2.close 
			
	
			
		acls = "todo_mellem"
		
		
	call todolist(1)
	
	oRec.movenext
	wend
	oRec.close %>
	
	
	<tr>
		<td colspan=4 style="border-bottom:1px #cccccc solid; padding:2px;">
	<%
	'*** Mere end 7 dage gamle ToDo's ***********************
	if cint(parent) = 0 then
	
			if request("FM_visalle") = "1" then
			Response.cookies("visallegamle") = "j"
			Response.cookies("visallegamle").expires = date + 10
			fmallechk0 = ""
			fmallechk1 = "CHECKED"
			else
			Response.cookies("visallegamle") = "n"
			Response.cookies("visallegamle").expires = date + 10
			fmallechk0 = "CHECKED"
			fmallechk1 = ""
			end if%>
			<br><br><br><form method=post action="webblik.asp?FM_parent=<%=parent%>">
			<b>Mere end en uge gamle ToDo's:</b><br>
			Vis kun maks 3 md. gamle  <input type="radio" name="FM_visalle" id="FM_visalle" value="0" <%=fmallechk0%>>&nbsp;&nbsp;
			Vis alle:<input type="radio" name="FM_visalle" id="FM_visalle" value="1" <%=fmallechk1%>>
			<input type="submit" value="Vis" style="width:30; font-size:10px;">
			</form></td>
			</tr><tr>
			<td colspan=4><div style="position:relative; height:400px; overflow:auto;">
			<%
			
			sqlDatostart = year(now)&"/"& month(now)&"/"&day(now) 
			datointervalslut = dateadd("d", -8, sqlDatostart)  
			sqlDatostart = year(datointervalslut)&"/"& month(datointervalslut) &"/"& day(datointervalslut)
			
			if request.cookies("visallegamle") = "j" then
			sqlDatoslut = "2002/1/1"
			else
			sqlDatoslut_temp = dateadd("d", -90, day(now)&"/"& month(now)&"/"&year(now)) 
			sqlDatoslut = year(sqlDatoslut_temp)&"/"& month(sqlDatoslut_temp)&"/"&day(sqlDatoslut_temp) 
			end if
			
				
			strSQL3 = "SELECT t.navn, t.id, t.parent, t.dato, t.level, t.delt, trn.medarbid, trn.todoid, t.afsluttet "_
			&" FROM todo_rel_new trn "_
			&" RIGHT JOIN todo_new t ON (t.id = trn.todoid "& visaflSQL &" AND (parent = "& parent &" "& hentklikketSQL &")"_
			&" AND dato BETWEEN '"& sqlDatoslut &"' AND '"& sqlDatostart &"') "_
			&" WHERE medarbid = "& session("mid") &" GROUP BY t.id ORDER BY t.navn" 
			
			'Response.write "<br><br>" & strSQL
			'Response.flush
			
			oRec3.open strSQL3, oConn, 3 
			while not oRec3.EOF 
				
			%>
			<a href="webblik.asp?FM_parent=<%=oRec3("id")%>&todolevel=<%=oRec3("level")%>&oldone=1" class="todo_lille">
			<% if oRec3("afsluttet") = 1 then%>
			<s>
			<%end if%>
			
			<%=oRec3("navn")%>
			
			<% if oRec3("afsluttet") = 1 then%>
			</s>
			<%end if%>
			</a>
			 
			 
			 &nbsp;<a href="webblik_oprtodo.asp?func=slet&id=<%=oRec3("id")%>&FM_parent=<%=parent%>&oldone=<%=oldones%>" class="red">[x]</a>&nbsp;&nbsp; 
			<br>
			<%
			oRec3.movenext
			wend
			oRec3.close 
	
	end if
	%>
	</div>
	</td></tr>
	</table>
	
	</td>
	
	
	
	<%
	'**********************************************************
	'**************** Milepæle *******************************'
	'**********************************************************
	%>
	<td valign=top style="border:2px #5582d2 solid; padding:10px;" bgcolor="#ffffff"><h4>Milepæle</h4>
	<%call periodeForm%>
	<font class=megetlillesort>(Job start- og slut -datoer er ikke vist)</font>
	<br>
	<%
	
	sqlDatostart = strAar&"/"&strMrd&"/"&strDag  
	sqlDatoslut = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut
		
	
	Response.write "<b>Periode:</b> " & formatdatetime(strDag&"/"&strMrd&"/"&strAar, 1) & " - " & formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 1)
	%>
	<table cellspacing=0 cellpadding=0 border=0 width=400>
	<tr>
		 <td colspan=3 align=right style="border-bottom:1px #999999 dashed; padding:2px;">
			<a href="javascript:popUp('milepale.asp?menu=job&func=opr&jid=0&rdir=webblik','650','500','250','120');" target="_self" class=todo_stor>Opret Ny milepæl&nbsp;<img src="../ill/soeg-knap.gif" width="16" height="16" alt="" border="0"></a>
		</td>
	</tr>
	<tr><td style="padding-top:0px;">
	
		
		<table cellspacing=0 cellpadding=0 border=0 width=100%>
		<%
		if func = "sletmilepal" then
		'*** Her slettes en milepal ***
		milepalid = request("milepalid")
		oConn.execute("DELETE FROM milepale WHERE id = "& milepalid &"")
		end if
		
		
		
		'datointervalstart = dateadd("d", -5, year(now)&"/"& month(now)&"/"&day(now)) 
		'sqlDatostart = year(datointervalstart)&"/"& month(datointervalstart)&"/"&day(datointervalstart) 
		'datointervalslut = dateadd("d", 65, sqlDatostart)  
		'sqlDatoslut = year(datointervalslut)&"/"& month(datointervalslut) &"/"& day(datointervalslut)
		
		lastDato = ""
		strSQL = "SELECT m.navn, m.milepal_dato, m.jid, m.beskrivelse, j.jobnavn, j.jobnr, m.id AS mileid, "_
		&" mt.id, mt.navn AS typenavn "_
		&" FROM milepale m "_
		&" LEFT JOIN  job j ON (j.id = m.jid "& strPgrpSQLkri &") "_
		&" LEFT JOIN milepale_typer mt ON (mt.id = m.type) "_
		&" WHERE milepal_dato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"' ORDER BY m.milepal_dato"
		
		'Response.write strSQL
		'Response.flush
		
		oRec.open strSQL, oConn, 3 
		x = 0
		while not oRec.EOF 
		
		if len(trim(oRec("jobnavn"))) <> 0 then
		
			if lastDato <> oRec("milepal_dato") then%>
			<tr><td valign=bottom style="padding-top:10px; border-bottom:1px #999999 dashed; padding:5px;" bgcolor="#d6dff5"><font class=stor-blaa><%=formatdatetime(oRec("milepal_dato"), 2)%></font></td></tr>
			<%end if%>
			
			<tr><td style="border-bottom:1px #999999 dashed; padding:5px;">
			<b>Job:</b> <%=oRec("jobnavn")%> (<%=oRec("jobnr")%>) <br>
			<b><%=oRec("typenavn")%>:</b> <a href="javascript:popUp('milepale.asp?func=red&id=<%=oRec("mileid")%>&jid=<%=oRec("jid")%>&rdir=webblik','650','500','250','120');" target="_self" class=todo_mellem><%=oRec("navn")%></a>&nbsp;&nbsp;<a href="webblik.asp?func=sletmilepal&milepalid=<%=oRec("id")%>" class=red>[x]</a><br>
			<%if len(oRec("beskrivelse")) <> 0 then%>
			<font size=1 color="#999999"><%=oRec("beskrivelse")%></font>
			<%end if%>
			</td></tr>
			
			<%
			lastDato = oRec("milepal_dato")
			x = x + 1
		
		end if
		
		oRec.movenext
		wend
		oRec.close 
		%>
		</table>
	
	</td></tr></table>
	 <!-- todo/milepæle -->