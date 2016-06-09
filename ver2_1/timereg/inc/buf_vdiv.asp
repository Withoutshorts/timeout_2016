for w = 0 to w - 1
		
		if w = g then
		varVisi = "visible"
		else
		varVisi = "hidden"
		end if
		varWeekTop = varTop - 32
		
		if request("print") <> "j" then%>
		<div id="weekdiv_<%=w%>" name="weekdiv_<%=w%>" style="position:absolute; top:<%=varWeekTop%>; left:0; z-index:100; visibility:<%=varVisi%>;">
		<%end if%>
		<table border="0" cellpadding="0" cellspacing="0" width="260" style="page-break-before:always;">
		<tr><td height=30 colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td></tr>
		<tr>
			<td colspan="6" height=20 style="padding-left : 4px;">&nbsp;</td>
		</tr>
		<tr>
			<td height=1 colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
		</tr>
		<%
				for a = 0 to a - 1
				%>
				<tr><td height=2 colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td></tr>
				<tr>
					<td height="20" colspan="2" width="185" style="!border: 1 px; background-color: LightBlue; border-color: silver; border-style: solid; padding-left : 4px;"><font size=1><%= jobnr(a) &"&nbsp;" & jobnavn(a) &"&nbsp;</font><br><b>" & left(aktiviteter(a), 16)%></b></td>
					<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					<td width="50" align=right valign="bottom" style="!border: 1 px; background-color: LightBlue; border-color: silver; border-style: solid; padding-left : 4px;"><font size=1>timer</td>
					<td width="1"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					<td width="54" align=right valign="bottom" style="!border: 1 px; background-color: LightBlue; border-color: silver; border-style: solid; padding-left : 4px;"><font size=1>omsætning</td>
				</tr>
					<%
					for t = 0 to t - 1 
					%>
					<tr>
						<tr><td height=1 colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td></tr>
						<td colspan=2 style="!border: 1 px; background-color: #FFFFFF; border-color: #d2691e; border-style: solid; padding-left : 4px;"><%=medarbejder(t)%></td>
						<td width=1><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
						<td width="50" align=right style="!border: 1 px; background-color: #FFFFFF; border-color: silver; border-style: solid; padding-left : 4px; padding-right : 1px;">
						<%
						
						'call weekberegning()
						
						Response.write SQLBless(intZtimer) &"</td><td width='1'><img src='ill/blank.gif' width='1' height='1' border='0' alt=''></td>"_
						&"<td width=54 align=right style='!border: 1 px; background-color: #FFFFFF; border-color: silver; border-style: solid; padding-left : 4px; padding-right : 1px;'>" & SQLBless(ccur(totOms)) &""
						intZtimer = 0
						totOms = 0
						%></td></tr>
					<%
					next
					%>
					<tr><td height=2 colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td></tr>
					<tr><td colspan=2 style="!border: 1 px; background-color: #FFFFFF; border-color: Darkred; border-style: solid; padding-left : 4px;">Ialt: </td>
					<td width=1><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					<td align=right style="!border: 1 px; background-color: #FFFFFF; border-color: Darkred; border-style: solid; padding-left : 4px; padding-right : 1px;"><%
					Response.write SQLBless(intZtimerTotalAkt)
					intZtimerTotalAkt = 0%></td><td width=1><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					<td align="right" style='!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px; padding-right : 1px;'><%=SQLBless(ccur(totOmsAkt))%></td>
					</tr>
					<%totOmsAkt = 0
				next
			%>
			<tr><td height=20 colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td></tr>
					<tr><td colspan=2 style="!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px;"><b>Uge <%=weeknumbers(w)%> total:</b></td>
					<td width=1><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					<td align="right" style="!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px; padding-right : 4px;"><b><%
					Response.write SQLBless(intZtimerTotalWeek)
					intZtimerTotalWeek = 0%></b></td><td width=1><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					<td align="right" style='!border: 1 px; background-color: #FFFFFF; border-color: DarkRed; border-style: solid; padding-left : 4px; padding-right : 1px;'><b><%=SQLBless(ccur(totOmsWeek))%></b></td>
					</tr>
				</TABLE><br><br><br>
				<%if request("print") <> "j" then%>
				</div>
				<%end if%>
			<%totOmsWeek = 0
		next
