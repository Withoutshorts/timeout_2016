<%
'** Uge knapper **
		g = (w - 1)
		antalWeeks = w
		varleft = 32
		varTop = 129
		
				for w = 0 to w - 1
					if w = g then
					varBColor = "darkRed"
					else
					varBColor = "Lightblue"
					end if
					%>
					<div id="knap_<%=w%>" name="knap_<%=w%>" style="position:absolute; top:<%=varTop%>; left:<%=varleft%>; width:45; z-index:1000; visibility:visible; background-color:#ffffff; !border: 1px; border-color: <%=varBColor%>;  border-bottom:0; border-style: solid;">
					<table border="0" cellpadding="0" cellspacing="0" width="45">
					<tr><td align="center" height="18" width="45">
					<a href="#" onclick="javascript:showweek('<%=w%>', '<%=antalWeeks%>');"><%=weeknumbers(w)%></a></td></tr></table>
					</div>
					<%
					select case w
					case 3, 7, 11, 15, 19, 23, 27, 31, 35, 39, 43, 47, 51
					varTop = varTop + 18
					varleft = 32 - w
					case else
					varTop = varTop
					varleft = varleft + 49
					end select
				next
		
	%>