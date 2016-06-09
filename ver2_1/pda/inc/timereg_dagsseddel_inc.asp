<%

for intcounter = 1 to 7
select case intcounter
case 1
nameandid = "timedetailson"
underopr = "underoprson"
usedatod = varTjDatoUS_son
case 2
nameandid = "timedetailman"
underopr = "underoprman"
usedatod = varTjDatoUS_man
case 3
nameandid = "timedetailtir"
underopr = "underoprtir"
usedatod = varTjDatoUS_tir
case 4
nameandid = "timedetailons"
underopr = "underoprons"
usedatod = varTjDatoUS_ons
case 5
nameandid = "timedetailtor"
underopr = "underoprtor"
usedatod = varTjDatoUS_tor
case 6
nameandid = "timedetailfre"
underopr = "underoprfre"
usedatod = varTjDatoUS_fre
case 7
nameandid = "timedetaillor"
underopr = "underoprlor"
usedatod = varTjDatoUS_lor
end select
%>
<span name=<%=nameandid%> id=<%=nameandid%> style="position:absolute; left:510; top:108; visible:hidden; display:none; border:1px #003399 solid; width:220; height:400; overflow:auto; background-color:#FFFFF1; filter:alpha(opacity=95); z-index:20000; padding-left:3;">
<table border=0 cellpadding=0 cellspacing=1 bgcolor="#fffff1" width=190><tr><td colspan=2>
<b><%=writedagnavn(intcounter)%> d. <%=tjekdag(intcounter)%>:</b>&nbsp;<a href="#" onClick="hidetimedetail()">[X]</a><br>
&nbsp;</td>
</tr>
<%
strSQL = "SELECT timer, tjobnavn, tjobnr, tknavn, tknr, TAktivitetNavn, tfaktim FROM timer WHERE Tmnr = "& intMnr &" AND Tdato='"& usedatod &"' ORDER BY tknavn"
		
	oRec.open strSQL, oConn, 3
	totaltimer = 0
	totalkm = 0
	while not oRec.EOF
	Response.write "<tr><td colspan=2><b>" & oRec("Tknavn") & "</b><br>"
		
		if oRec("tfaktim") <> 0 then
		Response.write oRec("tjobnavn") &"&nbsp;("& oRec("tjobnr")&")"
		end if
	
		Response.write "</td></tr><tr><td width=120>"&oRec("TAktivitetNavn") &"</td><td align=right style='padding-right:5px;' valign=bottom>"& formatnumber(oRec("timer"), 2) &""
		if oRec("tfaktim") <> 5 then 
		Response.write " timer"
		else
		Response.write " km."
		end if
		Response.write "</td></tr>"
		Response.write "<tr><td colspan=2 bgcolor=#003399 height=1><img src='../ill/blank.gif' width='1' height='1' alt='' border='0'></td></tr>"
		
		if oRec("tfaktim") <> 5 then
		totaltimer = totaltimer + oRec("timer") 
		else
		totalkm = totalkm + oRec("timer") 
		end if
	oRec.movenext
	wend
	oRec.close
	Response.write "<tr><td align=right valign=bottom><b>Total:</b>&nbsp;</td><td align=right style='padding-right:5px;'>"& formatnumber(totaltimer,2) &" timer</td></tr>"
	Response.write "<tr><td align=right valign=bottom><b>Km:</b>&nbsp;</td><td align=right style='padding-right:5px;'>"& formatnumber(totalkm,2) &" km</td></tr>"
	Response.write "<tr><td colspan=2 bgcolor=#003399 height=1><img src='../ill/blank.gif' width='1' height='1' alt='' border='0'></td></tr>"
	Response.write "<tr><td colspan=2 bgcolor=#003399 height=1><img src='../ill/blank.gif' width='1' height='1' alt='' border='0'></td></tr>"
	%>
</table>
</span>
<%
'd = d + 1
next%>
