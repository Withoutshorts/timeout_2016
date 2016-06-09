<!--#include file="../../inc/connection/conn_db_inc.asp"-->

<table cellspacing="0" cellpadding="0" border="0" width="260">
<tr>
	<td>Dato</td>
	<td>Jobnr</td>
	<td>Jobnavn</td>
	<td>Aktivitet</td>
	<td>Timer</td>
	<td>&nbsp;</td>
</tr>
 
<!-------------------------------Sideindhold------------------------------------->
<br>
<!--#include file="convertDate.asp"--> 
<%
	oRec.Open "SELECT MTdato, MTjobnr, MTjobnavn, MTknavn, MTimer, Mtid.id AS MT_id, MTaktivitetId, aktiviteter.navn AS Anavn FROM Mtid, aktiviteter WHERE MTmnavn= '"& Session("user") &"' AND aktiviteter.id = MTaktivitetId", oConn, 3
	Response.write  "<br>"&"Indtastninger: "&"<br><br>"
	
	While not oRec.EOF 
	%>
	<tr>
		<td bgcolor="#cccccc" colspan="6"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr onmouseover="mOvr('gift',this,'lightBlue');" onmouseout="mOut(this,'#fff8dc');">
		<td><font class="tdlysblaa"><%=convertDate(oRec("MTdato"))%></TD>
    	<td align="center"><font class="tdlysblaa"><%=oRec("MTjobnr")%></TD>
		<td><font class="tdlysblaa"><%=oRec("MTjobnavn")%></TD>
		<td><font class="tdlysblaa"><%=oRec("Anavn")%></TD>
		<td align="center"><font class="tdlysblaa"><%=oRec("MTimer")%></TD>
		<td><a href="slet_db.asp?id=<%=oRec("MT_id")%>"><font class="tdslet">Slet</a></td>
	</TR>
	
	<%
	oRec.movenext
	wend
		  
	oRec.Close
%>
<tr>
	<td colspan="6" align="center"><br><br><br><a href='afslut.asp'><img src='ill/afslut.gif' width='200' height='25' alt='' border='0'></a></td>
</tr>
</TABLE>
</body>
</html>
