
<table cellspacing="0" cellpadding="0" border="0" bgcolor="#003399" width=100%>
<tr>
<td><a href="../login.asp"><img src="../ill/logo_topbar.gif" alt="" border="0"></a></td>
</tr>
</table>
<%
level = session("rettigheder")
%>

<%
strSQL = "SELECT * FROM dbversion, licens"
oRec.open strSQL, oConn, 3
if not oRec.EOF then
licenshaver = oRec("licens")

'** tjek om der er tilvalgt CRM ***
len_licensType = instr(oRec("key"), "B")
	if len_licensType > 0 then
	licensType = "CRM"
	end if
end if
oRec.close
%>

		<span name="keys" id="keys" style="position:absolute; top:42; left:5; z-index:0;">
		<table cellpadding="2" cellspacing="0">
		<tr>
		<td valign="top"><img src="../ill/keys.gif" alt="" border="0"></td>
		<td>
		<font style="color:#ffffff;font-family:arial,sans-serif;font-size:11px;">
		<%=session("user")%>&nbsp;&nbsp;<font color="#FFCC00">(<%=licenshaver%>)</font> 
			<%
			strSQL = "SELECT lastlogin FROM medarbejdere WHERE Mnavn ='" & session("user") &"'"
				oRec.open strSQL, oConn, 3
				if not oRec.EOF then
				strDato = oRec("lastlogin")
				end if
				oRec.close
			%>
			&nbsp;<a href="../sesaba.asp" STYLE="font-family : arial, sans-serif; font-size : 11px; color : ffffff;"><strong>Log ud</strong></a> &nbsp;&nbsp;</font>
		</td>
		</tr>
		</table>
		</span>


