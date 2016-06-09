<!--#include file="inc/regular/header_lysblaa_inc.asp"-->



<body topmargin="0" leftmargin="0" class="login">
<table cellpadding="0" cellspacing="0" border="0" bgcolor="#003399" width="100%">
<tr>
<td><img src="ill/ur.gif" width="152" height="53" alt="" border="0"></td>
</tr>
</table>

<div id="a" style="position:absolute; left:250; top:120; background-color:#ffff99; width:400px; height:200px; padding:20px;">
<table cellpadding="0" cellspacing="0" border="0" width="350">
<tr>
<td class=alt><br />
<img src="ill/login_velk_header.gif" width="543" height="60" alt="" border="0">
<h3>Sessionen er udløbet</h3>
Sessionen er løbet ud, så du er blevet logget af systemet.<br>
<br>Nb: Session udløber ud hvis systemet ikke bliver brugt i 8 timer.
<br /><br />
<%
key = Request.Cookies("login")("lto")
%>

Klik <a href="login.asp?key=<%=key %>" target="_top">her</a> for at logge ind i jeres system igen.
<br /><br />
Med venlig hilsen<br />
OutZourCE dev. team.

</td>
</tr>
</table>
</div>
<!--#include file="inc/regular/footer_inc.asp"-->


