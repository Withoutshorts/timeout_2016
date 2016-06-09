

<html>
<head>
	<title>TimeOut</title>
	<LINK rel="stylesheet" type="text/css" href="inc/style/timeout_style_print.css">
</head>
<body topmargin="0" leftmargin="0" class="login">

<div id="error" style="position:absolute; left:200px; top:100px; width:400px; background-color:#ffffff; border:10px #CCCCCC solid; padding:20px;">
<table cellspacing="0" cellpadding="4" border="0" bgcolor="#ffffff" width=100%>
	<tr>
	    <td valign=bottom><h4>Der er opstået en fejl!</h4></td>
	    <td valign="top" align=right >&nbsp;
		
	
	
		
	</td>
	</tr>
	<tr>
	<td colspan=2>
	<%
	if Request.Cookies("loginlto")("lto") <> "" then
    thislto = Request.Cookies("loginlto")("lto")
    else
    thislto = ""
    end if
	%>
	<b>Sessionen er udløbet.</b><br>
Det er nu over 8 timer siden du sidst har brugt TimeOut, sessionen er afbrudt og
forbindelsen lukket ned.<br><br>

Log ind igen ved at <a href="https://outzource.dk/<%=thislto %>" class=vmenu target='_top'>klikke her.</a><br>
Hvis du fortsat har problemer med at logge ind, ring da til OutZourCE på 2684 2000, eller skriv til <a href="mailto:support@outzource.dk" class=vmenu>support@outzource.dk</a>.<br>
<br><br>
Med venlig hilsen<br>
OutZourCE dev. team<br><br><br>
        &nbsp;
	</td>
	</tr>
	
</table>
</div>


<!--#include file="inc/regular/footer_login_inc.asp"-->


