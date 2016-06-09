
<%
		
func = request("func")
if func = "dbopr" then
	function SQLBless2(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless2 = tmp
	end function


strConnSur = "driver={MySQL};server=192.168.1.33; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_admin;"
Set oConnSur = Server.CreateObject("ADODB.Connection")
oConnSur.open strConnSur

spg1 = SQLBless2(request("FM_spg1"))
spg2 = SQLBless2(request("FM_spg2"))
spg3 = SQLBless2(request("FM_spg3"))
spg4 = SQLBless2(request("FM_spg4"))
spg5 = SQLBless2(request("FM_spg5"))
user = request("FM_email")
useDato = year(now)&"/"&month(now)&"/"&day(now)
lto = request("lto")

strSQL = "INSERT INTO survey_may_2004 (spg1, spg2, spg3, spg4, spg5, dato, user, lto) VALUES "_
&" ('"& spg1 &"', '"& spg2 &"','"& spg3 &"','"& spg4 &"','"& spg5 &"', '"& useDato &"', '"& user &"', '"& lto &"') "
oConnSur.execute(strSQL)

oConnSur.close
Set oConnSur = nothing

Response.redirect "login.asp?dt=1"	

else
%>
<html>
<head>
	<title>TimeOut 2.1</title>
	<LINK rel="stylesheet" type="text/css" href="inc/style/timeout_style.css">
</head>
<body topmargin="0" leftmargin="0" class="login">
<table cellpadding="0" cellspacing="0" border="0" bgcolor="#003399" width="100%">
<tr>
<td><img src="ill/ur.gif" width="152" height="53" alt="" border="0"></td>
</tr>
</table>
<div id="a" style="position:absolute; left:250; top:80;">
<table cellpadding="0" cellspacing="0" border="0" width="350">
<form action="survey.asp?func=dbopr" method="post" name="survey" id="survey">
<input type="hidden" name="lto" value="<%=request("lto")%>">
<tr>
<td class=alt>Svar på disse <b>5 spørgsmål</b> og vær med til at gøre TimeOut til et endnu bedre produkt. 
Du deltager samtidig i lodtrækningen om 3 flasker rødvin og en CD efter eget valg.<br><br>

<b>1) Hvad synes du generelt om TimeOut?</b><br>
<textarea cols="40" rows="3" name="FM_spg1"></textarea><br><br>
<b>2) Hvad er den bedste funktion i TimeOut, eller hvad synes du der fungerer bedst?</b><br>
<textarea cols="40" rows="3" name="FM_spg2"></textarea><br><br>
<b>3) Hvad er det dårligeste ved TimeOut?</b><br>
<textarea cols="40" rows="3" name="FM_spg3"></textarea><br><br>
<b>4) Hvilken funktion mangler i TimeOut?</b><br>
<textarea cols="40" rows="3" name="FM_spg4"></textarea><br><br>
<b>5) Hvordan synes du drift og support fungerer?</b><br>
<textarea cols="40" rows="3" name="FM_spg5"></textarea><br><br>
<b>Email:</b><input type="text" name="FM_email" size=20 value=""><br>
Hvis du vil være anonym kan du lade være med at udfylde email.
<br><br>
<input type="submit" value="Send spørgeskema!">
</form>
Spørgeskema undersøgelsen løber frem til den 1/7-2004. Vinderen får direkte besked.<br><br>
Med venlig hilsen<br>
OutZourCE<br><br><br>&nbsp;
</td>
</tr>
</table>
</div>
<!--#include file="inc/regular/footer_login_inc.asp"-->
<%end if%>

