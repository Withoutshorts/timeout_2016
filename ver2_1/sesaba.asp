<%response.buffer = true%>
<!--#include file="inc/connection/conn_db_inc.asp"-->
<!--#include file="inc/regular/global_func.asp"-->


<%

if len(trim(session("mid"))) <> 0 then
uid = session("mid")
else
uid = 0
end if


call erStempelurOn()
               
if cint(stempelur_hideloginOn) = 1 then 'skriv ikke login, men åben komme/gå tom ==> Opdater ikke logind historiken ved logud

'select case lcase(lto)
'case "kejd_pb", "kejd_pb2", "fk", "fk_bpm"

logudDone = 1

'case else
else

if len(request("logudDone")) then
logudDone = request("logudDone")
else
logudDone = 0
end if

end if
'end select




'*** Opdater login_historik (Stempelur) hvis stempel ur er slået til ****
if cint(logudDone) = 0 then


'*** Opdaterer login_historik AUTOmatisk hvis stempelur ikke er slået til ****
LogudDateTime = year(now)&"/"& month(now)&"/"&day(now)&" "& datepart("h", now) &":"& datepart("n", now) &":"& datepart("s", now) 
logudTid = LogudDateTime 'year(loginDato) &"/"& month(loginDato)&"/"& day(loginDato) & " " & logudHH &":"& logudMM & ":00"
	

strSQL = "SELECT id, login FROM login_historik WHERE mid = " & uid &" AND stempelurindstilling <> -1 ORDER BY id DESC"
oRec.open strSQL, oConn, 3 
if not oRec.EOF then

    loginTid = oRec("login") 'year(loginDato) &"/"& month(loginDato)&"/"& day(loginDato) & " " & loginHH &":"& loginMM & ":00"
	
	if len(loginTid) <> 0 AND len(logudTid) <> 0 then

	loginTidAfr = left(formatdatetime(loginTid, 3), 5)
	logudTidAfr = left(formatdatetime(logudTid, 3), 5)

	minThisDIFF = datediff("s", loginTidAfr, logudTidAfr)/60
	
    else

    minThisDIFF = 0

	end if
	
	strSQL2 = "UPDATE login_historik SET logud = '"& logudTid &"', logud_first = '"& logudTid &"', minutter = "& minThisDIFF &" WHERE id = " & oRec("id")
	'Response.Write strSQL2
    'Response.flush	
    oConn.execute(strSQL2)
	
	
end if
oRec.close  
'***********************************************

end if



'*** Bruges på timereg siden ***'
session("forste") = "j"
session.abandon


'response.flush
'Response.Write "<br><br><br><br><br><br><br><br><br><br><br><br><br>&nbsp;<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"

              
				
if lto = "tec" OR lto = "intranet - local" then

    %>
        

<!--#include file="inc/regular/header_login_inc.asp"-->

            <link href='//fonts.googleapis.com/css?family=Open+Sans:400,300,600,700,800' rel='stylesheet' type='text/css'>
            <link href="../inc/menu/css/chronograph_01.less" rel="stylesheet/less" type="text/css" />

       

<body topmargin="0" leftmargin="0" class="login">

<center>
<div id="error" style="position:relative; left:0px; top:250px; width:400px; padding:10px; background-color:#ffffff; border:10px #CCCCCC solid;">
<table cellspacing="0" cellpadding="10" border="0" bgcolor="#ffffff" width=100%>
	<tr>
	    <td valign=bottom style="padding-top:15px;" colspan="2"><h4>Tak for idag.</h4></td>
	   
	</tr>
	<tr>
	<td colspan=2>
	
<b>Du er nu logget ud af TimeOut.</b><br />
Alle dine registreringer er modtaget og sessionen er afsluttet.<br /><br /><br />
<a href="Javascript:window.close()" style="color:red;">[X]</a><a href="Javascript:window.close()" style="color:#5582d2;">
<b>Du kan lukke din browser ned ved at klikke her..</b></a><br /><br /><br />
        &nbsp;
	</td>
	</tr>
	
</table>
</div>
</center>




</body></html>



    <%
else
'Response.Write("<script language=""JavaScript"">window.opener.close();</script>")

'response.write "her"
'response.end
call browsertype()
if browstype_client = "ch" then
        %>
         <br /><br />Tak for dine registreringer Du kan nu lukke browseren. 
          <br /><br />Du bruger Chrome. Derfor lukker dette vindue ikke selv.

        <%
else
Response.Write("<script language=""JavaScript"">window.open('', '_self', '');</script>")
Response.Write("<script language=""JavaScript"">window.close();</script>")
end if

end if

'else				
				
%>
