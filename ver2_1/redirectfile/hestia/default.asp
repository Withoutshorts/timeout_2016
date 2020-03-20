
<%
'response.buffer = true
'Response.addHeader "Cache-Control", "no-cache, no-store, must-revalidate" ' HTTP 1.1.
'Response.addHeader "Pragma", "no-cache" ' HTTP 1.0.
'Response.addHeader "Expires", "0" ' Proxies.

'** Lokal **'
'** Brug kun i specel tilfæle eller ved indstallation på egen server **'
'response.redirect "../timeout_xp/wwwroot/ver2_1/login.asp?key=2.2009-1201-TO100&lto=cst"

if len(trim(request("tomobjid"))) <> 0 then
    tomobjid = request("tomobjid")
else
    tomobjid = ""
end if

'** Brug denne på produktions server **'
response.redirect "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/login.asp?key=2.2014-0130-TO147&lto=hestia&tomobjid="&tomobjid

%>
<!--
<img src="outzource_logo_200.gif" /><br /><br />
Hej TimeOut brugere<br />
Timeout er flyttet til ny og hurtigere server.<br />
I skal i <a href="https://timeout.cloud/hestia">klikke her for at logge ind.</a> 

<br /><br />
Mvh.
Outzource og Hestia
<br /><span style="font-size:8px;">default</span>
-->

<%'response.end
'response.redirect "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/login.asp?key=2.2014-0130-TO147&lto=hestia&tomobjid="&tomobjid

'** Version lukket ned ***'
'response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
%>



