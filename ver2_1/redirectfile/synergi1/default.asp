<%
'** Lokal **'
'** Brug kun i specel tilf�le eller ved indstallation p� egen server **'
'response.redirect "../timeout_xp/wwwroot/ver2_1/login.asp?key=2.2009-1201-TO100&lto=cst"

'** Brug denne p� produktions server **'

if len(trim(request("tomobjid"))) <> 0 then
    tomobjid = request("tomobjid")
else
    tomobjid = ""
end if

'response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_14/login.asp?key=2.2011-0512-TO128&lto=synergi1&tomobjid="&tomobjid
response.redirect "http://timeout.cloud/timeout_xp/wwwroot/ver2_14/login.asp?key=2.2011-0512-TO128&lto=synergi1&tomobjid="&tomobjid

'** Version lukket ned ***'
'response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
%>


