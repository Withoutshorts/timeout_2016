
<%
'** Lokal **'
'** Brug kun i specel tilfæle eller ved indstallation på egen server **'
'response.redirect "../timeout_xp/wwwroot/ver2_1/login.asp?key=2.2009-1201-TO100&lto=cst"

if len(trim(request("tomobjid"))) <> 0 then
    tomobjid = request("tomobjid")
else
    tomobjid = ""
end if

'** Brug denne på produktions server **'
response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_14/login.asp?key=2.2014-0331-TO1482016&lto=epi_uk&tomobjid="&tomobjid

'** Version lukket ned ***'
'response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
%>



