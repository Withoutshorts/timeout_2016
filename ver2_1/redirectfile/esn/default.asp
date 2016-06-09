
<%
'** Lokal **'
'** Brug kun i specel tilfæle eller ved indstallation på egen server **'
'response.redirect "../timeout_xp/wwwroot/ver2_1/login.asp?key=2.2009-1201-TO100&lto=cst"



if len(trim(request("tomobjid"))) <> 0 then
    tomobjid = request("tomobjid")
else
    tomobjid = ""
end if

if len(trim(request("usr"))) <> 0 then
    'usr = mid(request("usr"), 3, 1) & mid(request("usr"), 11, 1) & mid(request("usr"), 17, 1)
    usr = request("usr") 
else
   usr = ""
end if

'** Brug denne på produktions server **'
'response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_14/login.asp?key=2.2014-1113-TO156&lto=esn&tomobjid="&tomobjid&"&usr="&usr
'response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_14/login.asp?key=2.2014-1113-TO156&lto=esn"
response.redirect "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/login.asp?key=2.2014-1113-TO156&lto=esn&tomobjid="&tomobjid&"&usr="&usr

'** Version lukket ned ***'
'response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
%>



