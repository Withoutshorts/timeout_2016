
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


    
    
    '**** Mobil login *****************************************************************
    userAgent = request.servervariables("HTTP_USER_AGENT")

    if (instr(lcase(userAgent), "iphone") <> 0 OR instr(lcase(userAgent), "iemobile") <> 0 _
    OR instr(lcase(userAgent), "android") <> 0 OR instr(lcase(userAgent), "mobile") <> 0 _ 
    OR inStr(1, userAgent, "iphone", 1) > 0 OR inStr(1, userAgent, "windows ce", 1) > 0 OR inStr(1, userAgent, "blackberry", 1) > 0 OR inStr(1, userAgent, "opera mini", 1) > 0 _
    OR inStr(1, userAgent, "mobile", 1) > 0 OR inStr(1, userAgent, "palm", 1) > 0 OR inStr(1, userAgent, "portable", 1) > 0) _
    AND (instr(lcase(userAgent), "ipad") = 0 OR ((lto = "hestia" OR instr(lto, "epi") <> 0) AND instr(lcase(userAgent), "ipad") <> 0)) then
	'** Iphone **'
    
    browstype_client = "ip"
    else


	if instr(userAgent , "Firefox") <> 0 then
	browstype_client = "mz"
	else
            
            if instr(userAgent , "Chrome") <> 0 then
	        browstype_client = "ch"
	        else

                if instr(userAgent , "Safari") <> 0 then
	            browstype_client = "sf"
	            else
                browstype_client = "ie"
	            end if

            end if

	
	end if

    end if
   
	user_agent_txt = userAgent
    '*********************** END mobil tjek '****************************************



if browstype_client = "ip" then 'mobil direkte til Stempelur
'response.redirect "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/to_2015/monitor.asp?func=startside&key=9K2018-2506-TO181&lto=miele&fromlogin=1"
response.redirect "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/login.asp?key=2.2014-1113-TO156&lto=esn"
else


'** Brug denne på produktions server **'
'response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_14/login.asp?key=2.2014-1113-TO156&lto=esn&tomobjid="&tomobjid&"&usr="&usr
'response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_14/login.asp?key=2.2014-1113-TO156&lto=esn"
response.redirect "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/login.asp?key=2.2014-1113-TO156&lto=esn&tomobjid="&tomobjid&"&usr="&usr

'** Version lukket ned ***'
'response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"

end if

%>



