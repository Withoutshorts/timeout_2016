
<%
'** Lokal **'
'** Brug kun i specel tilf�le eller ved indstallation p� egen server **'
'response.redirect "../timeout_xp/wwwroot/ver2_1/login.asp?key=2.2009-1201-TO100&lto=cst"

'** Brug denne p� produktions server **'

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


if browstype_client = "ip" then 'mobil direkte til Stempelur
'response.redirect "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/to_2015/monitor.asp?func=startside&key=9K2018-2506-TO181&lto=miele&fromlogin=1"
response.redirect "https://timeout.cloud/timeout_xp/wwwroot/ver4_22/login.asp?key=9K2018-2506-TO181&lto=miele"
else
response.redirect "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/login.asp?key=9K2018-2506-TO181&lto=miele"
end if

'** Version lukket ned ***'
'response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/tak_login.asp"
%>


