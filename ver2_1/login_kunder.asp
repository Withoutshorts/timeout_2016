<!--#include file="inc/connection/conn_db_inc.asp"-->
<!--#include file="inc/errors/error_inc.asp"-->
<!--#include file="inc/regular/global_func.asp"-->
<%



thisfile = "login_kunder"


session("spmettanigol") = session("spmettanigol") + request("attempt")

	If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
	if year(now) > 2043 then
	%>
	<html>
	<head>
	
	<title>TimeOut 2.1</title>
	<LINK rel="stylesheet" type="text/css" href="inc/style/timeout_style.css">
	</head>
	<body>
	<div id="a" style="position:absolute; left:250; top:220;">
	<table cellspacing="4" cellpadding="2" border="0">
		<tr>
	    	<td colspan="2">TimeOut2.1 er indtil videre ikke udviklet til at kunne køre længere end til og med år 2044. <br>
			Kontakt <a href="mailto:support@outzource.dk">support@outzource.dk</a> for at får opgraderet jeres version af timeOut2.1<br>
			<br>
			Med venlig hilsen<br>
			OutZourCE</td>
		</tr>
	</table>
	</div>
	<%
	else
	'************************************ Login side *********************************************
	%>
	<!--#include file="inc/regular/header_login_inc.asp"-->
	<script language="JavaScript">
	var url="https://outzource.dk/<%=lto%>/default_kunder.asp" 
	var title="TimeOut kundelogin"
	function favorites(){
	if (document.all)
	window.external.AddFavorite(url,title)
	}
	</script>

	<%
	strSQL = "SELECT licens FROM licens WHERE id = 1"
	oRec.Open strSQL, oConn, 0, 1, 1
	if not oRec.EOF then
	licensto = oRec("licens")
	end if
	oRec.close
	
	'*** tjekker om det er PDA eller PC ****
	if instr(request.servervariables("HTTP_USER_AGENT"), "Smartphone") <> 0 then
	pixLeft = 20
	pixTop = 40
	tboxsize = 10
	topgif = ""
	else
	
		select case lto
		case "kringit"
		%>
		<span name="waterm" id="waterm" style="position:absolute; top:142; left:0; z-index:-1000;">
		<img src='ill/kringit/kring_waterm_st.gif' alt='' border='0'>
		</span>
		<%
		'topgif = "<img src='ill/login_timeout.gif' alt='' width='1' height='1' border='0'>"
		pixLeft = 40
		pixTop = 80
		tboxsize = 15
		bgcol = "#ffffff"
		cl = "alm"
		payoff = "<img src='ill/kringit/kring_login_velk_header.gif' alt='' border='0'>"
		footerCl = "lillesort"
		
		case else
		
		'topgif = "<img src='ill/login_timeout.gif' alt='' border='0'>"
		pixLeft = 120
		pixTop = 150
		tboxsize = 15
		bgcol = "#004E90"
		cl = "alt"
		payoff = "<img src='ill/outzource_logo_neg_300.gif' alt='' border='0'>"
		footerCl = "lillehvid"
		end select
		
	end if%>
	
	
	
	<table cellspacing="0" cellpadding="0" border="0" width="100%">
	
	<tr bgcolor="<%=bgcol%>">
	<td><%=payoff%></td>
	</tr>
	<tr bgcolor="#ffffff">
	    <td><img src="ill/blank.gif" width="1" height="5" alt="" border="0"><br /></td>
	</tr>
	</table>
	
    <br /><br /><br /><br /><br /><br /><br /><br />&nbsp;

	<table cellspacing="0" cellpadding="10" border="0" width="600">
	<form action="login_kunder.asp" method="POST">
	<tr><td colspan=2 style="padding-top:10; padding-left:146px;">&nbsp;</td>
		
		
		<td colspan="2" rowspan="2" class='<%=cl%>' valign="top" align=right style="padding-top:20px; padding-right:50px;">
		<%if lto <> "kringit" then
			strSQL = "SELECT useasfak, logo, id, filnavn FROM kunder, filer WHERE useasfak = 1 AND filer.id = kunder.logo"
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
			logonavn = "<img src='inc/upload/"&lto&"/"&oRec("filnavn")&"' alt='' border='0'>"
			else
			logonavn = "<img src='ill/blank.gif' width='150' height='50' alt='' border='0'>"
			end if
			oRec.close
			%><br><br>
			Denne service er leveret af: <b><%=licensto%></b><br><br>
			<%=logonavn%>
		<%else%>
		<img src='ill/blank.gif' width='103' height='60' alt='' border='0'>
		<%end if%>
		</td>
	</tr>
	<tr>
		<td class='<%=cl%>' colspan="2" valign="top" style="padding-left:200px;" align="right">
		
			<table border=0 cellspacing=0 cellpadding=2>
			<tr>
				<td class='<%=cl%>' colspan="2" valign="top" align="right">
				<input type="hidden" name="FM_ansat_kunde" value="2">
				</td>
			</tr>
			<tr>
				<td class='<%=cl%>' valign=top style="padding-top:3px;">Login:</td>
				<td valign="top" align=right><input type="Text" id="login" name="login" value="<%=Request.Cookies("usrname")("usrval")%>"  style="width:150px;"></td>
			</tr>
			<tr>
				<td class='<%=cl%>' valign=top style="padding-top:7px;">Password:</td>
				<td valign="top" align=right><input type="Password" name="pw" value="" style="width:150px;">
				<input type="hidden" name="attempt" value="1">
                   <br /><input id="loginsubmit" type="submit" value="Logind >>" />
			<br><br>
				<!-- GeoTrust QuickSSL [tm] Smart Icon tag. Do not edit. -->
                <!--<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="//smarticon.geotrust.com/si.js"></SCRIPT>-->
                <!-- end GeoTrust Smart Icon tag -->
			<br>
			
			
			
			</td>
			</tr></form>
			</table>
		
		</td>
		</tr>
		</table>
		
		
		<br><br>&nbsp;
		
	
		
		
		
		<%
	
	end if
	
	
Else
	'************ Hvis der er brugt mere end 3 login forsøg *********
	if session("spmettanigol") > 3 then
	%>
	<!--#include file="inc/regular/header_login_inc.asp"-->
	<%
	errortype = 25
	call showError(errortype)
	else
	varLogin = Request.Form("login")
	varPW = Request.Form("pw")
	
	If varlogin = "" Then
	%>
	<!--#include file="inc/regular/header_login_inc.asp"-->
	<%
	errortype = 1
	call showError(errortype)
	
	Else



	'***************** login og pw modtages og valideres ***********************************************
		If varPW = "" then
		%>
		<!--#include file="inc/regular/header_login_inc.asp"-->
		<%
		errortype = 2
		call showError(errortype)
		
		Else
		
	    call TimeOutVersion()
    
    	'*** Er det en ansat eller en kunde der logger på?***
		stransatkunde = request("FM_ansat_kunde") 'altid 2 på kundelogin
		
		strSQL = "SELECT id, kundeid, email, password AS pw, navn AS Mnavn FROM kontaktpers WHERE email='"& request.Form("login") &"'"
		oRec.Open strSQL, oConn, 0, 1, 1
		
		'Response.write strSQL
		'Response.flush
		
		if not oRec.EOF Then
		
		if Trim(request.Form("pw")) = Trim(oRec("pw")) then
			
			'*** Sætter sesions var. alt efter kunde der logger på****
			strUsrId = 0
			thisKpid = oRec("id")
			session("user") = Trim(oRec("Mnavn"))
			session("kontaktpersEmail") = oRec("email")
			session("login") = strUsrId
			session("Mid") = oRec("kundeid") 'Kundeid
			kundeId = oRec("kundeid")
			session("rettigheder") = 0
			startside = 2 'joblog_k.asp
			'*******************************************************************************************
		
		
		'********************* Skriver til logfil ********************************************************
		'Response.write request.servervariables("PATH_TRANSLATED")
		'Response.flush
		if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\login_kunder.asp" then
		
      
        		
				Set objFSO = server.createobject("Scripting.FileSystemObject")
					
					Set objF = objFSO.GetFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\log\logfile_timeout_"&lto&".txt")
					Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\log\logfile_timeout_"&lto&".txt", 8)

				
				objF.writeLine(session("user") &chr(009)&chr(009)&chr(009)& date &chr(009)& time&chr(009)& request.servervariables("REMOTE_ADDR"))
				objF.close	
		end if
		'*******************************************************************************************
		
		
		
		'******************* Sætter sidste login dato *********************************************
		strMthNow = month(now)
		
		select case strMthNow
		case 1
		strMth = "Jan. "
		case 2
		strMth = "Feb. "
		case 3
		strMth = "Mar. "
		case 4
		strMth = "Apr. "
		case 5
		strMth = "May "
		case 6
		strMth = "Jun. "
		case 7
		strMth = "Jul. "
		case 8
		strMth = "Aug. "
		case 9
		strMth = "Sep."
		case 10
		strMth = "Okt. "
		case 11
		strMth = "Nov. "
		case 12
		strMth = "Dec. "
		end select
		
		strLastLogin =  day(now)& " " & strMth & " " & year(now)
		session("dato") = strLastLogin
		
		oRec.close
		
		strSQL = "SELECT lastlogin AS lastlogin FROM kontaktpers WHERE id="& thisKpid 
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then
		session("strLastlogin") = oRec("lastlogin")
		end if
		oRec.close
		
		oConn.execute("UPDATE kontaktpers SET lastlogin = '"& strLastLogin &"' WHERE id="& thisKpid  &"")
		
		'*****************************************************************************************
		
		'*********** redirect til den valgte side *************************************************
		Response.Cookies("login")("ansatkunde") = request("FM_ansat_kunde")
		
		Response.Cookies("usrname") = request("login")
		Response.Cookies("login").Expires = date + 365
		
		
		session("spmettanigol") = 0
		
			'*** tjekker om det er PDA eller PC ****
			if instr(request.servervariables("HTTP_USER_AGENT"), "Smartphone") <> 0 then
				if startside = 1 then
				response.redirect "pda/crmkalender.asp?menu=crm&shokselector=1&ketype=e&selpkt=kal"
				else
				response.redirect "pda/pda_timereg.asp"
				end if
			else
				
				'*********************************************************
				'**** Tjekker om det er localhost (udviklingsmaskine) ****
				'**** Hvis nej: SSL 								  ****
				'*********************************************************
				if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\login_kunder.asp" then
					select case lto
					case "spritelab", "kringit", "abusiness", "kits", "outz"
					response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_10/timereg/sdsk.asp?usekview=j&FM_kontakt="&kundeId
				    case else
				    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_10/timereg/joblog_k.asp?func=aft&useKid="&kundeId
				    end select
				else
					response.redirect "timereg/joblog_k.asp?useKid="&kundeId
				end if
				
				
			end if
		'********************************************************************************************
		Else
		%>
		<!--#include file="inc/regular/header_login_inc.asp"-->
		<%
		errortype = 4
		call showError(errortype)
		
		End if
		Else
		%>
		<!--#include file="inc/regular/header_login_inc.asp"-->
		<%
		errortype = 3
		call showError(errortype)
		
		End if
		
		call closeDB
		
		End if
		End if
	End if
	

End if
%>
<!--#include file="inc/regular/footer_login_inc.asp"-->


