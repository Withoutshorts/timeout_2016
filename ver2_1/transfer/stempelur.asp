<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->
<!--#include file="inc/smiley_inc.asp"-->


<%


                            

 function nyePauser(p, thison)
            
            'Response.Write "her"
            
            if (thison = 1) then
             %>
             <br /><br /><b>Pause <%=p %>:</b> 
            <select name="p<%=p %>" id="p<%=p %>" style="font-family:arial; font-size:9px;">
            
             <% 
            
             if p = 1 then
             caseVal = stPause1
             else
             caseVal = stPause2
             end if
            
             
             Select case caseVal
             case 0
             selM0 = "SELECTED"
             selM10 = ""
             selM15 = ""
             selM20 = ""
             selM30 = ""
             case 10
             selM0 = ""
             selM10 = "SELECTED"
             selM15 = ""
             selM20 = ""
             selM30 = ""
             case 15
             selM0 = ""
             selM10 = ""
             selM15 = "SELECTED"
             selM20 = ""
             selM30 = ""
             case 20
             selM0 = ""
             selM10 = ""
             selM15 = ""
             selM20 = "SELECTED"
             selM30 = ""
             case 30
             selM0 = ""
             selM10 = ""
             selM15 = ""
             selM20 = ""
             selM30 = "SELECTED"
             end select
             
               
             
             %>
             <option value="0" <%=selM0%>>0 min</option>
             <option value="10" <%=selM10%>>10 min</option>
             <option value="15" <%=selM15%>>15 min</option>
              <option value="20" <%=selM20%>>20 min</option>
              <option value="30" <%=selM30%>>30 min</option>
              
            </select> min.<br />
        
            
            Kommentar pause <%=p %>:
	        <br /><textarea id="FM_komm_p<%=p %>" name="FM_komm_p<%=p %>" style="font-size:9px; font-family:arial; width:350px; height:46px;"></textarea>
	        <%
            end if
            
            end function
            


sub crmaktstatheader
%>
<table border=0 cellpadding=0 cellspacing=0 width="450">
	<tr>
	<td valign="top" width="163"><img src="../ill/logo_bg.gif" width="163" height="53" alt="" border="0"></td>
	<td valign="bottom"><b>Timeout Kontrolpanel - Stempelur indstillinger</b><br>
	Tilføj, fjern eller rediger Stempelur indstillinger.</td>
	</tr>
	</table><br>
<%
end sub


if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	if len(request("lastid")) <> 0 then
	lastID = request("lastid")
	else
	lastID = 0
	end if 
	
	if len(request("showonlyone")) <> 0 then
	showonlyone = request("showonlyone")
	else
	showonlyone = 0
	end if
	
	if len(request("medarbSel")) <> 0 then
	medarbSel = request("medarbSel")
	else
	medarbSel = 0
	end if
	
	if showonlyone = 1 then
	showalle = 0
	medarbSQLkri = " AND m.mid = " & medarbSel
	else
	showalle = 1
	medarbSQLkri = " AND m.mid <> 0 "
	end if
	
	if len(request("hidemenu")) <> 0 then
	hidemenu = request("hidemenu")
	else
	hidemenu = 0
	end if
	
	rdir = request("rdir")
		
	thisfile = "stempelur"
	
	select case func
	case "sletlogind"
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	<%
	slttxt = "<b>Slet dette logind?</b><br />"_
	&"Du er ved at slette et logind. Er dette korrekt?"
	slturl = "stempelur.asp?menu="&request("menu")&"&func=sletlogindok&id="&id&"&medarbSel="&medarbSel&"&showonlyone="&showonlyone&"&hidemenu="&hidemenu&"&rdir="&rdir
	
	call sltquePopup(slturl,slttxt,slturlalt,slttxtalt,20,10)
	
	case "sletlogindok"
	
	
	strSQLdel = "DELETE FROM login_historik WHERE id = "& id
	oConn.execute(strSQLdel)
	
	select case rdir
	case "treg"
	Response.Write("<script language=""JavaScript"">window.opener.location.href('timereg_akt_2006.asp?showakt=3');</script>")
    case else
	Response.Write("<script language=""JavaScript"">window.opener.location.href('stempelur.asp?menu=stat&func=stat&medarbSel="&medarbSel&"&showonlyone="&showonlyone&"&hidemenu="&hidemenu&"');</script>")
    end select						
	
	Response.Write("<script language=""JavaScript"">window.close();</script>")
	
	
	case "slet"
	'*** Her spørges om det er ok at der slettes en medarbejder ***
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20; top:10; visibility:visible;">
	<%call crmaktstatheader%>
	<table cellspacing="2" cellpadding="2" border="0">
	<tr>
	    <td>Du er ved at <b>slette</b> en Stempelur indstilling. Er dette korrekt?</td>
	</tr>
	<tr>
	   <td><a href="stempelur.asp?menu=tok&func=sletok&id=<%=id%>">Ja</a>&nbsp;&nbsp;&nbsp;<a href="javascript:history.back()">Nej</a></td>
	</tr>
	</table>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%
	case "sletok"
	'*** Her slettes en stempelurs indstilling ***
	oConn.execute("DELETE FROM stempelur WHERE id = "& id &"")
	Response.redirect "stempelur.asp?menu=tok&shokselector=1"
	
	case "dbopr", "dbred"
	'*** Her indsættes en ny type i db ****
		if len(request("FM_navn")) = 0 then
		%>
		<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
		<%
		errortype = 8
		call showError(errortype)
		
		else
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function
		
		strNavn = SQLBless(request("FM_navn"))
		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
		
		if len(request("FM_faktor")) <> 0 then
		intFaktor = replace(request("FM_faktor"), ",", ".")
		else
		intFaktor = 0
		end if
		
		if len(request("FM_minimum")) <> 0 then
		intMinimum = replace(request("FM_minimum"), ",", "") 
		intMinimum = replace(intMinimum, ".", "")
		else
		intMinimum = 0
		end if
		
		if len(request("FM_forvalgt")) <> 0 then
		intForvalgt = request("FM_forvalgt")
		else
		intForvalgt = 0
		end if
		
			if intForvalgt = 1 then
			oConn.execute("UPDATE stempelur SET forvalgt = 0")
			end if
		
		if func = "dbopr" then
		oConn.execute("INSERT INTO stempelur (navn, editor, dato, faktor, forvalgt, minimum) VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', "& intFaktor &", "& intForvalgt &", "& intMinimum &")")
		else
		oConn.execute("UPDATE stempelur SET navn ='"& strNavn &"', "_
		&" forvalgt = "& intForvalgt &", editor = '" &strEditor &"', dato = '" & strDato &"', faktor = "& intFaktor &", minimum = "& intMinimum &" WHERE id = "&id&"")
		end if
		
		
		Response.redirect "stempelur.asp?menu=tok&shokselector=1"
		end if
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	
	else
	strSQL = "SELECT navn, editor, dato, faktor, forvalgt, minimum FROM stempelur WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	intFaktor = oRec("faktor")
	intForvalgt = oRec("forvalgt")
	intMinimum = oRec("minimum")
	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "opdaterpil" 
	end if
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20; top:10; visibility:visible;">
	<%call crmaktstatheader%>
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr><form action="stempelur.asp?menu=tok&func=<%=dbfunc%>" method="post">
	<input type="hidden" name="id" value="<%=id%>">
    	<td valign="top" colspan="2"><font class="pageheader">Kontrolpanel | Stempelur indstillinger | <%=varbroedkrumme%></font></td>
	</tr>
	<%if dbfunc = "dbred" then%>
	<tr>
		<td colspan="2" valign="bottom" style="height:30;">Sidst opdateret den <b><%=formatdatetime(strDato, 0)%></b> af <b><%=strEditor%></b></td>
	</tr>
	<%end if%>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td><b>Navn:</b></td>
		<td><input type="text" name="FM_navn" value="<%=strNavn%>" size="30" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td><b>Omregnings faktor:</b></td>
		<td><input type="text" name="FM_faktor" value="<%=intFaktor%>" size="4" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
	</tr>
	<tr>
		<td valign=top style="padding-top:3px;"><b>Minimums indtastning:</b></td>
		<td><input type="text" name="FM_minimum" value="<%=intMinimum%>" size="4" style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;">&nbsp;<b>(Angivet i hele min.)</b><br>
		<font class=lillesort>Hvis minimum er sat til 60 min.<br> 
		vil alle login's på mindre end 60 min. blive takseret til 1 time.</font></td>
	</tr>
	<tr>
		<td colspan=2>
		<%
		if intForvalgt = 1 then
		chk = "CHECKED"
		else
		chk = ""
		end if
		%>
		<br><input type="checkbox" name="FM_forvalgt" value="1" <%=chk%>>&nbsp;<b>Denne indstilling skal stå som forvalgt ved login.</b></td>
	</tr>
	<tr>
		<td colspan="2"><br><br><img src="ill/blank.gif" width="100" height="1" alt="" border="0"><input type="image" src="../ill/<%=varSubVal%>.gif"></td>
	</tr>
	</form>
	</table>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	
	<%
	case "dbloginhist"
	
	
	'**** ignorer tidsinterval / fast st. tid *********
	
	stpause = -1
    stpause2 = -1
	
	strSQL = "SELECT ignorertid_st, ignorertid_sl, stpause, stpause2 FROM licens WHERE id = 1"
	oRec.open strSQL, oConn, 3
    if not oRec.EOF then
    
    ig_sttid = oRec("ignorertid_st")
    ig_sltid = oRec("ignorertid_sl")
    
    stpause = oRec("stpause")
    stpause2 = oRec("stpause2")
   
    end if
    oRec.close
    
    if len(ig_sttid) <> 0 then
    ig_sttid = formatdatetime(ig_sttid, 3)
    end if
    
    if len(ig_sltid) <> 0 then
    ig_sltid = formatdatetime(ig_sltid, 3)
    end if
    
   
    
	
	ids = split(request("id"), ",")
	useMid = split(request("mid"), ",")
	loginDato = split(request("logindato"), ",")
	'Response.Write request("logindato")
	'Response.end
	
	
	'*** Logind og logud tider ****'
    login_hh = replace(request("FM_login_hh"), ",","")
	login_mm = replace(request("FM_login_mm"), ",","")
    logud_hh = replace(request("FM_logud_hh"), ",","")
	logud_mm = replace(request("FM_logud_mm"), ",","")
	
	login_hh = split(login_hh, " #")
	login_mm = split(login_mm, " #")
	logud_hh = split(logud_hh, " #")
	logud_mm = split(logud_mm, " #")
	

	
	
	oprlogin_hh = split(request("oprFM_login_hh"), ",")
	oprlogin_mm = split(request("oprFM_login_mm"), ",")
	
	oprlogud_hh = split(request("oprFM_logud_hh"), ",")
	oprlogud_mm = split(request("oprFM_logud_mm"), ",")
	
	'kommThis = split(request("FM_kommentar"), ",")
    'Response.Write "K "& request("FM_kommentar")
	'Response.Flush
	
	stur = split(request("FM_stur"), ",")
	
	ipn = request.servervariables("REMOTE_ADDR")
	
	if len(request("p1")) <> 0 then
    p1 = request("p1")
    else
    p1 = 0
    end if
    
    if len(request("p2")) <> 0 then
    p2 = request("p2")
    else
    p2 = 0
    end if
    
    p1on = request("p1on")
    'Response.Write "request(p2on) " & request("p2on")
    p2on = request("p2on")
    'Response.Write "p2on " & p2on
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="inc/isint_func.asp"-->
	<%
	
	                function SQLBless2(s)
				    dim tmp
				    tmp = s
				    tmp = replace(tmp, "'", "''")
				    SQLBless2 = tmp
				    end function
	
	
	
	'**** Pauser *****
	p1_komm = SQLBless2(trim(request("FM_komm_p1")))
    p2_komm = SQLBless2(trim(request("FM_komm_p2")))
    
    

	useleftdiv = "c"
	
	
	
	for a = 0 to UBOUND(ids) 
	
	
	if len(trim(loginDato(a))) = 0 OR len(trim(login_hh(a))) = 0 OR len(trim(login_mm(a))) = 0 then
		'OR len(trim(logud_hh(a))) = 0 OR _
	    'len(trim(logud_mm(a))) = 0
			
			errortype = 60
			call showError(errortype)
			Response.end
			
			else
			
			
			'Response.Write logud_hh(a)
			'Response.end
			
			'call erDetInt(trim(logud_hh(a)))
			'if cint(isInt) > 0 OR instr(logud_hh(a), ".") <> 0 OR logud_hh(a) > 23 then
				
			'	errortype = 58
			'	call showError(errortype)
			'	Response.end
			
			'isInt = 0
			
			'else
					
					if isdate(trim(loginDato(a))) <> true then 
					'OR cint(len(trim(loginDato(a)))) <> 10 then
					
					'Response.Write "loginDato(a)" & loginDato(a) & " len:" & cint(len(trim(loginDato(a))))
					'Response.end
					
					errortype = 66
					call showError(errortype)
					Response.end    
					        
					       
					
					else
			        
			        
					'call erDetInt(trim(logud_mm(a))) 
					'if cint(isInt) > 0 OR instr(trim(logud_mm(a)), ".") <> 0 OR logud_mm(a) > 59 then
						
					'	errortype = 58
					'	call showError(errortype)
					'	Response.end
						
                    'isInt = 0
					
					'else
					        
					        'Response.Write "login_hh(a) " & login_hh(a) & "<br>"
							'Response.flush
					        
							call erDetInt(trim(login_hh(a))) 
							if cint(isInt) > 0 OR instr(login_hh(a), ".") <> 0 OR login_hh(a) > 23 then
								
								
								
								errortype = 59
								call showError(errortype)
								Response.end
							
							isInt = 0
							
							else
							
									call erDetInt(trim(login_mm(a))) 
									if isInt > 0 OR instr(login_mm(a), ".") <> 0 OR login_mm(a) > 59 then
										
										errortype = 59
										call showError(errortype)
										Response.end
									
									isInt = 0
									
									else
					
					
					loginHH = left(trim(login_hh(a)), 2)
					loginMM = left(trim(login_mm(a)), 2)	
	                
	                '** Tjekker om der er angivet et logud tidspunkt ***'
	                if len(trim(logud_hh(a))) <> 0 AND len(trim(logud_mm(a))) <> 0 then
					logudHH = left(trim(logud_hh(a)), 2)
					logudMM = left(trim(logud_mm(a)), 2)
					else
					logudHH = "99"
					logudMM = "99"
					end if
					
					oprloginHH = left(trim(oprlogin_hh(a)), 2)
					oprloginMM = left(trim(oprlogin_mm(a)), 2)	
	
					oprlogudHH = left(trim(oprlogud_hh(a)), 2)
					oprlogudMM = left(trim(oprlogud_mm(a)), 2)
					
					
					
					'****** Kommentar **************'
					'komm = kommThis(a)
					komm = SQLBless2(trim(request("FM_kommentar_"&a&"")))
					
					
					
					
					
					'*******************************************
					'*** Test for ignorer periode v. login *****
					testloginTid = year(loginDato(a)) &"/"& month(loginDato(a))&"/"& day(loginDato(a)) & " " & loginHH &":"& loginMM & ":00"
					testig_sttid = year(loginDato(a)) &"/"& month(loginDato(a))&"/"& day(loginDato(a)) & " " & ig_sttid
					testig_sltid = year(loginDato(a)) &"/"& month(loginDato(a))&"/"& day(loginDato(a)) & " " & ig_sltid
					
					'Response.Write cdate(testloginTid) &" > "& cdate(testig_sttid) &" AND "& cdate(testloginTid) &" < "& cdate(testig_sltid) &" then <br>"
					
					use_ig_sltid = 0
					
					if cdate(testloginTid) > cdate(testig_sttid) AND cdate(testloginTid) < cdate(testig_sltid) then
					use_ig_sltid = 1
					loginTid = year(loginDato(a)) &"/"& month(loginDato(a))&"/"& day(loginDato(a)) & " " & ig_sltid
					else
					use_ig_sltid = 0
					loginTid = year(loginDato(a)) &"/"& month(loginDato(a))&"/"& day(loginDato(a)) & " " & loginHH &":"& loginMM & ":00"
					end if
					'Response.Write loginTid & "<hr>"
					
					
					'*** Skal ignorer periode SLUT på her ***? 
				    logudTid = year(loginDato(a)) &"/"& month(loginDato(a))&"/"& day(loginDato(a)) & " " & logudHH &":"& logudMM & ":00"
					'Response.Write logudTid &"<br>"
					
					'******************************************
					
					
					
					
					
					'************ ?? ***********
					if a = 0 then
					loginTidp = year(loginDato(a)) &"/"& month(loginDato(a))&"/"& day(loginDato(a)) & " 00:00:00"
					logudTidp = year(loginDato(a)) &"/"& month(loginDato(a))&"/"& day(loginDato(a)) & " 00:00:00"
					loginDTp = year(loginDato(a)) &"/"& month(loginDato(a))&"/"& day(loginDato(a))
					end if
					'************************
					
					
					
					            
					            
					            '**** Beregner Minutter mellem logind og logud ***********
					            if len(loginTid) <> 0 AND len(logudTid) <> 0 then
					                    
					                    '*** Er der angivet et logud ****'
					                    if logudHH <> "99" then
                     					
		                                loginTidAfr = left(formatdatetime(loginTid, 3), 5)
					                    logudTidAfr = left(formatdatetime(logudTid, 3), 5)
                    		
					                    minThisDIFF = datediff("s", loginTidAfr, logudTidAfr)/60
                    					
					                    else
                    					
					                    minThisDIFF = 0
					                    end if
					                    
					            end if
					            '***
					        
					        
					        '*** er logud tid mindre end logind tidspunkt **'
					        if minThisDIFF < 0 then
					        
					            errortype = 139
								call showError(errortype)
								Response.end
					        
					        
					        
					        else
					        
					        
					        
					        '** Er logud tid ændret (Ved rediger) ?? **************'
					        '*** Eller er det første logud på dette logind? *******'    
					        logudrettetJa = 0
					        'manuelt_afsluttet = 0
					        tiderRettet = 0
					       
					        strSQLA = "SELECT logud, login, manuelt_afsluttet, login_first, logud_first FROM login_historik WHERE logud IS NOT NULL AND id = " & ids(a)
					        oRec.open strSQLA, oConn, 3
                            if not oRec.EOF then
                                oprLogud = oRec("logud_first")
                                oprLogin = oRec("login_first")
                                logudrettetJa = 1
                                manuelt_afsluttet = oRec("manuelt_afsluttet")
                            end if
                            oRec.close
                            
                            if cint(logudrettetJa) = 1 then
                            len_oprLogud = len(oprLogud)
                                
                                if len_oprLogud <> 0 then
                                oprLogud_left = left(oprLogud, len_oprLogud- 3)
                                oprLogud = oprLogud_left &":00"
                                else
                                oprLogud = year(loginDato(a)) &"/"& month(loginDato(a))&"/"& day(loginDato(a)) & " " & oprlogud_hh(a) &":"& oprlogud_mm(a) &":00"
                                end if
                            
                            
                            len_oprLogin = len(oprLogin)
                                if len_oprLogin <> 0 then
                                oprLogin_left = left(oprLogin, len_oprLogin - 3)
                                oprLogin = oprLogin_left &":00"
                                else
                                oprLogin = year(loginDato(a)) &"/"& month(loginDato(a))&"/"& day(loginDato(a)) & " " & oprlogin_hh(a) &":"& oprlogin_mm(a) &":00"
                                end if
                            'Response.Write "her"
                            
                            else
                            'oprLogud = "00:00:00"
                            oprLogud = year(loginDato(a)) &"/"& month(loginDato(a))&"/"& day(loginDato(a)) & " " & oprlogud_hh(a) &":"& oprlogud_mm(a) &":00"
                            oprLogin = year(loginDato(a)) &"/"& month(loginDato(a))&"/"& day(loginDato(a)) & " " & oprlogin_hh(a) &":"& oprlogin_mm(a) &":00"
                            'Response.Write "der loginDato(a):" & loginDato(a) & "<br>"
                             end if
                            
                          
                            '*** use_ig_sltid = ignorer periode ***
                            
                            '** er der angivet logud ***'
                            if logudHH <> "99" then
                            'Response.Write "her"
                            'Response.Write oprLogin& " - " & loginTid & "OR logud: "& oprLogud &" <>" & logudTid & "<br>"
                            if (cDate(oprLogin) <> cDate(loginTid)) OR (cdate(oprLogud) <> cdate(logudTid)) then
                            tiderRettet = 1
                            end if
                            
                            else
                                 '*** Ellers tjekkes kun logind tid **'
                                 if (cDate(oprLogin) <> cDate(loginTid)) then
                                 tiderRettet = 1
                                 end if
                            
                            end if
					
					        
					        
					       
					
					'Response.end
					
					
					'*** Validere at login og logud tider er blevet ændret ved rediger. **'
					'*** I så fald skal der angives en kommentar ***
				    if cint(tiderRettet) = 1 OR (lto = "dencker" AND (stur(a) = 5 OR stur(a) = 6)) OR (lto = "outz" AND stur(a) = 2) then
				    
				    'Response.Write "oprLogud: " & cDate(oprLogud) & "<br>"
					'Response.Write "logud nu: " & cDate(logudTid) & "<br>"
					
					'Response.Write "oprLogin: " & cDate(oprLogin) & "<br>"
					'Response.Write "login nu: " & cDate(loginTid)
					
					'Response.end
					
					    if len(trim(komm)) = 0 AND use_ig_sltid <> 1 then
    					
					    errortype = 130
					    'if lto = "outz" then
					    'Response.Write "stur: "& stur(a) & "tiderRettet "& tiderRettet
					    'call showError(errortype)
					    'else
					    call showError(errortype)
					    'end if
					    
					    Response.end
    					
					    end if
					
					end if
					'************************************************************************
					
					
					    '*****************************************************************'
				        '*** Skal der afsendes mail (Ved ændring af logind / logud tid)??? 
				        '*** Hvis logud er manuelt indtastet fra logud stat siden. *** 
				        if (logudrettetJa = 1 AND tiderRettet = 1) AND use_ig_sltid <> 1 OR (lto = "dencker" AND stur(a) = 5 OR stur(a) = 6) OR (lto = "outz" AND stur(a) = 2) then
				        
    					
				        	
					                        '***** Oprettter Mail object ***''
			                                if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\stempelur.asp" then
                                			
                                			
			                                Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
			                                ' Sætter Charsettet til ISO-8859-1 
			                                Mailer.CharSet = 2
			                                ' Afsenderens navn 
			                                Mailer.FromName = "TimeOut Stempelur Service"
			                                ' Afsenderens e-mail 
			                                Mailer.FromAddress = "timeout_no_reply@outzource.dk"
			                                Mailer.RemoteHost = "webmail.abusiness.dk"
			                                Mailer.ContentType = "text/html"
                                		    
                        		            '*** Finder medarbejder ***'
                        		            mnavn = "-"
                        		            strSQLmedarb = "SELECT mnavn, mnr FROM medarbejdere WHERE mid = "& medarbSel
                                		    
                        		            'Response.Write strSQLmedarb
                        		            'Response.flush
                                		    
                        		            oRec.open strSQLmedarb, oConn, 3
                        		            if not oRec.EOF then
                                		    
                        		            mnavn = oRec("mnavn") & " ("& oRec("mnr") &")"
                                		    
                        		            end if
                        		            oRec.close
                                		    
                                		    '*** Type navn ***'
                                		    stempelur_typenavn = ""
                                		    strSQL_stemptypnavn = "SELECT navn FROM stempelur WHERE id = "& stur(a)
                                		    oRec.open strSQL_stemptypnavn, oConn, 3
                                		    
                                		    if not oRec.EOF then
                                		    
                                		    stempelur_typenavn = oRec("navn")
                                		    
                                		    end if
                                		    oRec.close
                                		    
        		                            
                                            'Mailens emne
                                            if (lto = "dencker" AND (stur(a) = 5 OR stur(a) = 6)) OR (lto = "outz" AND stur(a) = 2) then
                                            Mailer.Subject = "Medarbejder "& mnavn &" har haft et "& stempelur_typenavn  
                                            else
			                                Mailer.Subject = "Medarbejder "& mnavn &" har ændret stempelur tid. (type: "& stempelur_typenavn &")" 
        			                        end if
        			                        
			                                'Modtagerens navn og e-mail
			                                select case lto
			                                case "dencker" 
			                                    
			                                    if (lto = "dencker" AND (stur(a) = 5 OR stur(a) = 6)) then
			                                    Mailer.AddRecipient "Karen", "krt@dencker.net"
			                                    end if
			                                    
			                                    Mailer.AddRecipient "Anders Dencker", "ad@dencker.net"
			                                    
			                                case "outz"
			                                Mailer.AddRecipient "OutZourCE", "sk@outzource.dk"
			                                case "cst"
			                                Mailer.AddRecipient "C.S.T", "th@copst.com"
			                                case else
			                                Mailer.AddRecipient "OutZourCE", "timeout_no_reply@outzource.dk"
			                                end select
                                			
        			                       
                                			
			                                        ' Selve teksten
					                                Mailer.BodyText = "Der er blevet ændret i logind eller logud tid, eller medarbejderen har brugt en logind-indstilling der skal begrundes."_
					                                &"<br><br><b>Type:</b> "& stempelur_typenavn &""_
					                                &"<br><br><b>Medarbejder:</b><br>"& mnavn &"<br><br><b>Logind dato og klokkeslet:</b> "& formatdatetime(loginTid, 1) &" kl. "& formatdatetime(loginTid, 3) &" til "& formatdatetime(logudTid, 3) &"<br><br>"_ 
					                                &"<b>Kommentar:</b><br> "_
					                                &""& komm &"<br><br>"_
					                                &"Med venlig hilsen<br><br>TimeOut Stempelur Service<br><br>&nbsp;" 
        					                        
					                                If Mailer.SendMail Then
                                    				
					                                Else
    					                                Response.Write "Fejl...<br>" & Mailer.Response
  					                                End if	
                                				
                                			
			                                end if ''** Mail
        					        
        					
        					
					       
				      end if
					
					
					
					thisDato = year(loginDato(a)) &"/"& month(loginDato(a)) &"/"& day(loginDato(a))
					oprLogud_first = thisDato &" "& formatdatetime(oprLogud, 3)
					
					        
					        '*** Findes der allerede en indtastning på denne dato i dette tidsrum ****'
					        
					        strSQL = "SELECT dato, login, logud FROM login_historik WHERE mid = "& useMid(a) &" AND dato = '"& thisDato &"'"_
					        &" AND stempelurindstilling <> -1 AND id <> "& ids(a) &" AND minutter <> 0 AND ((login < '"& loginTid &"' AND logud > '"& loginTid &"') OR (login < '"& logudTid &"' AND logud > '"& logudTid &"') "_
					        &" OR (login > '"& loginTid &"' AND logud < '"& logudTid &"')) "
					        
					        'Response.Write strSQL & "<br>"
					        'Response.flush
					        
					        oRec.open strSQL, oConn, 3
					        if not oRec.EOF then
					        strLogindkonflikt = oRec("login") & " til " & oRec("logud") 
					        errortype = 134
					        call showError(errortype)
					        Response.end
					        
					        end if
					        oRec.close
					        
					
					    '*** er logud angivet`eller tomt? ***'
					    if logudHH <> "99" then
					    logudTid = logudTid 
					    else
					    logudTid = ""
					    end if
					
					
					if ids(a) <> 0 then
					
					    
					   
					        
					    strSQL = "UPDATE login_historik SET dato = '"& thisDato &"', login = '"& loginTid &"', logud = '"& logudTid &"', "_
					    &" stempelurindstilling = "& stur(a) &", minutter = "& minThisDIFF &", "_
					    &" manuelt_afsluttet = "& tiderRettet &", kommentar = '"& komm &"'"
    					
					    '*** Er det første logud ***?
					    '*** Ellers må logud_fisrt og login_fisrt ikke rettes ***'
					    if logudrettetJa <> 1 AND logudHH <> "99" then
					    strSQL = strSQL & ", logud_first = '"& oprLogud_first &"'"
					    end if
					
					
					    strSQL = strSQL & " WHERE id = " & ids(a)
					
					
					'Response.end
					
					else
					
					        
					        '*** Dette er alligevel ok, 
					        '*** dog tjekkes der for tidspunkt ovenfor ***'
					        '*** Deak 7/7-2009 **'
					        
					        '*** Findes der allerede en indtastning på denne dato ****'
					        'strSQL = "SELECT dato FROM login_historik WHERE mid = "& useMid(a) &" AND dato = '"& thisDato &"'"
					        'oRec.open strSQL, oConn, 3
					        'if not oRec.EOF then
					        
					        'errortype = 133
					        'call showError(errortype)
					        'Response.end
					        
					        'end if
					        'oRec.close
					        
	                         
					
					strSQL = "INSERT INTO login_historik  "_
					&"(dato, login, mid, stempelurindstilling, minutter, "_
					&" login_first "
					
					if logudHH <> "99" then 
					strSQL = strSQL & ", logud, logud_first"
					end if
					
					strSQL = strSQL & ", manuelt_oprettet, manuelt_afsluttet, kommentar, ipn) "_
					&" VALUES ('"& thisDato &"', '"& loginTid &"', "_
					&" "& useMid(a) &", "& stur(a) &", "& minThisDIFF &", '"& loginTid &"'"
					
					if logudHH <> "99" then 
					strSQL = strSQL & ", '"& logudTid &"', '"& logudTid &"'"
					end if
					
					
					strSQL = strSQL & ", 1, 0, '"& komm &"', '"& ipn &"')"
				    
				    
				    'Response.write strSQL & "<br>"
					'Response.end
					
					end if
					
					'Response.write strSQL & "<br>"
					'Response.end
				    oConn.execute(strSQL)
				    
				    
		            idLast = ids(s)
		    
		    
					
					       
	                        
	                        
							    end if '** validering
						    end if
						end if
					'end if
				'end if 
			end if
		end if
	    
	   next
	                        
	                        
	                        
	                        
	                        
	                        '*** Tilføj Pauser *****
	                        strSQLpDel = "DELETE FROM login_historik WHERE stempelurindstilling = -1  AND dato = '"& loginDTp &"' AND mid = "& medarbSel
	                        oConn.execute(strSQLpDel)
	                        
	                        'Response.Write " p1on " & p1on
	                        
	                        if p1on <> 0 then
	                        
	                        strSQLp = "INSERT INTO login_historik SET dato = '"& loginDTp &"', "_
	                        &" login = '"& loginTidp &"', "_
	                        &" logud = '"& logudTidp &"', "_
					        &" stempelurindstilling = -1, minutter = "& p1 &", "_
					        &" manuelt_afsluttet = 0, kommentar = '"& p1_komm &"', mid = "& medarbSel
					        
					        'Response.Write strSQLp & "<br>"
					        
					        oConn.execute(strSQLp)
					        
					        end if
					        
					        'Response.Write " p2on " & p2on
					        
					        if p2on <> 0 then
					        strSQLp = "INSERT INTO login_historik SET dato = '"& loginDTp &"', "_
	                        &" login = '"& loginTidp &"', "_
	                        &" logud = '"& logudTidp &"', "_
					        &" stempelurindstilling = -1, minutter = "& p2 &", "_
					        &" manuelt_afsluttet = 0, kommentar = '"& p2_komm &"', mid = "& medarbSel
					       
					       
        					 'Response.Write strSQLp & "<br>"
        				       oConn.execute(strSQLp)
	                        
	                        end if
	                        
	                        'Response.end
	                        
	                        
	                        'if rdir <> "sesaba" then
					        'Response.redirect "stempelur.asp?menu=stat&func=stat&lastid="&idLast&"&medarbSel="&medarbSel&"&showonlyone="&showonlyone&"&hidemenu="&hidemenu
	                        'else
	                        'Response.redirect "../sesaba.asp?logudDone=1"
	                        'end if 
	                        
	                        
	                        select case rdir
	                        case "treg"
	                        Response.Write("<script language=""JavaScript"">window.opener.location.href('timereg_akt_2006.asp?showakt=3');</script>")
                            case "sesaba"
                            
                            Response.redirect "../sesaba.asp?logudDone=1"
                            Response.end
                            case else
	                        Response.Write("<script language=""JavaScript"">window.opener.location.href('stempelur.asp?menu=stat&func=stat&medarbSel="&medarbSel&"&showonlyone="&showonlyone&"&hidemenu="&hidemenu&"&lastid="&idLast&"');</script>")
                            end select						
	
	                        Response.Write("<script language=""JavaScript"">window.close();</script>")
	                        
	    
	
	case "redloginhist", "oprloginhist"
	
	
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	
	
	    if rdir = "sesaba" then
	    sideDivTop = 20
	    sideDivLeft = 160
	    opdAfsl = "Stempelur, angiv logind og logud tid"
	    bgcol = "#ffffff"
	    else
	    
	    sideDivTop = 0
	    sideDivLeft = 10
	    opdAfsl = "Logind historik (Stempelur) - Opdater"
	    bgcol = "#ffffff"
	    end if
	
	%>
	
	
	<div id="sindhold2" style="position:relative; width:600px; left:<%=sideDivLeft%>px; top:<%=sideDivTop%>px; visibility:visible;">
   
	
	<% 
	
	oimg = "ikon_stempelur_48.png"
	oleft = 0
	otop = 0
	owdt = 400
	oskrift = opdAfsl
	
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	
	
	dim loginDT, logudDT, datoDT, stur, kommentar, loginDato, idThis, firsttime
	useDato = year(now) & "/" & month(now) &"/" & day(now)
	
	if func = "redloginhist" then
	
	redim loginDT(50), logudDT(50), datoDT(50), stur(50)
	redim kommentar(50), loginDato(50), idThis(50), firsttime(50)
	

	
	
	if id <> 0 then
	
	
	    
	    strSQL2 = "SELECT id, dato, mid FROM login_historik WHERE id = "& id 
	    oRec2.open strSQL2, oConn, 3
        if not oRec2.EOF then
        
        useDato = oRec2("dato")
        useMid = oRec2("mid")
        
        end if
        oRec2.close
	
	
	else
	
	
	
	    strSQL2 = "SELECT id, dato, mid FROM login_historik WHERE mid = "& medarbSel &" AND dato = '"& useDato &"' ORDER BY id DESC limit 0, 1"
	    oRec2.open strSQL2, oConn, 3
        if not oRec2.EOF then
        
        useDato = oRec2("dato")
        useMid = oRec2("mid")
        
        end if
        oRec2.close
	
        
    end if
    
    useDato = year(useDato) & "-" & month(useDato) &"-" & day(useDato)
	
	
	strSQL = "SELECT l.login, l.logud, l.dato, l.stempelurindstilling, l.kommentar, l.id FROM login_historik l "
	strSQL = strSQL & " WHERE l.dato = '"& useDato &"' AND l.mid = "& useMid &" AND l.stempelurindstilling <> -1 ORDER BY id DESC"
	
	'Response.Write strSQL
	'Response.Flush
	
	a = 0
	oRec.open strSQL, oConn, 3 
	while not oRec.EOF 
	
	loginDT(a) = oRec("login")
	
	if len(trim(oRec("logud"))) <> 0 then
	'Response.Write "<br>"& a & " her:" & oRec("logud") & "<br>" 
	logudDT(a) = oRec("logud")
	firsttime(a) = 0
	else
	logudDT(a) = now '"00:00:00"
	'Response.Write "<br>"& a & " der:" & oRec("logud") & "<br>"
	firsttime(a) = 1 
	end if
	
	datoDT(a) = oRec("dato")
	stur(a) = oRec("stempelurindstilling")
	kommentar(a) = oRec("kommentar")
	idThis(a) = oRec("id")
	
	a = a + 1
	oRec.movenext
	wend
	oRec.close 
	
	
	
	else
	
	a = 1
	
	redim loginDT(a), logudDT(a), datoDT(a), stur(a), kommentar(a), loginDato(a), idThis(a)
	redim firsttime(a)
	
	i = 0
	
	datoDT(i) = day(now) & "-" & month(now) & "-"& year(now)
	useMid = medarbSel
	idThis(i) = 0
	loginDT(i) = now
	logudDT(i) = now
	stur(i) = 1
	kommentar(i) = ""
	loginDato(i) = useDato
	firsttime(i) = 2 
	
	
	
	end if
	
	
	
	tTop = 10
    tLeft = 0
    tWdth = 560


    call tableDiv(tTop,tLeft,tWdth)
	
	
	%>
	
	
	<table cellspacing=0 cellpadding=0 border=0 width=100%>
	<tr>
	    <td style="padding:20px 0px 0px 30px;">
	Login historik for: 
	<b><%= weekdayname(weekday(useDato, 1))%> d. <%=formatdatetime(useDato, 1)%></b>
        &nbsp;
   <form action="stempelur.asp?menu=stat&func=dbloginhist&FM_usedatokri=1&medarbSel=<%=useMid%>&showonlyone=<%=showonlyone%>&hidemenu=<%=hidemenu%>&rdir=<%=rdir %>" method="post">
   
	
	<%
	
	for a = 0 TO a - 1
	
	if len(trim(loginDT(a))) <> 0 then
	loginDT(a) = loginDT(a)
	'loginDato(a) = formatdatetime(loginDT(a), 2)
	else
	loginDT(a) = "01-01-2001"
	'loginDato(a) = "01-01-2001"
	end if
	
	'Response.write DatoDT(a)
	'Response.flush
	
	
	
	if cint(idThis(a)) = cint(id) then
	bgthis = "#ffffe1"
	bdcol = "orange" 
	else
	bgthis = "#eff3ff"
	bdcol = "silver" 
	end if%>
	
	    <%select case firsttime(a) 
	    case 1 
	    strTxt = "<img src=../ill/stempelur_ind.gif /> <b>Dette logind er aktivt!<br />"_
	    &"Du vil få den logud tid der er angivet nedenfor.</b>"
	    bdcol = "orange"
	    bgthis = bgthis
	    case 0
        strTxt = "<img src=../ill/stempelur_ok.gif /><b> Dette logind er afsluttet. Logind og logud tid er gemt i databasen.</b>"
	    bgthis = bgthis
	    bdcol = bdcol
	    case 2
	    strTxt = "<b>Opret nyt logind</b>"
	    bgthis = bgthis
	    bdcol = bdcol
	   end select %>
	
	<br />
	 <table cellspacing="0" cellpadding="0" border="0" width=500>
	 <tr bgcolor="<%=bgthis %>">
	    <td colspan=4 style="padding:10px 0px 10px 10px; border-top:1px <%=bdcol%> solid; border-left:1px <%=bdcol%> solid; border-right:1px <%=bdcol%> solid;">
	    <%= strTxt %>
	   </td></tr>
	
	<tr bgcolor="<%=bgthis %>">
		<td valign=top style="padding:0px 0px 0px 10px; border-left:1px <%=bdcol%> solid;"><b>Dato:</b>
		<input type="hidden" name="id" id="id" value="<%=idThis(a)%>">
		<input type="hidden" name="mid" id="mid" value="<%=useMid%>">
		<br />
		<%
		't = 1000
		'if t = 1000 then
		if level = 1 OR id = 0 then%>
		<input type="text" size="8" name="logindato" id="logindato" value="<%=datoDT(a)%>" style="font-size:9px; font-family:arial;"><font class=megetlillesort><br />Eks: 03-04-2006</font>
		<%else%>
		
		<%=formatdatetime(datoDT(a),1) %> <!--left(formatdatetime(loginDT, 3), 5)--> 
		<input type="hidden" name="logindato" id="logindato" value="<%=datoDT(a)%>">
		<%end if%></td>
	
		<td valign=top style="padding:0px 0px 0px 10px;"><b>Logind:</b><br />
		
		<%if level <= 2 OR level = 6 then%>
		<input type="text" name="FM_login_hh" id="FM_login_hh" size=1 value="<%=left(formatdatetime(loginDT(a), 3), 2)%>" style="font-size:9px; font-family:arial;"> 
		<b>:</b> <input type="text" name="FM_login_mm" id="FM_login_mm" size=1 value="<%=mid(formatdatetime(loginDT(a), 3), 4, 2)%>" style="font-size:9px; font-family:arial;">
		&nbsp;&nbsp;tt:mm
		<%else%>
		
		<%=left(formatdatetime(loginDT(a), 3), 5)%>
		<input type="hidden" name="FM_login_hh" id="FM_login_hh" value="<%=left(formatdatetime(loginDT(a), 3), 2)%>">
		<input type="hidden" name="FM_login_mm" id="FM_login_mm"  value="<%=mid(formatdatetime(loginDT(a), 3), 4, 2)%>">
	
		<%end if%></td>
		
		<input type="hidden" name="oprFM_login_hh" id="Hidden1" value="<%=left(formatdatetime(loginDT(a), 3), 2)%>">
		<input type="hidden" name="oprFM_login_mm" id="Hidden2"  value="<%=mid(formatdatetime(loginDT(a), 3), 4, 2)%>">
	    
	    <!-- bruges til arrray split -->
	    <input type="hidden" name="FM_login_hh" id="Hidden3" value="#">
		<input type="hidden" name="FM_login_mm" id="Hidden4"  value="#">
	    
	
	
		<td valign=top style="padding:0px 0px 0px 10px;"><b>Logud:</b><br />
		
		
		<input type="text" name="FM_logud_hh" id="FM_logud_hh" size=1 value="<%=left(formatdatetime(logudDT(a), 3), 2)%>" style="font-size:9px; font-family:arial;"> <b>:</b>
		<input type="text" name="FM_logud_mm" id="FM_logud_mm" size=1 value="<%=mid(formatdatetime(logudDT(a), 3), 4, 2)%>" style="font-size:9px; font-family:arial;">&nbsp;&nbsp;tt:mm<% %>
		<% 
		'** Hvis der ikke før er angivet en logud tid ***'
		if len(trim(logudDT(a))) = 0 then
		logudDT(a) = formatdatetime(now, 3)
		else
		logudDT(a) = logudDT(a) 
		end if
		%>
		
		<input type="hidden" name="oprFM_logud_hh" id="oprFM_logud_hh" value="<%=left(formatdatetime(logudDT(a), 3), 2)%>">
		<input type="hidden" name="oprFM_logud_mm" id="oprFM_logud_mm" value="<%=mid(formatdatetime(logudDT(a), 3), 4, 2)%>">
		
		
		<!-- bruges til arrray split -->
		 <input type="hidden" name="FM_logud_hh" id="Hidden5" value="#">
		<input type="hidden" name="FM_logud_mm" id="Hidden6"  value="#">
		
		</td>
		
		
		
	
		<td valign=top style="padding:0px 0px 0px 10px; border-right:1px <%=bdcol%> solid;"><b>Stempelurindst:</b><br />
		<select name="FM_stur" style="width:100px; font-size:9px; font-family:arial;">
		<option value="0">Ingen</option>
		<%
		strSQL3 = "SELECT id, navn FROM stempelur ORDER BY navn"
		oRec3.open strSQL3, oConn, 3 
		while not oRec3.EOF 
		
		if stur(a) = oRec3("id") then
		sel = "SELECTED"
		else
		sel = ""
		end if
		%>
		<option value="<%=oRec3("id")%>" <%=sel%>><%=oRec3("navn")%></option>
		
		<%
		oRec3.movenext
		wend
		oRec3.close 
		%></select>
		</td>
	</tr>
	<tr bgcolor="<%=bgthis %>">
	    <td colspan=4 style="padding-bottom:30px; padding-left:10px; border-bottom:1px <%=bdcol%> solid; border-left:1px <%=bdcol%> solid; border-right:1px <%=bdcol%> solid;"><br /><b>Kommentar:</b>
	    <br /><textarea id="FM_kommentar_<%=a %>" name="FM_kommentar_<%=a %>" style="font-size:9px; font-family:arial; width:440px; height:60px;"><%=kommentar(a) %></textarea></td>
	</tr>
	</table>
	<%next %>
	
	<br />
	
	
	    
	   
        
        <%
            dagidag = weekday(now, 2)
            p1on = 0
            p2on = 0
            
            select case dagidag
            case 1
            feltNavne = "p1_man AS p1on, p2_man AS p2on"
            case 2
            feltNavne = "p1_tir AS p1on, p2_tir AS p2on"
            case 3
            feltNavne = "p1_ons AS p1on, p2_ons AS p2on"
            case 4
            feltNavne = "p1_tor AS p1on, p2_tor AS p2on"
            case 5
            feltNavne = "p1_fre AS p1on, p2_fre AS p2on"
            case 6
            feltNavne = "p1_lor AS p1on, p2_lor AS p2on"
            case 7
            feltNavne = "p1_son AS p1on, p2_son AS p2on"
            end select
        
            '*** Henter forvalgte standard pauser ***
            strSQL = "SELECT stpause, stpause2, "& feltNavne &""_
            &" FROM  "_
            &" licens WHERE id = 1" 
            
            oRec.open strSQL, oConn, 3
            if not oRec.EOF then
            
            stPause1 = oRec("stpause")
            stPause2 = oRec("stpause2")
            p1on = oRec("p1on")
            p2on = oRec("p2on")
            
            end if
            oRec.close
        
        if p1on <> 0 then %>
        <table cellspacing=0 cellpadding=0 border=0 width=500>
	    <tr bgcolor="#FFFFFF">
	    <td style="padding-left:10px;"><br />
	    <b>Tilføj pauser på denne dato:</b>
	    <br />Pauser bliver automatisk fratrukket login timer på denne dato.
        <br />Standard pauser er: <b><%=stPause1 %></b> og <b><%=stPause2 %></b> min.
        <%
            
            
            '*** Henter allerede oprettee pauser på denne dato
            strSQL = "SELECT minutter, kommentar FROM  "_
            &" login_historik WHERE stempelurindstilling = -1 AND "_
            &" dato = '"& useDato &"' AND mid = " & useMid &" ORDER BY minutter" 
            
            'Response.Write strSQL
            'Response.Flush
            
            
            p = 1
            oRec.open strSQL, oConn, 3
            while not oRec.EOF
            %>
        
            <br /><br /><b>Pause <%=p %>:</b> 
            <select name="p<%=p %>" id="p<%=p %>" style="font-family:arial; font-size:9px;">
            
             <% 
             
             if id = 0 then
             
             if p = 1 then
             caseVal = stPause1
             else
             caseVal = stPause2
             end if
             
             else
             
             caseVal = oRec("minutter")
             
             end if
             
             Select case caseVal
             case 0
             selM0 = "SELECTED"
             selM10 = ""
             selM15 = ""
             selM20 = ""
             selM30 = ""
             case 10
             selM0 = ""
             selM10 = "SELECTED"
             selM15 = ""
             selM20 = ""
             selM30 = ""
             case 15
             selM0 = ""
             selM10 = ""
             selM15 = "SELECTED"
             selM20 = ""
             selM30 = ""
             case 20
             selM0 = ""
             selM10 = ""
             selM15 = ""
             selM20 = "SELECTED"
             selM30 = ""
             case 30
             selM0 = ""
             selM10 = ""
             selM15 = ""
             selM20 = ""
             selM30 = "SELECTED"
             end select
             
               
             
             %>
             <option value="0" <%=selM0%>>0 min</option>
             <option value="10" <%=selM10%>>10 min</option>
             <option value="15" <%=selM15%>>15 min</option>
              <option value="20" <%=selM20%>>20 min</option>
              <option value="30" <%=selM30%>>30 min</option>
              
            </select> min.<br />
        
            
            Kommentar pause <%=p %>:
	        <br /><textarea id="FM_komm_p<%=p %>" name="FM_komm_p<%=p %>" style="font-size:9px; font-family:arial; width:350px; height:46px;"><%=oRec("kommentar") %></textarea>
	       
            <% 
            p = p + 1
            oRec.movenext
            wend
            oRec.close 
            
            
            
            '*** Der er ikke oprettet pauser på denne dato endnu ***
            
            
            if p = 1 then
            call nyePauser(1, p1on)
            call nyePauser(2, p2on)
            end if
            
            
            if p = 2 then
            call nyePauser(2, p2on)
            end if
       
        %>
            <input id="p1on" name="p1on" value="<%=p1on %>" type="hidden" />
             <input id="p2on" name="p2on" value="<%=p2on %>" type="hidden" />
	    </td></tr>
	    
	    <%end if %>
	<tr>
		<td align=right style="padding-right:30px;"><br><br />&nbsp;
		<%if id <> 0 then %>
            <input id="Submit1" type="submit" value="Opdater >>" />
		<%else %>
            <input id="Submit1" type="submit" value="Gem >>" />
        <%end if %>		
		</td>
	</tr>
	</table>
    </form>
    
    
    </td></tr></table>
	</div><!-- table div -->
	</div><!-- sidediv -->
	
	
     <br /><br /> <br /><br /> <br /><br /> <br /><br />
    &nbsp;
    
	
	<%case "stat"
	
	
	
	
	if hidemenu = 0 then%>
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<!--<h4>Timeregistrering - Jobliste</h4>-->
	<%call tsamainmenu(7)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	if showonejob <> 1 then
		call stattopmenu()
	end if
	%>
	</div>
	<%
	sideDivTop = 132
	sideDivLeft = 20
	
	else
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	
	sideDivTop = 20
	sideDivLeft = 20
	
	end if
	
	if len(request("FM_sumprmed")) <> 0 then
	showTot = 1
	sumChk = "CHECKED"
	else
	sumChk = ""
	showTot = 0
	end if
	%>
	
	<div id="sindhold" style="position:absolute; left:<%=sideDivLeft%>; top:<%=sideDivTop%>; visibility:visible;">
	
	<% 
	oimg = "ikon_stempelur_48.png"
	oleft = 0
	otop = 0
	owdt = 600
	oskrift = "Logind-historik (Stempelur)"
	
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	%>
	
	<%call filterheader(0,0,600,pTxt)%>
	<table cellspacing="0" cellpadding="0" border="0" width=100%>
	<form action="stempelur.asp?menu=stat&func=stat&FM_usedatokri=1&lastid=<%=lastid%>&showonlyone=<%=showonlyone%>&hidemenu=<%=hidemenu%>" method="post">
	<tr>
			<td valign=top><b>Medarbejder:</b><br>
			<%
			strSQL = "SELECT mnavn, mid, mnr FROM medarbejdere m WHERE m.mid <> 0 AND mansat <> 2 AND mansat <> 3 "& medarbSQLkri &" ORDER BY mnavn"
			'Response.write strSQL
			'Response.flush
			oRec.open strSQL, oConn, 3 
			%><select name="medarbSel" id="medarbSel" style="width:200px;">
			
			<%if showalle = 1 then%>
			<option value="0">Alle</option>
			<%end if
			
			
			while not oRec.EOF 
			
			 if cint(medarbSel) = oRec("mid") then
			 mSEL = "SELECTED"
			 else
			 mSEL = ""
			 end if %>
			 
			<option value="<%=oRec("mid")%>" <%=mSEL%>><%=oRec("mnavn")%>&nbsp;(<%=oRec("mnr")%>)</option>
			<%
			oRec.movenext
			wend
			oRec.close %>
			</select><br><br>
			<input type="checkbox" name="FM_sumprmed" id="FM_sumprmed" value="1" <%=sumChk%>> Vis sum fordelt pr. medarbejder og logind type.&nbsp;&nbsp;</td>
			<td valign=top><b>Periode:</b><br>
			<!--#include file="inc/weekselector_s.asp"-->
			</td>
			<td>
                <input id="Submit2" type="submit" value="Kør >> " /></td>
	</tr></form>
	</table>
	
	
	<!-- filter header -->
	</td></tr></table>
	</div>
	
	<%
	
	Response.flush
	
	
	'*** Alle timer, uanset fomr. ***
	sqlDatoStart = strAar&"/"&strMrd&"/"&strDag
	sqlDatoSlut = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut
	
	
	if datediff("d", strDag&"/"&strMrd&"/"&strAar, strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 2, 2) > 365 then
	%>
	<br /><br /><br />
	<div style="position:relative; background-color:#ffffff; border:1px red dashed; padding:20px;">
	Datointerval er mere 365 dage. Vælg et midre interval.
	</div>
	<%
	
	Response.end
	
	end if
	
	call ersmileyaktiv()
	
	call stempelurlist(medarbSel, showtot, layout, sqlDatoStart, sqlDatoSlut)
	
	%>
	
	
	</div>
	<br /><br /><br /><br /><br /><br /><br />&nbsp;
	<%case else%>
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<script language="javascript">
	<!--
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	//-->
	</script>
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20; top:10; visibility:visible;">
	<%call crmaktstatheader%>
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr>
    <td valign="top"><font class="pageheader">Kontrolpanel | Stempelur indstillinger</font>
	<br>
	<br>
	Sortér efter <a href="stempelur.asp?menu=tok&sort=navn">Navn</a> eller <a href="stempelur.asp?menu=tok&sort=nr">Id nr.</a>
	<img src="../ill/blank.gif" width="200" height="1" alt="" border="0">
	<a href="stempelur.asp?menu=tok&func=opret">Opret ny Stempelur indstilling <img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a><br>
	<br>&nbsp;
	</td>
	</tr>
	</table>
	
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr bgcolor="#5582D2">
		<td width="8" valign=top rowspan=2><img src="../ill/tabel_top_left.gif" width="8" height="32" alt="" border="0"></td>
		<td colspan=5 valign="top" style="border-top:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right valign=top rowspan=2><img src="../ill/tabel_top_right.gif" width="8" height="32" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td width="50" class=alt><b>Id</b></td>
		<td width="300"class=alt><b>Stempelur indstillinger</b></td>
		<td><b>Faktor</b></td>
		<td><b>Minimums indtast.</b></td>
		
		<td class=alt>&nbsp;</td>
	</tr>
	<%
	sort = Request("sort")
	if sort = "navn" then
	strSQL = "SELECT id, navn, faktor, forvalgt, minimum FROM stempelur ORDER BY navn"
	else
	strSQL = "SELECT id, navn, faktor, forvalgt, minimum FROM stempelur ORDER BY id"
	end if
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	%>
	<tr>
		<td bgcolor="#003399" colspan="7"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="#ffffff" onmouseover="mOvr('gift',this,'#B4C7EF');" onmouseout="mOut(this,'#ffffff');">
		<td style="border-left:1px #003399 solid;">&nbsp;</td>
		<td><%=oRec("id")%></td>
		<td><a href="stempelur.asp?menu=tok&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%> </a> 
		<%if oRec("forvalgt") = 1 then%>
		&nbsp;<b>(Forvalgt)</b>
		<%end if%>
		</td>
		<td><%=formatnumber(oRec("faktor"), 2)%></td>
		<td><%=formatnumber(oRec("minimum"), 0)%> min.</td>
		<td><a href="stempelur.asp?menu=tok&func=slet&id=<%=oRec("id")%>"><img src="../ill/slet.gif" width="20" height="20" alt="" border="0"></a></td>
		<td style="border-right:1px #003399 solid;">&nbsp;</td>
	</tr>
	<%
	x = 0
	oRec.movenext
	wend
	%>	
	<tr bgcolor="#5582D2">
		<td width="8" valign=top height=20 style="border-bottom:1px #003399 solid; border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td colspan=5 valign="top" style="border-bottom:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right valign=top style="border-bottom:1px #003399 solid; border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	</table>
	
	<br><br><br>
	<a href="Javascript:window.close()" class=red>[Luk vindue]</a><br><br>&nbsp;
	</div>
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
