
<%
'key = request("key")
'lto = request("lto")
'Response.redirect "../ver2_10/login.asp?key="&key&"&lto="&lto


'**** Sletter alle cookies ***'
'Dim x,y

'For Each x in Request.Cookies
'    Response.Write("<p>")
'    If Request.Cookies(x).HasKeys Then
'        For Each y in Request.Cookies(x)
'            'Response.Write(x & ":" & y & "=" & Request.Cookies(x)(y))
'            'Response.Cookies(x).expires = date - 1
'            'Response.Write("<br />")
'        Next
'    Else
'        'Response.Write(x & "=" & Request.Cookies(x) & "<br />")
'        'Response.Cookies(x).expires = date - 1
'    End If
'    Response.Write "</p>"
'Next
%>

<noscript>
  <META HTTP-EQUIV="Refresh" CONTENT="0;URL=nojava.asp">
</noscript>







<%
thisfile = "login.asp" %>
<!--#include file="inc/connection/conn_db_inc.asp"-->
<!--#include file="inc/errors/error_inc.asp"-->
<!--#include file="inc/regular/global_func.asp"-->
<%

'*** Sætter lokal dato/kr format. *****
Session.LCID = 1030




session("spmettanigol") = session("spmettanigol") + request("attempt")

	If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
	
	'************************************ Login side *********************************************
	%>
	<!--#include file="inc/regular/header_login_inc.asp"-->



	
	
	<!--
	
	
	<div style='position:absolute; top:50px; left:400px; margin:0px; z-index : 1250;'>
	<OBJECT classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" WIDTH="750" HEIGHT="53" id="godjul" ALIGN="">
 	<PARAM NAME=movie VALUE="godjul.swf"> 
	<PARAM NAME=quality VALUE=high>
	<PARAM NAME=wmode VALUE=transparent> 
	<PARAM NAME=bgcolor VALUE=#000000>
	<EMBED src="godjul.swf" quality=high wmode=transparent bgcolor=#000000 WIDTH="750" HEIGHT="53" NAME="godjul" ALIGN="" TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.macromedia.com/go/getflashplayer"></EMBED>
	</OBJECT>
	</div>
	-->
	
	
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
	tboxsize = 100
	'topgif = ""
	else
	pixLeft = 120
	pixTop = 120
	tboxsize = 200
	'topgif = "<img src='ill/login_timeout.gif' alt='' border='0'>"
	end if%>
	
	

	
	
	<center>
	<div id="content" style="position:relative; width:850px; padding:0px; left:0px; top:40px; background-color:#FFFFFF; border:1px #5582d2 solid;">
	
		<table cellspacing="0" cellpadding="0" border="0" width=100%>
	<form action="login.asp" method="POST">
   
	<tr bgcolor="#004E90">
		<td align=left style="padding:20px;"><img src="ill/outzource_logo_neg_300.gif" /><!--<img src="ill/jul_outz_2012.png" />-->
		</td>
		
		<td>
		<!--<img SRC="ill/Segl_100X72.gif" BORDER="0" ALT="Denne side er SSL 128 bit krypteret." width="97" height="67">-->
		
		
		<!-- GeoTrust QuickSSL [tm] Smart Icon tag. Do not edit. -->
        <!--
        <SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="//smarticon.geotrust.com/si.js"></SCRIPT>
        -->
        <!-- end GeoTrust Smart Icon tag -->
		

		
		</td>
	</tr>
	<tr>
	    <td bgcolor="#ffffff" style="height:5px;"><img src="ill/blank.gif" alt="" border="0"></td>
	</tr>
	</table>
		   
	<table cellspacing="0" cellpadding="0" border="0" width=100%>
	<tr>
		<td valign="top" style="padding:20px 10px 10px 40px; width:350px;">
		
		
		<div style="position:relative; left:0px; top:0px; width:340px; overflow:hidden; border:10px yellowgreen solid; padding:20px; background-color:#FFFFFF;">
		
        <%select case lto 
        case "demo", "intranet - local" 
       ' Response.Write "<b>Velkommen til vores online TimeOut demo.</b>"_
		'	&" <br>"_
	'		&" Login og password til demoen er: <b>guest / 1234 </b> .<br><br>"_
'			&" Husk! I er altid velkommen til at kontakte OutZourCE på "_
	'		&" +45 2684 2000 eller på email: <a href='mailto:timeout@outzource.dk' class=vmenu>timeout@outzource.dk</a>"
			'&" <br><br>God fornøjelse...<br>&nbsp;"
        case "kejd_pb2", "wwf2", "epi_cati"
        Response.write "<h4 style=""color:red;"">DEMO - DEMO - DEMO - DEMO</h4>"
        case else
        end select'OR lto = "intranet - local"  %>
        
        <table border=0 cellspacing=0 cellpadding=0 width=100%>
	    <tr><td colspan=2><h4>Logind:</h4></td></tr>
	
			

            	<%if lto = "demo" OR lto = "intranet - local" then %>

                <tr><td colspan=2 style="padding:2px 0px 20px 0px; line-height:14px;">Kære besøgende, angiv venligst dit navn, firma og email, og du vil få adgang til vores demo:</td></tr>

                     <tr><td align="right" valign=top style="padding:4px 0px 0px 0px;"><b>Dit navn:</b></td>
                	<td valign="top" style="padding:0px 0px 2px 0px;">&nbsp;<input type="text" name="FM_ref" id="FM_ref" value="" style="width:<%=tboxsize%>;">
                </td>
        </tr>

			   <tr><td align=right valign=top style="padding:4px 0px 0px 0px;"><b>Firma:</b></td>
                	<td valign="top" style="padding:0px 0px 2px 0px;">&nbsp;<input type="text" name="FM_firma" id="FM_firma" value="" style="width:<%=tboxsize%>;">
                </td>
        </tr>

          <tr><td align="right" valign=top style="padding:4px 0px 0px 0px;"><b>Email:</b></td>
                	<td valign="top" style="padding:0px 0px 2px 0px;">&nbsp;<input type="text" name="FM_email" id="FM_email" value="" style="width:<%=tboxsize%>;">
                </td>
        </tr>
  <tr><td align="right" valign=top style="padding:4px 0px 0px 0px;"><b>Telefon:</b></td>
                	<td valign="top" style="padding:0px 0px 2px 0px;">&nbsp;<input type="text" name="FM_tel" id="Text1" value="" style="width:<%=tboxsize%>;">
                </td>
        </tr>


     

            <%else %>

            	<!--<tr>
		    <td colspan=2 style="padding:10px;" class=lille>Logind <%=weekdayname(weekday(now, 1)) %> d. <%=formatdatetime(now, 1) %> kl. <%=formatdatetime(now, 3) %> </td>
		</tr>-->
			
			<%end if %>
		
        <input type="hidden" id="lto" value="<%=lto %>" />
       

        <tr>
			<td align=right valign=top style="padding:16px 0px 0px 0px;"><b>Brugernavn:</b>
			
			
			<%if lto = "demo" then
			uval = "Guest"
			else
			    if Request.Cookies("login")("usrval") <> "" then
			    uval = Request.Cookies("login")("usrval")
                'shownyhb = 0
			    else
			    uval = ""
			    'shownyhb = 1
                end if
			end if%></td>
			<td valign="top" style="padding:12px 0px 2px 0px;">
                &nbsp;<input type="Text" name="login" id="login" value="<%=uval%>" style="width:<%=tboxsize%>;"></td>
		</tr>
		<tr>
			<td align=right valign=top style="padding-top:4px;"><b>Password:</b>
			<%if lto = "demo" then
			pawval = "1234"
			else
			    if Request.Cookies("login")("pwval") <> "" then
			    pawval = Request.Cookies("login")("pwval")
			    else
			    pawval = ""
			    end if
			end if%></td>
			<td valign="top" style="padding:0px 0px 2px 0px;">
                &nbsp;<input type="Password" name="pw" id="pw" value="<%=pawval %>" style="width:<%=tboxsize%>;">
			&nbsp;<input type="hidden" name="attempt" value="1">
			
			
			
			    <%if len(pawval) <> 0 then
			    huskCHK = "CHECKED"
			    else
			    huskCHK = ""
			    end if %>
                <br /><input id="huskmig" name="huskmig" type="checkbox" value="1" <%=huskCHK %> /> Husk mig


               
			    
			</td>
			</tr>
			
            <!-- Nyhedsbrev 
            <tr>
            <td valign=top colspan=2 style="padding:4px 0px 2px 20px;"> <%
                
                monthnow = month(now)
                
                select case monthnow
                case 2,4,8,12
                shownyhb = 1
                case else
                shownyhb = 0
                end select

                if request.cookies("tsa")("nyhansw") <> "" then
                shownyhb = 0
                end if
                
                if cint(shownyhb) = 1 then %>
                <br /><br />
               
                <b>Ønsker du at modtage TimeOut nyhedsbrev:</b><br />
                <input id="Radio1" name="nyhedsbrev" type="checkbox" value="1"/> Ja tak, det vil jeg gerne.
                (afmelding sker fra din medarbejderprofil)
                <%end if%></td>
            </tr>
            -->
	
			
			
			
		<tr>
	
		<td valign=top colspan=2 style="padding-top:30px;">	
		
			
			<%
			'**************** Stempelur ********************************
			session("stempelur") = 0
			strSQL = "SELECT stempelur FROM licens WHERE id = 1"
			oRec.open strSQL, oConn, 3 
			if not oRec.EOF then
			
				session("stempelur") = oRec("stempelur")
			
			end if
			oRec.close  
			
			if session("stempelur") <> 0 then
			
			%>
		    <h4>Stempelur indstilling:</h4>
			<%
			
				
				strSQL = "SELECT id, navn, faktor, forvalgt FROM stempelur ORDER BY navn"
				oRec.open strSQL, oConn, 3 
				st = 0
                while not oRec.EOF 
				
                

				if oRec("forvalgt") = 1 then
				chk = "CHECKED"
				else
				chk = ""
				end if
				%>
				<input type="radio" name="FM_stempelur" value="<%=oRec("id")%>" <%=chk%>> <%=oRec("navn")%><br />
				<%

                st = st + 1
				oRec.movenext
				wend
				oRec.close 

                if st = 0 then
                %>
                <br />(ingen)
			    <%
                end if
			
			else
			%>
			<input type="hidden" name="FM_stempelur" id="FM_stempelur" value="0">
			
			<%
			end if
			'************************************************************
			
            
              select case lto 
              case "demo", "intranet - local" 
              submitDsp = "none"
              submitVzb = "hidden"
              case else
              submitDsp = ""
              submitVzb = "visible"
              end select
            %>
		
		
		
		
		</td>
		</tr>
        <tr>
		<td align=right colspan=2 style="padding:30px 20px 2px 0px;">
        <div id="logindiv" style="visibility:<%=submitVzb%>; display:<%=submitDsp%>;">
		<input id="loginsubmit" type="submit" value="Login >>" />
        
        </div>	
		<br />&nbsp;
		
		</td>
		</tr>

        <tr><td colspan=2><br /><br />
        <b>Genveje til TimeOut:</b> <span style="font-size:10px; color:#999999;">(IE 4+)</span><br />
        <a href="javascript:window.external.AddFavorite('https://outzource.dk/<%=lto%>','TimeOut - Tid, Overblik & Fakturering')" style="color:#999999; font-size:10px; text-decoration:underline;">Tilføj TimeOut til favoritter</a>&nbsp;<span style="font-size:8px;">[CTRL] + [D]</span> 
        
        
 <br />
 <a  href="#" onclick="this.style.behavior='url(#default#homepage)';this.sethomepage('https://outzource.dk/<%=lto%>');" style="color:#999999; font-size:10px; text-decoration:underline;">Gør TimeOut til startside</a><br />

  <a  href="https://outzource.dk/<%=lto%>" style="color:#999999; font-size:10px; text-decoration:underline;">TimeOut genvej</a> &nbsp;<span style="font-size:10px;"> << Træk dette link ud på skrivebordet</span>
  <br /><br />&nbsp;
        
        </td></tr>
		
		
			<% '*** Logo og licensindehaver ***'
			strSQL = "SELECT useasfak, logo, id, filnavn FROM kunder, filer WHERE useasfak = 1 AND filer.id = kunder.logo"
			
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
			logonavn = "<img src='inc/upload/"&lto&"/"&oRec("filnavn")&"' alt='' style='border:0px #000000 solid;'>"
			else
			logonavn = ""
			end if
			oRec.close
			
		%>
		
			 
		    <%if logonavn <> "" then %>
		    <tr>
			<td colspan=2 bgcolor="#ffffff" style="padding:10px;" align=center>
			<%=logonavn %>
		    </td>
		    </tr>
		    
		    <%end if %>
		    <tr>
			<td colspan=2 style="padding:4px 0px 2px 15px;">
			<br />Licenshaver: <b><%=licensto%></b>
			</td>
			</tr>
		
		</table>
		</div>
		
		
		<br />&nbsp;
		
			
	   </td>
		</form>
		
		<td style="width:20px">&nbsp;</td>	
			
		
		
		
		
		
		<td valign="top" style="padding:20px 40xp 10px 0px;">
		    
		    <table cellpadding=0 cellspacing=0 border=0 width=100%><tr><td><br />
		            <h4>Seneste:</h4>
		            

                    <b class=news>12 april 2013</b> 
                
                      <p><b>Vi er igang med..</b><br />
                      <br />
                          
                          <span style="color:#999999;">- Importere timer via excel. Afsluttet!</span><br /><br /> 
                          <span style="color:#999999;">- Afstemning lønperiode og overføre flekssaldo. Afsluttet!</span><br /><br />
                          - Forbedre laodtiden på Grandtotal, samt mulighed for at rotere udprintet, så det bliver nemmere at få det hele med ud på papir. <br />Forv klar: uge 15<br /><br />
                          - Ny kørsels registrering, integreret med Google maps, så TimeOut selv regner distancen ud. <br />Forv klar: uge 17-18<br /><br />
                          - Mulighed for at få igangværende job vist i en forenklet version, med mulighed for selv at vælge kolonner, så det er nemmere at få overblik over projekterne. <br />Forv klar: uge 18-19<br /><br />
                          - Ny navigation, så det bliver nemmere at springe fra et hjørne til et andet i TimeOut. <br /><br />
                          - Diverse kunde specifikke tilretninger<br /> <br />
                          
                           

                      
                     <b>Vi planlægger..</b>
                    <br /><br />
                          - Upload af udlæg og materialer via .csv fil.<br /><br />
                          - Søgning blandt medarbejdere på timereg. siden uden re-load for markant bedre performance for de virksomheder der har +100 brugere<br /><br />

                      Mvh.<br />
                      OutZourCE
                       
                      <br /><br />
                      </p>

                    <!--

                       <b class=news>13 marts 2013</b> 
                      <p><b>Ny release</b><br />
                      Vi har netop lanceret vores nye re-lease, der bl.a indeholder:<br /><br />
                      - Afstemning lønperiode og overføre flekssaldo. Dette medfører også langt hurtigere visning af afstemningen (<a href="help_and_faq/afstemning_lon_20130313.pdf" target="_blank">læs mere her</a>)<br /><br />
                      - Mulighed for uplaod af timer via excel.<br />  
                      - Optimering af loadtid på timereg.siden<br />
                      <br />
                      Vores næste relaese kommer bl.a til at indeholde en forbederet navigation, optimering af smartphone versionen.<br /><br /> 

                      Mvh.<br />
                      OutZourCE
                       
                      <br /><br />
                      </p>

                      -->

			
			    <hr style="width:260px; border:1px #cccccc dashed; height:1px;" />
			
			</td></tr></table>
			
		 
        
		
		
		</td>
			<td style="width:20px">&nbsp;</td>	
		
		</tr>
		</table>
		
		</div> <!-- cointent -->
		</center>
		<br /><br />
		
		
        <%
        call browsertype()
        if browstype_client <> "ip" then %>
        
        <script>
		
        
        
        function NewWin(url)    {
		window.open(url, 'Disclaimer', 'width=600,height=580,scrollbars=yes,toolbar=no');    
		}
		</script>
		<br /><br /><br />


		<center>
		<font class=lillehvid>© 2002 - <%=year(now) %> OutZourCE og OutZourCE -underleverandører. Alle rettigheder forbeholdes.&nbsp;&nbsp;|&nbsp;&nbsp; 
		Tænkt og produceret af OutZourCE&nbsp;&nbsp;|&nbsp;&nbsp; Design af Margin Media<br>
		Graphical elements copyright ©Microsoft Corp&nbsp;&nbsp;|&nbsp;&nbsp;<a href="javascript:NewWin('timereg/disclaimer.asp');" target="_self"><font class="lillehvid"><u>Betingelser</u></a></font>&nbsp;&nbsp;|&nbsp;&nbsp;Optimeret til IE 5+, FireFox 3+ og Chrome 
        
		<% 
		select case lto
		case "infow_demo", "infow", "assurator", "viatech", "pcmanden", "dreist", "rage", "enho"
		%>
		&nbsp;&nbsp;|&nbsp;&nbsp;<b>Support og Salg:</b> salg@informationworker.dk &nbsp;&nbsp;|&nbsp;&nbsp;Tlf: 70212128 
		<%case else%>
		&nbsp;&nbsp;|&nbsp;&nbsp;<b>Support og Salg:</b> timeout@outzource.dk&nbsp;&nbsp;|&nbsp;&nbsp;Tlf: +45 26 84 20 00&nbsp;&nbsp;|&nbsp;&nbsp;Skype: OutZourCE
		
		<%end select%>
		
        <br /><span style="font-size:9px; color:#ffffff;"><%=request.servervariables("HTTP_USER_AGENT") %></span>
        <br /><br>&nbsp;
		</center>

        <%end if %>
		
		
		
		<%
		
		
	
	
else
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

    
	if varlogin = "" Then
	%>
	<!--#include file="inc/regular/header_login_inc.asp"-->
	<%
	errortype = 1
	call showError(errortype)
	
	else
	'***************** login og pw modtages og valideres ***********************************************
		
		if varPW = "" then
		%>
		<!--#include file="inc/regular/header_login_inc.asp"-->
		<%
		errortype = 2
		call showError(errortype)
		
		else
		
		
		
		'**** Bruges på timregside til at nulstille bruger ***'
		session("forste") = "j"
		
		'*** Er det en ansat eller en kunde der logger på?***'
		'*** Altid ansat, Kunder har deres  egen login fil **'

        call TimeOutVersion()
	
		
		stransatkunde = "1" 'request("FM_ansat_kunde")
		
		if stransatkunde = "1" then
		'DECODE(
		strSQL = "SELECT m.mansat, m.mnavn, b.rettigheder, m.mid, m.tsacrm, m.pw, m.timereg FROM medarbejdere m, brugergrupper b WHERE "_
		&" login = '"&trim(request("login"))&"' AND (mansat <> 2 AND mansat <> 3) AND m.pw = MD5('"& trim(request("pw")) &"') AND b.id = m.brugergruppe"
		else
		strSQL = "SELECT id, kundeid, email, password AS pw, navn AS Mnavn FROM kontaktpers WHERE email='"& request.Form("login") &"'"
		end if
		
		founduser = 0
		
		oRec.Open strSQL, oConn, 3
		if not oRec.EOF Then
		
		            founduser = 1		
            			
			            '*******************************************************************************************'
			            '*** Sætter sesions var. alt efter om det er en medarbejder eller en kunde der logger på****'
            			
            			
			            if stransatkunde = "1" then
			            strUsrId = Trim(oRec("Mid"))
			            session("user") = Trim(oRec("Mnavn"))
			            session("login") = strUsrId
			            session("mid") = strUsrId
			            session("rettigheder") = oRec("rettigheder")
            				
				            startside = oRec("tsacrm")
            				
            				
				            '*** Hent timrregside fra db eller Form *** 
				            if len(trim(request("FM_0206"))) <> 0 then
				            treg0206 = request("FM_0206")
				            else
				            treg0206 = oRec("timereg")
				            end if
            				
            			
            				
            				
			            else
			                strUsrId = 0
			                thisKpid = oRec("id")
			                session("user") = Trim(oRec("Mnavn"))
			                session("login") = strUsrId
			                session("mid") = oRec("kundeid") 'Kundeid
			                session("rettigheder") = 0
			                startside = 2 'joblog_k.asp
			                treg0206 = 0
			            end if
			            '*******************************************************************************************'
            		
            		
            		
            		
            		
            		
		            '********************* Skriver til logfil ********************************************************'
		            if request.servervariables("PATH_TRANSLATED") <> "E:\www\timeout_xp\wwwroot\ver2_1\login.asp" AND request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\login.asp" then

                    

				            Set objFSO = server.createobject("Scripting.FileSystemObject")
				            select case lto
				            '*** Bruger case else undt. for nedenstående, der dog burde kunne skiftes unden probs. Minus måske "net" pga lto.
            				
				            case "titoonic"
					            Set objF = objFSO.GetFile("d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\log\logfile_timeout_titoonic.txt")
					            Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\log\logfile_timeout_titoonic.txt", 8)
				            case "skybrud"
					            Set objF = objFSO.GetFile("d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\log\logfile_timeout_skybrud.txt")
					            Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\log\logfile_timeout_skybrud.txt", 8)
            				
				            case else
					            Set objF = objFSO.GetFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\logfile_timeout_"&lto&".txt")
					            Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\logfile_timeout_"&lto&".txt", 8)
            				
				            end select
            				
				            objF.writeLine(session("user") &chr(009)&chr(009)&chr(009)& date &chr(009)& time&chr(009)& request.servervariables("REMOTE_ADDR"))
				            objF.close	
		            end if
		            '*******************************************************************************************'
            		
            		
            		
		            '******************* Sætter sidste login dato *********************************************'
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
		            strMth = "Maj. "
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
            		
		  end if
		  oRec.close
            		
		
		
		if founduser = 1 then
		
		
    		
		    if stransatkunde = "1" then
		    strSQL = "SELECT lastlogin FROM medarbejdere WHERE Mnavn ='" & session("user") &"'"
		    else
		    strSQL = "SELECT lastlogin AS lastlogin FROM kontaktpers WHERE id="& thisKpid 
		    end if
    		
		    oRec.open strSQL, oConn, 3
		    if not oRec.EOF then
		    session("strLastlogin") = oRec("lastlogin")
		    end if
		    oRec.close
		    
		    
		     
		    
		    
		    '************************************************************************'
		    '***** Medarbejder kunde login ***'
		    if stransatkunde = "1" then
			
			    oConn.execute("UPDATE medarbejdere SET lastlogin = '"& strLastLogin &"', Mnavn = '"& session("user") &"', timereg = "& treg0206 &" WHERE Mid ="& strUsrId &"")
		        
                '** Tilmeld til nyhedsbrev ***'
		        if request("nyhedsbrev") = "1" then
                oConn.execute("UPDATE medarbejdere SET nyhedsbrev = 1 WHERE Mid ="& strUsrId &"")
                response.cookies("tsa")("nyhansw") = 1
                response.cookies("tsa").expires = date + 180
		        end if
		       
		       '*** Ferie Pl / Ferie afholdt **' 
		        call opdaterFeriePl(session("rettigheder"))
                
               '*** Kosoliderer timer *****'

               'Response.Write "lto:" & lto & "dt: "& datepart("d", now, 2,2)  & " //last: "& datepart("d", session("strLastlogin"), 2,2) &" level: " & session("rettigheder")
               'Response.end

               if session("rettigheder") = 1 AND ((datepart("d", now, 2,2) = 18 AND _
               datepart("d", session("strLastlogin"), 2,2) <> 18) OR (datepart("d", now, 2,2) = 27 AND datepart("d", session("strLastlogin"), 2,2) <> 27)) then
                   select case lto
                   case "xx"
                   case "dencker", "epi", "epi_osl", "epi_sta", "epi_ab"
                   call timer_konsolider(lto,0)
                   case "intranet - local", "outz"
                   call timer_konsolider(lto,0)
                   end select
               end if

                
		        'Response.end
		        
		        
		       
		        
				
				'*** Opdater login_historik (Stempelur) ****'		
				if len(request("FM_stempelur")) <> 0 then
				intStempelur = request("FM_stempelur")
				else
				intStempelur = 0
				end if
				
                session("dontLogind") = 0

                '**** Selvom stempelur er slået til skal følgende grupper ikke have et logind **'
                select case lto 
				case "cst"
                        
                        strSQLmansat = " mansat <> 0"
                        call medarbiprojgrp(21, session("mid"), 0, -1)
                        if instr(instrMedidProgrp, "#"& session("mid") &"#,") <> 0 then
                        session("dontLogind") = 1
                        else
                        session("dontLogind") = 0
                        end if

                
              
                        
                case else
                session("dontLogind") = 0
                end select


                '**** Log ikke ind automatisk ****'
                select case lcase(lto)
                case "xintranet - local", "fk_bpm", "kejd_pb", "kejd_pb2"
                case else
              
				    
                if session("dontLogind") <> 1 then
				'**** Tjekker om der skla laves en auto-logud eller om der er et igangværende logind på dagen ***'    
				call logindStatus(strUsrId, intStempelur, 1, now)
				end if


                end select
				
				
				
				'Response.End
		
		else
		oConn.execute("UPDATE kontaktpers SET lastlogin = '"& strLastLogin &"' WHERE id="& thisKpid  &"")
		end if
		'***********************************'
		'***** Medarbejder kunde login  END '
		
		
        '********************* Sebnder mail hvis det er demo **************************************
        if lto = "demo" then
        email = request("FM_email")
        firma = request("FM_firma")
        ref = request("FM_ref")
        tel = request("FM_tel")

                                    

                                    if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\login.asp" then

                
                                    Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		                            ' Sætter Charsettet til ISO-8859-1
		                            Mailer.CharSet = 2
		                            Mailer.FromName = "TimeOut | " & ref &"("& firma &")" 
		                            Mailer.FromAddress = email
		                            Mailer.RemoteHost = "webmail.abusiness.dk" '"pasmtp.tele.dk"
						            Mailer.AddRecipient "Søren Karlsen", "sk@outzource.dk"
		                          

                                 
            					    Mailer.Subject = "Der har været en ny potentiel kunde i demoen"
		                            strBody = "Hej OutZourCE" & vbCrLf & vbCrLf
                                    strBody = strBody &"Der har været en ny potentiel kunde i demoen : "& vbCrLf & vbCrLf
                                    strBody = strBody &"Firma: "& firma & vbCrLf 
						            strBody = strBody &"Navn: "& ref & vbCrLf & vbCrLf
                                    strBody = strBody &"Email: "& email &""& vbCrLf & vbCrLf
                                    strBody = strBody &"Tel: "& tel &""& vbCrLf & vbCrLf
		                            strBody = strBody &"Med venlig hilsen" & vbCrLf
		                            strBody = strBody & session("user") & vbCrLf & vbCrLf
            		                
            		
            		
		                            Mailer.BodyText = strBody
            		
		                            Mailer.sendmail()
		                            Set Mailer = Nothing


                                    end if

        end if


			
		
		'*********** redirect til den valgte side *************************************************
		'Response.Cookies("usrname")("ansatkunde") = request("FM_ansat_kunde")
		
		
		Response.Cookies("loginlto")("lto") = lto
		Response.Cookies("loginlto").expires = date + 1
       
		
		if len(request("huskmig")) <> 0 then
		Response.Cookies("login")("usrval") = request("login")
		Response.Cookies("login")("pwval") = request("pw")
		Response.Cookies("login").expires = date + 45
		else
		Response.Cookies("login").expires = date - 1
		end if
		
		
		
		
		
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
				'if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\login.asp" then
				'	if stransatkunde = "1" then
				'		select case startside 
				'		case 1 
				'			if lto = "kasters" then
				'			response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/timereg/kunder.asp?menu=crm"
				'			else
				'			response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/timereg/crmkalender.asp?menu=crm&shokselector=1&ketype=e&selpkt=kal"
				'			end if
				'		case 2
				'		response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/timereg/jbpla_w.asp?menu=res"
				'		case else '0
				'			'if lto = "bowe" then
				'			'Response.write "treg0206" & treg0206
				'			if cint(treg0206) = 1 then
				'			response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/timereg/timereg_2006_fs.asp"
				'			'response.redirect "http://81.19.249.35/timeout_xp/wwwroot/ver2_1/timereg/timereg_2006_fs.asp"
				'			else
				'			response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/timereg/timereg.asp?showstempelur=1"
				'			end if
				'		end select
				'	else
				'		response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/timereg/joblog_k.asp"
				'	end if
				'else
					if stransatkunde = "1" then
						select case startside 
						case 1 
						    if lto = "essens" OR lto = "fkfk" then
						    response.redirect "timereg/crmhistorik.asp?menu=crm&ketype=e&func=hist&selpkt=hist"
						    else
						    response.redirect "timereg/crmkalender.asp?menu=crm&shokselector=1&ketype=e&selpkt=kal"
						    end if
						case 2
						response.redirect "timereg/jbpla_w.asp?menu=job"
                        case 3

                        datoNu = now
                        dagIuge = datepart("w", datoNu, 2,2)
                        select case dagIuge
                        case 1
                        add = 6
                        case 2
                        add = 5
                        case 3
                        add = 4
                        case 4
                        add = 3
                        case 5
                        add = 2
                        case 6
                        add = 1
                        case 7
                        add = 0
                        end select

                        sondagIuge = dateAdd("d", add, datoNu)
                        mandagIuge = dateAdd("d", -6, sondagIuge)

                        sondagIuge = year(sondagIuge) &"/"& month(sondagIuge) &"/"& day(sondagIuge) 
                        mandagIuge = year(mandagIuge) &"/"& month(mandagIuge) &"/"& day(mandagIuge) 

                        'Response.Write "timereg/logindhist_2011.asp?usemrn="&session("mid")&"&varTjDatoUS_man="&mandagIuge&"&varTjDatoUS_son="&sondagIuge
                        'Response.end
                        response.redirect "timereg/logindhist_2011.asp?usemrn="&session("mid")&"&varTjDatoUS_man="&mandagIuge&"&varTjDatoUS_son="&sondagIuge    
						case else '0
							if cint(treg0206) = 1 then
                            
                            

                            'Response.write timeout_version
                            'Response.end

                            if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\login.asp" then

                            select case timeout_version
                            case "ver2_1"
							response.redirect "timereg/timereg_akt_2006.asp"
                            case "ver3_99"
                            response.redirect "../ver3_99/timereg/timereg_akt_2006.asp"
                            case "ver2_10"
                            'response.redirect "../ver2_10/timereg/timereg_akt_2006.asp?hideallbut_first=1"
                            response.redirect "../ver2_10/timereg/timereg_akt_2006.asp"
                            case else
                            response.redirect "timereg/timereg_akt_2006.asp"
                            end select
							'response.redirect "timereg/timereg_2006_fs.asp"

                            else

                            'select case lto
                            'case "dencker", "intranet - local"
                            response.redirect "http://localhost/timeout_xp/timereg/timereg_akt_2006.asp"
                            'case else
                            'response.redirect "http://localhost/timeout_xp/timereg/timereg_akt_2006.asp?hideallbut_first=1"
                            'end select

                            end if


							else
							response.redirect "timereg/timereg.asp?showstempelur=1"
							end if
						end select
					else
						response.redirect "timereg/joblog_k.asp"
					end if
				'end if
				
				
			end if
		'********************************************************************************************
		
		
		
		else 'founduser
		%>
		<!--#include file="inc/regular/header_login_inc.asp"-->
		<%
		errortype = 4
		call showError(errortype)
		
		End if
		
		
		
		
		
		'call closeDB
		
		End if '* Validering '
		End if '* Validering '
	end if '* Validering '
	

end if
%>
<!--#include file="inc/regular/footer_login_inc.asp"-->


