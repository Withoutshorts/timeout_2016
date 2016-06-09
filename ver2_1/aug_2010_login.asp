
<%
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
	tboxsize = 10
	'topgif = ""
	else
	pixLeft = 120
	pixTop = 120
	tboxsize = 25
	'topgif = "<img src='ill/login_timeout.gif' alt='' border='0'>"
	end if%>
	
	

	
	
	<center>
	<div id="content" style="position:relative; width:800px; padding:0px; left:0px; top:40px; background-color:#FFFFFF; border:1px #5582d2 solid;">
	
		<table cellspacing="0" cellpadding="0" border="0" width=100%>
	<form action="login.asp" method="POST">
	<tr bgcolor="#004E90">
		<td align=left style="padding:20px;"><img src="ill/outzource_logo_neg_300.gif" />
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
		
		<%if lto = "demo" then 'OR lto = "intranet - local"  %>
		
		<% 
		Response.Write "<b>Velkommen til vores online TimeOut demo.</b>"_
			&" <br>"_
			&" Login og password til demoen er: <b>guest / 1234 </b> .<br><br>"_
			&" Husk! I er altid velkommen til at kontakte OutZourCE på "_
			&" tel.: +45 2684 2000 eller på email: <a href='mailto:timeout@outzource.dk' class=vmenu>timeout@outzource.dk</a>"_
			&" <br><br>God fornøjelse...<br>&nbsp;"
			%>
		
		<%end if %>
		
		
		<div style="position:relative; left:0px; top:0px; border:1px #5582d2 solid; padding:0px; background-color:#D6Dff5;">
		<table border=0 cellspacing=0 cellpadding=0 width=100%>
		<tr>
		    <td colspan=2 style="padding:10px;" class=lille>Logind <%=weekdayname(weekday(now, 1)) %> d. <%=formatdatetime(now, 1) %> kl. <%=formatdatetime(now, 3) %></td>
		   
		</tr>
		<tr>
			
		
		
			
		
			<%if request("usedemofromweb") = "1" then %>
			<input type="hidden" name="FM_firma" value="<%=request("fm_firma")%>">
			<input type="hidden" name="FM_email" value="<%=request("fm_email")%>">
			<input type="hidden" name="FM_ref" value="<%=request("fm_ref")%>">
			<%end if %>
			
			<td align=right valign=top style="padding:16px 0px 2px 0px;"><b>Brugernavn:</b>
			
			
			<%if lto = "demo" then
			uval = "Guest"
			else
			    if Request.Cookies("login")("usrval") <> "" then
			    uval = Request.Cookies("login")("usrval")
			    else
			    uval = ""
			    end if
			end if%></td>
			<td valign="top" style="padding:12px 0px 2px 0px;">
                &nbsp;<input type="Text" name="login" id="login" value="<%=uval%>" size="<%=tboxsize%>" style="border:1px #5582d2 solid;"></td>
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
			<td valign="top">
                &nbsp;<input type="Password" name="pw" id="pw" value="<%=pawval %>" size="<%=tboxsize%>" style="border:1px #5582d2 solid;">
			&nbsp;<input type="hidden" name="attempt" value="1">
			
			
			
			    <%if len(pawval) <> 0 then
			    huskCHK = "CHECKED"
			    else
			    huskCHK = ""
			    end if %>
                <br /><input id="huskmig" name="huskmig" type="checkbox" value="1" <%=huskCHK %> /> Husk mig.
			    
			</td>
			</tr>
			
	
			
			
			
		<tr>
	
		<td valign=top colspan=2 style="padding:4px 0px 2px 20px;">	
		
			
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
		
			
			<br />
			<img src="ill/ikon_stempelur_24.png" alt="" border="0">&nbsp;&nbsp;<b>Stempelur indstilling:</b>
			<%
			
				
				strSQL = "SELECT id, navn, faktor, forvalgt FROM stempelur ORDER BY navn"
				oRec.open strSQL, oConn, 3 
				while not oRec.EOF 
				
				if oRec("forvalgt") = 1 then
				chk = "CHECKED"
				else
				chk = ""
				end if
				%>
				<br><input type="radio" name="FM_stempelur" value="<%=oRec("id")%>" <%=chk%>> <%=oRec("navn")%>
				<%
				oRec.movenext
				wend
				oRec.close 
			
			
			else
			%>
			<input type="hidden" name="FM_stempelur" id="FM_stempelur" value="0">
			
			<%
			end if
			'************************************************************
			%>
		
		
		
		
		</td>
		</tr>
		<tr>
		<td align=right colspan=2 style="padding:4px 20px 2px 0px;">
		<input id="loginsubmit" type="submit" value="Login >>" />	
		<br />&nbsp;
		
		</td>
		</tr>
		
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
		
		<td style="width:100px">&nbsp;</td>	
			
		
		
		
		
		
		<td valign="top" style="padding:20px 40xp 10px 10px;">
		    
		    <div style="position:relative; width:350px; padding:10px; height:280px; background-color:#ffffff; border:0px #5582d2 solid; overflow:auto;">
		    <table cellpadding=0 cellspacing=0 border=0 width=90%><tr><td>
		            
		            <!--
		               <b class=news>19 maj. 2010</b> 
		            <p><b>Problemer med at huske cookies i IE 8.0</b><br />
		            Vi oplever flere brugere der har problemer med at deres browser ikke husker
		            indstillingerne fra sidste gang man var inde i TimeOut. <br /><br />
		            Problemet løses ved at sætte "privacy"  internet indstillingerne korrekt og 
		            tilføje outzource.dk domænet under "trustet sites". Du kan læse mere om at indstille browseren korrekt 
		            på denne <a href="http://www.timeanddate.com/custom/cookiesie.html" target="_blank">her (åbner i nyt vindue)</a>.
		            
		            </p>
		            -->
                     <b class=news>5 jul. 2010</b> 
		            <p><b>Sommerferie</b><br />
		            OutZourCE holder sommerferie fra mandag d. 5/7 til fredag d. 23/7 2010 (uge 27-29) begge dage incl. <br /><br />
                    Ved nedbrud og andre uopsættelige hændelser kan i kontakte OutZourCE 
                    på tel. +45 26842000 eller på sk@outzource.dk <br /><br />

                    Vi tjekker mails lejlighedsvis, og svarer tilbage ved først kommende lejlighed.<br /><br />

                    God sommer.


		            </p>
                    <br /><br />


		            <b class=news>14 jun. 2010</b> 
		            <p><b>Næste release</b>
		            
		            Vi forventer at være klar med vores juli release lige efter sommerferien, dvs omkring 1 august. Den primære tilretning i vores nye release vil være en ny forkalkulation "on the fly" med mulighed for at tilføje stam-aktiviteter 
                    ved joboprettelse uden re-load af siden.<br /><br />
                    Herudover vil der være en forbedret print job funktion, ny pipeline, samt diverse forbedringer til afstemning, feriekalender og overblik over mest lønsomme kunder mm. 
                    </p>
                    <br /><br />
		            
		            
		              <b class=news>19 apr. 2010</b> 
		            <p><b>Stade indmelding på igangværende job</b>
		            
		            Vi har opdateret listen over <a href="login_screenshots.asp" target="_blank">igangværende job</a> (TSA --> Projektplan & Ressourcer --> Igangværende job)
                    Således at man nu kan angive ”Stade indmelding” ved at angive et forventet rest-estimat på jobbet.
                    <br /><br />
                    Stade beregnes udfra:  forkalkulerede timer – (realiserede timer + rest. estimat) = Stade i %
                    <br /><br />
                    Desuden vises jobbet afsluttet %, i en progess-bar og
                    Det er blevet muligt at trække/prioitere i job via drag’n drop.
                    </p>
                    <br /><br />
		            
		            
		             <b class=news>21 mar. 2010</b> 
		            <p><b>Teamleder funktion</b><br />
		            Vi har idag implementeret vores nye teamleder funktion, der betyder at man som teamleder for en projektgruppe, kan følge timeforbrug, udlæg og selv varetage timeregistreringer for medarbejdere i gruppen.
		            Man vælger Teamledere under <b>Job - Projektgrupper - Se medlemmer</b>
		            <br /><br />Hvis i har problemer med  der ikke længere kan se stat. for andre medarbejdere, kan I ringe eller <a href="mailto:support@outzource.dk">skrive</a> til os og vi vil hjælpe jer med at få oprettet de korrekte teamledere med det samme.
		             </p>
		            
		             <b class=news>25 feb. 2010</b> 
		            <p><b>Joboverblik</b><br />
		            Vi har inde på selv jobbet tilføjet et joboverbliks-billede, således at man kan se alle filer (tilbud), fakturaer og aktiviteter på jobbet i et samlet overblik.
		            </p>
		    
		    
		            <b class=news>23 feb. 2010</b> 
		            <p><b>Nyhedsbrev feb 2010</b><br />
		            Læs om vores seneste opdateringer, bl.a fakturering, timeregistrering og mere dynamisk aktivitetsliste og tildeling af stam-aktivitetsgrupper.
		            <br /><br />Få også mere at vide om hvad der er på vej i næste opdatering.
		            <a href="http://www.outzource.dk/nyhedsbrev_feb_2010.htm" target="_blank">Se vores nyhedsbrev her</a> 
		            </p>
		    
		            <b class=news>6 jan. 2010</b> 
		            <p><b>Vi har idag lanceret vores opdaterede fakturering der har følgende hovedpunkter:</b>
		           
		            <ul>
		            <li class=news> Hurtigere loadtid gennem bl.a kun at vise de medarbejdere der har timeforbrug i den valgte periode.
		            <li class=news> Tilføje alternative fakturaadresser.
		            <li class=news> Oprette materialeforbrug/udlæg undervejs, samt fakturere udlæg i hovedoverskrifter.
		            <li class=news> Nemmere ændring af faktura systemdato/periode.
		            <li class=news> Oversigt over forkalk. timer på aktiviteter.
		            <li class=news> Samle aktivitet til at fakturere fastpris job/samle sum-aktiviterne til en linie.
		            </ul>
		            <a href="help_and_faq/TimeOut_fakturering_rev_091220-2.pdf" target="_blank" class=vmenu>Læs mere i faktura manualen her..</a>
		            <br /><br />
		          
		            <p>
		            I samme forbindelse er det blevet muligt at prissætte hver enkelt aktivitet udfra enhedsprisen på aktiviteten, samt at bruge aktiviteterne som grundlag for det samlede budget på et job.
		            
			</p>
				 <hr style="width:260px; border:1px #cccccc dashed; height:1px;" />
		       
		          <b class=news>23 dec. 2009</b> <p>
		        OutZourCE ønsker alle vores kunder og samarbejdspartnere en rigtig god jul og et godt nytår.<br /><br />
		        Vi glæder os til det nye år, og for at komme frisk fra start, lancerer vi vores nye fakturerings del fuld af brugervenlig teknologi, lige efter nytår. 
		        <br /><br />
		        Vi holder lukket mellem jul og nytår, men har I problemer eller spørgsmål kan i ringe på +45 2684 2000.<br /><br />
		        Med velig hilsen<br /><br />
		        OutZourCE
			</p>
				 <hr style="width:260px; border:1px #cccccc dashed; height:1px;" />
		       
		         <b class=news>8 dec. 2009</b> <p>
		        Fra idag af er sortering af aktiviter indbyrdes mellem hinanden, indenfor hver fase, blevet muligt via drag'n drop funktionalitet fra aktivitetslisten under job. 
		        <br /><br />
			</p>
				 <hr style="width:260px; border:1px #cccccc dashed; height:1px;" />
		       
		       
		       
		       
		       <b class=news>4 dec. 2009</b> <p>
		       Vi er igang med at forenkle fakturerings-processen i TimeOut.
		       Vi har forbedret loadtiden med ca. 200%, hvilket især kommer til udtryk ved mange medarbejdere eller mange aktiviteter.
		       <br /><br />Herudover har vi tilføjet mere fleksibilitet og flere funktioner. Vi forventer at tilretinnger er klar i løbet af en uges tid.
		       <br /><br />
			</p>
				 <hr style="width:260px; border:1px #cccccc dashed; height:1px;" />
		       
		       
		       <b class=news>20 nov. 2009</b> <p>
		       For at gøre overblikket bedre i det nye søge filter har vi sorteret listen efter <b>kundenavne</b> først, og gjort søgningen "loose", dvs. den søger 
		       også på en del af navnet <b>midt i jobnavnet</b>. Desuden søger den også på <b>kundenavne</b>.
		       <br /><br />
		       Vi har desuden gjort det muligt for jer selv at vælge om i vil bruge den nye version af timereg. siden eller gå
		       tilbage til den gamle et stykke tid endnu.<br /><br />
		       
			   Æ,ø,å problem i det nye søgefilter rettet. <br /><br />
			   
			   
			    
			   <br /><br />
			</p>
				 <hr style="width:260px; border:1px #cccccc dashed; height:1px;" />
		       
		       <b class=news>18 nov. 2009</b> <p>
			   Det nye søgefilter på timeregistrerings siden er implementeret. 
			   <br /><br />
			</p>
				 <hr style="width:260px; border:1px #cccccc dashed; height:1px;" />
		      
		        <b class=news>16 nov. 2009</b> <p>
			    Vi er ved at lægge sidste hånd på vores nye søge-filter på timeregistrerings siden. <a href="ill/timereg_09.pdf" target="_blank">Pre. release-notes </a> <br /><br />
			    Det nye søgefilter er en del af en proces hvor vi har forsøgt at forenkle timeregistrering siden.<br /><br />
			    Vi har opdelt sidens informationer og funktioner i primære og sekundære, så 
			    det væsentlige kommer til at fylde mest og det sekundære bliver gemt lidt af vejen. 
			    <br /><br />Dette skulle gerne ende ud med et resultat, hvor vi står med en forenklet og mere funktionel
			    version af timeregisterings siden til glæde for alle.<br /><br />
			</p>
				 <hr style="width:260px; border:1px #cccccc dashed; height:1px;" />
		    
		    
		    <b class=news>3 nov. 2009</b> <p>
			Som en del af vores efterår 2009 re-lease er vi nu klar med vores nye print job funktion.<br />
			<br />Det er muligt at printe et job som et tilbud, en arb. besk. eller revisition, og
			I kan vælge at gemme dokumentet i fil-arkivet, printe ud, eller gemme som PDF.<br /><br />
		    Man kan gemme dokumentet som skabelon, for at sikre ens-artethed i de materialer der sende til kunder.
			<br /><br />
			<a href="http://www.screentoaster.com/watch/stV0JTRkJIR1xYQllaXF5aV1JQ/timeout_ny_print_job_funktion">Se det hele her i vores video toast. </a>
			 (der er lidt støj på optagelsen)
			</p>
				 <hr style="width:260px; border:1px #cccccc dashed; height:1px;" />
		    
		    		    
		    <b class=news>19 okt. 2009</b> <p>
			TimeOut har været igennem en ikke planlagt sessions-genstart idag kl. 14.00. Dette har betydet at I har modtaget en besked om at sessionen har været løbet ud, og er blevet bedt om at logge på igen.
			</p>
				 <hr style="width:260px; border:1px #cccccc dashed; height:1px;" />
		    
		    <b class=news>1 okt. 2009</b> <p>
			Vores næste re-lease er sat til at blive lanceret i slutningen af oktober. Den indeholder bl.a en række opdateringer til fakturerings processen, materiale registreringen, 
			samt et række tilretinger der skal give bedre overblik over timepriser på job for de enkelte medarbejdere. Senere kommer også faser på aktiviteter og den jobansvarliges forecast af tid pr. medarbejder på job. 
			</p>
				 <hr style="width:260px; border:1px #cccccc dashed; height:1px;" />
		    
		     <b class=news>25 sep. 2009</b> <p>
			Arbejder på at få vores udskudte september re-lease klar. Har pre-lanceret en opdateret Afstemnings År -> Dato 
			oversigt, så den viser total for periode, samt udviklingen måned for måmed.<br /><br />
			</p>
				 <hr style="width:260px; border:1px #cccccc dashed; height:1px;" />
		   
		    
		    <b class=news>22 sep. 2009</b> <p>
			Pga. sikkerheds problem med twitter er vi gået tilbage til at servere vores nyheder via alm. nyhedsfeed her på siden.<br /><br />
			</p>
				 <hr style="width:260px; border:1px #cccccc dashed; height:1px;" />
		   
			<b class=news>22 sep. 2009</b> <p>
			Arbejder på focus på kommentarfeltet på timereg. siden, samt at få inført faser på aktiviteter.<br /><br />
			</p>
				 <hr style="width:260px; border:1px #cccccc dashed; height:1px;" />
			
			<b class=news>18 sep. 2009</b> <p>
			Vores september re-lease er blevet udsat ca. 2 uger. Vi forventer at være klar ca. d. 1/10-1009.
			Vi vil løbende holde jer informeret her på siden.
			</p>
			
			    <hr style="width:260px; border:1px #cccccc dashed; height:1px;" />
			
			</td></tr></table>
			
		  </div>
        
		
		
		</td>
		
		
		</tr>
		</table>
		
		</div> <!-- cointent -->
		</center>
		<br /><br />
		
		
		<script>
		function NewWin(url)    {
		window.open(url, 'Disclaimer', 'width=600,height=580,scrollbars=yes,toolbar=no');    
		}
		</script>
		<br /><br /><br />
		<center>
		<font class=lillehvid>© 2002 - <%=year(now) %> OutZourCE og OutZourCE -underleverandører. Alle rettigheder forbeholdes.&nbsp;&nbsp;|&nbsp;&nbsp; 
		Tænkt og produceret af OutZourCE&nbsp;&nbsp;|&nbsp;&nbsp; Design af Margin Media<br>
		Graphical elements copyright ©Microsoft Corp&nbsp;&nbsp;|&nbsp;&nbsp;<a href="javascript:NewWin('timereg/disclaimer.asp');" target="_self"><font class="lillehvid"><u>Betingelser</u></a></font>&nbsp;&nbsp;|&nbsp;&nbsp;Optimeret til IE 5+ 1024 x 768
		<% 
		select case lto
		case "infow_demo", "infow", "assurator", "viatech", "pcmanden", "dreist", "rage", "enho"
		%>
		&nbsp;&nbsp;|&nbsp;&nbsp;<b>Support og Salg:</b> salg@informationworker.dk &nbsp;&nbsp;|&nbsp;&nbsp;Tlf: 70212128 
		<%case else%>
		&nbsp;&nbsp;|&nbsp;&nbsp;<b>Support og Salg:</b> timeout@outzource.dk&nbsp;&nbsp;|&nbsp;&nbsp;Tlf: +45 26 84 20 00&nbsp;&nbsp;|&nbsp;&nbsp;Support: +45 36 965 565 ell. Skype: OutZourCE
		
		<%end select%>
		<br><br>&nbsp;
		</center>
		
		
		
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
					            Set objF = objFSO.GetFile("d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\log\logfile_timeout_"&lto&".txt")
					            Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\log\logfile_timeout_"&lto&".txt", 8)
            				
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
			
			    oConn.execute("UPDATE medarbejdere SET lastlogin = '"& strLastLogin &"', Mnavn = '"&session("user")&"', timereg = "& treg0206 &" WHERE Mid ="& strUsrId &"")
		        
		        
		       
		       '*** Ferie Pl / Ferie afholdt **' 
		        call opdaterFeriePl(session("rettigheder"))
                
                
                
		        'Response.end
		        
		        
		       
		        
				
				'*** Opdater login_historik (Stempelur) ****'		
				if len(request("FM_stempelur")) <> 0 then
				intStempelur = request("FM_stempelur")
				else
				intStempelur = 0
				end if
				
				
				    
				'**** Tjekker om der skla laves en auto-logud eller om der er et igangværende logind på dagen ***'    
				call logindStatus(strUsrId, intStempelur)
				
				
				
				
				'Response.End
		
		else
		oConn.execute("UPDATE kontaktpers SET lastlogin = '"& strLastLogin &"' WHERE id="& thisKpid  &"")
		end if
		'***********************************'
		'***** Medarbejder kunde login  END '
		
		
			
		
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
						    if lto = "essens" then
						    response.redirect "timereg/crmhistorik.asp?menu=crm&ketype=e&func=hist&selpkt=hist"
						    else
						    response.redirect "timereg/crmkalender.asp?menu=crm&shokselector=1&ketype=e&selpkt=kal"
						    end if
						case 2
						response.redirect "timereg/jbpla_w.asp?menu=job"
						case else '0
							if cint(treg0206) = 1 then
							response.redirect "timereg/timereg_akt_2006.asp"
							'response.redirect "timereg/timereg_2006_fs.asp"
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


