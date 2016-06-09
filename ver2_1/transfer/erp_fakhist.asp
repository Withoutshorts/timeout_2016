<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/dato2.asp"-->
<!--#include file="../inc/regular/erp_func.asp"-->

<%

call TimeOutVersion()

if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	menu = "erp"
	
	%>
	<script>
	
	function vaelgkontakt(){
    var kid = 0;
	//document.getElementById("FM_aftale").value = "-1"
	//document.getElementById("FM_job").value = "-1"
	//document.getElementById("FM_sog").value = ""
	kid = document.getElementById("FM_kunde").value
	window.location.href = "erp_fakhist.asp?FM_job=0&FM_aftale=0&FM_sog=9999&FM_kunde="+kid
    }

    </script>
	
	<%
	
	
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	
	func = request("func")
    thisfile = "erp_tilfakturering.asp"
    print = request("print")
	
	select case func 
	case "opdbetalt"
	
	
	'Response.Write request("erfakbetalt") & "<br>"
	
	fids = split(trim(request("erfakbetalt")), ", ")
	thisKom = split(trim(request("FM_fakbetkom")), ", #, ")
	for x = 0 to UBOUND(fids)
	
	thisval = trim(request("fakid_"&fids(x)&""))
	
	
	strSQL = "UPDATE fakturaer SET erfakbetalt = "& thisval &", fakbetkom = '"& replace(thisKom(x), "'", "''") &"' WHERE fid = "& fids(x)
	'Response.Write strSQL & "<br>"
	oConn.execute(strSQL)
	
	next
	
	'Response.end
	
	Response.redirect "erp_fakhist.asp?FM_job="&request("jobid")&"&FM_aftale="&request("aftid")&"&FM_sog="&request("FM_sog")&"&sort="&request("sort")
   
	
	
	case else
	
	
	if print <> "j" then
	
	dTop = "132"
	dLeft = "20" 
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<script>

	    $(document).ready(function() {

	        $("#FM_kunde").change(function() {

	        $("#FM_sog").val('')
	        $("#fakhist").submit()

	        });

	        $("#FM_aftale").change(function() {

	        $("#FM_sog").val('')
	        $("#fakhist").submit()

	        });

	        $("#FM_job").change(function() {

	        $("#FM_sog").val('')
	        $("#fakhist").submit()

	        });




	    });
	
	</script>
	
	
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%call erpmainmenu(1)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	call erptopmenu()
	%>
	</div>
	
	<%else 
	
	dTop = "20"
	dLeft = "20" 
	
	%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	
	<%end if %>
	
	
	<div id="sindhold" style="position:absolute; left:<%=dLeft %>px; top:<%=dTop %>px; visibility:visible;">
	<%
	'**********************************************************
	'**************** Job til fakturering *********************
	'**********************************************************
	
	
	
	
	%>
	<h4><img src="../ill/ac0010-24.gif" /> Faktura historik (søg i fakturaer)</h4>
	
	<%
	
	ptop = 0
	pleft = 0
	pwdt = 802
	
	call filterheader(ptop,pleft,pwdt,pTxt) %>
	<table cellspacing=0 cellpadding=2 border=0 width=100%>
	<form action="erp_fakhist.asp" id="fakhist" method="POST">
	<tr>
	<td valign=top width=100>
	
    <% 
    if len(request("FM_kunde")) <> 0 then
			
			'if request("FM_kunde") = 0 then
			'kundeid = 0
			'sqlKundeKri = "t.tknr <> 0"
			'sqlKundeKri2 = "k.kid <> 0"
			'else
			kundeid = request("FM_kunde")
			'sqlKundeKri = "t.tknr = "& valgtKunde &""
			'sqlKundeKri2 = "k.kid = "& valgtKunde &""
			'end if
			
	else
		
			if len(request.cookies("erp")("kid")) <> 0 AND request.cookies("erp")("kid") <> 0 then
			kundeid = request.cookies("erp")("kid")
			'sqlKundeKri = "t.tknr = "& valgtKunde &""
			'sqlKundeKri2 = "k.kid = "& valgtKunde &""
			else
			kundeid = 0
			'sqlKundeKri = "t.tknr <> 0"
			'sqlKundeKri2 = "k.kid <> 0"
			end if
			
	end if
	
	
	response.Cookies("erp")("kid") = kundeid
	response.Cookies("erp").expires = date + 60
	
	if len(request("FM_sog")) <> 0 AND request("FM_sog") <> "9999" then
	showSog = request("FM_sog")
	else
	showSog = ""
	end if
	
	if len(request("FM_job")) <> 0 then
	jobid = request("FM_job")
	else
	jobid = 0
	end if
	
	if len(request("FM_aftale")) <> 0 then
	aftid = request("FM_aftale")
	else
	aftid = 0
	end if
	
	if len(request("FM_viskunabne")) <> 0 AND request("FM_viskunabne") <> "0" then
	viskunabne = 1
	else
	viskunabne = 0
	end if
	
	if len(trim(showSog)) <> 0 then
	aftid = 0
	jobid = 0
	kundeid = 0
	end if
	
	select case jobid
	case -1
	vlgtJob = "Ingen"
	case 0
	vlgtJob = "Alle"
	end select 
	
	select case aftid
	case -1
	vlgtAft = "Ingen"
	case 0
	vlgtAft = "Alle"
	end select 
	
	if len(trim(request("FM_jobikkedelafaftale"))) <> 0 then
	jidaftCHK = "CHECKED"
	jidaft = 1
	else
	jidaftCHK = ""
	jidaft = 0
	end if
	
    if len(trim(request("FM_viskun_interne"))) <> 0 then
	viskintCHK = "CHECKED"
	viskint = 1
	else
	viskintCHK = ""
	viskint = 0
	end if
	
    
    %>
    <b>Kontakter:</b></td><td valign=top>
        <%if print <> "j" then %>
        <select id="FM_kunde" name="FM_kunde" style="font-size : 9px; width:325px;"> 
        
		<option value="0">Alle</option>
		<%end if
		ketypeKri = " ketype <> 'e'"
		strKundeKri = " AND kid <> 0 "
		vlgtKunde = " Alle "
		
				strSQL = "SELECT Kkundenavn, Kkundenr, Kid, COUNT(f.fid) AS antalfak FROM kunder "_
				&" LEFT JOIN fakturaer f ON (f.fakadr = kid AND shadowcopy = 0) "_
				&" WHERE "& ketypeKri &" "& strKundeKri &" AND f.fid <> 0 GROUP BY kid ORDER BY Kkundenavn"
				
				
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(kundeid) = cint(oRec("Kid")) then
				isSelected = "SELECTED"
				vlgtKunde = oRec("Kkundenavn") & " ("& oRec("Kkundenr") &")"
				else
				isSelected = ""
				end if
				
				
				if print <> "j" then
				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%> (<%=oRec("kkundenr") %>)</option>
				<%
				end if
				oRec.movenext
				wend
				oRec.close
				
				if print <> "j" then%>
		        </select><br>
		        <% 
		        else
		        %>
				<%=vlgtKunde %>
				<% 
				end if
				%>
	    
	    <!--
	    <b>Kontaktansvarlige (medarbejder):</b><br> <select name="FM_kans" size="1" style="font-size : 9px; width:285px;">
		<option value="0">Alle</option>
	    </select>-->
	    
	    
	    
	    </td>
	    
	    <%
	    if len(trim(request("sort"))) <> 0 then
	    sort = request("sort")
	    

        sortCHK4 = ""
        sortCHK3 = "" 
	    sortCHK2 = ""
	    sortCHK1 = ""
	            
	            select case sort
	            case 2
	           
	            sortCHK2 = "CHECKED"
	            sortOrderKri = "f.fakdato DESC, f.faknr"
	            case 3
	            sortCHK3 = "CHECKED"
	             sortOrderKri = "f.b_dato DESC, f.faknr"
                case 4
	            sortCHK4 = "CHECKED"
	            sortOrderKri = "f.labeldato DESC, f.faknr"
	            case else
	            sortCHK1 = "CHECKED"
	            sortOrderKri = "f.faknr"
	            end select
	            
	   
	    else
          select case lto
          case "epi", "intranet - local"
            sort = 4
            sortCHK4 = "CHECKED"
	        sortCHK3 = ""
	        sortCHK2 = ""
	        sortCHK1 = ""
	        sortOrderKri = "f.labeldato DESC, f.faknr"
          case else
	        sort = 2
            sortCHK4 = ""
	        sortCHK3 = ""
	        sortCHK2 = "CHECKED"
	        sortCHK1 = ""
	        sortOrderKri = "f.fakdato DESC, f.faknr"
         end select
	    end if
	    
	    
	    'if request("showlabel") = "1" OR (request.Cookies("erp")("showlabel") = "1" AND len(trim(request("sort"))) = 0) then
	    'showlabel = 1
	    'showlabelCHK1 = "CHECKED"
	    'showlabelCHK0 = ""
	    'else
	    'showlabel = 0
	    'showlabelCHK1 = ""
	    'showlabelCHK0 = "CHECKED"
	    'end if
	    
	    'response.cookies("erp")("showlabel") = showlabel
        
        if cint(viskunabne) <> 0 then
        viskunabneCHK = "CHECKED"
        else
        viskunabneCHK = ""
        end if
            
            
        if print <> "j" then%>
        <td rowspan=3 style="padding:2px 2px 2px 20px;" valign=top>
	    <b>Periode:</b> (fakturadato)<br /> 
	    <!--#include file="inc/weekselector_s.asp"-->&nbsp;<br />
	    
        <input id="FM_usedatokri" name="FM_usedatokri" value="1" type="hidden" />
        <br />
        <b>Sortér efter</b><br />
        
        <input id="sort" name="sort" type="radio" value="1" <%=sortCHK1 %> />Faktura Nr.<br /> 
        <input id="sort" name="sort" type="radio" value="2" <%=sortCHK2 %> />Faktura Sys. Dato<br />
        <input id="sort" name="sort" type="radio" value="4" <%=sortCHK4 %> />Faktura Labeldato<br />
        <input id="sort" name="sort" type="radio" value="3" <%=sortCHK3 %> />Forfaldsdato<br />
        
            <!--<br />
            <br />Vis fakturaer der benytter <b>faktura "label" dato</b>: <br />
            <input id="showlabel" name="showlabel" value="1" type="radio" <=showlabelCHK1 %> /> Kun hvis <b>label-dato</b> ligger i den valgte periode.<br />
            <input id="showlabel" name="showlabel" type="radio" value="0" <=showlabelCHK0 %> /> Uanset om <b>label-dato</b> ligger i den valgte per.
            
            -->
            
            <br />
            
            
            <input id="FM_viskunabne" name="FM_viskunabne" value="1" type="checkbox" <%=viskunabneCHK%> /> Vis kun åbne poster.
            


            <br />
            <input id="Checkbox1" name="FM_viskun_interne" value="1" type="checkbox" <%=viskintCHK%> /> Vis kun interne.

        </td>
	    <%else

        %>
        <td>
        <b>Kun kun åbne/interne fakturaer:</b>
        <%if cint(viskunabne) <> 0 then %>
        <br />- Vis kun åbne
        <%end if %>

        <%if cint(viskint) <> 0 then %>
        <br />- Vis kun interne
        <%end if %>
        
        </td>
        <%

	    dontshowDD = 1
	    %>
	    <!--#include file="inc/weekselector_s.asp"-->
	    <%end if %>
	    
	 </tr>
	 
	  <tr><td valign=top>
       
        <b>Aftaler:</b>
        </td><td valign=top>
	    
	    <% 
	    strSQLaft = "SELECT s.id AS sid, s.navn, s.aftalenr, COUNT(f.fid) AS antalfak "_
	    &" FROM serviceaft s "_
	    &" LEFT JOIN fakturaer f ON (f.aftaleid = s.id AND shadowcopy = 0) "_
	    &" WHERE s.kundeid = " & kundeid & " AND f.fid <> 0 GROUP BY s.id ORDER BY s.navn"
	  
	  'Response.Write strSQLjob
	  'Response.flush
	    %>
	    
	        <% if print <> "j" then%>
	        <select id="FM_aftale" name="FM_aftale" size=3 style="font-size : 9px; width:325px;">
            
            <%if cint(aftid) = -1 then
            ingSEL = "SELECTED"
            else
            ingSEL = ""
            end if
            %>
            
            <option value=-1 <%=ingSEL %>>Ingen</option>
            <%if cint(aftid) = 0 then
            nulSEL = "SELECTED"
            else
            nulSEL = ""
            end if
            %>
            <option value=0 <%=nulSEL %>>Alle</option>
            
            <%
            end if
            
            oRec.open strSQLaft, oConn, 3
				while not oRec.EOF
				
				if cint(aftid) = cint(oRec("sid")) then
				isSelected = "SELECTED"
				vlgtAft = oRec("navn") & " ("& oRec("aftalenr") &")"
				else
				isSelected = ""
				end if
				
				
				if print <> "j" then
				%>
				<option value="<%=oRec("sid")%>" <%=isSelected%>><%=oRec("navn")%> (<%=oRec("aftalenr")%>) - <%=oRec("antalfak") %></option>
				<%
				end if
				oRec.movenext
				wend
				oRec.close
				
				if print <> "j" then%>
                </select>
                <%else %>
                <%=vlgtAft %>
                <%end if %>
        
        </td></tr>
	    
	    </tr>
        <tr><td valign=top><b>Job:</b>
        <br />Kun kontakter, job og aftaler<br /> med fakturaer på vises.</td><td valign=top>
	    <%
	    if cdbl(aftid) > 0 then
	    aftidKri = " AND serviceaft = " & aftid
	    else
	    aftidKri = ""
	    end if
	    
	    if kundeid <> 0 then
	    strjobKidSQL = "jobknr = " & kundeid
	    else
	    strjobKidSQL = "jobknr <> 0 "
	    end if
	    
	    strSQLjob = "SELECT j.id AS jid, j.jobnavn, j.jobnr, COUNT(f.fid) AS antalfak "_
	    &" FROM job j "_
	    &" LEFT JOIN fakturaer f ON (f.jobid = j.id AND shadowcopy = 0) "_
	    &" WHERE "& strjobKidSQL &" "& aftidKri &" AND f.fid <> 0 GROUP BY j.id ORDER BY j.jobnavn"
	  
	  'Response.Write strSQLjob
	  'Response.flush
	  
	        if print <> "j" then %>
            <select id="FM_job" name="FM_job" size=8 style="font-size : 9px; width:325px;">
             <%if cint(jobid) = -1 then
            ingSEL = "SELECTED"
            else
            ingSEL = ""
            end if
            %>
            
            <option value="-1" <%=ingSEL %>>Ingen</option>
            
            <%if cint(jobid) = 0 then
            nulSEL = "SELECTED"
            else
            nulSEL = ""
            end if
            %>
            
            <option value=0 <%=nulSEL %>>Alle</option>
            <%
            end if
            
            oRec.open strSQLjob, oConn, 3
				while not oRec.EOF
				
				if cint(jobid) = cint(oRec("jid")) then
				isSelected = "SELECTED"
				vlgtJob = oRec("jobnavn") & " ("& oRec("jobnr") &")"
				else
				isSelected = ""
				end if
				
				
				if print <> "j" then
				%>
				<option value="<%=oRec("jid")%>" <%=isSelected%>><%=oRec("jobnavn")%> (<%=oRec("jobnr")%>) - <%=oRec("antalfak") %></option>
				<%
				end if
				oRec.movenext
				wend
				oRec.close
				%>
         
        <% if print <> "j" then%>   
        </select><br /><input id="FM_jobikkedelafaftale" name="FM_jobikkedelafaftale" value="1" type="checkbox" <%=jidaftCHK%> />Vis kun job der ikke er del af aftale.
        
        <%else%>
        <%=vlgtJob %>
        <%end if %>
        </td></tr>
       
      
	  <tr>
	  <td valign=top>
	    <br /><b>Søg:</b>&nbsp;<br />
	    Faktura nr., aft.nr eller jobnr.<br />
	    Ignorerer periode og valgt kontakt.</td><td>
	    <br /> 
	    <%if print <> "j" then %>
	    <input id="FM_sog" name="FM_sog" value="<%=showSog%>" type="text" style="font-size : 9px; width:285px;" /><br />
	    Faktura nr, eks. 54 - 103, % = wildcard <br />
	    <%else %>
	    <%=showSog %>
	    <%end if %>
	    </td>
	    <td align=right>
	    <%if print <> "j" then %>
	    <input id="Submit1" type="submit" value="Vis Fakturaer" />
	    <%end if %></td>
	 </tr>
	 
	 
	 </form></table>
	 </div>
	 
	 </td></tr>
	 </table>
	 </div>
	
	<%
	sqlDatostart = strAar&"/"&strMrd&"/"&strDag  'year(datointervalstart)&"/"& month(datointervalstart)&"/"&day(datointervalstart) 
	sqlDatoslut = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut 'year(datointervalslut)&"/"& month(datointervalslut) &"/"& day(datointervalslut)
	%>
	
	<br />
	<br><b>Periode:</b>&nbsp;
	<%=formatdatetime(strDag&"/"& strMrd &"/"& strAar, 1) & " - " & formatdatetime(strDag_slut &"/"& strMrd_slut &"/"& strAar_slut, 1)%>
	
	<br />
    
    <%
    'Response.Write "jobid: " & jobid
    'Response.Write "<br>aftid: " & aftid 
    
    'Opretter en instans af fil object **'
    Set fs=Server.CreateObject("Scripting.FileSystemObject")



    
    
    tTop = 20
    tLeft = 0
    tWdth = 1104
    
    call tableDiv(tTop,tLeft,tWdth)
    
    %>
    
    <table cellspacing=0 cellpadding=2 border=0 width=100% bgcolor="#ffffff">
    <form action="erp_fakhist.asp?func=opdbetalt&FM_job=<%=jobid%>&FM_aftale=<%=aftid%>&FM_sog=<%=showSog%>&sort=<%=sort %>" method="post">
    <tr bgcolor="#8cAAe6">
        <td class=alt height=25 style="padding:0px 5px 0px 5px;"><b>Kontakt</b><br />Job / Aftale</td>
        <!--<td class=alt><b>Job</b></td>
        <td class=alt><b>Aftale</b></td>-->
        <td class=alt colspan=3><b>Faktura nr.</b></td>
        <td class=alt><b>Faktura dato</b><br />L: Label dato<br />
        F: Fak. sys. dato</td>
        <td class=alt><b>Forfaldsdato</b></td>
        <td class=alt style="padding:0px 0px 0px 2px;"><b>Status</b></td>
        <td class=alt><b>Type</b></td>
         <!--<td class=alt align=right style="padding:0px 5px 0px 0px;"><b>Antal</b></td>-->
        <td class=alt width=100 align=right style="padding:0px 5px 0px 0px;"><b>Fak. Beløb <br />(excl. moms)</b></td>
        
        
        <td class=alt align=right style="padding:0px 5px 0px 0px;"><b>Vente t.</b><br />
        Brugt / Ultimo</td>
         <td class=alt><b>Konto</b><br /><b>Modkonto</b></td>
         <td class=alt><b>Lukket?<br />(betalt)</b></td>
         <td class=alt><b>Kommantar</b></td>
         <td>
             &nbsp;</td>
        
    </tr>
	<%  
	
	
	
	if len(trim(showSog)) <> 0 then
	    
	    if instr(showSog, "-") <> 0 then
	    
	    whereInstr = instr(showSog, "-")
	    showSog1 = mid(showSog, 1, whereInstr-1)
	    showSog1 = trim(showSog1)
	    
	    lenInstr = len(showSog)
	    showSog2 = mid(showSog, whereInstr+1, lenInstr)
	    showSog2 = trim(showSog2)
	    
	    jobaftSQL = " f.faknr BETWEEN "& showSog1 &" AND "& showSog2 &""_
	    &" OR (j.jobnr BETWEEN "& showSog1 &" AND "& showSog2  &""_
	    &" OR s.aftalenr BETWEEN "& showSog1 &" AND "& showSog2  &") AND shadowcopy <> 1"
	    
	    else
	    
	    jobaftSQL = " (f.faknr LIKE '"& showSog &"' AND shadowcopy <> 1) OR "_
	    &" (j.jobnr LIKE '"& showSog &"' AND shadowcopy <> 1) OR "_
	    &" (s.aftalenr LIKE '"& showSog &"' AND shadowcopy <> 1)"
	    
	    end if
	    
	        %>
			<!--#include file="inc/isint_func.asp"-->
			<%
			call erDetInt(showSog) 
			if (isInt > 0 OR instr(showSog, ".") <> 0) AND instr(showSog, "%") = 0 AND instr(showSog, "-") = 0 then
				
		    jobaftSQL = " f.faknr = 9999999 "
			
			isInt = 0
			end if
	
	else
	    
	    if cint(kundeid) <> 0 then
	    kundeIDKri = " AND f.fakadr = " & kundeid 
	    else
	    kundeIDKri = " AND f.fakadr <> 0 "
	    end if
	    
		
		    
		    if cint(aftid) <> 0 then
		    jobaftSQL = "(f.aftaleid = "& aftid &")"
		    else
		    jobaftSQL = "(f.aftaleid <> 0 AND f.jobid = 0 "& kundeIDKri  &") " 
		    end if
		
		
		
		    if cint(jobid) = 0 then
		    jobaftSQL = jobaftSQL & " OR (f.jobid <> 0 AND shadowcopy <> 1 "& kundeIDKri &" ) "
		    else
		    jobaftSQL = jobaftSQL & " OR ((f.jobid = "& jobid &" AND shadowcopy <> 1) "& kundeIDKri &") "
		    end if
		    

            if cint(viskint) = 1 then
            intSQLKri = " AND medregnikkeioms = 1"
            else
            intSQLKri = ""
            end if
	
	
	end if
	    
	    
	      
	    
	    
	    kundeidSQL = "k.kid = f.fakadr"
		
		strSQLFak = "SELECT k.kid, f.jobid, f.aftaleid, f.fid, f.fakdato, f.b_dato, f.faknr, f.betalt, "_
		&" f.faktype, f.beloeb, f.timer AS fak, f.shadowcopy, f.erfakbetalt, f.kurs, "_
		&" k.kkundenavn, k.kkundenr, f.brugfakdatolabel, f.istdato2, f.fakbetkom, "_
		&" j.jobnavn, j.jobnr, j.jobslutdato, j.jobstartdato, j.fastpris, j.serviceaft, "_
		&" s.navn AS aftnavn, s.advitype , s.aftalenr, s.pris, "_
		&" s.stdato, s.sldato, sj.navn AS sjaftnavn, sj.aftalenr AS sjaftalenr, "
		
	    strSQLFak = strSQLFak &" k1.navn AS konto, k1.kontonr AS kontonr, "_
		&" k2.navn AS modkonto, k2.kontonr AS modkontonr, "_
		&" sum(fms.venter) AS ventetimer, sum(venter_brugt) AS ventetimer_brugt, v.valutakode, v.id AS valid, labeldato, medregnikkeioms, fak_laast "_
		&" FROM fakturaer f "
		
		strSQLFak = strSQLFak &" LEFT JOIN kunder k on ("&kundeidSQL&")"_
		&" LEFT JOIN job j ON (j.id = f.jobid)"_
		&" LEFT JOIN kontoplan k1 ON (k1.kontonr = f.konto )"_
		&" LEFT JOIN kontoplan k2 ON (k2.kontonr = f.modkonto )"_
		&" LEFT JOIN fak_med_spec fms ON (fms.fakid = f.fid) "_
	    &" LEFT JOIN serviceaft s ON (s.id = f.aftaleid)"_
	    &" LEFT JOIN serviceaft sj ON (sj.id = j.serviceaft)"_
	    &" LEFT JOIN valutaer v ON (v.id = f.valuta)"
	    
		
	    strSQLFak = strSQLFak &" WHERE ("& jobaftSQL &")"
		
		if len(trim(showSog)) = 0 then 
		   
		    strSQLFak = strSQLFak &" AND (f.fakdato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"'"
		   
		    
		    'if showlabel = 1 then 'Faktura label
		    'strSQLFak = strSQLFak &" AND brugfakdatolabel = 0 " 
		    'end if
		    
		    strSQLFak = strSQLFak &" OR (brugfakdatolabel = 1 AND (f.labeldato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"'))"
		    
		    
		    strSQLFak = strSQLFak &")"
		end if
		
        '** vis kun åbne
		if cint(viskunabne) <> 0 then
		strSQLFak = strSQLFak & " AND erfakbetalt = 0"
		end if
	    
        '** vis kun interne
		strSQLFak = strSQLFak & intSQLKri

		strSQLFak = strSQLFak &" GROUP BY f.fid ORDER BY "& sortOrderKri &" DESC"
	        
	        'Response.Write strSQLFak &"<br>" 
	        'Response.Flush
	        
		f = 0
		eksportFid = 0
		faksumialt = 0
		fakantalIalt = 0
		belobGrundValTot = 0
		antalAbneposter = 0
		febkbelobAbneposter = 0
        internTotFak = 0
        jobTotFak = 0
        aftTotFak = 0
        kreTotFak = 0
		oRec3.open strSQLFak, oConn, 3
        while not oRec3.EOF
        
        
        
        if oRec3("jobid") <> 0 then 
        bgThis = "#FFFFFF"
        else
        bgThis = "#EFf3ff"
        end if
        
       
            '** Henter Job / aftale oplysninger på faktura ***
            '** A) jobid <> 0 = Faktura oprettet på job.
            
            '** B) aftaleid <> 0 Faktura oprettet på aftale.
            
            '** C) jobid og shadowcopy <> 0 Faktura oprettet på aftale, men
            '** der er et job tilknyttet denne aftale, og derofr bliver der også oprettet en faktura på jobbet.
            
            
            
           
            
            if (oRec3("jobid") <> 0 AND jidaft = 0) OR (oRec3("jobid") <> 0 AND jidaft = 1 AND oRec3("serviceaft") = 0) OR oRec3("aftaleid") <> 0 then 
            
           
            %>
            <tr bgcolor="<%=bgThis %>">
            <td valign=top style="border-bottom:1px #C4C4C4 dashed; padding:5px 5px 3px 5px;">
            <span style="color:darkred;"><%=oRec3("kkundenavn") %> (<%=oRec3("kkundenr") %>)</span><br />
            
            
            <%if oRec3("jobid") <> 0 then %>
            
            <u>Job:</u> <b><%=oRec3("jobnavn") %> (<%=oRec3("jobnr") %>)</b>
            <font class=megetlillesort>
            <%if oRec3("fastpris") = "1" then %>
            <i>- fastpris</i>
            <%else %>
            <i>- lbn. timer</i>
            <%end if %>
            
            <br />
            <%if isdate(oRec3("jobstartdato")) AND isDate(oRec3("jobslutdato")) then  %>
            <%=replace(formatdatetime(oRec3("jobstartdato"), 2),"-",".") %> til <%=replace(formatdatetime(oRec3("jobslutdato"), 2),"-",".") %>
            <%else %>
            <%=oRec3("jobstartdato") %> til <%=oRec3("jobslutdato") %>
            <%end if %>
            </font>
            
            <%if oRec3("serviceaft") <> 0 then %>
            <br /><font class=medlillesilver><%=oRec3("sjaftnavn") %> (<%=oRec3("sjaftalenr") %>)</font>
            <%end if
           
           
            end if
            
            
            
            '*** Aftaler ***
            if oRec3("aftaleid") <> 0 then
            
            %>
            
            <u>Aftale:</u> <b><%=oRec3("aftnavn") %> (<%=oRec3("aftalenr") %>)</b>
            <font class=megetlillesort>
            <%
             if oRec3("advitype") <> 2 then
	        aftType = "Enh./Klip"
	        else
	        aftType = "Periode"
	        end if
             %>
             
             - <%=aftType %> - 
             
             <%if len(oRec3("stdato")) <> 0 AND len(oRec3("sldato")) <> 0 then %>
             <%=replace(formatdatetime(oRec3("stdato"), 2),"-",".")  %> til <%=replace(formatdatetime(oRec3("sldato"), 2),"-",".")  %>
             <%end if%>
            </font>
            
            <%strSQLaftjob = "SELECT jobnavn, jobnr FROM job WHERE serviceaft = "& oRec3("aftaleid")
            oRec.open strSQLaftjob, oConn, 3
            while not oRec.EOF 
            %>
            <br /><font class=medlillesilver><%=oRec("jobnavn") %> (<%=oRec("jobnr") %>)</font>
            <%
            oRec.movenext
            wend
            oRec.close %>
            <%end if %>
            
            
        
        
            
        </td>
        
        
        
        <td valign=top align=right style="border-bottom:1px #C4C4C4 dashed;">
        <% if print <> "j" then%>
        <a href="erp_fak_godkendt_2007.asp?id=<%=oRec3("fid")%>&aftid=<%=oRec3("aftaleid")%>&jobid=<%=oRec3("jobid")%>" class="vmenu" target="_blank"><%=oRec3("faknr") %></a>
        <%else %>
        <%=oRec3("faknr") %>
        <%end if %>
        
        </td>
         <td valign=top style="border-bottom:1px #C4C4C4 dashed; padding-right:5px;">
        <%
        '*** Findes PDF? ***
        if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\"&toVer&"\timereg\erp_fakhist.asp" then
	        pdfurl = "d:\\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\upload\"&lto&"\faktura_"&lto&"_"&oRec3("faknr")&".pdf"
	    else
	        pdfurl = "c:\\www\timeout_xp\wwwroot\"&toVer&"\inc\upload\"&lto&"\faktura_"&lto&"_"&oRec3("faknr")&".pdf"
	    end if
        
        	If (fs.FileExists(pdfurl))=true Then
                  Response.Write("<a href=""https://outzource.dk/timeout_xp/wwwroot/"&toVer&"/inc/upload/"&lto&"/faktura_"&lto&"_"&oRec3("faknr")&".pdf"" target=""blank""><img src=""../ill/ikon_pdf.gif"" border=""0""></a>")
            Else
                  Response.Write("&nbsp;")
            End If
            
          %>
          </td>
         <td valign=top style="border-bottom:1px #C4C4C4 dashed; padding:4px 5px 0px 0px;">
          <%
        
        '*** Findes XML ***
         if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\"&toVer&"\timereg\erp_fakhist.asp" then
	        xmlurl = "d:\\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\upload\"&lto&"\faktura_xml_"&lto&"_"&oRec3("faknr")&".xml"
	    else
	        xmlurl = "c:\\www\timeout_xp\wwwroot\"&toVer&"\inc\upload\"&lto&"\faktura_xml_"&lto&"_"&oRec3("faknr")&".xml"
	    end if
        
        	If (fs.FileExists(xmlurl))=true Then
                  Response.Write("<a href=""https://outzource.dk/timeout_xp/wwwroot/"&toVer&"/inc/upload/"&lto&"/faktura_xml_"&lto&"_"&oRec3("faknr")&".xml"" target=""blank""><img src=""../ill/ikon_xml.gif"" border=""0""></a>")
            Else
                  Response.Write("&nbsp;")
            End If
        
       
        
        %>
        
        
        <%if oRec3("jobid") <> 0 then
        thisJobId = oRec3("jobid")
        else
        thisJobId = 0
        end if
        
        if oRec3("aftaleid") <> 0 then
        thisAftaleId = oRec3("aftaleid")
        else
        thisAftaleId = 0
        end if
         %>
         
           
           
        </td>
        <td valign=top style="border-bottom:1px #C4C4C4 dashed; white-space:nowrap;">
        
        <%if cint(oRec3("brugfakdatolabel")) = 1 then %>
        L: <b><%=replace(formatdatetime(oRec3("labeldato"),2),"-",".")  %></b><br />
        <span style="font-size:9px; color:#999999;">(<%=replace(formatdatetime(oRec3("fakdato"),2),"-",".") %>)</span>
        <%else %>
        F: <b><%=replace(formatdatetime(oRec3("fakdato"),2),"-",".") %></b>
        <%end if %>
        
        </td>
        <td valign=top style="border-bottom:1px #C4C4C4 dashed;">
        <%if len(oRec3("b_dato")) <> 0 then %>
        <%=replace(formatdatetime(oRec3("b_dato"),2),"-",".") %>
        <%else %>
        -
        <%end if %>
        
        </td>
        <td valign=top style="border-bottom:1px #C4C4C4 dashed; padding:2px 5px 0px 5px;">
        <%if oRec3("betalt") = 1 then %>
       <img src="../ill/godkend_icon_16.gif" border="0" /> <font color=forestgreen>Godkendt</font>
        <%else %>
        
         <%if cint(oRec3("betalt")) = 0 AND (cdate(oRec3("fakdato")) > cdate("25/8/2007")) AND print <> "j" then %>
           <a href="../timereg/erp_opr_faktura_fs.asp?func=red&fid=<%=oRec3("fid")%>&FM_kunde=<%=oRec3("kid")%>&FM_job=<%=thisJobId%>&FM_aftale=<%=thisAftaleId%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd %>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>&reset=1">
           <img src="../ill/redigerfak.gif" border="0" alt="Rediger faktura" /> </a>
           <%end if %>
        
        <font color=silver>Kladde</font>
        <%end if %>
        </td>
        <td valign=top style="border-bottom:1px #C4C4C4 dashed;">
        <%select case oRec3("faktype")
        case 1
        strType = "Kreditnota"
        minus = "-"
        case else
        strType = "Faktura"
        minus = ""
        end select
        
        Response.Write strType
        %>

          <%if cint(oRec3("medregnikkeioms")) <> 0 then %>
         <br />(intern)
        <%end if %>

        </td>
        <!--
        <td valign=top align=right style="border-bottom:1px #C4C4C4 dashed; padding:1px 5px 0px 0px;">
        <formatnumber(minus&oRec3("fak"), 2)%>
        </td>
        -->
       
        <td valign=top align=right style="border-bottom:1px #C4C4C4 dashed; padding:1px 5px 0px 0px;"> 

        <%fakbelob = replace(oRec3("beloeb"), "-", "") %>

        <b><%=formatnumber(minus&fakbelob, 2)&" "&oRec3("valutakode") %></b>
      
        <br />
              
        <%
        '*** Beløb i grundvaluta == DKK ***'
        if oRec3("valid") <> 1 then 
            
            call beregnValuta(minus&(fakbelob),oRec3("kurs"),100)
            belobGrundVal = valBelobBeregnet %>
            
            ~ <%=formatnumber(belobGrundVal, 2)%> DKK
        
            <% 
        else
            
            belobGrundVal = minus&fakbelob/1
        end if 

        '** job ell. aftale **'
        
        



        if oRec3("faktype") = 1 then '** kreditnota
        kreTotFak = kreTotFak/1 + belobGrundVal
        else

            '** kreditnotaer skal ikke medregnes ***'
            if oRec3("jobid") <> 0 then '** jobfak
            jobTotFak = jobTotFak/1 + belobGrundVal 
            else
            aftTotFak = aftTotFak/1 + belobGrundVal
            end if

            if oRec3("medregnikkeioms") <> 0 then
            internTotFak = internTotFak/1 + belobGrundVal/1
            else
            belobGrundValTot = belobGrundValTot/1 + belobGrundVal/1 
            end if

            

        end if
        
        
        '*** Åbne poster ****'
        if cint(oRec3("erfakbetalt")) = 1 then
        febkbelobAbneposter = febkbelobAbneposter 
        else
            if oRec3("medregnikkeioms") = 0 then
            febkbelobAbneposter = febkbelobAbneposter + (belobGrundVal/1)  
            antalAbneposter = antalAbneposter + 1
            else
            febkbelobAbneposter = febkbelobAbneposter 
            end if      
        end if%>
        
       
        </td>
         
        <td valign=top align=right style="border-bottom:1px #C4C4C4 dashed; padding:1px 5px 0px 0px;"><%=oRec3("ventetimer_brugt")%> / <%=oRec3("ventetimer") %></td>
        <td valign=top style="border-bottom:1px #C4C4C4 dashed; padding:2px 5px 0px 0px;" class=lillegray><%=oRec3("konto") %>&nbsp;<%=oRec3("kontonr") %><br /><%=oRec3("modkonto") %>&nbsp;<%=oRec3("modkontonr") %></td>
        <td valign=top style="border-bottom:1px #C4C4C4 dashed; padding:3px 5px 0px 0px;">
            <%if oRec3("erfakbetalt") = 1 then
            efbCHK = "SELECTED"
            ikbCHK = ""
            else
            ikbCHK = "SELECTED"
            efbCHK = ""
            end if 
            
            if print <> "j" then%>
            <!--
            <input id="fakid_<%=oRec3("fid")%>" name="fakid_<%=oRec3("fid")%>" value="1" type="radio" <%=efbCHK %> /> <b>Ja</b><br />
            <input id="fakid_<%=oRec3("fid")%>" name="fakid_<%=oRec3("fid")%>" value="0" type="radio" <%=ikbCHK %> /> Nej
            -->
            <select id="fakid_<%=oRec3("fid")%>" name="fakid_<%=oRec3("fid")%>" style="width:50px; font-size:9px;">
                <option value="0" <%=ikbCHK %>>Nej</option>
                <option value="1" <%=efbCHK %>>Ja</option>
            </select>
            <input id="erfakbetalt" name="erfakbetalt" type="hidden" value="<%=oRec3("fid")%>" />
           
            <%else 
                if oRec3("erfakbetalt") = 1 then
                Response.Write "<i>V</i>"
                else
                Response.Write "-"
                end if
            end if
            %>
        </td>
        
        <td valign=top class=lille style="border-bottom:1px #C4C4C4 dashed;">
		<%if print <> "j" then %>
		<textarea id="FM_fakbetkom" name="FM_fakbetkom" cols="20" rows="3" style="width:133px; font-size:9px; font-family:Arial;" ><%=oRec3("fakbetkom") %></textarea>	
            <input id="FM_fakbetkom" name="FM_fakbetkom" value="#" type="hidden" />	
            <%else %>
            <%=oRec3("fakbetkom") %>
            <%end if %>		
		</td>
        
        <td valign=top style="border-bottom:1px #C4C4C4 dashed;">
            <%if cint(oRec3("betalt")) = 0 AND (cdate(oRec3("fakdato")) > cdate("25/8/2007")) AND print <> "j" AND oRec3("fak_laast") = 0 then %>
            <a href="javascript:popUp('erp_fak.asp?func=slet&rdir=hist&id=<%=oRec3("fid")%>', '320', '200','100', '50')" target="_self"><img src="../ill/slet_16.gif" border=0 /></a>
            <%else %>
            &nbsp;
            <%end if %>
            </td>
        </tr>
      
        <% 
        
        'if oRec3("jobid") <> 0 AND oRec3("serviceaft") <> 0 then
        'fakantalIalt = fakantalIalt 
        'else 
        'fakantalIalt = fakantalIalt + (minus&oRec3("fak"))
        'end if
        
        f = f + 1
        eksportFid = eksportFid &","& oRec3("fid")
        
        end if 
        
        oRec3.movenext
        wend
        oRec3.close
	
	
	
	
	if f = 0 then%>
	<tr><td colspan=14>Der blev ikke fundet fakturaer der matcher de valgte kriterier</td></tr>
	</table>
	<%else %>
	<tr><td>
        &nbsp;</td>
        <td colspan=3 style="white-space:nowrap;">&nbsp;</td>
     
        <!--<td align=right style="padding:0px 5px 0px 0px;">&nbsp;</td>-->
        <td colspan=5 align=right style="padding:10px 5px 0px 0px;">
        Antal fakturaer: <b><%=f %></b> stk.<br /><br />
        Faktureret ialt: <B><%=formatnumber((belobGrundValTot+internTotFak+(kreTotFak)), 2) &" DKK"  %></B> <br />
        (kreditnotaer udgør: <%=formatnumber(kreTotFak, 2) & " DKK" %>)<br /><br />
        Interne udgør: <%=formatnumber(internTotFak, 2) & " DKK" %><br />
        Udfaktureret ialt: (- interne) <B><%=formatnumber(belobGrundValTot+(kreTotFak), 2) &" DKK"  %></B> <br /><br />
        
        Job: <%=formatnumber(jobTotFak, 2) & " DKK" %><br />
        Aftaler: <%=formatnumber(aftTotFak, 2) & " DKK" %><br /><br />
        
        Åbne poster: (ikke interne) <b><%=antalAbneposter %></b> stk., <%=formatnumber(febkbelobAbneposter, 2) &" DKK"  %>
        <br /><br />
        Alle beløb er Excl. moms.
        </td>
        
        
        <td colspan=4 align=right valign=top><br />
        <%if print <> "j" then %>
            <input id="Submit1" type="submit" value="Opdater >>" />
            <%end if %>
        &nbsp;</td>
    </tr>
     <input id="FM_fakbetkom" name="FM_fakbetkom" value="xc" type="hidden" />	
    </form>
    </table>
    <br /><br /><br />&nbsp;

    </div>
    
    
    
    
    <%
    
    if print <> "j" then
    
    itop = 80
    ileft = 0
    iwdt = 400
    
    call sideinfo(itop,ileft,iwdt)
    %>
    Interne fakturaer medregnes ikke i omsætningen.
 
   </td>
    </tr>
    </table>
    </div>
    <br /><br /> <br /><br /><br /> <br /><br /><br />
    <br /><br /><br />
    &nbsp;
   <%
        
        
       ptop = 47
       pleft = 805 
       pwdt = 200 
        
       call eksportogprint(ptop,pleft,pwdt)
        
        %>
        
      
       <form action="erp_fakturaer_eksport_2007.asp?visning=0" method="post" target="_blank">
     <input id="Hidden2" name="fakids" value="<%=eksportFid%>" type="hidden" />
     <tr>
    <td valign=top align=center><input type=image src="../ill/export1.png" /></td>
    <td><input id="Submit4" type="submit" value="A) Detail .csv fil eksport" style="font-size:9px; width:160px;" />
    <font class=megetlillesort>Eksporterer alle aktivitets-, medarbejder- og materiale -linjer på de viste fakturaer. (Pivot)
    </td>
    </tr>
    </form>

    
      <form action="erp_fakturaer_eksport_2007.asp?visning=1" method="post" target="_blank">
     <input id="fakids" name="fakids" value="<%=eksportFid%>" type="hidden" />
     <tr>
    <td valign=top align=center><input type=image src="../ill/export1.png" /></td>
    <td><input id="Submit2" type="submit" value="B) Standard .csv fil eksport." style="font-size:9px; width:160px;" />
    <font class=megetlillesort>Eksporterer hovedkolonner, Kontakt, id, adresse, beløb, moms, fakturadato, konti mm.
    </td>
    </tr>
    </form>

        <form action="erp_fakturaer_eksport_2007.asp?visning=2" method="post" target="_blank">
     <input id="Hidden1" name="fakids" value="<%=eksportFid%>" type="hidden" />
     <tr>
    <td valign=top align=center><input type=image src="../ill/export1.png" /></td>
    <td><input id="Submit3" type="submit" value="C) Minimun .txt fil eksport." style="font-size:9px; width:160px;" />
    <font class=megetlillesort>Eksporterer kun kontakt id, beløb incl. moms, fakturadato, valuta.
    </td>
    </tr>
    </form>

   
     <tr>
    <td align=center><a href="erp_fakhist.asp?print=j&FM_job=<%=jobid%>&FM_aftale=<%=aftid%>&FM_sog=<%=showSog%>&FM_viskunabne=<%=viskunabne%>&sort=<%=sort %>&FM_viskun_interne=<%=viskint%>" class=rmenu target=blank><img src="../ill/printer3.png" border=0 alt="" /></a>
    </td><td><a href="erp_fakhist.asp?print=j&FM_job=<%=jobid%>&FM_aftale=<%=aftid%>&FM_sog=<%=showSog%>&FM_viskunabne=<%=viskunabne%>&sort=<%=sort %>&FM_viskun_interne=<%=viskint%>" class=rmenu target=blank>Print faktura historik</a></td>
   </tr>
   
   <tr>
    <td align=center valign=top><a href="#" onclick="Javascript:window.open('erp_make_pdf_multi.asp?fakids=<%=eksportFid%>&lto=<%=lto%>&doctype=pdfxml', '', 'width=350,height=120,resizable=no,scrollbars=no')" class=rmenu><img src="../ill/ikon_sendemail_24.png" border=0 alt="" /></a>
    </td><td><a href="#" onclick="Javascript:window.open('erp_make_pdf_multi.asp?fakids=<%=eksportFid%>&lto=<%=lto%>&doctype=pdfxml', '', 'width=350,height=120,resizable=no,scrollbars=no')" class=rmenu>Mail med PDF og XML filer fra liste.</a>
    <font class=megetlillesort>Afsender mail med vedh. pdf'er og xml filer fra nedenstående liste. (Sendes til egen email, maks. ca 400 fakturaer)
    <br />&nbsp;</td>
   </tr>
   
 
   
     </table>
        </div>
    <%else %>   
     <% 
                Response.Write("<script language=""JavaScript"">window.print();</script>")
                %> 
        
    <%end if%>
	
	
	
	
	<%end if 'f = 0%>
	
	<%


    set fs=nothing
	
	end select 'func %>
                
         
	<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->