<%Response.Buffer = true %>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->
<!--#include file="inc/isint_func.asp"-->
<!--#include file="inc/joblog_inc.asp"-->
<%
if session("user") = "" then
%>
<!--#include file="../inc/regular/header_inc.asp"-->
<%
	errortype = 5
	call showError(errortype)
	else
	
	
	 %>
	
	
	
	<!--#include file="inc/convertDate.asp"-->
	
	<script language="javascript">
	<!--
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	
	function clearJobsog(){
	document.getElementById("FM_jobsog").value = ""
	}
	
	
	
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	
	
	//function checkAll(field)
	//{
	//field.checked = true;
	//for (i = 0; i < field.length; i++)
	//	field[i].checked = true ;
	//}
	
	//function uncheckAll(field){
	//field.checked = false;
	//for (i = 0; i < field.length; i++)
	//	field[i].checked = false ;
	//}
	
	
	function BreakItUp2()
	{
	  //Set the limit for field size.
	  var FormLimit = 302399
	
	  //Get the value of the large input object.
	  var TempVar = new String
	  TempVar = document.theForm2.BigTextArea.value
	  
	  //alert(TempVar.length)
	
	  //If the length of the object is greater than the limit, break it
	  //into multiple objects.
	  if (TempVar.length > FormLimit)
	  {
	    document.theForm2.BigTextArea.value = TempVar.substr(0, FormLimit)
	    TempVar = TempVar.substr(FormLimit)
	
	    while (TempVar.length > 0)
	    {
	      var objTEXTAREA = document.createElement("TEXTAREA")
	      objTEXTAREA.name = "BigTextArea"
	      objTEXTAREA.value = TempVar.substr(0, FormLimit)
	      document.theForm2.appendChild(objTEXTAREA)
	
	      TempVar = TempVar.substr(FormLimit)
	    }
	  }
	}
	
	
	function showjoblogsubdiv(){
	document.getElementById("joblogsub").style.visibility = "visible"
	document.getElementById("joblogsub").style.display = "" 
	}
	
	function hidejoblogsubdiv(){
	document.getElementById("joblogsub").style.visibility = "hidden"
	document.getElementById("joblogsub").style.display = "none" 
	}
	
	function closeprogrpdiv(){
	document.getElementById("progrpdiv").style.visibility = "hidden"
	document.getElementById("progrpdiv").style.display = "none" 
	}
	
	function setshowprogrpval(){
	document.getElementById("showprogrpdiv").value = 1
	}
	
	function curserType(tis){
	document.getElementById(tis).style.cursor='hand'
	}
	
	
	function checkAll(val){
	
	antal_v = document.getElementById("antal_v").value 
	
	if (val == 1) {
	document.getElementById("chkalle1").checked = true
	document.getElementById("chkalle2").checked = false
	tval = true
	} else {
	tval = false
	document.getElementById("chkalle1").checked = false
	document.getElementById("chkalle2").checked = true
	}
	
    for (i = 0; i < antal_v ; i++)
        document.getElementById("FM_godkendt_"+i).checked = tval
	}
	
	//-->
	</script>
	<% 
	
    '**************************************************'
	'***** Faste Filter kriterier *********************'
	'**************************************************'
		
	
		
	'*** Job og Kundeans ***
	call kundeogjobans()
	
	'*** Medarbejdere / projektgrupper '
    selmedarb = session("mid")
	
	call medarbogprogrp("oms")
	
	
	
	'**** Kundekri ***
	if len(request("FM_kunde")) <> 0 then
	kundeid = request("FM_kunde")
	else
	kundeid = 0
	end if
	
	'*** Kundeans ***
	strKnrSQLkri = ""
	
	'***** Valgt job eller søgt på Job ****
	'** hvis Sog = Ja
	call jobsog()

	
	'*** Aftale ****
	if len(request("FM_aftaler")) <> 0 then ' AND jobid <= 0 AND len(jobSogVal) = 0 then
	    aftaleid = request("FM_aftaler")
	else
        aftaleid = 0
	end if
	
	'*** Typer **'
	chktype_1_2 = ""
	chktype_5 = ""
	chktype_11 = ""
	chktype_12 = ""
	chktype_13 = ""
	chktype_14 = ""
	chktype_2021 = ""
	
	if len(request("viskunabnejob")) <> 0 then
	viskunabnejob = request("viskunabnejob")
	    
	    if viskunabnejob = 0 then
	    jost0CHK = "CHECKED"
	    jost1CHK = ""
	    else
	    jost1CHK = "CHECKED"
	    jost0CHK = ""
	    end if
	    
	else
	    if len(trim(request.cookies("stat")("viskunabnejob"))) <> 0 then
	    viskunabnejob = request.cookies("stat")("viskunabnejob")
	    
	            
	            if viskunabnejob = 0 then
                jost0CHK = "CHECKED"
                jost1CHK = ""
                else
                jost1CHK = "CHECKED"
                jost0CHK = ""
                end if
	            
	            
	    else
	    jost0CHK = "CHECKED"
	    jost1CHK = ""
	    viskunabnejob = 0
	    end if
	end if
	
	
	response.cookies("stat")("viskunabnejob") = viskunabnejob
	response.cookies("stat").expires = date + 40
	
	
	
	
	
	if len(request("FM_typer")) <> 0 then
	vartyper = request("FM_typer")
	dim typer
    typerSQL = " tfaktim = -1 "
	typer = split(request("FM_typer"), ",")
	
	for t = 0 to UBOUND(typer)
	typerSQL = typerSQL & " OR tfaktim = "& typer(t) 
	    
	    select case typer(t) 
	    case 1
	    chktype_1_2 = "CHECKED"
	    case 5
	    chktype_5 = "CHECKED"
	    case 6
	    chktype_6 = "CHECKED"
	    case 10
	    chktype_10 = "CHECKED"
	    case 11
	    chktype_11 = "CHECKED"
	    case 12
	    chktype_12 = "CHECKED"
	    case 13
	    chktype_13 = "CHECKED"
	    case 14
	    chktype_14 = "CHECKED"
	    case 20
	    chktype_2021 = "CHECKED"
	    end select
	    
	next 
	else
	    vartyper = "1,2"
	    typerSQL = " tfaktim = 1 OR tfaktim = 2 OR tfaktim = 6 "
	    chktype_1_2 = "CHECKED"
        chktype_5 = ""
        chktype_10 = ""
        chktype_11 = ""
	    chktype_12 = ""
	    chktype_13 = ""
	    chktype_14 = ""
	    chktype_2021 = ""
	end if 
	
	'Response.Write typerSQL
	
	
	'**** Alle SQL kri starter med NUL records ****'
	jobidFakSQLkri = " OR jobid = -1 "
	jobnrSQLkri = " OR tjobnr = -1 "
	jidSQLkri = " OR id = -1 "
	seridFakSQLkri = " OR aftaleid = -1 "
	
	    '**** Skjul timepriser ****'
	     if level <= 2 OR level = 6 then
	    
	        
		    if len(trim(request("joblog_uge"))) <> 0 then
		        if len(trim(request("hidetimepriser"))) <> 0 then    
		        hidetimepriser = 1
		        hidetpCHK = "CHECKED"
		        else
		        hidetimepriser = 0
		        hidetpCHK = ""
		        end if
		    else
		        if len(trim(request.cookies("stat")("hidetimepriser"))) <> 0 then
		        hidetimepriser = request.cookies("stat")("hidetimepriser")
		        if hidetimepriser = 1 then
		        hidetpCHK = "CHECKED"
		        else
		        hidetpCHK = ""
		        end if
		        
		        else
		        hidetimepriser = 0
		        hidetpCHK = ""
		        
		        end if
		        
		    end if
		
		end if
		response.cookies("stat")("hidetimepriser") = hidetimepriser
        '****'
        
    
            '**** Enheder **'
	        if len(trim(request("joblog_uge"))) <> 0 then
		        if len(trim(request("hideenheder"))) <> 0 then    
		        hideenheder = 1
		        hideehCHK = "CHECKED"
		        else
		        hideenheder = 0
		        hideehCHK = ""
		        end if
		    else
		        if len(trim(request.cookies("stat")("hideenheder"))) <> 0 then
		        
		            hideenheder = request.cookies("stat")("hideenheder")
		            if hideenheder = 1 then
		            hideehCHK = "CHECKED"
		            else
		            hideehCHK = ""
		            end if
    		        
		        else
		        hideenheder = 0
		        hideehCHK = ""
		        
		        end if
		        
		    end if
		
		
		    response.cookies("stat")("hideenheder") = hideenheder
	
	        '***'
	        
	        
	        '**** GK status og faktureret **'
	        if len(trim(request("joblog_uge"))) <> 0 then
		        if len(trim(request("hidegkfakstat"))) <> 0 then    
		        hidegkfakstat = 1
		        hidegkCHK = "CHECKED"
		        else
		        hidegkfakstat = 0
		        hidegkCHK = ""
		        end if
		    else
		        if len(trim(request.cookies("stat")("hidegkfakstat"))) <> 0 then
		        
		            hidegkfakstat = request.cookies("stat")("hidegkfakstat")
		            if hidegkfakstat = 1 then
		            hidegkCHK = "CHECKED"
		            else
		            hidegkCHK = ""
		            end if
    		        
		        else
		        hidegkfakstat = 0
		        hidegkCHK = ""
		        
		        end if
		        
		    end if
		
		
		    response.cookies("stat")("hidegkfakstat") = hidegkfakstat
	
	        '***'  
	        
	        
	        
   
	menu = request("menu")
	
	thisfile = "joblog"
	
	func = request("func")
	print = request("print")
	
	rdir = request("rdir")
	
	if rdir = "treg" OR rdir = "nwin" then
	globalWdt = 915
	else
	globalWdt = 1004
	end if
	
	if request("bruguge") = "1" then
	bruguge = 1
	brugugeCHK = "CHECKED"
	    
	    bruguge_week = request("bruguge_week") 
	    bruguge_year = request("bruguge_year") 
	    
	    'Response.Write "bruguge_week: " &  bruguge_week & "<br>"
	    'Response.Write "bruguge_year: " & bruguge_year
	    
	    stDato = "1/1/"&bruguge_year
	    
	    
	    datoFound = 0
	    for u = 1 to 53 AND datoFound = 0
	    
	    if u = 1 then
	    tjkDato = stDato
	    else
	    tjkDato = dateadd("d",7,tjkDato)
	    end if
	    
	    tjkDatoW = datepart("ww", tjkDato, 2,2)
	    
	    if cint(bruguge_week) = cint(tjkDatoW) then
	    
	    wDay = datepart("w", tjkDato, 2,2)
	       
	        
	        select case wDay
	        case 1
	        tjkDato = dateAdd("d", 0, tjkDato)
	        case 2
	        tjkDato = dateAdd("d", -1, tjkDato)
	        case 3
	        tjkDato = dateAdd("d", -2, tjkDato)
	        case 4
	        tjkDato = dateAdd("d", -3, tjkDato)
	        case 5
	        tjkDato = dateAdd("d", -4, tjkDato)
	        case 6
	        tjkDato = dateAdd("d", -5, tjkDato)
	        case 7
	        tjkDato = dateAdd("d", -6, tjkDato)
	        end select
	    
	    stDaguge = day(tjkDato)
	    stMduge = month(tjkDato)
	    stAaruge = year(tjkDato)
	    
	    tjkDato_slut = dateadd("d", 6, tjkDato)
	    
	    slDaguge = day(tjkDato_slut)
	    slMduge = month(tjkDato_slut)
	    slAaruge = year(tjkDato_slut)
	    
	       
	    datoFound = 1
	    
	    end if
	    
	    next
	    
	    
	else
	    
	    if rdir = "treg" AND len(trim(request("bruguge_week"))) = 0 then
	    
	    stMduge = request("FM_start_mrd")
        stDaguge = request("FM_start_dag")
        stAaruge = right(request("FM_start_aar"),2) 
        slDaguge = request("FM_slut_dag")
        slMduge = request("FM_slut_mrd")
        slAaruge = right(request("FM_slut_aar"),2)
        
        bruguge = 1
	    brugugeCHK = "CHECKED"
    	
	    bruguge_week = datepart("ww", stDaguge&"/"&stMduge&"/"&stAaruge, 2,2)
	    bruguge_year = datepart("yyyy", stDaguge&"/"&stMduge&"/"&stAaruge, 2,2) 
        
        else
	    
	    bruguge = 0
	    brugugeCHK = ""
    	
    	if len(trim(request("bruguge_week"))) <> 0 then
    	bruguge_week = request("bruguge_week") 
	    bruguge_year = request("bruguge_year") 
	    else
    	
	    bruguge_week = datepart("ww", now, 2,2)
	    bruguge_year = datepart("yyyy", now, 2,2)
	    end if 
	    
	    end if
	
	end if
	
	'Response.Write "her:" & request("bruguge")
	
	
	'***** Slut faste var **'
	
	
	
	'**** Annuler kunde/overordnet godkender timer *** '
	if func = "opdaterliste" then
	
	  
	
	  ujid = split(request("ids"), ",")
	  uGodkendt = split(trim(request("FM_godkendt")), "#, ")
       

     for u = 0 to UBOUND(ujid)
	
	
	if trim(left(uGodkendt(u), 1)) = "1" then
	uGodkendt(u) = trim(left(uGodkendt(u), 1))
	editor = session("user")
	else
	uGodkendt(u) = 0
	editor = ""
	end if
	
				intGodkendt = request("tid") 
				strSQL = "UPDATE timer SET godkendtstatus = "& uGodkendt(u) &", "_
				&"godkendtstatusaf = '"& editor &"' WHERE tid = " & ujid(u)
				
				'Response.write strSQL &"<br>"
				
				oConn.execute(strSQL)
				
	next
		%>
		
			
	    
	    
	  
		
	<% 
	
	
	strLink = "joblog.asp?rdir="&rdir&"&menu="&menu&"&FM_medarb="&selmedarb&""_
	&"&lastFakdag="&lastFakdag&"&FM_job="&jobid&"&FM_start_dag="&strDag&""_
	&"&FM_start_mrd="&strMrd&"&FM_start_aar="&strAar&"&FM_slut_dag="&strDag_slut&""_
	&"&FM_slut_mrd="&strMrd_slut&"&FM_slut_aar="&strAar_slut&""_
	&"&FM_jobsog="&jobSogValPrint&""_
	&"&FM_aftaler="&aftaleid&""_
	&"&FM_radio_projgrp_medarb="&progrp_medarb&""_
	&"&FM_progrupper="&progrp&""_
	&"&FM_kunde="&kundeid&""_
	&"&FM_kundeans="&kundeans&""_
	&"&FM_jobans="&jobans&""_
	&"&FM_kundejobans_ell_alle="&visKundejobans&""_
	&"&FM_typer="&vartyper&""_
	&"&FM_orderby_medarb="&fordelpamedarb&""_
	&"&bruguge="&bruguge&"&bruguge_week="&bruguge_week&"&bruguge_year="&bruguge_year
	
	'Response.Write strLink
	'Response.end
	Response.redirect strLink
	
	
	end if
	
	
	

	
	
	if print <> "j" AND rdir <> "treg" AND rdir <> "nwin" then
	pleft = 20
	ptop = 132
	jobbgcol = "#ffffff"
	else
	pleft = 20
	ptop = 0
	jobbgcol = "#ffffff"
	end if
	
	if print <> "j" AND rdir <> "treg" AND rdir <> "nwin" then
	%>
	
    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	
	<%select case level
	case 1,2,6 %>
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
	<%end select%>
	
	<%
	else
	
	level = session("rettigheder") '*** Sættes normalt i topmenu **'
	
	    if print = "j" then
	    %>
	    <!--#include file="../inc/regular/header_hvd_inc.asp"-->
	    <% 
	    else
	    %>
	    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	    <% 
	    
	    end if
	end if
	
	
	
	'*** Var. til komrpimeret visning ***'
        if len(request("FM_komprimeret")) <> 0 then
            komprimer = request("FM_komprimeret")
            if komprimer = "1" then 
            komCHK1 = "CHECKED"
            komCHK0 = ""
            else
            komCHK1 = ""
            komCHK0 = "CHECKED"
            end if
            Response.Cookies("stat")("komprimerliste") = komprimer
        else
            if request.Cookies("stat")("komprimerliste") = "1" then
             komCHK1 = "CHECKED"
             komCHK0 = ""
             komprimer = 1
            else
            komprimer = 0
            komCHK1 = ""
            komCHK0 = "CHECKED"
            end if
        end if
        
	    
	    '**** Vis ugeseddel ell. joblog ***'
	    if rdir = "treg" then
	        joblog_uge = 2 '** Viser ugeseddel **'
	    else
	    
	        if len(request("joblog_uge")) <> 0 then
			joblog_uge = request("joblog_uge")
			    
			    
			else
			        if request.Cookies("stat")("joblog_uge") <> "" then
			        joblog_uge = request.Cookies("stat")("joblog_uge")
			        else
			        joblog_uge = 1
			        end if
			
			end if
		
		
			
			
			    if cint(joblog_uge) = 1 then
			    joblog_ugeCHK1 = "CHECKED"
			    joblog_ugeCHK2 = ""
			    
			    joblogsubWZB = "visible"
			    joblogsubDSP = ""
			    
			    else
			    joblog_ugeCHK1 = ""
			    joblog_ugeCHK2 = "CHECKED"
			    
			    joblogsubWZB = "hidden"
			    joblogsubDSP = "none"
			    
			    end if
			
			response.Cookies("stat")("joblog_uge") = joblog_uge
	
	    end if 
	
	
	
	'**** Vis subtotaler pr akt ell. medarb. ***'
	if cint(joblog_uge) = 1 then		
	
	if len(trim(request("FM_orderby_medarb"))) <> 0 AND request("FM_orderby_medarb") <> "0" then
	fordelpamedarb = request("FM_orderby_medarb")
	    
	    if fordelpamedarb = 1 then
	    orderByKri = "Tjobnavn, Tjobnr, Tmnavn"
	    else
	    orderByKri = "Tjobnavn, Tjobnr, TAktivitetId"
	    end if
	
	else
	fordelpamedarb = 0
	orderByKri = "Tjobnavn, Tjobnr"
	end if
	
	
	orderByKri = orderByKri & ", Tdato DESC"
	
	else
	
	'** Timesedde altide pr. medarb, tdato ***'
	orderByKri = orderByKri & " Tmnr, Tdato DESC"
	komprimer = 1
	
	end if
	
	
	 
	%>
    
    
    





<div id="sindhold" style="position:absolute; left:<%=pleft%>; top:<%=ptop%>; visibility:visible; z-index:100;">
<!-------------------------------Sideindhold------------------------------------->


<%if print <> "j" then%>



	<% 
    
    
    oimg = "ikon_stat_joblog_48.png"
	oleft = 0
	otop = 0
	owdt = 300
	oskrift = "Joblog & Ugesedler"
	
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	
	
	if rdir = "treg" then
	%>
	<a href="timereg_akt_2006.asp?showakt=1" class="vmenu"><< Tilbage</a><br /><br />
	<%
	else
	%>
	<br /><br />
	<%
	end if
	
    	    
   
        call filterheader(0,0,700,pTxt)%>
    	    
	  
	    <table cellspacing=0 cellpadding=5 border=0 width=100% bgcolor="#FFFFFF">
	    <form action="joblog.asp?rdir=<%=rdir %>" method="post">
	    
	    <!--FM_jobsog=<%=jobSogValPrint%>&FM_job=<%=jobid%>&FM_aftaler=<%=aftaleid%>&seomsfor=<%=strYear%>" >
	    
	    <!--<input type="hidden" name="FM_kunde" id="FM_kunde" value="<%=kundeid%>">
	    <input type="hidden" name="FM_kundeans" id="FM_kundeans" value="<%=kundeans%>">
	    <input type="hidden" name="FM_jobans" id="FM_jobans" value="<%=jobans%>">
	    <input type="hidden" name="FM_kundejobans_ell_alle" value="<%=visKundejobans%>">
	    <input type="hidden" name="FM_typer" value="<%=vartyper%>">
	    <input type="hidden" name="FM_orderby_medarb" value="<%=fordelpamedarb%>">
	    -->
    	
    	
	    <tr>
	    <td><input type="radio" name="FM_radio_projgrp_medarb" value="1" <%=projgrp_medarb1%>>&nbsp;<b>Medarbejder(e):</b><br>
	    <select name="FM_medarb" style="font-size : 11px; width:205px;">
	    
	    <%
	    select case level
	    case 1,2,6
	    %>
	    <option value="0">Alle</option>
	    <%
	    medarbSelKri = " mansat <> 2 "
	    case else
	    medarbSelKri = " mid = " & selmedarb
	    end select
	    
	    strSQL = "SELECT Mid, Mnavn, Mnr, mansat, init FROM medarbejdere WHERE "& medarbSelKri &" ORDER BY Mnavn"
	    oRec.open strSQL, oConn, 3
	    while not oRec.EOF 
    	
    		
		    if cint(selmedarb) = cint(oRec("mid")) then
		    thisChecked = "SELECTED"
		    else
		    thisChecked = ""
		    end if
		    %>
		    <option value="<%=oRec("mid")%>" <%=thisChecked%>><%=oRec("mnavn")%>&nbsp;(<%=oRec("mnr")%>) - <%=oRec("init")%></option>
	    <%
	    oRec.movenext
	    wend
	    oRec.close
	    %></select>
	    </td>
    	
	    <td>
        &nbsp;
	    
	    <% 
	    if rdir <> "treg" then
	    
	    
	    select case level
	    case 1,2,6
	    %>
	    
	    
	    <input type="radio" name="FM_radio_projgrp_medarb" value="2" <%=projgrp_medarb2%> onclick="setshowprogrpval();">&nbsp;<b>Projektgruppe(r):</b><br>
	    <%
	    strSQL = "SELECT p.id AS id, navn, count(medarbejderid) AS antal FROM projektgrupper p "_
	    &" LEFT JOIN progrupperelationer ON (ProjektgruppeId = p.id) "_
	    &" GROUP BY p.id ORDER BY navn"
	    oRec.open strSQL, oConn, 3
    	
	    'Response.write strSQL
	    'Response.flush
	    %>
    	
	    <select name="FM_progrupper" style="font-size : 11px; width:205px;" onchange="setshowprogrpval();">
	    <%
    			
    			
			    while not oRec.EOF
    			
			    if cint(progrp) = cint(oRec("id")) then
			    isSelected = "SELECTED"
			    else
			    isSelected = ""
			    end if
    			
			    if cint(oRec("antal")) > 0 then%>
			    <option value="<%=oRec("id")%>" <%=isSelected%>><%=oRec("navn")%></option>
			    <%end if
    			
			    oRec.movenext
			    wend
			    oRec.close
			    %>
	    </select>
	    
	    <%
	    end select
	    
	    end if
	    
	    %>
	    &nbsp;
	    </td>
	    <td valign=middle align=right>&nbsp;<input type="submit" value="Vælg medarb."></td>
	    </tr>
	    
	    <!--
	    </form>
	    -->
	    </table>

	    <%if progrp_medarb = 2 ANd request("showprogrpdiv") = "1" then %>
	    <div id="progrpdiv" style="position:absolute; left:625px; top:200px; width:200px; height:250px; overflow:auto; background-color:#ffffff; border:1px #c4c4c4 solid; padding:10px; font-size:9px; z-index:20000;" onclick="closeprogrpdiv();" onmouseover="curserType('progrpdiv');">
	    <b>Medarbejdere i valgt projektgruppe:</b><br />
	    <%=medarbigrp %><br />
	    [Klik her for at lukke]
	    </div>
	    <%end if %>
	    
        <input id="showprogrpdiv" name="showprogrpdiv" value="0" type="hidden" />
<%end if %>
	
	
	
<% if print <> "j" AND rdir <> "treg" then%>
	
	<table cellspacing=0 cellpadding=5 border=0 width=100% bgcolor="#FFFFFF">
	
	<!--<form action="joblog.asp?FM_job=0&FM_aftaler=0&seomsfor=<%=strYear%>" method="post">
	<input type="hidden" name="FM_radio_projgrp_medarb" id="FM_radio_projgrp_medarb" value="<%=progrp_medarb%>">
		<input type="hidden" name="FM_progrupper" id="FM_progrupper" value="<%=progrp%>">
		<input type="hidden" name="FM_medarb" id="FM_medarb" value="<%=selmedarb%>">
		<input type="hidden" name="FM_typer" value="<%=vartyper%>">
		<input type="hidden" name="FM_orderby_medarb" value="<%=fordelpamedarb%>">
		-->
	
	</tr>
	<tr bgcolor="#Eff3ff">
		<td colspan=2>
		<input type="radio" name="FM_kundejobans_ell_alle" value="1" <%=kundejobansCHK1%>><b>A) For Key account.</b> Vis kun job hvor..<br>
		<input type="checkbox" name="FM_kundeans" id="FM_kundeans" value="1" <%=kundeansChk%>>&nbsp;Valgte ovenstående medarbejder(e) er kundeansvarlig.<br>
		<input type="checkbox" name="FM_jobans" id="FM_jobans" value="1" <%=jobansChk%>>&nbsp;Valgte ovenstående medarbejder(e) er jobansvarlig.
		</td>
	</tr>
	<tr bgcolor="#Eff3ff">
		<td><input type="radio" name="FM_kundejobans_ell_alle" value="0" <%=kundejobansCHK0%>><b>B) Kontakt(er):</b>
<%end if
		
		
		
		
		strKnrSQLkri = " OR jobknr = -1 "
		strAftKidSQLkri = " OR kundeid = -1 "
		
		strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE (" & kundeAnsSQLKri & ") AND ketype <> 'e' ORDER BY Kkundenavn"
		'Response.write strSQL & "<br>"
		'Response.flush		
		
		
if print <> "j" AND rdir <> "treg" then%>
	<select name="FM_kunde" id="FM_kunde" style="font-size : 11px; width:305px;">
<%end if
				
				oRec.open strSQL, oConn, 3
				k = 0
				while not oRec.EOF
				
				if k = 0 then
				
				    if print <> "j" AND rdir <> "treg" then
				    %>
				    <option value="0">Alle</option>
				    <%
				    end if
				    
				end if
				
				if cint(kundeans) = 1 then
				strKnrSQLkri = strKnrSQLkri & " OR jobknr = "& oRec("kid")
				strAftKidSQLkri = strAftKidSQLkri & " OR kundeid = " & oRec("kid")
				else
				strAftKidSQLkri =  " OR kundeid <> 0 "
				strKnrSQLkri = " OR jobknr <> 0 "
				end if
				
				
				if cint(kundeid) = cint(oRec("kid")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if
				
				
				if print <> "j" AND rdir <> "treg" then
				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%>&nbsp;(<%=oRec("Kkundenr")%>)</option>
				<% 
				end if
				
				k = k + 1
				oRec.movenext
				wend
				oRec.close
				
		   
		   
		   if print <> "j" AND rdir <> "treg" then
				
				if cint(kundeid) = -1 then
				ingChk = "SELECTED"
				else
				ingChk = ""
				end if%>
				<option value="-1" <%=ingChk%>>Ingen</option>
				
				
				
		    </select>
		    <%end if 
		
		
				
				len_strKnrSQLkri = len(strKnrSQLkri)
				right_strKnrSQLkri = right(strKnrSQLkri, len_strKnrSQLkri - 3)
				strKnrSQLkri = right_strKnrSQLkri
				
				len_strAftKidSQLkri = len(strAftKidSQLkri)
				right_strAftKidSQLkri = right(strAftKidSQLkri, len_strAftKidSQLkri - 3)
				strAftKidSQLkri = right_strAftKidSQLkri
				
				%>
				</td><td align=right>
		<%if print <> "j" AND rdir <> "treg" then%>
		<input type="submit" value="Vælg Kontakt(er)">
		<%else %>
		&nbsp;
		<%end if%>
		</td>
		
		<!--
		</form>
		-->
	
	</tr>
	</table><br>
	
	<%
	
	
	    '*** strKnrSQLkri Bruges også i job functionen ***'
		if cint(kundeid) = 0 then
		strKnrSQLkri = strKnrSQLkri '" jobknr = 0" '
		else
		strKnrSQLkri = " jobknr = "& kundeid
		end if
		
		'*** Er jobansvarlig tilvalgt ? ** 
		if cint(jobans) = 1 then
		strKnrSQLkri = "("&strKnrSQLkri &") "& jobAnsSQLkri
		end if
		
		
		if aftaleid > 0 then
		aftKri = " AND serviceaft = "& aftaleid
		else
		aftKri = ""
		end if
		
		if len(trim(jobSogVal)) <> 0 then
		sogKri = " AND (jobnavn LIKE '" & jobSogVal &"%' OR jobnr = '"& jobSogVal &"')"
		else
		sogKri = ""
		end if
		
	
	    if viskunabnejob <> 0 then
	    jobstKri = " AND j.jobstatus = 1 "
	    else
	    jobstKri = " AND j.jobstatus <> 3 "
	    end if
	    
	    
	    if cint(kundeid) = 0 then
	    strAftKidSQLkri = strAftKidSQLkri '" kundeid <> 0" '
		else
		strAftKidSQLkri = " kundeid = "& kundeid
		end if
	    
	
	
	if print <> "j" then
	%>
	<table cellspacing=0 cellpadding=5 border=0 width=100% bgcolor="#FFFFFF">
	
	<%if rdir <> "treg" then %>
	
	<tr>
	
	<!--
	<form action="joblog.asp" method="post">
		<input type="hidden" name="FM_radio_projgrp_medarb" id="FM_radio_projgrp_medarb" value="<%=progrp_medarb%>">
		<input type="hidden" name="FM_progrupper" id="FM_progrupper" value="<%=progrp%>">
		<input type="hidden" name="FM_medarb" id="FM_medarb" value="<%=selmedarb%>">
		<input type="hidden" name="FM_kunde" id="FM_kunde" value="<%=kundeid%>">
		<input type="hidden" name="FM_kundeans" id="FM_kundeans" value="<%=kundeans%>">
		<input type="hidden" name="FM_jobans" id="FM_jobans" value="<%=jobans%>">
		<input type="hidden" name="FM_kundejobans_ell_alle" value="<%=visKundejobans%>">
		-->
		
	    
		<td colspan=3><b>A) Aftaler:</b> (Og tilhørende job)&nbsp;
		<%
			
		strSQL = "SELECT s.navn, s.aftalenr, s.id, k.kkundenavn, k.kkundenr FROM serviceaft s "_
		&" LEFT JOIN kunder k ON (k.kid = s.kundeid) "_
		&" WHERE "& strAftKidSQLkri &" ORDER BY k.kkundenavn, s.navn"
		
		'Response.write strSQL
		%>
		
		<select name="FM_aftaler" id="FM_aftaler" style="font-size : 11px; width:383px;" onChange="clearJobsog();">
		<option value="0">Alle - eller vælg fra liste...</option>
		<%
		
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(aftaleid) = cint(oRec("id")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if
				%>
				<option value="<%=oRec("id")%>" <%=isSelected%>><%=oRec("id")%> - <%=oRec("kkundenavn")%>&nbsp;(<%=oRec("kkundenr")%>) ][ <%=oRec("navn")%>&nbsp;(<%=oRec("aftalenr")%>)</option>
				<%
				oRec.movenext
				wend
				oRec.close
				
				if cint(aftaleid) = -1 then
				aChk = "SELECTED"
				else
				aChk = ""
				end if%>
				<option value="-1" <%=aChk%>>Ingen</option>
		</select></td>
	</tr>
	    
	
	<tr>
	<td colspan=3>
	
		<br />
		<b>B) Søg på jobnavn ell. jobnr.:</b>&nbsp;&nbsp;<input type="text" name="FM_jobsog" id="FM_jobsog" value="<%=jobSogVal%>" style="font-size:11px; width:200px;">&nbsp;(% wildcard, ell. <b>231, 269</b> for flere job)
		
		<br><br />
		<input name="viskunabnejob" id="viskunabnejob" type="radio" value=0 <%=jost0CHK %> onClick="submit()" />Vis alle job <input name="viskunabnejob" id="Radio3" type="radio" value=1 <%=jost1CHK %> onClick="submit()"/>Vis kun åbne job. (kør stat. igen)
		<br /><b>C) Vælg job:</b>&nbsp;
		<%
		
		
		strSQL = "SELECT jobnavn, jobnr, j.id, k.kkundenavn, k.kkundenr, j.serviceaft,"_
		&" s.navn, s.aftalenr FROM job j "_
		&" LEFT JOIN kunder k ON (k.kid = j.jobknr) "_
		&" LEFT JOIN serviceaft s ON (s.id = j.serviceaft) "_
		&" WHERE "& strKnrSQLkri &" "& jobstKri &" "& aftKri &" "& sogKri &" ORDER BY k.kkundenavn, j.jobnavn"
		
		'3 = tilbud
		'Response.write strSQL
		'Response.flush
		%>
        
		
		<select name="FM_job" id="FM_job" style="font-size : 11px; width:466px;" onChange="clearJobsog(), submit();">
		<option value="0">Alle</option>
		<%
				
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(jobid) = cint(oRec("id")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if
				
				if oRec("serviceaft") <> 0 then
				aftnavnognr = " -- aftale: "& oRec("navn") & "("& oRec("aftalenr") &")"
				else
				aftnavnognr = ""
				end if
				
				%>
				<option value="<%=oRec("id")%>" <%=isSelected%>><%=oRec("kkundenavn")%>&nbsp;(<%=oRec("kkundenr")%>) ][ <%=oRec("jobnavn")%>&nbsp;(<%=oRec("jobnr")%>) <%=aftnavnognr %></option>
				<%
				oRec.movenext
				wend
				oRec.close
				
				if cint(jobid) = -1 then
				mChk = "SELECTED"
				else
				mChk = ""
				end if%>
				<option value="-1" <%=mChk%>>Ingen</option>
		</select>
        </td>
	</tr>
	
	<%end if 'rdir %>
	
	<tr bgcolor="#Eff3ff">
			<!-- Brug altid datointerval, FM_usedatokri = 1 -->
			<input type="hidden" name="FM_usedatokri" id="FM_usedatokri" value="1">
			<td valign=top style="width:300px;">
			<br /><b>Vis periode ell. uge:</b><br>
			<!--#include file="inc/weekselector_s.asp"--> <!-- b -->
			
			<br />
			<br />
			
            <input id="bruguge" name="bruguge" type="checkbox" value="1" <%=brugugeCHK %> /> Vis uge: 
			<select name="bruguge_week">
			<% for w = 1 to 52
			
			if w = 1 then
			useWeek = 1
			else
			useWeek = dateadd("ww", 1, useWeek)
			end if
			
			if cint(datepart("ww", useWeek, 2,2)) = cint(bruguge_week) then
			wSel = "SELECTED"
			else
			wSel = ""
			end if
			
			%>
			<option value="<%=datepart("ww", useWeek, 2,2) %>" <%=wSel%>><%=datepart("ww", useWeek, 2,2) %></option>
			<%
			next %>
			
			</select>
			
			
			<select name="bruguge_year">
			<%for y = 1 to 10
			
			if y = 1 then
			useYear = dateadd("yyyy", -5, now)
			else
			useYear = dateadd("yyyy", 1, useYear)
			end if 
			
			if cint(datepart("yyyy", useYear, 2,2)) = cint(bruguge_year) then
			ySel = "SELECTED"
			else
			ySel = ""
			end if
			
			%>
			<option value="<%=datepart("yyyy", useYear, 2,2) %>" <%=ySel %>><%=datepart("yyyy", useYear, 2,2) %></option>
			<%
			
			next %>
			
			</select><br /><br />
			
			
		<%if rdir <> "treg" then %>
			
        <input id="joblog_uge1" name="joblog_uge" type="radio" value="1" <%=joblog_ugeCHK1 %> onclick=showjoblogsubdiv() /> <b>Vis som joblog</b> (en lang liste)<br />
	    
	    <div id="joblogsub" style="position:relative; background-color:#ffffff; top:0px; left:20px; width:175px; border:1px #cccccc solid; padding:10px; visibility:<%=joblogsubWZB %>; display:<%=joblogsubDSP %>;">
	    Vis Subtotaler:<br />
        
        <%if fordelpamedarb = 1 then
		chkFordelmedarb = "CHECKED"
		else
		chkFordelmedarb = ""
		end if%>
		
		<input type="radio" name="FM_orderby_medarb" id="FM_orderby_medarb" value="1" <%=chkFordelmedarb%>>Fordel på medarbejdere<br>
	    
	    <%if fordelpamedarb = 2 then
		chkFordelpaakt = "CHECKED"
		else
		chkFordelpaakt = ""
		end if%>
		
		<input type="radio" name="FM_orderby_medarb" id="Radio1" value="2" <%=chkFordelpaakt%>>Fordel på aktiviteter<br>
	    
	    <%if fordelpamedarb = 0 then
		chkFordelingen = "CHECKED"
		else
		chkFordelingen= ""
		end if%>
		
		<input type="radio" name="FM_orderby_medarb" id="Radio2" value="0" <%=chkFordelingen%>>Ingen<br>
	    <br />
	    
        
        
        Vis Komprimeret visning<br />
                <input name="FM_komprimeret" id="FM_komprimeret" type="radio" value="1" <%=komCHK1 %> /> Ja 
			    <input name="FM_komprimeret" id="FM_komprimeret" type="radio" value="0" <%=komCHK0 %> /> Nej
	    </div>
	    
	   
	   <input id="joblog_uge2" name="joblog_uge" type="radio" value="2" <%=joblog_ugeCHK2 %> onclick=hidejoblogsubdiv() /> <b>Vis som ugesedler for medarbejdere</b><br />(liste opdelt på ugebasis pr. medarb.)<br /><br />
		
		
			
		<%
		end if '** rdir **' 
		
	    %>
		
		
			
			<input id="hidetimepriser" name="hidetimepriser" value="1" type="checkbox" <%=hidetpCHK %> /> Skjul Timepriser
		<input id="hideenheder" name="hideenheder" value="1" type="checkbox" <%=hideehCHK %> /> Skjul Enheder
		
		<br />
		<input id="hidegkfakstat" name="hidegkfakstat" value="1" type="checkbox" <%=hidegkCHK %> /> Skjul Godkendt- og Faktureret -status
		
		
		
			</td>
			<td style="width:300px;" valign=top>
			
			<br><br><b>Vis aktivitets type(r):</b>
	        <br />
        <input name="FM_typer" id="type" value="1,2" type="checkbox" <%=chktype_1_2 %> /> <%=global_txt_142 & " ("& global_txt_129 &" & "& global_txt_131 & ")" %><br />
        <input name="FM_typer" id="type" value="5" type="checkbox" <%=chktype_5 %> /> <%=global_txt_130 %> <br />
        <input name="FM_typer" id="type" value="6" type="checkbox" <%=chktype_6 %> /> <%=replace(global_txt_132, "|", "&") %> (<%=global_txt_131 %>) <br />
        <input name="FM_typer" id="type" value="10" type="checkbox" <%=chktype_10 %> /> <%=global_txt_133 %>*<br />
        <input name="FM_typer" id="type" value="11" type="checkbox" <%=chktype_11 %> /> <%=global_txt_134 %>* <br />
        <input name="FM_typer" id="type" value="14" type="checkbox" <%=chktype_14 %> /> <%=global_txt_135 %> <br />
        <input name="FM_typer" id="type" value="12" type="checkbox" <%=chktype_12 %> /> <%=global_txt_136 %>* <br />
        <input name="FM_typer" id="type" value="13" type="checkbox" <%=chktype_13 %> /> <%=global_txt_137 %> <br />
        <input name="FM_typer" id="type" value="20,21" type="checkbox" <%=chktype_2021 %> /> <%=global_txt_138 &" & "& global_txt_139 %> <br />
       
       </td><td align=right valign=bottom style="width:100px;">
       
       
	<input type="submit" value="Kør statistik.">
	</td>
	</tr>
	
	
	</form>
	</table>
	
	<!-- filterDiv -->
	</td></tr>
	</table>
	</div>
	<%else 
	dontshowDD = 1%>
	<!--#include file="inc/weekselector_s.asp"--> 
	<%
	level = session("rettigheder")
	end if 'print%>
	




<%
startDatoKriSQL = strAar &"/"& strMrd &"/"& strDag
slutDatoKriSQL = strAar_slut &"/"& strMrd_slut &"/"& strDag_slut
	
	
	
	'*****************************************************************************************************
		'**** Job der skal indgå i omsætning, budget og forbrugstal *******************************
		'*****************************************************************************************************
		
		
		'*** valgte job ***
		call valgtejob()
		
		
		'**** Valgte aftaler *****
		call valgteaftaler()
				
				
				
				'*** For at spare (trimme) på SQL hvis der vælges alle job alle kunder og vis kun for jobanssvarlige ikke er slået til ****
				'*** Og der ikke er søgt på jobnavn ***
				if cint(kundeid) = 0 AND cint(jobid) = 0 AND cint(jobans) = 0 AND cint(kundeans) = 0 AND len(trim(jobSogVal)) = 0 AND cint(aftaleid) = 0  then 
						
					'jobidFakSQLkri = " OR jobid <> 0 "
					jobnrSQLkri = " OR tjobnr <> 0 "
					jidSQLkri =  " OR id <> 0 "
					seridFakSQLkri = " OR aftaleid <> 0 "
						
				end if
	
	
		'**************** Trimmer SQL states ************************
		'len_jobidFakSQLkri = len(jobidFakSQLkri)
		'right_jobidFakSQLkri = right(jobidFakSQLkri, len_jobidFakSQLkri - 3)
		'jobidFakSQLkri =  right_jobidFakSQLkri
		
		len_jobnrSQLkri = len(jobnrSQLkri)
		right_jobnrSQLkri = right(jobnrSQLkri, len_jobnrSQLkri - 3)
		jobnrSQLkri =  right_jobnrSQLkri
		
		len_jidSQLkri = len(jidSQLkri)
		right_jidSQLkri = right(jidSQLkri, len_jidSQLkri - 3)
		jidSQLkri =  right_jidSQLkri
		
		
		len_seridFakSQLkri = len(seridFakSQLkri)
		right_seridFakSQLkri = right(seridFakSQLkri, len_seridFakSQLkri - 3)
		seridFakSQLkri =  right_seridFakSQLkri
		
	
		

	
	if request("print") = "j" then
	%>
	Periode <%=formatdatetime(strDag&"/"&strMrd&"/"&strAar, 1)%> - <%=formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 1)%> </font><br>
	<%else%>
	<!--include file="inc/stat_submenu.asp"-->
	<%end if%>



        <%if cint(hidegkfakstat) <> 1 then %>
        <form action="joblog.asp?func=opdaterliste&menu=<%=menu %>&rdir=<%=rdir %>" method="post">
	    <!--<input type="hidden" name="seomsfor" id="hidden10" value="<%=strYear%>">-->
	    <input type="hidden" name="FM_aftaler" id="hidden9" value="<%=aftaleid%>">
	   <input type="hidden" name="FM_job" id="hidden8" value="<%=jobid%>">
	    <input type="hidden" name="FM_jobsog" id="hidden7" value="<%=jobSogVal%>">
	   <input type="hidden" name="FM_radio_projgrp_medarb" id="hidden4" value="<%=progrp_medarb%>">
		<input type="hidden" name="FM_progrupper" id="hidden5" value="<%=progrp%>">
		<input type="hidden" name="FM_medarb" id="hidden6" value="<%=selmedarb%>">
	    <input type="hidden" name="FM_kunde" id="hidden1" value="<%=kundeid%>">
	    <input type="hidden" name="FM_kundeans" id="hidden2" value="<%=kundeans%>">
	    <input type="hidden" name="FM_jobans" id="hidden3" value="<%=jobans%>">
	   <input type="hidden" name="FM_kundejobans_ell_alle" value="<%=visKundejobans%>">
	    <input type="hidden" name="FM_typer" value="<%=vartyper%>">
	    <input type="hidden" name="FM_orderby_medarb" value="<%=fordelpamedarb%>">
	    <input type="hidden" name="bruguge" value="<%=bruguge%>">
	    <input type="hidden" name="bruguge_week" value="<%=bruguge_week%>">
	    <input type="hidden" name="bruguge_year" value="<%=bruguge_year%>">
	    
	   <%end if %>
	    
      
       


        <%
        
        
        visopdaterknap = 0

		'*** Smiley ***''
		call ersmileyaktiv()
	    
	   
	    '*** Afsluttede uger ***''
		dim medabAflugeId
		redim medabAflugeId(200)
		strSQL2 = "SELECT u.status, u.afsluttet, WEEK(u.uge, 1) AS ugenr, YEAR(u.uge) AS aar, u.id, u.mid FROM ugestatus u WHERE "& ugeAflsMidKri &" AND uge BETWEEN '"& startDatoKriSQL &"' AND '"& slutDatoKriSQL &"' GROUP BY u.mid, uge"
		'Response.write strSQL2
		'Response.end
		oRec2.open strSQL2, oConn, 3
		while not oRec2.EOF
			medabAflugeId(oRec2("mid")) = medabAflugeId(oRec2("mid")) & "#"& oRec2("ugenr") &"_"& oRec2("aar") &"#,"
		oRec2.movenext
		wend
		oRec2.close 
		
	
	'*******************'
	'**** MaIN SQL *****'
	'*******************'
	
	
	lastjobnavn = ""
	lastmedarb = 0
	strSQL = "SELECT Tdato, TasteDato, Tjobnr, Tjobnavn, TAktivitetNavn AS Anavn, TAktivitetId, "_
	&" a.fakturerbar,"_
	&" Tknavn, Tmnr, Tmnavn, Timer, Tid, Tfaktim, TimePris, Timerkom, Tknr,"_
	&" sttid, sltid, a.faktor, godkendtstatus, godkendtstatusaf, kkundenr, "_
	&" m.mnr, m.init, mnavn, m.mid, a.id AS aid, v.valutakode, t.valuta, k.kid"_
	&" FROM timer t "_
	&" LEFT JOIN aktiviteter a ON (a.id = TAktivitetId)"_
	&" LEFT JOIN kunder k ON (k.kid = Tknr)"_
	&" LEFT JOIN medarbejdere m ON (m.mid = tmnr)"_
	&" LEFT JOIN valutaer v ON (v.id = t.valuta)"_
	&" WHERE ("& jobnrSQLkri &") AND ("& medarbSQlKri &") AND ("& typerSQL &") AND "_
	&" (Tdato BETWEEN '"& startDatoKriSQL &"' AND '"& slutDatoKriSQL &"') ORDER BY "& orderByKri
	
	'&" WHERE "& jobMedarbKri &" "& selaktidKri &""_
    'Response.write "kundeid:"& kundeid &"  "&  strSQL & "<br><br>"
	'Response.flush
	
	oRec.open strSQL, oConn, 0, 1
	v = 0
	while not oRec.EOF
		
		
		strWeekNum = datepart("ww", oRec("Tdato"),2,2)
		id = 1
		
			
			if lto = "bowe" then
			
			
					
					'*** Akt navn (Dag / Aften / Nat) ****
					
					akttid = ""
					if instr(oRec("Anavn"), "(Nat)") <> 0 OR instr(oRec("Anavn"), "(Aften)") <> 0 OR instr(oRec("Anavn"), "(Dag)") <> 0 then 
					
					if instr(oRec("Anavn"), "(Nat)") <> 0 then
					akttid_st = instr(oRec("Anavn"), "(N")
					akttid_sl = instr(oRec("Anavn"), "t)")
					end if
					
					if instr(oRec("Anavn"), "(Aften)") <> 0 then
					akttid_st = instr(oRec("Anavn"), "(A")
					akttid_sl = instr(oRec("Anavn"), "n)")
					end if
					
					if instr(oRec("Anavn"), "(Dag)") <> 0 then
					akttid_st = instr(oRec("Anavn"), "(D")
					akttid_sl = instr(oRec("Anavn"), "g)")
					end if
					
					if akttid_st <> 0 AND akttid_sl <> 0 then
					akttid_length = (akttid_sl - akttid_st) 
					akttid_str = mid(oRec("Anavn"), akttid_st+1, akttid_length)
					akttid = akttid_str
					else
					akttid = ""
					end if
					
					
			        end if
					
					
					
					'*** KA nr ****
					kanr_st = instr(oRec("Tjobnavn"), "(")
					kanr_sl = instr(oRec("Tjobnavn"), ")")
					
					if kanr_st <> 0 AND kanr_sl <> 0 then
					kanr_length = (kanr_sl - kanr_st) 
					kanr_str = mid(oRec("Tjobnavn"), kanr_st+1, kanr_length-1)
					kanr = kanr_str
					else
					kanr = ""
					end if
					
					'** bowekode **
					bowekode_st = instr(oRec("Tjobnavn"), ")")
					
					if bowekode_st <> 0 then
					bowekode_length = 2
					bowekode = mid(oRec("Tjobnavn"), bowekode_st+1, 3)
					else
					bowekode = ""
					end if
					
					ekspTxt = ekspTxt & oRec("Tknavn")&";"&oRec("kkundenr")&";"&oRec("Tjobnavn") &";"&kanr&";"&bowekode&";"& oRec("Tjobnr")&";"
			
			else
		
			ekspTxt = ekspTxt & oRec("Tknavn")&";"&oRec("kkundenr")&";"&oRec("Tjobnavn") &";"& oRec("Tjobnr")&";"
		
			end if 'bowe
			
			
			
			if cint(joblog_uge) = 1 then
			
			
		
		        
		        'Response.write lastmedarb &" <> " & oRec("Tmnr") & " job: " & lastjobnavn &" <> " & oRec("Tjobnavn") 
		        'Response.Write lastaktid & "<>"& oRec("Taktivitetid")
				
		        if fordelpamedarb = 1 then
				
				    if (lastmedarb <> oRec("Tmnr") OR lastjobnr <> oRec("Tjobnr")) then
					    if v <> 0 then
    					
					    call fordelpamedarbejdere()
    					
					    medarbFeriePlan = 0
					    medarbFrokot = 0
					    medarbAfspadOp = 0
					    medarbKm = 0
					    medarbtimer = 0
					    medarbEnheder = 0
    					
					    end if
				    end if
				
				end if
				
				
				
				if fordelpamedarb = 2 then
				
				    if (lastaktid <> oRec("Taktivitetid") OR lastjobnr <> oRec("Tjobnr")) then
					    if v <> 0 then
    					
					    call fordelpaakt()
    					
    					
					    akttimer = 0
					    aktEnheder = 0
    					
    					
					    end if
				    end if
				
				end if
		        
		        
				
				if lastjobnr <> oRec("Tjobnr") then
				
				
				        if v <> 0 then
        				
				        call jobtotaler()
        				
				        kmTotpajob = 0
				        timerTotpajob = 0
				        enhederTotpajob = 0
				        afspadseringOpsTot = 0
				        timerFrokostTot = 0
				        ferieplanTot = 0
        				
				        end if
				
				
				
				 
				   
				
				tTop = 40
                tLeft = 0
                tWdth = globalWdt


                call tableDiv(tTop,tLeft,tWdth)

                
                       call jobansoglastfak()
                
                

                %>
                <table border="0" width=100% cellpadding="0" cellspacing="0" bgcolor="#ffffff">
				
				
				
				<%if komprimer <> "1" then %>
				<tr>
				    
					<td colspan="11" style="padding:5px; height:100px;">
					
					
					
					
					<table cellpadding=0 cellspacing=0 border=0 width=100%>
					<tr><td valign=top>
					<img src="../ill/ikon_kunder_16.png" alt="Kontakt" border="0">&nbsp;&nbsp;<%=oRec("Tknavn")%>&nbsp;(<%=oRec("kkundenr")%>)<br>
					<img src="../ill/ikon_job_16.png" alt="Job" border="0">&nbsp;&nbsp;<b><%=oRec("Tjobnavn")%> (<%=oRec("Tjobnr")%>)</b><br />
					
					
						Jobansvarlig 1: <b><%=jobans1txt%></b>
						<%if len(trim(jobans2txt)) <> 0 then%>
						<br>
	   					Jobansvarlig 2: <b><%=jobans2txt%></b>
						<%end if
						jobans1txt = ""
						jobans2txt = ""
						%>
						</td><td valign=top style="padding-left:20px;">
						
						
						<%select case jobstatus
						case 1
						thisjobstatus = "Aktiv"
						case 2
						thisjobstatus = "Passiv"
						case 0
						thisjobstatus = "Lukket"
						end select%>
						Jobstatus: <b><%=thisjobstatus%></b><br>
						Jobperiode: <b><%=formatdatetime(jobstartdato, 1)%> - <%=formatdatetime(jobslutdato, 1)%></b>
						<br>
						Forkalkuleret timer: <b><%=formatnumber(jobForkalkulerettimer, 2)%></b><br>
						
						<%if len(trim(rekvnr)) <> 0 then %>
						Rekvisitions nr: <b><%=rekvnr %></b><br />
						<%end if %>
						
						Sidste fakturadato: 
						<%if formatdatetime(lastfakdato, 2) <> "01-01-2001" then%>
						<b><%=formatdatetime(lastfakdato, 2)%></b>
						<%else%>
						-
						<%end if%>
						
						</td></tr>
						</table>
					    
						
					</td>
					
				</tr>
				<%end if '*** End Komprimer **'
				
				call tableheader()
				
				
				
				end if '*** job <> lastjob **'
				
				
				
		else '** joblog_uge = 2 **'
		
		
		            if cdate(lastdate) <> cdate(oRec("tdato")) then
		            
		             if v <> 0 then
		                    
		                call dagstotaler()
        				
				        
				        timerTotdag = 0
				        enhederTotdag = 0
				        
		             end if
		            
		            
		            end if
		            
				
				    if lastWeekNum <> strWeekNum OR lastmedarb <> oRec("Tmnr") then
				    
				    
				    
				   if v <> 0 then
				   
				
        				
				        call jobtotaler()
        				
				        kmTotpajob = 0
				        timerTotpajob = 0
				        enhederTotpajob = 0
				        afspadseringOpsTot = 0
				        timerFrokostTot = 0
				        ferieplanTot = 0
        				
				        
				    
				   
				   end if
				    
				
				tTop = 40
                tLeft = 0
                tWdth = globalWdt


                call tableDiv(tTop,tLeft,tWdth)



                %>
                <table border="0" width=100% cellpadding="0" cellspacing="0" bgcolor="#ffffff">
              
				    
					
				    <tr>
				        <td colspan=11 bgcolor="#ffffff" style="padding:10px 10px 0px 10px;">
				        <h4>Uge: <%=datepart("ww", oRec("tdato"), 2,2)&" - "& datepart("yyyy", oRec("tdato"), 2,2)%> - <%=oRec("tmnavn") %> (<%=oRec("tmnr") %>)</h4>
				        
				        <%
				        call erugeAfslutte(datepart("yyyy", oRec("tdato"), 2,2), datepart("ww", oRec("tdato"), 2,2), oRec("tmnr")) 
				        %>
				        
				        <%if showAfsuge = 0 then %>
				        Denne uge er afsluttet via smiley d. <%=formatdatetime(cdAfs, 2)%> kl. <%=formatdatetime(cdAfs, 3)%> (uge <%=datepart("ww", cdAfs, 2, 2)%>)
                        
		                <%end if%>
				        
				        </td>
				    </tr>
				    
				    
				    
				    <%
				    
				    
				    call tableheader()

				    end if
				    
				    
				
				
				
				
		end if '*** End Joblog_uge **'
		
		
				
		        
		         '*** Finder jobans og last fak hvis liste vises    ***'
                 '*** som ugeseddel, ellers er                      ***'
                 '*** jobans og last fak allrrede fundet            ***'
                 if cint(joblog_uge) = 2 then
                 call jobansoglastfak()
                 end if 
                 
                 '***'
		        
				
				%>
				
				
				
				
				
				<tr>
					<td bgcolor="#d6DFf5" colspan="11"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
				</tr>
				<tr>
				<td valign="top">&nbsp;</td>
				<td style="padding-top:3px;" valign="top"><%=strWeekNum%></td>
				<td align=right style="padding-top:3px; padding-right:10px;" valign="top"><b><font color="darkred"><b><%=left(weekdayname(weekday(oRec("Tdato"))), 3) %>. <%=day(oRec("Tdato")) &" "&left(monthname(month(oRec("Tdato"))), 3) &" "& right(year(oRec("Tdato")), 2)%></b></td>
				
				
				<%if komprimer = "1" then %>
				<td style="padding-top:3px; padding-right:5px;" valign="top">
				<b><%=oRec("Tknavn")%></b>&nbsp;(<%=oRec("kkundenr")%>)<br />
				
				<%=oRec("tjobnavn")%> (<%=oRec("tjobnr") %>)
				
				<font class=megetlillesilver>
				<%if jobans1 <> 0 then %>
				<br /><%=jobans1txt %>
				<%end if %>
				
				<%if jobans2 <> 0 then %>
				        
				        <%if jobans1 <> 0 then %>
				        <%=", " %>
				        <%end if %>
				        <%=jobans2txt %>
				<%end if %>
				</font>
				
				</td>
				<%else %>
				<td valign="top">
                    &nbsp;</td>
				<%end if %>
				<%
				 
				ekspTxt = ekspTxt & strWeekNum &";"& formatdatetime(oRec("Tdato"), 2) &";"
				
                '*** Er uge afsluttet af medarb, er smiley og autogk slået til ***'
                erugeafsluttet = instr(medabAflugeId(oRec("mid")), "#"&datepart("ww", oRec("Tdato"),2,2)&"_"&datepart("yyyy", oRec("Tdato"))&"#")
                 
                 
                
              
                 if (cint(erugeafsluttet) <> 0 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", oRec("Tdato")) = year(now) AND DatePart("m", oRec("Tdato")) < month(now)) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", oRec("Tdato")) < year(now) AND DatePart("m", oRec("Tdato")) = 12)) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND DatePart("yyyy", oRec("Tdato")) < year(now) AND DatePart("m", oRec("Tdato")) <> 12) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (year(now) - DatePart("yyyy", oRec("Tdato")) > 1))) then
              
                ugeerAfsl_og_autogk_smil = 1
                else
                ugeerAfsl_og_autogk_smil = 0
                end if 
				
				
				%>
				<td style="padding:3px 10px 5px 0px;" valign="top">
				<%
				if ((oRec("godkendtstatus") = 0 AND request("print") <> "j" AND ugeerAfsl_og_autogk_smil = 0) _
				OR (oRec("godkendtstatus") = 0 AND request("print") <> "j" AND ugeerAfsl_og_autogk_smil = 1 AND level = 1)) _
				AND (cdate(lastfakdato) <  cdate(oRec("Tdato"))) then %>
					
					
					
					<a href="rediger_tastede_dage_2006.asp?id=<%=oRec("Tid")%>" target="_blank" class=vmenu>
					
					<%=oRec("Anavn")%>
					
					<img src="../ill/blyant.gif" width="12" height="11" alt="Tilføj kommentar" border="0">
					</a>
					<%
					visning = 1
					call akttyper(oRec("fakturerbar"), visning)
					%>
				    &nbsp;<font class=megetlillesort>(<%=akttypenavn%>)</font>
					
					
					
				
				<%else%>
					<b><%=oRec("Anavn")%></b>
					<%
					visning = 1
					call akttyper(oRec("fakturerbar"), visning)
					%>
				    &nbsp;<font class=megetlillesort>(<%=akttypenavn%>)</font>
					
					
					
				<%end if
				
				if len(trim(oRec("timerkom"))) <> 0 then%>
				<br /><i><%=oRec("timerkom")%></i>
				<%end if%>
				
				
				&nbsp;</td>
				<%
				aktnavnEksp = ""
                if len(oRec("Anavn")) <> 0 then
                aktnavnEksp = replace(oRec("Anavn"), Chr(34), "''")
                else
                aktnavnEksp = ""
                end if
                
               
				call akttyper2009Prop(oRec("fakturerbar"))
				ekspTxt = ekspTxt & aktnavnEksp &";"& akttypenavn &";"& aty_fakbar &";" 
				
				if lto = "bowe" then
					ekspTxt = ekspTxt & akttid &";" 
				end if
				%>
				
				<td align="right" style="padding-top:3px; padding-right:5px;" valign="top"><b><%=formatnumber(oRec("Timer"), 2)%></b>
				<%
				ekspTxt = ekspTxt & oRec("Timer") &";"
				
				if len(oRec("sttid")) <> 0 then
				
					if left(formatdatetime(oRec("sttid"), 3), 5) <> "00:00" then
					Response.write "<font class=megetlillesort><br>"& left(formatdatetime(oRec("sttid"), 3), 5) & " - " & left(formatdatetime(oRec("sltid"), 3), 5)
					ekspTxt = ekspTxt & left(formatdatetime(oRec("sttid"), 3), 5) & " - " & left(formatdatetime(oRec("sltid"), 3), 5) &";"
					else
					ekspTxt = ekspTxt &";"
					end if
				
				else
				ekspTxt = ekspTxt &";"
				end if
				
				%></td>
				<%
				
				
				if cint(hideenheder) = 0 then 
				
				    enheder = 0
				    enheder = oRec("faktor") * oRec("timer")
    				
				    if enheder <> 0 then
				    enheder = cdbl(enheder)
				    else
				    enheder = 0
				    end if
				    %>
				    <td align="right" style="padding-top:3px; padding-right:5px;" valign="top"><%=formatnumber(enheder, 2)%></td>
				    <%
				    ekspTxt = ekspTxt & enheder &";"
				
				else
				%>
				<td>&nbsp;</td>
				<%
				end if
				
				
				select case oRec("tfaktim")
				case 1,2,6,13,14,20,21
				medarbtimer = medarbtimer + oRec("Timer")
				medarbEnheder = medarbEnheder + enheder
				timerTotpajob = timerTotpajob + oRec("Timer")
				enhederTotpajob = enhederTotpajob + enheder
				timerTotdag = timerTotdag + oRec("Timer")
				enhederTotdag = enhederTotdag + enheder
				case 5
				medarbKm = medarbKm + oRec("Timer") 
				kmTotpajob = kmTotpajob + oRec("Timer")
				case 10
				medarbFrokot = medarbFrokot + oRec("timer")
			    timerFrokostTot = timerFrokostTot + oRec("Timer")
				case 12
				afspadseringOpsTot = afspadseringOpsTot + oRec("timer")
				medarbAfspadOp = medarbAfspadOp + oRec("timer")
				case 11
				medarbFeriePlan = medarbFeriePlan + oRec("timer")
				ferieplanTot = ferieplanTot + oRec("timer")
				end select
				
				akttimer = akttimer + oRec("timer")
				aktEnheder = aktEnheder + enheder
				
				%>
				<td style="padding-top:3px; padding-left:10px;" valign="top" class=lille><%=left(oRec("Tmnavn"),16)%>&nbsp;(<%=oRec("tmnr")%>)</td>
				<%
				ekspTxt = ekspTxt & oRec("mnavn") &";"&oRec("mnr")&";"&oRec("init")&";"
				%>
				
				<td style="padding-top:3px; padding-right:5px;" align="right" valign="top">
				<%if (level <=2 OR level = 6) AND cint(hidetimepriser) = 0 then%>
				
				    <%
				    '*** Fakbar og ikke fakbare + salg ***
				    if oRec("tfaktim") = 1 OR oRec("tfaktim") = 2 OR oRec("tfaktim") = 6 then %>
    				
				    <b><%=formatnumber(oRec("TimePris"), 2)&" "&oRec("valutakode")%></b><br />
				    <% 
				    ekspTxt = ekspTxt & formatnumber(oRec("TimePris"), 2) &";"&oRec("valutakode")&";"
    				
				    end if
				
				end if%>
				<font class="megetlillesilver"><%=convertDate(oRec("TasteDato"))%></font></td>
				
				
				<%if cint(hidegkfakstat) <> 1 then %>
				<td class=lille valign=top align=center style="padding:3px 5px 3px 3px;">
                 <%
                 if cint(oRec("godkendtstatus")) = 0 then
                 gkCHK = ""
                 erGk = 0
                 else
                 gkCHK = "CHECKED"
                 erGk = 1
                 end if
                 
                 erGkaf = ""
                    
                 if len(trim(jobans1)) <> 0 then
                 jobans1 = jobans1
                 else
                 jobans1 = 0
                 end if
                 
                 if len(trim(jobans2)) <> 0 then
                 jobans2 = jobans2
                 else
                 jobans2 = 0
                 end if
                 
                 
                 
					if level = 1 OR _
					cint(session("mid")) = cint(jobans1) OR _
					cint(session("mid")) = cint(jobans2) then
					
					    '*** Godkendelse ***'
					     if print <> "j" then%>
    					
					     <input name="ids" id="ids" value="<%=oRec("tid")%>" type="hidden" />
					    <input type="checkbox" name="FM_godkendt" id="FM_godkendt_<%=v%>" value="1" <%=gkCHK %>>
                        <input name="FM_godkendt" id="FM_godkendt_hd_<%=v%>" value="#" type="hidden" />
					    <%
					    else
    					    
					         if cint(oRec("godkendtstatus")) = 1 then
				            %>
				            <b>Godkendt!</b>
				            <%
				           
				            end if
    					    
					    end if
					
					visopdaterknap = 1
					else
					    if cint(oRec("godkendtstatus")) = 1 then
				        %>
				        <b>Godkendt!</b>
				        <%
				        
				        end if
					end if
					
					if len(trim(oRec("godkendtstatusaf"))) <> 0 then%>
					<br /><i><%=left(oRec("godkendtstatusaf"), 15)%></i>
					<%
					erGkaf = oRec("godkendtstatusaf")
					end if %>
				
				</td>
				<td valign="top" style="padding:3px 5px 3px 3px;">
				
				<%
				erFaktureret = 0
				
				if cdate(lastfakdato) >=  cdate(oRec("Tdato")) then 
				
				erFaktureret = 1%>
				    
				    <%if (level = 1 OR _
					cint(session("mid")) = cint(jobans1) OR _
					cint(session("mid")) = cint(jobans2)) AND (print <> "j") then%>
					<a href="erp_fakhist.asp?FM_kunde=<%=oRec("kid")%>&FM_job=<%=jobid%>" class=vmenu target="_blank">Ja</a>
					<%else %>
					Ja
					<%end if %>
				
				
				
				<%end if %>
				</td>
				</tr>
				
				
				
				<%
				
				if lto <> "execon" then
				ekspTxt = ekspTxt & erGk &";"& erGkaf &";"& erFaktureret &";"
				end if
				
				else %>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
				    <%
				
				end if '** gkstat **'
				
				
				if len(oRec("timerkom")) <> 0 then
				
				kommEksp = ""
                
                if len(oRec("timerkom")) <> 0 then
                kommEksp = replace(oRec("timerkom"), Chr(34), "''")
                kommEksp = replace(kommEksp, Chr(39), "'")
                kommEksp = replace(kommEksp, Chr(45), "-")
                kommEksp = replace(kommEksp, "<br>", " ")
                kommEksp = replace(kommEksp, Chr(10), " ")
                kommEksp = replace(kommEksp, Chr(13), " ")
                else
                kommEksp = ""
                end if
                
               
				
				ekspTxt = ekspTxt & kommEksp &";"
				
				end if%>
				
				
				
				<!--
				'**** Kundekommentarer ****
					
					
					'strSQLin = "SELECT note FROM incidentnoter WHERE timerid = "& oRec("tid") &" ORDER BY id"
					'oRec3.open strSQLin, oConn, 3 
					'i = 0
					'Timergodkendt = 0
					'while not oRec3.EOF 
					'if i = 0 then%>
					'<tr>
					'<td valign="top" height="20">&nbsp;</td>
					'<td colspan=2>
                    '    &nbsp;</td>
					'<td valign="top" class=lillegray colspan=5 style="padding-left:3px;">
                    '    &nbsp;
					'<end if%>
					
					'<i><=oRec3("note")%></i><br><br>
                    '    &nbsp;
					
					'<
					
					
					'ekspTxt = ekspTxt & oRec3("note") &";"
					
					
				
					
					''if oRec3("status") = 1 then
					''Timergodkendt = 1
					''Response.write "<font class=roed><b>Registrering godkendt!</b></font>"
					''end if
					
					
					'i = i + 1
					'oRec3.movenext
					'wend
					'oRec3.close  
					
					'if i <> 0 then%>
					'	</td>
					'	<td colspan=2>
                     '   &nbsp;</td>
					'	
					'	<td valign="top">&nbsp;</td>
					'</tr>
					
					'<end if%>
					-->
				
				
			<%
			ekspTxt = ekspTxt & chr(013) 
			
			v = v + 1
			lastmedarbnavn = oRec("Tmnavn")
			lastmedarb = oRec("Tmnr")
			lastjobnr = oRec("Tjobnr")
			lastjobnavn = oRec("Tjobnavn")
			lastaktid = oRec("Taktivitetid")
			lastaktnavn = oRec("Anavn")
			lastakttype = akttypenavn
			lastWeekNum = strWeekNum 
			lastdate = oRec("tdato")
			'lastYearNum = 
	        
	        Response.flush
	       
	        oRec.movenext
	        wend
	        oRec.Close
	        'Set oRec = Nothing

    
    if cint(joblog_uge) = 1 then
    
    
	        if v > 0 AND fordelpamedarb = 1 then
        	
	        call fordelpamedarbejdere()
	        medarbtimer = 0
	        medarbEnheder = 0
	        medarbFeriePlan = 0
	        medarbFrokot = 0
	        medarbAfspadOp = 0
	        medarbKm = 0
        	
	        end if
        	
	        if v > 0 AND fordelpamedarb = 2 then
        				
        				
	        call fordelpaakt()
        	
	        akttimer = 0
	        aktEnheder = 0
        	
	        end if
	
	
	end if
	
	
	if v = 0 then
	%>
	<tr>
		<td colspan=11 valign="top" style="padding:20px; border:1px #5582d2 solid;"><br /><br /><b>Info:</b><br>
		Der blev ikke fundet registreringer der matcher de valgte kriterier.<br>
            &nbsp;</td>
	</tr>
	<%
	end if



if v <> 0 then
    if cint(joblog_uge) = 2 then
    call dagstotaler()
    end if
call jobtotaler()
call grandTotal




%>

<br /><br /><br />




	



<%if print <> "j" then%>
<br><br>
<% 
itop = 40
ileft = 0
iwdt = 400
call sideinfo(itop,ileft,iwdt)
%>
			<tr><td style="padding:10px;">
			<b>Side note(r):<br></b>
			*)Timepris er angivet som den timepris der er registret for den enkelte medarbejder. Der er <u>ikke taget højde</u> for om jobbet er et <b>Fastpris</b> job.
	        <br /><br />Aktivitetetstyper markeret med * er ikke medtaget i den daglige timeberegning.</td></tr></table>
			</div>
	<br><br><br><br><br><br><br><br><br>&nbsp;
<%end if 

if print <> "j" then

ptop = 100
pleft = 710
pwdt = 120

call eksportogprint(ptop,pleft,pwdt)
%>

<form action="joblog_eksport.asp" target="_blank" method="post" name=theForm2 onsubmit="BreakItUp2()"> <!--  -->
			<input type="hidden" name="datointerval" id="datointerval" value="<%=strDag&"/"&strMrd&"/"&strAar & " - " & strDag_slut&"/"&strMrd_slut&"/"&strAar_slut%>">
			<input type="hidden" name="txt1" id="txt1" value="">
			<input type="hidden" name="BigTextArea" id="BigTextArea" value="<%=ekspTxt%>">
            <input type="hidden" name="txt20" id="txt20" value="">

<tr>
    
    <%if len(ekspTxt) > 302399 then %>
    <td class=lille colspan=2>
    Mængden af data er for stor til eksport. Vælg et mindre interval ell. færre 
    medarbejdere. Størrelsen på data er: <%=len(ekspTxt) %> og den må ikke overstige 302399.
    <%else %>
    <td align=center><input type="image" src="../ill/export1.png"></td>
    </td><td>.csv fil eksport</td>
    <%end if %>
    </tr>
    <tr>
    <td align=center><a href="joblog.asp?rdir=<%=rdir %>&FM_kunde=<%=kundeid %>&menu=stat&FM_orderby_medarb=<%=fordelpamedarb%>&print=j&jobnr=<%=intJobnr%>&eks=<%=request("eks")%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=selmedarb%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd%>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>&FM_typer=<%=vartyper%>&FM_kundejobans_ell_alle=<%=visKundejobans%>&FM_jobsog=<%=jobSogVal%>&FM_radio_projgrp_medarb=<%=request("FM_radio_projgrp_medarb")%>&FM_komprimeret=<%=komprimer%>" target="_blank"  class='vmenu'>
   &nbsp;<img src="../ill/printer3.png" border=0 alt="" /></a>
</td><td>Print version</td>
   </tr>
   
	</form>
   </table>
</div>
<%else%>

<% 
Response.Write("<script language=""JavaScript"">window.print();</script>")
%>
<%end if%>



<%end if 'v <> 0 %>


</div>
<br>
<br>
&nbsp;
<%end if 'validering %>

<!--#include file="../inc/regular/footer_inc.asp"-->