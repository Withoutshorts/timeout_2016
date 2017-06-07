<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/webblik_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->
<!--#include file="inc/isint_func.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->




<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else
	
    menu = request("menu")
    thisfile = "pipeline"

    media = request("media")
    if media = "print" OR media = "export" then
        print = "j"
    end if

	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	'*** År ***'
		if len(Request("seomsfor")) <> 0 then
		seomsfor = Request("seomsfor")
		strYear = seomsfor
		strReqAar = strYear 
		else
			'if len(request("year")) <> 0 then
			'seomsfor = year("1/1/"&request("year"))
			'strYear = request("year")
			'strReqAar = strYear
			'else
			seomsfor = year(now)
			strYear = seomsfor
			strReqAar = strYear
			'end if 
		end if
	
    if len(trim(request("FM_start_mrd"))) <> 0 then
	strMrd = request("FM_start_mrd")
	else
	strMrd = month(now)
	end if

    if len(trim(request("seomsfor_antalmd"))) <> 0 then
    seomsfor_antalmd = request("seomsfor_antalmd")
    else
    seomsfor_antalmd = 12
    end if

    
    if len(trim(request("directexcel"))) <> 0 then
    directexcel = 1
    else
    directexcel = 0
    end if

    if cint(directexcel) = 1 then
        directexcelChk = "CHECKED"
    else
        directexcelChk = ""
    end if
        

	dim seomsfor_antalmdHigend
    seomsfor_antalmdHigend = seomsfor_antalmd - 1

    if len(trim(request("FM_start_mrd"))) <> 0 then
	
        if len(trim(request("FM_ign_per"))) <> 0 then    
          ign_per = 1
        else
          ign_per = 0
        end if

    else

        ign_per = 0

    end if


    if cint(ign_per) = 1 then
        ign_perCHK = "CHECKED"
    else
        ign_perCHK = ""
    end if

	'**************************************************
	'***** Faste Filter kriterier *********************
	'**************************************************
		
	
		
	'*** Job og Kundeans ***
	call kundeogjobans()
	
	'*** Medarbejdere / projektgrupper
	selmedarb = session("mid")
	'call medarbogprogrp("oms")
	medarbSQlKri = ""
	kundeAnsSQLKri = ""
	jobAnsSQLkri = ""
	jobAns2SQLkri = ""
	fakmedspecSQLkri = ""
    salgsansSQLkri = ""
	
	if len(trim(request("FM_progrp"))) <> 0 then
	progrp = request("FM_progrp")
	else
	progrp = 0
	end if
	
	'*** Rettigheder på den der er logget ind **'
	medarbid = session("mid")
	
	if len(request("FM_medarb")) <> 0 OR func = "export" then
	
	    if left(request("FM_medarb"), 1) = "0" then
	        if media <> "print" then
	        thisMiduse = request("FM_medarb_hidden")
    	    else
    	    thisMiduse = request("FM_medarb")
    	    end if
    	
    	intMids = split(thisMiduse, ", ")
	    else
	    thisMiduse = request("FM_medarb")
	    intMids = split(thisMiduse, ", ")
	    end if
	
	else
	    
	    thisMiduse = session("mid") 
	    intMids = split(thisMiduse, ", ")
	   
	end if
	
	
	'**** Kundekri ***
	if len(request("FM_kunde")) <> 0 then
	kundeid = request("FM_kunde")
	else
	kundeid = 0
	end if

    call salgsans_fn()


    

	
	'*** Kundeans ***
	strKnrSQLkri = ""
	
	'*** finder udfra valgte projektgrupper og medarbejdere
	'medarbSQlKri 
	'kundeAnsSQLKri
	
			    for m = 0 to UBOUND(intMids)
			    
			     if m = 0 then
			     medarbSQlKri = "(m.mid = " & intMids(m)
			     kundeAnsSQLKri = "kundeans1 = "& intMids(m) &" OR kundeans2 = "& intMids(m) &""
			     jobAnsSQLkri = "jobans1 = "& intMids(m)  
			     jobAns2SQLkri = "jobans2 = "& intMids(m)
                 jobAns3SQLkri = "jobans3 = "& intMids(m) & " OR jobans4 = "& intMids(m) & " OR jobans5 = "& intMids(m)

                   if cint(showSalgsAnv) = 1 then
                   salgsansSQLkri = "salgsans1 = "& intMids(m) & " OR salgsans2 = "& intMids(m) & " OR salgsans3 = "& intMids(m) & " OR salgsans4 = "& intMids(m) & " OR salgsans5 = "& intMids(m) 
                   end if

			     fakmedspecSQLkri = "fms.mid = "& intMids(m)
			     else
			     medarbSQlKri = medarbSQlKri & " OR m.mid = " & intMids(m)
			     kundeAnsSQLKri = kundeAnsSQLKri & " OR kundeans1 = "& intMids(m) &" OR kundeans2 = "& intMids(m) &""
			     jobAnsSQLkri = jobAnsSQLkri & " OR jobans1 = "& intMids(m)  
			     jobAns2SQLkri = jobAns2SQLkri & " OR jobans2 = "& intMids(m)
                 jobAns3SQLkri = jobAns3SQLkri & " OR jobans3 = "& intMids(m) & " OR jobans4 = "& intMids(m) & " OR jobans5 = "& intMids(m)

                  if cint(showSalgsAnv) = 1 then
                   salgsansSQLkri = salgsansSQLkri & " OR salgsans1 = "& intMids(m) & " OR salgsans2 = "& intMids(m) & " OR salgsans3 = "& intMids(m) & " OR salgsans4 = "& intMids(m) & " OR salgsans5 = "& intMids(m) 
                   end if

			     fakmedspecSQLkri = fakmedspecSQLkri & " OR fms.mid = "& intMids(m)
			     end if
			     
			    next
			    
			    medarbSQlKri = medarbSQlKri & ")"
			    
			    jobAnsSQLkri =  " ("& jobAnsSQLkri &")"
			    jobAns2SQLkri =  "xx (" & jobAns2SQLkri &")"
			    jobAns3SQLkri =  "xx (" & jobAns3SQLkri &")"
            
                   if cint(showSalgsAnv) = 1 then
                   salgsansSQLkri =" (" & salgsansSQLkri &")"
                   end if

            fakmedspecSQLkri = " AND ("& fakmedspecSQLkri &")"
			
	
	'** Er key acc og kundeansvarlig valgt **?
	if cint(kundeans) = 1 then
	kundeAnsSQLKri = kundeAnsSQLKri
	else
	kundeAnsSQLKri = " Kundeans1 <> -1 AND kundeans2 <> -1 "
	end if
	
	
	'***** Valgt job eller søgt på Job ****
	'** hvis Sog = Ja
	call jobsog()

	
	'*** Aftale ****
	'if len(request("FM_aftaler")) <> 0 then ' AND jobid <= 0 AND len(jobSogVal) = 0 then
	'aftaleid = request("FM_aftaler")
		
		'if cint(aftaleid) > 0 then 
		'jobid = 0
		'end if
		
	'else
		'if cint(jobid) > 0 OR len(jobSogVal) <> 0  then
		aftaleid = -1
		'else
	'	aftaleid = 0
		'end if
	'end if
	
	
	
	'**** Alle SQL kri starter med NUL records ****
	jobidFakSQLkri = " OR jobid = -1 "
	jobnrSQLkri = " OR tjobnr = '-1' "
	jidSQLkri = " OR id = -1 "
	seridFakSQLkri = " OR aftaleid = -1 "
	
	
    
       
	   
	
	
	
	
	
	
	'************ slut faste filter var **************		
	

	if media = "print" OR media = "export" then
	'leftPos = 20
	'topPos = 62
    %>
    <!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<%
	else
	'leftPos = 90
	'topPos = 102
    %>
    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	

  
    <script src="inc/pipeline_jav.js"></script>
	<!--include file="../inc/regular/topmenu_inc.asp"-->

	

	<%

        call menu_2014()

	end if
	%>


    <!--#include file="inc/dato2_b.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	
	<%if media <> "print" AND media <> "export" then%>
	
	<%
    'oleftpx = 1065
	'otoppx = 115
	'owdtpx = 120
	
	'call opretNy("obs.asp?menu=job&func=opret&id=0&int=1&rdir=pipe", "Opret nyt job", otoppx, oleftpx, owdtpx) 
	
    %>

	
	
	
    
    
    <div id="loadbar" style="position:absolute; display:; visibility:visible; top:260px; left:220px; width:300px; background-color:#ffffff; border:1px #cccccc solid; padding:2px; z-index:100000000;">

	<table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	<br />
	Forventet loadtid:
	<%
	
	exp_loadtid = 0
	'exp_loadtid = (((len(akttype_sel) / 3) * (len(antalvlgM) / 3)) / 50)  %> 
	ca.: <b>3-5 sek.</b>
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" />
	
	</td></tr></table>

	</div>
	
    <%end if%>

	
	<%
	if media <> "print" AND media <> "export" then
	pleft = 90
	ptop = 102 '132
	'ptopgrafik = 348
	else
	pleft = 10
	ptop = 20
	'ptopgrafik = 90
	end if

    if media <> "export" then
	%>	
	<div id="sidediv" style="position:absolute; left:<%=pleft%>px; top:<%=ptop%>px; visibility:visible;">
	
	
	<%
	
	oimg = "ikon_gantt_48.png"
	oleft = 0
	otop = 0
	owdt = 200
	oskrift = "Pipeline"
	
    if media = "print" then
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
    end if
    end if
	
	if media <> "print" AND media <> "export" then 
	
	call filterheader_2013(0,0,800,oskrift)%>
    	    
	  
	    <table cellspacing=0 cellpadding=0 border=0 width=100% bgcolor="#FFFFFF">
	    <form action="pipeline.asp?rdir=<%=rdir %>" method="post">
	    
	<%end if %>
	    
        
        <%call medkunderjob %>
	    
	    <%if media <> "print" AND media <> "export" then %>
	    
	    </td>
	    </tr>
	   
	
	
	<tr>
		<td valign=top>
        <br />
        <b>Startdato fra:</b> 
	<select name="FM_start_mrd">
		<option value="<%=strMrd%>"><%=strMrdNavn%></option>
		<option value="1">jan</option>
	   	<option value="2">feb</option>
	   	<option value="3">mar</option>
	   	<option value="4">apr</option>
	   	<option value="5">maj</option>
	   	<option value="6">jun</option>
	   	<option value="7">jul</option>
	   	<option value="8">aug</option>
	   	<option value="9">sep</option>
	   	<option value="10">okt</option>
	   	<option value="11">nov</option>
	   	<option value="12">dec</option></select>
	&nbsp;
        
     <b>År:</b>&nbsp;
	<select name="seomsfor" id="seomsfor" style="width:85px;">
			
			<%
			useaar = 2001
			for x = 0 to 20
			useaar = useaar + 1
			
			if cint(strYear) = cint(useaar) then
			aChk = "SELECTED"
			else
			aChk = ""
			end if
			%>
			<option value="<%=useaar%>" <%=aChk%>><%=useaar%></option>
			<%
			'x = x + 1
			next%>
		</select>

            og

        <%
        
            seomsfor_antalmd12sel = ""
            seomsfor_antalmd24sel = ""
            seomsfor_antalmd36sel = ""    
            
        select case seomsfor_antalmd
        case 12
        seomsfor_antalmd12sel = "SELECTED" 
        case 24
        seomsfor_antalmd24sel = "SELECTED"
        case 36
        seomsfor_antalmd36sel = "SELECTED"
        end select
            %>
     <select name="seomsfor_antalmd">
         <option value="12" <%=seomsfor_antalmd12sel%>>12 md frem.</option>
         <option value="24" <%=seomsfor_antalmd24sel%>>24 md frem.</option>
         <option value="36" <%=seomsfor_antalmd36sel%>>36 md frem.</option>
     </select>  <br />
            <span style="font-size:10px; color:#999999;">Ved 12 md. bliver fakturaplan udspec. pr. md i valgte interval</span>
     <br />
   <input type="checkbox" name="FM_ign_per" value="1" <%=ign_perCHK %> /> Ignorer periode<br />
   <input type="checkbox" name="directexcel" value="1" <%=directexcelChk %> /> Eksporter direkte til excel
   </td>
	<td align=right>
	&nbsp;<input type="submit" value=" Kør >> ">
	</td></tr>
    </form>
	</table>
	
	
	<!--filter div-->
	</td></tr>
	</table>
	</div>
	
	<%end if %>
	
	
	
	
	
        <%
	    '*****************************************************************************************************
		'**** Job og aftaler der skal indgå i omsætning, budget og forbrugstal *******************************
		'*****************************************************************************************************
		
		
		'*** valgte job ***
		call valgtejob()
		
        'Response.write "jobnrSQLkri" & jobnrSQLkri

		
		'**** Valgte aftaler *****
		'call valgteaftaler()
		
		
		
		'*** For at spare (trimme) på SQL hvis der vælges alle job alle kunder og vis kun for jobanssvarlige ikke er slået til ****
		'*** Og der ikke er søgt på jobnavn ***
		if cint(kundeid) = 0 AND cint(jobid) = 0 AND cint(jobans) = 0 AND cint(jobans2) = 0 AND cint(jobans3) = 0 AND cint(kundeans) = 0 AND len(trim(jobSogVal)) = 0 AND cint(aftaleid) = 0 AND cint(segment) = 0 then 
				
			'jobidFakSQLkri = " OR jobid <> 0 "
			jobnrSQLkri = " OR tjobnr <> '0' "
			jidSQLkri =  " OR id <> 0 "
			'seridFakSQLkri = " OR aftaleid <> 0 "
				
		end if
	
	
		'**************** Trimmer SQL states ************************
		len_jobidFakSQLkri = len(jobidFakSQLkri)
		right_jobidFakSQLkri = right(jobidFakSQLkri, len_jobidFakSQLkri - 3)
		jobidFakSQLkri =  right_jobidFakSQLkri
		
		len_jobnrSQLkri = len(jobnrSQLkri)
		right_jobnrSQLkri = right(jobnrSQLkri, len_jobnrSQLkri - 3)
		jobnrSQLkri =  right_jobnrSQLkri
		
		len_jidSQLkri = len(jidSQLkri)
		right_jidSQLkri = right(jidSQLkri, len_jidSQLkri - 3)
		jidSQLkri =  right_jidSQLkri
		
		
		len_seridFakSQLkri = len(seridFakSQLkri)
		right_seridFakSQLkri = right(seridFakSQLkri, len_seridFakSQLkri - 3)
		seridFakSQLkri =  right_seridFakSQLkri
		'*****************************************************************************************************
	
	
	


	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	
	

    if media <> "export" then
	
    tTop = 20
	tLeft = 0
	tWdth = 1300
	
	
	call tableDiv(tTop,tLeft,tWdth)

    end if


    call basisValutaFN()


    if cint(showSalgsAnv) = 1 then
        cps = 35
    else
        cps = 20
    end if

    
    
     

    redim strMrdno(seomsfor_antalmdHigend)
    redim strMrdnm(seomsfor_antalmdHigend)
    redim strYearno(seomsfor_antalmdHigend)
    redim strDayno(seomsfor_antalmdHigend)
    redim tableHTML(seomsfor_antalmdHigend)
    redim mdIdTot(seomsfor_antalmdHigend)
    redim mdIdBruttoOmsTot(seomsfor_antalmdHigend)
    redim mdIdFakturaTot(seomsfor_antalmdHigend)
    redim mdIdStadeTot(seomsfor_antalmdHigend)
    redim stadesum_prmd(seomsfor_antalmdHigend)
    redim stadesum_prmdGT(seomsfor_antalmdHigend)
    stadesum_prGT = 0


    for m = 0 to seomsfor_antalmdHigend

    strMrdno(m) = strMrd + m
    if strMrdno(m) <= 12 then
    strMrdno(m) = strMrdno(m)
    strYearno(m) = strYear
    else
    strMrdno(m) = strMrdno(m) - 12
    strYearno(m) = strYear + 1
    end if

    if m = seomsfor_antalmdHigend then
    strMrdEnd = strMrdno(m)
    strYearEnd = strYearno(m) 
    end if

    if seomsfor_antalmdHigend = 11 then
    strMrdnm(m) = monthname(strMrdno(m))
    end if

    select case strMrdno(m)
    case 1,3,5,7,8,10,12
    strDayno(m) = 31
    case 2
        select case strYearno(m)
        case 2000, 2004,2008,2012,2016,2020,2024,2028,2032,2036,2040,2044
        strDayno(m) = 29
        case else
        strDayno(m) = 28
        end select
    case else
    strDayno(m) = 30
    end select


    if m = seomsfor_antalmdHigend then
    strDayEnd = strDayno(m)
    end if

    if cint(directexcel) <> 1 then
    tableHTML(m) = "<table cellspacing=0 cellpadding=0 border=0 width='100%'>"
    end if

    next
    
    sqlDatostartFY = year(now) & "-1-1" 
    sqlDatoslutFY = year(now) & "-12-31"

    stadeStDato = strYear & "-" & strMrd & "-1"
   

    if media <> "export" AND cint(directexcel) <> 1 then %>
	
    <table cellspacing="0" cellpadding="0" border="0">
        
     <tr><td valign=top>
	
	<%

    


	'strBudget = "<table cellspacing=1 cellpadding=0 border=1>"
	strJob = "<table cellspacing=0 cellpadding=0 border=0><tr bgcolor=""#5582d2""><td class=alt valign=bottom style=""padding:2px; border-right:1px #cccccc solid; width:250px;""><b>Jobnavn (nr)</b><br>Kunde (knr)</td>"_
    &"<td class=alt valign=bottom style=""padding:2px; border-right:1px #cccccc solid;""><b>Brutto oms. "&basisValISO&"</b>"_
    &"<td class=alt valign=bottom style=""padding:2px; border-right:1px #cccccc solid;"" class=alt><b>Startdato</b> (måneder)</td>"_
    &"<td class=alt valign=bottom style=""padding:2px; border-right:1px #cccccc solid;"" class=alt><b>Status</b></td>"_
    &"<td class=alt valign=bottom style=""padding:2px; border-right:1px #cccccc solid;"" class=alt><b>Lukkedato</b></td>"_
    &"<td class=alt valign=bottom style=""padding:2px; border-right:1px #cccccc solid;""><b>Sand. %</b></td>"_
    &"<td class=alt style=""padding:2px; border-right:1px #cccccc solid;""><b>Pipeline værdi</b><br>(Brutto Oms - Udgifter ulev. * Sand.) "& basisValISO &"</td>"_
    &"<td class=alt valign=bottom style=""padding:2px; border-right:1px #cccccc solid;""><b>Stade / Fak. plan</b> "& basisValISO &"</td>"

        if seomsfor_antalmdHigend = 11 then
          for m = 0 to seomsfor_antalmdHigend

            strJob = strJob &"<td class=alt valign=bottom style=""padding:2px; background-color:#999999; border-right:1px #cccccc solid;""><b>Fak. plan</b><br>"& basisValISO &"<br> "& left(monthname(strMrdno(m)), 3) &" "& right(strYearno(m), 2) &"</td>"

          next
        end if

    strJob = strJob &"<td class=alt valign=bottom style=""padding:2px; border-right:1px #cccccc solid;""><b>Faktureret</b> "& basisValISO &"</td>"_
    &"<td class=alt valign=bottom style=""padding:2px; border-right:1px #cccccc solid;""><b>Faktureret FY "& year(now) &"</b> "& basisValISO &"</td>"_
    &"<td class=alt valign=bottom style=""padding:2px; border-right:1px #cccccc solid;""><b>Faktureret før FY "& year(now) &"</b> "& basisValISO &"</td>"_
    &"<td class=alt style=""padding:2px; border-right:1px #cccccc solid;"" valign=bottom colspan=3><b>Job-ansv.</b><br>Værdi "& basisValISO &"</td>"_
    &"<td class=alt style=""padding:2px; border-right:1px #cccccc solid;"" valign=bottom colspan=3><b>Job-ejer</b><br>Værdi "& basisValISO &"</td>"_
    &"<td class=alt style=""padding:2px; border-right:1px #cccccc solid;"" valign=bottom colspan=3><b>Job medansv. 3</b><br>Værdi "& basisValISO &"</td>"_
    &"<td class=alt style=""padding:2px; border-right:1px #cccccc solid;"" valign=bottom colspan=3><b>Job medansv. 4</b><br>Værdi "& basisValISO &"</td>"_
    &"<td class=alt style=""padding:2px; border-right:1px #cccccc solid;"" valign=bottom colspan=3><b>Job medansv. 5</b><br>Værdi "& basisValISO &"</td>"
    
    if cint(showSalgsAnv) = 1 then

    strJob = strJob &"<td class=alt style=""padding:2px; border-right:1px #cccccc solid;"" valign=bottom colspan=3><b>Salgsansvr. 1</b><br>Værdi "& basisValISO &"</td>"_
    &"<td class=alt style=""padding:2px; border-right:1px #cccccc solid;"" valign=bottom colspan=3><b>Salgsansvr. 2</b><br>Værdi "& basisValISO &"</td>"_
    &"<td class=alt style=""padding:2px; border-right:1px #cccccc solid;"" valign=bottom colspan=3><b>Salgsansvr. 3</b><br>Værdi "& basisValISO &"</td>"_
    &"<td class=alt style=""padding:2px; border-right:1px #cccccc solid;"" valign=bottom colspan=3><b>Salgsansvr. 4</b><br>Værdi "& basisValISO &"</td>"_
    &"<td class=alt style=""padding:2px; border-right:1px #cccccc solid;"" valign=bottom colspan=3><b>Salgsansvr. 5</b><br>Værdi "& basisValISO &"</td>"

    end if

    strJob = strJob &"<td class=alt style=""padding:2px; border-right:1px #cccccc solid;"" valign=bottom><b>Forretningsområder (units)</b>"
    'strJob = strJob &"<td class=alt style=""padding:2px;"" valign=bottom><b>Kode</b></td>"_
    strJob = strJob &"</tr>"
	
    end if
    
	select case viskun
	case 1
	strViskunKri = " AND jobstatus = 1 "
	case 2
	strViskunKri = " AND (jobstatus = 1 OR jobstatus = 3)" 'alle aktive + tilbud
	case 3
	strViskunKri = " AND jobstatus = 3 " 'tilbud
	end select
	
    select case lto
    case "epi2017"
	kvotient = 10000
    mdpeer = 5000000
    case else 
    kvotient = 100
    mdpeer = 100000
    end select
	
    jobnrSQLkri = replace(jobnrSQLkri, "tjobnr", "jobnr")
    
  

	strSQL = "SELECT job.id AS jobid, jobnr, jobnavn, jobstartdato, jobslutdato, lukkedato, jo_bruttooms, jobTpris, fakturerbart, budgettimer, fastpris, jobstatus, sandsynlighed, jo_udgifter_ulev, jo_valuta, jo_valuta_kurs, "_
    &" jobans1, jobans2, jobans3, jobans4, jobans5, jobans_proc_1, jobans_proc_2, jobans_proc_3, jobans_proc_4, jobans_proc_5, "


    if cint(showSalgsAnv) = 1 then

    strSQL = strSQL &" salgsans1, salgsans2, salgsans3, salgsans4, salgsans5, salgsans1_proc, salgsans2_proc, salgsans3_proc, salgsans4_proc, salgsans5_proc, "

     end if

    strSQL = strSQL &" m1.mnavn AS m1navn, m1.mnr AS m1mnr, m1.init AS m1init, "_
    &" m2.mnavn AS m2navn, m2.mnr AS m2mnr, m2.init AS m2init, "_
    &" m3.mnavn AS m3navn, m3.mnr AS m3mnr, m3.init AS m3init, "_
    &" m4.mnavn AS m4navn, m4.mnr AS m4mnr, m4.init AS m4init, "_
    &" m5.mnavn AS m5navn, m5.mnr AS m5mnr, m5.init AS m5init, "


    if cint(showSalgsAnv) = 1 then

        strSQL = strSQL &" s1.mnavn AS s1navn, s1.mnr AS s1mnr, s1.init AS s1init, "_
        &" s2.mnavn AS s2navn, s2.mnr AS s2mnr, s2.init AS s2init, "_
        &" s3.mnavn AS s3navn, s3.mnr AS s3mnr, s3.init AS s3init, "_
        &" s4.mnavn AS s4navn, s4.mnr AS s4mnr, s4.init AS s4init, "_
        &" s5.mnavn AS s5navn, s5.mnr AS s5mnr, s5.init AS s5init, "


     end if

    strSQL = strSQL &" kkundenavn, kkundenr "_
    &" FROM job"_
    &" LEFT JOIN kunder k ON (k.kid = jobknr) "_
    &" LEFT JOIN medarbejdere m1 ON (m1.mid = jobans1) "_
    &" LEFT JOIN medarbejdere m2 ON (m2.mid = jobans2) "_
    &" LEFT JOIN medarbejdere m3 ON (m3.mid = jobans3) "_
    &" LEFT JOIN medarbejdere m4 ON (m4.mid = jobans4) "_
    &" LEFT JOIN medarbejdere m5 ON (m5.mid = jobans5) "

     if cint(showSalgsAnv) = 1 then

    strSQL = strSQL &" LEFT JOIN medarbejdere s1 ON (s1.mid = salgsans1) "_
    &" LEFT JOIN medarbejdere s2 ON (s2.mid = jobans2) "_
    &" LEFT JOIN medarbejdere s3 ON (s3.mid = jobans3) "_
    &" LEFT JOIN medarbejdere s4 ON (s4.mid = jobans4) "_
    &" LEFT JOIN medarbejdere s5 ON (s5.mid = jobans5) "


     end if

        strSQL = strSQL &" WHERE ("& jobnrSQLkri &") "& replace(jobstKri, "j.", "job.")         
         
        stDatoSQLkri  = strYear &"/"& strMrd &"/1"
        endDatoSQLkri = strYearEnd&"/"& strMrdEnd &"/"& strDayEnd 
        
        if cint(ign_per) <> 1 then
        strSQL = strSQL &" AND jobstartdato BETWEEN '"& stDatoSQLkri &"' AND '"& endDatoSQLkri &"'"
        end if

        strSQL = strSQL &" GROUP BY job.id ORDER BY jobstartdato, jobnavn" 

    '" AND "& jobAnsSQLkri &""

   ' Response.Write " " & slagsansSQLkri &" "& jobAnsSQLkri &"<br>"

    'if session("mid") = 1 then
    'Response.Write strSQL
    'Response.flush
    'end if

    'strExport = "xx99123sy#z"
    lastMonth = 0
	mdTot = 0
    mdBruttoOmsTot = 0		
			
			
    
    oRec.open strSQL, oConn, 3
	
	y = 1
	while not oRec.EOF
	select case right(y, 1)
	case 2,4,6,8,0
	bgcolthis = "#EFF3ff"
	case else
	bgcolthis = "#FFFFFF"
	end select
	
   
    

	stmonth = month(oRec("jobstartdato"))
	endmonth = month(oRec("jobslutdato"))
	thisdays = datediff("m", oRec("jobstartdato"), oRec("jobslutdato"))
	thisbudget = pipelinevalue 'oRec("jobtpris")
	
        if cint(directexcel) <> 1 then

   
            if datepart("m", oRec("jobstartdato"),2,2) <> lastMonth then

                if y <> 1 then
                'strJob = strJob & "<tr bgcolor='#FFFFFF'><td style='padding:2px 10px 2px 2px; border-bottom:1px #cccccc solid;' class=lille><br>&nbsp;</td><td colspan=4 align=right style='padding:2px 10px 2px 2px; border-bottom:1px #cccccc solid;' class=lille valign=top><b>"& formatnumber(mdTot, 2) &"</b></td><td colspan="& cps &" style='padding:3px; border-bottom:1px #cccccc solid;'>&nbsp;</td></tr>"
                strJob = strJob & "<tr bgcolor='#FFFFFF'><td style='padding:2px 10px 2px 2px;' class=lille>Total:</td>"
                strJob = strJob & "<td align=right style='padding:2px 10px 2px 2px;' class=lille><b>"& formatnumber(mdBruttoOmsTot, 2) &"</b></td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>"
                strJob = strJob & "<td align=right style='padding:2px 10px 2px 2px;' class=lille><b>"& formatnumber(mdTot, 2) &"</b></td>"
        
                    
                     if seomsfor_antalmdHigend = 11 then
                      for m = 0 to seomsfor_antalmdHigend

                        strJob = strJob &"<td>&nbsp;</td>"

                      next
                    end if


                strJob = strJob &"<td colspan="& cps &" style='padding:3px; border-bottom:0px #cccccc solid;'>&nbsp;</td></tr>"
    
        
                mdTot = 0
                mdBruttoOmsTot = 0
                end if


            if seomsfor_antalmdHigend = 11 then
            cpsMonthTr = 37
            else
            cpsMonthTr = 25
            end if

            'strJob = strJob & "<tr bgcolor='#FFFFFF'><td colspan="& cpsMonthTr &" style='padding:3px;'>&nbsp;</td></tr>"
            strJob = strJob & "<tr bgcolor='#D6dff5'><td colspan="& cpsMonthTr &" style='padding:3px;'><b>"& monthname(datepart("m",oRec("jobstartdato"),2,2)) &" "& datepart("yyyy",oRec("jobstartdato"), 2,2) &"</b></td></tr>"
    
            end if


        end if
    
    
    '*** Tilbud Pipelineværdi = bruttooms - udgifter * sand *** / Ellers = bruttooms - udgifter
    if oRec("jobstatus") = 3 then
    pipelinevalue = formatnumber((oRec("jo_bruttooms") - oRec("jo_udgifter_ulev")) * (oRec("sandsynlighed")/100),2)
    else
    pipelinevalue = formatnumber((oRec("jo_bruttooms") - oRec("jo_udgifter_ulev")),2)
    end if 

    jo_bruttooms = oRec("jo_bruttooms")

    if cint(oRec("jo_valuta")) <> cint(basisValId) then

    call valutakode_fn(oRec("jo_valuta"))
    call beregnValuta(pipelinevalue,oRec("jo_valuta_kurs"),100)
    pipelinevalue = formatnumber(valBelobBeregnet, 2)
   
   
    call beregnValuta(jo_bruttooms,oRec("jo_valuta_kurs"),100)
    jo_bruttooms = formatnumber(valBelobBeregnet, 2)

    end if

    'pipelinevalue = oRec("jo_bruttooms")

	if cint(datepart("yyyy",oRec("jobstartdato"))) <= cint(strYear) AND cint(datepart("yyyy",oRec("jobslutdato"))) >= cint(strYear) then
	
    if cint(directexcel) <> 1 then
    strJob = strJob & "<tr bgcolor="& bgcolthis &"><td valign=top style='padding:2px; border-right:1px #cccccc solid;' class=lille><input type='hidden' name='FM_colthis_"&y&"' value='"&col&"'>"
	strJob = strJob & left(oRec("kkundenavn"), 15) & " ("& oRec("kkundenr") &")" & "<br>"

	    if media <> "print" AND level <=2 OR level = 6 then
	    strJob = strJob & "<a href='jobs.asp?menu=job&func=red&int=1&id="&oRec("jobid")&"&rdir=pipe&nomenu=1' class='vmenu' target=""_blank"">" & left(oRec("jobnavn"), 35) & " ("& oRec("jobnr")  &")</a>"
	    else
	    strJob = strJob & "<b>"& left(oRec("jobnavn"), 35)  & "</b> ("& oRec("jobnr")  &")"
	    end if

    end if

    strExport = strExport & oRec("jobnavn") & ";"& oRec("jobnr") &";"
	
    Select case oRec("jobstatus")
    case 3 
    jobstatusExpTxt = "Tilbud"
    'strJob = strJob & " Tilbud"
    sandsynlighedTxt = oRec("sandsynlighed") & "%"
    case 0
    jobstatusExpTxt = "Lukket"
    'strJob = strJob & " Lukket"
    sandsynlighedTxt = ""
    case 2
    jobstatusExpTxt = "Passivt"
    'strJob = strJob & " Passivt"
    sandsynlighedTxt = ""
    case else
    jobstatusExpTxt = "Aktivt"
    'strJob = strJob & "Aktivt"
    sandsynlighedTxt = ""
    end select
    
   strExport = strExport & jobstatusExpTxt &";"

    if cint(directexcel) <> 1 then
	strJob = strJob & "</td>" 
    end if

   strExport = strExport & oRec("kkundenavn") & ";"& oRec("kkundenr") &";"
	
    if cint(directexcel) <> 1 then
	    strJob = strJob & "<td valign=top align=right style='padding:2px 5px 2px 2px; border-right:1px #cccccc solid; white-space:nowrap;' class=lille>" & jo_bruttooms &"</td>"_
	    &"<td valign=top align=right style='padding:2px; white-space:nowrap; border-right:1px #cccccc solid;' class=lille>"& formatdatetime(oRec("jobstartdato"), 2) &" ("& thisdays &") </td>"_
         &"<td valign=top style='padding:2px; white-space:nowrap; border-right:1px #cccccc solid;' class=lille>"& jobstatusExpTxt &"</td>"_
        &"<td valign=top align=right style='padding:2px; white-space:nowrap; border-right:1px #cccccc solid;' class=lille>"
        
            if oRec("jobstatus") = 0 then
             strJob = strJob & formatdatetime(oRec("lukkedato"), 2)
             else
            strJob = strJob &"&nbsp;"
            end if 
    
        strJob = strJob &"</td>"_
        &"<td valign=top align=right style='padding:2px; border-right:1px #cccccc solid;' class=lille>"& sandsynlighedTxt &"</td>"_
	    &"<td valign=top align=right style='padding:2px 10px 2px 2px; border-right:1px #cccccc solid; white-space:nowrap;' class=lille>"& pipelinevalue &"</td>"
    end if


    '*** Stade / fakturering likividitet 'ALTID DKK
    stadesum = 0
    strSQLmilepale = "SELECT SUM(belob) AS stadesum FROM milepale WHERE type = 1 AND jid = "& oRec("jobid") &" AND milepal_dato >= '"& stadeStDato &"' GROUP BY jid"
    oRec6.open strSQLmilepale, oConn, 3
    if not oRec6.EOF then
        
    '** Omregn til basis valuta
    stadesum = formatnumber(oRec6("stadesum")*(oRec("jo_valuta_kurs")/100), 2)
    'stadesum = oRec6("stadesum")

    end if 
    oRec6.close

   stadesum_prGT = stadesum_prGT + stadesum

   if stadesum <> 0 then
    stadesumTxt = formatnumber(stadesum, 2)
    stadesumExp = formatnumber(stadesum, 2)
    else
    stadesumTxt = ""
    stadesumExp = 0
    end if


    '*** Faktureret Omregnes til DKK
    fakturasum = 0
    strSQLfaktureret = "SELECT SUM(IF(faktype = 0, f.beloeb * (f.kurs/100), f.beloeb * -1 * (f.kurs/100))) AS beloeb FROM fakturaer AS f WHERE shadowcopy <> 1 AND jobid = "& oRec("jobid") &" GROUP BY jobid"
   
    'if session("mid") = 1 then
    '    Response.write strSQLfaktureret
    '    Response.flush
    'end if   
        
    oRec6.open strSQLfaktureret, oConn, 3
    if not oRec6.EOF then
        
    fakturasum = oRec6("beloeb")

    end if 
    oRec6.close

    if fakturasum <> 0 then
    fakturasumTxt = formatnumber(fakturasum, 2)
    fakturasumExp = formatnumber(fakturasum, 2)
    else
    fakturasumTxt = ""
    fakturasumExp = 0
    end if



    '*** Faktureret Omregnes til DKK FINANSÅR
    strSQLperFak = " AND ((brugfakdatolabel = 0 AND f.fakdato BETWEEN '"& sqlDatostartFY &"' AND '"& sqlDatoslutFY &"')"
    strSQLperFak = strSQLperFak &" OR (brugfakdatolabel = 1 AND f.labeldato BETWEEN '"& sqlDatostartFY &"' AND '"& sqlDatoslutFY &"'))"

    fakturasumFY = 0
    strSQLfaktureret = "SELECT SUM(IF(faktype = 0, f.beloeb * (f.kurs/100), f.beloeb * -1 * (f.kurs/100))) AS beloeb FROM fakturaer AS f WHERE "_
    &" shadowcopy <> 1 AND jobid = "& oRec("jobid") &" "& strSQLperFak &" GROUP BY jobid"
   
    'if session("mid") = 1 then
    '    Response.write strSQLfaktureret
    '    Response.flush
    'end if   
        
    oRec6.open strSQLfaktureret, oConn, 3
    if not oRec6.EOF then
        
    fakturasumFY = oRec6("beloeb")

    end if 
    oRec6.close

    if fakturasumFY <> 0 then
    fakturasumFYTxt = formatnumber(fakturasumFY, 2)
    fakturasumFYExp = formatnumber(fakturasumFY, 2)
    else
    fakturasumFYTxt = ""
    fakturasumFYExp = 0
    end if

    fakturasumForFY = formatnumber((fakturasum - fakturasumFY), 2)
    if fakturasumForFY <> 0 then
    fakturasumForFYTxt = formatnumber(fakturasumForFY, 2)
    fakturasumForFYExp = formatnumber(fakturasumForFY, 2)
    else
    fakturasumForFYTxt = ""
    fakturasumForFYExp = 0
    end if
    

    if cint(directexcel) <> 1 then
    strJob = strJob & "<td valign=top align=right style='padding:2px 5px 2px 2px; border-right:1px #cccccc solid; white-space:nowrap;' class=lille>" & stadesumTxt &"</td>"
    end if

    strExport = strExport & jo_bruttooms &";"& formatdatetime(oRec("jobstartdato"), 2) &";"& formatdatetime(oRec("jobslutdato"), 2) &";"& formatdatetime(oRec("lukkedato"), 2) &";"& sandsynlighedTxt &";"& pipelinevalue & ";" & stadesumExp &";"
    
     
    '*** Stade / fakturering likividitet 'ALTID DKK PR MD 
    '*** ALTID 12 md
    if seomsfor_antalmdHigend = 11 then
        for m = 0 to seomsfor_antalmdHigend
    
        stadesum_prmd(m) = 0
        strSQLmilepale = "SELECT SUM(belob) AS stadesum FROM milepale WHERE type = 1 AND jid = "& oRec("jobid") &" AND (MONTH(milepal_dato) = "& strMrdno(m) &" AND YEAR(milepal_dato) = "& strYearno(m) &") GROUP BY jid"
    
            'response.write strSQLmilepale & "<br>"
            'response.flush
        
        oRec6.open strSQLmilepale, oConn, 3
        if not oRec6.EOF then
        
        '** Omregn til basis valuta
        stadesum_prmd(m) = formatnumber(oRec6("stadesum")*(oRec("jo_valuta_kurs")/100), 2)
        'stadesum_prmd(m) = oRec6("stadesum")
        

        end if 
        oRec6.close

        
       stadesum_prmdGT(m) = stadesum_prmdGT(m) + stadesum_prmd(m)


       if cint(directexcel) <> 1 then

            if stadesum_prmd(m) <> 0 then
            stadesumTxt_prmd = formatnumber(stadesum_prmd(m), 2)
            else
            stadesumTxt_prmd = ""
            end if

        strJob = strJob & "<td valign=top align=right style='padding:2px 5px 2px 2px; border-right:1px #cccccc solid; white-space:nowrap;' class=lille>" & stadesumTxt_prmd &"</td>"
        end if

            if stadesum_prmd(m) <> 0 then
            stadesumTxt_prmdExp = formatnumber(stadesum_prmd(m), 2)
            else
            stadesumTxt_prmdExp = 0
            end if

        strExport = strExport & stadesumTxt_prmdExp & ";"
    
        next
    end if

   
    

    if cint(directexcel) <> 1 then
    strJob = strJob & "<td valign=top align=right style='padding:2px 5px 2px 2px; border-right:1px #cccccc solid; white-space:nowrap;' class=lille>" & fakturasumTxt &"</td>"
    strJob = strJob & "<td valign=top align=right style='padding:2px 5px 2px 2px; border-right:1px #cccccc solid; white-space:nowrap;' class=lille>" & fakturasumFYTxt &"</td>"
    strJob = strJob & "<td valign=top align=right style='padding:2px 5px 2px 2px; border-right:1px #cccccc solid; white-space:nowrap;' class=lille>" & fakturasumForFYTxt &"</td>"
    end if

    strExport = strExport & fakturasumExp & ";" & fakturasumFYExp & ";" & fakturasumForFYExp &";"

    for m = 1 to 5

    if oRec("jobans"&m) <> 0 then 

    if len(trim(oRec("m"&m&"init"))) <> 0 then
    uNit5 = oRec("m"&m&"init")
    else
    uNit5 = "("& oRec("m"&m&"mnr") &")"
    end if

    if cint(directexcel) <> 1 then
    strJob = strJob &"<td valign=top style='padding:2px 2px 2px 2px;' class=lille>"& uNit5 &"</td>"
    end if

    jobans5proc = oRec("jobans_proc_"&m)
    if jobans5proc <> 0 then
    jobans5val = pipelinevalue * (jobans5proc/100)
    else
    jobans5val = 0
    end if

    if cint(directexcel) <> 1 then
        strJob = strJob &"<td valign=top align=right style='padding:2px 10px 2px 2px;' class=lille>"& oRec("jobans_proc_"&m) &"%</td>"_
        &"<td valign=top align=right style='padding:2px 10px 2px 2px; border-right:1px #cccccc solid;' class=lille>"& formatnumber(jobans5val, 2) &"</td>"
    end if

    strExport = strExport & uNit5 &";"& jobans5proc &";"& jobans5val & ";"

    else
        
        if cint(directexcel) <> 1 then
        strJob = strJob &"<td colspan=3 style=""border-right:1px #cccccc solid;"">&nbsp;</td>"
        end if

        strExport = strExport & ";;;"
    end if


    next

    if cint(showSalgsAnv) = 1 then

    for s = 1 To 5

    

                 if oRec("salgsans"&s) <> 0 then 

                    if len(trim(oRec("s"&s&"init"))) <> 0 then
                    sNit = oRec("s"&s&"init")
                    else
                    sNit = "("& oRec("s"&s&"mnr") &")"
                    end if

                    if cint(directexcel) <> 1 then
                    strJob = strJob &"<td valign=top style='padding:2px 2px 2px 2px;' class=lille>"& sNit &"</td>"
                    end if

                    salgsansProc = oRec("salgsans"&s&"_proc")
                    if salgsansProc <> 0 then
                    salgsansProcval = pipelinevalue * (salgsansProc/100)
                    else
                    salgsansProcval = 0
                    end if

                    if cint(directexcel) <> 1 then
                    strJob = strJob &"<td valign=top align=right style='padding:2px 10px 2px 2px;' class=lille>"& oRec("salgsans"&s&"_proc") &"%</td>"_
                    &"<td valign=top align=right style='padding:2px 10px 2px 2px; border-right:1px #cccccc solid;' class=lille>"& formatnumber(salgsansProcval, 2) &"</td>"
                    end if

                    strExport = strExport & sNit &";"& salgsansProc &";"& salgsansProcval & ";"

                    else
                        if cint(directexcel) <> 1 then
                        strJob = strJob &"<td colspan=3 style=""border-right:1px #cccccc solid;"">&nbsp;</td>"
                        end if
                        strExport = strExport & ";;;"
                    end if


     Next

     end if


    '*** Fomr ****'
    fomrTxt = ""
    fomrTxtTD = ""
    'strSQLfaktureret = "SELECT SUM(beloeb) AS fakturasum FROM fakturaer WHERE faktype = 1 AND shadowcopy <> 1 AND jobid = "& oRec("jobid") &" GROUP BY jobid"
    strSQLfomr = "SELECT for_fomr, for_jobid, fomr.business_unit FROM fomr_rel "_
    &"LEFT JOIN fomr ON (fomr.id = for_fomr) WHERE for_jobid = " & oRec("jobid") & " GROUP BY business_unit"
    
    'response.write strSQLfomr
    'response.flush
        
    oRec6.open strSQLfomr, oConn, 3
    if not oRec6.EOF then
        
    fomrTxt = fomrTxt & oRec6("business_unit") &";"
    fomrTxtTD = fomrTxtTD & oRec6("business_unit") &"<br>"

    end if 
    oRec6.close

    if cint(directexcel) <> 1 then
    strJob = strJob &"<td style=""border-right:1px #cccccc solid; vertical-align:top;"">"& fomrTxtTD &"</td>"
    end if
    strExport = strExport & fomrTxt
    
    if cint(directexcel) <> 1 then
    'strJob = strJob &"<td bgcolor="& col &" style='padding:2px 10px 2px 2px;'>&nbsp;</td></tr>"
    strJob = strJob &"</tr>"
    end if

    if cint(directexcel) <> 1 then

    if media <> "print" then
    strJob = strJob  &"<tr><td colspan=100 bgcolor=""#cccccc"" style=""height:1px;""><img src=""../ill/blank.gif"" width=""30"" height=""1"" alt="""" border=""0"" style=""border:0px;""></td></tr>"
    end if

    end if

    strExport = strExport & "xx99123sy#z"
	
    lastMonth = datepart("m", oRec("jobstartdato"),2,2) 


	
	thispipeline = formatnumber(pipelinevalue/kvotient,0) '(thisbudget/(thisdays+1))

    thisheight = thispipeline 'formatnumber(((thisbudget/(thisdays+1))/kvotient),0)
	if thisheight <> 0 then
	thisheight = thisheight
	else
	thisheight = 0
	end if
	%>
	
	<!-- januar -->
	<%

    'for m = 0 to 11

   

	'if (cdate(oRec("jobstartdato")) <= cdate(""& strDayno(m) &"/"& strMrdno(m)  &"/"& strYearno(m))) AND (cdate(oRec("jobstartdato")) >= cdate("1/"& strMrdno(m) &"/"&strYearno(m))) then
	'tableHTML(m) = tableHTML(m) & "<tr><td><div id='j"&y&"' name='j"&y&"' style='position:relative;visibility:visible; background-color:"& col &"; border:1px "& col &" dashed;  height="&thisheight&" z-index:1000;'><img src='ill/blank.gif' width='30' height="&thisheight&" alt='"& oRec("jobnr") &"&nbsp;" & oRec("jobnavn") &"  " &  formatcurrency(thispipeline,2) &"' border='0'></div></td></tr>"
	'mdIdTot(m) = mdIdTot(m) + cdbl(pipelinevalue)
	'else
	'Response.write "<div id='j"&y&"' name='j"&y&"' style='position:relative;visibility:hidden;  height:0; width:0; z-index:1000; display:none;'></div>"
	'end if
    
    'next
    %>
	

	
	
	
	<%
    mdIdTot(datepart("m",oRec("jobstartdato"),2,2)-1) = mdIdTot(datepart("m",oRec("jobstartdato"),2,2)-1) + cdbl(pipelinevalue)    
    mdIdBruttoOmsTot(datepart("m",oRec("jobstartdato"),2,2)-1) = mdIdBruttoOmsTot(datepart("m",oRec("jobstartdato"),2,2)-1) + cdbl(jo_bruttooms)

    mdIdFakturaTot(datepart("m",oRec("jobstartdato"),2,2)-1) = mdIdFakturaTot(datepart("m",oRec("jobstartdato"),2,2)-1) + cdbl(fakturasum)
    mdIdStadeTot(datepart("m",oRec("jobstartdato"),2,2)-1) = mdIdStadeTot(datepart("m",oRec("jobstartdato"),2,2)-1) + cdbl(stadesum)

    mdTot = mdTot + cdbl(pipelinevalue)
    mdTotGT = mdTotGT + cdbl(pipelinevalue)
    mdBruttoOmsTot = mdBruttoOmsTot + cdbl(jo_bruttooms)
    mdBruttoOmsTotGT = mdBruttoOmsTotGT + cdbl(jo_bruttooms)
    lastMthName = monthname(datepart("m",oRec("jobstartdato"),2,2)) &" "& datepart("yyyy",oRec("jobstartdato"), 2,2)
	y = y + 1
	end if
	
	oRec.movenext
	wend
	oRec.close
	
  
    if cint(directexcel) <> 1 then

        if y > 1 then
        strJob = strJob & "<tr bgcolor='#FFFFFF'><td style='padding:2px 10px 2px 2px; border-bottom:1px #cccccc solid;' class=lille>Total:</td>"
        strJob = strJob & "<td align=right style='padding:2px 10px 2px 2px; border-bottom:1px #cccccc solid;' class=lille><b>"& formatnumber(mdBruttoOmsTot, 2) &"</b></td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>"
        strJob = strJob & "<td align=right style='padding:2px 10px 2px 2px; border-bottom:1px #cccccc solid;' class=lille><b>"& formatnumber(mdTot, 2) &"</b></td>"
        
        if seomsfor_antalmdHigend = 11 then
        for m = 0 TO seomsfor_antalmdHigend

         strJob = strJob & "<td style='padding:3px; border-bottom:1px #cccccc solid;'>&nbsp;</td>"

        next
        end if

        strJob = strJob & "<td colspan="& cps &" style='padding:3px; border-bottom:1px #cccccc solid;'>&nbsp;</td></tr>"
    
        strJob = strJob & "<tr bgcolor='lightpink'><td style='padding:2px 10px 2px 2px;'>Grandtotal:</td>"
        strJob = strJob & "<td align=right style='padding:2px 10px 2px 2px;'><b>"& formatnumber(mdBruttoOmsTotGT, 2) &"</b></td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>"
        strJob = strJob & "<td align=right style='padding:2px 10px 2px 2px;'><b>"& formatnumber(mdTotGT, 2) &"</b></td><td><b>"& formatnumber(stadesum_prGT,2) &"</b></td>"
        
        if seomsfor_antalmdHigend = 11 then
            for m = 0 TO seomsfor_antalmdHigend

             strJob = strJob & "<td align=right style='padding:2px 10px 2px 2px; font-size:10px;'>"& left(monthname(strMrdno(m)), 3) &" "& right(strYearno(m), 2) &"<br>"& formatnumber(stadesum_prmdGT(m), 2) &"</td>"

            next
        end if
        
        
        strJob = strJob & "<td colspan="& cps &" style='padding:3px; border-bottom:1px #cccccc solid;'>&nbsp;</td></tr>"
        end if


	
	'strBudget = strBudget & "</table>" 
	strJob = strJob & "</table>" 
    end if
    %>
    
  
  
	
	<% '*** job table ***
    if media <> "export" AND cint(directexcel) <> 1 then

	if media <> "print" then
	bgcoltable = "#d6dff5"
	else
	bgcoltable = "#ffffff"
	end if
    %>
    <!-- Job Pipeline Table -->
    <table cellspacing="1" cellpadding="0" border="0" width=100% bgcolor="<%=bgcoltable%>">
	<tr><td valign="top"><%=strJob%></td>
	</tr>
	</table>

         <br /><br /><br /><br /><br /><br />


    <%if cint(seomsfor_antalmdHigend) = 11 AND cint(ign_per) <> 1 then %>
    
    <!--- BAR CHART --->
    <div style="padding:20px; border:10px #CCCCCC solid; background-color:#FFFFFF;">
    <table cellspacing="1" cellpadding="0" border="0">
	<tr>
     <td>Måned:</td>
    <%
    for m = 0 to 11
    %>
    
    <td colspan="4" style="padding:2px;" valign=bottom align="center"><b><%=left(strMrdnm(m), 3) &" "& right(strYearno(m), 2) %></b></td>
    <%
    next

    %>
    </tr>
        <tr><td>&nbsp;</td>
    <%
     
    for m = 0 to 11

        if mdIdBruttoOmsTot(m) <> 0 then
            mdIdBruttoOmsTotHgt = formatnumber((mdIdBruttoOmsTot(m)/mdpeer) * 100, 0)
            'if mdIdBruttoOmsTotHgt > 100 then
            '    mdIdBruttoOmsTotHgt = 100 - mdIdBruttoOmsTotHgt
            'end if
        else
        mdIdBruttoOmsTotHgt = 0
        end if

        if mdIdBruttoOmsTotHgt > 100 then
        mdIdBruttoOmsTotHgt = 100
        else
        mdIdBruttoOmsTotHgt = mdIdBruttoOmsTotHgt
        end if

        'if mdIdBruttoOmsTotHgt > 99 then
        'bgthisDv = "green"
        'end if

        'if mdIdBruttoOmsTotHgt >= 70 AND mdIdBruttoOmsTotHgt < 99 then
        'bgthisDv = "#5582d2"
        'end if 

        'if mdIdBruttoOmsTotHgt >= 30 AND mdIdBruttoOmsTotHgt < 70 then
        bgthisDv = "#d6dff5"
        'end if

        'if mdIdBruttoOmsTotHgt >= 10 AND mdIdBruttoOmsTotHgt < 30 then
        'bgthisDv = "orange"
        'end if

        'if mdIdBruttoOmsTotHgt < 10 then
        'bgthisDv = "darkred"
        'end if


        if mdIdFakturaTot(m) <> 0 then
        mdIdFakturaTotHgt = formatnumber((mdIdFakturaTot(m)/mdpeer) * 100, 0)
        else
        mdIdFakturaTotHgt = 0
        end if

        if mdIdFakturaTotHgt > 100 then
        mdIdFakturaTotHgt = 100
        else
        mdIdFakturaTotHgt = mdIdFakturaTotHgt
        end if

        
        bgFakDv = "yellowgreen"
       

     
        if mdIdstadeTot(m) <> 0 then
        mdIdstadeTotHgt = formatnumber((mdIdstadeTot(m)/mdpeer) * 100, 0)
        else
        mdIdstadeTotHgt = 0
        end if

        if mdIdstadeTotHgt > 100 then
        mdIdstadeTotHgt = 100
        else
        mdIdstadeTotHgt = mdIdstadeTotHgt
        end if

        
        bgStadeDv = "#CCCCCC"

        
        if mdIdBruttoOmsTotHgt >= 100 then
        mdIdBruttoOmsTotHgtBdrTop = "10"
        else
        mdIdBruttoOmsTotHgtBdrTop = "0"
        end if

        if mdIdStadeTotHgt >= 100 then
        mdIdStadeTotHgtBdrTop = "10"
        else
        mdIdStadeTotHgtBdrTop = "0"
        end if

        if mdIdFakturaTotHgt >= 100 then
        mdIdFakturaTotHgtBdrTop = "10"
        else
        mdIdFakturaTotHgtBdrTop = "0"
        end if
    %>
    
    <td style="padding:0px; border:1px #CCCCCC solid;" valign=bottom>
        <div style="height:<%=mdIdBruttoOmsTotHgt*2%>px; border-top:<%=mdIdBruttoOmsTotHgtBdrTop%>px #5582d2 solid; width:20px; vertical-align:bottom; background-color:<%=bgthisDv%>;">&nbsp;</div>
      </td>
     <td style="padding:0px; border:1px #CCCCCC solid;" valign=bottom>
         <div style="height:<%=mdIdstadeTotHgt*2%>px; border-top:<%=mdIdStadeTotHgtBdrTop%>px #999999 solid; width:20px; vertical-align:bottom; background-color:<%=bgStadeDv%>;">&nbsp;</div>
      </td>
       <td style="padding:0px; border:1px #CCCCCC solid;" valign=bottom>
         <div style="height:<%=mdIdFakturaTotHgt*2%>px; border-top:<%=mdIdFakturaTotHgtBdrTop%>px green solid; width:20px; vertical-align:bottom; background-color:<%=bgFakDv%>;">&nbsp;</div>
       </td>
         <td style="padding:0px; border:0px #CCCCCC solid; width:10px;">&nbsp;</td>
    <%
    next
     %>
    </tr>
        <tr><td>&nbsp;</td>
    <%

  

    for m = 0 to seomsfor_antalmdHigend
    %>
    
    <td colspan="4" style="padding:2px; font-size:9px; text-align:right;" valign=top>
        <%if mdIdBruttoOmsTot(m) <> 0 then %>
        Bgt.: <%=formatnumber(mdIdBruttoOmsTot(m), 0) %><br />
        <%end if %>

         <%if mdIdFakturaTot(m) <> 0 then %>
        Pro.:<%=formatnumber(mdIdFakturaTot(m), 0) %><br />
        <%end if %>

          <%if mdIdStadeTot(m) <> 0 then %>
        Inv.:<%=formatnumber(mdIdStadeTot(m), 0) %><br />
        <%end if %>

        
    </td>
    <%
    next

   

    %>

    </tr>
        </table><br><br />
	<font class=megetlillesort>Antal job og tilbud: <%=y-2%> <!-- Kvo. <%=kvotient%>,  -->
        <br />
        Progress stated og invoiced er fordelt udfra de måneder hvor jobbbet har startdato, så der kan sammenlignes med budget. 
        </font>

    <table>
        <tr>
    <td style="padding:0px; border:1px #CCCCCC solid;" valign=bottom>
        <div style="height:20px; width:20px; vertical-align:bottom; background-color:<%=bgthisDv%>;">Bgt.</div>
      </td>
     <td style="padding:0px; border:1px #CCCCCC solid;" valign=bottom>
         <div style="height:20px; width:20px; vertical-align:bottom; background-color:<%=bgStadeDv%>;">Pro.</div>
      </td>
       <td style="padding:0px; border:1px #CCCCCC solid;" valign=bottom>
         <div style="height:20px; width:20px; vertical-align:bottom; background-color:<%=bgFakDv%>;">Inv.</div>
       </td>
        </tr>


        </table>
    </div>
    <%else %>
         <br /><br />
         <font class=megetlillesort>Antal job og tilbud: <%=y-2%> <!-- Kvo. <%=kvotient%>,  -->
	<%end if 'antalmd + ign_per %>
    
    </td>
    </tr>
    </table>
    
	
     </div><!-- table div -->
	
	
	<br><br>
	<br>
	<%
    
    
    if media <> "print" then%>
	<br /><a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br><br>
    <!--
	<div style="position:relative; width:420px; background-color:#eff3ff; top:0px; left:0px; visibility:visible; border:1px red dashed; border-right:2px red dashed; border-bottom:2px red dashed; padding:15px;"><img src="../ill/ac0005-24.gif" width="24" height="24" alt="" border="0">&nbsp;<b>Side note(r):<br></b>
		Klik på jobnavnet til venstre for at for vist hvilke blokke er hører til dette job.<br>
		Beløb bliver beregnet udfra det angivede budget / antal måneder jobbet løber over.<br>
		Hold musen over de markedrede blokke for at få vis den månedlige forventede omsætning på dette job.

        Jobstartdato,
        Jobansvarlige 1-5
        Jobstatus
        Pipelineværdi



		</div>
        -->
		<br><br>&nbsp;
	<%end if

    end if 'export



    '******************* Eksport **************************' 

if media = "export" OR cint(directexcel) = 1 then

    
  call TimeOutVersion() 

  
    ekspTxt = strExport
	ekspTxt = replace(ekspTxt, "xx99123sy#z", vbcrlf)
	
    datointerval = request("datointerval")
	
	
	filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
				Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\pipeline.asp" then
					Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\pipelineexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\pipelineexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				else
					Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\pipelineexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\pipelineexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				end if
				
				
				
				file = "pipelineexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
				
				
				'**** Eksport fil, kolonne overskrifter ***
				
				
				strOskrifter = "Job;Job Nr;Jobstatus;Kontakt;Knr;Brutto Oms "& basisValISO &";Startdato;Slutdato;Lukkedato;Sandsynlighed %;Pipelineværdi "& basisValISO &";Faktureringsplan (stade) "& basisValISO &";"
        
                  if seomsfor_antalmdHigend = 11 then
                      for m = 0 to seomsfor_antalmdHigend

                        strOskrifter = strOskrifter &" "& left(monthname(strMrdno(m)), 3) &" - "& strYearno(m) &";"

                      next
                   end if

        
                strOskrifter = strOskrifter & "Faktureret "& basisValISO &";Faktureret FY "& basisValISO &";Faktureret før FY "& basisValISO &";Jobansv.;Andel %;Værdi "& basisValISO &";"_
                & "Jobejer.;Andel %;Værdi "& basisValISO &";Jobmedansv1.;Andel %;Værdi "& basisValISO &";Jobmedansv2.;Andel %;Værdi "& basisValISO &";Jobmedansv3.;Andel %;Værdi "& basisValISO &";"

                
				if cint(showSalgsAnv) = 1 then

                for s = 1 to 5
                    strOskrifter = strOskrifter & "Salgsansv. "&s&";Andel %;Værdi "& basisValISO &";"
                next

                end if
				
                strOskrifter = strOskrifter &"Fomr;Fomr;Fomr;Fomr;Fomr;"
				
				
				if cint(ign_per) <> 1 then
				objF.writeLine("Periode afgrænsning: "& monthname(strMrdno(0)) &" "& strYearno(0) &" + "& seomsfor_antalmd &" md "& vbcrlf)
                end if
				
                objF.WriteLine(strOskrifter & chr(013))
				objF.WriteLine(ekspTxt)
				objF.close
				
				%>
				
	            <table border=0 cellspacing=1 cellpadding=0 width="200">
	            <tr><td valign=top bgcolor="#ffffff" style="padding:5px;">
	            <img src="../ill/outzource_logo_200.gif" />
	            </td>
	            </tr>
                <%if cint(directexcel) = 1 then%>
                    
                    <tr>
	            <td valign=top bgcolor="#ffffff" style="padding:5px 5px 5px 15px;">
	            <a href="../inc/log/data/<%=file%>" class=vmenu target="_blank">Din CSV. fil er klar >></a>
	            </td></tr>
                    
                <%else%>
	            <tr>
	            <td valign=top bgcolor="#ffffff" style="padding:5px 5px 5px 15px;">
	            <a href="../inc/log/data/<%=file%>" class=vmenu target="_blank" onClick="Javascript:window.close()">Din CSV. fil er klar >></a>
	            </td></tr>
                <%end if %>
	            </table>
    
	            
	          
	            
	            <%
                
                
                Response.end
	            'Response.redirect "../inc/log/data/"& file &""	
				



end if






    if media <> "print" then

ptop = 0
pleft = 840
pwdt = 140

call eksportogprint(ptop,pleft,pwdt)

lnk = "&FM_kunde="&kundeid&"&FM_job="&jobid&"&FM_medarb="&thisMiduse&"&FM_medarb_hidden="&thisMiduse&"&FM_progrp="&progrp&"&seomsfor="&seomsfor&"&FM_start_mrd="&strMrd&""_
&"&FM_jobsog="&jobSogVal&"&viskunabnejob0="&viskunabnejob0&"&viskunabnejob1="&viskunabnejob1&"&viskunabnejob2="&viskunabnejob2&""_
&"&FM_kundejobans_ell_alle="&visKundejobans&"&FM_kundeans="&kundeans&"&FM_jobans="&jobans&"&FM_jobans2="&jobans2&"&FM_jobans3="&jobans3&"&FM_segment="&segment&"&FM_ign_per="&ign_per
%>

        
      <tr>
        <td align=center>
        <a href="#" onclick="Javascript:window.open('pipeline.asp?media=export<%=lnk%>', '', 'width=350,height=120,resizable=no,scrollbars=no')" class=vmenu><img src="../ill/export1.png" border=0></a>
        </td><td><a href="#" onclick="Javascript:window.open('pipeline.asp?media=export<%=lnk%>', '', 'width=350,height=120,resizable=no,scrollbars=no')" class=vmenu>.csv fil eksport</a></td>
       </tr>
    <tr>
   <td align=center><a href="pipeline.asp?media=print<%=lnk%>" target="_blank"  class='vmenu'>
   &nbsp;<img src="../ill/printer3.png" border=0 alt="" /></a>
    </td><td><a href="pipeline.asp?media=print<%=lnk%>" target="_blank" class="vmenu">Print version</a></td>
   </tr>
   
    <tr><td colspan="2"><br /><br />
        <%
            
                nWdt = 120
                nTxt = "Opret nyt job"
                nLnk = "jobs.asp?menu=job&func=opret&id=0&int=1&rdir=pipe"
                nTgt = ""
                call opretNy_2013(nWdt, nTxt, nLnk, nTgt) %>    
            
        

    </td></tr>
	
   </table>
</div>
<%else%>

<% 
Response.Write("<script language=""JavaScript"">window.print();</script>")
%>
<%end if%>



	<br>
	<br>
	
	&nbsp;
	
	
	</div>
	


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
