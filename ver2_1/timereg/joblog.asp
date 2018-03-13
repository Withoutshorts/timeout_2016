<%Response.Buffer = true %>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->
<!--#include file="inc/isint_func.asp"-->
<!--#include file="inc/joblog_inc.asp"-->
<!--#include file="inc/convertDate.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->

<%
if session("user") = "" then
%>
<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->  
<%
	errortype = 5
	call showError(errortype)
	else
	
	call smileyAfslutSettings()
	
	
    '**************************************************'
	'***** Faste Filter kriterier *********************'
	'**************************************************'
		
	
		
	'*** Job og Kundeans ***
	call kundeogjobans()
	
	'*** Medarbejdere / projektgrupper '
    selmedarb = session("mid")
	
	'call medarbogprogrp("oms")
	medarbSQlKri = ""
	kundeAnsSQLKri = ""
	jobAns2SQLkri = ""
	jobAnsSQLkri = ""
	ugeAflsMidKri = ""
	'fakmedspecSQLkri = ""
	
	if len(trim(request("FM_progrp"))) <> 0 then
	progrp = request("FM_progrp")
	else
    progrp = 0
	end if

    if len(trim(request("cur"))) <> 0 then
	cur = request("cur")
	else
    cur = 0
	end if

    if len(trim(request("ver"))) <> 0 then
	ver = request("ver")
	else
    ver = 0
	end if
    

	'Response.Write "medid first: "& left(request("FM_medarb"), 1)
    'Response.Write request("FM_medarb_hidden")
	'Response.end
	
	'*** Rettigheder p� den der er logget ind **'
	medarbid = session("mid")
	
	
    if len(trim(request("nomenu"))) <> 0 AND request("nomenu") <> "0" then
    nomenu = 1
    else 
	nomenu = 0
    end if

    if len(request("FM_medarb")) <> 0 then 'OR func = "export"
	
	    if left(request("FM_medarb"), 1) = "0" then
	        
            if media <> "print" then
	        thisMiduse = replace(request("FM_medarb_hidden"), "-", "0")
    	    else
    	    thisMiduse = replace(request("FM_medarb"), "-", "0")
    	    end if
    	
    	intMids = split(thisMiduse, ", ")
	    else
	    thisMiduse = replace(request("FM_medarb"), "-", "0")
	    intMids = split(thisMiduse, ", ")
	    end if
	
	else
	    
	    thisMiduse = session("mid") 
	    intMids = split(thisMiduse, ", ")
	   
	end if
	
	
	media = request("media")
	
	'**** Kundekri ***
	if len(request("FM_kunde")) <> 0 then
	kundeid = request("FM_kunde")
	else
	kundeid = 0
	end if

    
	
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
			     ugeAflsMidKri = " u.mid = "& intMids(m)
			     else
			     medarbSQlKri = medarbSQlKri & " OR m.mid = " & intMids(m)
			     kundeAnsSQLKri = kundeAnsSQLKri & " OR kundeans1 = "& intMids(m) &" OR kundeans2 = "& intMids(m) &""
			     jobAnsSQLkri = jobAnsSQLkri & " OR jobans1 = "& intMids(m)  
			     jobAns2SQLkri = jobAns2SQLkri & " OR jobans2 = "& intMids(m)
                 jobAns3SQLkri = jobAns3SQLkri & " OR jobans3 = "& intMids(m) & " OR jobans4 = "& intMids(m) & " OR jobans5 = "& intMids(m)
			     ugeAflsMidKri = ugeAflsMidKri &  " OR u.mid = "& intMids(m)
			     end if
			     
			    next
			    
			    medarbSQlKri = medarbSQlKri & ")"
			    
		    jobAnsSQLkri =  " ("& jobAnsSQLkri &")"
			jobAns2SQLkri =  "xx (" & jobAns2SQLkri &")"
			jobAns3SQLkri =  "xx (" & jobAns3SQLkri &")"
			ugeAflsMidKri = " ("& ugeAflsMidKri &")"
			
	
	'Response.Write "ugeAflsMidKri" & ugeAflsMidKri
	'Response.end
	
	
	'** Er key acc og kundeansvarlig valgt **?
	if cint(kundeans) = 1 then
	kundeAnsSQLKri = kundeAnsSQLKri
	else
	kundeAnsSQLKri = " Kundeans1 <> -1 AND kundeans2 <> -1 "
	end if
	
	'***** Valgt job eller s�gt p� Job ****
	'** hvis Sog = Ja
	call jobsog()

	
	'*** Aftale ****
	if len(request("FM_aftaler")) <> 0 then ' AND jobid <= 0 AND len(jobSogVal) = 0 then
	    aftaleid = request("FM_aftaler")
	else
        aftaleid = 0
	end if
	
	
	
	
	
	
	
	
	if len(request("FM_akttype")) <> 0 AND len(request("FM_kunde")) <> 0 then 
	akttype_sel = request("FM_akttype")
	else
	    if request.Cookies("stat")("fm_akttype_sel") <> "" then
	    akttype_sel = request.Cookies("stat")("fm_akttype_sel") 
        'akttype_sel = akttype_sel & "#1#, #2#, "
	    else
	    'call akttyper2009(6)
	    'akttype_sel = aktiveTyper
	    'akttype_sel = "#-99#, #-1#, #-2#, #-3#, #-4#, #-5#, "
        akttype_sel = akttype_sel & "#1#, #2#, "
	    end if
	end if
	
    'akttype_sel = "#1#, #2#, "
	akttype_sel = replace(akttype_sel, "#-99#, ", "")
	vartyper = akttype_sel
	response.Cookies("stat")("fm_akttype_sel") = akttype_sel
	response.cookies("stat").expires = date + 80
	
	
	'Response.Write "akttype_sel" & akttype_sel

	'Response.Write typerSQL
	
	
	'**** Alle SQL kri starter med NUL records ****'
	jobidFakSQLkri = " OR jobid = -1 "
	jobnrSQLkri = " OR tjobnr = '-1' "
	jidSQLkri = " OR id = -1 "
	seridFakSQLkri = " OR aftaleid = -1 "

   

	    '**** Skjul timepriser ****'
	     if level <= 2 OR level = 6 OR (lto = "oko" AND level = 1) then
	    
	        
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

        
                '** DEFAULT ***
		        select case lto 
                case "glad", "lyng", "oko"
                hidetimepriser = 1
		        hidetpCHK = "CHECKED"
                case else
                hidetimepriser = 0
		        hidetpCHK = ""
                end select    
		        
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
		       

                '** DEFAULT ***
		        select case lto 
                case "glad", "lyng"
                hideenheder = 1
		        hideehCHK = "CHECKED"
                case else
                hideenheder = 1
		        hideehCHK = "CHECKED"
                end select    

		        
		        end if
		        
		    end if
		
		
		    response.cookies("stat")("hideenheder") = hideenheder
		    
		    '***'
		    
		    
		    '**** Benyt Enheder **'
	        if len(trim(request("joblog_uge"))) <> 0 then
		        if len(trim(request("useenheder"))) <> 0 then    
		        useenheder = 1
		        useehCHK = "CHECKED"
		        else
		        useenheder = 0
		        useehCHK = ""
		        end if
		    else
		        if len(trim(request.cookies("stat")("useenheder"))) <> 0 then
		        
		            useenheder = request.cookies("stat")("useenheder")
		            if useenheder = 1 then
		            useehCHK = "CHECKED"
		            else
		            useehCHK = ""
		            end if
    		        
		        else

		        useenheder = 0
		        useehCHK = ""
		        
		        end if
		        
		    end if
		
		
		    response.cookies("stat")("useenheder") = useenheder
		    '***'
		    
		    
		    '**** Vis kostpriser **'
	        if len(trim(request("joblog_uge"))) <> 0 then
		        if len(trim(request("visKost"))) <> 0 then    
		        visKost = 1
		        visKostCHK = "CHECKED"
		        else
		        visKost = 0
		        visKostCHK = ""
		        end if
		    else
		        if len(trim(request.cookies("stat")("visKost"))) <> 0 AND level = 1 then
		        
		            visKost = request.cookies("stat")("visKost")
		            if visKost = 1 then
		            visKostCHK = "CHECKED"
		            else
		            visKostCHK = ""
		            end if
    		        
		        else

    
    
                 '** DEFAULT ***
		        select case lto 
                case "glad", "lyng"
                visKost = 0
		        visKostCHK = ""
                case else
                visKost = 0
		        visKostCHK = ""
                end select                


		        
		        
		        end if
		        
		    end if
		
		
		    response.cookies("stat")("useenheder") = useenheder
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
		       

                    
                '** DEFAULT ***
		        select case lto 
                case "glad", "lyng"
                hidegkfakstat = 1
		        hidegkCHK = "CHECKED"
                case else
                hidegkfakstat = 1
		        hidegkCHK = "CHECKED"
                end select            

                

		        
		        end if
		        
		    end if
		
		
		    response.cookies("stat")("hidegkfakstat") = hidegkfakstat
	
	        '***'  
	        
            
            '*** Skjul fase kolonne
            if len(trim(request("joblog_uge"))) <> 0 then
		        if len(trim(request("hidefase"))) <> 0 then    
		        hidefase = 1
		        hideFASCHK = "CHECKED"
		        else
		        hidefase = 0
		        hideFASCHK = ""
		        end if
		    else
		        if len(trim(request.cookies("stat")("hidefase"))) <> 0 then
		        
		            hidefase = request.cookies("stat")("hidefase")
		            if hidefase = 1 then
		            hidefase = 1
                    hideFASCHK = "CHECKED"
		            else
		            hideFASCHK = ""
		            hidefase = 0
                    end if
    		        
		        else
		        

                '** DEFAULT ***
		        select case lto 
                case "glad", "lyng"
                hidefase = 1
		        hideFASCHK = "CHECKED"
                case else
                hidefase = 0
		        hideFASCHK = ""
                end select            


		        
		        end if
		        
		    end if
		
		
		    response.cookies("stat")("hidefase") = hidefase
            
                    

            '*** Show CPR ****
            if len(trim(request("joblog_uge"))) <> 0 then
		        if len(trim(request("showcpr"))) <> 0 then    
		        showcpr = 1
		        showcprCHK = "CHECKED"
		        else
		        showcpr = 0
		        showcprCHK = ""
		        end if
		    else
		        if len(trim(request.cookies("stat")("showcpr"))) <> 0 then
		        
		            showcpr = request.cookies("stat")("showcpr")
		            if showcpr = 1 then
		            showcprCHK = "CHECKED"
		            else
		            showcprCHK = ""
		            end if
    		        
		        else
		        showcpr = 0
		        showcprCHK = ""
		        
		        end if
		        
		    end if
		
		
		    response.cookies("stat")("showcpr") = showcpr
	        
	        
             '*** Show Forretningsomr�de ****
            if len(trim(request("joblog_uge"))) <> 0 then
		        if len(trim(request("showfor"))) <> 0 then    
		        showfor = 1
		        showforCHK = "CHECKED"
		        else
		        showfor = 0
		        showforCHK = ""
		        end if
		    else
		        if len(trim(request.cookies("stat")("showfor"))) <> 0 then
		        
		            showfor = request.cookies("stat")("showfor")
		            if showfor = 1 then
		            showforCHK = "CHECKED"
		            else
		            showforCHK = ""
		            end if
    		        
		        else
		       
		        
                '** DEFAULT ***
		        select case lto 
                case "glad", "lyng"
                showfor = 0
		        showforCHK = ""
                case else
                showfor = 0
		        showforCHK = ""
                end select

		        end if
		        
		    end if
		
		
		    response.cookies("stat")("showfor") = showfor
	        


	menu = request("menu")
	
	thisfile = "joblog"
	
	func = request("func")
	print = request("print")
	
	rdir = request("rdir")
	
	if print = "j" then
	

	    select case lto
        case "mmmi", "unik"
	    globalWdt = 815 '1160
	    case else
	    globalWdt = 1060
	    end select
	
    else
	
	    if rdir = "treg" OR rdir = "nwin" then
	    globalWdt = 915
	    else
	    globalWdt = 1160
	    end if
	
	end if

    if len(trim(request("brugmd_md"))) <> 0 then
    brugmd_md = request("brugmd_md")
    else
    brugmd_md = month(now)
    end if
	
    '** Uge eller m�ned forvalgt
	if request("bruguge") = "1" OR request("brugmd") = "1" then
	
                    if request("bruguge") = "1" then
                    bruguge = 1
	                brugugeCHK = "CHECKED"
	    
	                    bruguge_week = request("bruguge_week") 
	                    bruguge_year = request("bruguge_year") 
                        'brugmd_md = request("brugmd_md") 'bruges ikke til beregning da, der er valgt uge
	    
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
	    

                    else 'brugmd = 1

                
                    brugmd = 1
	                brugmdCHK = "CHECKED"
	    
	                    bruguge_week = request("bruguge_week") 'bruges ikke til beregning da, der er valgt md
	                    'brugmd_md = request("brugmd_md") 
                        bruguge_year = request("bruguge_year") 
	   
	                    stDato = "1/"& brugmd_md &"/"&bruguge_year
	                    tjkDato = stDato 

	                    tjkDato_slut = dateadd("m", 1, tjkDato)
                        tjkDato_slut = dateadd("d", -1, tjkDato_slut)

                   
	                    stDaguge = day(tjkDato)
	                    stMduge = month(tjkDato)
	                    stAaruge = year(tjkDato)
	    
	              
	    
	                    slDaguge = day(tjkDato_slut)
	                    slMduge = month(tjkDato_slut)
                        slAaruge = year(tjkDato_slut)
	                    
                        if cint(slMduge) = 2 then
    
                                select case right(slAaruge, 2) 
                                case "04","08", "12", "16", "20", "24", "28", "32", "36", "40", "44", "48", "52", "56", "60"


                                end select

    
                        end if 
        
    
    
                        
	                   
                

                    end if
	    
	else
	    
	    if rdir = "treg" AND len(trim(request("bruguge_week"))) = 0 then
	    
	    stMduge = request("FM_start_mrd")
	    
	    'Response.Write "stMduge =" & request("FM_start_mrd")
	    'Response.flush
	    
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
	  'uGodkendt = split(trim(request("FM_godkendt")), "#, ")
       
     uGodkendtDato = year(now) & "-" & month(now) & "-" & day(now) 
     editor = session("user")

     for u = 0 to UBOUND(ujid)
	
	
	'if trim(left(uGodkendt(u), 1)) = "1" then
	'uGodkendt(u) = trim(left(uGodkendt(u), 1))
	
	'else
	'uGodkendt(u) = 0
	'editor = ""
	'end if
	
				'intGodkendt = request("tid") 

                if len(trim(request("FM_godkendt_"& trim(ujid(u))))) <> 0 then
                uGodkendt = request("FM_godkendt_"& trim(ujid(u)))
                else
                uGodkendt = 0
                end if
              

				strSQL = "UPDATE timer SET godkendtstatus = "& uGodkendt &", "_
				&"godkendtstatusaf = '"& editor &"', godkendtdato = '"& uGodkendtDato &"' WHERE tid = " & ujid(u) & "  AND overfort = 0"
				
				'Response.write strSQL &"<br>"
				
				oConn.execute(strSQL)
				
	next

    'Response.end
		%>
		
			
	    
	    
	  
		
	<% 
	
	dontshowDD = 1%>
	<!--#include file="inc/weekselector_s.asp"--> 
	<%
	
	'*** LINK og faste VAR ****
	strLink = "joblog.asp?rdir="&rdir&"&menu="&menu&"&FM_medarb="&thisMiduse&"&FM_medarb_hidden="&thisMiduse&""_
	&"&lastFakdag="&lastFakdag&"&FM_job="&jobid&"&FM_start_dag="&strDag&""_
	&"&FM_start_mrd="&strMrd&"&FM_start_aar="&strAar&"&FM_slut_dag="&strDag_slut&""_
	&"&FM_slut_mrd="&strMrd_slut&"&FM_slut_aar="&strAar_slut&""_
	&"&FM_jobsog="&jobSogValPrint&""_
	&"&FM_aftaler="&aftaleid&""_
	&"&FM_radio_projgrp_medarb="&progrp_medarb&""_
	&"&FM_progrp="&progrp&""_
	&"&FM_kunde="&kundeid&""_
	&"&FM_kundeans="&kundeans&""_
	&"&FM_jobans="&jobans&""_
	&"&FM_kundejobans_ell_alle="&visKundejobans&""_
	&"&FM_akttype="&vartyper&""_
	&"&FM_orderby_medarb="&fordelpamedarb&""_
	&"&bruguge="&bruguge&"&bruguge_week="&bruguge_week&"&bruguge_year="&bruguge_year&""_
    &"&brugmd="&brugmd&"&brugmd_md="&brugmd_md&""_
    &"&viskunabnejob0="&viskunabnejob0&"&viskunabnejob1="&viskunabnejob1&"&viskunabnejob2="&viskunabnejob2&"&FM_segment="&segment&"&nomenu="&nomenu

	'Response.Write strLink
	'Response.end
	Response.redirect strLink
	
	
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
	  
          joblog_ugeCHK1 = ""
	    joblog_ugeCHK2 = ""
        joblog_ugeCHK3 = ""
        
        
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
		
		
			
			
			    select case cint(joblog_uge) 
                case 1 
			    joblog_ugeCHK1 = "CHECKED"
			   
			    joblogsubWZB = "visible"
			    joblogsubDSP = ""
			    case 3

                joblog_ugeCHK3 = "CHECKED"
			    
			    joblogsubWZB = "hidden"
			    joblogsubDSP = "none"

			    case else
			   
			    joblog_ugeCHK2 = "CHECKED"
			    
			    joblogsubWZB = "hidden"
			    joblogsubDSP = "none"
			    
			    end select
			
			response.Cookies("stat")("joblog_uge") = joblog_uge
	
	    end if 
	
	
	
	'**** Vis subtotaler pr akt ell. medarb. ***'
	select case cint(joblog_uge) 
    case 1 
	
	if len(trim(request("FM_orderby_medarb"))) <> 0 AND request("FM_orderby_medarb") <> "0" then
	fordelpamedarb = request("FM_orderby_medarb")
	    
	    if fordelpamedarb = 1 then
	    orderByKri = "Tjobnavn, Tjobnr, Tmnavn, Tdato "
	    else
	    orderByKri = "Tjobnavn, Tjobnr, fase, TAktivitetId, Tdato "
	    end if
	
	else
	fordelpamedarb = 0
	orderByKri = "Tdato, Tjobnavn, Tjobnr "
	end if
	
	
	orderByKri = orderByKri & " DESC" ', Tdato
	
    case 3

    orderByKri = "Tmnr, Tknr, Tjobnavn, Tjobnr, fase, a.sortorder, taktivitetnavn, tdato"

	case else
	
	'** Timeseddel altid pr. medarb, tdato ***'
	orderByKri = orderByKri & " Tmnr, Tdato DESC"
	komprimer = 1
	
	end select
	
	
	
    if request("print") = "j" then
            media = "print"
    end if



   if print <> "j" AND rdir <> "treg" AND rdir <> "nwin" AND media <> "export" AND nomenu <> "1" then
	pleft = 90
	ptop = 102 '202
	jobbgcol = "#ffffff"
    ov_top = 70
    pr_top = 0
    hlp_top = -16
	else
            if media = "export" then
	        pleft = 20
	        ptop = 10
	        jobbgcol = "#ffffff"
            ov_top = 20
            pr_top = 0
            hlp_top = -16
            else
            pleft = 20
	        ptop = -60
	        jobbgcol = "#ffffff"
            ov_top = 20
            pr_top = 0
            hlp_top = -16
            end if
	end if
	
	if print <> "j" AND rdir <> "treg" AND rdir <> "nwin" AND media <> "export" then
	%>
	
    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
     <script src="inc/joblog_jav.js"></script>
	
	
	
    <%if nomenu <> "1" then %>
   


    <%
        call menu_2014()
        
    end if'nomenu %>
	
	<%
	
	else
	
	level = session("rettigheder") '*** S�ttes normalt i topmenu **'
	
	    if print = "j" OR media = "export" then
	    %>
	    <!--#include file="../inc/regular/header_hvd_inc.asp"-->
	    <% 
	    else
	    %>
	    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	   
	    <% 
	    
	    end if
        %>
         <script src="inc/joblog_jav.js"></script>
        <%
	end if
	

            

    'if media <> "export" then
    if print <> "j" then         
	ldTop = 200
	ldLft = 240
	else
	ldTop = 15
	ldLft = 10
	end if

    if print <> "j" then
	%>
	<div id="load" style="position:absolute; display:; visibility:visible; top:<%=ldTop%>px; left:<%=ldLft%>px; width:300px; background-color:#ffffff; border:10px #cccccc solid; padding:10px; z-index:100000000;">
    <table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" /><br />
	&nbsp;&nbsp;&nbsp;&nbsp;<%=joblog2_txt_001 %>
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" />
	</td></tr></table>
	
	</div>
	<%end if

        'response.flush

    
   
    oimg = "ikon_stat_joblog_48.png"
	'oimg = "blank.gif"
    oleft = 20
	otop = ov_top
	owdt = 300
	oskrift = joblog2_txt_002 &" & "& joblog2_txt_003
	
    if media = "print" then
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
    end if

  

	
    %>
    <!-------------------------------Sideindhold------------------------------------->
    <div id="sindhold" style="position:absolute; left:<%=pleft%>px; top:<%=ptop%>px; visibility:visible; z-index:100;">
    
    <%


    if print <> "j" AND media <> "export" then
	

	
    	    
   
        call filterheader_2013(0,0,804,oskrift)%>
    	    
	  
	    <table cellspacing=0 cellpadding=0 border=0 width=100% bgcolor="#FFFFFF">
	    <form action="joblog.asp?rdir=<%=rdir %>&nomenu=<%=nomenu %>" method="post">
	    
	<%end if %> <!-- Print -->
	    
	    <%call medkunderjob %>
	    
	    </td>
	    </tr>
	    
	    
	
	<%if print <> "j" AND media <> "export" then%>
	
	<tr><td colspan=2>
	
	<table border=0 width=100% cellpadding=5 cellspacing=0>
	<tr>
			<!-- Brug altid datointerval, FM_usedatokri = 1 -->
			<input type="hidden" name="FM_usedatokri" id="FM_usedatokri" value="1">
			<td valign=top style="border-bottom:0px solid #cccccc;">

            <br /><b><%=joblog2_txt_004 %>:</b>
            <br />
			<!--#include file="inc/weekselector_s.asp"--> <!-- b -->
			
			<br /><br />
                <%=joblog2_txt_005 %>:<br />
                <input id="bruguge" name="bruguge" type="checkbox" value="1" <%=brugugeCHK %> />  <%=joblog2_txt_006 %>:
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

             &nbsp;&nbsp;&nbsp;&nbsp;
              <input id="brugmd" name="brugmd" type="checkbox" value="1" <%=brugmdCHK %> />    <%=joblog2_txt_007 %>:
               
			<select name="brugmd_md">
			<% for m = 1 to 12
			
			if m = 1 then
			useMd = 1
			else
			useMd = dateadd("m", 1, useMd)
			end if
			
			if cint(datepart("m", useMd, 2,2)) = cint(brugmd_md) then
			mSel = "SELECTED"
			else
			mSel = ""
			end if
			
			%>
			<option value="<%=datepart("m", useMd, 2,2) %>" <%=mSel%>><%=monthname(datepart("m", useMd, 2,2)) %></option>
			<%
			next %>
			
			</select>
			
			
                

			&nbsp;&nbsp;&nbsp;&nbsp; <%=joblog2_txt_008 %>: <select name="bruguge_year">
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
			
			</select>

         
		
		
			</td>
        <td>&nbsp;</td>
        </tr>


             <tr><td colspan="5" style="padding-top:20px;"><span id="sp_ava" style="color:#5582d2;">[+] <%=joblog2_txt_009 %></span></td></tr>
         

        <tr id="tr_ava" style="display:none; visibility:hidden;">

			<td colspan="5">
			<br /><b><%=joblog2_txt_010 %>:</b> (<%=joblog2_txt_011 %>)<br />
         
                    
            <table cellspacing=2 cellpadding=1 border=0 width=100%>
			<tr><td valign=top style="width:1px;">
            <%
			call akttyper2009(5)
			%>
			</table>
			</td></tr>


			</table>
       
        </td>
        </tr>
             <tr><td colspan="5" style="padding-top:20px;"><span id="sp_pre" style="color:#5582d2;">[+]<%=" "& joblog2_txt_012 &" " %>&<%=" "& joblog2_txt_013 %></span></td></tr>
         

        <tr id="tr_pre" style="display:none; visibility:hidden;">
			<td colspan="2" valign=top> 
         
            <%if rdir <> "treg" then %>
		
		 <div style="width:250px; padding:10px 10px 10px 10px; background-color:#F7F7F7;">	
   
        <input id="joblog_uge1" name="joblog_uge" type="radio" value="1" <%=joblog_ugeCHK1 %> onclick=showjoblogsubdiv() /> <b><%=joblog2_txt_014 %></b><%=" "& joblog2_txt_015 %><br />
	    
	    <div id="joblogsub" style="position:relative; background-color:#EFf3ff top:0px; left:20px; width:200px; border:0px #cccccc solid; padding:10px; visibility:<%=joblogsubWZB %>; display:<%=joblogsubDSP %>;">
	         <%=joblog2_txt_016 %>:<br />
        
        <%if fordelpamedarb = 1 then
		chkFordelmedarb = "CHECKED"
		else
		chkFordelmedarb = ""
		end if%>
		
		<input type="radio" name="FM_orderby_medarb" id="FM_orderby_medarb" value="1" <%=chkFordelmedarb%>><%=joblog2_txt_017 %><br>
	    
	    <%if fordelpamedarb = 2 then
		chkFordelpaakt = "CHECKED"
		else
		chkFordelpaakt = ""
		end if%>
		
		<input type="radio" name="FM_orderby_medarb" id="Radio1" value="2" <%=chkFordelpaakt%>><%=joblog2_txt_018 %><br>
	    
	    <%if fordelpamedarb = 0 then
		chkFordelingen = "CHECKED"
		else
		chkFordelingen= ""
		end if%>
		
		<input type="radio" name="FM_orderby_medarb" id="Radio2" value="0" <%=chkFordelingen%>><%=joblog2_txt_019 %>
	    
        <!--
        
        Vis Komprimeret visning<br />
                <input name="FM_komprimeret" id="FM_komprimeret" type="radio" value="1" <%=komCHK1 %> /> Ja 
			    <input name="FM_komprimeret" id="FM_komprimeret" type="radio" value="0" <%=komCHK0 %> /> Nej
	    
	    
	    -->
            <input id="FM_komprimeret" name="FM_komprimeret" value="1" type="hidden" />
	    
	    </div> 
	    
	   
	   <input id="joblog_uge2" name="joblog_uge" type="radio" value="2" <%=joblog_ugeCHK2 %> onclick=hidejoblogsubdiv() /> <b><%=joblog2_txt_020 %></b> <br />
	    
         <input id="joblog_uge3" name="joblog_uge" type="radio" value="3" <%=joblog_ugeCHK3 %> onclick=hidejoblogsubdiv() /> <b><%=joblog2_txt_021 %></b> <br /><br />
			
	
			</div>
		<%
		end if '** rdir **' 
            %>

                </td><td>
            <b><%=joblog2_txt_022 %></b><br />
            <%select case lto
            case "oko"
                    if level = 1 then%>
                    <input id="hidetimepriser" name="hidetimepriser" value="1" type="checkbox" <%=hidetpCHK %> /> <%=joblog2_txt_023 %><br />
                    <%else %>
                    <input id="hidetimepriser" name="hidetimepriser" value="1" type="hidden"/>
                    <%end if %>
            <%case else %>
			<input id="hidetimepriser" name="hidetimepriser" value="1" type="checkbox" <%=hidetpCHK %> /> <%=joblog2_txt_023 %><br />
		    <%end select %>

        <input id="hideenheder" name="hideenheder" value="1" type="checkbox" <%=hideehCHK %> /><%=" "& joblog2_txt_024 %> <br />
        <input id="hidegkfakstat" name="hidegkfakstat" value="1" type="checkbox" <%=hidegkCHK %> /><%=" "& joblog2_txt_025 %><br />
		<input id="Checkbox2" name="hidefase" value="1" type="checkbox" <%=hideFASCHK %> /><%=" "& joblog2_txt_026 %><br /><br />


		<input id="useenheder" name="useenheder" value="1" type="checkbox" <%=useehCHK %> /><%=" "& joblog2_txt_027 %>
		
		<%if level = 1 then %>
		<br />
		<input id="visKost" name="visKost" value="1" type="checkbox" <%=visKostCHK %> /> <%=joblog2_txt_028 %>
		<%end if %>

        <br /><input id="Checkbox3" name="showfor" value="1" type="checkbox" <%=showForCHK %> /> <%=joblog2_txt_029 %>

		<br /><input id="Checkbox1" name="showcpr" value="1" type="checkbox" <%=showCPRCHK %> /> <%=joblog2_txt_030 %>
		
     

       </td>
       </tr>
       </table>
       
       </td></tr>
       
       <tr>
       
       <td colspan=2 align=right valign=bottom>
       
       
	<input type="submit" value=" <%=joblog2_txt_031 %> >> ">
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
	end if 'print / media
	









startDatoKriSQL = strAar &"/"& strMrd &"/"& strDag
slutDatoKriSQL = strAar_slut &"/"& strMrd_slut &"/"& strDag_slut
	
	
	
	'*****************************************************************************************************
		'**** Job der skal indg� i oms�tning, budget og forbrugstal *******************************
		'*****************************************************************************************************
		
		
		'*** valgte job ***
		call valgtejob()
		
		
		'**** Valgte aftaler *****
		call valgteaftaler()
				
				
				    'if lto = "essens" then
					'Response.Write "her"
					'Response.Write "kid "& cint(kundeid) &" jobid "& cint(jobid) &" jobans "& cint(jobans) _ 
				    '& " kans " & cint(kundeans) & " sogval "& len(trim(jobSogVal)) & " aftid "& cint(aftaleid) 
				    'end if
				
                'Response.write "<br>jobstKri2: " & jobstKri & "<br>"
                'Response.write "jobnrSQLkri: OPR" & jobnrSQLkri & "<br><br>"

				'*** For at spare (trimme) p� SQL hvis der v�lges alle job alle kunder og vis kun for jobanssvarlige ikke er sl�et til ****
				'*** Og der ikke er s�gt p� jobnavn ***
				'if cint(kundeid) = 0 AND cint(jobid) = 0 AND cint(jobans) = 0 AND cint(jobans2) = 0 AND cint(jobans3) = 0 _
				' AND cint(kundeans) = 0 AND len(trim(jobSogVal)) = 0 AND cint(aftaleid) = 0 AND cint(segment) = 0 then 
						
					
					''jobidFakSQLkri = " OR jobid <> 0 "
					'jobnrSQLkri = " OR tjobnr <> '0' "
					'jidSQLkri =  " OR id <> 0 "
					'seridFakSQLkri = " OR aftaleid <> 0 "
						
				'end if
	
	
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
		
	
		

	
	if request("print") = "j" and  cint(joblog_uge) <> 3 then
	%>
	<h4>Periode <%=formatdatetime(strDag&"/"&strMrd&"/"&strAar, 1)%> - <%=formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 1)%> </h4>
	<%else
	
	    if media <> "export" then%>
	    <!--include file="inc/stat_submenu.asp"-->
	    <%end if%>
	<%end if%>



        <%if cint(hidegkfakstat) <> 1 then %>
       
          
           <form action="joblog.asp?func=opdaterliste&menu=<%=menu %>&rdir=<%=rdir %>&nomenu=<%=nomenu %>" method="post">
	
	    <input type="hidden" name="FM_aftaler" id="hidden9" value="<%=aftaleid%>">
	   <input type="hidden" name="FM_job" id="hidden8" value="<%=jobid%>">
	    <input type="hidden" name="FM_jobsog" id="hidden7" value="<%=jobSogVal%>">
	   <input type="hidden" name="FM_radio_projgrp_medarb" id="hidden4" value="<%=progrp_medarb%>">
		<input type="hidden" name="FM_progrp" id="hidden5" value="<%=progrp%>">
		<input type="hidden" name="FM_medarb" id="hidden6" value="<%=thisMiduse%>">
        <input type="hidden" name="FM_medarb_hidden" id="Text1" value="<%=thisMiduse%>">
	    <input type="hidden" name="FM_kunde" id="hidden1" value="<%=kundeid%>">
	    <input type="hidden" name="FM_kundeans" id="hidden2" value="<%=kundeans%>">
	    <input type="hidden" name="FM_jobans" id="hidden3" value="<%=jobans%>">
	   <input type="hidden" name="FM_kundejobans_ell_alle" value="<%=visKundejobans%>">
	    <input type="hidden" name="FM_akttype" value="<%=vartyper%>">
	    <input type="hidden" name="FM_orderby_medarb" value="<%=fordelpamedarb%>">
	    <input type="hidden" name="bruguge" value="<%=bruguge%>">
	    <input type="hidden" name="bruguge_week" value="<%=bruguge_week%>">
        <input type="hidden" name="brugmd" value="<%=brugmd%>">
	    <input type="hidden" name="brugmd_md" value="<%=brugmd_md%>">
	    <input type="hidden" name="bruguge_year" value="<%=bruguge_year%>">
	    
	    <input type="hidden" name="FM_start_mrd" value="<%=strMrd%>">
	    <input type="hidden" name="FM_start_dag" value="<%=strDag%>">
	    <input type="hidden" name="FM_start_aar" value="<%=strAar%>">
	    <input type="hidden" name="FM_slut_dag" value="<%=strDag_slut%>">
	    <input type="hidden" name="FM_slut_mrd" value="<%=strMrd_slut%>">
	    <input type="hidden" name="FM_slut_aar" value="<%=strAar_slut%>">
	    
	    
       
        
	    
	   <%end if %>
	   

          
                          
	   
      


        <%      
           dim timerTotprDag
        redim timerTotprDag(31)   
            
         end if 'cint(joblog_uge) = 3
        
        dim vlgt_typerTotaler, vlgt_typerTotalerGrand, vlgt_typerTotalerMed, vlgt_typerTotalerAkt
        redim vlgt_typerTotaler(200), vlgt_typerTotalerGrand(200), vlgt_typerTotalerMed(200), vlgt_typerTotalerAkt(200)
        
        dim vlgt_typerTotalerEnh, vlgt_typerTotalerGrandEnh, vlgt_typerTotalerMedEnh, vlgt_typerTotalerAktEnh
        redim vlgt_typerTotalerEnh(200), vlgt_typerTotalerGrandEnh(200), vlgt_typerTotalerMedEnh(200), vlgt_typerTotalerAktEnh(200)

      
        
         
        visopdaterknap = 0

		'*** Smiley ***''
		call ersmileyaktiv()
	    
        '*** BRUGES DENNE MERE? ???
	   
	    '*** Afsluttede uger ***''
		'dim medabAflugeId
		'redim medabAflugeId(2250)
		'strSQL2 = "SELECT u.status, u.afsluttet, WEEK(u.uge, 1) AS ugenr, YEAR(u.uge) AS aar, u.id, u.mid FROM ugestatus u WHERE "& ugeAflsMidKri &" AND uge BETWEEN '"& startDatoKriSQL &"' AND '"& slutDatoKriSQL &"' GROUP BY u.mid, uge"
		''Response.write strSQL2
		''Response.end
		'oRec2.open strSQL2, oConn, 3
		'while not oRec2.EOF
		'	medabAflugeId(oRec2("mid")) = medabAflugeId(oRec2("mid")) & "#"& oRec2("ugenr") &"_"& oRec2("aar") &"#,"
		'oRec2.movenext
		'wend
		'oRec2.close 
        '**************************
		
	
	'*******************'
	'**** MaIN SQL *****'
	'*******************'

    select case lto 'SHOW REPORT IN LOCAL CUR
    case "bf"
    call valutaKurs(5)
    tilkurs = dblKurs
    end select
	
	'*** Vis alle timer hvor man er jobanssv. **'
	if cint(visKundejobans) = 1 then
	medarbSQlKri = "tmnr <> 0" 
	else
	medarbSQlKri = replace(medarbSQlKri, "m.mid", "t.tmnr")
	end if
	
	call akttyper2009(7)

    if cint(joblog_uge) = 3 then
    dim medarbtimerArr
    m = 10000
    j = 13
    redim medarbtimerArr(m,j) 
    end if
	
    strExportOskriftDage = ""

	lastjobnavn = ""
	lastmedarb = 0
	strSQL = "SELECT Tdato, TasteDato, Tjobnr, Tjobnavn, a.navn AS Anavn, TAktivitetId, "_
	&" a.fakturerbar,"_
	&" Tknavn, Tmnr, Tmnavn, Timer, Tid, Tfaktim, TimePris, Timerkom, Tknr, t.kurs, "_
	&" sttid, sltid, a.faktor, godkendtstatus, godkendtstatusaf, kkundenr, kkundenavn, t.timerkom, "_
	&" m.mnr, m.init, mnavn, m.mid, a.id AS aid, a.sortorder, v.valutakode, t.valuta, k.kid, t.kostpris, kpvaluta, kpvaluta_kurs, m.mcpr, "_
	&" a.fase, a.antalstk, a.aktbudgetsum, a.bgr, a.aktbudget, t.editor, a.avarenr, overfort "_
	&" FROM timer t "_
	&" LEFT JOIN aktiviteter a ON (a.id = TAktivitetId)"_
    &" LEFT JOIN kunder k ON (k.kid = Tknr)"_
	&" LEFT JOIN medarbejdere m ON (m.mid = tmnr)"_
    &" LEFT JOIN valutaer v ON (v.id = t.valuta)"_
	&" WHERE ("& jobnrSQLkri &") AND ("& medarbSQlKri &") AND ("& aty_sql_sel &") AND "_
	&" (Tdato BETWEEN '"& startDatoKriSQL &"' AND '"& slutDatoKriSQL &"') ORDER BY "& orderByKri
	
	'&" WHERE "& jobMedarbKri &" "& selaktidKri &""_
    'Response.write strSQL & "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"
	
	
	'Response.Write "visKundejobans "& visKundejobans &"<br>jobAnsSQLkri:" & jobAnsSQLkri & "<br>"& jobAns2SQLkri & "<br> kundeans"& kundeAnsSQLkri & "<br><br>"
	
	'if lto = "sdeo" and session("mid") = 1 then
	'Response.Write strSQL
	'Response.end
	'end if
	
	oRec.open strSQL, oConn, 0, 1
	'showFaseTot = 0
	v = 0
    m = 0
    x = 0
    antalM = 0
	lastFase = ""
	thisFase = ""
	while not oRec.EOF
		
		
		
		
            if cint(joblog_uge) = 3 then
            medarbtimerArr(m,0) = oRec("Tmnavn")
            medarbtimerArr(m,1) = oRec("Tmnr")
            medarbtimerArr(m,2) = oRec("kkundenavn")
            medarbtimerArr(m,3) = oRec("kkundenr")
            medarbtimerArr(m,4) = oRec("Tjobnavn")
            medarbtimerArr(m,5) = oRec("Tjobnr")
            medarbtimerArr(m,6) = oRec("anavn")
            medarbtimerArr(m,7) = oRec("aid")
            medarbtimerArr(m,8) = oRec("Tdato")
            medarbtimerArr(m,9) = oRec("timer")
            medarbtimerArr(m,10) = oRec("kid")
            medarbtimerArr(m,11) = oRec("timerkom")
            medarbtimerArr(m,12) = oRec("init")
            medarbtimerArr(m,13) = oRec("mcpr")
            end if



        strWeekNum = datepart("ww", oRec("Tdato"),2,2)
		thisFase = oRec("fase")
		id = 1
		
			
            '** P� kasserapoort BF skal linjeskift laves inden ny linje s� der ikke kommer et skift for meget til sidst
            if cint(ver) = 1 then
                        if v > 0 then
                        ekspTxt = ekspTxt & "xx99123sy#z"
                        else
                        ekspTxt = ekspTxt 
                        end if
            'else
            'ekspTxt = ekspTxt & "xx99123sy#z"
            end if

			
		    select case ver 
            case 1 
            case else
			ekspTxt = ekspTxt & oRec("kkundenavn")&";"&oRec("kkundenr")&";"&oRec("Tjobnavn") &";"& oRec("Tjobnr") &";"
		    end select
		
			
			
			if media <> "export" then
			
			
			'** Fordel p� medarb **'
			select case cint(joblog_uge)
            case 1
			
			
		
		        
		        'Response.write lastmedarb &" <> " & oRec("Tmnr") & " job: " & lastjobnavn &" <> " & oRec("Tjobnavn") 
		        'Response.Write lastaktid & "<>"& oRec("Taktivitetid")
				
		        if fordelpamedarb = 1 then
				
				    if (lastmedarb <> oRec("Tmnr") OR lastjobnr <> oRec("Tjobnr")) then
					    if v <> 0 then
    					
					    call fordelpamedarbejdere()
    					
    					
					    end if
				    end if
				
				end if
				
				
				'** Fordel p� akt **'
				if fordelpamedarb = 2 then
				
				    if (lastaktid <> oRec("Taktivitetid") OR lastjobnr <> oRec("Tjobnr")) then
					    if v <> 0 then
					    
					    '** Nulstiller fase ved n�ste job ***'
					    if lastjobnr <> oRec("Tjobnr") then
    					call fordelpaakt(1)
    					else
    					call fordelpaakt(0)
    					end if
    					
					    
    					
    					
					   
    					
					    end if
				    end if
				
				end if
		        
		        
				
				if lastjobnr <> oRec("Tjobnr") then
				
				
				        if v <> 0 then
        				
				        call jobtotaler()
        				
				      
				        end if
				
				
				
				 
				   
				
				tTop = 40
                tLeft = 0
                tWdth = globalWdt


                call tableDiv(tTop,tLeft,tWdth)

                
                       call jobansoglastfak()
                
                
                
                %>
                <table border="0" width="100%" cellpadding="0" cellspacing="0" bgcolor="#ffffff">
				
				
				
				<%
				komprimer = "1"
				if komprimer <> "1" then %>
				<tr>
				    
					<td colspan="16" style="padding:5px; height:100px;">
					
					
					
					
					<table cellpadding=0 cellspacing=0 border=0 width=100%>
					<tr><td valign=top>
					<img src="../ill/ikon_kunder_16.png" alt="Kontakt" border="0">&nbsp;&nbsp;<%=oRec("kkundenavn")%>&nbsp;(<%=oRec("kkundenr")%>)<br>
					<img src="../ill/ikon_job_16.png" alt="Job" border="0">&nbsp;&nbsp;<b><%=oRec("Tjobnavn")%> (<%=oRec("Tjobnr")%>)</b><br />
					
					
						<%=joblog2_txt_032 &" "%>1: <b><%=jobans1txt%></b>
						<%if len(trim(jobans2txt)) <> 0 then%>
						<br>
	   					<%=joblog2_txt_032 &" "%> 2: <b><%=jobans2txt%></b>
						<%end if
						jobans1txt = ""
						jobans2txt = ""
						%>
						</td><td valign=top style="padding-left:20px;">
						
						
						<%select case jobstatus
						case 1
						thisjobstatus = joblog2_txt_033
						case 2
						thisjobstatus = joblog2_txt_034
						case 0
						thisjobstatus = joblog2_txt_035
						end select%>
						<%=joblog2_txt_026 %>: <b><%=thisjobstatus%></b><br>
						<%=joblog2_txt_037 %>: <b><%=formatdatetime(jobstartdato, 1)%> - <%=formatdatetime(jobslutdato, 1)%></b>
						<br>
						<%=joblog2_txt_038 %>: <b><%=formatnumber(jobForkalkulerettimer, 2)%></b><br>
						
						<%if cint(fastpris) = 1 then %>
						<%=joblog2_txt_039 %>: <b><%=joblog2_txt_040 %></b>
						
						<%if cint(usejoborakt_tp) <> 0 then   %>
						(<%=joblog2_txt_041 %>)
						<% else %>
						(<%=joblog2_txt_042 %>)
						<%end if %>
						
						<%else %>
						<%=joblog2_txt_043 %>: <b><%=joblog2_txt_044 %></b>
						<%end if %><br />
						
						<%if len(trim(rekvnr)) <> 0 then %>
						<%=joblog2_txt_045 %>: <b><%=rekvnr %></b><br />
						<%end if %>
						
						<%=joblog2_txt_046 %>: 
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
				
				
		 		
		 case 2 '** joblog_uge = 2 **'
		
		            
		            
		            'Response.Write "lastdate: " & cdate(lastdate) &" <> "& cdate(oRec("tdato")) & " v: "& v &"<br>" 
		            
		            if cdate(lastdate) <> cdate(oRec("tdato")) OR (lastmedarb <> oRec("Tmnr"))  then
		            
		             if v <> 0 then
		                    
		                call dagstotaler(2)
		                
		                timerTotdag = 0
				        enhederTotdag = 0
				     else
				        
				        call dagstotaler(1)
				        
		             end if
		             
		           
		            
		            end if
		            
				
		            
		            
		            
		            if lastWeekNum <> strWeekNum OR lastmedarb <> oRec("Tmnr") then
				    
				    
				    
				   if v <> 0 then
				   
				
        				
				        call jobtotaler()
        				
				       
				   
				   end if
				    
				
				tTop = 40
                tLeft = 0
                tWdth = globalWdt


                call tableDiv(tTop,tLeft,tWdth)



                %>
                
                <!--lastWeekNum &" <> "& strWeekNum &" OR "& lastmedarb &" <>"& oRec("Tmnr")  -->
                
                
                <table border="0" width=100% cellpadding="0" cellspacing="0" bgcolor="#ffffff">
              
				    
					
				    <tr>
				        <td colspan=16 bgcolor="#ffffff" style="padding:10px 10px 0px 10px;">
				        <h4><%=joblog2_txt_006 %>: <%=datepart("ww", oRec("tdato"), 2,2)&" - "& datepart("yyyy", oRec("tdato"), 2,2)%> - <%=oRec("tmnavn") %> (<%=oRec("tmnr") %>)</h4>
				        
				        <%
				        call erugeAfslutte(datepart("yyyy", oRec("tdato"), 2,2), datepart("ww", oRec("tdato"), 2,2), oRec("tmnr"), SmiWeekOrMonth, 0) 
				        %>
				        
				        <%if showAfsuge = 0 then %>
				        <%=joblog2_txt_047 %> <%=formatdatetime(cdAfs, 2)%> kl. <%=formatdatetime(cdAfs, 3)%> (uge <%=datepart("ww", cdAfs, 2, 2)%>)
                        
		                <%end if%>
				        
				        </td>
				    </tr>
				    
				    
				    
				    <%
				    
				    
				    call tableheader()
                    
                    call dagstotaler(0)
                                
				    end if
				    
				    
				
				
		case 3 'm�nedsoversigt
                
				
		end select '*** End Joblog_uge **'

        'else
         'call jobansoglastfak()
		
		end if '** media	  
				

        if cint(joblog_uge) <> 3 then 'm�nedsoversigt

                     
		                        call akttyper2009Prop(oRec("tfaktim"))
		                     
                        
                              '*** Finder jobans og last fak hvis liste vises    ***'
                             '*** som ugeseddel, ellers er                      ***'
                             '*** jobans og last fak allerede fundet            ***'
                             if cint(joblog_uge) = 2 OR media = "export" then
                             call jobansoglastfak()
                             end if 
                 
                             '***'
	   
	   
	         
				
				
				              if (cint(oRec("tfaktim")) = 20 OR cint(oRec("tfaktim")) = 21 OR _
				               cint(oRec("tfaktim")) = 8 OR cint(oRec("tfaktim")) = 81) AND _
				                level <> 1 AND cdbl(session("mid")) <> cdbl(oRec("tmnr")) then 
	                           hidesygtyper = 1
	                           else 
	                           hidesygtyper = 0
	                          end if
	              
	              
	                          '*** Skjul sygdmstyper for andre end end selv + admin ****
	                         if hidesygtyper <> 1 then 
				

                            if len(trim(oRec("fase"))) <> 0 then
                            thisFase = replace(oRec("fase"), "_", " ") 
                            else
                            thisFase = ""
                            end if
				
				            if media <> "export" then
				            %>
				
				
				
				
				            <%if print = "j" then %>
				            
                                <%end if %>
				            <tr>
				            <td valign="top">&nbsp;</td>
				            <td style="padding-top:3px; width:25px; border-top:1px #cccccc solid;" valign="top"><%=strWeekNum%>
				            </td>
                            <%

                                dayTxtTrans = "xxx"
                                if left(weekdayname(weekday(oRec("Tdato"))), 3) = "man" then 
                                dayTxtTrans = joblog2_txt_130
                                end if
                                if left(weekdayname(weekday(oRec("Tdato"))), 3) = "tir" then 
                                dayTxtTrans = joblog2_txt_131
                                end if
                                if left(weekdayname(weekday(oRec("Tdato"))), 3) = "ons" then 
                                dayTxtTrans = joblog2_txt_132
                                end if
                                if left(weekdayname(weekday(oRec("Tdato"))), 3) = "tor" then 
                                dayTxtTrans = joblog2_txt_133
                                end if
                                if left(weekdayname(weekday(oRec("Tdato"))), 3) = "fre" then 
                                dayTxtTrans = joblog2_txt_134
                                end if
                                if left(weekdayname(weekday(oRec("Tdato"))), 3) = "l�r" then 
                                dayTxtTrans = joblog2_txt_135
                                end if
                                if left(weekdayname(weekday(oRec("Tdato"))), 3) = "s�n" then 
                                dayTxtTrans = joblog2_txt_136
                                end if

                                monthTxtTrans = "xxx"
                                if left(monthname(month(oRec("Tdato"))), 3) = "jan" then
                                monthTxtTrans = joblog2_txt_118
                                end if
                                if left(monthname(month(oRec("Tdato"))), 3) = "feb" then
                                monthTxtTrans = joblog2_txt_119
                                end if
                                if left(monthname(month(oRec("Tdato"))), 3) = "mar" then
                                monthTxtTrans = joblog2_txt_120
                                end if
                                if left(monthname(month(oRec("Tdato"))), 3) = "apr" then
                                monthTxtTrans = joblog2_txt_121
                                end if
                                if left(monthname(month(oRec("Tdato"))), 3) = "maj" then
                                monthTxtTrans = joblog2_txt_122
                                end if
                                if left(monthname(month(oRec("Tdato"))), 3) = "jun" then
                                monthTxtTrans = joblog2_txt_123
                                end if
                                if left(monthname(month(oRec("Tdato"))), 3) = "jul" then
                                monthTxtTrans = joblog2_txt_124
                                end if
                                if left(monthname(month(oRec("Tdato"))), 3) = "aug" then
                                monthTxtTrans = joblog2_txt_125
                                end if
                                if left(monthname(month(oRec("Tdato"))), 3) = "sep" then
                                monthTxtTrans = joblog2_txt_126
                                end if
                                if left(monthname(month(oRec("Tdato"))), 3) = "okt" then
                                monthTxtTrans = joblog2_txt_127
                                end if
                                if left(monthname(month(oRec("Tdato"))), 3) = "nov" then
                                monthTxtTrans = joblog2_txt_128
                                end if
                                if left(monthname(month(oRec("Tdato"))), 3) = "dec" then
                                monthTxtTrans = joblog2_txt_129
                                end if

                            %>
				            <td style="padding-top:3px; white-space:nowrap; border-top:1px #cccccc solid; padding-right:10px; padding-left:10px;" valign="top"><%=dayTxtTrans %>. <%=day(oRec("Tdato")) &" "& monthTxtTrans &". "& right(year(oRec("Tdato")), 4)%></td>
				
				
				            <%if komprimer = "1" then %>
				            <td style="padding-top:3px; padding-right:15px; border-top:1px #cccccc solid; width:175px;" valign="top">
                            <%if print <> "j" then %>
				            <b><%=left(oRec("Tknavn"), 40)%></b> <label style="font-size:9px;"><%=oRec("kkundenr") %></label><br />
                            <%else %>
				            <b><%=left(oRec("Tknavn"), 20)%></b><br />
                            <%end if %>
             

                
                            <%if print <> "j" then %>
				            <%=left(oRec("tjobnavn"), 40)%> (<%=oRec("tjobnr") %>)
                            <%else %>
                            <%=left(oRec("tjobnavn"), 40)%> (<%=oRec("tjobnr") %>)
                            <%end if %>
               
				

                             <%if print <> "j" then %>
				           
                            <!--<span style="color:#999999; font-size:9px;">
				
				            <%'if cint(fastpris) = 1 then %>
				            (fastpris)
                            <%'else %>
                            (lbn. timer)
				            <%'end if %>
				            </span>
				            -->
                
                            <span style="color:#999999; font-size:9px;">
                            <%if jobans1 <> 0 then %>
				            <br /><%=jobans1txt %>
				            <%end if %>
				
				            <%if jobans2 <> 0 then %>
				        
				                    <%if jobans1 <> 0 then %>
				                    <%=", " %>
				                    <%end if %>
				                    <%=jobans2txt %>
				            <%end if %>
				            </span>
				
                            <%end if %>

				            </td>
				            <%else %>
				            <td valign="top" style="width:10px; border-top:1px #cccccc solid;">
                                &nbsp;</td>
				            <%end if %>
				
              
                            <%if cint(hidefase) <> 1 then %>
				            <td valign=top style="padding:3px 10px 5px 0px; width:100px; border-top:1px #cccccc solid;"><%=thisFase%>&nbsp;</td>
                            <%
                
                            else %>
                            <td style="border-top:1px #cccccc solid;"><img src="ill/blank.gif" width="1" height="1" border="0" /></td>
                            <%end if %>
				
				            <%
				
				            end if 'media
				

                            select case ver 
                            case 1 
                            
                            if isNull(oRec("avarenr")) <> true AND len(trim(oRec("avarenr"))) > 5 AND instr(oRec("avarenr"), "M") <> 0 AND lto = "bf" then

                            kontoTxt = trim(oRec("avarenr"))
                            kontonrLen = len(kontoTxt)
                            kontonrM = instr(kontoTxt, "M") 
                            kontonrLeft = mid(kontoTxt, 2, kontonrM-2)
                            kontonrRight = mid(kontoTxt, kontonrM+1, kontonrLen)

                            else

                            kontonrLeft = oRec("avarenr")
                            kontonrRight = 0

                            end if

                            call meStamdata(oRec("tmnr"))
                            bilagsnr = "" 'year(oRec("tdato"))&month(oRec("tdato"))&day(oRec("tdato")) 

                            ekspTxt = ekspTxt & chr(34) & formatdatetime(oRec("tdato"), 2) & chr(34) &";"& chr(34) & oRec("tjobnavn") &" ["& meInit &"]" & chr(34) &";" & bilagsnr &";"& chr(34) & kontonrLeft & chr(34) &";" & chr(34) & oRec("anavn") & chr(34) &";"
                                
                            case else

                            if cint(hidefase) <> 1 then 
                            ekspTxt = ekspTxt & thisFase &";"
				            end if
				 
				            ekspTxt = ekspTxt & strWeekNum &";"& formatdatetime(oRec("tdato"), 2) &";"
				
                            end select

                            '*** Er uge afsluttet af medarb, er smiley og autogk sl�et til ***'
                            'Denne kan lukkes ned / fjernes 
                            'erugeafsluttet = instr(medabAflugeId(oRec("mid")), "#"&datepart("ww", oRec("Tdato"),2,2)&"_"&datepart("yyyy", oRec("Tdato"))&"#")
                
                     
                

                            strMrd_sm = datepart("m", oRec("Tdato"), 2, 2)
                            strAar_sm = datepart("yyyy", oRec("Tdato"), 2, 2)
                            strWeek = datepart("ww", oRec("Tdato"), 2, 2)
                            strAar = datepart("yyyy", oRec("Tdato"), 2, 2)

                            if cint(SmiWeekOrMonth) = 0 then
                            usePeriod = strWeek
                            useYear = strAar
                            else
                            usePeriod = strMrd_sm
                            useYear = strAar_sm
                            end if

                
                            call erugeAfslutte(useYear, usePeriod, oRec("mid"), SmiWeekOrMonth, 0)
		        
		                    'Response.Write "smilaktiv: "& smilaktiv & "<br>"
		                    'Response.Write "SmiWeekOrMonth: "& SmiWeekOrMonth &" ugeNrAfsluttet: "& ugeNrAfsluttet & " tjkDag: "& tjkDag &"<br>"
		                    'Response.Write "autolukvdatodato: "& autolukvdatodato & "<br>"
		                    'Response.Write "tjkDag: "& tjkDag & "<br>"
		                    'Response.Write "autolukvdato: "& autolukvdato & "<br>"
		                    'Response.Write "erugeafsluttet:" & erugeafsluttet & "<br>"
		        
		                    call lonKorsel_lukketPer(oRec("Tdato"), jobRisiko)
		         
                            'if ( (( datepart("ww", ugeNrAfsluttet, 2, 2) = usePeriod AND cint(SmiWeekOrMonth) = 0) OR (datepart("m", ugeNrAfsluttet, 2, 2) = usePeriod AND cint(SmiWeekOrMonth) = 1 )) AND cint(ugegodkendt) = 1 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                            '(smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", oRec("Tdato")) = year(now) AND DatePart("m", oRec("Tdato")) < month(now)) OR _
                            '(smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", oRec("Tdato")) < year(now) AND DatePart("m", oRec("Tdato")) = 12)) OR _
                            '(smilaktiv = 1 AND autolukvdato = 1 AND DatePart("yyyy", oRec("Tdato")) < year(now) AND DatePart("m", oRec("Tdato")) <> 12) OR _
                            '(smilaktiv = 1 AND autolukvdato = 1 AND (year(now) - DatePart("yyyy", oRec("Tdato")) > 1))) OR cint(lonKorsel_lukketIO) = 1 then
              
                            'ugeerAfsl_og_autogk_smil = 1
                            'else
                            'ugeerAfsl_og_autogk_smil = 0
                            'end if 
				
                             '*** tjekker om uge er afsluttet / lukket / l�nk�rsel
                            call tjkClosedPeriodCriteria(oRec("tdato"), ugeNrAfsluttet, usePeriod, SmiWeekOrMonth, splithr, smilaktiv, autogk, autolukvdato, lonKorsel_lukketIO)

				            'if print <> "j" then
				            'awdt = 350
				            'else
				            awdt = 250
				            'end if
				            'oRec("fakturerbar")
				            call akttyper(oRec("fakturerbar"), 1)

				            if media <> "export" then
				            %>
				            <td style="padding:3px 10px 5px 0px; width:<%=awdt%>px; border-top:1px #cccccc solid;" valign="top">
				            <%
				                if ((oRec("godkendtstatus") <> 1 AND oRec("godkendtstatus") <> 3 AND request("print") <> "j" AND ugeerAfsl_og_autogk_smil = 0) _
				                OR (oRec("godkendtstatus") <> 1 AND oRec("godkendtstatus") <> 3 AND request("print") <> "j" AND ugeerAfsl_og_autogk_smil = 1 AND level = 1)) _
				                AND (cdate(lastfakdato) <  cdate(oRec("Tdato")) AND oRec("overfort") = 0) then %>
    					
    					
    					            <a href="#" onclick="Javascript:window.open('../to_2015/rediger_tastede_dage_2006.asp?id=<%=oRec("tid") %>', '', 'width=450,height=675,resizable=yes,scrollbars=yes')" class=vmenu>
					                <!--<a href="../to_2015/rediger_tastede_dage_2006.asp?id=<%=oRec("Tid")%>" target="_blank" class=vmenu>-->
    					
					                <%=oRec("anavn")%>
    					
					                <img src="../ill/blyant.gif" width="12" height="11" alt="Tilf�j kommentar" border="0">
					                </a>
					    
				                    &nbsp;<span style="color:#999999; font-size:9px;">(<%=lcase(akttypenavn)%>)</span>
    				    
    					
    					
    					
    				
				                <%else%>
					                <b><%=oRec("anavn")%></b>
					                <%
					    
					                %>
				                    &nbsp;<span style="color:#999999; font-size:9px;">(<%=lcase(akttypenavn)%>) <!--: <%=oRec("fakturerbar") %>/<%=oRec("tfaktim") %>--></span> 
    				    
    					
    					
    					
				                <%end if


                                  '*** varenr **'
                                  if isNull(oRec("avarenr")) <> true AND len(trim(oRec("avarenr"))) <> 0 then 
                                  %><br /><span style="color:#5582d2; font-size:9px;"><%=oRec("avarenr")%></span><%  
                                  end if
                                      
                                


				

                            timerKom = ""
                            timerKom = oRec("timerkom")
				            if len(trim(timerKom)) <> 0 then%>
				            <br /><i><%=timerKom%></i>
				            <%end if%>
				
			                <%
				            '**** Kundekommentarer ****
					
					            kundeKomm = ""
					            strSQLin = "SELECT note FROM incidentnoter WHERE timerid = "& oRec("tid") &" ORDER BY id"
					            oRec3.open strSQLin, oConn, 3 
					            i = 0
					
					            while not oRec3.EOF 
					            %>
					
					            <br><br><i><%=oRec3("note")%></i>
                   
                                <%
					            kundeKomm = kundeKomm & oRec3("note") & "<br /><br />"
					
					
					
					
					            oRec3.movenext
					            wend
					            oRec3.close  
					            %>
					
				
				             &nbsp;</td>
				            <%
				
                
                
                            end if '** media
                



                            aktnavnEksp = ""
                            if len(oRec("Anavn")) <> 0 then
                            'aktnavnEksp = replace(oRec("Anavn"), Chr(34), "&quot;")
                            aktnavnEksp = replace(oRec("Anavn"), "''", "")
                            aktnavnEksp = replace(oRec("Anavn"), "'", "")
                            else
                            aktnavnEksp = ""
                            end if
                

                            select case ver
                            case 1
                            case else
               
				                    'call akttyper2009Prop(oRec("tfaktim"))
				                    ekspTxt = ekspTxt & aktnavnEksp &";"& Chr(34) & akttypenavn & Chr(34) &";"& aty_fakbar &";" 
				
				                    if lto = "bowe" then
					                    ekspTxt = ekspTxt & akttid &";" 
				                    end if
				
                            end select


                            ekspTxt = ekspTxt & Chr(34) & oRec("avarenr") & Chr(34) &";" 


                            if cint(showfor) = 1 then
                    

                                forr = ""
                                strSQLfor = "SELECT f.navn AS forr FROM fomr_rel AS fr "_
                                &" LEFT JOIN fomr AS f ON (f.id = fr.for_fomr) WHERE for_aktid = " & oRec("TAktivitetId")

                                fo = 0
                                oRec5.open strSQLfor, oConn, 3
                                while not oRec5.EOF 

                                if fo = 0 then
                                forr = oRec5("forr") 
                                else
                                forr = forr &"<br>"& oRec5("forr") 
                                end if


                                fo = fo + 1
                                oRec5.movenext
                                wend
                                oRec5.close

                            if media <> "export" then%>
				            <td style="padding-top:3px; padding-left:5px; border-top:1px #cccccc solid;" valign="top"><span style="color:#999999; font-size:10px;"><%=forr %></span>

                            </td>
				            <%

                            else

                            %>
				            <td style="padding-top:3px; padding-left:5px; border-top:1px #cccccc solid;" valign="top"><img src="ill/blank.gif" width="1" height="1" border="0" /></td>
				            <%

				            end if

                            forrExp = replace(forr, "<br>", ", ") 

                            select case ver
                            case 1
                            case else
                            ekspTxt = ekspTxt & forrExp &";"
                            end select

                            end if


                            



				            if media <> "export" then%>
				            <td style="padding-top:3px; padding-left:5px; border-top:1px #cccccc solid;" valign="top"><%=left(oRec("Tmnavn"),25)%>&nbsp;[<%=oRec("init")%>]
                    
                                <%if cint(showcpr) = 1 then%>
                                <br /><span style="color:#999999;">CPR: <%=oRec("mcpr") %></span> 
                                <%end if %>

                            </td>
				            <%
				            end if
				
                            select case ver
                            case 1
                            case else

				            ekspTxt = ekspTxt & oRec("mnavn") &";"&oRec("mnr")&";"&oRec("init")&";"

                                if cint(showcpr) = 1 then
                                ekspTxt = ekspTxt & oRec("mcpr") & ";" 
                                end if
				
                             end select
				
				            if media <> "export" then%>
				            <td align="right" style="padding-top:3px; padding-right:5px; border-top:1px #cccccc solid;" valign="top"><b> <%=formatnumber(oRec("Timer"), 2)%></b>
                    
                  

				            <%
				            end if

                            select case ver
                            case 1
                                
                                    
                                  if cint(cur) = 0 then '* BF Basis currency OMreng altid til DKK

                                  
                                  belob = formatnumber(oRec("timer") * oRec("timepris"), 2)
                                  frakurs = oRec("kurs")
	                              call beregnValuta(belob,frakurs,100)

                                  ekspTxt = ekspTxt & formatnumber(valBelobBeregnet, 2) &";;;"& chr(34) & kontonrRight & chr(34) &""
                                  
                                  else '*BF KEEP local currency, use the actual currency

                                  belob = formatnumber(oRec("timer") * oRec("timepris"), 2)
                                  

                                  select case lto 
                                  case "bf"
                                  frakurs = oRec("kurs")

                                  'if cdbl(frakurs) <> cdbl(tilkurs) then 'ONLY IF NOT IN LOCAL CUR
                                  'call beregnValuta(belob,frakurs,tilkurs)
                                  'else
                                  valBelobBeregnet = belob
                                  'end if

                                  case else
                                  valBelobBeregnet = belob
                                  end select

	                              

	                              ekspTxt = ekspTxt & formatnumber(valBelobBeregnet, 2) &";;;"& chr(34) & kontonrRight & chr(34) &""
                                  
                                  end if

                            case else
				            ekspTxt = ekspTxt & oRec("timer") &";" ';91000
                            end select
				

                            select case ver 
                            case 1

                            case else
				
				                if len(oRec("sttid")) <> 0 then
				
					                if left(formatdatetime(oRec("sttid"), 3), 5) <> "00:00" then
					                    if media <> "export" then
					                    Response.write "<span style=""color:#999999; font-size:9px;""><br>"& left(formatdatetime(oRec("sttid"), 3), 5) & " - " & left(formatdatetime(oRec("sltid"), 3), 5)
					                    end if
					                    ekspTxt = ekspTxt & left(formatdatetime(oRec("sttid"), 3), 5) & " - " & left(formatdatetime(oRec("sltid"), 3), 5) &";"
					                else
					                    ekspTxt = ekspTxt &";"
					                end if
				
				                else
				                ekspTxt = ekspTxt &";"
				                end if

                            end select
				
				            if media <> "export" then%>
				            </td>
				            <%end if
				
				
				            enheder = 0
				                enheder = oRec("faktor") * oRec("timer")
    				
				                if enheder <> 0 then
				                enheder = cdbl(enheder)
				                else
				                enheder = 0
				                end if
				
				            if cint(hideenheder) = 0 then 
				
				    
				                if media <> "export" then%>
				                <td align="right" style="padding-top:3px; padding-right:5px; border-top:1px #cccccc solid; white-space:nowrap;" valign="top"><%=formatnumber(enheder, 2)%></td>
				                <%end if

                                select case ver 
                                case 1
                                case else
				                ekspTxt = ekspTxt & enheder &";"
                                end select
				
				            else
				                if media <> "export" then%>
				                <td style="border-top:1px #cccccc solid;"><img src="ill/blank.gif" width="1" height="1" border="0" /></td>
				                <%end if
				            end if
				
				            vlgt_typerTotaler(oRec("tfaktim")) = vlgt_typerTotaler(oRec("tfaktim")) + oRec("Timer")
				            vlgt_typerTotalerAkt(oRec("tfaktim")) = vlgt_typerTotalerAkt(oRec("tfaktim")) + oRec("Timer")
				            vlgt_typerTotalerMed(oRec("tfaktim")) = vlgt_typerTotalerMed(oRec("tfaktim")) + oRec("Timer")
				            vlgt_typerTotalerGrand(oRec("tfaktim")) = vlgt_typerTotalerGrand(oRec("tfaktim")) + oRec("Timer")
				
				            vlgt_typerTotalerEnh(oRec("tfaktim")) = vlgt_typerTotalerEnh(oRec("tfaktim")) + enheder
				            vlgt_typerTotalerAktEnh(oRec("tfaktim")) = vlgt_typerTotalerAktEnh(oRec("tfaktim")) + enheder
				            vlgt_typerTotalerMedEnh(oRec("tfaktim")) = vlgt_typerTotalerMedEnh(oRec("tfaktim")) + enheder
				            vlgt_typerTotalerGrandEnh(oRec("tfaktim")) = vlgt_typerTotalerGrandEnh(oRec("tfaktim")) + enheder
				
				
				
				
				
				           
				
				            '*** Timepriser ***'
				            if (level <=2 OR level = 6) AND cint(hidetimepriser) = 0 then
				
				            if media <> "export" then%>
				            <td style="padding-top:3px; padding-right:5px; white-space:nowrap; border-top:1px #cccccc solid;" align="right" valign="top">
				            <%
				            end if
				    
				    
				    
				             
				                '*** Alle fakturerbare & Med p� faktura ***'
				    
				                if cint(aty_fakbar) = 1 OR cint(aty_medpafak) = 1 OR cint(oRec("tfaktim")) = 61 then 'stk antal
				    
				                    'if cint(fastpris) = 0 then '** lbn. timer
    				                tpris = formatnumber(oRec("TimePris"), 2)
    				                'else
    				                '    if cint(usejoborakt_tp) = 0 then
    				                '    tpris = formatnumber(fasttimePris, 2) 'Fra job
    				                '    else
    				                '    tpris = formatnumber(oRec("aktbudget"), 2) 'Enheds pris timer eller. stk. fra aktivitet
    				                '    end if
    				                'end if 
    				    
    				    
    				                if media <> "export" then%>
    				                <%=tpris %>
				                    <%end if 
				                
                                    select case ver 
                                    case 1
                                    case else
				                    ekspTxt = ekspTxt & formatnumber(tpris, 2) &";"
                                    end select
    				            else
    				
    				            if media <> "export" then%>
    				            &nbsp;
    				            <%end if

                                    select case ver 
                                    case 1
                                    case else
    				                ekspTxt = ekspTxt & ";"
                                    end select

				                end if
				
				                if media <> "export" then%>
				                </td>
			                    <td style="padding-top:3px; padding-left:15px; padding-right:5px; border-top:1px #cccccc solid; white-space:nowrap;" align="right" valign="top">
				                <%
				                end if
				    
				                if cint(aty_fakbar) = 1 OR cint(aty_medpafak) = 1 OR cint(oRec("tfaktim")) = 61 then 'stk. antal%>
				        
				                <%if useenheder = 1 then
				                tprisTot = tpris*enheder
				                else
				                tprisTot = tpris*oRec("timer")
				                end if 
				    
				                if media <> "export" then%>
				                <%=formatnumber(tprisTot , 2)&" "&oRec("valutakode")%>
			                    <%end if 
			        
			                    select case ver 
                                case 1
                                case else
			                    ekspTxt = ekspTxt & formatnumber(tprisTot, 2) &";"&oRec("valutakode")&";"
                                end select


    			                else
    			                
                                if media <> "export" then%>
    			                &nbsp;
    			                <%end if


                                select case ver 
                                case 1
    			                case else
                                ekspTxt = ekspTxt & ";;"
                                end select


				                end if
				    
				                if media <> "export" then%>
				                </td>
				                <%end if
				
				                else
				                    if media <> "export" then %>
				                    <td style="border-top:1px #cccccc solid;"><img src="ill/blank.gif" width="1" height="1" border="0" /></td>
				                    <td style="border-top:1px #cccccc solid;"><img src="ill/blank.gif" width="1" height="1" border="0" /></td>
				                    <%end if
				            end if %>
				
				
				            <%
				            '*** Kostpriser ***'
				            if level = 1 AND cint(visKost) = 1 then
				
				            if media <> "export" then%>
				            <td style="padding-top:3px; padding-right:5px; white-space:nowrap; border-top:1px #cccccc solid;" align="right" valign="top">
				            <%end if
				    
				                'call akttyper2009Prop(oRec("tfaktim"))
				                '*** Alle fakturerbare ***'
				   
				                if cint(aty_fakbar) = 1 OR cint(aty_medpafak) = 1 then

				                    if media <> "export" then
				                    %>
    				                <%=formatnumber(oRec("kostpris"), 2)%>
				                    <%end if 

                                select case ver 
                                case 1
    			                case else
				                ekspTxt = ekspTxt & formatnumber(oRec("kostpris"), 2) &";"
                                end select

    				            else
    				                if media <> "export" then%>
    				                &nbsp;
    				                <%end if

                                    select case ver 
                                    case 1
    			                    case else
    				                ekspTxt = ekspTxt & ";"
                                    end select

				                end if
				
				             if media <> "export" then
			                 %>
				             </td>
			                 <td style="padding-top:3px; padding-right:5px; white-space:nowrap; border-top:1px #cccccc solid;" align="right" valign="top">
				             <%
				             end if
				 
				             if cint(aty_fakbar) = 1 OR cint(aty_medpafak) = 1 then%>
				    
				                <%if useenheder = 1 then
				                kostTot = oRec("kostpris")*enheder
				                else
				                kostTot = oRec("kostpris")*oRec("timer")
				                end if 


                                call valutakode_fn(oRec("kpvaluta"))

				    
				                 if media <> "export" then%>
				                 <%=formatnumber(kostTot, 2)&" "& valutaKode_CCC%> 
			                     <%
                                 'basisValISO    
                                 end if

                                 select case ver 
                                case 1
    			                case else
			                     ekspTxt = ekspTxt & formatnumber(kostTot, 2) &";"& valutaKode_CCC &";" 'basisValISO
                                end select
    			     
    			                 else
    			                  if media <> "export" then%>
    			                  &nbsp;
    			                  <%end if

                                 select case ver 
                                case 1
    			                case else
    			                ekspTxt = ekspTxt & ";;"
                                end select
				  
				              end if 'aty_fakbar
				    
				                 if media <> "export" then%>
				                </td>
				                <%
				                end if
				            else 
				
				                 if media <> "export" then%>
				                 <td style="border-top:1px #cccccc solid;"><img src="ill/blank.gif" width="1" height="1" border="0" /></td>
				                <td style="border-top:1px #cccccc solid;"><img src="ill/blank.gif" width="1" height="1" border="0" /></td>
				                <%end if
				
				            end if
				
				
				             if media <> "export" then
				    
				                %>
				                 <!-- Taste dato -->
				                <%if cint(hidegkfakstat) <> 1 then  %>
				                <td style="padding-top:3px; padding-right:5px; border-top:1px #cccccc solid;" align="right" valign="top"><font class="megetlillesilver"><%=oRec("TasteDato")%><br />
                                <i><%=oRec("editor") %></i></font></td>
				                <%else %>
				                <td style="border-top:1px #cccccc solid;"><img src="ill/blank.gif" width="1" height="1" border="0" /></td>
				                <%end if 
				
                            else
                            
                              if cint(hidegkfakstat) <> 1 then

                                select case ver 
                                case 1
    			                case else
                                ekspTxt = ekspTxt & oRec("TasteDato") &";"
                                end select

                              end if

				            end if%>
				
				            <%if cint(hidegkfakstat) <> 1 then 
				                if media <> "export" then%>
				                <td class=lille valign=top style="padding:3px 1px 3px 3px; border-top:1px #cccccc solid; white-space:nowrap;">
                                <%end if
                    
                                 erGk = 0
                                 gkCHK0 = ""
                                 gkCHK1 = ""
                                 gkCHK2 = ""
                                 gkCHK3 = ""
                                 gk3bgcol = "#999999"
                                 gk2bgcol = "#999999"
                                 gk1bgcol = "#999999"
                                 gk0bgcol = "#999999"

                             select case cint(oRec("godkendtstatus"))
                             case 2
                             gkCHK2 = "CHECKED"
                             erGk = 2
                             gk2bgcol = "red"
                             case 1
                             gkCHK1 = "CHECKED"
                             erGk = 1
                             gk1bgcol = "green"
                             case 3
                             gkCHK3 = "CHECKED"
                             erGk = 3
                             gk3bgcol = "orange"
                             case else
                             gkCHK0 = "CHECKED"
                             erGk = 0
                             gk0bgcol = "#000000"
                             end select
                 
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
                 
                    
                                if media <> "export" then
                 
					            if level = 1 OR _
					            cint(session("mid")) = cint(jobans1) OR _
					            cint(session("mid")) = cint(jobans2) then


                                '** + evt. temaledere ***'
					
					                '*** Godkendelse ***'
					                 if print <> "j" then%>
    					
					                 <input name="ids" id="ids" value="<%=oRec("tid")%>" type="hidden" />
					               <span style="color:<%=gk1bgcol%>;"><%=joblog2_txt_048 %>:</span><input type="radio" name="FM_godkendt_<%=oRec("tid")%>" id="FM_godkendt1_<%=v%>" class="FM_godkendt_1" value="1" <%=gkCHK1 %>><br />
                                   <span style="color:<%=gk3bgcol%>;"><%=joblog2_txt_049 %>:</span><input type="radio" name="FM_godkendt_<%=oRec("tid")%>" id="FM_godkendt3_<%=v%>" class="FM_godkendt_3" value="3" <%=gkCHK3 %>><br />
                                    <span style="color:<%=gk2bgcol%>;"><%=joblog2_txt_050 %>:</span><input type="radio" name="FM_godkendt_<%=oRec("tid")%>" id="FM_godkendt2_<%=v%>" class="FM_godkendt_2" value="2" <%=gkCHK2 %>><br />
                                   <span style="color:<%=gk0bgcol%>;"><%=joblog2_txt_051 %>:</span><input type="radio" name="FM_godkendt_<%=oRec("tid")%>" id="FM_godkendt0_<%=v%>" value="0" class="FM_godkendt_0" <%=gkCHK0 %>><br />

                                    <!--<input name="FM_godkendt" id="FM_godkendt_hd_<=v>" value="#" type="hidden" />-->
					                <%
					                else
    					    
					        
    					    
					                end if
					
					            visopdaterknap = 1
					            else
					                select case oRec("godkendtstatus")
                                    case 1
				                    %>
				                    Godkendt
				                    <%
                                    case 2
                                     %>
				                    <span style="color:red;"><%=joblog2_txt_050 %>!</span>
                                   <% case 3
                                     %>
				                    <%=joblog2_txt_052 %>
				                    <%
                                    case else
				        
				                    end select
					            end if
					
					            end if 'Media

					

					            'if len(trim(oRec("godkendtstatusaf"))) <> 0 then
					                'if media <> "export" then
					    
					                '<span style="font-size:9px;"><i><%=left(oRec("godkendtstatusaf"), 15)></i></span>
					     
					                'end if
					            erGkaf = oRec("godkendtstatusaf")
					            'end if 
					

			                    if media <> "export" then%>
				                </td>
				                <td valign="top" style="padding:3px 5px 3px 3px; border-top:1px #cccccc solid;">
				                <%
				                end if
				
				            erFaktureret = 0
				
				            if cdate(lastfakdato) >= cdate(oRec("Tdato")) then 
				
				            erFaktureret = 1%>
				    
				                <%
				                 if media <> "export" then
				     
				                    if (level = 1 OR _
					                cint(session("mid")) = cint(jobans1) OR _
					                cint(session("mid")) = cint(jobans2)) AND (print <> "j") then%>
					                <a href="erp_fakhist.asp?FM_kunde=<%=oRec("kid")%>&FM_job=<%=jobid%>" class=vmenu target="_blank"><%=joblog2_txt_054 %></a>
					                <%else %>
					                <%=joblog2_txt_054 %>
					                <%
					                end if
					
					            end if %>
				
				
				
				            <%end if 
				
				             if media <> "export" then%>
				            </td>
				            </tr>
				            <%end if %>
				
				
				            <%
				
				
                             select case ver 
                                case 1
    			                case else
				            
				                ekspTxt = ekspTxt & erGk &";"& erGkaf &";"& erFaktureret &";"
				            end select
				
				            else 
				                 if media <> "export" then
				                %>
				                <td style="border-top:1px #cccccc solid;"><img src="ill/blank.gif" width="1" height="1" border="0" /></td>
				                <td style="border-top:1px #cccccc solid;"><img src="ill/blank.gif" width="1" height="1" border="0" /></td>
				                <%end if
				
				            end if '** gkstat **'
				
				
				
				            komm_note_Txt = ""
                            timerKom = ""
                            timerKom = oRec("timerkom")
				            if len(timerKom) <> 0 then
				            komm_note_Txt = timerKom & kundeKomm
				            end if
                
                
                
                           'call htmlreplace(komm_note_Txt)
                           'htmlparseCSVtxt = htmlparseTxt

                           select case ver 
                           case 1
    			           case else

                                   if len(komm_note_Txt) <> 0 then
                                    'komm_note_Txt = replace(komm_note_Txt, " -", "-")
                                    call htmlparseCSV(komm_note_Txt)
                                    'komm_note_Txt = replace(htmlparseCSVtxt, "vbcrlf", "xxxxx")
                                    'komm_note_Txt = replace(komm_note_Txt, "vbcrlf", "xxxxx")
                                    'komm_note_Txt = replace(komm_note_Txt, "''", "")
                                    'if right(trim(komm_note_Txt), 1) = "&quot;" OR right(trim(komm_note_Txt), 1) = "&#34;" OR right(trim(komm_note_Txt), 1) = chr(34) OR right(trim(komm_note_Txt), 1) = "''" then
                                    'addPkt = "#"
                                    'else
                                    'addPkt = ""
                                    'end if
                                    
                                    ekspTxt = ekspTxt & ""& Chr(34) & komm_note_Txt & " " & Chr(34) &";" 'Der skal v�re et mellemum inden sidste chr(34) da der eller vil komme ombrydninsfejl hvis teksten slutter med et "
                                   else
                                    ekspTxt = ekspTxt & ";"
                                   end if
                                    
                           end select%>
				
				
				
				
				
				
			            <%
                        if cint(ver) = 1 then
                        ekspTxt = ekspTxt 
                        else
                        ekspTxt = ekspTxt & "xx99123sy#z"
                        end if

			             
			
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
			            lastFase = oRec("fase")
			            'lastYearNum = 
	        
	        
	        
	                    Response.flush
	       
	       
	                   end if 'hidesygtyper = 1

            end if 'if cint(joblog_uge) = 3

           
            m = m + 1
            

            if lastM <> oRec("tmnr") then
            antalM = antalM + 1
            end if
	       
            lastM = oRec("tmnr")
        
            x = x + 1
	        oRec.movenext
	        wend
	        oRec.Close
	        'Set oRec = Nothing
            m = 0

      
        



     if media <> "export" then
    
        if cint(joblog_uge) = 1 then
    
    
	            if v > 0 AND fordelpamedarb = 1 then
        	        call fordelpamedarbejdere()
	            end if
        	
	            if v > 0 AND fordelpamedarb = 2 then
                    call fordelpaakt(1)
        	    end if
	
	
	    end if
	
    end if


    
      if cint(joblog_uge) = 3 then 'm�nedsoversigt

        if x <> 0 then


        call maanedsoversigtArr
        end if
     end if


       



                if media <> "export" then

	
	                if x = 0 then
	                %>
                    <br /><br /><br />
                    <table border=0 cellspacing=0 cellpadding=0 width="925">
	                <tr>
                        <td style="padding: 20px; background-color: #FFFFFF;">
                            <h4><%=joblog2_txt_055 %>:</h4>
                            <%=joblog2_txt_056 %><br>
                            <br />
                            <%=joblog2_txt_057 &" " %><b><%=joblog2_txt_058 &" "%></b> <%=joblog2_txt_059 %>
                            &nbsp;</td>
	                </tr>
                    </table>
                                    <br /><br><BR />&nbsp;
	                <%

                    response.end
	                end if
	
	
	

                end if 'media

      

if x <> 0 then

    if media <> "export" then

        if cint(joblog_uge) = 2 then
        call dagstotaler(3)
        end if
        

          if cint(joblog_uge) <> 3 then

                    call jobtotaler()

                    if media = "print" then

                        select case lto 
                        case "mmmi"
                        case else 
                        call grandTotal
                        end select

                    else
            
                    call grandTotal

                    end if

        else


                    

            
                    'select case lto 
                    'case "mmmi", "intranet - local"
                    'case else

                    if print = "j" then

                    call underskrift
                    
                    end if
                        
                        
                      

                    'end select
            

          end if

        %><br /><br /><br /><%
    
    end if





                '******************* Eksport **************************' 
                if media = "export" then



                  
                                         if cint(joblog_uge) = 3 then

                                         select case lto 
                                          case "mmmi", "intranet - local"
                                         call underskriftExcel  

                        
                                         end select

                                         end if
                      

    
   
    
                    call TimeOutVersion()
    


                    hidetimepriser = request.cookies("stat")("hidetimepriser")
                    hideenheder = request.cookies("stat")("hideenheder")
                    hidegkfakstat = request.cookies("stat")("hidegkfakstat")
                    hidefase = request.cookies("stat")("hidefase")

                    'Response.Write "hidefase: " & hidefase
                    'Response.end 
        
                    if cint(ver) = 1 then
                    ekspTxt = replace(ekspTxt, "xx99123sy#z", vbcrlf)
                    else
                    ekspTxt = replace(ekspTxt, "xx99123sy#z", vbcrlf)
                    end if

	                
	
	                datointerval = request("datointerval")
	
	
	                filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	                filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)


                    if cint(ver) = 1 then
                    fileext = "txt"
                    else
                    fileext = "csv"
                    end if

				    
                    
	
				                Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				                if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\joblog.asp" then
					                Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\joblogexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"."& fileext, True, False)
					                Set objNewFile = nothing
					                Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\joblogexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"."& fileext, 8)
				                else
					                Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\joblogexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"."& fileext, True, False)
					                Set objNewFile = nothing
					                Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\joblogexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"."& fileext, 8)
				                end if
				
				
                                file = "joblogexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"."& fileext
				
				                '**** Eksport fil, kolonne overskrifter ***
                                if cint(joblog_uge) = 3 then 'm�nedsoversigt
                                strOskrifter = joblog2_txt_081&"; "&joblog2_txt_151&"; "&joblog2_txt_152&"; "&joblog2_txt_083&"; "&joblog2_txt_153&"; "&joblog2_txt_019&"; "&joblog2_txt_154&"; "&joblog2_txt_086&"; "&joblog2_txt_114&";"

                                      strOskrifter = strOskrifter & strExportOskriftDage


                                else

				
				                    if cint(ver) = 1 then
                                    strOskrifter = "" 'Bf: Blank
				                    'strOskrifter = "Dato;tekst;bilag;konto;kontonavn;debit bel�b;kredit bel�b;modkonto"
                                    else

				                    strOskrifter = joblog2_txt_083&";"&joblog2_txt_155&";"&joblog2_txt_019&";"&joblog2_txt_156&";"
                
                                            if cint(hidefase) <> 1 then
                                            strOskrifter = strOskrifter &joblog2_txt_082&";"
                                            end if
                
                                                strOskrifter = strOskrifter & joblog2_txt_157&";"&joblog2_txt_072&";"&joblog2_txt_113&";"&joblog2_txt_158&";"&joblog2_txt_159&";Task No.;"
                 

                                                if cint(showfor) = 1 then
                                                strOskrifter = strOskrifter & joblog2_txt_160&";"
                                                end if
                 
                                                strOskrifter = strOskrifter & joblog2_txt_081&";"&joblog2_txt_161&";"&joblog2_txt_162&";"
                
                                                if cint(showcpr) = 1 then
                                                strOskrifter = strOskrifter & joblog2_txt_152&";"
                                                end if

                                            strOskrifter = strOskrifter & joblog2_txt_092&";"&joblog2_txt_163&";"
				                    end if


                        
                                        select case ver
                                        case 1

                                        case else                
				
				                                if cint(hideenheder) = 0 then
				                                strOskrifter = strOskrifter & joblog2_txt_093&";"
				                                end if
				
				
				                                if (level <=2 OR level = 6) AND cint(hidetimepriser) = 0 then
				                                strOskrifter = strOskrifter & joblog2_txt_164&";"&joblog2_txt_165&";"&joblog2_txt_166&";"
				                                end if 
				
				                                if level = 1 AND visKost = 1 then 
				                                strOskrifter = strOskrifter & joblog2_txt_096&";"&joblog2_txt_097&";"&joblog2_txt_166&";"
				                                end if
				

                                                if cint(hidegkfakstat) <> 1 then
                                                  strOskrifter = strOskrifter & joblog2_txt_098&";"
                                                end if
                                

				                                if lto <> "execon" AND cint(hidegkfakstat) = 0 then
				                                strOskrifter = strOskrifter  & joblog2_txt_167&";"&joblog2_txt_168&";"&joblog2_txt_169&";"
				                                end if
				
				                                strOskrifter = strOskrifter  & joblog2_txt_114&";"

                                        end select

			
                                end if	

				                select case ver
                                case 1
				                case else
                                objF.writeLine(joblog2_txt_170&": "& datointerval & vbcrlf)
				                objF.WriteLine(strOskrifter) '& chr(013)
                                end select
				                objF.WriteLine(ekspTxt)
				                objF.close
				
				                %>
				
	                            <table border=0 cellspacing=1 cellpadding=0 width="200">
	                            <tr><td valign=top bgcolor="#ffffff" style="padding:5px;">
	                            <img src="../ill/outzource_logo_200.gif" />
	                            </td>
	                            </tr>
	                            <tr>
	                            <td valign=top bgcolor="#ffffff" style="padding:5px 5px 5px 15px;">
	                            <a href="../inc/log/data/<%=file%>" class=vmenu target="_blank" onClick="Javascript:window.close()"><%=joblog2_txt_060 %> >></a>
	                            </td></tr>
	                            </table>
	            
	          
	            
	                            <%
                
                
                                Response.end
	                            'Response.redirect "../inc/log/data/"& file &""	
				



                end if 'media%>






	



            <%if print <> "j" then%>

              <!--pagehelp-->

                            <%

                            itop = hlp_top
                            ileft = 635
                            iwdt = 120
                            ihgt = 0
                            ibtop = 40 
                            ibleft = 150
                            ibwdt = 600
                            ibhgt = 310
                            iId = "pagehelp"
                            ibId = "pagehelp_bread"
                            call sideinfoId(itop,ileft,iwdt,ihgt,iId,phDsp,phVzb,ibtop,ibleft,ibwdt,ibhgt,ibId)
                            %>
                       
			                            <b><%=joblog2_txt_061 %>: </b><br> <%=joblog2_txt_062 %><br>
			                            <%=joblog2_txt_063 %><br><br>
			                            <b><%=joblog2_txt_064 %></b><br /> <%=joblog2_txt_065 %><br>
			                            <%=joblog2_txt_066 &" "%> 
			                            <%=joblog2_txt_067 %><br />
			                            <br />
			                            <%=joblog2_txt_068 %><br />
			                            <br />
			                            <b>Key-account</b><br />
                                        <%=joblog2_txt_069 %> 
                                        <br /><br />
                                        <b><%=joblog2_txt_070 %></b><br />
                                        <%=joblog2_txt_071 &" "%> <%=basisValISO %>.&nbsp;
                		
                		
		                            </td>
	                            </tr></table></div>
	



	            <br><br><br><br><br><br><br><br><br>&nbsp;
            <%end if 




            if print <> "j" then

            ptop = pr_top
            pleft = 840
            pwdt = 140

            pnteksLnk = "FM_segment="&segment&"&viskunabnejob0="&viskunabnejob0&"&viskunabnejob1="&viskunabnejob1&"&viskunabnejob2="&viskunabnejob2
            pnteksLnk = pnteksLnk & "&FM_orderby_medarb="&fordelpamedarb&"&datointerval="&strDag&"/"&strMrd&"/"&strAar & " - " & strDag_slut&"/"&strMrd_slut&"/"&strAar_slut
            pnteksLnk = pnteksLnk & "&rdir="&rdir &"&FM_kunde="&kundeid &"&menu=stat&jobnr="&intJobnr&"&eks="&request("eks")&"&lastFakdag="&lastFakdag&"&selmedarb="&selmedarb&"&selaktid="&selaktid
            pnteksLnk = pnteksLnk & "&FM_job="&request("FM_job")&"&FM_start_dag="&strDag&"&FM_start_mrd="&strMrd&"&FM_start_aar="&strAar&"&FM_slut_dag="&strDag_slut&"&FM_slut_mrd="&strMrd_slut&"&FM_slut_aar="&strAar_slut
            pnteksLnk = pnteksLnk & "&FM_kundejobans_ell_alle="&visKundejobans&"&FM_jobsog="&jobSogVal&"&FM_akttype="&vartyper&"&nomenu="&nomenu

            call eksportogprint(ptop,pleft,pwdt)

            %>

        
        
        
                   
                    
                  <tr>
                    <td>
                   <!-- <a href="#" onclick="Javascript:window.open('joblog.asp?media=export&<%=pnteksLnk%>', '', 'width=350,height=150,resizable=no,scrollbars=no')" class=rmenu><img src="../ill/export1.png" border=0></a>
                    </td><td><a href="#" onclick="Javascript:window.open('joblog.asp?media=export&<%=pnteksLnk%>', '', 'width=350,height=150,resizable=no,scrollbars=no')" class=rmenu>.csv fil eksport</a>
                       -->

                       

                         

                        <%select case lto 
                        case "bf",  "intranet - local"

                        %>

                         <form action="joblog.asp?media=export&ver=0&cur=0&<%=pnteksLnk%>" target="_blank" method="post"> 
                                <input type="hidden" name="FM_medarb_hidden" value="<%=thisMiduse%>" />
                             <input type="hidden" name="FM_medarb" value="<%=thisMiduse%>" />
                        <input type="submit" id="sbm_csv" value="<%=joblog2_txt_143 %> >>" style="font-size:9px;" />
                               </form>

                         <form action="joblog.asp?media=export&ver=1&cur=0&<%=pnteksLnk%>" target="_blank" method="post"> <br />

                             <input type="hidden" name="FM_medarb_hidden" value="<%=thisMiduse%>" />
                             <input type="hidden" name="FM_medarb" value="<%=thisMiduse%>" />
                             <input type="submit" id="sbm_csv" value="<%=joblog2_txt_144 %> >>" style="font-size:9px;" />
                        </form>
                   
                        <form action="joblog.asp?media=export&ver=1&cur=1&<%=pnteksLnk%>" target="_blank" method="post"> 
                           <input type="hidden" name="FM_medarb_hidden" value="<%=thisMiduse%>" />
                             <input type="hidden" name="FM_medarb" value="<%=thisMiduse%>" />
                            <br />
                        <input type="submit" id="sbm_csv" value="<%=joblog2_txt_145 %> >>" style="font-size:9px;" />
                        <%case else %>
                        <form action="joblog.asp?media=export&ver=0&cur=0&<%=pnteksLnk%>" target="_blank" method="post"> 
                           <input type="hidden" name="FM_medarb_hidden" value="<%=thisMiduse%>" />
                             <input type="hidden" name="FM_medarb" value="<%=thisMiduse%>" />
                        <input type="submit" id="sbm_csv" value="<%=joblog2_txt_143 %> >>" style="font-size:9px;" />
                        <%end select%>

                    </td>
                   </tr>
               
                </form>
                  <form action="joblog.asp?print=j&cur=0&<%=pnteksLnk%>" target="_blank" method="post"> 
                       <input type="hidden" name="FM_medarb_hidden" value="<%=thisMiduse%>" />
                             <input type="hidden" name="FM_medarb" value="<%=thisMiduse%>" />
                 <tr>
               <td valign="top" style="padding-top:10px;">
                    <input type="submit" value="<%=joblog2_txt_146 %> >>" style="font-size:9px;" />
                   
                    <!--
                   <a href="joblog.asp?print=j&<%=pnteksLnk%>" target="_blank"  class='rmenu'>
                  &nbsp;<img src="../ill/printer3.png" border=0 alt="" /></a>
                  </td><td><a href="joblog.asp?print=j&<%=pnteksLnk%>" target="_blank" class="rmenu">Print version</a>

                       -->
                     </td>
               </tr>
                  </form>
	
               </table>
            </div>
            <%else%>

            <% 
            Response.Write("<script language=""JavaScript"">window.print();</script>")
            %>
            <%end if%>



<%end if 'x <> 0 %>


</div>
<br>
<br>
&nbsp;
<%'end if 'validering %>

<!--#include file="../inc/regular/footer_inc.asp"-->