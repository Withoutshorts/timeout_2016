


<%response.buffer = true 
Session.LCID = 1030
%>
			        

<!--#include file="../inc/connection/conn_db_inc.asp"-->


<%'** JQUERY START ************************* %>

<%'** JQUERY END ************************* %>
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../timereg/inc/isint_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->



<%
if len(session("user")) = 0 then
	errortype = 5
	call showError(errortype)
        response.End
	end if
	

	
	'** Faste filter kri ***'
	thisfile = "fomr"
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
    call GetSytemSetup()
	
	call fomr_account_fn()         
%>
    
    <%call menu_2014 %>
    <div id="wrapper">
        <div class="content">

    <%
  call fomr_account_fn()
    select case func 
    case "slet"

        oskrift = fomr_txt_001
        slttxta = fomr_txt_002 & " <b>"&fomr_txt_002&"</b> " & fomr_txt_004
        slttxtb = "" 
        slturl = "fomr.asp?func=sletok&id="& id

        call sletcnf_2015(oskrift, slttxta, slttxtb, slturl)



	case "sletok"
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM fomr WHERE id = "& id &"")
    oConn.execute("DELETE FROM fomr_rel WHERE for_fomr = "& id &"")

	Response.redirect "fomr.asp?menu=tok&shokselector=1"
%>


<%

    case "dbopr", "dbred"
	'*** Her inds�ttes en ny type i db ****
		if len(request("FM_navn")) = 0 then
		
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

        business_area_label = SQLBless(request("FM_business_area_label"))
        business_unit = SQLBless(request("FM_business_unit"))

		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)

        if len(trim(request("FM_aktok"))) <> 0 then
        aktok = 1
		else
        aktok = 0
        end if


        if len(trim(request("FM_jobok"))) <> 0 then
        jobok = 1
		else
        jobok = 0
        end if

        if len(trim(request("FM_konto"))) <> 0 then
        konto = request("FM_konto")
        else
        konto = 0
        end if


		fomr_segment = request("FM_fomr_segment")


		if func = "dbopr" then
		oConn.execute("INSERT INTO fomr (navn, editor, dato, jobok, aktok, konto, business_unit, business_area_label, fomr_segment) VALUES "_
        &" ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', "& jobok &", "& aktok &", "& konto &", '"& business_unit &"', '"& business_area_label &"', '"& fomr_segment &"')")
		else
		oConn.execute("UPDATE fomr SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"', jobok = "& jobok &", aktok = "& aktok &", "_
        &" konto = "& konto &", business_unit = '"& business_unit &"', business_area_label = '"& business_area_label &"', fomr_segment = '"& fomr_segment &"' WHERE id = "& id &"")
		end if
		
		Response.redirect "fomr.asp?menu=tok&shokselector=1"
		end if



    case "red", "opret"
    '*** Her indl�ses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	'varSubVal = "Opret" 
	'varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
    select case lto
    case "a27"
    jobokCHK = ""
	aktokCHK = "CHECKED"
    business_area_label = "progress-bar-warning"
    case else
    jobokCHK = "CHECKED"
	aktokCHK = ""
    business_area_label = ""
    end select
    konto = 0

    fomr_segment = ""

	else
	strSQL = "SELECT navn, editor, dato, konto, aktok, jobok, business_unit, business_area_label, fomr_segment FROM fomr WHERE id= " & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("navn")
	strDato = formatdatetime(oRec("dato"), 2)
	strEditor = oRec("editor")


    business_unit = oRec("business_unit")
    business_area_label = oRec("business_area_label")

    if oRec("jobok") = 1 then
    jobokCHK = "CHECKED"
    else
    jobokCHK = ""
    end if

	if oRec("aktok") = 1 then
    aktokCHK = "CHECKED"
    else
    aktokCHK = ""
    end if

    konto = oRec("konto")

    fomr_segment = oRec("fomr_segment")

	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "Opdater" 
	end if
    
%>

<%if useTrainerlog = 1 then %>
    <style>
        .removefromtrailerlog {
            display:none;
        }
    </style>
<%end if %>

<!--Forretningsomr�der redigering-->
<div class="container">
    <div class="portlet">
        <h3 class="portlet-title">
          <u><%=fomr_txt_005 %></u>
        </h3>

        <form action="fomr.asp?menu=tok&func=<%=dbfunc%>" method="post">
             <input type="hidden" name="id" value="<%=id%>">
        <div class="row">
            <div class="col-lg-10">&nbsp</div>
            <div class="col-lg-2 pad-b10">
            <button type="submit" class="btn btn-success btn-sm pull-right"><b><%=fomr_txt_005 %></b></button>
            </div>
        </div>

        <div class="portlet-body">
            
            <section>
              <div class="well well-white">  
            <div class="row">
                <div class="col-lg-1">&nbsp</div>
                <div class="col-lg-2"><%=fomr_txt_007 %>:&nbsp<span style="color:red;">*</span></div>
                <div class="col-lg-3"><input type="text" name="FM_navn" class="form-control input-small" value="<%=strNavn %>" placeholder="<%=fomr_txt_007 %>"/></div>
                <div class="col-lg-1"><%=fomr_txt_008 %>:</div><div class="col-lg-2"><input type="text" name="FM_business_area_label" class="form-control input-small" value="<%=business_area_label %>" placeholder="<%=fomr_txt_008 %>"/></div>
            </div>
            <div class="row">
                <div class="col-lg-1">&nbsp</div>
                <div class="col-lg-2"><%=fomr_txt_009 %>:</div>
                <div class="col-lg-3"><input type="text" name="FM_business_unit" class="form-control input-small" value="<%=business_unit %>" placeholder="<%=fomr_txt_010 %>"/></div>
            </div>



            <div class="row">
                <div class="col-lg-1">&nbsp</div>
                <div class="col-lg-2 pad-t20"><b><%=fomr_txt_011 %>:</b></div>
            </div>
            <div class="row">
                <div class="col-lg-1">&nbsp</div>
                    <div class="col-lg-2"><%=fomr_txt_012 %></div> <div class="col-lg-1"><input type="checkbox" name="FM_jobok" <%=jobokCHK %>></div>
            </div>
            <div class="row">
                <div class="col-lg-1">&nbsp</div>
                    <div class="col-lg-2"><%=fomr_txt_013 %></div> <div class="col-lg-1"><input type="checkbox" name="FM_aktok" <%=aktokCHK %>></div>
            </div>


            <%if cint(fomr_account) = 0 then %>
            <div class="row pad-b20">
                <div class="col-lg-1">&nbsp</div>
                <div class="col-lg-2 pad-t20"><%=fomr_txt_014 %>:</div>
              <div class="col-lg-3 pad-t20"><select class="form-control input-sm" name="FM_konto">
              <option value="0"><%=fomr_txt_015 %></option>

                <%
                    licensindehaverKid = 0
                    call licKid()

                strSQLkonto = "SELECT navn, kontonr, id FROM kontoplan WHERE kid = "& licensindehaverKid
                    oRec2.open strSQLkonto, oConn, 3
                    while not oRec2.EOF 
            
                    if cint(konto) = oRec2("id") then
                    kontoSEL = "SELECTED"
                    else
                    kontoSEL = ""
                    end if

           
                    kontonrVal = " ("& oRec2("kontonr") &")"
            
         
                    %><option value="<%=oRec2("id") %>" <%=kontoSEL %>><%=oRec2("navn") & kontonrVal%></option><%

                     oRec2.movenext
                     wend
                     oRec2.close  %>
                  
                    </select></div>
                    </div>
                  <% end if %>

                  <div class="row pad-t20 removefromtrailerlog">
                      <div class="col-lg-1">&nbsp</div>
                      <div class="col-lg-6"><%=fomr_txt_016 %>:</div>
                  </div>
                        <div class="row removefromtrailerlog">
                            <div class="col-lg-1">&nbsp</div>
                            <div class="col-lg-4">
                            <%strSQL = "SELECT id, navn FROM kundetyper WHERE id <> 0 ORDER BY navn"
                            seg = 0
                            oRec.open strSQL, oConn, 3
                            while not oRec.EOF 

            
                            if instr(fomr_segment, "#"& oRec("id") &"#") <> 0 then
                            fomr_segmentCHK = "CHECKED"
                            else
                            fomr_segmentCHK = ""
                            end if

                        %>
                        <input type="checkbox" name="FM_fomr_segment" value="#<%=oRec("id") %>#" <%=fomr_segmentCHK %> /> <%=oRec("navn") %><br />
                        <%
                        seg = seg + 1
                        oRec.movenext
                        wend 
                        oRec.close

                        if seg = 0 then
                        %>

                        - <%=fomr_txt_017 %>

                        <%
                        end if

                        %>
                            </div>
                        </div>
                        
                
                 </div>
               
                    </section>
            <%if dbfunc = "dbred" then%> 
             <br /><br /><br /><div style="font-weight: lighter;"><%=fomr_txt_018 %> <b><%=strDato%></b> <%=fomr_txt_019 %> <b><%=strEditor%></b></div>
            <%end if  %>    
                </div>
             </form>
            </div>
        </div>
  <%case "stat"%>
	<!--Skal arbejdes videre med-->
	
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	



    <script src="inc/fomr_jav.js"></script>


	<%

        call menu_2014

    
	
	
	'**************************************************
	'***** Faste Filter kriterier *********************
	'**************************************************
		
	
	'*** Medarbejdere / projektgrupper
	selmedarb = session("mid")
	
	'*** Job og Kundeans ***
	call kundeogjobans()
	
    'call medarbogprogrp("oms")
	medarbSQlKri = ""
	kundeAnsSQLKri = ""
	jobAnsSQLkri = ""
	jobAns2SQLkri = ""
	
	if len(trim(request("FM_progrp"))) <> 0 then
	progrp = request("FM_progrp")
	else
	progrp = 0
	end if
	
    if len(trim(request("FM_fordel_jobakt"))) <> 0 then
    fordel_jobakt = request("FM_fordel_jobakt")
    else
    fordel_jobakt = 0
    end if

    if cint(fordel_jobakt) = 0 then
        fordel_jobakt0CHK = "CHECKED"
    else
        fordel_jobakt1CHK = "CHECKED"
    end if

	'Response.Write "medid first: "& left(request("FM_medarb"), 1)
	'Response.end
	
	'*** Rettigheder p� den der er logget ind **'
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
	
	
	media = request("media")
	
	'**** Kundekri ***
	if len(request("FM_kunde")) <> 0 then
	kundeid = request("FM_kunde")
	else
	kundeid = 0
	end if
	
	'*** Kundeans ***
	'strKnrSQLkri = ""
	strKnrSQLkri = " OR jobknr = 0 "
	
	'*** finder udfra valgte projektgrupper og medarbejdere
	'medarbSQlKri 
	'kundeAnsSQLKri
	
			    for m = 0 to UBOUND(intMids)
			    
			     if m = 0 then
			     medarbSQlKri = "(m.mid = " & intMids(m)
			     kundeAnsSQLKri = "kundeans1 = "& intMids(m) &" OR kundeans2 = "& intMids(m) &""
			     jobAnsSQLkri = "jobans1 = "& intMids(m)  
			     jobAns2SQLkri = "jobans2 = "& intMids(m)
			     else
			     medarbSQlKri = medarbSQlKri & " OR m.mid = " & intMids(m)
			     kundeAnsSQLKri = kundeAnsSQLKri & " OR kundeans1 = "& intMids(m) &" OR kundeans2 = "& intMids(m) &""
			     jobAnsSQLkri = jobAnsSQLkri & " OR jobans1 = "& intMids(m)  
			     jobAns2SQLkri = jobAns2SQLkri & " OR jobans2 = "& intMids(m)
			     end if
			     
			    next
			    
			    medarbSQlKri = medarbSQlKri & ")"
			    
			jobAnsSQLkri =  "AND ("& jobAnsSQLkri &")"
			jobAns2SQLkri =  "AND (" & jobAns2SQLkri &")"
			
	
	'** Er key acc og kundeansvarlig valgt **?
	if cint(kundeans) = 1 then
	kundeAnsSQLKri = kundeAnsSQLKri
	else
	kundeAnsSQLKri = " Kundeans1 <> -1 AND kundeans2 <> -1 "
	end if
	
	if len(request("FM_ignorerperiode")) <> 0 then
	ignper = request("FM_ignorerperiode")
	else
	ignper = 0
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
	
	
	
	'**** Alle SQL kri starter med NUL records ****
	jobidFakSQLkri = " OR jobid = -1 "
	jobnrSQLkri = " OR tjobnr = '-1' "
	jidSQLkri = " OR id = -1 "
	seridFakSQLkri = " OR aftaleid = -1 "
	
	
	
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
	
	
	
	
	
	
	'************ slut faste filter var **************		



    if media <> "export" then 
	ldTop = 400
	ldLft = 240
	else
	ldTop = 15
	ldLft = 10
	end if
	%>
	<!--<div id="load" style="position:absolute; display:; visibility:visible; top:<%=ldTop%>px; left:<%=ldLft%>px; width:300px; background-color:#ffffff; border:10px #9ACD32 solid; padding:10px; z-index:100000000;">
    <table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" /><br />
	&nbsp;&nbsp;&nbsp;&nbsp;Vent veligst. <br />&nbsp;&nbsp;&nbsp;&nbsp;Forventet loadtid: 3 - 20 sek...
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" />
	</td></tr></table>
	
	</div>-->
	<%
	
	


	
	timermedarbSQLkri = replace(medarbSQLkri, "m.mid", "tmnr")	
	
	
	if request("print") <> "j" then
	pleft = 90
	ptop = 102
	'ptopgrafik = 348
	else
	pleft = 20
	ptop = 20
	'ptopgrafik = 90
	end if
	%>	

	
	
	<%
	
	oimg = "pie_chart_48_hot.png"
	oleft = 0
	otop = 0
	owdt = 300
	oskrift = fomr_txt_001
	
    if request("print") = "j" then
	media = "print"
        call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
    end if
	
	if request("print") <> "j" then 
	
	call filterheader_2013(0,0,800,oskrift) %>
    <%end if %>
	    <%call medkunderjob %>	    
	  <div class="container">
	   <div class="portlet">
        

	    <form action="fomr.asp?rdir=<%=rdir %>&menu=stat&func=stat" method="post">
	
	    
	   <div class="portlet-body">
	        <div class="row">
            <div class="col-lg-1">&nbsp</div>
            <div class="col-lg-5"><%=fomr_txt_020 %>:</div>
	        </div>
        
            <div class="row">
            <div class="col-lg-1">&nbsp</div>
               <div class="col-lg-2"><input type="radio" name="FM_fordel_jobakt" value="0" <%=fordel_jobakt0CHK %> /> <%=fomr_txt_021 %><br /></div>
            </div>

            <div class="row">
            <div class="col-lg-1">&nbsp</div>
                <div class="col-lg-2"><input type="radio" name="FM_fordel_jobakt" value="1" <%=fordel_jobakt1CHK %> /> <%=fomr_txt_022 %></div>
            </div>
       
			<!-- Brug altid datointerval, FM_usedatokri = 1 -->
            <div class="row pad-t20">
                <div class="col-lg-1">&nbsp</div>
			    <div class="col-lg-2"><%=fomr_txt_023 %>:</div>
			</div>
            <div class="row">
                <div class="col-lg-1">&nbsp</div>
                <div class="col-lg-4"><input type="hidden" class="form-control input-small" name="FM_usedatokri" id="FM_usedatokri" value="1">
			    <!--#include file="../timereg/inc/weekselector_s.asp"--> <!-- b -->
                </div>
                <div class="col-lg-1">
                    <%if request("print") <> "j" then%>
			        <button class="btn btn-sm btn-success pull-right"><b> <%=fomr_txt_024 %> </b></button><br />&nbsp;
                    </div>
			        <%end if%>
                </div>
			</div>
			

			</form>

		
		
		<!-- fiulter DIV-->
		
		

	
	<%
	
	    '*** valgte job ***
		call valgtejob()
		
		
		'**** Valgte aftaler *****
		call valgteaftaler()
				
	
	
	    '*** For at spare (trimme) p� SQL hvis der v�lges alle job alle kunder og vis kun for jobanssvarlige ikke er sl�et til ****
		'*** Og der ikke er s�gt p� jobnavn ***
		'if cint(kundeid) = 0 AND cint(jobid) = 0 AND cint(jobans) = 0 AND cint(kundeans) = 0 AND len(trim(jobSogVal)) = 0  then 
		'if cint(kundeid) = 0 AND cint(jobid) = 0 AND cint(jobans) = 0 AND cint(jobans2) = 0 AND cint(jobans3) = 0 _
		'		 AND cint(kundeans) = 0 AND len(trim(jobSogVal)) = 0 AND cint(aftaleid) = 0 AND cint(segment) = 0 then 
				
				
		'	jidSQLkri =  " OR id <> 0 "
		'	jobnrSQLkri = " OR tjobnr <> '0' "
		'	'seridFakSQLkri = " OR aftaleid <> 0 
			
		'end if
	
	
		'**************** Trimmer SQL states ************************
		
		len_jobnrSQLkri = len(jobnrSQLkri)
		right_jobnrSQLkri = right(jobnrSQLkri, len_jobnrSQLkri - 3)
		jobnrSQLkri =  right_jobnrSQLkri
		
		len_jidSQLkri = len(jidSQLkri)
		right_jidSQLkri = right(jidSQLkri, len_jidSQLkri - 3)
		jidSQLkri =  right_jidSQLkri
		
		jidSQLKri = replace(jidSQLkri, "id", "aktiviteter.job")
		
		
		'*****************************************************************************************************
	
    Response.flush

	
	
	'*** Alle timer, uanset fomr. ***
	call akttyper2009(2)
	totaltimer = 0
	sqlDatoStart = strAar&"/"&strMrd&"/"&strDag
	sqlDatoSlut = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut
	
	'if ignper <> 1 then
	tdatoSQLkri = " AND tdato BETWEEN '"& sqlDatoStart &"' AND '"& sqlDatoSlut &"'"
	'else
	'tdatoSQLkri = ""
	'end if
	
    strJobSQLKri = " AND (j.jobnr = 0 "
    strJobTxt = ""

	strSQL = "SELECT sum(timer.timer) AS timertot, tjobnr, tjobnavn, tknavn FROM job AS j "_
    &" LEFT JOIN timer ON (tjobnr = j.jobnr) WHERE ("& aty_sql_realhours &") "_
    &" AND ("& timermedarbSQLkri &") AND ("& jobnrSQLkri &") "& jobstKri &" " & tdatoSQLkri & " GROUP BY tjobnr ORDER BY tknavn, tjobnavn"
	
    'if lto = "jttek" then
    'Response.Write strSQL
    'Response.flush
    'end if

	'Response.write "timermedarbSQLkri: "& timermedarbSQLkri & "<br>"& strSQL
	'Response.flush
	totaltimer = 0
    lastKnavn = ""
	oRec.open strSQL, oConn, 3 
	While not oRec.EOF 
	totaltimer = totaltimer + oRec("timertot")
	
    '** Led kun efter forretningeomr�der p� job der er registreret p� i periode
    strJobSQLKri = strJobSQLKri & " OR j.jobnr = '"& oRec("tjobnr") & "'"

    if lastKnavn <> oRec("tknavn") then
    strJobTxt = strJobTxt & "<br><br><b>"& oRec("tknavn") & "</b><br>"
    end if

    strJobTxt = strJobTxt & oRec("tjobnavn") & " ("& oRec("tjobnr") &"): " & oRec("timertot") & " timer<br>" 

    lastKnavn = oRec("tknavn")
    oRec.movenext
    wend
	oRec.close
	
	if totaltimer <> 0 then
	totaltimer = totaltimer
	else
	totaltimer = 1
	end if

    strJobSQLKri = strJobSQLKri & ")"
	
	'Response.write "Indtastet i alt i den valgte periode, p� de valgte medarbejdere og job,<br> uanset om aktivitet er tilknyttet et forretningsomr�de: <b>"& formatnumber(totaltimer, 2)&"</b> timer.<br>"

	
	%>
        <div class="portlet-body">
	    <div class="row">
            <div class="col-lg-1">&nbsp</div>
            <div class="col-lg-8"><%Response.write fomr_txt_025 & "<br>" & fomr_txt_046 & "<b>"& formatnumber(totaltimer, 2)&"</b> "& fomr_txt_047 &" <br>"%></div>
	    </div>
      </div>
	   </div>
     </div>
    </div>
	
    <%
	tTop = 0
    tLeft = 0
    tWdth = 810


    call tableDiv(tTop,tLeft,tWdth)

	%>
	<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<tr bgcolor="#5582D2">
		<td width="8" valign=top rowspan=2><img src="../ill/blank.gif" width="8" height="32" alt="" border="0"></td>
		<td colspan=3 valign="top"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td align=right width="8" valign=top rowspan=2><img src="../ill/blank.gif" width="8" height="32" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td class=alt><b><%=fomr_txt_001 %></b></td>
		<td class=alt align=right><b><%=fomr_txt_026 %></b></td>
		<td class=alt align=right><b>%-<%=fomr_txt_027 %></b> (<%=formatnumber(totaltimer, 2)%>)</td>
	</tr>
	<%
	
	call akttyper2009(2)

    dim fomrjobakt, ktypeid, fomridSum, ktypenavn, fomrnavn, fomridSumTot 
    redim fomrjobakt(50), fomrnavn(50), ktypeid(500), ktypenavn(500), fomridSum(500,50),  fomridSumTot(50) '** ktyper , Forretningsomr

    '*** Forretningsomr�der Navne **'
     '&" LEFT JOIN fomr AS f ON (f.id = fr.for_id) "_
     '  f.navn AS forromrnavn 
    strSQLf = "SELECT f.navn AS forromrnavn, f.id FROM fomr AS f WHERE f.id <> 0"
     oRec6.open strSQLf, oConn, 3
    while not oRec6.EOF 

    fomrnavn(oRec6("id")) = oRec6("forromrnavn")

    oRec6.movenext
    wend
    oRec6.close


    jobnrWrt = 0
    isjobnrWrt = "#0#" 

    ja = fordel_jobakt
    for ja = ja to ja 'fordeling mellem job og aktiviteter 


        if cint(ja) = 0 then 
        %><h4><%=fomr_txt_028 %></h4><%
        else
        %><h4><%=fomr_txt_029 %></h4><%
        end if


    '*** Forretningsomr�der Relationer **'
    jobnrSQLkri = replace(jobnrSQLkri, "tjobnr", "j.jobnr")

    strSQLfor = "SELECT fr.for_fomr, fr.for_aktid, fr.for_jobid, for_faktor, j.jobnr, kt.id AS ktypeid, kt.navn AS ktypenavn FROM job AS j"_
    &" LEFT JOIN fomr_rel AS fr ON (fr.for_jobid = j.id) "_
    &" LEFT JOIN kunder k ON (k.kid = j.jobknr)"_
    &" LEFT JOIN kundetyper kt ON (kt.id = k.ktype)"_
    &" WHERE fr.for_jobid <> 0 "& strJobSQLKri &" GROUP BY kt.id, fr.for_fomr, fr.for_jobid, fr.for_aktid"

    'AND ("& jobnrSQLkri &")
    'strSQL = strSQL & ", kt.id AS ktypeid, kt.navn AS ktypenavn"
    '&" LEFT JOIN aktiviteter AS a ON (a.id = fr.for_aktid) "_
    'fr.for_aktid, fr.for_jobid'

    'if lto = "jttek" then
    'Response.Write "<br><br>Forretningsomr�der Relationer<br>"& strSQLfor & "<br><br>"
    'Response.flush
    'end if

    x = 0
    oRec6.open strSQLfor, oConn, 3
    while not oRec6.EOF 
                
                
                if cint(ja) = 1 then
                fomrjobakt(oRec6("for_fomr")) = " taktivitetid = " & oRec6("for_aktid")
                grpBYja = "tjobnr"

                else
                fomrjobakt(oRec6("for_fomr")) = " tjobnr = '"& oRec6("jobnr") & "'"
                grpBYja = "taktivitetid"

                if instr(isjobnrWrt, ",#"& oRec6("jobnr") &"#") = 0 then
                jobnrWrt = 0
                else
                jobnrWrt = 1
                end if

                end if

                        if (ja = 1 AND oRec6("for_aktid") <> 0) OR (ja = 0 AND oRec6("for_jobid") AND cint(jobnrWrt) = 0) then
                
                        if isNULL(oRec6("ktypeid")) = true then
                        ktypeIDthis = 0
                        else
                        ktypeIDthis = oRec6("ktypeid")
                        end if

                        'Response.write "ktypeIDthis: "& ktypeIDthis & "<br>"
                        'Response.flush

                        if isNull(oRec6("ktypenavn")) = true then
                        ktypenavn(0) = ""
                        else
                        ktypenavn(oRec6("ktypeid")) = oRec6("ktypenavn")
                        end if

                        'fomrjobakt(oRec6("for_fomr")) = oRec6("forromrnavn")
                        '"& replace(oRec6("for_faktor"), ",", ".") &"/100), 2)
                        strSQLt = "SELECT ROUND(SUM(t.timer), 2) AS timerfomr FROM timer AS t WHERE " & fomrjobakt(oRec6("for_fomr")) & " AND ("& aty_sql_realhours &") AND ("& timermedarbSQLkri &") " & tdatoSQLkri & " GROUP BY "& grpBYja &"" 

                        'Response.write "<br><br>"& strSQLt
                        'Response.flush
                        oRec5.open strSQLt, oConn, 3
                        if not oRec5.EOF then

               
                        fomridSum(ktypeIDthis, oRec6("for_fomr")) = fomridSum(ktypeIDthis, oRec6("for_fomr")) + (oRec5("timerfomr"))


                        'oRec5.movenext
                        end if
                        oRec5.close


                        isjobnrWrt = isjobnrWrt & ",#"& oRec6("jobnr") &"#" 

                        end if 'ja = 0/1

    x = x + 1
    oRec6.movenext
    wend
    oRec6.close


   


    l = 0
    for k = 0 to 500

    
     if ktypenavn(k) <> "" then
     
     

            'if x = 0 then
            'lastKtypeNavn = oRec("ktypenavn")
            'end if

        %>
        <tr>
		    <td bgcolor="#cccccc" colspan="5"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	    </tr>
        <tr>
            <td bgcolor="#8cAAe6" class=alt colspan=5 style="padding:6px 4px 2px 7px;"><%=fomr_txt_030 %>: <b> <%= ktypenavn(k)%></b></td>
        </tr>

        <%
        end if
      


            for f = 0 to 50

            if fomridSum(k,f) > 0 then
            

            select case right(l, 1)
	        case 0, 2, 4, 6, 8
	        bgthis = "#eff3ff"
	        case else
	        bgthis = "#FFFFFF"
	        end select
            
            %>
	        <tr>
		        <td bgcolor="#cccccc" colspan="5"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	        </tr>
	        <tr bgcolor="<%=bgthis%>">
		        <td style="height:20px;">&nbsp;</td>
		        <td><%=fomrnavn(f)%>:</td>
		        <td align=right>
		
		        <b><%=formatnumber(fomridSum(k,f))%></b>&nbsp;<%=fomr_txt_048 %></td>
		        <td align=right><%=formatpercent((fomridSum(k,f)/totaltimer), 2)%></td>
		        <td>&nbsp;</td>
	         </tr>
	        <%
            fomridSumTot(f) = fomridSumTot(f) + fomridSum(k,f) 
            timerfomr = timerfomr + fomridSum(k,f)
       
            l = l + 1     
            'Response.Write "fomr: "& fomrnavn(f) &":: "& fomridSum(k,f) & "<br>"
            end if

            next

    next


     next 'ja

    %>
      
       <tr>
		    <td bgcolor="#cccccc" colspan="5"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	    </tr>
       <tr>
		    <td colspan="5"><img src="ill/blank.gif" width="1" height="30" border="0" alt=""></td>
	    </tr>
    
        <tr>
            <td bgcolor="#FFFFFF" colspan=5 style="padding:16px 4px 2px 7px;">&nbsp;</td>
        </tr>
        <tr>
            <td bgcolor="#FFFFFF" colspan=5 style="padding:6px 4px 2px 7px;"><b><%=fomr_txt_031 %>:</b></td>
        </tr>

     <%


    '************ totaler ********************'

    l = 0
   
            for f = 0 to 50

            if fomridSumTot(f) > 0 then
            

            select case right(l, 1)
	        case 0, 2, 4, 6, 8
	        bgthis = "#eff3ff"
	        case else
	        bgthis = "#FFFFFF"
	        end select
            
            %>
	        <tr>
		        <td bgcolor="#cccccc" colspan="5"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	        </tr>
	        <tr bgcolor="<%=bgthis%>">
		        <td style="height:20px;">&nbsp;</td>
		        <td><%=fomrnavn(f)%>:</td>
		        <td align=right>
		
		        <b><%=formatnumber(fomridSumTot(f))%></b>&nbsp;<%=fomr_txt_038 %></td>
		        <td align=right><%=formatpercent((fomridSumTot(f)/totaltimer), 2)%></td>
		        <td>&nbsp;</td>
	         </tr>
	        <%

            fomridSumGTot = fomridSumGTot + fomridSumTot(f)
            l = l + 1
           'Response.Write "fomr: "& fomrnavn(f) &":: "& fomridSum(k,f) & "<br>"
            end if

            next




    	if l = 0 then%>
	
	<tr bgcolor="#eff3ff">
		<td height=20>&nbsp;</td>
		<td colspan=3><br><br><%=fomr_txt_032 %> <b><%=fomr_txt_033 %></b> <%=fomr_txt_034 %> <br>
		<%=fomr_txt_035 %><br><br>&nbsp;</td>
		<td>&nbsp;</td>
	 </tr>
	<%end if%>
	
	
	<tr bgcolor="#FFFFFF">
		<td colspan=2 style="padding:6px 4px 2px 7px;"><b><%=fomr_txt_036 %>:</b></td>
		
     
        <td align=right><b><%=fomridSumGTot %></b> <%=fomr_txt_048 %></td>
        <td>&nbsp;</td>
	
		<td align=right valign=top><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	</table>
	</div>
	
    <br />
	<%=fomr_txt_037 %> ~ <b><%=formatnumber(totaltimer - timerfomr, 2) %></b> <%=fomr_txt_048 %><br />
    <%=fomr_txt_038 %>
	
	<br><br>&nbsp;


    <br /><br />
    <b><span id="sp_grundlag" style="color:#5582d2;">+ <%=fomr_txt_039 %></span></b>
    <div style="width:780px; padding:20px; background-color:#ffffff; visibility:hidden; display:none;" id="div_grundlag" >
    
    <%=strJobTxt %>
	
	</div>

    <br><br><br><br><br><br>&nbsp;


     <!--<script>
         document.getElementById("load").style.visibility = "hidden";
         document.getElementById("load").style.display = "none";
			    </script>
         -->







    <%
    ' ****  siden slutter her PERFORMANCE ***' 
    Response.end

    for t = 0 to 1
	'for_faktor
	strSQL = "SELECT aktiviteter.id AS aid, t.timer/100 AS timerfomr, fomr.id, fomr.navn AS fomrnavn"

    if t = 0 then
    strSQL = strSQL & ", kt.id AS ktypeid, kt.navn AS ktypenavn"
    end if

	strSQL = strSQL &" FROM fomr "_
    &" LEFT JOIN fomr_rel ON (for_fomr = fomr.id) "_
    &" LEFT JOIN job AS j ON (j.id = for_jobid) "_
    &" LEFT JOIN aktiviteter ON (aktiviteter.id = for_aktid) "_
	&" LEFT JOIN timer AS t ON (taktivitetid = aktiviteter.id)" 

    if t = 0 then
    strSQL = strSQL &" LEFT JOIN kunder k ON (k.kid = tknr)"_
    &" LEFT JOIN kundetyper kt ON (kt.id = k.ktype)"
    grpBY = "k.ktype, kid, fomr.id"
    ordBY = "k.ktype, fomrnavn"
	else
    grpBY = "fomr.id"
    ordBY = "fomrnavn"
    end if
	
     strSQL = strSQL &" WHERE fomr.navn <> '' AND t.timer IS NOT NULL AND t.timer > 0 AND ("& aty_sql_realhours &") AND ("& timermedarbSQLkri &") AND ("& jobnrSQLkri &") " & tdatoSQLkri &" "& jobstKri &" GROUP BY "& grpBY &" ORDER BY "& ordBY


    'if lto = "jttek" then
    'response.write strSQL
	'Response.flush
    'end if
	
    lastKtype = -1
    lastKtypeNavn = ""
    x = 0
	timerfomr = 0
	oRec.open strSQL, oConn, 3 
	while not oRec.EOF 
	select case right(x, 1)
	case 0, 2, 4, 6, 8
	bgthis = "#eff3ff"
	case else
	bgthis = "#FFFFFF"
	end select

    if t = 0 then

        if (lastKtype <> oRec("ktypeid") OR x = 0) AND oRec("ktypenavn") <> "" then

            'if x = 0 then
            'lastKtypeNavn = oRec("ktypenavn")
            'end if

        %>
        <tr>
		    <td bgcolor="#cccccc" colspan="5"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	    </tr>
        <tr>
            <td bgcolor="#8cAAe6" class=alt colspan=5 style="padding:6px 4px 2px 7px;">Segment: <b><%=oRec("ktypenavn")%></b></td>
        </tr>

        <%
        end if
    
    else

        if x = 0 then

         %>
        <tr>
		    <td bgcolor="#cccccc" colspan="5"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	    </tr>
        <tr>
            <td bgcolor="#FFFFFF" colspan=5 style="padding:16px 4px 2px 7px;">&nbsp;</td>
        </tr>
        <tr>
            <td bgcolor="lightpink" colspan=5 style="padding:6px 4px 2px 7px;"><b><%=fomr_txt_031 %>:</b></td>
        </tr>

        <%

        end if

    end if
	
    if t = 1 OR (oRec("timerfomr") <> 0 AND t = 0) then
	%>
	<tr>
		<td bgcolor="#cccccc" colspan="5"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="<%=bgthis%>">
		<td style="height:20px;">&nbsp;</td>
		<td><%=oRec("fomrnavn")%>:</td>
		<td align=right>
		
		<b><%=formatnumber(oRec("timerfomr"))%></b>&nbsp;<%=fomr_txt_048 %></td>
		<td align=right><%=formatpercent((oRec("timerfomr")/totaltimer), 2)%></td>
		<td>&nbsp;</td>
	 </tr>
	<%
	
	timerfomr = timerfomr + oRec("timerfomr")
    
    if t = 0 then

    if oRec("ktypeid") <> "" then
    lastKtype = oRec("ktypeid")
    else
	lastKtype = -1
    end if

    lastKtypeNavn = oRec("ktypenavn")

    end if

	x = x + 1

    end if
	
	
	oRec.movenext
	wend
	oRec.close 

    next

case else 
%>

<script src="js/fomr.js" type="text/javascript"></script>

<!--Forretningsomr�der liste-->
<div class="container">
    <div class="portlet">
        <h3 class="portlet-title">
            <u><%=fomr_txt_001 %></u>
        </h3>
                
        <form action="fomr.asp?menu=tok&func=opret" method="post">
          <section>
                         <div class="row">
                             <div class="col-lg-10">&nbsp;</div>
                             <div class="col-lg-2">
                          <button class="btn btn-sm btn-success pull-right"><b><%=fomr_txt_040 %> +</b></button><br />&nbsp;
                 <!--<a href="kunder.asp?menu=<%=menu%>&func=opret&ketype=<%=ketype%>&FM_soeg=<%=thiskri%>&medarb=<%=medarb%>" class="alt">Opret ny +</a>-->
                                     </div>
                 </div>
              </section>
         </form>

        <div class="portet-body">
            <table id="example" class="table dataTable table-striped table-bordered table-hover ui-datatable"> 
                <thead>
                    <tr>
                        <th><%=fomr_txt_041 %></th>
                        <th><%=fomr_txt_006 %></th>
                        <th><%=fomr_txt_042 %></th>
                        <%if cint(fomr_account) <> 1 then %>
                        <th><%=fomr_txt_014 %></th>
                        <%end if %>
                        <th><%=fomr_txt_043 %></th>
                        <th><%=fomr_txt_044 %></th>
                        <th style="width: 5%"><%=fomr_txt_045 %></th>
                    </tr>
                </thead>
                <tbody>
                    <%
                        strSQL = "SELECT f.id, f.navn, COUNT(relakt.for_fomr) AS relakt_antal, kp.kontonr AS kkontonr, kp.navn AS kpnavn, jobok, aktok, business_unit, business_area_label FROM fomr AS f"
                        strSQL = strSQL & " LEFT JOIN kontoplan AS kp ON (kp.id = f.konto) "
                        strSQL = strSQL & " LEFT JOIN fomr_rel AS relakt ON (relakt.for_fomr = f.id AND relakt.for_aktid <> 0) GROUP BY f.id "
	                    oRec.open strSQL,oConn, 3
                        while not oRec.EOF 
                    %>
                    <tr>
                        <td><%=oRec("id") %></td>
                        <td><a href="fomr.asp?func=red&id=<%=oRec("id") %>"><%=oRec("navn") %></a>
                            <%if oRec("business_area_label") <> "" then
                            %>

                        (<%=oRec("business_area_label")%>)
                
                        <%
                         end if %>
                        </td>
                        <td><%=oRec("business_unit") %></td>
                        <% if cint(fomr_account) <> 1 then %>
                        <td><%=oRec("kpnavn") &" "& oRec("kkontonr")%></td>
                        <% end if %>
                            
                        <td>
                            <%if cint(oRec("jobok")) = 1 then %>
                            <i>V</i>&nbsp;&nbsp;
                                <%
                                    antalfomrJob = 0
                                    strSQLantaljob = " SELECT COUNT(reljob.for_fomr) AS reljob_antal FROM fomr_rel AS reljob WHERE (for_fomr = "& oRec("id") &" AND for_jobid <> 0) GROUP BY for_jobid "  
                    
                                    oRec2.open strSQLantaljob, oConn, 3
	                                if not oRec2.EOF then

                                    antalfomrJob = oRec2("reljob_antal")

                                    end if
                                    oRec2.close
                                    %>
                
                            (<a href="../timereg/jobs.asp?FM_fomr=X234<%=oRec("id") %>X234" class="vmenu" target="_blank"><%=antalfomrJob%></a>)
            
                            <%else %>

                            <%end if %>
                        </td>
                        <td>
                             <%if cint(oRec("aktok")) = 1 then %>
                            <i>V</i>&nbsp;&nbsp;(<%=oRec("relakt_antal")%>)
                            <%else %>

                            <%end if %>
                        </td>
                        <td><a href="fomr.asp?menu=tok&func=slet&id=<%=oRec("id")%>"><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></td>
                    </tr>
                    <%
					oRec.movenext
					wend
					oRec.close
					%>
                </tbody>
            </table>
        </div>
    </div>
</div>



<%end select %>
</div> <!--Content-->
</div> <!--Wrapper-->

<!--#include file="../inc/regular/footer_inc.asp"-->
