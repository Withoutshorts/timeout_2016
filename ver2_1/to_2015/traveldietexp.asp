
<%response.buffer = true 
Session.LCID = 1030
%>
			        

<!--#include file="../inc/connection/conn_db_inc.asp"-->


<%'** JQUERY START ************************* %>

<%'** JQUERY END ************************* %>
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../timereg/inc/isint_func.asp"-->


<%'** ER SESSION UDLØBET  ****
    if len(session("user")) = 0 then
	
	errortype = 5
	call showError(errortype)
    response.end
    end if %>

<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<%
    
    if level = 1 then
        medarbSQLkri = " mid <> 0 "
        showalle = 1
        approveActive = ""
        approveActiveInt = 1
        afregnetActive = ""
        afregnetActiveInt = 1
    else
        afregnetActive = "DISABLED"
        approveActiveInt = 0
        afregnetActiveInt = 0
        call fTeamleder(session("mid"), 0, 1)
        if strPrgids = "ingen" then
            medarbSQLkri = " mid = "& session("mid")
            showalle = 0
            approveActive = "DISABLED"
        else
            medarbSQLkri = ""
            strSQL = "SELECT medarbejderid FROM progrupperelationer WHERE projektgruppeid IN ("& strPrgids &") GROUP BY medarbejderid"                                       
            oRec.open strSQL, oConn, 3
            a = 0
            while not oRec.EOF
                if a = 0 then
                    medarbSQLkri = medarbSQLkri & "(mid = "& oRec("medarbejderid")
                    allmedids = oRec("medarbejderid")
                else
                    medarbSQLkri = medarbSQLkri & " OR mid = "& oRec("medarbejderid")
                    allmedids = allmedids & "," & oRec("medarbejderid")
                end if

            a = a + 1
            oRec.movenext
            wend
            oRec.close
                                        
            medarbSQLkri = medarbSQLkri & ")"
            showalle = 1
            approveActive = ""
            approveActiveInt = 1
        end if
    end if 'level = 1

%>



<%
    public strJobOptions
    sub jobliste


    select case lto
    case "lm" '**LAAAANGSOM mange mange job
    strSQLjobkri = " AND j.id = 3"
    case else
    strSQLjobkri = ""
    end select


     '*** Job liste
         strJobOptions = "<option value='0'>"&diet_txt_001&"</option>"
         strJobOptions = strJobOptions & "<option value='0'>"&diet_txt_002&"</option>"
         strSQLjobliste = "SELECT jobnavn, jobnr, j.id, kkundenavn, kid FROM timereg_usejob AS tu "_
         &" LEFT JOIN job AS j ON (j.id = tu.jobid) "_
         &" LEFT JOIN kunder ON (kid = j.jobknr) "_
         &" WHERE tu.medarb = "& usemrn &" AND tu.aktid = 0 AND jobstatus = 1 AND risiko >= 0 "& strSQLjobkri &" ORDER BY kkundenavn, jobnavn"
       
        '** Behøver ikke være på aktiv joblisten
        'AND tu.forvalgt = 1
        
         'response.Write strSQLjobliste
         'response.flush

         oRec.open strSQLjobliste, oConn, 3
         j = 0
         while not oRec.EOF

         if lastKid <> oRec("kid") then

            if j > 0 then
            strJobOptions = strJobOptions & "<option value='0' DISABLED></option>"
            end if

         strJobOptions = strJobOptions & "<option value='0' DISABLED>"& left(oRec("kkundenavn"), 25) &"</option>" 
         end if

         strJobOptions = strJobOptions & "<option value="& oRec("id") &">"& left(oRec("jobnavn"), 25) &" ("& oRec("jobnr") &")</option>" 

         lastKid = oRec("kid")
         j = j + 1
         oRec.movenext
	     wend
         oRec.close


    end sub


   	media = request("media")

	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if  


    if len(trim(request("aar"))) <> 0 then
    aar = request("aar")
    else
    aar = "1-1-"& year(now)
    end if
    

    if len(trim(request("medarb"))) <> 0 then
    usemrn = request("medarb")
    else
    usemrn = session("mid")
    end if
    
     if media <> "print" then    
        select case func
        case "addmoorejobs", "addmoorejobsOpd", "slet", "sletok"

        case else
        call menu_2014
        end select
    end if
     
%>    

    
        <div id="wrapper">
            <div class="content">


    <%
    select case func
    case "addmoorejobsOpd"

        if len(trim(request("trvlid"))) <> 0 then
        trvlid = request("trvlid")
        else
        trvlid = 0
        end if

        
        diet_jobid_0 = request("FM_diet_jobid_0")
        diet_jobid_1 = request("FM_diet_jobid_1")
        diet_jobid_2 = request("FM_diet_jobid_2")
        diet_jobid_3 = request("FM_diet_jobid_3")
        diet_jobid_4 = request("FM_diet_jobid_4")
        
        diet_jobid_pro_0 = request("FM_diet_jobid_pro_0")
        diet_jobid_pro_1 = request("FM_diet_jobid_pro_1")
        diet_jobid_pro_2 = request("FM_diet_jobid_pro_2")
        diet_jobid_pro_3 = request("FM_diet_jobid_pro_3")
        diet_jobid_pro_4 = request("FM_diet_jobid_pro_4") 

        '*** RENSER UD **
        strSQLdel = "DELETE FROM traveldietexp_jobrel WHERE tj_trvlid = "& trvlid
        oConn.execute(strSQLdel)

        for r = 0 TO 4

        select case r
        case 0
        thisJobid = diet_jobid_0
        thisPercent = diet_jobid_pro_0
        case 1
        thisJobid = diet_jobid_1
        thisPercent = diet_jobid_pro_1
        case 2
        thisJobid = diet_jobid_2
        thisPercent = diet_jobid_pro_2
        case 3
        thisJobid = diet_jobid_3
        thisPercent = diet_jobid_pro_3
        case 4
        thisJobid = diet_jobid_4
        thisPercent = diet_jobid_pro_4
        end select

        if len(trim(thisJobid)) <> 0 then
        thisJobid = thisJobid
        else
        thisJobid = 0
        end if

        call erDetInt(thisPercent)


        'response.Write "<br>thisJobid: " & thisJobid
        'response.flush

            if cint(thisJobid) <> 0 AND len(trim(thisPercent)) <> 0 AND isInt = 0 then

            thisPercent = replace(thisPercent, ",", ".")
            strSQLins = "INSERT INTO traveldietexp_jobrel (tj_trvlid, tj_jobid, tj_percent) VALUES ("& trvlid &", "& thisJobid &", "& thisPercent &")"
            oConn.execute(strSQLins)

            end if

        next

        Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
		Response.Write("<script language=""JavaScript"">window.close();</script>")

        'response.Redirect "traveldietexp.asp?func=addmoorejobs&medarb="& usemrn &"&trvlid="&trvlid &"&eropda=1"


	case "addmoorejobs"

        if len(trim(request("trvlid"))) <> 0 then
        trvlid = request("trvlid")
        else
        trvlid = 0
        end if

        if len(trim(request("eropda"))) <> 0 then
        eropda = request("eropda")
        else
        eropda = 0
        end if
        
        

        
        %>
          <div class="container">
        <div class="porlet">
            
            <h3 class="portlet-title">
               <%=diet_txt_003 %>
            </h3>
            
            <div class="portlet-body">
                <form method="post" action="traveldietexp.asp?func=addmoorejobsOpd&trvlid=<%=trvlid%>&medarb=<%=usemrn %>">

                  

                <table class="table-stribed table" width="100%">
                    
                    <thead>
                        <tr>
                           
                            <th><%=diet_txt_004 %></th>
                            <th><%=diet_txt_005 %></th>
                        </tr>

                    </thead>
                    <tbody>
              
               
                <%
                call jobliste

                    f = 0
                    strSQLjobrel = "SELECT tj_trvlid, tj_jobid, tj_percent, jobnavn, jobnr, j.id AS jid FROM traveldietexp_jobrel "_
                    &" LEFT JOIN job AS j ON (j.id = tj_jobid) WHERE tj_trvlid = "& trvlid
                    oRec2.open strSQLjobrel, oConn, 3
                    while not oRec2.EOF
        
                       tj_percent = formatnumber(oRec2("tj_percent"), 0)

                      %>
                        <tr>
                           
                            <td><select name="FM_diet_jobid_<%=f %>" class="form-control input-small" style="width:400px;">
                                <option value="<%=oRec2("jid") %>" SELECTED><%=oRec2("jobnavn") & " ("& oRec2("jobnr") & ")"%></option>
                                <%=strJobOptions %>
                                </select></td>
                            <td><input name="FM_diet_jobid_pro_<%=f %>" type="text" class="form-control input-small" value="<%=tj_percent %>" style="width:60px;" /></td>
                        </tr>
                    <%
                    f = f + 1
                    oRec2.movenext
                    wend
                    oRec2.close 

                    if f > 0 then
                    f = f
                    else 
                    f = 0
                    end if


                for j = f To 4

                    if j = 0 then
                    jprocval = 100
                    else 
                    jprocval = ""
                    end if
                    %>
                        <tr>
                           
                            <td><select name="FM_diet_jobid_<%=j %>" class="form-control input-small" style="width:400px;"><%=strJobOptions %></select></td>
                            <td><input name="FM_diet_jobid_pro_<%=j %>" type="text" class="form-control input-small" value="<%=jprocval %>" style="width:60px;" /></td>
                        </tr>
                    <%
                 next     
                %>
                        </tbody>
                  </table>


                      <div class="row">
                             <div class="col-lg-10">&nbsp;

                                 
                                  <%if cint(eropda) = 1 then %>
                                <p style="color:green;"><%=diet_txt_006 %>!</p>
                                <%end if %>

                             </div>
                             <div class="col-lg-1">
                            <input type="submit" value="<%=diet_txt_007 %> >>" class="btn btn-sm btn-success pull-right"><br />&nbsp;
                            </div>
                        </div>

                </form>
                </div>
            </div>
        </div>
        	<%
    case "slet"
    %>
    <!--slet sidens indhold-->
 
    <div class="container" style="width:600px;">
        <div class="porlet">
            
            <h3 class="portlet-title">
               <%=diet_txt_008 %>
            </h3>
            
            <div class="portlet-body" style="width:400px;">
                <div align="center"><%=diet_txt_009 %> <b><%=diet_txt_010 %></b> <%=diet_txt_011 %>
                </div><br />
                <div align="center"><b><a href="traveldietexp.asp?func=sletok&id=<%=id%>&aar=<%=aar%>&medarb=<%=usemrn%>"><%=diet_txt_012 %></a></b>&nbsp&nbsp&nbsp&nbsp<b><a href="javascript:history.back()"><%=diet_txt_013 %></a></b>
                </div>
                <br /><br />
                </div>
            </div>
        </div>
    
    <%
    case "sletok"
        oConn.execute("DELETE FROM traveldietexp WHERE diet_id = "& id &"")
	    Response.redirect "traveldietexp.asp?aar="&aar&"&medarb="&usemrn
    %>    

    
    <%

        case "dbopr", "dbred"


        case "opdaterlist"
	    '*** Her indsættes en ny type i db ****
		
		
		
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function
		
       
        'Response.write "HER: "& request("FM_diet_namedest")
      
            
        strIds = split(request("FM_diet_id"), ", ")
        strDiet_medids = split(request("FM_diet_medid"), ", ")
		strNavn = split(request("FM_diet_namedest"), ", ##, ")

        strStdato = split(request("FM_diet_stdato"),", ##, ")
        strSldato = split(request("FM_diet_sldato"),", ##, ")   

        strJobids = split(request("FM_diet_jobid"), ", ")
        strkontoids = split(request("FM_diet_konto"), ", ")

        strMorgenAntal = split(request("FM_diet_morgen"),", ##, ") 
        strMiddagAntal = split(request("FM_diet_middag"),", ##, ")  
        strAftenAntal = split(request("FM_diet_aften"),", ##, ")        
		
        strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)

        diet_daypriceVal = request("FM_diet_dayprice")
        if len(trim(diet_daypriceVal)) <> 0 then
        diet_daypriceVal = diet_daypriceVal
        else
        diet_daypriceVal = 0
        end if

        if len(trim(request("FM_diet_dayprice"))) <> 0 then
        diet_dayprice = replace(request("FM_diet_dayprice"), ",", ".")
        else
        diet_dayprice = 0
        end if


        diet_dayprice_halfVal = request("FM_diet_dayprice_delvis")
        if len(trim(diet_dayprice_halfVal)) <> 0 then
        diet_dayprice_halfVal = diet_dayprice_halfVal
        else
        diet_dayprice_halfVal = 0
        end if

        if len(trim(request("FM_diet_dayprice_delvis"))) <> 0 then
        diet_dayprice_half = replace(request("FM_diet_dayprice_delvis"), ",", ".")
        else
        diet_dayprice_half = 0
        end if
        
        diet_morgenamountVal = request("FM_diet_morgenamount")
        if len(trim(diet_morgenamountVal)) <> 0 then
        diet_morgenamountVal = diet_morgenamountVal
        else
        diet_morgenamountVal = 0
        end if

        diet_middagamountVal = request("FM_diet_middagamount")
        if len(trim(diet_middagamountVal)) <> 0 then
        diet_middagamountVal = diet_middagamountVal
        else
        diet_middagamountVal = 0
        end if

        diet_aftenamountVal = request("FM_diet_aftenamount")
        if len(trim(diet_aftenamountVal)) <> 0 then
        diet_aftenamountVal = diet_aftenamountVal
        else
        diet_aftenamountVal = 0
        end if

        
        if len(trim(request("FM_diet_morgenamount"))) <> 0 then
        diet_morgenamount = replace(request("FM_diet_morgenamount"), ",", ".")
        else
        diet_morgenamount = 0
        end if

        
        if len(trim(request("FM_diet_middagamount"))) <> 0 then
        diet_middagamount = replace(request("FM_diet_middagamount"), ",", ".")
        else
        diet_middagamount = 0
        end if

        
        if len(trim(request("FM_diet_aftenamount"))) <> 0 then
        diet_aftenamount = replace(request("FM_diet_aftenamount"), ",", ".")
        else
        diet_aftenamount = 0
        end if

       
        'diet_godkend = split(request("FM_diet_godkend"),"##, ") 

        'response.write "<br>"& request("FM_diet_bilag") & "<br>"

        strBilag = split(request("FM_diet_bilag"),"##, ") 


        call akttyper2009(2)
        call traveldietexp_fn()


		for i = 0 TO UBOUND(strIds)

        strNavn(i) = SQLBless(strNavn(i))
        strNavn(i) = replace(strNavn(i), "#", "")
        strNavn(i) = replace(strNavn(i), ",", "")
        strNavn(i) = trim(strNavn(i))

        
        strStdato(i) = replace(strStdato(i), "#", "")
        strStdato(i) = replace(strStdato(i), ",", "")

        strSldato(i) = replace(strSldato(i), "#", "")
        strSldato(i) = replace(strSldato(i), ",", "")
        
        'response.Write "<br> strSldatostrSldatostrSldatostrSldatostrSldato - " & strSldato(i)
        if (isDate(strStdato(i)) = true OR isDate(strSldato(i)) = true) AND len(trim(strNavn(i))) = 0 then
            useleftdiv = "to_2015"
			errortype = 206
			call showError(errortype)
		    Response.End
        end if
       

        if isDate(strStdato(i)) = true then
            strStdato(i) = year(strStdato(i)) &"-"& month(strStdato(i)) &"-"& day(strStdato(i)) '&" "& formatdatetime(strStdato(i), 3)
        else
            strStdato(i) = year(now) &"-01-01"
            if len(trim(strNavn(i))) <> 0 then 
                useleftdiv = "to_2015"
			    errortype = 175
			    call showError(errortype)
		        Response.End
            end if
        end if

        if cint(request("brug_tidspunkt")) = 1 then
            if (len(trim(request("FM_diet_sttime_"& strIds(i))))) <> 0 then
                strSttime = request("FM_diet_sttime_" & strIds(i))
            else
                if len(trim(strNavn(i))) <> 0 then 
                    useleftdiv = "to_2015"
			        errortype = 205
			        call showError(errortype)
		            Response.End
                end if
            end if
        else
            strSttime = "00:00:00"
        end if


        strStdato(i) = strStdato(i) & " " & strSttime

        'response.Write "<br> herherherherherherherherherherherherher -  strStdato(i) " & strStdato(i) & "<br>"

         
        if isDate(strSldato(i)) = true then
            strSldato(i) = year(strSldato(i)) &"-"& month(strSldato(i)) &"-"& day(strSldato(i)) '&" "& formatdatetime(strSldato(i), 3)
        else
            strSldato(i) = year(now) &"-01-01"
            if len(trim(strNavn(i))) <> 0 then 
                useleftdiv = "to_2015"
			    errortype = 175
			    call showError(errortype)
		        Response.End
            end if
        end if

        if cint(request("brug_tidspunkt")) = 1 then
            if (len(trim(request("FM_diet_sltime_"& strIds(i))))) <> 0 then
                strSltime = request("FM_diet_sltime_" & strIds(i))
            else
                if len(trim(strNavn(i))) <> 0 then 
                    useleftdiv = "to_2015"
			        errortype = 205
			        call showError(errortype)
		            Response.End
                end if
            end if
        else
            strSltime = "00:00:00"
        end if

        

        strSldato(i) = strSldato(i) & " " & strSltime

        'response.Write "<br> herherherherherherherherherherherherher -  strStdato(i) " & strStdato(i) & " slutdato " & strSldato(i)




        if lto = "outz" OR lto = "plan" then
        strDelVis = request("FM_delvis_"&strIds(i))
        else
        strDelVis = 0
        end if


        'response.Write "<br> herherherherherherherherherherherherher strDelVis " & strDelVis
        'response.End

        '****************************************************************************
        '*** Tjekker om der er indtastet med end på de dage der oprettes diæter på
        '****************************************************************************

        antalDagArr = dateDiff("d", strStdato(i), strSldato(i), 2,2)
        timerForbrugprDagOverskreddet = 0

        for d = 0 TO antalDagArr

            if d = 0 then
            tjkThisDatp = year(strStdato(i)) & "-" & month(strStdato(i)) & "-" & day(strStdato(i))
            else
            tjkThisDatp = dateAdd("d", 1, tjkThisDatp)
            tjkThisDatp = year(tjkThisDatp) & "-" & month(tjkThisDatp) & "-" & day(tjkThisDatp)
            end if 

            timerBrugt = 0
            strSQLtimreal = "SELECT COALESCE(SUM(timer), 0) AS timerbrugt FROM timer WHERE tmnr = "& strDiet_medids(i) &" AND tdato = '"& tjkThisDatp &"' AND ("& aty_sql_realhours &") GROUP BY tdato, tmnr"
        
        
            'response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;antalDagArr: "& antalDagArr &" strSQLtimreal: " & strSQLtimreal & "<br>"
                      
            oRec5.open strSQLtimreal, oConn, 3
            while not oRec5.EOF
               
            timerBrugt = oRec5("timerbrugt")
        
            if cdbl(timerBrugt) > cdbl(traveldietexp_maxhours) then
            timerForbrugprDagOverskreddet = 1
            end if         
        
            oRec5.movenext
            wend
            oRec5.close

        next
        
        select case lto 
        case "plan"
        timerForbrugprDagOverskreddet = 0
        case else
        timerForbrugprDagOverskreddet = timerForbrugprDagOverskreddet
        end select

        
        if cint(timerForbrugprDagOverskreddet) = 1 then

                    useleftdiv = "to_2015"
					errortype = 178
					call showError(errortype)
		            Response.End

        end if
        
        '*****************************************************************************

        if len(trim(request("FM_diet_godkend_" & strIds(i)))) <> 0 then
            diet_godkend = request("FM_diet_godkend_" & strIds(i))
        else
            diet_godkend = 0
        end if

        if len(trim(request("FM_afreget_" & strIds(i)))) <> 0 then
            diet_afregnet = request("FM_afreget_" & strIds(i))            
        else
            diet_afregnet = 0
        end if

        if diet_afregnet = 1 then
            diet_godkend = 1 'Hvis afregnet bliver diæet automatisk godkendt
        end if
        
        response.Write "<br> ---------------------------------------- " & diet_afregnet

        if cint(request("FM_godkend_last_" & strIds(i))) <> cint(diet_godkend) then
            diet_godkenddate = year(now) &"-"& month(now) &"-"& day(now)
            strgodkendtaf = session("user")
        else
            if len(trim(request("FM_godkend_date_" & strIds(i)))) <> 0 then
                diet_godkenddate = request("FM_godkend_date_" & strIds(i))
            else
                diet_godkenddate = "2002-01-01"
            end if

            if len(trim(request("FM_godkendtaf_last_" & strIds(i)))) <> 0 then
                strgodkendtaf = request("FM_godkendtaf_last_" & strIds(i))
            else
                strgodkendtaf = ""
            end if
        end if

        diet_godkenddate = year(diet_godkenddate) &"-"& month(diet_godkenddate) &"-"& day(diet_godkenddate)

        if cint(request("FM_last_afregnet_" & strIds(i))) <> cint(diet_afregnet) then
            dietafrengetdate = year(now) &"-"& month(now) &"-"& day(now)
        else
            if len(trim(request("FM_afregnet_dato_" & strIds(i)))) <> 0 then
                dietafrengetdate = request("FM_afregnet_dato_" & strIds(i))
            else
                dietafrengetdate = "2002-01-01"
            end if
        end if

        dietafrengetdate = year(dietafrengetdate) &"-"& month(dietafrengetdate) &"-"& day(dietafrengetdate)


        strBilag(i) = replace(strBilag(i), "#", "")
        strBilag(i) = replace(strBilag(i), ",", "")
      
        if len(trim(strBilag(i))) <> 0 then
        strBilag(i) = strBilag(i) 
        else
        strBilag(i) = 0
        end if


        diet_dageIalt = dateDiff("h", strStdato(i), strSldato(i), 2,2) / 24
        if diet_dageIalt <> 0 then
            if instr(diet_dageIalt, ",") <> 0 then
            komma = instr(diet_dageIalt, ",")
            diet_dageIalt = left(diet_dageIalt, komma+2)
            else
            diet_dageIalt = diet_dageIalt
            end if

        else
        diet_dageIalt = 0
        end if

        'Lægger en dag til for tia, fordi de altid får penge uanset om de sover der eller ej
        if lto = "tia" then
            diet_dageIalt = diet_dageIalt + 1
        end if
        


        if len(trim(strJobids(i))) <> 0 then
        strJobids(i) = strJobids(i) 
        else
        strJobids(i) = 0
        end if

        
        if len(trim(strkontoids(i))) <> 0 then
        strkontoids(i) = strkontoids(i) 
        else
        strkontoids(i) = 0
        end if

                
       
        strMorgenAntal(i) = replace(strMorgenAntal(i), "#", "")
        strMorgenAntal(i) = replace(strMorgenAntal(i), ",", "")
        strMorgenAntal(i) = replace(strMorgenAntal(i), ".", "")
        strMorgenAntal(i) = trim(strMorgenAntal(i))

        if len(trim(strMorgenAntal(i))) <> 0 then
        strMorgenAntal(i) = strMorgenAntal(i)
        else
        strMorgenAntal(i) = 0
        end if

        strMiddagAntal(i) = replace(strMiddagAntal(i), "#", "")
        strMiddagAntal(i) = replace(strMiddagAntal(i), ",", "")
        strMiddagAntal(i) = replace(strMiddagAntal(i), ".", "")
        strMiddagAntal(i) = trim(strMiddagAntal(i))

        if len(trim(strMiddagAntal(i))) <> 0 then
        strMiddagAntal(i) = strMiddagAntal(i)
        else
        strMiddagAntal(i) = 0
        end if

        strAftenAntal(i) = replace(strAftenAntal(i), "#", "")
        strAftenAntal(i) = replace(strAftenAntal(i), ",", "")
        strAftenAntal(i) = replace(strAftenAntal(i), ".", "")
        strAftenAntal(i) = trim(strAftenAntal(i))
        
        if len(trim(strAftenAntal(i))) <> 0 then
        strAftenAntal(i) = strAftenAntal(i)
        else
        strAftenAntal(i) = 0
        end if

        diet_total = (strMorgenAntal(i)/1+strMiddagAntal(i)/1+strAftenAntal(i)/1)
        if len(trim(diet_total)) <> 0 then
        diet_total = diet_total
        else
        diet_total = 0
        end if

        'response.Write " diet_daypriceVal " & diet_daypriceVal

        if lto = "plan" or lto = "outz" then
            if cint(strDelvis) = 1 then
            makskost = (diet_dayprice_halfVal/1)*diet_dageIalt
            else
            makskost = (diet_daypriceVal/1)*diet_dageIalt
            end if
        else
        makskost = (diet_daypriceVal/1)*diet_dageIalt
        end if

        
        totalefterReduktion = makskost/1 - ((strMorgenAntal(i)*diet_morgenamountVal/1) + (strMiddagAntal(i)*diet_middagamountVal/1) + (strAftenAntal(i)*diet_aftenamountVal/1))
        
        makskost = replace(makskost, ",", ".")
        diet_dageIalt = replace(diet_dageIalt, ",", ".")
        totalefterReduktion = replace(totalefterReduktion, ",", ".")
        
        if len(trim(strNavn(i))) <> 0 then 

		    if strIds(i) = 0 then
		    strSQlins = ("INSERT INTO traveldietexp (diet_namedest, diet_stdato, diet_sldato, diet_jobid, diet_konto, diet_morgen, diet_middag, diet_aften, diet_total, diet_dayprice, diet_dayprice_half, "_
            &" diet_morgenamount, diet_middagamount, diet_aftenamount, diet_mid, diet_bilag, diet_maksamount, diet_rest, diet_traveldays, diet_delvis, diet_approved, diet_approveddate, diet_approvedby, diet_settled, diet_settleddate) "_
            &" VALUES ('"& strNavn(i) &"', '"& strStdato(i) &"', '"& strSldato(i) &"', "& strJobids(i) &", "& strkontoids(i) &", "_
            &" "& strMorgenAntal(i) &", "& strMiddagAntal(i) &", "& strAftenAntal(i) &", "& diet_total &", "& diet_dayprice &", "& diet_dayprice_half &", "_
            &" "& diet_morgenamount &","& diet_middagamount &","& diet_aftenamount &", "& strDiet_medids(i) &", "& strBilag(i) &", "& makskost &", "& totalefterReduktion &", "& diet_dageIalt &", "& strDelVis &", "& diet_godkend &", '"& diet_godkenddate &"', '"& strgodkendtaf &"', "& diet_afregnet &", '"& dietafrengetdate &"')")

            'response.write strSQlins
            'response.flush

            oConn.execute(strSQlins) 

                
           

		    else
		    strSQlupd = ("UPDATE traveldietexp SET "_
            &" diet_namedest ='"& strNavn(i) &"', diet_stdato = '"& strStdato(i) &"', diet_sldato = '"& strSldato(i) &"', "_
            &" diet_jobid = "& strJobids(i) &", diet_konto = "& strkontoids(i) &", "_
            &" diet_morgen = "& strMorgenAntal(i) &", diet_middag = "& strMiddagAntal(i) &", diet_aften = "& strAftenAntal(i) &", diet_total = "& diet_total &", diet_dayprice = "& diet_dayprice &", diet_dayprice_half = "& diet_dayprice_half &", "_
            &" diet_morgenamount ="& diet_morgenamount &", diet_middagamount = "& diet_middagamount &", "_
            &" diet_aftenamount =  "& diet_aftenamount &", diet_mid = "& strDiet_medids(i) &", diet_bilag = "& strBilag(i) &", diet_maksamount = "& makskost &", diet_rest = "& totalefterReduktion &", diet_traveldays = "& diet_dageIalt &", diet_delvis = "& strDelVis &", "_
            &" diet_approved = "& diet_godkend &", diet_approveddate = '"& diet_godkenddate &"', diet_approvedby = '"& strgodkendtaf &"', diet_settled = "& diet_afregnet &", diet_settleddate = '"& dietafrengetdate &"'" _ 
            &" WHERE diet_id = "& strIds(i) &"")

            'response.write strSQlupd
            'response.flush

            oConn.execute(strSQlupd)

		    end if

        else

             oConn.execute("DELETE FROM traveldietexp WHERE diet_id = "& strIds(i) &"")

        end if


            'response.Write "<br> --------------------------------  godkend " & diet_godkend & " date " & diet_godkenddate & " afrenget " & diet_afregnet & " set. date " & dietafrengetdate & "GODKEDNT AF " & strgodkendtaf

        next
		
            

        'response.end
		Response.redirect "traveldietexp.asp?medarb="&usemrn&"&aar="&aar
		


    case "opret", "red"
	


case else    


          
            select case lto
            case "plan", "intranet - local"
                brug_godkend = 0
                vis_reduktion = 1
                hide_klokkeslet = 0
                brug_tidspunkt = 1
            case "tia"
                brug_godkend = 0
                vis_reduktion = 0
                hide_klokkeslet = 0
                brug_tidspunkt = 0
            case "lm"
                brug_godkend = 1
                vis_reduktion = 1 
                hide_klokkeslet = 0
                brug_tidspunkt = 1
            case else    
                brug_godkend = 0
                vis_reduktion = 1 
                hide_klokkeslet = 0
                brug_tidspunkt = 1
            end select  

            if cint(vis_reduktion) = 0 then
            jobsize = "25%"
            else
            jobsize = "7%"
            end if


%>

<script src="js/traveldietexp_jav202001.js" type="text/javascript"></script>

<div class="container" style="width:1500px;">
    <div class="porlet">
        <h3 class="portlet-title"><u><%=diet_txt_014 %></u></h3>

        <%if media <> "export" then %>

                
                       <form method="post" action="traveldietexp.asp?issubmitted=1">
                            
                    
                           
                        <!-- NAVN / SORTERING / ID -->
                        <section>
                            <div class="well">
                                 <div class="row">
                                         <div class="col-lg-4 pad-t5"><%=diet_txt_015 %>:<br /> 
                                            <%                                              
                                            strSQlmedarblist = "SELECT mid, mnavn, init FROM medarbejdere WHERE mansat = 1 AND "& medarbSQLkri &" ORDER BY mnavn"
                                            %>
                                             
                                             <select name="medarb" class="form-control input-small" onchange="submit();">
                                             

                                                 <%oRec.open strSQlmedarblist, oConn, 3
                                                 while not oRec.EOF 
                                                     
                                                     if cint(oRec("mid")) = cint(usemrn) then
                                                     medidSel = "SELECTED" 
                                                     else
                                                     medidSel = ""
                                                     end if

                                                     
                                                     
                                                     %>
                                                    <option value="<%=oRec("mid") %>" <%=medidSel %>><%=oRec("mnavn") &" ["& oRec("init") &"]" %></option>
                                                     <%
                                                 oRec.movenext
                                                 wend 
                                                 oRec.close
                                                         
                                              if showalle = 1 then
                                                         
                                                         if cint(usemrn) = 0 then
                                                         visAlleSel = "SELECTED"
                                                         else
                                                         visAlleSel = ""
                                                         end if
                                                         %>

                                             <option value="0" <%=visAlleSel%>><%=diet_txt_016 %></option>
                                            <%end if %>
                                         </select>    
                                         </div>
                                        
                                       
                                                 
                                                <div class="col-lg-2 pad-t5"><%=diet_txt_044 %>:
                                                    <div class='input-group date' id='datepicker_stdato'>
                                                    <input type="text" class="form-control input-small" autocomplete="off" name="aar" value="<%=aar %>" placeholder="dd-mm-yyyy" />
                                                          <span class="input-group-addon input-small">
                                                                    <span class="fa fa-calendar">
                                                                    </span>
                                                                </span>
                                                           </div>
                                                </div>

                                             
                                            
                                              <div class="col-lg-1 pad-t20">
                                                    <button type="submit" class="btn btn-secondary btn-sm pull-right"><b><%=diet_txt_017 %> >></b></button>
                                             </div>
                                
                                      
                                 </div>

                            </div>

                        </section>
                         </form>


     
        
        <form action="traveldietexp.asp?func=opdaterlist" method="post">
            <input type="hidden" name="medarb" value="<%=usemrn%>"/> 
            <input type="hidden" name="aar" value="<%=aar%>"/>
            <input type="hidden" name="brug_tidspunkt" value="<%=brug_tidspunkt %>" />
            <section>
                         <div class="row">
                             <div class="col-lg-10">&nbsp;</div>
                             <div class="col-lg-2">
                            <input type="submit" value="<%=diet_txt_007 %> >>" class="btn btn-sm btn-success pull-right"><br />&nbsp;
                            </div>
                        </div>
                </section>
        
            
            

        <div class="porlet-body">

            <%
            strSQL = "SELECT tdf_diet_name FROM travel_diet_tariff WHERE tdf_diet_current = 1"
            oRec2.open strSQL, oConn, 3
            if not oRec2.EOF then
                response.Write "<span style='font-size:9px;'><i>"& diet_txt_058 &" - " & oRec2("tdf_diet_name") & "</i></span>"
            end if
            oRec2.close
            %>
          
           <table id="kundetyper" class="table dataTable table-striped table-bordered table-hover ui-datatable">                 
               <thead>

                   <%if lto <> "tia" then %>
                    <tr style="background-color:#FFFFFF;">
                    <%if lto = "plan" or lto = "outz" then %>
                      <th style="border:0;">&nbsp;</th>
                      <th style="border:0;">&nbsp;</th>
                    <%end if %>
                        <th style="border:0;">&nbsp;</th>
                       <th style="border:0;">&nbsp;</th>
                    <%if cint(brug_tidspunkt) = 1 then %>
                      <th style="border:0;">&nbsp;</th>
                       <th style="border:0;">&nbsp;</th>
                    <%end if %>
                       <th style="border:0;">&nbsp;</th>
                        <th style="border:0;"></th>
                        <th style="border:0;">&nbsp;</th>
                       <!-- <th style="border:0;">&nbsp;</th>-->
                        <%if lto <> "plan" AND lto <> "outz" then %>
                        <th style="border:0;">&nbsp;</th>
                        <th style="border:0;">&nbsp;</th>
                        <%else %>
                        <th colspan="2" style="border:0; background-color:#D6DFf5;">&nbsp;</th>
                        <%end if %>
                      
                        <%if cint(vis_reduktion) = 1 then %>
                            <th colspan="4" style="border-bottom:0; background-color:#D6DFf5;"><%=diet_txt_018 %></th>
                        <%end if %>


                        <%if lto <> "plan" AND lto <> "outz" then %>
                        <th style="border:0;">&nbsp;</th>
                        <th style="border:0;">&nbsp;</th>
                        <%else %>
                        <th style="border:0; background-color:#ff9797;">&nbsp;</th>
                        <th style="border:0; background-color:#ff9797;">&nbsp;</th>
                        <%end if %>

                        <%if lto <> "tia" then %>
                        <th style="border:0;">&nbsp;</th>
                        <%end if %>

                        <%if cint(vis_reduktion) = 1 then %>
                            <%if lto = "dencker" OR lto = "outz" then %>
                                <th style="border:0;">&nbsp;</th>
                                <th style="border:0;">&nbsp;</th>                           
                            <%end if %>
                       <%end if %>
                   </tr>
                   <%end if %>

                   <tr>
                       <%if lto = "plan" or lto = "outz" then %>
                       <th><%=diet_txt_041 %></th>
                       <th><%=diet_txt_042 %></th>
                       <%end if %>
                       <th style="width: 5%"><%=diet_txt_019 %></th>
                       <th style="width: 19%"><%=diet_txt_021 %> <span style="color:red;">*</span><br /><span style="font-size:9px;"><%=diet_txt_020 %></span></th>

                       <%if cint(brug_tidspunkt) = 1 then %>
                       <th style="width:5px;"><%=diet_txt_021 %> <span style="color:red;">*</span><br /><span style="font-size:9px;"><%=diet_txt_022 %></span></th>
                       <%end if %>

                       <th style="width: 19%"><%=diet_txt_023 %> <span style="color:red;">*</span><br /><span style="font-size:9px;"><%=diet_txt_020 %></span></th>

                       <%if cint(brug_tidspunkt) = 1 then %>
                       <th style="width:5px;"><%=diet_txt_023 %> <span style="color:red;">*</span><br /><span style="font-size:9px;"><%=diet_txt_022 %></span></th>
                       <%end if %>

                       <th style="width: 15%"><%=diet_txt_024 %> <span style="color:red;">*</span></th>
                       <th style="width:<%=jobsize%>;"><%=diet_txt_025 %><br /><span style="font-size:9px;"><%=diet_txt_026 %></span></th>
                       <!--
                       <th style="width: 10%">Kontonr</th>
                       -->

                       <%if lto = "plan" or lto = "outz" then %>
                       <th style="width: 5%; background-color:#D6DFf5;"><%=diet_txt_027 %><br /><span style="font-size:9px;"><%=diet_txt_028 %></span></th>
                       <th style="background-color:#D6DFf5;"><%=diet_txt_029 %></th>
                       <%else %>
                       <th style="width: 7%;"><%=diet_txt_027 %><br /><span style="font-size:9px;"><%=diet_txt_028 %></span></th>
                           <%if lto <> "tia" then %>
                            <th><%=diet_txt_029 %></th>
                           <%end if %>
                       <%end if %>
                      
                        <%if cint(vis_reduktion) = 1 then %>
                       <th style="width: 8%; background-color:#D6DFf5;"><%=diet_txt_030 %><br /><span style="font-size:9px;"><%=diet_txt_033 %></span></th>
                       <th style="width: 8%; background-color:#D6DFf5;"><%=diet_txt_031 %><br /><span style="font-size:9px;"><%=diet_txt_033 %></span></th>
                       <th style="width: 8%; background-color:#D6DFf5;"><%=diet_txt_032 %><br /><span style="font-size:9px;"><%=diet_txt_033 %></span></th>
                       <th style="background-color:#D6DFf5;"><%=diet_txt_034 %></th>

                            <%if lto = "plan" or lto = "outz" then %>
                               <th style="width: 8%; background-color:#ff9797;"><%=diet_txt_027 %><br /><span style="font-size:9px;"><%=diet_txt_028 %></span></th>
                               <th style="background-color:#ff9797;"><%=diet_txt_029 %></th>
                            <%end if %>

                       <%end if %>
                       <th><%=diet_txt_035 %></th>
                    
                       <th><%=diet_txt_036 %></th>


                       <%
                        if cint(brug_godkend) = 1 then
                        %>

                            <th style="text-align:center;"><%=expence_txt_044 %> <br /> <input type="radio" name="approveall" id="approveall" value="1" <%=approveActive %> /></th>
                            <th style="text-align:center;"><%=expence_txt_046 %> <br /> <input type="radio" name="approveall" id="disapproveall" value="1" <%=approveActive %> /></th>
                            <th style="text-align:center;"><%=expence_txt_048 %> <br /> <input type="checkbox" id="afregnAlle" <%=afregnetActive %> /></th>
                        <%end if %>


                       <th><%=diet_txt_037 %> </th>


                   </tr>

                
    <%

         end if 'media

     

       call jobliste


        call licKid()

        '*** Kontoplan
        if licensindehaverKid <> 0 then
        licensindehaverKid = licensindehaverKid
        else
        licensindehaverKid = 0
        end if

         strKontoOptions = "<option value='0'>"&diet_txt_043&"</option>"
         strSQLKontoliste = "SELECT navn, kontonr, id FROM kontoplan "_
         &" WHERE kid = "& licensindehaverKid &" ORDER BY navn"
         oRec.open strSQLkontoliste, oConn, 3
         k = 0
         while not oRec.EOF

         strKontoOptions = strKontoOptions & "<option value="& oRec("id") &">"& oRec("kontonr") &"</option>" 

         
         k = k + 1
         oRec.movenext
	     wend
         oRec.close

        if cint(usemrn) <> 0 then
            selMedarbSQlKri = "AND diet_mid = "& usemrn 
        else
            if level <> 1 then
                selMedarbSQlKri = "AND diet_mid IN ("& allmedids &")"
            else
                selMedarbSQlKri = "AND diet_mid <> 0 "
            end if       
        end if

        strTxtExport = ""

        sqlSTDatoKri = year(aar) & "-" & month(aar) & "-" & day(aar)

	    strSQL = "SELECT diet_id, diet_mid, diet_namedest, diet_stdato, diet_sldato, diet_jobid, diet_approved, diet_approvedby, diet_approveddate, diet_settled, diet_settleddate, "_
        &" diet_konto, diet_dayprice, diet_dayprice_half, diet_delvis, diet_maksamount, diet_morgen, diet_morgenamount, "_
        &" diet_middag, diet_middagamount, diet_aften, diet_aftenamount, diet_rest, diet_min25proc, diet_total, diet_bilag, diet_traveldays,"_
        &" j.id AS jid, j.jobnavn, j.jobnr, k.navn AS kontonavn, k.kontonr"_
        &" FROM traveldietexp "_
        &" LEFT JOIN job AS j ON (j.id = diet_jobid)"_
        &" LEFT JOIN kontoplan AS k ON (k.id = diet_konto)"_
        &" WHERE diet_stdato >= '"& sqlSTDatoKri &"' AND YEAR(diet_stdato) = '"& year(aar) &"' "& selMedarbSQlKri &" ORDER BY diet_mid, diet_stdato"
	

        'if session("mid") = 1 then
        'response.write strSQL & "<br>"
        'response.Flush
        'end if


        d = 0
	    diet_dayprice = 0
        diet_dayprice_half = 0
        diet_morgenamount = 0
        diet_middagamount = 0
        diet_aftenamount = 0

        if cint(level) = 1 then
        mainAmountBoxes = ""
        mainAmountBoxesName = ""
        else
        mainAmountBoxes = "DISABLED"
        mainAmountBoxesName = "X"
        end if



        oRec.open strSQL, oConn, 3
        while not oRec.EOF

        'response.Write "<br> her " & diet_dayprice
        if d = 0 then

            
        'diet_dayprice = formatnumber(oRec("diet_dayprice"), 2)
        'diet_dayprice_half = formatnumber(oRec("diet_dayprice_half"), 2)
        'diet_morgenamount = formatnumber(oRec("diet_morgenamount"), 2)
        'diet_middagamount = formatnumber(oRec("diet_middagamount"), 2)
        'diet_aftenamount = formatnumber(oRec("diet_aftenamount"), 2)
       
     
        strSQL = "SELECT tdf_diet_dayprice, tdf_diet_dayprice_half, tdf_diet_morgenamount, tdf_diet_middagamount, tdf_diet_aftenamount FROM travel_diet_tariff WHERE tdf_diet_current = 1"
        oRec2.open strSQL, oConn, 3
        if not oRec2.EOF then
            diet_dayprice = formatnumber(oRec2("tdf_diet_dayprice"), 2)
            diet_dayprice_half = formatnumber(oRec2("tdf_diet_dayprice_half"), 2)
            diet_morgenamount = formatnumber(oRec2("tdf_diet_morgenamount"), 2)
            diet_middagamount = formatnumber(oRec2("tdf_diet_middagamount"), 2)
            diet_aftenamount = formatnumber(oRec2("tdf_diet_aftenamount"), 2)
        end if
        oRec2.close
       
        'response.Write "<br> diet_dayprice " & diet_dayprice
        'response.Write "<br> diet_dayprice_half " & diet_dayprice_half
        'response.Write "<br> diet_morgenamount " & diet_morgenamount
        'response.Write "<br> diet_middagamount " & diet_middagamount
        'response.Write "<br> diet_aftenamount " & diet_aftenamount

        if media <> "export" then
        %>

                  
                    
                      <tr>
                          <%if lto = "plan" or lto = "outz" then %>
                          <th>&nbsp;</th>
                          <th>&nbsp;</th>
                          <%end if %>

                          <th>&nbsp;</th>

                          <%if cint(brug_tidspunkt) = 1 then  %>
                            <th>&nbsp;</th>
                            <th>&nbsp;</th>
                          <%end if %>

                       <th>&nbsp;</th>
                     <th>&nbsp;</th>
                      <th>&nbsp;</th>
                      <th>&nbsp;</th>
                     <!-- <th>&nbsp;</th> kontonr -->
                    
                       <th style="text-align:right;">
                           
                           <input type="hidden" value="<%=diet_dayprice %>" name="FM_diet_dayprice" class="form-control input-small" style="font-size:9px; font-weight:lighter;" <%=mainAmountBoxes %>/>
                           <span><%=diet_dayprice %></span>

                       </th>
                      <%if lto <> "tia" then %>
                       <th>&nbsp;</th>
                       <%end if %>

                       <%if cint(vis_reduktion) = 1 then %>
                      <th style="text-align:right;">
                          <input type="hidden" value="<%=diet_morgenamount %>" name="FM_diet_morgenamount<%=mainAmountBoxesName %>" class="form-control input-small" style="font-size:9px; font-weight:lighter;" <%=mainAmountBoxes %> />
                          <span><%=diet_morgenamount %></span>
                      </th>

                       <th style="text-align:right;">
                           <input type="hidden" value="<%=diet_middagamount %>" name="FM_diet_middagamount<%=mainAmountBoxesName %>" class="form-control input-small" style="font-size:9px; font-weight:lighter;" <%=mainAmountBoxes %> />
                           <span><%=diet_middagamount %></span>
                       </th>

                        <th style="text-align:right;">
                            <input type="hidden" value="<%=diet_aftenamount %>" name="FM_diet_aftenamount<%=mainAmountBoxesName %>" class="form-control input-small" style="font-size:9px; font-weight:lighter;" <%=mainAmountBoxes %> />
                            <span><%=diet_aftenamount %></span>
                        </th>

                      <th>&nbsp;</th>

                        <%if lto = "plan" or lto = "outz" then%>
                        <th style="text-align:right;">
                            <input type="hidden" value="<%=diet_dayprice_half %>" name="FM_diet_dayprice_delvis" class="form-control input-small" style="font-size:9px; font-weight:lighter;" <%=mainAmountBoxes %>/>
                            <span><%=diet_dayprice_half %></span>
                        </th>
                       
                          <th>&nbsp;</th>
                        <%end if %>

                    <%else %>
                         <input type="hidden" value="0" name="FM_diet_morgenamount<%=mainAmountBoxesName %>"/>
                         <input type="hidden" value="0" name="FM_diet_middagamount<%=mainAmountBoxesName %>"/>
                         <input type="hidden" value="0" name="FM_diet_aftenamount<%=mainAmountBoxesName %>"/>
                      
                        <%end if %>

                     <th>&nbsp;</th>
                       <th>&nbsp;</th>


                        <%if cint(brug_godkend) = 1 then %>
                            <th></th>
                            <th></th>
                            <th></th>
                        <%end if %>


                       <th>&nbsp;</th>

                   </tr>

                    <%if mainAmountBoxes = "DISABLED" then %>
                   <input type="hidden" value="<%=diet_dayprice %>" name="FM_diet_dayprice" />
                   <input type="hidden" value="<%=diet_morgenamount %>" name="FM_diet_morgenamount" />
                   <input type="hidden" value="<%=diet_middagamount %>" name="FM_diet_middagamount"  />
                   <input type="hidden" value="<%=diet_aftenamount %>" name="FM_diet_aftenamount" />
                      <%end if %>


                        


                    </thead>
               <tbody>


        <%
        end if 'media

        end if




        if oRec("jid") <> 0 then
        strJobOptionsLoop = "<option value='"& oRec("jid") &"' SELECTED>"& left(oRec("jobnavn"), 25) &" ("& oRec("jobnr") &")</option>"  & strJobOptions
        else
        strJobOptionsLoop = strJobOptions
        end if

        if oRec("diet_konto") <> 0 then
        strKontoOptionsLoop = "<option value='"& oRec("diet_konto") &"' SELECTED>"& oRec("kontonr") &"</option>"  & strKontoOptions
        else
        strKontoOptionsLoop = strKontoOptions
        end if


        if cDate(oRec("diet_stdato")) = "01-01-2010" then
        diet_stdato = ""
        else
            if cint(hide_klokkeslet) = 0 then
            diet_stdato_len = len(oRec("diet_stdato"))
            diet_stdato_left = left(oRec("diet_stdato"), diet_stdato_len - 3)
            diet_stdato = diet_stdato_left
            else
            diet_stdato = formatdatetime(oRec("diet_stdato"), 2)
            end if
        
        end if

        str_ststhour = hour(oRec("diet_stdato"))
        str_stminute = minute(oRec("diet_stdato"))
                                        
        if str_ststhour < 10 then
        str_ststhour = "0" & str_ststhour
        else
        str_ststhour = str_ststhour
        end if

        if str_stminute < 10 then
        str_stminute = "0" & str_stminute
        else
        str_stminute = str_stminute
        end if

        str_sttime = str_ststhour &":"& str_stminute
        str_stdato = day(oRec("diet_stdato")) &"-"& month(oRec("diet_stdato")) &"-"& year(oRec("diet_stdato"))



        if cDate(oRec("diet_sldato")) = "01-01-2010" then
        diet_sldato = ""
        else
        
            if cint(hide_klokkeslet) = 0 then
            diet_sldato_len = len(oRec("diet_sldato"))
            diet_sldato_left = left(oRec("diet_sldato"), diet_sldato_len - 3)
            diet_sldato = diet_sldato_left
            else
            diet_sldato = formatdatetime(oRec("diet_sldato"), 2)
            end if

        end if

        str_stslhour = hour(oRec("diet_sldato"))
        str_slminute = minute(oRec("diet_sldato"))
                                        
        if str_stslhour < 10 then
        str_stslhour = "0" & str_stslhour
        else
        str_stslhour = str_stslhour
        end if

        if str_slminute < 10 then
        str_slminute = "0" & str_slminute
        else
        str_slminute = str_slminute
        end if

        str_sltime = str_stslhour &":"& str_slminute
        str_sldato = day(oRec("diet_sldato")) &"-"& month(oRec("diet_sldato")) &"-"& year(oRec("diet_sldato"))  

 
        if oRec("diet_morgen") <> 0 then
        diet_morgenTxt = formatnumber(oRec("diet_morgen"), 0)
        diet_morgenBel = formatnumber(oRec("diet_morgen")*oRec("diet_morgenamount"), 2)
        else
        diet_morgenTxt = ""
        diet_morgenBel = 0
        end if

        if oRec("diet_middag") <> 0 then
        diet_middagTxt = formatnumber(oRec("diet_middag"), 0)
        diet_middagBel = formatnumber(oRec("diet_middag")*oRec("diet_middagamount"), 2)
        else
        diet_middagTxt = ""
        diet_middagBel = 0
        end if

        if oRec("diet_aften") <> 0 then
        diet_aftenTxt = formatnumber(oRec("diet_aften"), 0)
        diet_aftenBel = formatnumber(oRec("diet_aften")*oRec("diet_aftenamount"), 2)
        else
        diet_aftenTxt = ""
        diet_aftenBel = 0
        end if

        if oRec("diet_total") <> 0 then
        diet_totalTxt = formatnumber(oRec("diet_total"), 0)
        else
        diet_totalTxt = ""
        end if
        
        diet_dageIalt = oRec("diet_traveldays")' dateDiff("h", oRec("diet_stdato"), oRec("diet_sldato"), 2,2) / 24
        if diet_dageIalt <> 0 then
            'if instr(diet_dageIalt, ",") <> 0 then
            'komma = instr(diet_dageIalt, ",")
            'diet_dageIaltTxt = left(diet_dageIalt, komma-1)
            'else
            diet_dageIaltTxt = diet_dageIalt
            'end if

        else
        diet_dageIaltTxt = ""
        end if


       'makskost =  formatnumber((diet_dayprice*diet_dageIalt), 0)
       
        if oRec("diet_maksamount") <> 0 then
            makskostTxt = formatnumber(oRec("diet_maksamount"),2)
            else
            makskostTxt = ""
        end if


        if oRec("diet_rest") <> 0 then
        diet_restTxt = formatnumber(oRec("diet_rest"), 2)
        else
        diet_restTxt = "0,00"
        end if

        'totalefterReduktion = formatnumber(makskost/1 - (diet_morgenBel/1 + diet_middagBel/1 + diet_aftenBel/1), 0)
         'if totalefterReduktion <> 0 then
          '  totalefterReduktionTxt = totalefterReduktion
          '  else
           ' totalefterReduktionTxt = ""
        'end if

        strMedabSel = ""
        strMedabSelInit = ""
                  
        call meStamdata(oRec("diet_mid"))
        strMedabSel = meNavn
        strMedabSelInit = meInit


                            if media <> "export" then
                            %>
                            <tr>
                                 <%if lto = "plan" or lto = "outz" then %>
                                <td style="text-align:center;">                                    
                                    <%
                                    select case oRec("diet_delvis")
                                    case 0
                                        chk1 = "CHECKED"
                                        chk2 = ""
                                    case else
                                        chk1 = ""
                                        chk2 = "CHECKED"
                                    end select
                                    %>
                                    
                                    <input type="radio" class="delvisradio" id="<%=oRec("diet_id") %>" name="FM_delvis_<%=oRec("diet_id") %>" value="0" <%=chk1 %> />                                    
                                </td>
                                <td style="text-align:center;">
                                    <input type="radio" class="delvisradio" id="<%=oRec("diet_id") %>" name="FM_delvis_<%=oRec("diet_id") %>" value="1" <%=chk2 %> /> 
                                </td>
                                <%end if %>

                                <td>[<%=meInit %>]</td>
                                <td>
                                    <input type="hidden" name="FM_diet_medid" value="<%=oRec("diet_mid")%>"/>
                                    <input type="hidden" name="FM_diet_id" value="<%=oRec("diet_id")%>"/>
                                   <!-- <input type="text" value="<%=diet_stdato%>" name="FM_diet_stdato" class="form-control input-small" /></td> -->

                                    <div class='input-group date' id='datepicker_stdato'>
                                        <input type="text" class="form-control input-small" name="FM_diet_stdato" value="<%=str_stdato%>" placeholder="dd-mm-yyyy" autocomplete="off" />
                                        <span class="input-group-addon input-small">
                                                <span class="fa fa-calendar">
                                                </span>
                                        </span>
                                    </div>

                                <%if cint(brug_tidspunkt) = 1 then %>
                                <td>
                                    <input type="time" class="form-control input-small" name="FM_diet_sttime_<%=oRec("diet_id") %>" value="<%=str_sttime %>" style="width:75px; text-align:left; vertical-align:middle;" />
                                </td>
                                <%end if %>

                                <td>
                                   <!-- <input type="text" value="<%=diet_sldato%>" name="FM_diet_sldato" class="form-control input-small" /> -->

                                    <div class='input-group date' id='datepicker_stdato'>
                                        <input type="text" class="form-control input-small" name="FM_diet_sldato" value="<%=str_sldato %>" placeholder="dd-mm-yyyy" autocomplete="off" />
                                        <span class="input-group-addon input-small">
                                                <span class="fa fa-calendar">
                                                </span>
                                        </span>
                                    </div>
                                </td>

                                <%if cint(brug_tidspunkt) = 1 then %>
                                <td>
                                    <input type="time" class="form-control input-small" name="FM_diet_sltime_<%=oRec("diet_id") %>" value="<%=str_sltime %>" style="width:75px; text-align:left; vertical-align:middle;" />
                                </td>
                                <%end if %>

                                <td><input type="text" value="<%=oRec("diet_namedest")%>" placeholder="Destination" name="FM_diet_namedest" class="form-control input-small" /></td>

                                <td class="small"><!--<select name="FM_diet_jobid" class="form-control input-small"><%=strJobOptionsLoop %></select>-->
                           <%
                           end if 'media

                           jj = 0
                           jobnrStrTxt = ""
                           jobnrStrTxtNr = ""
                           jobnrStrTxtProc = ""
                         
                           tj_percent = 0
                          

                            strSQLjobrel = "SELECT tj_trvlid, tj_jobid, tj_percent, jobnavn, jobnr, j.id AS jid FROM traveldietexp_jobrel "_
                            &" LEFT JOIN job AS j ON (j.id = tj_jobid) WHERE tj_trvlid = "& oRec("diet_id")
                            oRec2.open strSQLjobrel, oConn, 3
                            while not oRec2.EOF
        
                              tj_percent = formatnumber(oRec2("tj_percent"), 0)

                              jobnrStrTxt = oRec2("jobnavn")  
                              jobnrStrTxtNr = oRec2("jobnr")
                              jobnrStrTxtProc = tj_percent

                              
                              diet_restTxtBeregnet = formatnumber((oRec("diet_rest")) * (tj_percent/100), 2) 

                              strTxtExport = strTxtExport & strMedabSel &";"& strMedabSelInit &";"& str_stdato &";"

                              if cint(brug_tidspunkt) = 1 then
                                strTxtExport = strTxtExport & str_sttime &";"
                              end if

                              strTxtExport = strTxtExport & str_sldato &";"

                             if cint(brug_tidspunkt) = 1 then
                                strTxtExport = strTxtExport & str_sltime &";"
                             end if

                              
                              strTxtExport = strTxtExport& diet_dageIaltTxt &";"& Chr(34) & oRec("diet_namedest") & Chr(34) &";"& Chr(34) & jobnrStrTxt & Chr(34) &";"& jobnrStrTxtNr &";"& jobnrStrTxtProc &";"
                               
                                  if cint(vis_reduktion) = 1 then 
                                  strTxtExport = strTxtExport & makskostTxt &";"& diet_morgenTxt &";"& diet_middagTxt &";"& diet_aftenTxt &";"& diet_restTxtBeregnet &";"
                                  else
                                  strTxtExport = strTxtExport & makskostTxt &";"
                                  end if  


                               if cint(brug_godkend) = 1 then

                                    if cint(oRec("diet_approved")) = 1 then
                                        ergodkendt = 1
                                    else
                                        ergodkendt = 0
                                    end if

                                    if cint(oRec("diet_approved")) = 2 then
                                        erafvist = 1
                                    else
                                        erafvist = 0
                                    end if

                                    strTxtExport = strTxtExport & ergodkendt &";"& erafvist &";"& oRec("diet_settled") &";"& vbcrlf
                               else
                                    strTxtExport = strTxtExport & vbcrlf
                               end if
                              
                               
                                  '&makskostTxt &";"& diet_morgenTxt &";"& diet_middagTxt &";"& diet_aftenTxt &";"& diet_restTxtBeregnet &";"& vbcrlf


                              %>
                                <%=left(oRec2("jobnavn"), 10) & " ("& oRec2("jobnr") & ") : "& tj_percent & "%"%><br />
                              <%
                   

                            jj = jj + 1
                            oRec2.movenext
                            wend
                            oRec2.close
                                  
                                  
                                  if jj = 0 then

                                    'diet_restTxtBeregnet = formatnumber((oRec("diet_rest")) * (tj_percent/100), 2) 

                                    strTxtExport = strTxtExport & strMedabSel &";"& strMedabSelInit &";"& str_stdato &";"

                                    if cint(brug_tidspunkt) = 1 then
                                    strTxtExport = strTxtExport & str_sttime &";"
                                    end if

                                    strTxtExport = strTxtExport & str_sldato &";"

                                    if cint(brug_tidspunkt) = 1 then
                                    strTxtExport = strTxtExport & str_sltime &";"
                                    end if

                                    strTxtExport = strTxtExport & diet_dageIaltTxt &";"& Chr(34) & oRec("diet_namedest") & Chr(34) &";"& Chr(34) & jobnrStrTxt & Chr(34) &";"& jobnrStrTxtNr &";"& jobnrStrTxtProc &";"
                                  
                                    if cint(vis_reduktion) = 1 then 
                                    strTxtExport = strTxtExport & makskostTxt &";"& diet_morgenTxt &";"& diet_middagTxt &";"& diet_aftenTxt &";"& diet_restTxt &";"
                                    else
                                    strTxtExport = strTxtExport & makskostTxt &";"
                                    end if     


                                    if cint(brug_godkend) = 1 then

                                        if cint(oRec("diet_approved")) = 1 then
                                            ergodkendt = 1
                                        else
                                            ergodkendt = 0
                                        end if

                                        if cint(oRec("diet_approved")) = 2 then
                                            erafvist = 1
                                        else
                                            erafvist = 0
                                        end if

                                        strTxtExport = strTxtExport & ergodkendt &";"& erafvist &";"& oRec("diet_settled") &";"& vbcrlf
                                    else
                                        strTxtExport = strTxtExport & vbcrlf
                                    end if

                               
                                  end if 'jj

                            if media <> "export" then
                                    %>
                                       <input type="hidden" name="FM_diet_jobid" value="0" />
                                       <input type="hidden" name="FM_diet_konto" value="0" />
                                       <a href="#" onclick="Javascript:window.open('traveldietexp.asp?func=addmoorejobs&trvlid=<%=oRec("diet_id")%>&medarb=<%=oRec("diet_mid")%>', '', 'width=600,height=600,top=100,left=400,resizable=no')" target="_parent" class="small"><%=diet_txt_038 %> >></a>
                                   </td>
                                   <!--<td><select name="FM_diet_konto" class="form-control input-small"><%=strKontoOptionsLoop %></select></td>-->

                                    <%if lto = "plan" OR lto = "outz" then %>

                                        <%
                                        if cint(oRec("diet_delvis")) = 1 then
                                            clsfuld = "none"
                                            clsdelvis = ""
                                        else
                                            clsfuld = ""
                                            clsdelvis = "none"
                                        end if
                                        %>


                                        <td style="text-align:center; display:<%=clsdelvis%>;" class="delvis" id="<%=oRec("diet_id") %>">-</td>
                                        <td style="text-align:center; display:<%=clsdelvis%>;" class="delvis" id="<%=oRec("diet_id") %>">-</td>

                                        <td style="text-align:right; display:<%=clsfuld%>;" class="fuld" id="<%=oRec("diet_id") %>"><%=diet_dageIaltTxt %></td>
                                        <td style="text-align:right; display:<%=clsfuld%>;" class="fuld" id="<%=oRec("diet_id") %>"><%=makskostTxt %></td>  

                                   <%else %>
                                       <td style="text-align:right;"><%=diet_dageIaltTxt %></td>
                                        <%if lto <> "tia" then %>
                                       <td style="text-align:right;"><%=makskostTxt %></td>        
                                        <%end if %>
                                   <%end if %>

                                    
                                   <%if cint(vis_reduktion) = 1 then %>

                                        <%if lto = "plan" OR lto = "outz" then %>  
                                
                                            <td style="display:<%=clsfuld%>;" class="fuld" id="<%=oRec("diet_id") %>"><input style="text-align:right;" type="text" value="<%=diet_morgenTxt%>" name="FM_diet_morgen" class="form-control input-small fuld_input" id="<%=oRec("diet_id") %>" /></td>
                                            <td style="display:<%=clsfuld%>;" class="fuld" id="<%=oRec("diet_id") %>"><input style="text-align:right;" type="text" value="<%=diet_middagTxt%>" name="FM_diet_middag" class="form-control input-small fuld_input" id="<%=oRec("diet_id") %>" /></td>
                                            <td style="display:<%=clsfuld%>;" class="fuld" id="<%=oRec("diet_id") %>"><input style="text-align:right;" type="text" value="<%=diet_aftenTxt%>" name="FM_diet_aften" class="form-control input-small fuld_input" id="<%=oRec("diet_id") %>" /></td>

                                            <td style="text-align:center; display:<%=clsdelvis%>;" class="delvis" id="<%=oRec("diet_id") %>"> - </td>
                                            <td style="text-align:center; display:<%=clsdelvis%>;" class="delvis" id="<%=oRec("diet_id") %>"> - </td>
                                            <td style="text-align:center; display:<%=clsdelvis%>;" class="delvis" id="<%=oRec("diet_id") %>"> - </td>

                                            <td style="text-align:center"><span class="fuld" id="<%=oRec("diet_id") %>"><%=diet_totalTxt%></span></td>

                                        <%else %>

                                            <td><input style="text-align:right;" type="text" value="<%=diet_morgenTxt%>" name="FM_diet_morgen" class="form-control input-small" /></td>
                                            <td><input style="text-align:right;" type="text" value="<%=diet_middagTxt%>" name="FM_diet_middag" class="form-control input-small" /></td>
                                            <td><input style="text-align:right;" type="text" value="<%=diet_aftenTxt%>" name="FM_diet_aften" class="form-control input-small" /></td>
                                            <td style="text-align:center"><%=diet_totalTxt%></td>

                                        <%end if %>
                                    
                                    <%if lto = "plan" or lto = "outz" then %>
                                            <td style="text-align:right; display:<%=clsdelvis%>;" class="delvis" id="<%=oRec("diet_id") %>"><%=diet_dageIaltTxt %></td>
                                            <td style="text-align:right; display:<%=clsdelvis%>;" class="delvis" id="<%=oRec("diet_id") %>"><%=makskostTxt %></td>
                                       
                                            <td style="text-align:center; display:<%=clsfuld%>;" class="fuld" id="<%=oRec("diet_id") %>">-</td>
                                            <td style="text-align:center; display:<%=clsfuld%>;" class="fuld" id="<%=oRec("diet_id") %>">-</td>
                                    <%end if %>


                                   <%else %>
                                   <input type="hidden" value="0" name="FM_diet_morgen" />
                                   <input type="hidden" value="0" name="FM_diet_middag" />
                                   <input type="hidden" value="0" name="FM_diet_aften" />
                                   <%end if %>

                                   <td style="text-align:center;">
                                         <%if cint(oRec("diet_bilag")) = 1 then
                                         diet_bilagCHK = "CHECKED"
                                         else
                                         diet_bilagCHK = ""
                                         end if
                                          %>
                         
                                     <input type="checkbox" value="1" name="FM_diet_bilag" <%=diet_bilagCHK %>/></td>
                                   <td style="text-align:right;"><%=diet_restTxt %></td>
                   
                        
                                    <%if cint(brug_godkend) = 1 then %>

                                        <%
                                        select case cint(oRec("diet_approved"))
                                            case 1
                                                godkend_CHK = "CHECKED"
                                                afvist_CHK = ""
                                            case 2
                                                godkend_CHK = ""
                                                afvist_CHK = "CHECKED"
                                            case else
                                                godkend_CHK = ""
                                                afvist_CHK = ""
                                        end select
                                       
                            
                                        thisapproveActive = approveActiveInt
                                        afregnet_CHK = ""
                                        'response.Write "Setteled " & oRec("diet_settled")
                                        if cint(oRec("diet_settled")) = 1 then
                                            thisapproveActive = 0
                                            afregnet_CHK = "CHECKED"
                                        end if
                                        %>
                                        
                                        <input type="hidden" name="FM_godkendtaf_last_<%=oRec("diet_id") %>" value="<%=oRec("diet_approvedby") %>" />
                                        <input type="hidden" name="FM_godkend_date_<%=oRec("diet_id") %>" value="<%=oRec("diet_approveddate") %>" />
                                        <input type="hidden" name="FM_godkend_last_<%=oRec("diet_id") %>" value="<%=oRec("diet_approved") %>" />

                                        <input type="hidden" name="FM_last_afregnet_<%=oRec("diet_id") %>" value="<%=oRec("diet_settled") %>" />
                                        <input type="hidden" name="FM_afregnet_dato_<%=oRec("diet_id") %>" value="<%=oRec("diet_settleddate") %>" />

                                        <th style="text-align:center;">

                                            <%if thisapproveActive = 1 then %>
                                            <input class="godkendradio" type="radio" name="FM_diet_godkend_<%=oRec("diet_id") %>" <%=godkend_CHK %> value="1" /> 
                                            <%else %>
                                            <input DISABLED type="radio" <%=godkend_CHK %> />
                                            <input type="hidden" name="FM_diet_godkend_<%=oRec("diet_id") %>" value="<%=oRec("diet_approved") %>" />
                                            <%end if %>

                                            <%if cint(oRec("diet_approved")) = 1 then %>
                                            <br /> <span style="font-size:9px;"><%=oRec("diet_approveddate") %></span>
                                            <%end if %>
                                        </th>
                                        <th style="text-align:center;">

                                            <%if thisapproveActive = 1 then %>
                                            <input class="afvisradio" <%=thisapproveActive %> type="radio" name="FM_diet_godkend_<%=oRec("diet_id") %>" <%=afvist_CHK %> value="2" />
                                            <%else %>
                                            <input type="radio" disabled <%=afvist_CHK %> />
                                            <%end if %>

                                            <%if cint(oRec("diet_approved")) = 2 then %>
                                            <br /> <span style="font-size:9px;"><%=oRec("diet_approveddate") %></span>
                                            <%end if %>
                                        </th>
                                        <th style="text-align:center;">
                                            <%if afregnetActiveInt = 1 then %>
                                            <input class="afregndiet" type="checkbox" name="FM_afreget_<%=oRec("diet_id") %>" <%=afregnet_CHK %> value="1" />
                                            <%else %>
                                            <input type="checkbox" <%=afregnet_CHK %> value="1" DISABLED />
                                            <input type="hidden" <%=afregnet_CHK %> name="FM_afreget_<%=oRec("diet_id") %>" value="<%=oRec("diet_settled") %>" />
                                            <%end if %>

                                            <%if cint(oRec("diet_settled")) = 1 then %>
                                                <br /> <span style="font-size:9px;"><%=oRec("diet_settleddate") %></span>
                                            <%end if %>
                                        </th>
                                    <%end if %>


                                   <td style="text-align:center;">
                                       <%if cint(brug_godkend) <> 1 then %>
                                            <a href="traveldietexp.asp?menu=tok&func=slet&id=<%=oRec("diet_id")%>&aar=<%=aar%>&medarb=<%=oRec("diet_mid")%>"><span style="color:darkred;" class="fa fa-times"></span></a>
                                       <%else %>
                                            <%if cint(oRec("diet_settled")) <> 1 then %>
                                                <a href="traveldietexp.asp?menu=tok&func=slet&id=<%=oRec("diet_id")%>&aar=<%=aar%>&medarb=<%=oRec("diet_mid")%>"><span style="color:darkred;" class="fa fa-times"></span></a>                                       
                                            <%end if %>
                                       <%end if %>
                                   </td>

                                   <input type="hidden" value="##"  name="FM_diet_stdato" />
                                   <input type="hidden" value="##" name="FM_diet_sldato" />
                                   <input type="hidden" value="##" name="FM_diet_namedest"  />
                                   <input type="hidden" value="##" name="FM_diet_morgen"  />
                                   <input type="hidden" value="##" name="FM_diet_middag"  />
                                   <input type="hidden" value="##" name="FM_diet_aften"  />
                                   <input type="hidden" value="##" name="FM_diet_bilag"  />
                                   
                               </tr>
                           <% 
                            end if 'media

                               
                    d = d + 1
	                oRec.movenext
	                wend
                    oRec.close




                    '************************************************************************************************
                    '**** TILFØJ NYE LINJER
                    '************************************************************************************************

                    '************************************************************************************************
                    '**** Findes der inge rejsedage i valgt år for medarb. findes dagspris for andre medarbejdere    
                    '************************************************************************************************
                                      
                    if cint(d) = 0 AND usemrn <> 0 then

                               
                    if media <> "export" then


                    'Tager altid bare den sidste som anbefalte pris IKKE KUN INDEFOR ÅR Da der så kan opstå 0
	                'strSQLsatser = "SELECT diet_dayprice, diet_dayprice_half, diet_morgenamount, "_
                    '&" diet_middagamount, diet_aftenamount "_
                    '&" FROM traveldietexp "_
                    '&" WHERE diet_dayprice <> 0 AND diet_dayprice_half <> 0 ORDER BY diet_id DESC"
	
                    'YEAR(diet_stdato) = '"& aar &"' AND
                    'response.write strSQLsatser
                    'response.Flush
        
                    d = 0
	                diet_dayprice = 0
                    diet_dayprice_half = 0
                    diet_morgenamount = 0
                    diet_middagamount = 0
                    diet_aftenamount = 0


                    'oRec.open strSQLsatser, oConn, 3
                    'if not oRec.EOF then

                                   'if cint(d) = 0 then

                                   'strSQL_tdf = "SELECT tdf_diet_name, tdf_diet_current, tdf_diet_dayprice, tdf_diet_dayprice_half, tdf_diet_morgenamount, tdf_diet_middagamount, tdf_diet_aftenamount FROM travel_diet_tariff WHERE tdf_diet_current = 1"
                                    'tdf_fundet = 0
                                    'oRec6.open strSQL_tdf, oConn, 3
                                    'if not oRec6.EOF then

                                        'tdf_fundet = 1
                                        'diet_dayprice = formatnumber(oRec6("tdf_diet_dayprice"), 2)
                                        'diet_dayprice_half = formatnumber(oRec6("tdf_diet_dayprice_half"), 2)
                                        'diet_morgenamount = formatnumber(oRec6("tdf_diet_morgenamount"), 2)
                                        'diet_middagamount = formatnumber(oRec6("tdf_diet_middagamount"), 2)
                                        'diet_aftenamount = formatnumber(oRec6("tdf_diet_aftenamount"), 2)
                               
                                        'tdf_diet_name = oRec6("tdf_diet_name")

                                    'end if
                                    'oRec6.close

                                    'end if


                                'if cint(tdf_fundet) = 0 then 'Gl for dem der ikke har tastet current


                                'diet_dayprice = formatnumber(oRec("diet_dayprice"), 2)
                                'diet_dayprice_half = formatnumber(oRec("diet_dayprice_half"), 2)
                                'diet_morgenamount = formatnumber(oRec("diet_morgenamount"), 2)
                                'diet_middagamount = formatnumber(oRec("diet_middagamount"), 2)
                                'diet_aftenamount = formatnumber(oRec("diet_aftenamount"), 2)

                               'tdf_diet_name = "#"


                                'end if

                                'Henter tariff udendem er er diæter 0
                                tdf_diet_name = ""
                                strSQL = "SELECT tdf_diet_dayprice, tdf_diet_dayprice_half, tdf_diet_morgenamount, tdf_diet_middagamount, tdf_diet_aftenamount, tdf_diet_name FROM travel_diet_tariff WHERE tdf_diet_current = 1"
                                oRec2.open strSQL, oConn, 3
                                    if not oRec2.EOF then
                                        diet_dayprice = formatnumber(oRec2("tdf_diet_dayprice"), 2)
                                        diet_dayprice_half = formatnumber(oRec2("tdf_diet_dayprice_half"), 2)
                                        diet_morgenamount = formatnumber(oRec2("tdf_diet_morgenamount"), 2)
                                        diet_middagamount = formatnumber(oRec2("tdf_diet_middagamount"), 2)
                                        diet_aftenamount = formatnumber(oRec2("tdf_diet_aftenamount"), 2)
                                        tdf_diet_name = oRec2("tdf_diet_name")
                                    end if
                                oRec2.close
                    %>

                  

                                  <tr>
                                      <%if lto = "plan" or lto = "outz" then %>
                                    <th>&nbsp;</th>
                                      <th>&nbsp;</th>
                                      <%end if %>
                                    <th>&nbsp;</th>
                                   <th>&nbsp;</th>
                                      <th>&nbsp;</th>
                                   <th>&nbsp;</th>
                                 <th>&nbsp;</th>
                                  <th>&nbsp;</th>
                                <%if lto <> "tia" then %>
                                  <th>&nbsp;</th>
                                <%end if %>
                                 <!-- <th>&nbsp;</th> kontonr -->
                    
                                   <th style="text-align:right;">
                                       
                                       <%if mainAmountBoxes = "DISABLED" then %>
                                       <input type="hidden" value="<%=diet_dayprice %>" name="" class="form-control input-small" style="font-size:9px; font-weight:lighter;" <%=mainAmountBoxes %>/>
                                       <input type="hidden" value="<%=diet_dayprice %>" name="FM_diet_dayprice"/>
                                       <%else %>
                                       <input type="hidden" value="<%=diet_dayprice %>" name="FM_diet_dayprice" class="form-control input-small" style="font-size:9px; font-weight:lighter;" <%=mainAmountBoxes %>/>
                                       <%end if %>

                                       <span><%=diet_dayprice %></span>

                                       <%'if tdf_diet_name <> "" then %>
                                       <!-- <span style="font-size:9px;"> (<%=tdf_diet_name %>)</span> -->
                                       <%'end if %>
                                   
                                   </th>
                                   <th>&nbsp;</th>
                                     <%if cint(vis_reduktion) = 1 then %>

                                  <th style="text-align:right;">
                                      <input type="hidden" value="<%=diet_morgenamount %>" name="FM_diet_morgenamount<%=mainAmountBoxesName %>" class="form-control input-small" style="font-size:9px; font-weight:lighter;" <%=mainAmountBoxes %> />
                                      <span><%=diet_morgenamount %></span>
                                  </th>
                                  
                                      <th style="text-align:right;">
                                          <input type="hidden" value="<%=diet_middagamount %>" name="FM_diet_middagamount<%=mainAmountBoxesName %>" class="form-control input-small" style="font-size:9px; font-weight:lighter;" <%=mainAmountBoxes %> />
                                          <span><%=diet_middagamount %></span>
                                      </th>
                                   
                                      <th style="text-align:right;">
                                          <input type="hidden" value="<%=diet_aftenamount %>" name="FM_diet_aftenamount<%=mainAmountBoxesName %>" class="form-control input-small" style="font-size:9px; font-weight:lighter;" <%=mainAmountBoxes %> />
                                          <span><%=diet_aftenamount %></span>
                                      </th>
                                 
                                      <th>&nbsp;</th>

                                      <%if lto = "plan" or lto = "outz" then %>
                                            <th style="text-align:right;">
                                                <%if mainAmountBoxes = "DISABLED" then %>
                                               <input type="hidden" value="<%=diet_dayprice_half %>" name="" class="form-control input-small" style="font-size:9px; font-weight:lighter;" <%=mainAmountBoxes %>/>
                                               <input type="hidden" value="<%=diet_dayprice_half %>" name="FM_diet_dayprice_delvis"/>
                                                <%else %>
                                                <input type="hidden" value="<%=diet_dayprice_half %>" name="FM_diet_dayprice_delvis" class="form-control input-small" style="font-size:9px; font-weight:lighter;" <%=mainAmountBoxes %>/>
                                                <%end if %>

                                                <span><%=diet_dayprice_half %></span>
                                            </th>

                                           <th>&nbsp</th> 

                                      <%end if %>

                                   <%else %>
                                   <input type="hidden" value="0" name="FM_diet_morgenamount<%=mainAmountBoxesName %>" />
                                   <input type="hidden" value="0" name="FM_diet_middagamount<%=mainAmountBoxesName %>" />
                                   <input type="hidden" value="0" name="FM_diet_aftenamount<%=mainAmountBoxesName %>" />
                                  
                                <%end if %>

                                    
                                 <th>&nbsp;</th>

                                    <%if lto <> "tia" then %>
                                   <th>&nbsp;</th>

                                    <%if cint(brug_godkend) = 1 then %>
                                    <th></th>
                                    <th></th>
                                    <th></th>
                                   <%end if %>


                                   <th>&nbsp;</th>
                                    <%end if %>


                                    
                               </tr>

                                <%if mainAmountBoxes = "DISABLED" then %>
                               <input type="hidden" value="<%=diet_morgenamount %>" name="FM_diet_morgenamount" />
                               <input type="hidden" value="<%=diet_middagamount %>" name="FM_diet_middagamount"  />
                               <input type="hidden" value="<%=diet_aftenamount %>" name="FM_diet_aftenamount" />
                                  <%end if %>


                                </thead>
                           <tbody>


                    <%
                    'end if
                    'oRec.close

                    end if 'media


                end if

	                if cint(hide_klokkeslet) = 0 then
                        daytimeFormatPlaceholder = "dd-mm-yyyy tt:mm"
                    else
                        daytimeFormatPlaceholder = "dd-mm-yyyy"
                    end if

                       
                   if media <> "export" then
                        
                        
                   if cint(usemrn) <> 0 then
                   %>
                   <tr>
                       <%if lto = "plan" or lto = "outz" then %>
                       <td style="text-align:center;">
                           <input type="radio" class="delvisradio" id="0" name="FM_delvis_0" value="0" checked />                           
                       </td>
                       <td style="text-align:center;">
                           <input type="radio" class="delvisradio" id="0" name="FM_delvis_0" value="1" />
                       </td>
                       <%end if %>
                       <td>&nbsp;</td>
                       <td>
                           <input type="hidden" name="FM_diet_id" value="0"/>
                           <input type="hidden" name="FM_diet_medid" value="<%=usemrn%>"/>
                          <!-- <input class="form-control input-small" type="text" name="FM_diet_stdato" value="" placeholder="<%=daytimeFormatPlaceholder %>"/>-->
                           <div class='input-group date' id='datepicker_stdato'>
                                <input type="text" class="form-control input-small" name="FM_diet_stdato" value="" placeholder="dd-mm-yyyy" autocomplete="off" />
                                <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                </span>
                          </div>
                       </td>

                       <%if cint(brug_tidspunkt) = 1 then %>
                       <td><input class="form-control input-small" type="time" name="FM_diet_sttime_0" /></td>
                       <%end if %>

                       <td>
                           <!--<input class="form-control input-small" type="text" name="FM_diet_sldato" value="" placeholder="<%=daytimeFormatPlaceholder %>"/>-->
                           <div class='input-group date' id='datepicker_stdato'>
                                <input type="text" class="form-control input-small" name="FM_diet_sldato" value="" placeholder="dd-mm-yyyy" autocomplete="off" />
                                <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                </span>
                            </div>
                       </td>

                       <%if cint(brug_tidspunkt) = 1 then %>
                       <td><input class="form-control input-small" type="time" name="FM_diet_sltime_0" /></td>
                       <%end if %>

                       <td><input type="text" value="" placeholder="Destination" name="FM_diet_namedest" class="form-control input-small" /></td>
                       <td><!--<select name="FM_diet_jobid" class="form-control input-small"><%=strJobOptions %></select>-->
                           <input type="hidden" name="FM_diet_jobid" value="0" />
                           <input type="hidden" name="FM_diet_konto" value="0" />
                           &nbsp;</td>
                       <!--<td><select name="FM_diet_konto" class="form-control input-small"><%=strKontoOptions %></select></td>-->

                       <%if lto = "plan" or lto = "outz" then %>
                            <td class="fuld" id="0"></td>
                            <td class="fuld" id="0"></td>

                            <td style="text-align:center; display:none;" class="delvis" id="0">-</td>
                            <td style="text-align:center; display:none;" class="delvis" id="0">-</td>
                       <%else %>
                           <td>&nbsp;</td>
                         <%if lto <> "tia" then %>
                           <td>&nbsp;</td>
                         <%end if %>
                       <%end if %>

                       <%if cint(vis_reduktion) = 1 then %>


                           <%if lto = "plan" OR lto = "outz" then %>
                               <td class="fuld" id="0"><input style="text-align:right;" type="text" value="" name="FM_diet_morgen" class="form-control input-small fuld_input" id="0" /></td>
                               <td class="fuld" id="0"><input style="text-align:right;" type="text" value="" name="FM_diet_middag" class="form-control input-small fuld_input" id="0" /></td>
                               <td class="fuld" id="0"><input style="text-align:right;" type="text" value="" name="FM_diet_aften" class="form-control input-small fuld_input" id="0" /></td>
                        
                               <td style="text-align:center; display:none;" class="delvis" id="0">-</td>
                               <td style="text-align:center; display:none;" class="delvis" id="0">-</td>
                               <td style="text-align:center; display:none;" class="delvis" id="0">-</td>

                                <td>&nbsp</td>

                           <%else %>
                               <td><input style="text-align:right;" type="text" value="" name="FM_diet_morgen" class="form-control input-small" /></td>
                               <td><input style="text-align:right;" type="text" value="" name="FM_diet_middag" class="form-control input-small" /></td>
                               <td><input style="text-align:right;" type="text" value="" name="FM_diet_aften" class="form-control input-small" /></td>
                               <td>&nbsp;</td>
                            <%end if %>

                        

                       <%if lto = "plan" or lto = "outz" then %>
                            <th class="fuld" id="0" style="text-align:center;">-</th>
                            <th class="fuld" id="0" style="text-align:center;">-</th>

                            <th class="delvis" id="0" style="display:none;"></th>
                            <th class="delvis" id="0" style="display:none;"></th>
                       <%end if %>

                       <%else %>
                        <input type="hidden" value="" name="FM_diet_morgen" class="form-control input-small" />
                       <input type="hidden" value="" name="FM_diet_middag" class="form-control input-small" />
                       <input type="hidden" value="" name="FM_diet_aften" class="form-control input-small" />
                       <%end if %>
                       
                      
                       <td style="text-align:center;"><input type="checkbox" value="1" name="FM_diet_bilag"/></td>
                       <td>&nbsp;</td>


                        <%if cint(brug_godkend) = 1 then %>
                            <input type="hidden" name="FM_godkend_last_0" value="0" />
                            <input type="hidden" name="FM_last_afivst_0" value="0" />
                            <input type="hidden" name="FM_godkendtaf_last_0" value="" />

                            <th style="text-align:center;"><input type="radio" class="godkendradio" name="FM_diet_godkend_0" value="1" <%=approveActive %> /></th>
                            <th style="text-align:center;"><input type="radio" class="afvisradio" name="FM_diet_godkend_0" value="2" <%=approveActive %> /></th>
                            <th style="text-align:center;"><input class="afregndiet" type="checkbox" name="FM_afreget_0" value="1" <%=afregnetActive %> /></th>
                        <%end if %>


                         <td>&nbsp;</td>

                       
                          
                       <input  type="hidden" value="##"  name="FM_diet_stdato" />
                       <input  type="hidden" value="##" name="FM_diet_sldato" />
                       <input type="hidden" value="##" name="FM_diet_namedest"  />
                       
                       
                       <input type="hidden" value="##" name="FM_diet_morgen"  />
                       <input type="hidden" value="##" name="FM_diet_middag"  />
                       <input type="hidden" value="##" name="FM_diet_aften"  />

                      <input type="hidden" value="##" name="FM_diet_bilag"  />
                       

                   </tr>
                   <%end if %>
               </tbody>     
           </table>
            
            
            <script type="text/javascript">
                  $('.date').datepicker({

                });
            </script>

             <section>
                         <div class="row">
                             <div class="col-lg-10">&nbsp;</div>
                             <div class="col-lg-2">
                            <input type="submit" value="<%=diet_txt_007 %> >>" class="btn btn-sm btn-success pull-right"><br />&nbsp;
                            </div>
                        </div>
                </section>
            
         </form>


         <br /><br />
                
            <section>
                <div class="row">
                     <div class="col-lg-12">
                        <b><%=diet_txt_039 %></b>
                        </div>
                    </div>

                  
                <form action="traveldietexp.asp?media=export&aar=<%=aar%>&medarb=<%=usemrn %>" method="post" target="_blank">
                  
                 <div class="row">
                     <div class="col-lg-12 pad-r30">
                         <input id="Submit6" type="submit" value="<%=diet_txt_040 %>" class="btn btn-sm" />
                        <!--Eksporter viste kunder som .csv fil-->
                         
                         </div>


                </div>
                </form>
                <br /><br />&nbsp;
            </section>
            <%else 'export 
                
                 call TimeOutVersion()
	
	
	            filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	            filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
	            Set objFSO = server.createobject("Scripting.FileSystemObject")
	
	            if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\to_2015\traveldietexp.asp" then
							
		            Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\traveldietexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
		            Set objNewFile = nothing
		            Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\traveldietexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
	
	            else
		
		            Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\traveldietexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, false)
		            Set objNewFile = nothing
		            Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\traveldietexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8, -1)
		
	            end if
	
	            file = "traveldietexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
	
	
	
	            'strTxtExportHeader = "Navn;Init;Afrejsedato;Hjemrejse;Dage ialt;Destination;Job/Projekt;Jobnr;Procent %;"
                strTxtExportHeader =  diet_txt_019&";"&diet_txt_045&";"&diet_txt_046&";"

                if cint(brug_tidspunkt) = 1 then
                strTxtExportHeader = strTxtExportHeader &";"
                end if

                strTxtExportHeader = strTxtExportHeader & diet_txt_047&";"

                if cint(brug_tidspunkt) = 1 then
                strTxtExportHeader = strTxtExportHeader &";"
                end if

                strTxtExportHeader = strTxtExportHeader & diet_txt_048&";"&diet_txt_049&";"&diet_txt_050&";"&diet_txt_051&";"&diet_txt_052 &" %;"

                if cint(vis_reduktion) = 1 then 
                'strTxtExportHeader = strTxtExportHeader & "Maks.kost;Morgen;Middag;Aften;Ialt/Rest.;" & vbcrlf
                strTxtExportHeader = strTxtExportHeader & diet_txt_053&";"&diet_txt_054&";"&diet_txt_055&";"&diet_txt_056&";"&diet_txt_057&";"
                else
                strTxtExportHeader = strTxtExportHeader & diet_txt_057 &";"
                end if
                
                
                if cint(brug_godkend) = 1 then
                strTxtExportHeader = strTxtExportHeader & expence_txt_044 &";"& expence_txt_046 &";"& expence_txt_048 &";"& vbcrlf
                else
                strTxtExportHeader = strTxtExportHeader & vbcrlf
                end if

		
                objF.WriteLine(strTxtExportHeader)
                objF.WriteLine(strTxtExport)
                objF.close


        'Response.end
	    Response.redirect "../inc/log/data/"& file &""	


		end if 'media %>

        </div>
    </div>
</div>

<%end select  %>
</div> 
</div>

<!--#include file="../inc/regular/footer_inc.asp"-->