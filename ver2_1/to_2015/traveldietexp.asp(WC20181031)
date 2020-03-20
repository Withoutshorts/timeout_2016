
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
    public strJobOptions
    sub jobliste

     '*** Job liste
         strJobOptions = "<option value='0'>Vælg Job</option>"
         strJobOptions = strJobOptions & "<option value='0'>Intet valgt</option>"
         strSQLjobliste = "SELECT jobnavn, jobnr, j.id, kkundenavn, kid FROM timereg_usejob AS tu "_
         &" LEFT JOIN job AS j ON (j.id = tu.jobid) "_
         &" LEFT JOIN kunder ON (kid = j.jobknr) "_
         &" WHERE tu.medarb = "& usemrn &" AND tu.aktid = 0 AND jobstatus = 1 AND risiko >= 0 ORDER BY kkundenavn, jobnavn"
       
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

        

        response.Redirect "traveldietexp.asp?func=addmoorejobs&medarb="& usemrn &"&trvlid="&trvlid &"&eropda=1"


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
               Tilføj flere job til Rejseplan/Diæt
            </h3>
            
            <div class="portlet-body">
                <form method="post" action="traveldietexp.asp?func=addmoorejobsOpd&trvlid=<%=trvlid%>&medarb=<%=usemrn %>">

                  

                <table class="table-stribed table" width="100%">
                    
                    <thead>
                        <tr>
                           
                            <th>Job</th>
                            <th>Procent (heltal)</th>
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
                                <p style="color:green;">Opdateret!</p>
                                <%end if %>

                             </div>
                             <div class="col-lg-1">
                            <input type="submit" value="Opdater >>" class="btn btn-sm btn-success pull-right"><br />&nbsp;
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
               Rejseplan/Diæt
            </h3>
            
            <div class="portlet-body" style="width:400px;">
                <div align="center">Du er ved at <b>SLETTE</b> en Rejseplan/Diæt. Er dette korrekt?
                </div><br />
                <div align="center"><b><a href="traveldietexp.asp?func=sletok&id=<%=id%>&aar=<%=aar%>&medarb=<%=usemrn%>">Ja</a></b>&nbsp&nbsp&nbsp&nbsp<b><a href="javascript:history.back()">Nej</a></b>
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

        if isDate(strStdato(i)) = true then
        strStdato(i) = year(strStdato(i)) &"-"& month(strStdato(i)) &"-"& day(strStdato(i)) &" "& formatdatetime(strStdato(i), 3)
        else
        strStdato(i) = year(now) &"-01-01 00:00"
        end if

        strSldato(i) = replace(strSldato(i), "#", "")
        strSldato(i) = replace(strSldato(i), ",", "")
       
        if isDate(strSldato(i)) = true then
        strSldato(i) = year(strSldato(i)) &"-"& month(strSldato(i)) &"-"& day(strSldato(i)) &" "& formatdatetime(strSldato(i), 3)
        else
        strSldato(i) = year(now) &"-01-01 00:00"
        end if
        

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

        makskost = (diet_daypriceVal/1)*diet_dageIalt
        totalefterReduktion = makskost/1 - ((strMorgenAntal(i)*diet_morgenamountVal/1) + (strMiddagAntal(i)*diet_middagamountVal/1) + (strAftenAntal(i)*diet_aftenamountVal/1))
        
        makskost = replace(makskost, ",", ".")
        diet_dageIalt = replace(diet_dageIalt, ",", ".")
        totalefterReduktion = replace(totalefterReduktion, ",", ".")

        if len(trim(strNavn(i))) <> 0 then 

		    if strIds(i) = 0 then
		    strSQlins = ("INSERT INTO traveldietexp (diet_namedest, diet_stdato, diet_sldato, diet_jobid, diet_konto, diet_morgen, diet_middag, diet_aften, diet_total, diet_dayprice, "_
            &" diet_morgenamount, diet_middagamount, diet_aftenamount, diet_mid, diet_bilag, diet_maksamount, diet_rest, diet_traveldays) "_
            &" VALUES ('"& strNavn(i) &"', '"& strStdato(i) &"', '"& strSldato(i) &"', "& strJobids(i) &", "& strkontoids(i) &", "_
            &" "& strMorgenAntal(i) &", "& strMiddagAntal(i) &", "& strAftenAntal(i) &", "& diet_total &", "& diet_dayprice &", "_
            &" "& diet_morgenamount &","& diet_middagamount &","& diet_aftenamount &", "& strDiet_medids(i) &", "& strBilag(i) &", "& makskost &", "& totalefterReduktion &", "& diet_dageIalt &")")

            'response.write strSQlins
            'response.flush

            oConn.execute(strSQlins) 

                
           

		    else
		    strSQlupd = ("UPDATE traveldietexp SET "_
            &" diet_namedest ='"& strNavn(i) &"', diet_stdato = '"& strStdato(i) &"', diet_sldato = '"& strSldato(i) &"', "_
            &" diet_jobid = "& strJobids(i) &", diet_konto = "& strkontoids(i) &", "_
            &" diet_morgen = "& strMorgenAntal(i) &", diet_middag = "& strMiddagAntal(i) &", diet_aften = "& strAftenAntal(i) &", diet_total = "& diet_total &", diet_dayprice = "& diet_dayprice &", "_
            &" diet_morgenamount ="& diet_morgenamount &", diet_middagamount = "& diet_middagamount &", "_
            &" diet_aftenamount =  "& diet_aftenamount &", diet_mid = "& strDiet_medids(i) &", diet_bilag = "& strBilag(i) &", diet_maksamount = "& makskost &", diet_rest = "& totalefterReduktion &", diet_traveldays = "& diet_dageIalt &""_
            &" WHERE diet_id = "& strIds(i) &"")

            'response.write strSQlupd
            'response.flush

            oConn.execute(strSQlupd)

		    end if

        else

             oConn.execute("DELETE FROM traveldietexp WHERE diet_id = "& strIds(i) &"")

        end if

        next
		
            

        'response.end
		Response.redirect "traveldietexp.asp?medarb="&usemrn&"&aar="&aar
		


    case "opret", "red"
	


case else    


          
            select case lto
            case "plan", "intranet - local"
            vis_reduktion = 0
            hide_klokkeslet = 1
            case else    
            vis_reduktion = 1 
            hide_klokkeslet = 0
            end select  




%>

<script src="js/traveldietexp_jav.js" type="text/javascript"></script>

<div class="container">
    <div class="porlet">
        <h3 class="portlet-title"><u>Rejseplan / Diæter</u></h3>


        <%if media <> "export" then %>

                
                       <form method="post" action="traveldietexp.asp?issubmitted=1">
                          
                            
                    
                           
                        <!-- NAVN / SORTERING / ID -->
                        <section>
                            <div class="well">
                                 <div class="row">
                                         <div class="col-lg-4 pad-t5">Medarbejder:<br /> 
                                             <%
                                                
                                                if cint(level) = 1 then
                                                 midSQLkr = " AND mid <> 0"
                                                else
                                                 midSQLkr = " AND mid = " & session("mid")
                                                end if
                                                 
                                               strSQlmedarblist = "SELECT mid, mnavn, init FROM medarbejdere WHERE mansat = 1 "& midSQLkr &" ORDER BY mnavn"
                                                 
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
                                                         
                                              if level = 1 then
                                                         
                                                         if cint(usemrn) = 0 then
                                                         visAlleSel = "SELECTED"
                                                         else
                                                         visAlleSel = ""
                                                         end if
                                                         %>

                                             <option value="0" <%=visAlleSel%>>Vis alle</option>
                                            <%end if %>
                                         </select>    
                                         </div>
                                        
                                       
                                                 
                                                <div class="col-lg-2 pad-t5">Fra:
                                                    <div class='input-group date' id='datepicker_stdato'>
                                                    <input type="text" class="form-control input-small" name="aar" value="<%=aar %>" placeholder="dd-mm-yyyy" />
                                                          <span class="input-group-addon input-small">
                                                                    <span class="fa fa-calendar">
                                                                    </span>
                                                                </span>
                                                           </div>
                                                </div>

                                             
                                            
                                              <div class="col-lg-1 pad-t20">
                                                    <button type="submit" class="btn btn-secondary btn-sm pull-right"><b>Søg >></b></button>
                                             </div>
                                
                                      
                                 </div>

                            </div>

                        </section>
                         </form>


     
        
        <form action="traveldietexp.asp?func=opdaterlist" method="post">
            <input type="hidden" name="medarb" value="<%=usemrn%>"/> 
            <input type="hidden" name="aar" value="<%=aar%>"/>
            <section>
                         <div class="row">
                             <div class="col-lg-10">&nbsp;</div>
                             <div class="col-lg-2">
                            <input type="submit" value="Opdater >>" class="btn btn-sm btn-success pull-right"><br />&nbsp;
                            </div>
                        </div>
                </section>
        
            
            

        <div class="porlet-body">
          
           <table id="kundetyper" class="table dataTable table-striped table-bordered table-hover ui-datatable">                 
               <thead>

                    <tr style="background-color:#FFFFFF;">
                      <th style="border:0;">&nbsp;</th>
                       <th style="border:0;">&nbsp;</th>
                       <th style="border:0;">&nbsp;</th>
                        <th style="border:0;"></th>
                        <th style="border:0;">&nbsp;</th>
                       <!-- <th style="border:0;">&nbsp;</th>-->
                        <th style="border:0;">&nbsp;</th>
                      <th style="border:0;">&nbsp;</th>
                      
                        <%if cint(vis_reduktion) = 1 then %>
                       <th colspan="4" style="border-bottom:0; background-color:#D6DFf5;">Reduktion</th>
                        <%end if %>
                      <th style="border:0;">&nbsp;</th>
                        <th style="border:0;">&nbsp;</th>
                        <th style="border:0;">&nbsp;</th>
                   </tr>

                   <tr>
                      <th style="width: 5%">Navn</th>
                       <th style="width: 15%">Afrejse<br /><span style="font-size:9px;">Dato <%if cint(hide_klokkeslet) = 0 then %> & Tid<%end if %></span></th>
                       <th style="width: 15%">Hjem<br /><span style="font-size:9px;">Dato <%if cint(hide_klokkeslet) = 0 then %>& Tid<%end if %></span></th>
                       <th style="width: 15%">Destination</th>
                       <th style="width: 15%">Job/Projektnr.<br /><span style="font-size:9px;">Aktiv jobliste</span></th>
                       <!--
                       <th style="width: 10%">Kontonr</th>
                       -->
                       <th style="width: 5%">Døgn<br /><span style="font-size:9px;">Dagspris</span></th>
                       <th>Maks. beløb</th>
                      
                        <%if cint(vis_reduktion) = 1 then %>
                       <th style="width: 7%; background-color:#D6DFf5;">Mor.<br /><span style="font-size:9px;">Antal stk.</span></th>
                       <th style="width: 7%; background-color:#D6DFf5;">Fro.<br /><span style="font-size:9px;">Antal stk.</span></th>
                       <th style="width: 7%; background-color:#D6DFf5;">Aft.<br /><span style="font-size:9px;">Antal stk.</span></th>
                       <th style="background-color:#D6DFf5;">Ialt</th>
                       <%end if %>
                       <th>Bilag</th>
                    
                       <th>Total</th>
                       <th>Slet</th>
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

         strKontoOptions = "<option value='0'>Vælg Konto</option>"
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
        selMedarbSQlKri = "AND diet_mid <> 0 "
        end if

        strTxtExport = ""

        sqlSTDatoKri = year(aar) & "-" & month(aar) & "-" & day(aar)

	    strSQL = "SELECT diet_id, diet_mid, diet_namedest, diet_stdato, diet_sldato, diet_jobid, "_
        &" diet_konto, diet_dayprice, diet_maksamount, diet_morgen, diet_morgenamount, "_
        &" diet_middag, diet_middagamount, diet_aften, diet_aftenamount, diet_rest, diet_min25proc, diet_total, diet_bilag, diet_traveldays,"_
        &" j.id AS jid, j.jobnavn, j.jobnr, k.navn AS kontonavn, k.kontonr"_
        &" FROM traveldietexp "_
        &" LEFT JOIN job AS j ON (j.id = diet_jobid)"_
        &" LEFT JOIN kontoplan AS k ON (k.id = diet_konto)"_
        &" WHERE diet_stdato >= '"& sqlSTDatoKri &"' AND YEAR(diet_stdato) = '"& year(aar) &"' "& selMedarbSQlKri &" ORDER BY diet_mid, diet_stdato"
	

        'response.write strSQL & "<br>"
        'response.Flush
        
        d = 0
	    diet_dayprice = 0
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


        if d = 0 then
        diet_dayprice = formatnumber(oRec("diet_dayprice"), 2)
        diet_morgenamount = formatnumber(oRec("diet_morgenamount"), 2)
        diet_middagamount = formatnumber(oRec("diet_middagamount"), 2)
        diet_aftenamount = formatnumber(oRec("diet_aftenamount"), 2)
       
       

        if media <> "export" then
        %>

                  
                    
                      <tr>
                      <th>&nbsp;</th>
                       <th>&nbsp;</th>
                     <th>&nbsp;</th>
                      <th>&nbsp;</th>
                      <th>&nbsp;</th>
                     <!-- <th>&nbsp;</th> kontonr -->
                    
                       <th>
                           
                           <input type="text" value="<%=diet_dayprice %>" name="FM_diet_dayprice" class="form-control input-small" style="font-size:9px; font-weight:lighter;" <%=mainAmountBoxes %>/>
                      

                       </th>
                       <th>&nbsp;</th>
                       <%if cint(vis_reduktion) = 1 then %>
                      <th><input type="text" value="<%=diet_morgenamount %>" name="FM_diet_morgenamount<%=mainAmountBoxesName %>" class="form-control input-small" style="font-size:9px; font-weight:lighter;" <%=mainAmountBoxes %> /></th>
                       <th><input type="text" value="<%=diet_middagamount %>" name="FM_diet_middagamount<%=mainAmountBoxesName %>" class="form-control input-small" style="font-size:9px; font-weight:lighter;" <%=mainAmountBoxes %> /></th>
                        <th><input type="text" value="<%=diet_aftenamount %>" name="FM_diet_aftenamount<%=mainAmountBoxesName %>" class="form-control input-small" style="font-size:9px; font-weight:lighter;" <%=mainAmountBoxes %> /></th>
                      <th>&nbsp;</th>
                    <%else %>
                         <input type="hidden" value="0" name="FM_diet_morgenamount<%=mainAmountBoxesName %>"/>
                         <input type="hidden" value="0" name="FM_diet_middagamount<%=mainAmountBoxesName %>"/>
                         <input type="hidden" value="0" name="FM_diet_aftenamount<%=mainAmountBoxesName %>"/>
                      
                        <%end if %>

                     <th>&nbsp;</th>
                       <th>&nbsp;</th>
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
        diet_restTxt = ""
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
                                <td>[<%=meInit %>]</td>
                                <td>
                                    <input type="hidden" name="FM_diet_medid" value="<%=oRec("diet_mid")%>"/>
                                    <input type="hidden" name="FM_diet_id" value="<%=oRec("diet_id")%>"/>
                                    <input type="text" value="<%=diet_stdato%>" name="FM_diet_stdato" class="form-control input-small" /></td>
                                <td><input type="text" value="<%=diet_sldato%>" name="FM_diet_sldato" class="form-control input-small" /></td>
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

                              strTxtExport = strTxtExport & strMedabSel &";"& strMedabSelInit &";"& diet_stdato &";"& diet_sldato &";"& diet_dageIaltTxt &";"& Chr(34) & oRec("diet_namedest") & Chr(34) &";"& Chr(34) & jobnrStrTxt & Chr(34) &";"& jobnrStrTxtNr &";"& jobnrStrTxtProc &";"
                               
                                  if cint(vis_reduktion) = 1 then 
                                  strTxtExport = strTxtExport & makskostTxt &";"& diet_morgenTxt &";"& diet_middagTxt &";"& diet_aftenTxt &";"& diet_restTxtBeregnet &";"& vbcrlf
                                  else
                                  strTxtExport = strTxtExport & makskostTxt &";"& vbcrlf
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

                                  strTxtExport = strTxtExport & strMedabSel &";"& strMedabSelInit &";"& diet_stdato &";"& diet_sldato &";"& diet_dageIaltTxt &";"& Chr(34) & oRec("diet_namedest") & Chr(34) &";"& Chr(34) & jobnrStrTxt & Chr(34) &";"& jobnrStrTxtNr &";"& jobnrStrTxtProc &";"
                                  
                                  if cint(vis_reduktion) = 1 then 
                                  strTxtExport = strTxtExport & makskostTxt &";"& diet_morgenTxt &";"& diet_middagTxt &";"& diet_aftenTxt &";"& diet_restTxt &";"& vbcrlf
                                  else
                                  strTxtExport = strTxtExport & makskostTxt &";"& vbcrlf
                                  end if     
                               
                                  end if

                            if media <> "export" then
                                    %>
                                       <input type="hidden" name="FM_diet_jobid" value="0" />
                                       <input type="hidden" name="FM_diet_konto" value="0" />
                                       <a href="#" onclick="Javascript:window.open('traveldietexp.asp?func=addmoorejobs&trvlid=<%=oRec("diet_id")%>&medarb=<%=oRec("diet_mid")%>', '', 'width=600,height=600,top=100,left=400,resizable=no')" target="_blank" class="small">Tilføj job >></a>
                                   </td>
                                   <!--<td><select name="FM_diet_konto" class="form-control input-small"><%=strKontoOptionsLoop %></select></td>-->
                                   <td style="text-align:right;"><%=diet_dageIaltTxt %></td>
                                   <td style="text-align:right;"><%=makskostTxt %></td>

                                   <%if cint(vis_reduktion) = 1 then %>
                                   <td><input type="text" value="<%=diet_morgenTxt%>" name="FM_diet_morgen" class="form-control input-small" /></td>
                                   <td><input type="text" value="<%=diet_middagTxt%>" name="FM_diet_middag" class="form-control input-small" /></td>
                                   <td><input type="text" value="<%=diet_aftenTxt%>" name="FM_diet_aften" class="form-control input-small" /></td>
                                   <td style="text-align:right;"><%=diet_totalTxt%></td>
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
                   
                       
                                   <td style="text-align:center;"><a href="traveldietexp.asp?menu=tok&func=slet&id=<%=oRec("diet_id")%>&aar=<%=aar%>&medarb=<%=oRec("diet_mid")%>"><span style="color:darkred;" class="fa fa-times"></span></a></td>

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
	                strSQLsatser = "SELECT diet_dayprice, diet_morgenamount, "_
                    &" diet_middagamount, diet_aftenamount "_
                    &" FROM traveldietexp "_
                    &" WHERE diet_dayprice <> 0 ORDER BY diet_id DESC"
	
                    'YEAR(diet_stdato) = '"& aar &"' AND
                    'response.write strSQLsatser
                    'response.Flush
        
                    d = 0
	                diet_dayprice = 0
                    diet_morgenamount = 0
                    diet_middagamount = 0
                    diet_aftenamount = 0


                    oRec.open strSQLsatser, oConn, 3
                    if not oRec.EOF then

                               diet_dayprice = formatnumber(oRec("diet_dayprice"), 2)
                                diet_morgenamount = formatnumber(oRec("diet_morgenamount"), 2)
                                diet_middagamount = formatnumber(oRec("diet_middagamount"), 2)
                                diet_aftenamount = formatnumber(oRec("diet_aftenamount"), 2)

                    %>

                  

                                  <tr>
                     
                                    <th>&nbsp;</th>
                                   <th>&nbsp;</th>
                                 <th>&nbsp;</th>
                                  <th>&nbsp;</th>
                                  <th>&nbsp;</th>
                                 <!-- <th>&nbsp;</th> kontonr -->
                    
                                   <th>
                                       
                                       <%if mainAmountBoxes = "DISABLED" then %>
                                       <input type="text" value="<%=diet_dayprice %>" name="" class="form-control input-small" style="font-size:9px; font-weight:lighter;" <%=mainAmountBoxes %>/>
                                       <input type="hidden" value="<%=diet_dayprice %>" name="FM_diet_dayprice"/>
                                       <%else %>
                                       <input type="text" value="<%=diet_dayprice %>" name="FM_diet_dayprice" class="form-control input-small" style="font-size:9px; font-weight:lighter;" <%=mainAmountBoxes %>/>
                                       <%end if %>
                                   
                                   </th>
                                   <th>&nbsp;</th>
                                     <%if cint(vis_reduktion) = 1 then %>
                                  <th><input type="text" value="<%=diet_morgenamount %>" name="FM_diet_morgenamount<%=mainAmountBoxesName %>" class="form-control input-small" style="font-size:9px; font-weight:lighter;" <%=mainAmountBoxes %> /></th>
                                   <th><input type="text" value="<%=diet_middagamount %>" name="FM_diet_middagamount<%=mainAmountBoxesName %>" class="form-control input-small" style="font-size:9px; font-weight:lighter;" <%=mainAmountBoxes %> /></th>
                                    <th><input type="text" value="<%=diet_aftenamount %>" name="FM_diet_aftenamount<%=mainAmountBoxesName %>" class="form-control input-small" style="font-size:9px; font-weight:lighter;" <%=mainAmountBoxes %> /></th>
                                  <th>&nbsp;</th>
                                   <%else %>
                                   <input type="hidden" value="0" name="FM_diet_morgenamount<%=mainAmountBoxesName %>" />
                                   <input type="hidden" value="0" name="FM_diet_middagamount<%=mainAmountBoxesName %>" />
                                    <input type="hidden" value="0" name="FM_diet_aftenamount<%=mainAmountBoxesName %>" />
                                  
                                <%end if %>

                                 <th>&nbsp;</th>
                                   <th>&nbsp;</th>
                                   <th>&nbsp;</th>
                               </tr>

                                <%if mainAmountBoxes = "DISABLED" then %>
                               <input type="hidden" value="<%=diet_morgenamount %>" name="FM_diet_morgenamount" />
                               <input type="hidden" value="<%=diet_middagamount %>" name="FM_diet_middagamount"  />
                               <input type="hidden" value="<%=diet_aftenamount %>" name="FM_diet_aftenamount" />
                                  <%end if %>


                                </thead>
                           <tbody>


                    <%
                    end if
                    oRec.close

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
                       <td>&nbsp;</td>
                       <td>
                           <input type="hidden" name="FM_diet_id" value="0"/>
                           <input type="hidden" name="FM_diet_medid" value="<%=usemrn%>"/>
                           <input class="form-control input-small" type="text" name="FM_diet_stdato" value="" placeholder="<%=daytimeFormatPlaceholder %>"/></td>
                       <td><input class="form-control input-small" type="text" name="FM_diet_sldato" value="" placeholder="<%=daytimeFormatPlaceholder %>"/></td>
                       <td><input type="text" value="" placeholder="Destination" name="FM_diet_namedest" class="form-control input-small" /></td>
                       <td><!--<select name="FM_diet_jobid" class="form-control input-small"><%=strJobOptions %></select>-->
                           <input type="hidden" name="FM_diet_jobid" value="0" />
                           <input type="hidden" name="FM_diet_konto" value="0" />
                           &nbsp;</td>
                       <!--<td><select name="FM_diet_konto" class="form-control input-small"><%=strKontoOptions %></select></td>-->
                       <td>&nbsp;</td>
                       <td>&nbsp;</td>
                       <%if cint(vis_reduktion) = 1 then %>
                       <td><input type="text" value="" name="FM_diet_morgen" class="form-control input-small" /></td>
                       <td><input type="text" value="" name="FM_diet_middag" class="form-control input-small" /></td>
                       <td><input type="text" value="" name="FM_diet_aften" class="form-control input-small" /></td>
                       <td>&nbsp;</td>
                       <%else %>
                        <input type="hidden" value="" name="FM_diet_morgen" class="form-control input-small" />
                       <input type="hidden" value="" name="FM_diet_middag" class="form-control input-small" />
                       <input type="hidden" value="" name="FM_diet_aften" class="form-control input-small" />
                       <%end if %>
                       
                      
                       <td style="text-align:center;"><input type="checkbox" value="1" name="FM_diet_bilag"/></td>
                       <td>&nbsp;</td>
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
            
            
             <section>
                         <div class="row">
                             <div class="col-lg-10">&nbsp;</div>
                             <div class="col-lg-2">
                            <input type="submit" value="Opdater >>" class="btn btn-sm btn-success pull-right"><br />&nbsp;
                            </div>
                        </div>
                </section>
            
         </form>


         <br /><br />
                
            <section>
                <div class="row">
                     <div class="col-lg-12">
                        <b>Funktioner</b>
                        </div>
                    </div>

                  
                <form action="traveldietexp.asp?media=export&aar=<%=aar%>&medarb=<%=usemrn %>" method="post" target="_blank">
                  
                 <div class="row">
                     <div class="col-lg-12 pad-r30">
                         <input id="Submit6" type="submit" value="Eksport til .csv" class="btn btn-sm" />
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
	
	
	
	            strTxtExportHeader = "Navn;Init;Afrejsedato;Hjemrejse;Dage ialt;Destination;Job/Projekt;Jobnr;Procent %;"
                if cint(vis_reduktion) = 1 then 
                strTxtExportHeader = strTxtExportHeader & "Maks.kost;Morgen;Middag;Aften;Ialt/Rest.;" & vbcrlf
                else
                strTxtExportHeader = strTxtExportHeader & "Ialt;"& vbcrlf
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