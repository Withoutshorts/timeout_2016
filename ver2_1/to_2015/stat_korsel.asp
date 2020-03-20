

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<div class="wrapper">
<div class="content">

    <%
        if len(session("user")) = 0 then	
	    errortype = 5
	    call showError(errortype)
	    response.End
        end if

        call smileyAfslutSettings()

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
	
	    level = session("rettigheder")

        if len(request("medarbSel")) <> 0 then
	        medarbSel = request("medarbSel")
            response.Cookies("to_2015")("korselusemrn") = medarbSel
	    else
            if request.Cookies("to_2015")("korselusemrn") <> "" then
                medarbSel = request.Cookies("to_2015")("korselusemrn")
            else
	            medarbSel = session("mid")
            end if
	    end if
   
        allmedids = ""

        if level = 1 then
            medarbSQLkri = " m.mid <> 0 "
            showalle = 1
        else
            call fTeamleder(session("mid"), 0, 1)
            if strPrgids = "ingen" then
                medarbSQLkri = " m.mid = "& session("mid")
                showalle = 0
            else
                medarbSQLkri = ""
                strSQL = "SELECT medarbejderid FROM progrupperelationer WHERE projektgruppeid IN ("& strPrgids &") GROUP BY medarbejderid"                                       
                oRec.open strSQL, oConn, 3
                a = 0
                while not oRec.EOF
                    if a = 0 then
                        medarbSQLkri = medarbSQLkri & "(m.mid = "& oRec("medarbejderid")
                        allmedids = oRec("medarbejderid")
                    else
                        medarbSQLkri = medarbSQLkri & " OR m.mid = "& oRec("medarbejderid")
                        allmedids = allmedids & "," & oRec("medarbejderid")
                    end if

                a = a + 1
                oRec.movenext
                wend
                oRec.close
                                        
                medarbSQLkri = medarbSQLkri & ")"
                showalle = 1
            end if
        end if 'level = 1


        if level = 1 then
            afregnetActive = ""
            afregnetActiveInt = 1
        else
            afregnetActive = "DISABLED"
            afregnetActiveInt = 0
        end if


        rdir = request("rdir")
		
	    thisfile = "stat_korsel"
	
	    '*** Smiley ***''
        call ersmileyaktiv()
	
	    media = request("media")
	    

        if isdate(request("FM_startdato")) = true then
        strAar = year(request("FM_startdato"))
        strMrd = month(request("FM_startdato"))
        strDag = day(request("FM_startdato"))        
        else
        strAar = year(now)
        strMrd = month(now)
        strDag = "1"
        end if

        strdatoST = strDag &"-"& strMrd &"-"& strAar

        if isdate(request("FM_slutdato")) = true then
            strAar_slut = year(request("FM_slutdato"))
            strMrd_slut = month(request("FM_slutdato"))
            strDag_slut = day(request("FM_slutdato"))            
        else
            strAar_slut = year(now)
            strMrd_slut = month(now)
            strDag_slut = day(dateadd("m",1, now))'day(dateadd("d", -1, (year(now)&"-"&(month(now) + 1)&"-1")))        
        end if

        strdatoSL = strDag_slut &"-"& strMrd_slut &"-"& strAar_slut

        function ArTilDatoKm()
	        %>
	        <%=stat_korsel_txt_051 %> -> <%=stat_korsel_txt_018 %> (<%=formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 2) %>) :
      
       
        
            <% 
            '** Sum år til dato ***'
            strSQLsum = "SELECT SUM(timer) AS kmsum, tfaktim FROM timer "_
            &" WHERE tmnr = "& ltmnr &" AND "_
            &" tdato BETWEEN '"& strAar_slut & "/1/1" &"' AND '"& sqlDatoSlut &"'"_
            &" AND tfaktim = 5 GROUP BY tmnr, tfaktim"
            kmAarDato = 0
        
            'Response.Write strSQLsum
            'response.flush
        
            oRec3.open strSQLsum, oConn, 3
            if not oRec3.EOF then
        
            kmAarDato = oRec3("kmsum")
        
            end if
            oRec3.close
	    
	        Response.Write "<b>"& formatnumber(kmAarDato, 2) & "</b>"	
	    end function

        if media <> "print" then	    
            call menu_2014()

	        sideDivTop = 102
	        sideDivLeft = 90	
	    else
	        sideDivTop = 20
	        sideDivLeft = 20	
	    end if

        if len(request("FM_sumprmed")) <> 0 AND request("FM_sumprmed") <> 0 then
          showTot = 1
          sumChk = "CHECKED"
          else
          sumChk = ""
          showTot = 0
          end if


          select case lto
          case "lm"
          brug_afregnet = 1
          case else
          brug_afregnet = 0
          end select



          if func = "dbred" then
          tids = request("FM_kilometer_id")
          tids_split = split(tids, ", ")

          for i = 0 TO uBOUND(tids_split)

          if len(trim(request("FM_kilometer_afregnet_" & tids_split(i)))) <> 0 then
                    kilometer_afregnet = request("FM_kilometer_afregnet_" & tids_split(i))
                else
                    kilometer_afregnet = 0
                end if

                 if len(trim(request("FM_kiliometer_afrenget_lastvalue_" & tids_split(i)))) <> 0 then
                    kilometer_afregnet_before = request("FM_kiliometer_afrenget_lastvalue_" & tids_split(i))
                else
                    kilometer_afregnet_before = 0
                end if
                
                kilometer_afregnet_dato = year(now) &"-"& month(now) &"-"& day(now)

                if cint(kilometer_afregnet) <> cint(kilometer_afregnet_before) then
                    
                    oConn.execute("UPDATE timer SET afregnet = "& kilometer_afregnet &", afregnetdato = '"& kilometer_afregnet_dato &"' WHERE tid = "& tids_split(i))

                    'response.Write "<br><br><br><br><br> ------------------------------------- TID " & tids_split(i) & " afrenget " & kilometer_afregnet

                end if

            next

            response.Redirect "stat_korsel.asp"
        end if
        %>


    <script src="js/datepicker.js" type="text/javascript"></script>
    <script src="js/statkorsel_jav.js" type="text/javascript"></script>    
    <div class="container">
        <div class="portlet">

            <h3 class="portlet-title"><u><%=stat_korsel_txt_001 %></u></h3>


            <div class="portlet-body">

                <%
                if media <> "print" then
                filerDisplay = ""
                else
                filerDisplay = "none"
                end if
                %>

                <form action="stat_korsel.asp?menu=stat&FM_usedatokri=1&hidemenu=<%=hidemenu%>" method="post" style="display:<%=filerDisplay%>;">                    
                    <div class="well">
                        <h4 class="panel-title-well"><%=medarb_txt_098 %></h4>
                        <div class="row">
                            <div class="col-lg-3"><%=stat_korsel_txt_002 %>:</div>
                            <div class="col-lg-2"><%=stat_korsel_txt_003 %>:</div>
                            <div class="col-lg-2"><%=stat_korsel_txt_004 %>:</div>
                        </div>

                        <div class="row">
                            <div class="col-lg-3">

                                <select name="medarbSel" id="medarbSel" class="form-control input-small">
                                    <%if showalle = 1 then%>
			                        <option value="0"><%=stat_korsel_txt_005 %></option>
			                        <%end if %>

                                    <%
                                    strSQL = "SELECT mnavn, mid, mnr FROM medarbejdere m WHERE m.mid <> 0 AND mansat <> 2 AND mansat <> 3 AND "& medarbSQLkri &" ORDER BY mnavn"
			                        'Response.write strSQL
			                        'Response.flush
			                        oRec.open strSQL, oConn, 3
                                    
                                    while not oRec.EOF 

                                    if cint(medarbSel) = oRec("mid") then
			                         mSEL = "SELECTED"
			                         medarbPrVal = oRec("mnavn") &"&nbsp;("& oRec("mnr")&")"
			                         else
			                         mSEL = ""
			                         end if
                                    %>
                                    <option value="<%=oRec("mid")%>" <%=mSEL%>><%=oRec("mnavn")%>&nbsp;(<%=oRec("mnr")%>)</option>
                                    <%
                                    oRec.movenext
			                        wend
			                        oRec.close
                                    %>

                                </select>
                            </div>

                            <div class="col-lg-2">
                                <div class='input-group date'>
                                    <input type="text" autocomplete="off" class="form-control input-small" name="FM_startdato" value="<%=strdatoST %>" placeholder="dd-mm-yyyy" />
                                    <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                                </div>
                            </div>
                            <div class="col-lg-2">
                                <div class='input-group date'>
                                    <input type="text" autocomplete="off" class="form-control input-small" name="FM_slutdato" value="<%=strdatoSL %>" placeholder="dd-mm-yyyy" />
                                    <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                                </div>
                            </div>
                        </div>

                        <div class="row">
                            <div class="col-lg-6"><input type="checkbox" name="FM_sumprmed" id="FM_sumprmed" value="1" <%=sumChk%>> 
                            <%=stat_korsel_txt_006 &" "& stat_korsel_txt_007 %></div>
                            <div class="col-lg-6"><button type="submit" class="btn btn-secondary btn-sm pull-right"><b><%=stat_korsel_txt_008 %> >></b></button></div>
                        </div>

                    </div>
                    </form>

                    <%if media = "print" then 
                        if medarbSel <> 0 then
                        medarbPrVal = medarbPrVal
                        else
                        medarbPrVal = "Alle"
                        end if
                    %>
                    <h5><%=medarbPrVal & " - " & strdatoST & " - " & strdatoSL  %></h5>
                    <%end if %>
                
                    <%
                    '*** Alle timer, uanset fomr. ***
	                sqlDatoStart = strAar&"/"&strMrd&"/"&strDag
	                sqlDatoSlut = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut
	
	                datoStart = strDag&"/"&strMrd&"/"&strAar
	                datoSlut = strDag_slut&"/"&strMrd_slut&"/"&strAar_slut

	                if medarbSel <> 0 then
                    medarbSQLkri = " tmnr = " & medarbSel
	                else
	                
                    if level <> 1 then
                        medarbSQLkri = " tmnr IN ("& allmedids &")"
                    else
                        medarbSQLkri = " tmnr <> 0 "
                    end if

	                end if
                    %>

                    
                    <%if showTot = 0 then %>
                        <form action="stat_korsel.asp?func=dbred" method="post">

                        <%if brug_afregnet = 1 then %>
                        <div class="row">
                            <div class="col-lg-12">
                                <button type="submit" class="btn btn-success btn-sm pull-right"><b><%=medarb_txt_020 %></b></button>
                            </div>
                        </div>
                        <%end if %>

                        <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                            <thead>
                                <tr>
                                    <th><%=stat_korsel_txt_009 %></th>
                                    <th><%=stat_korsel_txt_002 %></th>
                                    <th><%=stat_korsel_txt_010 %><br /><%=stat_korsel_txt_011 %></th>
                                    <th><%=stat_korsel_txt_012 %></th>
                                    <th><%=stat_korsel_txt_013 %></th>

                                    <th><%=stat_korsel_txt_014 %>
                                    <%if lto <> "lm" then %>
                                        <br />
                                        (<%=stat_korsel_txt_015 %>)
                                    <%end if %>
                                    </th>

                                    <th><%=stat_korsel_txt_016 %>
                                    <%if lto <> "lm" then %>
                                        <br />
                                        (<%=stat_korsel_txt_015 %>)
                                    <%end if %>
                                    </th>

                                    <th><%=stat_korsel_txt_017 %></th>

                                    <%if brug_afregnet = 1 then %>

                                        <th><%=expence_txt_048 %>
                                            <br /> <input type="checkbox" id="afregnAlle" <%=afregnetActive %> />
                                        </th>
                                    <%end if %>
                                </tr>
                            </thead>

                            <tbody>
                                <%
                                strEksportTxt = stat_korsel_txt_018&";"&stat_korsel_txt_002&";"&stat_korsel_txt_019&";"&stat_korsel_txt_011&";"&stat_korsel_txt_012&";"&stat_korsel_txt_014&";"&stat_korsel_txt_016&";"&stat_korsel_txt_017&";"&stat_korsel_txt_013

                                if brug_afregnet = 1 then
                                    strEksportTxt = strEksportTxt & ";"&expence_txt_048
                                end if

                                strEksportTxt = strEksportTxt &";xx99123sy#z"


                                strSQL = "SELECT tid, taktivitetnavn, timer, tfaktim, tjobnavn, tjobnr, tdato, "_
                                &" tknr, tknavn, tmnavn, tmnr, kkundenr, godkendtstatus, timerkom, bopal, "_
                                &" m.mnr, m.init, m.mid, destination, overfort, afregnet, afregnetdato FROM timer "_
                                &" LEFT JOIN kunder K on (kid = tknr) "_
                                &" LEFT JOIN medarbejdere m ON (mid = tmnr) WHERE "& medarbSQLkri &" AND "_
                                &" tdato BETWEEN '"& sqlDatoStart &"' AND '"& sqlDatoSlut &"' AND tfaktim = 5 ORDER BY tmnr, tdato DESC "
   
                                'Response.Write strSQL
                                'Response.flush
   
                                at = 0
                                timertot = 0
                                lwedaynm = ""
                                timerDag = 0
                                ltmnr = 0
                                lastMid = 0
   
                                oRec.open strSQL, oConn, 3
                                while not oRec.EOF
                                    
                                if lastMid <> oRec("mid") then
                                call afsluger(oRec("mid"), sqlDatoStart, sqlDatoSlut)
                                end if

                                select case right(at, 1)
                                case 0,2,4,6,8
                                bgcol = "#ffffff"
                                case else
                                bgcol = "#EFF3ff"
                                end select
                                %>

                                <%if ltmnr <> oRec("tmnr") then %>
                                    <%if at <> 0 then %>
                                        <tr>
                                            <td align=right colspan=8><br /><%=stat_korsel_txt_050 %>: <b><%=formatnumber(kmTot, 2) %></b>
                                            <br />
                                            <%call ArTilDatoKm() %>        
                                            </td>

                                            <%if cint(brug_afregnet) = 1 then %>
                                            <td></td>
                                            <%end if %>
                                        </tr>
                                        <%
                                   kmTot = 0
                                   end if %>

                                    <tr bgcolor="#cccccc">
                                        <td colspan=9 height=30><b><%=oRec("tmnavn") %></b>
                                        <%if len(trim(oRec("init"))) <> 0 then
                                        %>
                                        (<%=oRec("init") %>)
                                        <%
                                        end if %>
                                        </td>
                                    </tr>

                                <%end if %>

                                <input type="hidden" name="FM_kilometer_id" value="<%=oRec("tid") %>" />

                                <tr>
                                    <td style="width:100px;">
                                        <%
                                            select case weekday(oRec("Tdato"), 2)
                                            case 1
                                            strweekdayname = stat_korsel_txt_052 
                                            case 2
                                            strweekdayname = stat_korsel_txt_053
                                            case 3
                                            strweekdayname = stat_korsel_txt_054
                                            case 4
                                            strweekdayname = stat_korsel_txt_055
                                            case 5
                                            strweekdayname = stat_korsel_txt_056
                                            case 6
                                            strweekdayname = stat_korsel_txt_057
                                            case 7
                                            strweekdayname = stat_korsel_txt_058
                                            end select

                                            select case month(oRec("Tdato"))
                                            case 1
                                            strkmonthname = stat_korsel_txt_059
                                            case 2
                                            strkmonthname = stat_korsel_txt_060
                                            case 3
                                            strkmonthname = stat_korsel_txt_061
                                            case 4
                                            strkmonthname = stat_korsel_txt_062
                                            case 5
                                            strkmonthname = stat_korsel_txt_063
                                            case 6
                                            strkmonthname = stat_korsel_txt_064
                                            case 7
                                            strkmonthname = stat_korsel_txt_065
                                            case 8
                                            strkmonthname = stat_korsel_txt_066
                                            case 9
                                            strkmonthname = stat_korsel_txt_067
                                            case 10
                                            strkmonthname = stat_korsel_txt_068
                                            case 11
                                            strkmonthname = stat_korsel_txt_069
                                            case 12
                                            strkmonthname = stat_korsel_txt_070
                                            end select

                                        %>
                                        

                                        <%=strweekdayname %>. <%=day(oRec("Tdato")) &" "& strkmonthname &". "& right(year(oRec("Tdato")), 2) %>

                                    </td>
                                    
                                    <td style="width:140px;"><%=oRec("tmnavn") %>
                                    <%if len(trim(oRec("init"))) <> 0 then
                                    %>
                                    (<%=oRec("init") %>)
                                    <%
                                    end if 
        
                                    if len(trim(oRec("timerkom"))) <> 0 then
                                    timerKom = replace(oRec("timerkom"), stat_korsel_txt_004&":", "<br><b>"&stat_korsel_txt_004&":</b> ")
                                    timerKom = replace(timerKom, stat_korsel_txt_047&":", "<b>"&stat_korsel_txt_047&":</b> ")
                                    timerKom = replace(timerKom, stat_korsel_txt_048&":", "<br><b>"&stat_korsel_txt_048&":</b> ")
                                    else
                                    timerKom = oRec("timerkom")
                                    end if        
                                    %>
                                    </td>

                                    <td style="width:200px;"><b><%=oRec("tknavn") %> (<%=oRec("kkundenr") %>)</b><br />
                                    <%=oRec("tjobnavn") %> (<%=oRec("tjobnr") %>)
                                    </td>

                                    <td><%=oRec("taktivitetnavn") %></td>
                                    <td><%=timerKom %></td>
                                    <td>
                                    <%if oRec("bopal") = 1 then %>
                                    Ja
                                    <%else %>
                                    &nbsp;
                                    <%end if %></td>

                                    <%
                                    strEksportTxt = strEksportTxt & oRec("Tdato") & ";"& oRec("tmnavn") &" ("& oRec("init") &");" & oRec("tknavn") & " ("& oRec("kkundenr") &");"
                                    strEksportTxt = strEksportTxt & oRec("tjobnavn") &" ("& oRec("tjobnr")  &");" & oRec("taktivitetnavn") &";"& oRec("bopal") & ";" 
                                    %>

                                    <td valign=top>    
                                     <%
     
                                     lastDestNavn = ""
     
                                     if lto <> "lm" then
                                         call destAdr(oRec("destination")) %>
     
                                         <%=lastDestNavn%>&nbsp;
                                    <%else %>
                                        <%=oRec("destination") %>
                                    <%end if %>

                                    </td>

                                    <td>
                                        <%
                                        '** er periode godkendt ***'
		                                        tjkDag = oRec("Tdato")
		                                        erugeafsluttet = instr(afslUgerMedab(oRec("tmnr")), "#"&datepart("ww", tjkDag,2,2)&"_"& datepart("yyyy", tjkDag) &"#")
                

                                                strMrd_sm = datepart("m", tjkDag, 2, 2)
                                                strAar_sm = datepart("yyyy", tjkDag, 2, 2)
                                                call thisWeekNo53_fn(tjkDag)
                                                strWeek = thisWeekNo53 'datepart("ww", tjkDag, 2, 2)
                                            
                                                strAar = datepart("yyyy", tjkDag, 2, 2)

                                                if cint(SmiWeekOrMonth) = 0 then
                                                usePeriod = strWeek
                                                useYear = strAar
                                                else
                                                usePeriod = strMrd_sm
                                                useYear = strAar_sm
                                                end if

                
                                                call erugeAfslutte(useYear, usePeriod, oRec("tmnr"), SmiWeekOrMonth, 0, tjkDag)
		        
		                                        'Response.Write "smilaktiv: "& smilaktiv & "<br>"
		                                        'Response.Write "SmiWeekOrMonth: "& SmiWeekOrMonth &" ugeNrAfsluttet: "& ugeNrAfsluttet & " tjkDag: "& tjkDag &"<br>"
		                                        'Response.Write "autolukvdatodato: "& autolukvdatodato & "<br>"
		                                        'Response.Write "tjkDag: "& tjkDag & "<br>"
		                                        'Response.Write "autolukvdato: "& autolukvdato & "<br>"
		                                        'Response.Write "erugeafsluttet:" & erugeafsluttet & "<br>"
		        
		                                         strSQL2 = "SELECT risiko FROM job WHERE jobnr = '"& oRec("Tjobnr") &"'"
						                                oRec2.open strSQL2, oConn, 3 
						                                if not oRec2.EOF then
						                                risiko = oRec2("risiko")
						                                end if
						                                oRec2.close

                                                 call lonKorsel_lukketPer(tjkDag, risiko, oRec("tmnr"))
              
		         
                                               
                                                '*** tjekker om uge er afsluttet / lukket / lønkørsel
                                                call tjkClosedPeriodCriteria(tjkDag, ugeNrAfsluttet, usePeriod, SmiWeekOrMonth, splithr, smilaktiv, autogk, autolukvdato, lonKorsel_lukketIO, ugegodkendt)


                                                if oRec("godkendtstatus") = 1 OR oRec("godkendtstatus") = 3 OR oRec("overfort") = 1 then '3: tentative
                                                ugeerAfsl_og_autogk_smil = 1
                                                else
                                                ugeerAfsl_og_autogk_smil = ugeerAfsl_og_autogk_smil
                                                end if
                
                
                                                if ugeerAfsl_og_autogk_smil = 0 OR (level = 10) then%>
                
                                                <a href="#" onclick="Javascript:window.open('rediger_tastede_dage_2006.asp?id=<%=oRec("tid") %>', '', 'width=650,height=500,resizable=yes,scrollbars=yes')" class=vmenu><%=formatnumber(oRec("timer"), 2) %></a>
	                                        
                                                <%else %>
                                                <%=formatnumber(oRec("timer"), 2) %>
                                                <%end if %>
                                    
                                    </td>

                                    <%if brug_afregnet = 1 then %>

                                        <%
                                        if cint(oRec("afregnet")) = 1 then
                                            kilometer_afrenget = "CHECKED"
                                            afregnetdatospan = "<br> <span style='font-size:9px;'>"& oRec("afregnetdato") &"</span>"
                                        else
                                            kilometer_afrenget = ""
                                            afregnetdatospan = ""
                                        end if

                                        %>
                                        
                                        <input type="hidden" name="FM_kiliometer_afrenget_lastvalue_<%=oRec("tid") %>" value="<%=oRec("afregnet") %>" />
                                        <td style="text-align:center;">
                                            <%if afregnetActiveInt = 1 then %>
                                            <input class="afregnkorsel" <%=afregnetActive %> type="checkbox" value="1" name="FM_kilometer_afregnet_<%=oRec("tid") %>" <%=kilometer_afrenget %> />
                                            <%else %>
                                            <input type="checkbox" <%=kilometer_afrenget %> disabled />
                                            <input type="hidden" value="<%=oRec("afregnet") %>" name="FM_kilometer_afregnet_<%=oRec("tid") %>" />
                                            <%end if %>
                                            <%=afregnetdatospan %>
                                        </td>
                                    <%end if %>
                                </tr>

                                <%
                                call htmlparseCSV(timerKom)
   
                                if lto = "lm" then
                                    lastDestnavn = oRec("destination")
                                end if

                                strEksportTxt = strEksportTxt & lastDestnavn & ";" & formatnumber(oRec("timer"), 2) & ";" & htmlparseCSVtxt & ";"

                                if brug_afregnet = 1 then
                                    strEksportTxt = strEksportTxt & oRec("afregnet") & ";"
                                end if

                                strEksportTxt = strEksportTxt & "xx99123sy#z"

                                lwedaynm = weekdayname(weekday(oRec("Tdato")))
                                ltmnr = oRec("tmnr")
   
                                kmTot = kmTot + oRec("timer")
    
                                lastMid = oRec("mid") 
    
                                at = at + 1
                                oRec.movenext
                                wend
                                oRec.close
                                %>

                                <%if at <> 0 then %>
                                <tr>
                                    <td align=right colspan=8><br />Total i periode: <b><%=formatnumber(kmTot, 2) %></b>
                                     <br />
                                    <%call ArTilDatoKm() %></td>

                                    <%if brug_afregnet = 1 then %>
                                    <td></td>
                                    <%end if %>
                               </tr>
                               <tr> 
                               <%else %>
                                <tr>
                                    <td colspan=9><%=stat_korsel_txt_020 %></td>
                                </tr>
                               <%end if %>

                            </tbody>

                        </table>
                        </form>

                    <%else 'show tot %>


                    <%
                    call akttyper2009(2)
    
                    dim overligger, rest
                    dim lastEndDato, dest_idnum
                    dim dest_id, dest_days, days_all
                    redim dest_id(400,1), dest_days(720,1), days_all(720,1)
                    redim lastEndDato(400,1), dest_idnum(400,1)
                    redim overligger(400,1), rest(400,m)
    
                    sqlDatoSlut = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut
                    datoSlut = strDag_slut&"/"&strMrd_slut&"/"&strAar_slut
                    datoStart = dateadd("m", -24, strDag_slut&"/"&strMrd_slut&"/"&strAar_slut)
                    sqlDatoStart = year(datoStart) &"/"& month(datoStart) &"/"& day(datoStart)
                    %>



                    <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                        <thead>
                            <tr>
                                <th style="width:120px;"><b><%=stat_korsel_txt_002 %></b></th>
	                            <th style="width:180px;"><b><%=stat_korsel_txt_021 %></b></th>
	                            <th style="width:150px;"><b><%=stat_korsel_txt_022 %></b></th>
	                            <th style="padding-right:20px;"><b><%=stat_korsel_txt_023 %><br /> <%=stat_korsel_txt_024 %></b></th>
	                            <th style="padding-right:20px;"><b><%=stat_korsel_txt_023 %> <br /><%=stat_korsel_txt_025 %></b></th>
	                            <th style="padding-right:20px;"><b><%=stat_korsel_txt_026 %></b><br />
	                            (<%=stat_korsel_txt_027 %><br /><%=stat_korsel_txt_028 %><br /> <%=stat_korsel_txt_029 %>)
	                            </th>
	                           <th style="padding-right:20px;"><b><%=stat_korsel_txt_030 %></b></th>
                            </tr>
                        </thead>
                        <%
                            'strEksportTxt = "Medarbejder;Mnr;Init;Destination;Antal arb. dage;Antal dage m. godtgørelse;Akkumuleret pr. dest.;Rest dage;xx99123sy#z"  
                            strEksportTxt = stat_korsel_txt_002&";"&stat_korsel_txt_031&";"&stat_korsel_txt_032&";"&stat_korsel_txt_016&";"&stat_korsel_txt_033&";"&stat_korsel_txt_034&";"&stat_korsel_txt_035&";"&stat_korsel_txt_036&";xx99123sy#z"
                        %>

                        <tbody>
                            <%
                            strSQL = "SELECT tknr, timer, tdato, tmnavn, tmnr, m.init, m.mnr, destination, tfaktim, bopal FROM timer AS t "_
                            &" LEFT JOIN medarbejdere AS m ON (m.mid = t.tmnr)"_
                            &" WHERE "& medarbSQLkri &" AND tdato BETWEEN '"& sqlDatoStart &"' AND '"& sqlDatoSlut &"' AND (tfaktim = 5 AND bopal = 1) ORDER BY tmnr, tdato, destination"  '(("& aty_sql_realhours &") AND tfaktim <> 5) OR 
    
                            'Response.Write strSQL
                            'Response.Flush
    
                            lastMed = 0
                            'lastDato = ""
                            lastDest = ""
                            lastTfaktim = 0
                            lastBopal = 0
                            lastDato = "2002-1-1"
                            'lastEndDatoPrev = lastDato
    
                            x = 0
                            y = 0
                            z = 0
                            m = 0
                            destIds = 0
    
                            isDestiWrt = ""
                            antalDageMSkatteFri = 0
    
                            antalDageDiff = 0
    
                            oRec.open strSQL, oConn, 3
                            while not oRec.EOF 

                                '** Desiutnation angivet eller ej **'
                                if len(trim(oRec("destination"))) <> 0 then
                                thisDest = trim(oRec("destination"))
                                else
                                thisDest = "k_0"
                                end if

                            %>

                            <%if z = 0 then %>
                                <tr>
                                    <td colspan=7>
                                    <b><%=oRec("tmnavn")%></b> (<%=oRec("mnr") %>)
       
                                    <%if len(trim(oRec("init"))) <> 0 then %>
                                    - <%=oRec("init") %>
                                    <%end if %>       
                                    </td>
                                </tr>
                            <%
                                perStDato = oRec("tdato")
                            end if 
                            %>


                            <%
                            if z <> 0 then
                                '** tjekker seneste indtasting, er det mere end 60 dage siden??
                                sidsteIndtast = dateDiff("d",LastDato, oRec("tdato"),2,2)
                                'Response.Write "sidsteIndtast" & sidsteIndtast
    
                                if lastMed <> oRec("tmnr") OR _
                                    (lastDest <> thisDest AND ((cint(lastBopal) = 1 AND cint(lastTfaktim) = 5))) OR sidsteIndtast >= 60 then 
            
            
                                     call destAdr(lastDest)

                                     select case right(y, 1)
                                     case 0,2,4,6,8
                                     bgcolor = "#eff3FF"
                                     case else
                                     bgcolor = "#ffffff"
                                     end select
            
                                     overliggerThis = 60
                                     antalDageDiff = 0
             
                                     tjkStDato = "2002-01-02"
                                     tjkSlDato = "2002-01-03"
             
                                     call beregnDage(m,x,lastDest,overliggerThis)
                                     call medarbDestLinie(m)
             
                                      '** Ny perstDato ***'
                                     perStDato = oRec("tdato")
                                     'korselsRecFundetpaDato = 1
                       
                                     y = y + 1
                                end if 'medarbid  / dest


                                '** Navn overskrift ***'
                                if lastMed <> oRec("tmnr") then
    
                                    call medarbDageiPer()
    
                                    %>
                                    <tr bgcolor="#8caae6">
                                        <td colspan=7 height=20>
                                        <b><%=oRec("tmnavn")%></b> (<%=oRec("mnr") %>)
        
                                        <%if len(trim(oRec("init"))) <> 0 then %>
                                        - <%=oRec("init") %>
                                        <%end if %>
        
                                        </td>
                                    </tr>
                                    <%
    
                                    antalArbIalt = 0
                                end if 


                                    '** nulstiller ***'
                                    if lastMed <> oRec("tmnr") then
                                        m = m + 1
                                        redim overligger(400,m), dest_idnum(400,m), rest(400,m)
                                        redim lastEndDato(400,m), dest_id(400,m), dest_days(720,m), days_all(720,m)
                                        isDestiWrt = ""
                                        destIds = 0
                                        x = 0
                                        y = 0
                                    end if


                                end if 'z <> 0


                                '** Er destination brugt før ellerer den ny? ***'    
                                if instr(isDestiWrt, ",#"& trim(thisDest) &"#") = 0 then
                                
                                    dest_id(y,m) = thisDest
               
                
                                    'Response.Write "<br>kid: "& dest_id(y,m) &" y: "& y 
                                    isDestiWrt = isDestiWrt & ",#"&trim(thisDest)&"#" 
                                    dest_days(y,m) = 0
                                    days_all(y,m) = 0
                                    dest_idnum(y,m) = destIds
                                    overligger(y,m) = 60
                                    destIds = destIds + 1 
                                    overliggerThis = overligger(y,m)
                                    rest(y,m) = overligger(y,m)
                
                                    first = 1 
                                    else
                                    first = 0 
               
                                end if


                                lastDest = thisDest
                                lastBopal = oRec("bopal")
                                lastTfaktim = oRec("tfaktim")
                                lastmedNavn = oRec("tmnavn")
                                lastMed = oRec("tmnr")
                                lastInit = oRec("init")
                                lastDato = oRec("tdato")
                                lastMnr = oRec("mnr")
                                'lastKid = oRec("tknr")


                                x = x + 1
                                z = z + 1
    
    
                                oRec.movenext
                                wend
                                oRec.close



                                if z <> 0 then
    
                                select case right(y, 1)
                                case 0,2,4,6,8
                                bgcolor = "#eff3FF"
                                case else
                                bgcolor = "#ffffff"
                                end select
   
    
                                call destAdr(lastDest)
                                call beregnDage(m,x,lastDest,overliggerThis)
                                call medarbDestLinie(m)
                                call medarbDageiPer()
    
                                end if


                                %>
                        </tbody>

                    </table>

                    <%end if 'showtot %>

                    <%

                    public tjkStDato, tjkSlDato, antalDageMSkatteFri, lastRest, antalDageDiff, days_All_thisDest 
                    function beregnDage(m,x,lastDest,overliggerThis)
	
	                     for d = 0 to UBOUND(dest_id)
        
                                        if dest_id(d,m) = lastDest then
                    
                                        '** Beregner dage ***'
                                         antalDageMSkatteFri = x 'dest_days(d,m)
                             
                                                    if dest_days(d,m) < overliggerThis then
                                    
                                                        days_All_thisDest = days_all(d,m) + antalDageMSkatteFri
                                                        
                                                    else
                                    
                                                        days_All_thisDest = antalDageMSkatteFri
                                    
                                                    end if    
                    
                    
                                                        '** Er vi stadigvæk under overligger ***'
                                                        if days_All_thisDest < overliggerThis then
                                                        days_All_thisDest = days_All_thisDest
                                                        antalDageMSkatteFri = antalDageMSkatteFri
                                                        else
                                                        'Response.Write "her" & overliggerThis & "<br>"
                                                        days_All_thisDest_diff = (days_All_thisDest - overliggerThis)
                                                        days_All_thisDest = days_All_thisDest - days_All_thisDest_diff
                                                        antalDageMSkatteFri = x - days_All_thisDest_diff 
                                                        end if
                    
                    
                                        'dest_days(d,m) = 0
                                        days_all(d,m) = days_all(d,m) + antalDageMSkatteFri 
                    
                    
                                        rest(d,m) = rest(d,m) - antalDageMSkatteFri
                                        lastRest = rest(d,m)
                    
                                        antalDageDiff = dateDiff("d",lastEndDato(dest_idnum(d,m),m),perStDato,2,2)
                                        lastEndDuse = lastEndDato(dest_idnum(d,m),m)
                    
                                        tjkStDato = year(lastEndDuse) &"-"& month(lastEndDuse) &"-"& day(lastEndDuse)
                                        tjkSlDato = year(perStDato) &"-"& month(perStDato) &"-"& day(perStDato)
                    
                                        lastEndDato(dest_idnum(d,m),m) = lastDato
                    
                                        end if
                                 next
	        
	            
	            
	                                 if antalDageDiff >= 60 then
                    
                    
                    
                                        '** Finder antal dage med registreringer siden sidste (skal der nulstilles?)
                                        strSQL2 = "SELECT tid AS antal FROM timer AS t "_
                                        &" WHERE tmnr = "& lastMed &" AND tdato BETWEEN '"& tjkStDato &"' AND '"& tjkSlDato &"' AND "_
                                        &" (("& aty_sql_realhours &") OR (tfaktim = 5 AND destination <> '"& lastDest &"'))" _
                                        &" GROUP BY tdato ORDER BY tdato"  '(("& aty_sql_realhours &") AND tfaktim <> 5) OR 
                    
                                        'Response.Write strSQL2 & "<br><br>"
                                        'Response.Flush
                                        antalArbDage = 0
                                        oRec2.open strSQL2, oConn, 3
                                        while not oRec2.EOF 
                    
                                        antalArbDage = antalArbDage + 1 'oRec2("antal")
                    
                                        oRec2.movenext
                                        wend 
                                        oRec2.close
                    
                    
                                             'if antalArbDage > 2 then
                                             if antalArbDage >= 60 AND lastEndDuse <> "" AND (cint(first) <> 1) then
                                             %>
                                             <tr bgcolor="#FFFFe1">
                                                <td colspan=7><br />
                                                <%=stat_korsel_txt_037 %>: <b><%=lastDestNavn %></b><br />
                                                <%=stat_korsel_txt_038 %><br /><br />
                           
                                                <%=stat_korsel_txt_039 %><br />
                                                <%=stat_korsel_txt_040 %>
                                                <br />
                                                <%=stat_korsel_txt_041 %>: <b><%=lastEndDuse %> (<%=stat_korsel_txt_042 %>: <%=antalDageDiff %> / <%=antalArbDage %> <%=stat_korsel_txt_043 %>)</b>
                                                </td>
                                             </tr>
                                             <%
                         
                                                for d = 0 to UBOUND(dest_id)
        
                                                    if dest_id(d,m) = lastDest then
                                                    '** tilføjer nye 60 dage da der har været 60 arb.dage imellem.
                                                    overligger(d,m) = 60
                                                    rest(d,m) = overligger(d,m) - antalDageMSkatteFri 
                                                    overliggerThis = overligger(d,m) 
                                                    lastRest = rest(d,m)
                                                    'nulstilling = 1
                                                    end if
                            
                                                next
                         
                                             end if
             
             
                                 end if
	        
	            
	        
	                end function




                    public antalArbIalt
                    function medarbDestLinie(m)

                        %>
                        <tr bgcolor="<%=bgcolor %>">
                        <td style="border-bottom:1px #cccccc solid;">
    
   
    
                        <%=lastmedNavn &" ("& lastMnr &") "%>
        
                            <%if len(trim(lastInit)) <> 0 then %>
                            - <%=lastInit%>
                            <%end if %>
        
       
                        </td>
    
                        <%          
                        strEksportTxt = strEksportTxt & lastmedNavn &";"& lastMnr &";"& lastInit &";" 
    
   
               
    
                
               
               
                        %>
    
                        <td style="border-bottom:1px #cccccc solid; white-space:nowrap;"><%=lastDestNavnTxt%></td>
                        <td style="border-bottom:1px #cccccc solid; white-space:nowrap;"><%=formatdatetime(perStDato, 1) %> - <%=formatdatetime(LastDato, 1) %></td>
    
                        <td align=right style="border-bottom:1px #cccccc solid; padding-right:20px;"><%=formatnumber(x, 2) %></td>
                        <td align=right style="border-bottom:1px #cccccc solid; padding-right:20px;"><%=formatnumber(antalDageMSkatteFri, 2) %></td>
                        <td align=right style="border-bottom:1px #cccccc solid; padding-right:20px;"><b><%=formatnumber(days_All_thisDest, 2) %></b></td>
                        <td align=right style="border-bottom:1px #cccccc solid; padding-right:20px;"><%=formatnumber(lastRest, 2) %></td>
    
                        <%
                        strEksportTxt = strEksportTxt & lastDestNavn &";"& formatnumber(x, 2) &";"& formatnumber(antalDageMSkatteFri, 2) &";"& formatnumber(days_All_thisDest, 2)  &";"& formatnumber(lastRest, 2) &";xx99123sy#z" 
                            %>
                        </tr>
    
                        <%
                        lastDestNavn = ""
                        antalArbIalt = antalArbIalt + x 
                        'antalDageMSkatteFri = 0
                        isDestiWrt = isDestiWrt & ",#"&lastDest&"#" 
                        'y = y + 1
                        x = 0
    
                        end function



                        public lastDestNavn, lastDestNavnTxt
                        function destAdr(dest)
                        if len(trim(dest)) <> 0 then
                        
                        
                            '** Er der kørt til en kontaktperson / filial eller til hovedkontakten.
                            '** KP = kontkapers
                            if left(dest,2) <> "kp" then
                            
                                if len(trim(dest)) > 2 then
                                lastDestId = 0
                                len_lastDest = len(dest)
                                lastDestId = right(dest, len_lastDest - 2)
                                else
                                lastDestId = 0
                                end if
                            
                                if len(trim(lastDestId)) = 0 OR lastDestId = "st" then
                                lastDestId = 0
                                end if

                                lastDestNavn = "-"
                                strSQLk = "SELECT kkundenavn FROM kunder WHERE kid = "& lastDestId
                            
                                'Response.Write dest & " -sql: "& strSQLK & "<br>"
                                'Response.flush
                            
                                oRec2.open strSQlk, oConn, 3
                            
                                if not oRec2.EOF then
                            
                                lastDestNavn = oRec2("kkundenavn")
                            
                                end if
                                oRec2.close
                        
                            else
                        
                                lastDestId = 0
                                len_lastDest = len(dest)
                                lastDestId = right(dest, len_lastDest - 3)
                            
                            
                                lastDestNavn = "-"
                                strSQLk = "SELECT navn FROM kontaktpers WHERE id = "& lastDestId
                            
                                'Response.write "kp tur: "& strSQLk & "<br>"
                                'Response.flush
                                oRec2.open strSQlk, oConn, 3
                            
                                if not oRec2.EOF then
                            
                                lastDestNavn = oRec2("navn")
                            
                                end if
                                oRec2.close
                            
                                end if 
                        
                            end if
    
    
    
                        if len(trim(lastDestNavn)) <> 0 then
                        lastDestNavnTxt = "<b>"&lastDestNavn&"</b>"
                        lastDestNavn = lastDestNavn
                        else
                        lastDestNavnTxt = "<font class=""lillegray"">"&stat_korsel_txt_049&"</font>"
                        lastDestNavn = stat_korsel_txt_049
                        end if
    
                    end function


                    function medarbDageiPer()
                    %>
                    <tr>
                        <td></td>
                        <td></td>
                        <td></td>
                        <td align=right style="padding-right:20px;"><b><%=formatnumber(antalArbIalt) %></b></td>
                        <td></td>
                        <td></td>
                        <td></td>
                        <td></td>
                    </tr>
                    <%   
                    end function

                        if media = "export" then 
    

                        call TimeOutVersion()
   
	                    ekspTxt = replace(strEksportTxt, "xx99123sy#z", vbcrlf)
	
	                    filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	                    filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
				                    Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				                    if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\stat_korsel.asp" then
					                    Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\stat_korsel_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					                    Set objNewFile = nothing
					                    Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\stat_korsel_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				                    else
					                    Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\log\data\stat_korsel_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					                    Set objNewFile = nothing
					                    Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\log\data\stat_korsel_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				                    end if
				
				
				
				                    file = "stat_korsel_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
				
				
				                    '**** Eksport fil, kolonne overskrifter ***
			                        'objF.writeLine("Periode afgrænsning: "& datointerval & vbcrlf)
				                    'objF.WriteLine(strOskrifter & chr(013))
				                    objF.WriteLine(ekspTxt)
				
				
				
				
	
	                    Response.redirect "../inc/log/data/"& file &""	

                        end if
                  %>




                  <%if media <> "print" then %>


                    <br />

                    <div class="row">
                        <div class="col-lg-12"><b><%=stat_korsel_txt_044 %></b></div>
                    </div>

                    <div class="row">
                        <div class="col-lg-12">
                            <a href="#" onclick="Javascript:window.open('stat_korsel.asp?media=export&FM_sumprmed=<%=showtot %>&medarbSel=<%=medarbSel %>&FM_startdato=<%=strdatoST %>&FM_slutdato=<%=strdatoSL %>', '', 'width=500,height=150,resizable=no,scrollbars=no')" class=rmenu><%=stat_korsel_txt_045 %></a>
                        </div>
                    </div>

                    <div class="row">
                        <div class="col-lg-12"><a href="stat_korsel.asp?media=print&FM_sumprmed=<%=showtot %>&medarbSel=<%=medarbSel %>&FM_startdato=<%=strdatoST %>&FM_slutdato=<%=strdatoSL %>" target="_blank" class=rmenu><%=stat_korsel_txt_046 %></a></div>
                    </div>

                  <%else %>

                    <%Response.Write("<script language=""JavaScript"">window.print();</script>") %>

                  <%end if %>

            </div>
        </div>
    </div>







</div>
</div>

    

<!--#include file="../inc/regular/footer_inc.asp"-->