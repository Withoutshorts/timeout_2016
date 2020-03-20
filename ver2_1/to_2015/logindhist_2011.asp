
<%thisfile = "logindhist_2018" %>



<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/regular/top_menu_mobile.asp"-->




<%
    
    if len(session("user")) = 0 then

	errortype = 5
	call showError(errortype)
    response.End
    end if

    function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
	end function
	
	func = request("func")
	id = request("id")
	thisfile = "logindhist_2011.asp"
    rdir = "lgnhist_2018"
	media = request("media")

    'usemrn = request("usemrn")
    if len(trim(request("usemrn"))) <> 0 then
    usemrn = request("usemrn")
    else
    usemrn = session("mid")
    end if

    level = session("rettigheder")
 

    

    if len(trim(request("medarbsel_form"))) <> 0 then

     

            if len(trim(request("FM_visAlleMedarb"))) <> 0 then
            visAlleMedarbCHK = "CHECKED"
            visAlleMedarb = 1
            else
            visAlleMedarbCHK = ""
            visAlleMedarb = 0
            end if

    
            if len(trim(request("FM_visAlleMedarb_pas"))) <> 0 then
            visAlleMedarb_pasCHK = "CHECKED"
            visAlleMedarb_pas = 1
            else
            visAlleMedarb_pasCHK = ""
            visAlleMedarb_pas = 0
            end if

    else


            if request.cookies("tsa")("visAlleMedarb") = "1" then
            visAlleMedarbCHK = "CHECKED"
            visAlleMedarb = 1
            else
            visAlleMedarbCHK = ""
            visAlleMedarb = 0
            'usemrn = session("mid")
            end if

                    
            if request.cookies("tsa")("visAlleMedarb_pas") = "1" then
            visAlleMedarb_pasCHK = "CHECKED"
            visAlleMedarb_pas = 1
            else
            visAlleMedarb_pasCHK = ""
            visAlleMedarb_pas = 0
            end if


    end if
    response.cookies("tsa")("visAlleMedarb") = visAlleMedarb
    response.cookies("tsa")("visAlleMedarb_pas") = visAlleMedarb_pas

    call medarb_teamlederfor


    if len(trim(request("nomenu"))) <> 0 AND request("nomenu") <> "0" then
    nomenu = 1
    else
    nomenu = 0
    end if

   

    varTjDatoUS_man = request("varTjDatoUS_man")
    varTjDatoUS_son = dateAdd("d", 6, varTjDatoUS_man)
    'varTjDatoUS_son = request("varTjDatoUS_son")

   

    lnk = "&usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&varTjDatoUS_son="&varTjDatoUS_son&"&rdir="&rdir&"&nomenu="&nomenu


    'Response.Write varTjDatoUS_man
    'Response.end

    datoMan = day(varTjDatoUS_man) &"/"& month(varTjDatoUS_man) &"/"& year(varTjDatoUS_man)
    datoSon = day(varTjDatoUS_son) &"/"& month(varTjDatoUS_son) &"/"& year(varTjDatoUS_son)

    next_varTjDatoUS_man = dateadd("d", 7, varTjDatoUS_man)
	next_varTjDatoUS_son = dateadd("d", 7, varTjDatoUS_son)
    next_varTjDatoUS_man = year(next_varTjDatoUS_man) &"/"& month(next_varTjDatoUS_man) &"/"& day(next_varTjDatoUS_man)
	next_varTjDatoUS_son = year(next_varTjDatoUS_son) &"/"& month(next_varTjDatoUS_son) &"/"& day(next_varTjDatoUS_son)


    prev_varTjDatoUS_man = dateadd("d", -7, varTjDatoUS_man)
	prev_varTjDatoUS_son = dateadd("d", -7, varTjDatoUS_son)
    prev_varTjDatoUS_man = year(prev_varTjDatoUS_man) &"/"& month(prev_varTjDatoUS_man) &"/"& day(prev_varTjDatoUS_man)
	prev_varTjDatoUS_son = year(prev_varTjDatoUS_son) &"/"& month(prev_varTjDatoUS_son) &"/"& day(prev_varTjDatoUS_son)


	'Response.Write "media " & media
	
    d_end = 6
	
	'*** Sætter lokal dato/kr format. *****
	Session.LCID = 1030
	
	
    select case func 
	case "-"
	
	case "xxxopdaterlist"

    Response.write Request("logindhist_id") & "<br>"
    Response.write Request("logindhist_dato") & "<br>"
    Response.Write Request("logindhist_sttid")  & "<br>"
    Response.Write Request("logindhist_sltid")  & "<br>"

    ids = split(Request("logindhist_id"), ", ")
    datoers = split(Request("logindhist_dato"), ", ")
    stTiders = split(Request("logindhist_sttid"), ", ")
    slTiders = split(Request("logindhist_sltid"), ", ")

    for i = 0 TO UBOUND(ids)

    '** start tid **'
    dato_tidSt = datoers(i) &" "& stTiders(i)
    if len(trim(stTiders(i))) <> 0 then
        call logindStatus(usemrn, 1, 2, dato_tidSt)
    end if

    Response.Flush
    
     '** slut tid **'
    dato_tidSl = datoers(i) &" "& slTiders(i)
    if len(trim(slTiders(i))) <> 0 then
        call logindStatus(usemrn, 1, 2, dato_tidSl)
    end if
    
    next


    Response.end
	

    case else
       

    call akttyper2009(2)
	
	
	
	if media <> "export" then
	
	
	if media <> "print" then
	%>

	

     <%call browsertype()
    
     %>

	
    <SCRIPT language=javascript src="../timereg/inc/smiley_jav.js"></script>
    <SCRIPT language=javascript src="../timereg/inc/logindhist_2011_jav.js"></script>

      

 
  

  <%
      if nomenu <> 1 then

	 
        if (browstype_client <> "ip") then

         call menu_2014() 

        else

            ddDato_ugedag = day(now) &"/"& month(now) &"/"& year(now)
            ddDato_ugedag_w = datepart("w", ddDato_ugedag, 2, 2)
            
            varTjDatoUS_man_tt = dateAdd("d", -(ddDato_ugedag_w-1), ddDato_ugedag)
        
           call mobile_header 


        end if


      end if

    
       
    else 
	
	'leftPos = 0
	'topPos = 0
	
	%>
	
   <%end if%>
	
	<%end if%>
	

	<%
        if media = "print" then
            mnavn = ""
            strSQLmedarb = "SELECT mnavn FROM medarbejdere WHERE mid = "& usemrn
            oRec2.open strSQLmedarb, oConn, 3
            if not oRec2.EOF then
                mnavn = oRec2("mnavn")
            end if
            oRec2.close

            printTxt = " - " & mnavn

        else
            printTxt = ""
        end if
    %>
	
	<!-------------------------------Sideindhold------------------------------------->
       <%if (browstype_client <> "ip") then %>
        <div class="wrapper">
            <div class="content">   
     <%end if %>


        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u>Komme / Gå tid <%=printTxt %></u></h3>


                <!-- link til timereg og ugeseddel skal ind -->


                <div class="portlet-body">


                     <%select case lto 
                    case "bf"
            
                    case else%>

                    <%
                            
                      
                            select case lto 
                            case "ddc"
                            wdth = 140
                             lft = 960 
                            case else
                            wdth = 240
                            lft = 860 
                            end select
                        
                            if cint(stempelurOn) = 1 then 
                                wdth = wdth + 60
                                lft = lft - 60
                          
                            end if

                                if cint(vis_favorit) = 1 then
                                wdth = wdth + 60
                                lft = lft - 60
                                end if

                    %>
          
           
                        <div style="position:relative; background-color:#ffffff; border:1px #cccccc solid; border-bottom:0; padding:4px; width:<%=wdth%>px; top:-60px; left:<%=lft%>px; z-index:0; font-size:11px;">
           
                              <%select case lto
                                    case "ddc", "cflow", "foa", "care", "kongeaa"
                                    case else
                                    %>
                                    <a href="../timereg/<%=lnkTimeregside%>" class="vmenu"><%=replace(tsa_txt_031, " ", "")%></a>
                                    &nbsp;&nbsp;|&nbsp;&nbsp;
                                    <%end select %>
                    
                             <%  
                            select case lto
                            case "cflow"
                            call medariprogrpFn(session("mid"))

                                if instr(medariprogrpTxt, "#14#") <> 0 OR instr(medariprogrpTxt, "#16#") <> 0 OR instr(medariprogrpTxt, "#3#") <> 0 OR instr(medariprogrpTxt, "#19#") <> 0 OR level = 1 then 
                                %>
                                 <a href="<%=lnkUgeseddel%>" class="vmenu"><%=tsa_txt_337 %></a>&nbsp;&nbsp;|&nbsp;&nbsp;
                                <%
                                end if

                            case else
                            %>
                            <a href="<%=lnkUgeseddel%>" class="vmenu"><%=tsa_txt_337 %></a>&nbsp;&nbsp;|&nbsp;&nbsp;
                            <%end select %>

                           
                            
                            
                       <%if cint(stempelurOn) = 1 then
                
               
                            select case lto
                            case "cflow"
                            call medariprogrpFn(session("mid"))

                                if instr(medariprogrpTxt, "#14#") <> 0 OR instr(medariprogrpTxt, "#16#") <> 0 OR instr(medariprogrpTxt, "#3#") <> 0 OR instr(medariprogrpTxt, "#19#") <> 0 OR level = 1 then 
                                %>
                                <a href="<%=lnkLogind%>" class="vmenu" style="background-color:azure"><%=tsa_txt_340 %></a>&nbsp;&nbsp;|&nbsp;&nbsp;
                                <%
                                end if

                            case else
                                    %>

                        <a href="<%=lnkLogind%>" class="vmenu" style="background-color:azure"><%=tsa_txt_340 %></a>&nbsp;&nbsp;|&nbsp;&nbsp;
                        <%end select %>
               
            
                        <%end if
            
                         
                        if cint(vis_favorit) = 1 then
                    
                            select case lto
                            case "cflow"
                            call medariprogrpFn(session("mid"))

                                if instr(medariprogrpTxt, "#14#") <> 0 OR instr(medariprogrpTxt, "#16#") <> 0 OR instr(medariprogrpTxt, "#3#") <> 0 OR instr(medariprogrpTxt, "#19#") <> 0 OR instr(medariprogrpTxt, "#21#") <> 0 then

                               %>
                                  <a href="<%=lnkFavorit%>"><%=favorit_txt_001 %></a>
                                <%
                                end if
                            case else
                            %>
                            <a href="<%=lnkFavorit%>"><%=favorit_txt_001 %></a>
                            <%end select %>

                        <%end if %>
                          

                            <%if lto = "foa" OR lto = "care" then %>
                            &nbsp;&nbsp;|&nbsp;&nbsp;
                            <a href="../timereg/<%=lnkAfstem%>"><%=afstem_txt_011 %></a>
                            <%end if %>
           
                        </div>



                    <%
                    end select



                    call erTeamLederForvlgtMedarb(level, usemrn, session("mid"))
                        
	                call ersmileyaktiv() %>
                   

                  <!--  <div class="well">

                        <div class="row">
                        <div class="col-lg-12">
                            <h4 class="panel-title-well">Søgefilter</h4>
                        </div>
                        </div>  -->

                    <%
                    if media <> "print" AND len(trim(strSQLmids)) > 0 AND browstype_client <> "ip" then 'Hvis man er level 1 eller teamleder vil len(trim(strSQLmids)) ALTID VÆRE > 16
                      %>
    
	                <form id="filterkri" method="post" action="logindhist_2011.asp">
                        <input type="hidden" name="FM_sesMid" id="FM_sesMid" value="<%=session("mid") %>">
                        <input type="hidden" name="medarbsel_form" id="medarbsel_form" value="1">
                        <input type="hidden" name="nomenu" id="nomenu" value="<%=nomenu %>">
                        <input type="hidden" name="varTjDatoUS_man" id="varTjDatoUS_man" value="<%=varTjDatoUS_man %>">
        
                        <div class="row">
                            <div class="col-lg-3">
                                <% 
				                call medarb_vaelgandre
                                %>
                            </div>
                            <div class="col-lg-4"><input type="CHECKBOX" name="FM_visallemedarb" id="FM_visallemedarb" value="1" <%=visAlleMedarbCHK %> /> <%=tsa_txt_388 %> (<%=tsa_txt_357 %>)
                            &nbsp
                            <input type="CHECKBOX" name="FM_visallemedarb_pas" id="FM_visallemedarb_pas" value="1" <%=visAlleMedarb_pasCHK %> onclick="submit();" /> <%=medarb_txt_031%></div>
                        </div>

                    </form>

	                <%end if 'media
                    %>

                  <!--  </div> -->

                    <%if (browstype_client <> "ip") then %>

                    <br />&nbsp;
                    <div class="row">
                    
                    
                    <div class="col-lg-6">
                    <!-- Afslut uge - skal nok laves om -->
                    <table>
                        <tr><td colspan="2" style="padding-top:10px;">
                        <%
                            call smileyAfslutSettings()
                            'Skal det være mandag for måned og søndag for uge??
                            select case cint(SmiWeekOrMonth) 
                            case 0 
                            call thisWeekNo53_fn(tjkDato)
                            periodeTxt = tsa_txt_005 & " "& thisWeekNo53 'datepart("ww", tjkDato, 2 ,2)
                            weekMonthDate = datepart("ww", varTjDatoUS_son,2,2)
                            case 1
                            periodeTxt = monthname(datepart("m", tjkDato, 2 ,2))
                            weekMonthDate = datepart("m", varTjDatoUS_son,2,2)
                            case 2
                            weekMonthDate = formatdatetime(now, 2) ' SKAL DET ALTID VÆRE DD?
                            weekMonthDate = year(weekMonthDate) & "-" & month(weekMonthDate) & "-" & day(weekMonthDate)
                            periodeTxt = weekdayname(weekday(weekMonthDate, 2), 0, 2)
                            end select




                        if cint(smilaktiv) <> 0 AND media <> "print" then

                            'call medrabSmilord(usemrn) '** Sættes på virksomhedsniveau
                            call smiley_agg_fn()
                            call meStamdata(usemrn)

                        call smileyAfslutSettings()

                            '** Henter timer i den uge der skal afsluttes ***'
                            call afsluttedeUger(year(now), usemrn, 0)

                        '** Er kriterie for afslutuge mødt? Ifht. medatype mindstimer og der må ikke være herreløse timer fra. f.eks TT
                            call timeKriOpfyldt(lto, sidsteUgenrAfsl, meType, usemrn, SmiWeekOrMonth)

                            timerdenneuge_dothis = 1
                            call timerDenneUge(usemrn, lto, tjkTimeriUgeSQL, akttypeKrism, timerdenneuge_dothis, SmiWeekOrMonth)
     
       
                            '** Periode aflsutning Dag / UGE / MD
                            select case cint(SmiWeekOrMonth) 
                            case 0
                            call thisWeekNo53_fn(varTjDatoUS_son)
                            weekMonthDate = thisWeekNo53 'datepart("ww", varTjDatoUS_son,2,2)
                            case 1
                            weekMonthDate = datepart("m", varTjDatoUS_son,2,2)
                            case 2
                            weekMonthDate = formatdatetime(now, 2) ' SKAL DET ALTID VÆRE DD?
                            weekMonthDate = year(weekMonthDate) & "-" & month(weekMonthDate) & "-" & day(weekMonthDate)
                            end select

                            'Response.Write "Hej der<br>"

                            call erugeAfslutte(datepart("yyyy", varTjDatoUS_son,2,2), weekMonthDate, usemrn, SmiWeekOrMonth, 0, tjkDato) 

         
                            '** Faneblade med afslutnings status
                        select case cint(SmiWeekOrMonth) 
                        case 0, 1
                        call smileyAfslutBtn(SmiWeekOrMonth)

                        call ugeAfsluttetStatus(varTjDatoUS_son, showAfsuge, ugegodkendt, ugegodkendtaf, SmiWeekOrMonth) 
                        end select
         
        


            
	                        '**** Afslut uge ***'
	                    '**** Smiley vises hvis sidste uge ikke er afsluttet, og dag for afslutninger ovwerskreddet ***' 12 
                        if cint(SmiWeekOrMonth) = 0 then
                        denneUgeDag = datePart("w", now, 2,2)
                        s0Show_sidstedagsidsteuge = dateAdd("d", -denneUgeDag, now) 'now
                        else
                        s0Show_sidstedagsidsteuge = now
                        end if

                        '** finder kriterie for rettidig afslutning
                        call rettidigafsl(s0Show_sidstedagsidsteuge, 0)

                        if cint(SmiWeekOrMonth) = 0 then
                            call thisWeekNo53_fn(s0Show_sidstedagsidsteuge)
                            s0Show_weekMd = thisWeekNo53 'datePart("ww", s0Show_sidstedagsidsteuge, 2,2)
                        else
                            s0Show_weekMd = datePart("m", s0Show_sidstedagsidsteuge, 2,2)
                        end if

        
                            '** HVORFOR DENNE IGEN?
                        '** tjekker om uge er afsluttet og viser afsluttet eller form til afslutning
		                call erugeAfslutte(year(s0Show_sidstedagsidsteuge), s0Show_weekMd, usemrn, SmiWeekOrMonth, 0, tjkDato)
      
      
      
            
             
                        if (cDate(formatdatetime(now, 2)) >= cDate(formatdatetime(cDateUge, 2)) AND cint(ugeNrStatus) = 0) OR cint(smiley_agg) = 1 then

	                    'if (datepart("w", now, 2,2) = 1 AND datepart("h", now, 2,2) <= 23 AND session("smvist") <> "j") OR cint(smiley_agg) = 1 then
    	                smVzb = "visible"
    	                smDsp = ""
    	                session("smvist") = "j"
    	                showweekmsg = "j"
    	                else
    	                smVzb = "hidden"
    	                smDsp = "none"
    	                showweekmsg = "n"
    	                end if
    	
    	

          
              

    	    
                            '*** Auto popup ThhisWEEK now SMILEY
	                        %>
	                        <!--<div id="s0" style="position:relative; left:0px; top:5px; width:725px; visibility:<%=smVzb%>; display:<%=smDsp%>; z-index:2; background-color:#FFFFFF; padding:20px; border:0px #CCCCCC solid;">-->
                            <br />&nbsp;<div id="s0" class="well" style="width:800px;">
	                        <%
                           
                            call afslutkri(varTjDatoUS_son, tjkTimeriUgeDt, usemrn, lto, SmiWeekOrMonth)

                            call afslutuge(weekSelectedTjk, 1, varTjDatoUS_son, "logindhist_2018", SmiWeekOrMonth)

                          
                            

                            '********************************************************************     
                            '*** Viser allerede afsluttede peridoer 
                            '********************************************************************
                            if cint(SmiWeekOrMonth) <> 2 then 'Afslut på dag. Giver ikke mening at vise. Uoverskueligt %>
                            <br /><br />
                            <span id="se_uegeafls_a" style="color:#5582d2;">[+] <%=tsa_txt_402 & " "& year(varTjDatoUS_son)%></span><br /><%

                            varTjDatoUS_ons = dateAdd("d", -3, varTjDatoUS_son)

	                        useYear = year(varTjDatoUS_ons)
                            '** Hvilke uger er afsluttet '***
	                        call smileystatus(usemrn, 1, useYear)
	        
                                %> <br />&nbsp;<%

                                end if
                
                
                                %>
              
	       

               
	                        </div>
        
                        <br /><br />
                        <%end if %>

                        <%if cint(smilaktiv) <> 0 AND media = "print" then 

           
                            call erugeAfslutte(datepart("yyyy", varTjDatoUS_son,2,2) , weekMonthDate, usemrn, SmiWeekOrMonth, 0, tjkDato) 

                            call ugeAfsluttetStatus(varTjDatoUS_son, showAfsuge, ugegodkendt, ugegodkendtaf, SmiWeekOrMonth) 

           

                        end if %>
                            
                        &nbsp;
                        </td>
                        </tr>

                    </table> </div> <!-- col-lg-6 -->

                    <%end if 'if (browstype_client <> "ip") then
                        
                        
                    if (browstype_client = "ip") then%>
                    <div class="row">
                    <%end if %>

                    <div class="col-lg-4"></div>


                        <!-- Uge skifter -->

                    <%if media <> "print" then %>

                    <%
                    call thisWeekNo53_fn(prev_varTjDatoUS_man)
                    prevWeek = thisWeekNo53 'datepart("ww", prev_varTjDatoUS_man, 2,2) 
                    call thisWeekNo53_fn(next_varTjDatoUS_man)
                    nextWeek = thisWeekNo53 'datepart("ww", next_varTjDatoUS_man, 2,2) 

                    call thisWeekNo53_fn(varTjDatoUS_man)
                    thisWeek = thisWeekNo53 'datepart("ww", varTjDatoUS_man, 2,2)
                   
                    %>
                    
                        <h4 class="col-lg-2">
                            <span class="pull-right">
                                <a href="logindhist_2011.asp?usemrn=<%=usemrn%>&varTjDatoUS_man=<%=prev_varTjDatoUS_man%>&varTjDatoUS_son=<%=prev_varTjDatoUS_son %>&rdir=<%=rdir %>&nomenu=<%=nomenu %>"><</a>
                                &nbsp Uge <%=thisWeek %> &nbsp
                                <a href="logindhist_2011.asp?usemrn=<%=usemrn%>&varTjDatoUS_man=<%=next_varTjDatoUS_man %>&varTjDatoUS_son=<%=next_varTjDatoUS_son %>&rdir=<%=rdir %>&nomenu=<%=nomenu %>">></a>
                            </span>
                        </h4>     
                        
                    <%end if %>



                    </div> <!-- Row -->






                    <%call stempelurlist_2018(useMrn, 0, 1, varTjDatoUS_man, varTjDatoUS_son, 0, d_end, lnk) %>


                    <%
                    if browstype_client <> "ip" then  
                    call fLonTimerPer(varTjDatoUS_man, 7, 0, useMrn)
                        
                    end if%>



                    <!---- Funktioner ---->

                    <%if media <> "print" AND browstype_client <> "ip" then %>
                    <br /><br /><br /><br />

                    

                    <section>
                    <div class="row">
                        <div class="col-lg-12"><b>Funktioner</b></div>
                    </div>

                     <form action="logindhist_2011.asp?usemrn=<%=usemrn%>&varTjDatoUS_man=<%=varTjDatoUS_man %>&varTjDatoUS_son=<%=varTjDatoUS_son %>&media=print" method="post" target="_blank">
                         <div class="row">
                             <div class="col-lg-4"><input id="Submit5" type="submit" value="Printvenlig version" class="btn btn-sm" /></div>
                         </div>                 
                    </form>



                    <% if (level <=2 OR level = 6) AND cint(smilaktiv) <> 0 then 

                    response.Write "<br><br>"

                    sm_aar = year(varTjDatoUS_son)
                    sm_mid = usemrn

                    'response.write "weekMonthDate: "& weekMonthDate

                    call erugeAfslutte(sm_aar, weekMonthDate, sm_mid, SmiWeekOrMonth, 0, tjkDato)
                     
                    fmlink = "ugeseddel_2011.asp?usemrn="& usemrn &"&varTjDatoUS_man="& varTjDatoUS_man &"&varTjDatoUS_son="& varTjDatoUS_son &"&nomenu="& nomenu &"&rdir=logindhist"
                    %>
           
                    <%call godkendugeseddel(fmlink, usemrn, varTjDatoUS_man, rdir) %>
        
                    <%end if 'level %>


             
                  </section>

                <%end if %>




                </div>
            </div>
        </div>



        <%end select %>


    </div>
</div>



    

<!--#include file="../inc/regular/footer_inc.asp"-->