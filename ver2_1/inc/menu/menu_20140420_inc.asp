


<%
level = session("rettigheder")

public lnkTimeregside, lnkUgeseddel, lnkAfstem, lnkLogind


function menu_2014()

           

            call TimeOutVersion()
            call erCRMaktiv()
			call erSDSKaktiv()
            call erERPaktiv()
            call erStempelurOn()
            call ersmileyaktiv()
			call smileyAfslutSettings()
            call smiley_agg_fn()

            '*** Skal der lukkes ned for timereg. cint(smiley_agg) = 1 
            if cint(smiley_agg) = 1 then
            call afsluttedeUger(year(now), session("mid"), 1)
            end if
           

         


select case lto
case "xintranet - local", "nt"
    tregmenu_2014 = 0
    projmenu_2014 = 2
    statmenu_2014 = 2
  

    if level <= 3 OR level = 6 then
    kundmenu_2014 = 1
    else
    kundmenu_2014 = 0
    end if

    marbmenu_2014 = 1
    
    if level = 1 then
    lagemenu_2014 = 1
    else
    lagemenu_2014 = 0
    end if

    
    
    if level = 1 then
    admimenu_2014 = 1
    else
    admimenu_2014 = 0
    end if 

    if cint(erpOnOff) = 1 AND (level <= 2 OR level = 6) then
    erpfmenu_2014 = 1 
    else
    erpfmenu_2014 = 0
    end if

    if cint(dsksOnOff) = 1 AND (level <= 2 OR level = 6) then
    helpmenu_2014 = 0
    else
    helpmenu_2014 = 0
    end if



case "epi", "epi_osl", "epi_sta", "epi_ab", "epi_uk", "ascendis"

    tregmenu_2014 = 1
    
     if level <= 3 OR level = 6 then
    projmenu_2014 = 1
    else
    projmenu_2014 = 0
    end if

    statmenu_2014 = 1

    if level <= 3 OR level = 6 then
    kundmenu_2014 = 1
    else
    kundmenu_2014 = 0
    end if

    marbmenu_2014 = 1
    lagemenu_2014 = 0
  

    
    if level = 1 then
    admimenu_2014 = 1
    else
    admimenu_2014 = 0
    end if 

    if cint(erpOnOff) = 1 AND (level <= 2 OR level = 6) then
    erpfmenu_2014 = 1 
    else
    erpfmenu_2014 = 0
    end if
    
    if cint(dsksOnOff) = 1 AND (level <= 2 OR level = 6) then
    helpmenu_2014 = 1
    else
    helpmenu_2014 = 0
    end if


case "tec", "esn", "xintranet - local"


    call medariprogrpFn(session("mid"))

    medariprogrpTxtDage = medariprogrpTxt


     tregmenu_2014 = 1
    
    if level = 1 OR (instr(medariprogrpTxtDage, "#71#") <> 0) then
    projmenu_2014 = 1
    else
    projmenu_2014 = 0
    end if

    if level = 1 OR (instr(medariprogrpTxtDage, "#71#") <> 0) then
    statmenu_2014 = 1
    else
    statmenu_2014 = 0
    end if

    if level = 1 then
    kundmenu_2014 = 0
    else
    kundmenu_2014 = 0
    end if
    

    marbmenu_2014 = 1
    
    if level = 1 then
    lagemenu_2014 = 0
    else
    lagemenu_2014 = 0
    end if

    
    if level = 1 then
    admimenu_2014 = 1
    else
    admimenu_2014 = 0
    end if 

    'if cint(erpOnOff) = 1 AND (level <= 2 OR level = 6) then
    'erpfmenu_2014 = 1 
    'else
    erpfmenu_2014 = 0
    'end if
    
    if cint(dsksOnOff) = 1 AND (level <= 2 OR level = 6) then
    helpmenu_2014 = 1
    else
    helpmenu_2014 = 0
    end if

case "nonstop"


    call medariprogrpFn(session("mid"))

    medariprogrpTxtDage = medariprogrpTxt


     tregmenu_2014 = 1
    
    if level = 1  then
    projmenu_2014 = 1
    else
    projmenu_2014 = 0
    end if

    if level = 1 then
    statmenu_2014 = 1
    else
    statmenu_2014 = 0
    end if

    if level = 1 then
    kundmenu_2014 = 1
    else
    kundmenu_2014 = 0
    end if
    

    marbmenu_2014 = 1
    
    if level = 1 then
    lagemenu_2014 = 0
    else
    lagemenu_2014 = 0
    end if

    
    if level = 1 then
    admimenu_2014 = 1
    else
    admimenu_2014 = 0
    end if 

    if cint(erpOnOff) = 1 AND (level <= 2 OR level = 6) then
    erpfmenu_2014 = 1 
    else
    erpfmenu_2014 = 0
    end if
    
    if cint(dsksOnOff) = 1 AND (level <= 2 OR level = 6) then
    helpmenu_2014 = 1
    else
    helpmenu_2014 = 0
    end if


case "bf"

    tregmenu_2014 = 1

    if (level <= 3 OR level = 6) then
    projmenu_2014 = 1
    else
    projmenu_2014 = 0
    end if

    statmenu_2014 = 1

    if level <= 3 OR level = 6 then
    kundmenu_2014 = 1
    else
    kundmenu_2014 = 0
    end if

    marbmenu_2014 = 1
    lagemenu_2014 = 0
    
    
    
    if level = 1 then
    admimenu_2014 = 1
    else
    admimenu_2014 = 0
    end if 

    if cint(erpOnOff) = 1 AND (level <= 2 OR level = 6) then
    erpfmenu_2014 = 1 
    else
    erpfmenu_2014 = 0
    end if
    
    if cint(dsksOnOff) = 1 AND (level <= 2 OR level = 6) then
    helpmenu_2014 = 1
    else
    helpmenu_2014 = 0
    end if


case "gd"

    kundmenu_2014 = 1
    marbmenu_2014 = 1
    admimenu_2014 = 1
    helpmenu_2014 = 1
    dsksOnOff = 0

case else
    tregmenu_2014 = 1

    if (level <= 3 OR level = 6) then
    projmenu_2014 = 1
    else
    projmenu_2014 = 0
    end if

    statmenu_2014 = 1

    if level <= 3 OR level = 6 then
    kundmenu_2014 = 1
    else
    kundmenu_2014 = 0
    end if

    marbmenu_2014 = 1
    
    if level <= 3 OR level = 6 then
    lagemenu_2014 = 1
    else
    lagemenu_2014 = 1
    end if
    
    
    if level = 1 then
    admimenu_2014 = 1
    else
    admimenu_2014 = 0
    end if 

    if cint(erpOnOff) = 1 AND (level <= 2 OR level = 6) then
    erpfmenu_2014 = 1 
    else
    erpfmenu_2014 = 0
    end if
    
    if cint(dsksOnOff) = 1 AND (level <= 2 OR level = 6) then
    helpmenu_2014 = 1
    else
    helpmenu_2014 = 0
    end if
end select
    
%>

    <div class="fixed-navbar-hoz">

          
        <h1 class="menu_logo-left">
            <a href="<%=toSubVerPath14 %>../login.asp"></a>
        </h1>
       

          <%if cint(slip_smiley_agg_lukper) <> 1 then %>
            <nav class="dropdown-right">
         <ul> 



                <li>
                    <a href="#"><span class="glyph icon-user"></span><span class="account-name"><%=session("user") %></span><b class="caret-down"></b></a>
                    <ul>

                        
                        <li><a href="<%=toSubVerPath14 %>help.asp">Hj&aelig;lp</a></li>
                        <%if lto <> "tec" AND lto <> "esn" then %>
                        <li><a href="https://www.islonline.net/start/ISLLightClient"><%=tsa_txt_433 %></a></li>
                        <%end if %>

                        <%if level = 1 then %>
                        <li><a href="<%=toSubVerPath14 %>kontrolpanel.asp"><%=tsa_txt_434 %></a></li>
                        <%end if %>
                        <%
                        
                        
                        'call erStempelurOn()
           
                        'if cint(stempelur_hideloginOn) = 1 then    
                            
                        if session("stempelur") <> 0 AND lto <> "tec" AND lto <> "esn"  then %>
                        <li style="background-color:red;"><a href="<%=toSubVerPath14 %>stempelur.asp?func=redloginhist&medarbSel=<%=session("mid")%>&showonlyone=1&hidemenu=1&id=0&rdir=sesaba" target="_top">Log ud</a></li>
                        <%else %>
                        <li style="background-color:red;"><a href="<%=toSubVerPath14 %>../sesaba.asp" target="_top"><%=tsa_txt_435 %></a></li>
                        <%end if %>
                        <li style="background-color:#999999;"><a href="#"><%=tsa_txt_436 %><br /><%=lto %></a></li>
                    </ul>
                </li>
            </ul>
        </nav>
       <%end if %>
    </div>
    <div class="fixed-navbar-vert">
        <nav>
            <ul>
                <%if cint(tregmenu_2014) <> 0 AND (lto = "demo" OR lto = "intranet - local") then %>
                <li id="menu_0" class="showLeft tooltip-right" tooltip-title="Home"><a href="<%=toSubVerPath15 %>home_dashboard_demo.asp"><span class="glyph icon-home"></span></a></li>
                <%end if %>

                <%if cint(tregmenu_2014) <> 0 then %>
                <li id="menu_1" class="showLeft tooltip-right" tooltip-title="<%=menu_txt_001 %>"><a href="#"><span class="glyph icon-stopwatch"></span></a></li><!-- stopwatch -->
                <%end if %>

                <%if cint(projmenu_2014) <> 0 AND (level <=2 OR level = 6) OR cint(projmenu_2014) = 2 then 
                    select case lto
                    case "nt", "intranet - local"
                    tooltip2 = "Orders"
                    case else
                    tooltip2 = menu_txt_002
                    end select
                    %>
                <li id="menu_2" class="showLeft tooltip-right" tooltip-title="<%=tooltip2 %>"><a href="#"><span class="glyph icon-file-cabinet"></span></a></li>
                <%end if %>

                  <%if cint(statmenu_2014) <> 0 then %>
                <li id="menu_3" class="showLeft tooltip-right" tooltip-title="<%=menu_txt_003 %>"><a href="#"><span class="glyph icon-chart-graph2"></span></a></li>
                <%end if %>

                  <%if cint(kundmenu_2014) <> 0 AND (level <=3 OR level = 6) then %>
                <li id="menu_4" class="showLeft tooltip-right" tooltip-title="<%=menu_txt_004 %>"><a href="#"><span class="glyph icon-building"></span></a></li>
                <%end if %>

                  <%if cint(marbmenu_2014) <> 0 then %>
                <li id="menu_5" class="showLeft tooltip-right" tooltip-title="<%=menu_txt_005 %>"><a href="#"><span class="glyph icon-user"></span></a></li>
                <%end if %>

                  <%if cint(lagemenu_2014) <> 0 AND (level <=2 OR level = 6) then %>
                <li id="menu_6" class="showLeft tooltip-right" tooltip-title="<%=menu_txt_006 %>"><a href="#"><span class="glyph icon-barcode "></span></a></li>
                <%end if %>

                  <%if cint(erpfmenu_2014) <> 0 then %>
                <li id="menu_7" class="showLeft tooltip-right" tooltip-title="<%=menu_txt_007 %>"><a href="#"><span class="glyph icon-piggy-bank"></span></a></li>
                <%end if %>

                       <%if cint(helpmenu_2014) <> 0 then %>
                <li id="menu_9" class="showLeft tooltip-right" tooltip-title="<%=menu_txt_008 %>"><a href="#"><span class="glyph icon-help"></span></a></li>
                <%end if %>

                  <%if cint(admimenu_2014) <> 0 then %>
                <li id="menu_8" class="showLeft tooltip-right" tooltip-title="<%=menu_txt_009 %>"><a href="#"><span class="glyph icon-tools"></span></a></li>
                <%end if %>

           
            </ul>


             
        </nav>

       

    </div>


 <nav id="menu-slider" class="menu-slider" style="overflow-y:scroll;">

             <span style="color:#999999; float:right; font-size:14px; padding-right:20px;" id="luk_menuslider">X</span>
         
            <!-- Timeregistrering -->
             <ul class="menupkt_n2" id="ul_menu-slider-1" style="display:none; visibility:visible;">
                <h3 class="menuh3" style="font-size:16px; font-weight:200;"><%=tsa_txt_437 %></h3>
                 <!--
                <summary>I denne sektion har vi samlet alt vedr. Timeregistrering. Du har her mulighed for at registrere timer eller få et overblik over dine nuværende registreringer</summary>
                 -->
              
         

                <%

                if len(trim(varTjDatoUS_man)) <> 0 then
                varTjDatoUS_man = varTjDatoUS_man
                else
                    if session("varTjDatoUS_man") <> "" then
                    varTjDatoUS_man = session("varTjDatoUS_man")
                    else
                    varTjDatoUS_man = year(now)&"/"&month(now)&"/"&day(now)
                    varTjDatoUS_man = year(varTjDatoUS_man)&"/"&month(varTjDatoUS_man)&"/"&day(varTjDatoUS_man)

                    end if
                end if


                 

                 if varTjDatoUS_now_WeekDay <> 1 then
                 varTjDatoUS_now_WeekDay = datePart("w", varTjDatoUS_man, 2,2)
                 'response.Write "varTjDatoUS_man_WeekDay: "& varTjDatoUS_now_WeekDay & "<br>"
                 stDato_ugeseddel_kommegaa_US_man = dateAdd("d", - (varTjDatoUS_now_WeekDay) + 1, varTjDatoUS_man)
                 else
                 stDato_ugeseddel_kommegaa_US_man = varTjDatoUS_man
                 end if

                 stDato_ugeseddel_kommegaa_US_man = year(stDato_ugeseddel_kommegaa_US_man)&"/"&month(stDato_ugeseddel_kommegaa_US_man)&"/"& day(stDato_ugeseddel_kommegaa_US_man)

                 varTjDatoUS_man = year(varTjDatoUS_man)&"/"&month(varTjDatoUS_man)&"/"&day(varTjDatoUS_man)

                 'response.Write "<br>varTjDatoUS_man: "& varTjDatoUS_man


                 strdagTreg = day(varTjDatoUS_man)
                 strmrdTreg = month(varTjDatoUS_man)
                 straarTreg = year(varTjDatoUS_man)

                session("varTjDatoUS_man") = varTjDatoUS_man


                if len(trim(usemrn)) <> 0 then
                    usemrn = usemrn
                else
                    usemrn = session("mid")
                end if

                'lnkTreg = "timereg_akt_2006.asp?showakt=1&hideallbut_first=2&strdag="&day(varTjDatoUS_man)&"&strmrd="& month(varTjDatoUS_man) &"&straar="&year(varTjDatoUS_man)
                lnkTimeregside = "timereg_akt_2006.asp?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&strdag="&strdagTreg&"&strMrd="&strMrdTreg&"&strAar="&strAarTreg
                lnkUgeseddel = "ugeseddel_2011.asp?usemrn="&usemrn&"&varTjDatoUS_man="&stDato_ugeseddel_kommegaa_US_man'&"&varTjDatoUS_son="&varTjDatoUS_son
	            lnkAfstem = "afstem_tot.asp?usemrn="&usemrn&"&show=5&varTjDatoUS_man="&varTjDatoUS_man'&"&varTjDatoUS_son="&varTjDatoUS_son
                lnkLogind = "logindhist_2011.asp?usemrn="&usemrn&"&varTjDatoUS_man="&stDato_ugeseddel_kommegaa_US_man'&"&varTjDatoUS_son="&varTjDatoUS_son 
                
                    
                 select case lto
                 case "cst"
                    %>   <li><a href="<%=toSubVerPath14 %><%=lnkTimeregside %>"><%=tsa_txt_116 &" "& tsa_txt_438 %></a></li><%
                 case "tec", "esn"
                       
                        %>   
                        <li><a href="<%=toSubVerPath14 %><%=lnkTimeregside %>"><%=tsa_txt_116 &" "& tsa_txt_438 %></a></li><%
                       
                 case else
                    %>   <li><a href="<%=toSubVerPath14 %><%=lnkTimeregside %>"><%=tsa_txt_116 %></a></li><%
                 end select
           
               

                 %>  <li><a href="<%=toSubVerPath15 %><%=lnkUgeseddel%>"><%=tsa_txt_337 %></a></li>
                      
                
                

               

                  <%if cint(stempelurOn) = 1 then %>
                 <li><a href="<%=toSubVerPath14 %><%=lnkLogind%>"><%=tsa_txt_340 %></a></li>
                 <%end if %>

                
                 <% select case lto
                
                 case "tec", "esn", "epi", "epi_uk", "epi_no", "epi_as"
                 case else
                    
                     if level <= 7 then %>
                     <li><a href="<%=toSubVerPath14 %>materialer_indtast.asp?id=0&fromsdsk=0&aftid=0"><%=tsa_txt_191 %></a></li>
	                 <%end if 
                 end select%>
                 
                  <% 
                      
                  call traveldietexp_fn()
                      
                  if cint(traveldietexp_on) = 1 then
                  %>
                  <li><a href="<%=toSubVerPath15 %>traveldietexp.asp">Rejse/Diæter</a></li>
                  <%end if%>
               

                 <li><a href="<%=toSubVerPath14 %><%=lnkAfstem %>"><%=tsa_txt_389 %></a></li>

                 <%select case lto
                 case "oko", "epi", "epi2017", "wilke", "intranet - local", "outz", "dencker", "essens", "synergi1", "jttek", "hidalgo", "demo", "bf", "plan", "acc", "assurator", "glad"
                     
                    if level = 1 OR (lto = "wilke") OR (lto = "outz") OR (lto = "dencker") OR (lto = "hidalgo") OR (lto = "acc") OR (lto = "epi2017") OR (lto = "assurator") OR (lto = "glad") then%>
                    <li><a href="<%=toSubVerPath15 %>medarbdashboard.asp"><%=tsa_txt_529 %></a></li>
                    <%end if %>
                 <%end select %>
                

             
                   

                   <%if level <= 7 then %>


                       <% 
                       '*** OVERSKRIFT projektleder funktioner    
                       %>

                        <h3 class="menuh3"><%=tsa_txt_439 %></h3>




                       <%'** Igangværende arbejde ***'
                        select case lto
                
                        case "tec", "esn"%>
                        <%case else 
                            
                                if level <= 2 OR level = 6 then%>
                               <li><a href="<%=toSubVerPath14 %>webblik_joblisten.asp"><%=tsa_txt_452 %></a></li>
                        <%      end if
                        end select %>
                           
                            
                      <% 
                     '**** Ressource forecast    
                     select case lto
                     case "tec", "esn", "nonstop", "epi2017"
                     case else
                    
                             if level <= 2 OR level = 6 then %>
                            <li><a href="<%=toSubVerPath14 %>ressource_belaeg_jbpla.asp"><%=tsa_txt_440 %></a></li>
                             <%end if%>

                    <%end select %>

	               
                  <%end if %>

                

                  <%'** Afslut periode ***'
                    if cint(SmiWeekOrMonth) = 0 then
                     afslutUgeseddelMenuTxt = tsa_txt_356
                     else

                        select case lto
                        case "tec", "esn"
                        afslutUgeseddelMenuTxt = tsa_txt_441 
                        case else
                        afslutUgeseddelMenuTxt = tsa_txt_442
                        end select

                        
                    end if
                      
                    if ((level <= 2 OR level = 6) OR lto = "tec" OR lto = "esn" OR lto = "intranet - local") AND cint(smilaktiv) <> 0 then %>
                   <li><a href="<%=toSubVerPath14 %>week_real_norm_2010.asp"><%=afslutUgeseddelMenuTxt %></a></li>
	               <%end if %>


                  <%
                   '** Ferie / Sygdoms kalender
                   '** Level 1 har adgang til sygdom OGSÅ TEC og ESN da der ikke tjekkes på medarbejderlinej niveau i feriekalender
                   '** Ellers kun ferie  
                   if level = 1 then %>

	                    <li><a href="<%=toSubVerPath14 %>feriekalender.asp?menu=job"><%=replace(tsa_txt_443, "|", "&") %></a></li>
		           
                 
                  <%else %>

                         <%select case lto
                
                         case "tec", "esn"

                                  if level <= 2 OR level = 6 then
                                  %>
                                    <li><a href="<%=toSubVerPath14 %>feriekalender.asp?menu=job"><%=replace(tsa_txt_443, "|", "&") %></a></li>
                                  <%end if

                         case "nonstop"

                         case else
                    
                          %>

		                     <li><a href="<%=toSubVerPath14 %>feriekalender.asp?menu=job"><%=replace(tsa_txt_444, "|", "&") %></a></li>

                        <%end select %>

		          <%end if %>


                 <%


            call stadeOn()

            if jobasnvigv = 1 then %>
            <li><a href="<%=toSubVerPath14 %>stat_opdater_igv.asp?func=opdater" target="_blank"><%=tsa_txt_462 %></a>
            <%end if 
               




     call meStamdata(session("mid")) 
     if cint(meVisskiftversion) = 1 then

     %>
    <h3 class="menuh3"><%=tsa_txt_461 %></h3> 
    <%
        
       if lto = "biofac" then
     
        %>
       <li><a href="https://outzource.dk/timeout_xp/wwwroot/ver2_14/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=support@outzource.dk&key=2.2013-0912-TO146" >Adminbruger </a></li>
       <li><a href="https://outzource.dk/timeout_xp/wwwroot/ver2_14/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=me@outzource.dk&key=2.2013-0912-TO146" >Medarbejder </a></li>
        <%
       end if

        if lto = "demo" then
        %>
       <li><a href="https://outzource.dk/timeout_xp/wwwroot/ver2_14/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=support@outzource.dk&key=2.052-xxxx-B000" >Adminbruger </a></li>
       <li><a href="https://outzource.dk/timeout_xp/wwwroot/ver2_14/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=me@outzource.dk&key=2.052-xxxx-B000" >Medarbejder </a></li>
        <%
       end if


        if lto = "outz" then
       
      %>
       <li><a href="https://outzource.dk/timeout_xp/wwwroot/ver2_14/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2012-1012-TO136" >EPI NO</a></li>
       <%
       end if

       if lto = "epi" OR lto = "intranet - local" then
     
       %>
        <li><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2012-1012-TO136" >EPI NO</a></li>
        <li><a href="https://outzource.dk/timeout_xp/wwwroot/ver2_14/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2013-0109-TO142" >EPI AB</a></li>
      <li><a href="https://outzource.dk/timeout_xp/wwwroot/ver2_14/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2014-0331-TO148" >EPI UK</a></li>
       <%
     
       end if

        if lto = "epi_osl" OR lto = "epi_no" then
     
      %>
       <li><a href="http://timeout.cloud/timeout_xp/wwwroot/ver2_14/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2010-2004-TO116" >EPI</a></li>
        <li><a href="https://outzource.dk/timeout_xp/wwwroot/ver2_14/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2013-0109-TO142" >EPI AB</a></li>
      <li><a href="https://outzource.dk/timeout_xp/wwwroot/ver2_14/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2014-0331-TO148" >EPI UK</a></li>
       <%
   
       end if

       if lto = "epi_ab" then
      
       %>
       <li><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2010-2004-TO116" >EPI</a></li>
       <li><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2012-1012-TO136" >EPI NO</a></li>
                 <li><a href="https://outzource.dk/timeout_xp/wwwroot/ver2_14/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2014-0331-TO148" >EPI UK</a></li>
       <%
       
       end if

        if lto = "xxepi_sta" then
     
    %>
       <li><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2012-1012-TO136" >EPI NO</a></li>
        <li><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2010-2004-TO116" >EPI</a></li>
        <li><a href="https://outzource.dk/timeout_xp/wwwroot/ver2_14/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2013-0109-TO142" >EPI AB</a></li>
      <li><a href="https://outzource.dk/timeout_xp/wwwroot/ver2_14/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2014-0331-TO148" >EPI UK</a></li>
       <%
       
       end if

    
          if lto = "epi_uk" then
     
    %>
        <li><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2012-1012-TO136" >EPI NO</a></li>
        <li><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2010-2004-TO116" >EPI</a></li>
        <li><a href="https://outzource.dk/timeout_xp/wwwroot/ver2_14/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2013-0109-TO142" >EPI AB</a></li>
      
       <%
       
       end if


       if lto = "mmmi" then
      
      %>
       <li><a href="https://outzource.dk/timeout_xp/wwwroot/ver2_10/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2011-0410-TO127" >UNIK</a></li>
       <%end if
       


       if lto = "unik" then
    
      %>
       <li><a href="https://outzource.dk/timeout_xp/wwwroot/ver2_10/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2010-0311-TO119" >MMMI</a></li>
       <%
       end if
       
         
    end if
     %>
                
                
                              
            </ul>   






            <!-- PROJEKTETER -->
              <ul class="menupkt_n2" id="ul_menu-slider-2" style="display:none; visibility:visible;">
               
                <!--<summary>I denne sektion har vi samlet alt vedr. Projekter.</summary>-->
              
                <%
                '*** NT

                    

                if cint(projmenu_2014) = 2 then%>
                     <h3 class="menuh3"><%=tsa_txt_445 %></h3>
                      <li><a href="<%=toSubVerPath15 %>job_nt.asp?rapporttype=0"><%=tsa_txt_445 %></a></li>
                      <li><a href="<%=toSubVerPath15 %>job_nt.asp?func=opret&id=0&rapporttype=0"><%=tsa_txt_446 %></a></li>
                    
                 <%
                 '** Standard    
                 else %>
                 <h3 class="menuh3"><%=tsa_txt_447 %></h3>
                <li><a href="<%=toSubVerPath14 %>jobs.asp"><%=tsa_txt_448 %></a></li>
                
                 <%'** Opret job ****'%>
                 <li><a href="<%=toSubVerPath14 %>jobs.asp?func=opret&id=0"><%=tsa_txt_449 %></a></li>


                  <%if lto = "outz" OR lto = "demo" OR lto = "intranet - local" then %>
                  <li><a href="<%=toSubVerPath15 %>jobs.asp?func=opret&id=0"><%=tsa_txt_449 %> simpel</a></li>
                  <%end if %>
              

                  
                  <%if level <= 2 OR level = 6 then %>
                  <li><a href="<%=toSubVerPath14 %>job_print.asp?menu=job&kid=0&id=0"><%=tsa_txt_450 %></a></li>
                  <li><a href="<%=toSubVerPath14 %>filer.asp"><%=tsa_txt_451 %></a></li>
                  <%end if %>

                    <%if level <= 2 OR level = 6 then %>
                  <h3 class="menuh3"><%=tsa_txt_439 %></h3>
                   
                  <li><a href="<%=toSubVerPath14 %>webblik_joblisten.asp"><%=tsa_txt_452 %></a></li>
                 

                 
                 

                  <li><a href="<%=toSubVerPath14 %>ressource_belaeg_jbpla.asp"><%=tsa_txt_440 %>

                      <%if lto = "wwf" then%>
                          
                      (<%=tsa_txt_453 %>)
                          
                      <%end if %>

                      </a></li>
                  

                 
                   <li><a href="<%=toSubVerPath14 %>webblik_joblisten21.asp"><%=tsa_txt_454 %></a></li>
                   <li><a href="<%=toSubVerPath14 %>webblik_milepale.asp"><%=tsa_txt_455 %></a></li>

                  <%if (level = 1 OR session("mid") = 35) then 'Kim B epinion %>
                  <li><a href="<%=toSubVerPath14 %>pipeline.asp?menu=webblik&FM_kunde=0&FM_progrupper=10"><%=tsa_txt_456 %></a></li>
                  <%end if %>

                  <%if lto = "outz" OR lto = "intranet - local" then %>
                    <li><a href="<%=toSubVerPath14 %>jbpla_w.asp">Planlægningskalender (TEST)</a></li>
                  <%end if %>

                     <%if lto = "outz" OR lto = "intranet - local" OR lto = "essens" OR lto = "hidalgo" then %>
                    <li><a href="../ressource_planner/ressplan_2017.aspx" target="_blank">Ressource Planner</a></li>
                  <%end if %>

                   <%end if %>
                  


                  <%end if %>
            </ul>   




             <!-- STATISTIK / RAPPORTER -->
             <ul class="menupkt_n2" id="ul_menu-slider-3" style="display:none; visibility:visible;">
                <h3 class="menuh3"><%=tsa_txt_457 %></h3>
            
                
                     <%if cint(statmenu_2014) = 2 then 'nt%>
                

                           <%if level = 1 then %>
                 
                                     <li><a href="<%=toSubVerPath14 %>oms.asp"><%=tsa_txt_459 %></a></li>

                            <%end if %>
                   
                          <!--
                         <li><a href="<%=toSubVerPath14 %>job_nt.asp?rapporttype=1">Production / Enq. Overview</a></li>
                         <li><a href="<%=toSubVerPath14 %>job_nt.asp?rapporttype=3">Overview (ext.)</a></li>
                         <li><a href="<%=toSubVerPath14 %>job_nt.asp?rapporttype=4">Overview</a></li>
                         <li><a href="<%=toSubVerPath14 %>oms.asp">Sales Budget</a></li> 
                         -->

                 <%else %>

                         <br /><br />
                         <!--<li><a href="../to_2015/progtotal_medarb.asp">Test</a></li> -->
                         <li><a href="<%=toSubVerPath14 %>joblog_timetotaler.asp"><%=tsa_txt_458 %></a></li>
                         <li><a href="<%=toSubVerPath14 %>joblog.asp"><%=tsa_txt_118 %></a></li>
                 
                         <%if level = 1 then %>
                 
                                     <li><a href="<%=toSubVerPath14 %>oms.asp"><%=tsa_txt_459 %></a></li>



                                     <li><a href="<%=toSubVerPath14 %>fomr.asp?func=stat"><%=tsa_txt_460 %></a></li>
                                
                        <%end if %>

                        <%if (level <= 2 OR level = 6) OR lto = "fk" then %>

                                    <li><a href="<%=toSubVerPath15 %>medarb_protid.asp"><%=menu_txt_010 %></a></li>

                        <%end if %>

                       <%if (lto = "epi2017" AND (level = 1 OR session("mid") = 2694) ) then
                           %>
                              <li><a href="<%=toSubVerPath14 %>stempelur.asp?func=stat"><%=tsa_txt_463 %></a></li>
                       <%end if%>


                       <%if level = 1 then %>

                                         <br /><br />
                                         <%if session("stempelur") <> 0 AND lto <> "epi2017" then %>
                                            <li><a href="<%=toSubVerPath14 %>stempelur.asp?func=stat"><%=tsa_txt_463 %></a></li>

                                        <%end if %>
                 
             

                             <li><a href="<%=toSubVerPath14 %>saleandvalue.asp"><%=replace(tsa_txt_464, "|", "&") %></a></li>

                             <%if cint(smilaktiv) <> 0 then  %>
                             <li><a href="<%=toSubVerPath14 %>smileystatus.asp"><%=tsa_txt_465 %></a></li>
                             <%end if %>

                         <%end if %>


                         <%if level = 1 OR ((level <= 2 OR level = 6) AND lto = "epi2017") then %>

                            <li><a href="<%=toSubVerPath14 %>bal_real_norm_2007.asp?dontdisplayresult=1"><%=replace(tsa_txt_466, "|", "&") %></a></li>
                         <%end if %>
                
                 
                         <%if level = 1 then %>
	                    <li><a href="<%=toSubVerPath14 %>feriekalender.asp?menu=job"><%=replace(tsa_txt_443, "|", "&") %></a></li>
		                <%else %>
		                 <li><a href="<%=toSubVerPath14 %>feriekalender.asp?menu=job"><%=replace(tsa_txt_444, "|", "&") %></a></li>
		                <%end if %>



                         <%if cint(smilaktiv) <> 0 AND level <= 2 OR level = 6 then  %>
                        <li><a href="<%=toSubVerPath14 %>week_real_norm_2010.asp?menu=stat"><%=tsa_txt_356 %></a></li>
                        <%end if %>

                        <%if level <= 2 OR level = 6 then '** Indtil teamleder er impl. på kørsels siden **"%>
                        <li><a href="<%=toSubVerPath14 %>stat_korsel.asp?menu=stat" ><%=tsa_txt_265 %></a></li>
                        <li><a href="<%=toSubVerPath14 %>pipeline.asp?menu=webblik&FM_kunde=0&FM_progrupper=10"><%=tsa_txt_456 %></a></li>
                        <%end if %>

                        <br /><br />


                        <%if level = 1 then %>
                        <li><a href="<%=toSubVerPath14 %>materiale_stat.asp?menu=stat" ><%=tsa_txt_156 %></a></li>
                        <%end if %>

                        


                        <%if level = 1 AND (lto = "intranet - local" OR lto = "epi" OR lto = "epi_cati" OR lto = "epi_no" OR lto = "epi_se" OR lto = "epi2017") then %>
                        
                         <br /><br />
                        <li><a href="<%=toSubVerPath14 %>timerimperr.asp?menu=stat" ><%=tsa_txt_467 %></a></li>
                        <li><a href="<%=toSubVerPath14 %>diff_timer_kons.asp?menu=stat" ><%=tsa_txt_468 %></a></li>
                        <%end if %>

                        <%


                        jobasnvigv = 0
                        strSQLigv = "SELECT jobasnvigv FROm licens WHERE id = 1"
                        oRec.open strSQLigv, oConn, 3
                        if not oRec.EOF then
                        jobasnvigv = oRec("jobasnvigv")
                        end if
                        oRec.close

                        if jobasnvigv = 1 AND level = 1 then %>
                        <li><a href="<%=toSubVerPath14 %>stat_opdater_igv.asp?menu=stat"><%=tsa_txt_462 %></a></li>
                        <%end if %>


                    <% 
                     'if (lto = "oko") OR lto = "intranet - local" then 'jobresume 
                        
                        
                         select case lto
                         case "bf"
                        
                         case else%>


                         <li><a href="<%=toSubVerPath14 %>jobprintoverblik.asp?menu=job&id=0" target="_blank">Joboverblik / Resume</a></li>

                         <%end select %>

                    <%'end if %>


                 <%end if %>

            </ul>   


        <!-- Kunder -->
       <ul class="menupkt_n2" id="ul_menu-slider-4" style="display:none; visibility:visible;">
                <h3 class="menuh3"><%=global_txt_124 %></h3>
            
                
	        <li><a href='<%=toSubVerPath15 %>kunder.asp?visikkekunder=1'><%=global_txt_124 %></a></li>
            <li><a href='<%=toSubVerPath15 %>kunder.asp?func=opret&ketype=k'><%=global_txt_180 %></a></li>           
            <li><a href='<%=toSubVerPath14 %>serviceaft_osigt.asp?id=0&func=osigtall'><%=global_txt_181 %></a></li>       
            <li><a href='<%=toSubVerPath15 %>kontaktpers.asp?func=list'><%=global_txt_182 %></a></li>

           
           <%if lto = "gd" then %>            
           <li><a href="<%=toSubVerPath14 %>filer.asp"><%=tsa_txt_451 %></a></li>
           <%end if %>

           <%if cint(crmOnOff) = 1 then %>
           <h3 class="menuh3"><%=tsa_txt_491 %></h3>
           <!--<li><a href='<%=toSubVerPath14 %>crmkalender.asp?menu=crm&shokselector=1&ketype=e&selpkt=kal&status=0&id=0&emner=0'>Kalender (gl.)</a></li>-->          
           <li><a href='<%=toSubVerPath14 %>kunder.asp?menu=crm'><%=tsa_txt_469 %></a></li>
           <li><a href='<%=toSubVerPath14 %>crmhistorik.asp?menu=crm&ketype=e&func=hist&id=0&selpkt=hist'><%=tsa_txt_470 %></a></li>
           <!--<li><a href='<%=toSubVerPath14 %>crmstat.asp?menu=crm'>CRM statistik</a></li>-->

           <%end if %>
		                
            </ul>   
  

       <!-- Medarbejdere -->
       <ul class="menupkt_n2" id="ul_menu-slider-5" style="display:none; visibility:visible;">
                <h3 class="menuh3"><%=tsa_txt_471 %></h3>
            
                
           <%if level = 1 then %>
	        <li><a href='<%=toSubVerPath15 %>medarb.asp?visikkemedarbejdere=1'><%=global_txt_125 %></a></li>
            <li><a href='<%=toSubVerPath15 %>medarb.asp?menu=medarb&func=opret'><%=global_txt_183 %></a></li>
            <li><a href='<%=toSubVerPath15 %>medarb.asp?menu=medarb&func=red&id=<%=session("mid")%>'><%=global_txt_184 %></a></li>
           <%else 
               
               
               
            %>

            <%if cint(create_newemployee) = 1 then %>
            <li><a href='<%=toSubVerPath15 %>medarb.asp?visikkemedarbejdere=1'><%=global_txt_125 %></a></li>
            <li><a href='<%=toSubVerPath15 %>medarb.asp?menu=medarb&func=opret'><%=global_txt_183 %></a></li>
            <%end if %>

            <li><a href='<%=toSubVerPath15 %>medarb.asp?menu=medarb&func=red&id=<%=session("mid")%>'><%=global_txt_184 %></a></li>

         

           <%end if %>

          
           
           
		                
            </ul>   


        <!-- Lager -->
       <ul class="menupkt_n2" id="ul_menu-slider-6" style="display:none; visibility:visible;">
                <h3 class="menuh3"><%=tsa_txt_472 %></h3>
            
                
	        <li><a href='<%=toSubVerPath14 %>materialer.asp'><%=tsa_txt_478 %></a></li>

           <%if level <= 2 OR level = 6 then %>
            <li><a href='<%=toSubVerPath14 %>materialer.asp?func=opret'><%=tsa_txt_473 %></a></li>

           <li><a href='<%=toSubVerPath14 %>materiale_grp.asp'><%=tsa_txt_474 %></a></li>
           
           <li><a href='<%=toSubVerPath14 %>materialer_ordrer.asp'><%=tsa_txt_475 %></a></li>

           <li><a href="<%=toSubVerPath14 %>produktioner.asp"><%=tsa_txt_476 %></a></li>

           <li><a href="<%=toSubVerPath14 %>leverand.asp"><%=replace(tsa_txt_477, "|", "&") %></a></li>
           <%end if %>

           

           
           
		                
            </ul>   

     


      <!-- Administration -->
       <ul class="menupkt_n2" id="ul_menu-slider-8" style="display:none; visibility:visible; height:1000px;">
      <h3 class="menuh3"><%=tsa_txt_479 %></h3>
           <%if lto = "xintranet - local" or lto = "gd" then
             else     
           %>
           <h3 class="menuh3"><%=tsa_txt_480 %></h3>
             <li><a href="<%=toSubVerPath14 %>abonner.asp"><%=tsa_txt_480 %></a></li>

           <h3 class="menuh3"><%=tsa_txt_448 %></h3>

                  <li><a href="<%=toSubVerPath15 %>projektgrupper.asp"><%=tsa_txt_481 %></a></li>
                  <li><a href="<%=toSubVerPath15 %>akt_gruppe.asp?menu=job&func=favorit"><%=tsa_txt_482 %></a></li>
                  <li><a href="<%=toSubVerPath15 %>ulev_gruppe.asp?menu=job&func=favorit"><%=tsa_txt_483 %></a></li>
                  <li><a href="<%=toSubVerPath15 %>fomr.asp"><%=tsa_txt_460 %></a></li>
                  <li><a href="<%=toSubVerPath15 %>milepale_typer.asp"><%=tsa_txt_484 %></a></li>
                                
                  <li><a href="<%=toSubVerPath15 %>stfolder_gruppe.asp?ketype=e"><%=tsa_txt_485 %></a></li>

                  <li><a href="<%=toSubVerPath15 %>job_movecustomer_multiple.asp">Flyt job til ny kunde</a></li>
                   <li><a href="<%=toSubVerPath15 %>akt_movejob_multiple.asp">Flyt aktivitet til nyt job</a></li> 

                    


           <%end if %>

           <h3 class="menuh3"><%=tsa_txt_486 %></h3>

                    <li><a href="<%=toSubVerPath15 %>kundetyper.asp?ketype=e"><%=tsa_txt_487 %></a></li>

                     <%
        
			if cint(crmOnOff) = 1 then
                %>
          

           
                        <li><a href="<%=toSubVerPath15 %>crmstatus.asp?menu=tok&ketype=e"><%=tsa_txt_488 %></a></li>
                            <li><a href="<%=toSubVerPath14 %>crmemne.asp?menu=tok&ketype=e"><%=tsa_txt_489 %></a></li>
                                <li><a href="<%=toSubVerPath14 %>crmkontaktform.asp?menu=tok&ketype=e"><%=tsa_txt_490 %></a></li>
                                   
                                             

           <%end if %>

           <%if lto = "gd" then


           else     
           %>
           <h3 class="menuh3"><%=tsa_txt_492 %></h3>
           
		
		<li><a href="<%=toSubVerPath14 %>javascript:NewWin_help('momskoder.asp');" target="_self"><%=tsa_txt_493 %></a></li>
		<li><a href="<%=toSubVerPath14 %>javascript:NewWin_help('nogletalskoder.asp');" target="_self"><%=tsa_txt_494 %></a></li>
		<li><a href="<%=toSubVerPath14 %>javascript:NewWin_help('momsafsluttet.asp');" target="_self"><%=tsa_txt_495 %></a></li>

              <%
            '**** INDLÆS / IMPORTER  funktioner     
            if (lto = "oko" OR lto = "intranet - local") AND level = 1 then %>

            <li><a href="../timereg_net/importer_job.aspx?lto=<%=lto%>&mid=<%=session("mid")%>&editor=<%=session("user") %>" target="_blank">NAV-importer sag/job</a></li>
            <li><a href="../timereg_net/importer_akt.aspx?lto=<%=lto%>&mid=<%=session("mid")%>&editor=<%=session("user") %>" target="_blank">NAV-importer sagslinje/aktivitet</a></li>

           <%end if


             '**** INDLÆS / IMPORTER  funktioner     
            if (lto = "wilke" OR lto = "intranet - local") AND level = 1 then %>

            <li><a href="../timereg_net/importer_job.aspx?lto=<%=lto%>&mid=<%=session("mid")%>&editor=<%=session("user") %>" target="_blank">C5-importer Job (sagslinjer & aktiviteter)</a></li>
           
           <%end if


            '**** INDLÆS / IMPORTER  funktioner     
            if (lto = "dencker" OR lto = "dencker_test" OR lto = "intranet - local") AND level = 1 then %>
             <li><a href="../timereg_net/importer_job.aspx?lto=<%=lto%>&mid=<%=session("mid")%>&editor=<%=session("user") %>&importtype=d1" target="_blank">Monitor-importer job</a></li>
            <li><a href="../timereg_net/importer_job_monitor.aspx?lto=<%=lto%>&mid=<%=session("mid")%>&editor=<%=session("user") %>" target="_blank">Monitor-importer aktiviteter</a></li>
           
           <%end if


             '**** INDLÆS / IMPORTER  funktioner     
            if (lto = "tia" OR lto = "intranet - local") AND level = 1 then %>
            
             <li><a href="../timereg_net/importer_job.aspx?lto=<%=lto%>&mid=<%=session("mid")%>&editor=<%=session("user") %>&importtype=t1" target="_blank">NAV-importer job/proj.</a></li>
            <li><a href="../timereg_net/importer_akt.aspx?lto=<%=lto%>&mid=<%=session("mid")%>&editor=<%=session("user") %>&importtype=t2" target="_blank">NAV-importer tasks/aktiviteter</a></li>
             <li><a href="../timereg_net/importer_med.aspx?lto=<%=lto%>&mid=<%=session("mid")%>&editor=<%=session("user") %>&importtype=t3" target="_blank">NAV-importer medarbejdere</a></li>
           
           <%end if


               '**** INDLÆS / IMPORTER  funktioner     
            if (lto = "epi2017") AND level = 1 then %>
            
            <li><a href="../timereg_net/importer_timer.aspx?lto=<%=lto%>&mid=<%=session("mid")%>&editor=<%=session("user") %>" target="_blank">CSV Import timer</a></li>
            <li><a href="../timereg_net/importer_akt.aspx?lto=<%=lto%>&mid=<%=session("mid")%>&editor=<%=session("user") %>&importtype=e2" target="_blank">CSV  Import omkostninger TEST</a></li>
            <li><a href="../timereg_net/importer_med.aspx?lto=<%=lto%>&mid=<%=session("mid")%>&editor=<%=session("user") %>&importtype=t2" target="_blank">CSV Import medarbejdere TEST</a></li>
           
           <%end if
                
            end if ' lto gd
            %>
	        


           <h3 class="menuh3"><%=tsa_txt_471 %></h3>

             <li><a href='<%=toSubVerPath15 %>bgrupper.asp'><%=tsa_txt_496 %></a></li>
            <li><a href='<%=toSubVerPath15 %>medarbtyper.asp'><%=tsa_txt_497 %></a></li>
            <li><a href="<%=toSubVerPath14 %>medarbtyper_grp.asp"><%=tsa_txt_498 %></a></li>

          


           <% if cint(dsksOnOff) = 1 then %>
           <h3 class="menuh3"><%=tsa_txt_499 %></h3>

           
            <li><a href="<%=toSubVerPath15 %>sdsk_prioitet.asp?menu=tok" target="_blank"><%=tsa_txt_500 %></a></li> 
            <li><a href="<%=toSubVerPath15 %>sdsk_prio_typ.asp?menu=tok" target="_blank"><%=tsa_txt_501 %></a></li>
            <li><a href="<%=toSubVerPath15 %>sdsk_prio_grp.asp?menu=tok" target="_blank"><%=tsa_txt_502 %></a></li>
            <li><a href="<%=toSubVerPath15 %>sdsk_status.asp?menu=tok"><%=tsa_txt_503 %></a></li>
            
            <li><a href="<%=toSubVerPath15 %>sdsk_typer.asp?menu=tok" target="_blank"><%=tsa_txt_504 %></a>
            <%
            end if
            %>

        </ul>   



       <!-- ERP -->
       <ul class="menupkt_n2" id="ul_menu-slider-7" style="display:none; visibility:visible;">
                <h3 class="menuh3"><%=tsa_txt_505 %></h3>
            
                <%select case lto 
                case "bf"
                 case else %>
	        <li><a href='<%=toSubVerPath14 %>erp_tilfakturering.asp'><%=tsa_txt_506 %></a></li>
            <%end select %>

           <li><a href='<%=toSubVerPath14 %>erp_opr_faktura_fs.asp?reset=0'><%=tsa_txt_507 %></a></li>
           <li><a href='<%=toSubVerPath14 %>erp_fakhist.asp'><%=tsa_txt_508 %></a></li>
           
             <%select case lto 
                case "bf"
                 case else %>
           <li><a href='<%=toSubVerPath14 %>erp_fakturaer_find.asp'><%=tsa_txt_509 %></a></li>

           <h3 class="menuh3"><%=tsa_txt_510 %></h3>
           <li><a href="<%=toSubVerPath14 %>erp_job_afstem.asp?menu=erp&show=joblog_afstem"><%=tsa_txt_511 %></a></li>
           <li><a href='<%=toSubVerPath14 %>erp_serviceaft_saldo.asp?menu=erp'><%=tsa_txt_512 %></a></li>

           <h3 class="menuh3"><%=tsa_txt_513 %></h3>

           <li><a href="<%=toSubVerPath14 %>kontoplan.asp?nolist=1"><%=tsa_txt_514 %></a></li>

           <li><a href="<%=toSubVerPath14 %>kontoplan.asp?menu=kon&func=opret"><%=tsa_txt_515 %></a></li>
           

		<li><a href="<%=toSubVerPath14 %>posteringer.asp?id=0"><%=tsa_txt_516 %></a></li>
	    <li><a href="<%=toSubVerPath14 %>posteringer.asp?id=0&kontonr=-1&kid=0"><%=tsa_txt_517 %></a></li>
		<li><a href="<%=toSubVerPath14 %>posteringer.asp?id=0&kontonr=-2&kid=0"><%=tsa_txt_518 %></a></li>
           <li><a href="<%=toSubVerPath14 %>resultatop.asp"><%=tsa_txt_519 %></a></li>
		
		
           <h3 class="menuh3"><%=tsa_txt_520 %></h3>
           <li><a href="<%=toSubVerPath14 %>budget_aar_dato.asp"><%=tsa_txt_521 %></a></li>

        <%if level <= 2 OR level = 6 then  %>
        <li><a href="<%=toSubVerPath14 %>budget_bruttonetto.asp"><%=replace(tsa_txt_522, "|", "&") %></a></li>
        <li><a href="<%=toSubVerPath14 %>budget_medarb.asp?menu=erp"><%=tsa_txt_523 %></a></li>
            
           <%call timesimon_fn() %>

            <%if cint(timesimon) = 1 then %>
            <li><a href="<%=toSubVerPath15 %>timbudgetsim.asp">Timebudget simulering</a></li>
            <%end if %>

            <li><a href="<%=toSubVerPath15 %>budget_firapport.asp">Budget & Finansiel rap.</a></li>

            <%'end if %>
        <%end if %>

             <%end select %>


            

         

		                
            </ul>   
 

         <!-- SDSK -->
       <ul class="menupkt_n2" id="ul_menu-slider-9" style="display:none; visibility:visible;">
                <h3 class="menuh3"><%=tsa_txt_524 %></h3>
            
                
	        <li><a href='<%=toSubVerPath14 %>sdsk.asp'><%=tsa_txt_524 %></a></li>
           <li><a href='<%=toSubVerPath14 %>sdsk.asp?func=opr&FM_kontakt=0'><%=tsa_txt_525 %></a></li>
           
            <li><a href='<%=toSubVerPath14 %>sdsk_stat.asp?menu=sdsk'><%=tsa_txt_526 %></a></li>
           <li><a href="<%=toSubVerPath14 %>sdsk_knowledge.asp?menu=sdsk&visikke=1"><%=tsa_txt_527 %></a></li>

              <li><a href="<%=toSubVerPath14 %>filer.asp"><%=tsa_txt_528 %></a></li>


         
           
		                
            </ul>   
 
        <ul><br /><br /><br /><br /><br /><br /><br /><br /><br />
            <br /><br /><br /><br /><br /><br /><br /><br /><br /><li>...</li>&nbsp;</ul>

         

        </nav>
        <!--<nav class="menu-slider menu-slider-left" id="menu-slider-1">-->
        
        <!--
        <nav>
            <h3 class="menuh3">Tidsregistrering</h3>
            <summary>I denne sektion har vi samlet alt vedr. Timeregistrering. Du har her mulighed for at registrere timer eller få et overblik over dine nuværende registreringer</summary>
            <ul id="ul_menu-slider-1">
                <li><a href="<%=toSubVerPath14 %>timereg_akt_2006.asp">Timeregistrering</a></li>
                <li><a href="<%=toSubVerPath14 %>index.html">Ugeplan</a></li>
                <li><a href="<%=toSubVerPath14 %>index.html">Afstemning</a></li>
            </ul>   
        </nav>
            -->

        
     <%if cint(tregmenu_2014) <> 0 AND (lto = "demo" OR lto = "intranet - local") then 
         
         if thisfile <> "home_dashboard_demo.asp" then
         %>
        <div style="position:absolute; top:50px; left:1102px; padding:5px; background-color:#ffd800; z-index:9000000000000;">
        
        <a href="<%=toSubVerPath15 %>home_dashboard_demo.asp"><span class="glyph icon-home"></span> Tilbage til start</a> </div>

    <%
        end if 
        end if %>

  
<%end function %>


  
   