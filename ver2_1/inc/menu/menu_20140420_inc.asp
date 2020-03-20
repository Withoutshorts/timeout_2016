<%
level = session("rettigheder")

public lnkTimeregside, lnkUgeseddel, lnkAfstem, lnkLogind, lnkFavorit


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
           

            call vis_resplanner_fn()
            call vis_favorit_fn()
            call vis_projektgodkend_fn()
         
            call vis_lager_fn()

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
    lagemenu_2014 = 1 'vis_lager2
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


    kundebudget = 1



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
    lagemenu_2014 = 1
    statmenu_2014 = 1

case "tia", "xintranet - local"
    tregmenu_2014 = 1
    projmenu_2014 = 1
    marbmenu_2014 = 1
    kundmenu_2014 = 1

    if level = 7 then
    statmenu_2014 = 0
    else
    statmenu_2014 = 1
    end if
    
    if level = 1 then
    admimenu_2014 = 1
    end if

    lagemenu_2014 = 0


case "alfanordic", "xintranet - local"

    tregmenu_2014 = 1
    statmenu_2014 = 2
    marbmenu_2014 = 1
    
    if level = 1 then
    projmenu_2014 = 1
    kundmenu_2014 = 1
   
    else
    projmenu_2014 = 0
    kundmenu_2014 = 0
    
    end if
    
    helpmenu_2014 = 0
    erpfmenu_2014 = 0
   

    if level = 1 then
    admimenu_2014 = 1
    else
    admimenu_2014 = 0
    end if

    lagemenu_2014 = vis_lager2
    
 case "wap"

    tregmenu_2014 = 1

    if level <> 1 then
        statmenu_2014 = 2
        marbmenu_2014 = 1
    else
        statmenu_2014 = 1
        marbmenu_2014 = 1
        projmenu_2014 = 1
        kundmenu_2014 = 1
        lagemenu_2014 = vis_lager2
        admimenu_2014 = 1
        helpmenu_2014 = 1
    end if

case "ddc"

    tregmenu_2014 = 1

    if (level <= 3 OR level = 6) then
    projmenu_2014 = 1
    else
    projmenu_2014 = 0
    end if

    statmenu_2014 = 1

    if level = 1 then
    kundmenu_2014 = 1
    else
    kundmenu_2014 = 0
    end if

    marbmenu_2014 = 1
    lagemenu_2014 = vis_lager2
   
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

case "a27"


    tregmenu_2014 = 2
    marbmenu_2014 = 1
    if level <= 2 OR level = 6 then
        admimenu_2014 = 2
    end if

    if level = 1 then
    projmenu_2014 = 0
    else
    projmenu_2014 = 0
    end if

    if level = 1 then
    kundmenu_2014 = 0
    else
    kundmenu_2014 = 0
    end if

    statmenu_2014 = 0


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
    
    if level <= 2 OR level = 6 then
    lagemenu_2014 = vis_lager2
    else
    lagemenu_2014 = 0
    end if
    
    
    if level = 1 OR lto = "demo" then
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

    kundebudget = 0

    if lto = "lm" then
        rejseafregingmenu = 1
    end if

end select

  
%>

    <style>
        .news:hover {
            background-color:inherit !important;
        }
    </style>


    <div class="fixed-navbar-hoz" style="z-index:300000;">

          
        <h1 class="menu_logo-left">
            <a href="<%=toSubVerPath14 %>../login.asp"></a>
        </h1>
       

          <%if cint(slip_smiley_agg_lukper) <> 1 then %>
            <nav class="dropdown-right">

                  

            <%if testDB = "1" then %>
            <ul><li><br /><br /><br /><br /><span style="color:red;">DEMO / TEST data: (<%=lto %>)</span></li></ul>
            <%end if %>


            <% 
            'Henter startside
            startsideint = 0
            startsidelink = ""
            strSQL = "SELECT tsacrm FROM medarbejdere WHERE mid = "& session("mid")
            oRec.open strSQL, oConn, 3
            if not oRec.EOF then
                startsideint = oRec("tsacrm")
            end if
            oRec.close

            
            select case cint(startsideint)
                case 1
                    if lto = "essens" OR lto = "fkfk" then
					startsidelink = "timereg/crmhistorik.asp?menu=crm&ketype=e&func=hist&selpkt=hist"
					else
					startsidelink = "timereg/crmkalender.asp?menu=crm&shokselector=1&ketype=e&selpkt=kal"
					end if
                case 2
					startsidelink = "timereg/ressource_belaeg_jbpla.asp?menu=webblik"
                case 3
                    call sonmaniuge(now)
                    if lto = "xtec" OR lto = "xcst" OR lto = "xesn" then
                    startsidelink = "timereg/logindhist_2011.asp?usemrn="&session("mid")&"&varTjDatoUS_man="&mandagIuge&"&varTjDatoUS_son="&sondagIuge
                    else
                    startsidelink = "to_2015/logindhist_2011.asp?usemrn="&session("mid")&"&varTjDatoUS_man="&mandagIuge&"&varTjDatoUS_son="&sondagIuge        
                    end if
                case 4
                    startsidelink = "timereg/webblik_joblisten.asp?menu=webblik&fromvmenu=1"
                case 5
                    if lto = "nt" OR lto = "intranet - local" then
                    startsidelink = "to_2015/job_nt.asp"
                    else
					startsidelink = "to_2015/job_list.asp"
                    end if 
                case 6
                    call sonmaniuge(now)
                    startsidelink = "to_2015/ugeseddel_2011.asp?usemrn="&session("mid")&"&varTjDatoUS_man="&mandagIuge&"&varTjDatoUS_son="&sondagIuge
                case 7 ' DASHBOARD
					startsidelink = "to_2015/medarbdashboard.asp"
                case 8
                    startsidelink = "to_2015/kunder.asp?visikkekunder=1"
                case 9
                    call sonmaniuge(now)
					startsidelink = "to_2015/favorit.asp?varTjDatoUS_man="&mandagIuge
                case 10
                    startsidelink = "to_2015/materialerudlaeg.asp"
                case else '0
                    startsidelink = "timereg/timereg_akt_2006.asp"

            end select
            %>

            <%if startsidelink <> "" then %>
            <ul>
                <li><a href="https://timeout.cloud/timeout_xp/wwwroot/<%=toVer %>/<%=startsidelink %>"><span style="font-size:165%;" class="glyph icon-home"></span></a></li>
            </ul>
            <%end if %>

            <ul>
                <% 
                strConnect_admin = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154;Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_admin;" 
                Set oConn_admin = Server.CreateObject("ADODB.Connection")
                Set oRec_admin = Server.CreateObject ("ADODB.Recordset")

                oConn_admin.open strConnect_admin

                sqlNow = year(now) &"-"& month(now) &"-"& day(now)

                'Finder antal nyheder
                antalNews = 0
                strSQL = "SELECT count(id) as antal FROM systemmeddelelser WHERE ('"& sqlNow &"' >= visfra AND '"& sqlNow &"' <= vistil)"
                oRec_admin.open strSQL, oConn_admin, 3
                if not oRec_admin.EOF then
                    antalNews = oRec_admin("antal")
                end if
                oRec_admin.Close

                strSQL = "SELECT id, overskrift, besked, editdate, visfra FROM systemmeddelelser WHERE ('"& sqlNow &"' >= visfra AND '"& sqlNow &"' <= vistil)"
                'sresponse.Write strSQL
                oRec_admin.open strSQL, oConn_admin, 3
                antalnye = 0
                an = 0
                strmeddelelser = ""
                systemedids = ""
                while not oRec_admin.EOF
                    an = an + 1
                    
                    'strmeddelelser = strmeddelelser & "<li class='news' style='text-align:center; width:250px; height:auto; padding-top:5px; padding-bottom:5px;'><span style='cursor:default !important; color:black;'><b>"& oRec_admin("visfra") &"</b><br>"& oRec_admin("besked") &"</span>"

                    strmeddelelser = strmeddelelser & "<li class='news' style='background-color:transparent !important; padding:15px 15px 0px 15px; height:auto; color:#2c2c2c'><span style='font-size:80%;'><i>"& oRec_admin("visfra") &"</i></span><br><b>"& oRec_admin("overskrift") &"</b><br><span>"& oRec_admin("besked") &"</span></li>"

                    strmeddelelser = strmeddelelser & "<br>"

                    if cint(an) = cint(antalNews) then
                        'strmeddelelser = strmeddelelser & "<br><br> <span style='font-size:150%;' class='fa fa-cog'></span> <br>"
                    end if

                    'strmeddelelser = strmeddelelser & "</li>"
                    
                    'Tjekker om den loggede ind allerede har set nyheden
                    erLaest = 0
                    strSQLSYSRead = "SELECT id FROM systemmeddelelser_rel WHERE medid = "& session("mid") & " AND sysmed = "& oRec_admin("id")
                    oRec.open strSQLSYSRead, oConn, 3
                    if not oRec.EOF then
                        erLaest = 1
                    end if
                    oRec.close

                    'if request.Cookies("systemmeddelelse_" & oRec_admin("id") & "_laest") <> "" then
                        'erLaest = 1
                    'end if

                    if erLaest = 0 then
                        antalnye = antalnye + 1
                        if antalnye = 1 then
                            systemedids = oRec_admin("id")
                        else
                            systemedids = systemedids & "," & oRec_admin("id")
                        end if
                    end if

                oRec_admin.MoveNext
                wend
                oRec_admin.Close

                oConn_admin.Close


                if antalnye = 0 then
                    antalnye = ""
                end if

                %>
                <input type="hidden" id="systemnewsids" value="<%=systemedids %>" />

                <li id="systemeddelelser">
                    <a href="#" id="showNews"><span style="font-size:155%;" class="glyph icon-info"></span><i style="color:#FF5252; font-size:12px; position:absolute; top:15px;"><b id="antalnyemeddelelser"><%=antalnye %></b></i></a>
                    <!-- background-color:white; border:1px solid #d9d9d9; border-radius:25px; margin-top:5px;" --> 
                    <ul id="listofnews" style="all:inherit; background-color:#f2f2f2; position:absolute; border:1px solid #d9d9d9; border-radius:25px; margin-top:5px; height:auto; width:250px; display:none;">
                        <%                 
                        if strmeddelelser <> "" then

                            response.Write strmeddelelser

                            'response.Write "<li style=visiblity:hidden;></li>"
                            'response.Write "<li class='news' style='text-align:center;'><span class='fa fa-times' style='color:darkred;'></span></li>"
                            'response.Write "<li id='newsread' style='text-align:center; cursor:pointer; padding-top:15px; color:black; background-color:#e3e3e3;'><b>Markér som læst</b></li>"
                        else
                            'response.Write "<li class='news' style='background-color:transparent !important;'><a style='cursor:default !important; color:#2c2c2c; text-align:center;'><b>Timeout er oppe</b><br><span style='font-size:9px;'>2020-01-29 10:00</span></a></li>"
                            'response.Write "<li class='news' style='background-color:transparent !important; border-top:1px solid #d9d9d9;'><a style='cursor:default !important; color:#2c2c2c; text-align:center;'><b>Timeout er nede</b><br><span style='font-size:9px;'>2020-01-29 09:45</span></a></li>"
                            %>
                            <li class='news' style='background-color:transparent !important; padding-top:5px; text-align:center; height:auto;'><div style='all:inherit; cursor:default !important; color:#2c2c2c; text-align:center; display:inline-block; white-space:pre-line;'><b><%=menu_txt_016 %></b><br>&nbsp;</div></li>
                            <%
                        end if              
                                                    
                        %>                           

                        <li class='news' style='background-color:transparent !important; padding-top:5px; text-align:center; height:auto; padding-bottom:5px;'><a style='all:inherit; cursor:default !important; color:#2c2c2c; text-align:center; display:inline-block; white-space:pre-line;'><span id="newsread" class="btn btn-default btn-sm" style="font-size:100%;">OK</span></a></li>
                      
                        

                    </ul>

                </li>
            </ul>


         <ul> 



                <li>
                    <a href="#"><span class="glyph icon-user"></span><span class="account-name"><%=session("user") %></span><b class="caret-down"></b></a>
                    <ul>


                       

                        
                        <li><a href="<%=toSubVerPath15 %>help_faq.asp" target="_blank">Hj&aelig;lp / Help / FAQ</a></li>
                        <%if lto <> "tec" AND lto <> "esn" then %>
                        <!--<li><a href="https://www.islonline.net/start/ISLLightClient"><%=tsa_txt_433 %></a></li>-->
                        <li><a href="https://download.teamviewer.com/download/TeamViewerQS.exe"><%=tsa_txt_433 %></a></li>
                        <%end if %>

                        <%if level = 1 then %>
                        <li><a href="<%=toSubVerPath14 %>kontrolpanel.asp"><%=tsa_txt_434 %></a></li>
                        <%end if %>
                        <%
                        
                        
                        'call erStempelurOn()
           
                        'if cint(stempelur_hideloginOn) = 1 then    
                            
                        if session("stempelur") <> 0 AND lto <> "tec" AND lto <> "esn" AND lto <> "cflow" then %>
                        <li style="background-color:red;"><a href="<%=toSubVerPath14 %>stempelur.asp?func=redloginhist&medarbSel=<%=session("mid")%>&showonlyone=1&hidemenu=1&id=0&rdir=sesaba" target="_top"><%=tsa_txt_435 %></a></li>
                        <%else %>
                        <li style="background-color:red;"><a href="<%=toSubVerPath14 %>../sesaba.asp?fromweblogud=1" target="_top"><%=tsa_txt_435 %></a></li>
                        <%end if %>
                        <li style="background-color:#999999;"><a href="#"><%=tsa_txt_436 %><br /><%=lto %></a></li>
                    </ul>
                </li>
            </ul>
        </nav>
       <%end if %>
    </div>
    <div class="fixed-navbar-vert" style="z-index:200000;">
        <nav>
            <ul>
                <%if cint(tregmenu_2014) <> 0 AND (lto = "demo" OR lto = "intranet - local") then %>
                <li id="menu_0" class="showLeft tooltip-right" tooltip-title="Home"><a href="<%=toSubVerPath15 %>home_dashboard_demo.asp"><span class="glyph icon-home"></span></a></li>
                <%end if %>

                <%if cint(tregmenu_2014) <> 0 then %>
                <li id="menu_1" class="showLeft tooltip-right" tooltip-title="<%=menu_txt_001 %>"><a href="#"><span class="glyph icon-stopwatch"></span></a></li><!-- stopwatch -->
                <%end if %>

                <%if cint(rejseafregingmenu) = 1 then %>
                <li id="menu_10" class="showLeft tooltip-right" tooltip-title="Rejseafregning"><a href="#"><span class="glyph icon-suitcase"></span></a></li><!-- stopwatch -->
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


 <nav id="menu-slider" class="menu-slider" style="overflow-y:scroll; z-index:100000;">

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

                 toSubVerPath15_422 = "../../ver4_22/"& replace(toSubVerPath15, "../", "")
                 toSubVerPath15_399 = "../../ver3_99/"& replace(toSubVerPath15, "../", "")

                if len(trim(usemrn)) <> 0 then
                    usemrn = usemrn
                else
                    usemrn = session("mid")
                end if

                'lnkTreg = "timereg_akt_2006.asp?showakt=1&hideallbut_first=2&strdag="&day(varTjDatoUS_man)&"&strmrd="& month(varTjDatoUS_man) &"&straar="&year(varTjDatoUS_man)
                lnkTimeregside = "timereg_akt_2006.asp?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&strdag="&strdagTreg&"&strMrd="&strMrdTreg&"&strAar="&strAarTreg
                lnkUgeseddel = "ugeseddel_2011.asp?usemrn="&usemrn&"&varTjDatoUS_man="&stDato_ugeseddel_kommegaa_US_man'&"&varTjDatoUS_son="&varTjDatoUS_son
                lnkUgetast = "ugetast.asp?varTjDatoUS_man="&stDato_ugeseddel_kommegaa_US_man
	            lnkAfstem = "afstem_tot.asp?usemrn="&usemrn&"&show=5&varTjDatoUS_man="&varTjDatoUS_man'&"&varTjDatoUS_son="&varTjDatoUS_son
                
                lnkLogind = "logindhist_2011.asp?usemrn="&usemrn&"&varTjDatoUS_man="&stDato_ugeseddel_kommegaa_US_man'&"&varTjDatoUS_son="&varTjDatoUS_son 
                lnkFavorit = "favorit.asp?FM_medid="&usemrn&"&usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man
                    

                '**** Timereg siden
     
                    
                 select case lto
                 case "a27"
                    %>
                         <li><a href="../to_workout/workout.asp">Workout <%=ucase(lto) %></a></li>
                    <%
                 case "ddc", "cflow", "care", "foa", "kongeaa"
                 case "cst"
                    %>   <li><a href="<%=toSubVerPath14 %><%=lnkTimeregside %>"><%=tsa_txt_116 &" "& tsa_txt_438 %></a></li><%
                 case "tec", "esn"
                       
                        %>   
                        <li><a href="<%=toSubVerPath14 %><%=lnkTimeregside %>"><%=tsa_txt_116 &" "& tsa_txt_438 %></a></li><%
                 
                case "xtia", "welcom", "alfanordic", "wap", "tbg"

                            if lto = "tbg" AND (session("mid") = 1 OR session("mid") = 3) then
                            %><li><a href="<%=toSubVerPath14 %><%=lnkTimeregside %>"><%=tsa_txt_116 %></a></li><%
                            end if 
                                  
                case "flash"
                                %>
                    <li><a href="<%=toSubVerPath14 %><%=lnkTimeregside %>"><%=tsa_txt_116 %></a></li>
                    <li><a href="<%=toSubVerPath15 %><%=lnkUgetast%>"><%=menu_txt_015 %></a></li>
                    <%

                 case else
                    %>
                    <li><a href="<%=toSubVerPath14 %><%=lnkTimeregside %>"><%=tsa_txt_116 %></a></li>
                  
                    <%
                 end select
                 %>
               



                <%'**** Ugeseddel
                    select case lto 

                    case "wap", "a27"

                        ''***** ingen ugeseddel ******''

                      case "cflow"

                        call medariprogrpFn(session("mid"))
                        if instr(medariprogrpTxt, "#14#") <> 0 OR instr(medariprogrpTxt, "#16#") <> 0 OR instr(medariprogrpTxt, "#3#") <> 0 OR instr(medariprogrpTxt, "#19#") <> 0 then
                        %>
                          <li><a href="<%=toSubVerPath15 %><%=lnkUgeseddel%>"><%=tsa_txt_337 %></a></li>
                        <%
                        end if

                    case else
                %>
                        <li><a href="<%=toSubVerPath15 %><%=lnkUgeseddel%>"><%=tsa_txt_337 %></a></li>
                       
                      
                <%
                    end select 'Ugeseddel
                    
                    
                    
                'select case lto
                 'case "outz", "intranet - local", "hidalgo", "tia", "dencker", "eniga", "welcom", "mmmi", "epi2017"
                    if cint(vis_favorit) = 1 then
                    
                    select case lto 
                    case "bf"

                        if level = 1 then
                        %>
                             <li><a href="<%=toSubVerPath15 %>favorit.asp?FM_medid=<%=usemrn %>&varTjDatoUS_man=<%=varTjDatoUS_man %>"><%=favorit_txt_001 %></a></li>
                        <%
                        end if  
                            
                   
                    case "cflow"

                        'call medariprogrpFn(session("mid"))
                        if instr(medariprogrpTxt, "#14#") <> 0 OR instr(medariprogrpTxt, "#16#") <> 0 OR instr(medariprogrpTxt, "#3#") <> 0 OR instr(medariprogrpTxt, "#19#") <> 0 OR instr(medariprogrpTxt, "#21#") <> 0 then

                        %>
                             <li><a href="<%=toSubVerPath15 %>favorit.asp?FM_medid=<%=usemrn %>&varTjDatoUS_man=<%=varTjDatoUS_man %>"><%=favorit_txt_001 %></a></li>
                        <%
                        end if  

                    case else%>
                    <li><a href="<%=toSubVerPath15 %>favorit.asp?FM_medid=<%=usemrn %>&varTjDatoUS_man=<%=varTjDatoUS_man %>"><%=favorit_txt_001 %></a></li>
                 <% end select
                     
                   end if 'vis_favorit 

                 

                 '*** Komme / Gå / Stempelur
                 select case lto
                 case "alfanordic", "a27"                   

                 case else
                 if cint(stempelurOn) = 1 then %>

                    <%select case lto 'ENS, CST, TEC venter til 1.8 2018
                    case "xcst", "xesn", "xtec", "intranet - local"%>
                    <li><a href="<%=toSubVerPath14 %><%=lnkLogind%>"><%=tsa_txt_340 %></a></li>
                    <%case else %>
                    <li><a href="<%=toSubVerPath15 %><%=lnkLogind%>"><%=tsa_txt_340 %></a></li>
                    <%end select %>

                     <%if lto = "cflow" OR lto = "miele" or lto = "outz" or lto = "cool" OR lto = "hidalgo" OR lto = "kongeaa" then%>

                            <%
                            'select case lto
                              '    case "cflow"
                              '        ltokey = "9K2017-1124-TO179"
                              '    case "outz"
                              '        ltokey = "2.152-1604-B001"
                              '    case "cool"
                              '        ltokey = "9K2019-2702-TO187"
                              'end select

                              ltokey = strLicenskey 'session("lto")
                              %>





                              <li><a href="<%=toSubVerPath15 %>monitor.asp?func=startside&ltokey=<%=ltokey %>" target="_blank">Terminal ind/ud stempling</a></li>



                            <%if level = 1 OR session("mid") = 81 then %>
                            <li><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/to_2015/monitor.asp?func=startside&ltokey=<%=ltokey %>" target="_blank" style="color:orange">NY! Terminal ind/ud stempling m. besøk</a></li>
                            <%end if %>
                     <%end if

              




                 end if
                 end select 
                 %>

                
                 <% select case lto
                
                 case "tec", "esn", "epi", "epi_uk", "epi_no", "epi_as", "tia", "xintranet - local", "wap", "ddc", "cflow", "a27", "kongeaa"

                 case "foa", "care"
                    if level <= 2 OR level = 6 then
                        %><li><a href="<%=toSubVerPath14 %>materialer_indtast.asp?id=0&fromsdsk=0&aftid=0"><%=tsa_txt_191 %></a></li><%
                    end if
                 case else
                    
                     if level <= 7 then %>
                     <li><a href="<%=toSubVerPath14 %>materialer_indtast.asp?id=0&fromsdsk=0&aftid=0"><%=tsa_txt_191 %></a></li>
	                 <%end if 
                 end select%>
                 
                  <% 
                   
                  select case lto
                  case "alfanordic", "xintranet - local", "a27", "kongeaa"
                    
                  case else       
                      call traveldietexp_fn()
                      
                      if cint(traveldietexp_on) = 1 OR (lto = "tia" AND (session("mid") = 9 OR session("mid") = 1)) then
                      %>
                      <li><a href="<%=toSubVerPath15 %>traveldietexp.asp"><%=diet_txt_014 %></a></li>
                      <%end if
                  end select    
                  %>
                
                 <%
                '** Afstemning
                 select case lto 
                 case "alfanordic", "xintranet - local", "a27"



                 case else
                 %>
                 <li><a href="<%=toSubVerPath14 %><%=lnkAfstem %>"><%=tsa_txt_389 %></a></li>
                 <%end select %>

                 <%select case lto
                 case "oko", "epi", "epi2017", "wilke", "outz", "dencker", "essens", "synergi1", "jttek", "hidalgo", "demo", "bf", "plan", "acc", "assurator", "glad", "tia", "eniga", "welcom", "alfanordic", "intranet - local", "cflow", "cisu", "sduuas", "adra", "ddc", "foa", "care", "sducei"
                     
                        if level = 1 OR (lto = "wilke") OR (lto = "outz") OR (lto = "dencker") OR (lto = "hidalgo") OR (lto = "acc") OR (lto = "epi2017") OR (lto = "assurator") OR (lto = "glad") _
                        or (lto = "tia") OR (lto = "eniga") OR (lto = "welcom") OR (lto = "alfanordic") OR (lto = "intranet - local") OR (lto = "cisu") OR (lto = "sduuas") OR (lto = "sducei") OR (lto = "adra") OR (lto = "ddc") OR (lto = "sdumikjoh") OR (lto = "care" AND (level <= 2 OR level = 6)) OR (lto = "foa" AND (level <= 2 OR level = 6)) then%>
                        <li><a href="<%=toSubVerPath15 %>medarbdashboard.asp"><%=tsa_txt_529 %></a></li>
                        <%end if 
                        
                        'if (level <= 2 OR level = 6) AND (lto = "cflow") then
                            %>
                           <!--  <li><a href="<%=toSubVerPath15 %>medarbdashboard.asp"><%=tsa_txt_529 %></a></li> -->
                            <%
                        'end if
                        

                 end select %>
                

            
                   
                   <%if level <= 2 OR level = 6 then %> 
                    
                       <%
                       '** INGEN Projektleder funktioner  
                       if (lto = "tia" AND level = 7) OR (lto = "alfanordic") OR lto = ("intranet - local") OR lto = ("wap") OR lto = ("cflow") OR lto = ("a27") OR lto = ("kongeaa") then  
                       
                       else    
                       %>

                       <% 
                       '*** OVERSKRIFT projektleder funktioner    
                       %>

                        <h3 class="menuh3"><%=tsa_txt_439 %></h3>


                            <%select case lto
                            case "outz", "intranet - local"
                            '**** Kanban **
                            %>
                            <li><a href="https://timeout.cloud/timeout_xp/wwwroot/ver4_22/to_2015/kanban2.asp">Kanban</a></li>
                            <%end select%>


                           <%'** Igangværende arbejde ***'
                            select case lto
                
                            case "tec", "esn" %>
                            <%case else 
                            
                                    if level <= 2 OR level = 6 then%>
                                   <li><a href="<%=toSubVerPath15 %>webblik_joblisten.asp"><%=tsa_txt_452 %></a></li>
                            <%      end if
                            end select %>
                           
                            
                      <% 
                     '**** Ressource forecast    
                         select case lto
                         case "tec", "esn", "nonstop", "epi2017", "ddc", "xoko"
                         case "oko"
                                 if (level <= 2 OR level = 6) AND (session("mid") = 1 OR session("mid") = 42 OR session("mid") = 6) then 'AND cint(vis_resplanner) = 1 %>
                                <li><a href="<%=toSubVerPath15 %>ressource_belaeg_jbpla.asp"><%=tsa_txt_440 %></a></li>
                                 <%end if

                         case else
                    
                                 if (level <= 2 OR level = 6) then 'AND cint(vis_resplanner) = 1 %>
                                <li><a href="<%=toSubVerPath15 %>ressource_belaeg_jbpla.asp"><%=tsa_txt_440 %></a></li>
                                 <%end if%>

                        <%end select %>

                 

                     <%end if %>


	               
                  <%end if 'level%>
                

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
                   '** Level 1 har adgang til sygdom OGSå TEC og ESN da der ikke tjekkes på medarbejderlinej niveau i feriekalender
                   '** Ellers kun ferie  
                   if lto = "alfanordic" OR lto = "intranet - local" OR lto = "ddc" OR lto = "cflow" OR lto = "a27" OR lto = "kongeaa" then
                      
                   else 
                       if level = 1 then %>

	                        <li><a href="<%=toSubVerPath15 %>feriekalender.asp?menu=job"><%=replace(tsa_txt_443, "|", "&") %></a></li>
		           
                 
                      <%else %>

                             <%select case lto
                
                             case "tec", "esn", "care", "foa"

                                      if level <= 2 OR level = 6 then
                                      %>
                                        <li><a href="<%=toSubVerPath15 %>feriekalender.asp?menu=job"><%=replace(tsa_txt_443, "|", "&") %></a></li>
                                      <%end if

                             case "nonstop"

                             case "tia", "xintranet - local", "wap"

                             case else
                    
                              %>

		                         <li><a href="<%=toSubVerPath14 %>feriekalender.asp?menu=job"><%=replace(tsa_txt_444, "|", "&") %></a></li>

                            <%end select %>

		         
                        <%end if %>
                    <%end if %>


                 <%


            call stadeOn()

            if jobasnvigv = 1 then %>
            <li><a href="<%=toSubVerPath14 %>stat_opdater_igv.asp?func=opdater" target="_blank"><%=menu_txt_013%></a>
            <%end if 

                    


             'select case lto 
                 
                     'case "tia", "intranet - local", "outz", "oliver", "hidalgo", "demo"
              
                
                       if cint(vis_projektgodkend) = 1 then 
                       %>
                      <li><a href="<%=toSubVerPath15%>godkend_job_timer_2017.asp"><%=godkend_txt_010 %></a></li>
                       <% end if
             'end select   



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
       <li><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2014-1007-TO155" >SDUTEK</a></li>
                 
       <%
        'UNIK '2.2011-0410-TO127   
        end if
       


      if lto = "sdutek" then
       %>
       <li><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/timereg/timereg_akt_2006.asp?eksterntlnk=aaQWEIOC345345DFNEfjsdf7890sdfv&email=<%=meEmail%>&key=2.2010-0311-TO119" >MMMI</a></li>
       <%
       end if
       
         
    end if 'skift version
     %>
                
                                              
            </ul>   


            <!-- Rejse afregning -->
            <%if cint(rejseafregingmenu) = 1 then %>
            <ul class="menupkt_n2" id="ul_menu-slider-10" style="display:none; visibility:visible;">
                <h3 class="menuh3" style="font-size:16px; font-weight:200;">Rejseafregning</h3> 

                <li><a href="<%=toSubVerPath15 %>materialerudlaeg.asp">Udlæg</a></li>

                <%call traveldietexp_fn()
                      
                if cint(traveldietexp_on) = 1 OR (lto = "tia" AND (session("mid") = 9 OR session("mid") = 1)) then %> 
                     <li><a href="<%=toSubVerPath15 %>traveldietexp.asp"><%=diet_txt_014 %></a></li>
                <%end if %>

                <li><a href="<%=toSubVerPath15 %>stat_korsel.asp?menu=stat" ><%=tsa_txt_265 %></a></li>

            </ul>
            <%end if %>



            <!-- PROJEKTETER -->
              <ul class="menupkt_n2" id="ul_menu-slider-2" style="display:none; visibility:visible;">
               
                <!--<summary>I denne sektion har vi samlet alt vedr. Projekter.</summary>-->
              
                <%
                '*** NT

                    

                if cint(projmenu_2014) = 2 then%>
                     <h3 class="menuh3"><%=tsa_txt_445 %></h3>
                      <li><a href="<%=toSubVerPath15 %>job_nt.asp?rapporttype=0"><%=tsa_txt_445 %></a></li>
                      <li><a href="<%=toSubVerPath15 %>job_nt.asp?func=opret&id=0&rapporttype=0"><%=tsa_txt_446 %></a></li>

                      <%if level <= 2 OR level = 6 then %>
               
                     <li><a href="<%=toSubVerPath14 %>filer.asp"><%=tsa_txt_451 %></a></li>
                     <%end if %>
                    
                 <%
                 '** Standard    
                 else %>
                 <h3 class="menuh3"><%=tsa_txt_447 %></h3>
                <li><a href="<%=toSubVerPath15 %>job_list.asp?hidelist=2"><%=tsa_txt_448 %></a></li>
                

              

                 <%'** Opret job ****'%>
                 <%if lto <> "care" then %>
                 <li><a href="<%=toSubVerPath14 %>jobs.asp?func=opret&id=0"><%=tsa_txt_449 %></a></li>
                  <%end if %>

                  <%if lto = "outz" OR lto = "demo" OR lto = "intranet - local" OR lto = "hidalgo" OR lto = "essens" OR lto = "care" then %>
                  <li><a href="<%=toSubVerPath15 %>jobs.asp?func=opret&id=0"><%=tsa_txt_449 %> simpel</a></li>

                                 <%if lto = "care" AND level = 1 then %>
                                 <li><a href="<%=toSubVerPath14 %>jobs.asp?func=opret&id=0"><%=tsa_txt_449 %> (Old)</a></li>
                                  <%end if %>

                  <%end if %>
              

                  <%select case lto 
                    case "tia", "xintranet - local", "ddc", "kongeaa"
                      
                    case else 
                  %>
                  <%if level <= 2 OR level = 6 then %>
                  <li><a href="<%=toSubVerPath14 %>job_print.asp?menu=job&kid=0&id=0"><%=tsa_txt_450 %></a></li>
                  <li><a href="<%=toSubVerPath14 %>filer.asp"><%=tsa_txt_451 %></a></li>
                  <%end if %>
                  <%end select %>




                    <%if lto <> "kongeaa" then %>
                        <%if level <= 2 OR level = 6 then %>
                              <h3 class="menuh3"><%=tsa_txt_439 %></h3>
                   

                          
                              <li><a href="<%=toSubVerPath15 %>webblik_joblisten.asp"><%=tsa_txt_452 %></a></li>

                             <%select case lto
                             case "xintranet - local", "xoko"
                                 %>
                                <li><a href="<%=toSubVerPath15_399 %>webblik_joblisten.asp" style="color:orange">Beta <%=tsa_txt_452 %></a></li> <!-- toSubVerPath15_399 -->
                                <%
                             end select%>
                 

                            <%select case lto
                             case "oko"

                             case else
                              %>
                              <li><a href="<%=toSubVerPath15 %>ressource_belaeg_jbpla.asp"><%=tsa_txt_440 %>

                                   <%if lto = "wwf" then%>
                          
                                  (<%=tsa_txt_453 %>)
                          
                                  <%end if %>

                                  </a></li>

                        
                             <%end select%>
                 
                     


                             
             
                                   <%select case lto 

                                     case "tia", "intranet - local", "outz", "oliver", "hidalgo", "demo", "ddc"
                                       %>
                                      <li><a href="<%=toSubVerPath15%>godkend_job_timer_2017.asp"><%=godkend_txt_010 %></a></li>
                                       <%
                                     case else   

                                            %> 
                                           <li><a href="<%=toSubVerPath14 %>webblik_joblisten21.asp"><%=tsa_txt_454 %></a></li>
                                           <li><a href="<%=toSubVerPath14 %>webblik_milepale.asp"><%=tsa_txt_455 %></a></li>

                                          <%if (level = 1 OR session("mid") = 35) then 'Kim B epinion %>
                                          <li><a href="<%=toSubVerPath14 %>pipeline.asp?menu=webblik&FM_kunde=0&FM_progrupper=10"><%=tsa_txt_456 %></a></li>
                                          <%end if %>

                                  <%end select %>

                                      <%if lto = "outz" OR lto = "intranet - local" then %>
                                        <li><a href="<%=toSubVerPath14 %>jbpla_w.asp">Planlægningskalender (TEST)</a></li>
                                      <%end if %>

                                         <%'if lto = "outz" OR lto = "intranet - local" OR lto = "essens" OR lto = "hidalgo" then 
                                       if cint(vis_resplanner) = 1 then%>
                                   
                                         <form id="ressourceplanner_sub" action="http://timeout.cloud/timeout_xp/wwwroot/ver4_22/ressource_planner/ressplan_2017.aspx?" method="post" target="_blank">
                                            <input type="hidden" name="FM_lto" value="<%=lto %>" />
                                            <li><a id="ressourceplanner" href="#" >Ressource Planner</a></li>
                                        </form>

                                        <script type="text/javascript">
                                            $("#ressourceplanner").click(function () {
                                                $("#ressourceplanner_sub").submit();
                                            });
                                        </script>


                                            <!--<a href="http://timeout.cloud/timeout_xp/wwwroot/ver4_22/ressource_planner/ressplan_2017.aspx?lto=234fsdf45t9xxx4cc34vdg56<%=lto %>HrtKvv8344" target="_blank">Ressource Planner</a>-->

                                            <!--<a href="../ressource_planner/ressplan_2017.aspx?lto=234fsdf45t9xxx4cc34vdg56<%=lto %>HrtKvv8344" target="_blank">Ressource Planner</a>-->
                                      <%end if 'vis_resplanner %>

                               <%end if 'level
                                  

                          select case lto 
                          case "tia", "care", "kongeaa"
                          case else
                           %>
                          <li><a href="<%=toSubVerPath15 %>eval_liste.asp">Project Evaluation</a></li>
                          <%end select %>


                      <%end if ' kongeaa%>
                  <%end if 'NT -standard menu %>
            </ul>   




             <!-- STATISTIK / RAPPORTER -->
             <ul class="menupkt_n2" id="ul_menu-slider-3" style="display:none; visibility:visible;">
                <h3 class="menuh3"><%=tsa_txt_457 %></h3>
            
                
                     <%if cint(statmenu_2014) = 2 then 'nt <> 2%>
                
                           <%if lto = "alfanordic" OR lto = "intranet - local" OR lto = "wap" then %>

                                <%if lto <> "wap" then %>        
                                    <li><a href="<%=toSubVerPath15 %>medarb_protid.asp"><%=menu_txt_010 %></a></li>                        

                                
                                    <li><a href="<%=toSubVerPath15 %>stat_korsel.asp?menu=stat" ><%=tsa_txt_265 %></a></li>

                                    <%if level = 2 OR level = 6 then %>
                                    <li><a href="<%=toSubVerPath14 %>joblog_timetotaler.asp"><%=tsa_txt_458 %></a></li>
                                    <li><a href="<%=toSubVerPath14 %>joblog.asp"><%=tsa_txt_118 %></a></li>
                                    <li><a href="<%=toSubVerPath14 %>bal_real_norm_2007.asp?dontdisplayresult=1"><%=replace(tsa_txt_466, "|", "&") %></a></li>
                                    <%end if %>


                                <%else %>
                        
                                <!-- hvis WAP skal de have grandtotal og afspadsering -->
                                <li><a href="<%=toSubVerPath14 %>joblog_timetotaler.asp"><%=tsa_txt_458 %></a></li>
		                        <li><a href="<%=toSubVerPath15 %>feriekalender.asp?menu=job"><%=replace(tsa_txt_444, "|", "&") %></a></li>


		                            

                                <%end if %>

                           <%else 'NT %>

                                <%if level = 1 AND (session("mid") = 3 OR session("mid") = 6 OR session("mid") = 1) then %>               
                                    <li><a href="<%=toSubVerPath14 %>oms.asp"><%=tsa_txt_459 %></a></li>

                                <%end if %>
                                        
                                <%if level = 1 AND (session("mid") = 3 OR session("mid") = 6 OR session("mid") = 1 OR session("mid") = 4) then %> 

                                    <li><a href="<%=toSubVerPath15 %>materialerudlaeg.asp">Expences</a></li>
                           
                                <%end if %>

                           <%end if %>
                          <!--
                         <li><a href="<%=toSubVerPath14 %>job_nt.asp?rapporttype=1">Production / Enq. Overview</a></li>
                         <li><a href="<%=toSubVerPath14 %>job_nt.asp?rapporttype=3">Overview (ext.)</a></li>
                         <li><a href="<%=toSubVerPath14 %>job_nt.asp?rapporttype=4">Overview</a></li>
                         <li><a href="<%=toSubVerPath14 %>oms.asp">Sales Budget</a></li> 
                         -->

                 <%else 'Standard Rapport menu%>

                         <br /><br />
                         <!--<li><a href="../to_2015/progtotal_medarb.asp">Test</a></li> -->

                         <%select case lto
                            case "kongeaa"

                            case "care", "foa", "plan" %>

                            <%if level <= 2 OR level = 6 then %>
                             <li><a href="<%=toSubVerPath14 %>joblog_timetotaler.asp"><%=tsa_txt_458 %></a></li>
                             <li><a href="<%=toSubVerPath14 %>joblog.asp"><%=tsa_txt_118 %></a></li>
                            <%end if %>
                    
                        <%case else %>

                            <li><a href="<%=toSubVerPath14 %>joblog_timetotaler.asp"><%=tsa_txt_458 %></a></li>
                            <li><a href="<%=toSubVerPath14 %>joblog.asp"><%=tsa_txt_118 %></a></li>

                        <%end select %>
                 

                        <%if lto <> "kongeaa" then %>
                             <%if level = 1 then %>
                 
                                         <li><a href="<%=toSubVerPath14 %>oms.asp"><%=tsa_txt_459 %></a></li>
                                         <li><a href="<%=toSubVerPath14 %>fomr.asp?func=stat"><%=tsa_txt_460 %></a></li>
                                
                            <%end if %>
                        <%end if %>


                        <%if lto <> "kongeaa" then %>
                            <%if (level <= 2 OR level = 6) OR lto = "fk" OR lto = "mmmi" OR lto = "adra" OR lto = "care" OR lto = "foa" then %>

                                        <%if lto = "care" OR lto = "foa" then %>

                                            <%if level <= 2 OR level = 6 then %>
                                                <li><a href="<%=toSubVerPath15 %>medarb_protid.asp"><%=menu_txt_014 %></a></li>
                                                <li><a href="<%=toSubVerPath15 %>medarb_protid.asp?version=1"><%=menu_txt_010 %></a></li>
                                            <%else %>
                                                <li><a href="<%=toSubVerPath15 %>medarb_protid.asp?version=1"><%=menu_txt_010 %></a></li>
                                                <li><a href="<%=toSubVerPath15 %>medarb_protid.asp"><%=menu_txt_014 %></a></li>
                                            <%end if %>   
                 
                                        <%else %>

                                        <li><a href="<%=toSubVerPath15 %>medarb_protid.asp"><%=menu_txt_014 %></a></li>
                                        <li><a href="<%=toSubVerPath15 %>medarb_protid.asp?version=1"><%=menu_txt_010 %></a></li>

                                        <%end if %>

                            <%end if %>
                        <%end if %>


                    <%if lto <> "kongeaa" then %>
                       <%if (lto = "epi2017" AND (level = 1 OR session("mid") = 2694) ) then
                           %>
                              <li><a href="<%=toSubVerPath14 %>stempelur.asp?func=stat"><%=tsa_txt_463 %></a></li>
                       <%end if%>
                   <%end if %>


                 <%if lto <> "kongeaa" then %>
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
                <%end if %>


                         <%if level = 1 OR ((level <= 2 OR level = 6) AND lto = "epi2017") OR (lto = "plan" AND session("mid") = 274) then %>

                            <li><a href="<%=toSubVerPath14 %>bal_real_norm_2007.asp?dontdisplayresult=1"><%=replace(tsa_txt_466, "|", "&") %></a></li>
                         <%end if %>
                
                                <%if lto = "care" or lto = "foa" then %>

                                    <%if level <= 2 OR level = 6 then %>                   
                                        <%if level = 1 then %>
	                                    <li><a href="<%=toSubVerPath15 %>feriekalender.asp?menu=job"><%=replace(tsa_txt_443, "|", "&") %></a></li>
		                                <%else %>
		                                <li><a href="<%=toSubVerPath15 %>feriekalender.asp?menu=job"><%=replace(tsa_txt_444, "|", "&") %></a></li>
		                                <%end if 
                                    end if %>

                                <%else %>

                                    <%if level = 1 then %>
	                                <li><a href="<%=toSubVerPath15 %>feriekalender.asp?menu=job"><%=replace(tsa_txt_443, "|", "&") %></a></li>
		                            <%else %>
		                            <li><a href="<%=toSubVerPath15 %>feriekalender.asp?menu=job"><%=replace(tsa_txt_444, "|", "&") %></a></li>
		                            <%end if
                                        
                                end if
                                
                         %>



                         <%if cint(smilaktiv) <> 0 AND level <= 2 OR level = 6 then  %>
                        <li><a href="<%=toSubVerPath14 %>week_real_norm_2010.asp?menu=stat"><%=tsa_txt_356 %></a></li>
                        <%end if %>



                        <%if level <= 2 OR level = 6 then '** Indtil teamleder er impl. på kørsels siden **"%>

                        <%'if session("stempelur") <> 0 OR lto = "hidalgo" then %>
                        <li><a href="<%=toSubVerPath15 %>godkend_request_timer.asp?FM_medarb=<%=session("mid")%>">Godkend Ferie, Overtid & Løntimer</a></li>
                        <%'end if %>

                        <%if lto = "wap" then %>
                        <li><a href="<%=toSubVerPath15 %>godkend_request_timer.asp?FM_medarb=<%=session("mid")%>" target="_blank">Godkend forespurgte timer</a></li>
                        <%end if %>
                        
                                <%select case lto
                                 case "ddc", "kongeaa"
                            
                                 case else%>
                                <li><a href="<%=toSubVerPath15 %>stat_korsel.asp?menu=stat" ><%=tsa_txt_265 %></a></li>
                                <li><a href="<%=toSubVerPath14 %>pipeline.asp?menu=webblik&FM_kunde=0&FM_progrupper=10"><%=tsa_txt_456 %></a></li>
                                <%end select %>

                        <%end if %>

                        <%if lto <> "kongeaa" then %>
                            <%if level = 1 OR lto = "ddc" then %>
                            <li><a href="<%=toSubVerPath15%>forecast_kapacitet.asp?menu=stat">Forecast & kapacitet</a></li>
                            <li><a href="<%=toSubVerPath15%>forecast_kapacitet_graf.asp?menu=stat">Forecast & kapacitet - Grafisk</a></li>
                            <%end if %>
                        <%end if %>

                        <br /><br />


                        <%if lto <> "kongeaa" then %>
                            <%if level = 1 AND cint(lagemenu_2014) = 1 OR lto = "immenso" then %>
                            <li><a href="<%=toSubVerPath14 %>materiale_stat.asp?menu=stat"><%=tsa_txt_156 %></a></li>
                            <%end if %>

                    
                                <%if level <= 2 OR level = 6 then 
                                %>
                                <li><a href="<%=toSubVerPath15 %>materialerudlaeg.asp">Udlæg</a></li>
                                <%end if%>
                        

                            <br /><br />
                        
                            <%if level = 1 AND (lto = "intranet - local" OR lto = "epi" OR lto = "epi_cati" OR lto = "epi_no" OR lto = "epi_se" OR lto = "epi2017" OR lto = "tia") then %>
                            <li><a href="<%=toSubVerPath14 %>timerimperr.asp?menu=stat" ><%=tsa_txt_467 %></a></li>
                            <%end if %>



                            <%if level = 1 AND (lto = "intranet - local" OR lto = "epi" OR lto = "epi_cati" OR lto = "epi_no" OR lto = "epi_se" OR lto = "epi2017") then %>
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
                            <li><a href="<%=toSubVerPath14 %>stat_opdater_igv.asp?menu=stat"><%=menu_txt_013 %></a></li>
                            <%end if %>
                        <%end if %>


                    <% 

                     'if (lto = "oko") OR lto = "intranet - local" then 'jobresume 
                        
                        
                         select case lto
                         case "bf", "kongeaa"
                        
                         case else
                        
                         if (level <= 2 OR level = 6) OR lto = "oko" then%>
                           <li><a href="<%=toSubVerPath14 %>jobprintoverblik.asp?menu=job&id=0" target="_blank">Joboverblik / Resume</a></li>
                        <%end if %>

                         <%end select %>

                    <%'end if %>


                    <%if lto <> "foa" AND lto <> "care" then %>

                         <h3 class="menuh3">Your Reports</h3>
                        <%
                            yr = 0

                            'if yr = 1000 then
                            strSQLyourRap = "SELECT rap_mid, rap_navn, rap_url, rap_criteria, rap_dato, rap_editor FROM your_rapports WHERE rap_mid = " & session("mid") & " OR rap_mid = 0 ORDER BY rap_navn LIMIT 10"
                            oRec6.open strSQLyourRap, oConn, 3
                            while not oRec6.EOF 
                        
                            %>
                            <li><a href="<%=toSubVerPath14 %><%=oRec6("rap_url") & oRec6("rap_criteria")%>" target="_blank" class="a_yourrep"><%=oRec6("rap_navn") %></a></li>
                            <%
                        
                            yr = yr + 1
                            oRec6.movenext
                            wend
                            oRec6.close

                            'end if
                        
                        if yr <> 0 then%> 
                        <li><a href="yourrep.asp" target="_blank" style="color:#999999;">Go to "Your Reports"</a></li>
                        <%else %>
                            <li><a href="#" style="color:#999999;">- None</a></li>
                        <%end if %>

                    <%end if 'foa, care%>

                 <%end if %>

            </ul>   


        <!-- Kunder -->
       <ul class="menupkt_n2" id="ul_menu-slider-4" style="display:none; visibility:visible;">
                <h3 class="menuh3"><%=global_txt_124 %></h3>
            
                
	        <li><a href='<%=toSubVerPath15 %>kunder.asp?visikkekunder=1'><%=global_txt_124 %></a></li>
            <li><a href='<%=toSubVerPath15 %>kunder.asp?func=opret&ketype=k'><%=global_txt_180 %></a></li>
           
           <%if cint(kundebudget) = 1 OR lto = "outz" then %>
            <li><a href='<%=toSubVerPath15 %>kunder_budget.asp?'><%=kunder_txt_119 %></a></li>
           <%end if %>

           <%if lto <> "kongeaa" then %>
            <li><a href='<%=toSubVerPath14 %>serviceaft_osigt.asp?id=0&func=osigtall'><%=global_txt_181 %></a></li>       
            <li><a href='<%=toSubVerPath15 %>kontaktpers.asp?func=list'><%=global_txt_182 %></a></li>
           <%end if %>

           
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
                <%if lto <> "a27" then %>
                <h3 class="menuh3"><%=tsa_txt_471 %></h3>
                <%else %>
                <h3 class="menuh3"><%=global_txt_125 %></h3>
                <%end if %>

            <%
            select case lto 
            case "alfanordic"       
            %>
               <li><a href='<%=toSubVerPath15 %>medarb.asp?menu=medarb&func=red&id=<%=session("mid")%>'><%=global_txt_184 %></a></li>
            <%
            case else
            %> 
       
           
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

           <%end select %>
          

           <% if session("stempelur") <> 0 then %>

                <%if lto = "foa" AND (level > 2 AND level <> 6) then %> 
                
                <%else %>

                    <li><a href='<%=toSubVerPath15 %>infoscreen.asp' target="_blank">Infoskærm (hvem er ikke mødt)</a></li>

                    <%if (level = 1 AND lto = "cflow") OR lto = "foa" then%>
                
                    <li><a href='<%=toSubVerPath15 %>infoscreen.asp' target="_blank" style="color:orange;">Infoskærm (m. SMS alarm)</a></li>
                   
                    <%end if %>

                    <%if lto = "cool" then %>
                        <li><a href='<%=toSubVerPath15 %>infoscreen.asp?frokostvisning=1' target="_blank">Infoskærm - Kantine</a></li>
                    <%end if %>

                <%end if %>


           <%end if %>
           
           
		                
            </ul>   


        <!-- Lager -->

       <%
           if cint(lagemenu_2014) <> 0 then
            %>
            <ul class="menupkt_n2" id="ul_menu-slider-6" style="display:none; visibility:visible;">
            <h3 class="menuh3"><%=tsa_txt_472 %></h3>
            
                
	        <li><a href='<%=toSubVerPath15 %>materialer.asp'><%=tsa_txt_478 %></a></li>

                   <%if level <= 2 OR level = 6 then %>
                    <li><a href='<%=toSubVerPath15 %>materialer.asp?func=opret'><%=tsa_txt_473 %></a></li>

                   <li><a href='<%=toSubVerPath15 %>materiale_grp.asp'><%=tsa_txt_474 %></a></li>
           
                   <li><a href='<%=toSubVerPath14 %>materialer_ordrer.asp'><%=tsa_txt_475 %></a></li>

                   <li><a href="<%=toSubVerPath14 %>produktioner.asp"><%=tsa_txt_476 %></a></li>

                   <li><a href="<%=toSubVerPath15 %>leverand.asp"><%=replace(tsa_txt_477, "|", "&") %></a></li>
                   <%end if %>

            </ul>   
            <%end if %>
     


      <!-- Administration -->
       <ul class="menupkt_n2" id="ul_menu-slider-8" style="display:none; visibility:visible; height:1000px;">
      <h3 class="menuh3"><%=tsa_txt_479 %></h3>


           <%if admimenu_2014 = 2 then 'Trainer log %>

                <li><a href='../to_workout/aktiviteter.asp'>Activities</a></li>
                <li><a href="<%=toSubVerPath15 %>projektgrupper.asp"><%=progrp_txt_040 %></a></li>
                <li><a href="<%=toSubVerPath15 %>fomr.asp"><%=fomr_txt_001 %></a></li>
                <li><a href='<%=toSubVerPath15 %>medarbtyper.asp'><%=medarbtyp_txt_006 %></a></li>

           <%else %>


           <%if lto = "xintranet - local" or lto = "gd" then
             else     
           %>
           <h3 class="menuh3"><%=tsa_txt_480 %></h3>
             <li><a href="<%=toSubVerPath15 %>abonner.asp"><%=tsa_txt_480 %></a></li>

            <%if lto = "outz" AND level = 1 then %>
             <li><a href="<%=toSubVerPath15 %>systemmeddelelser.asp">System meddelelser</a></li>
            <%end if %>

            <%if (session("stempelur") <> 0 AND lto = "dencker" OR lto = "adra" OR lto = "outz" OR lto = "cool" OR lto = "hidalgo" OR lto = "kongeaa") then %>             
                <li><a href="<%=toSubVerPath15 %>infoscreen_news.asp">Nyheder</a></li>
                <li><a href="<%=toSubVerPath15 %>udeafhuset.asp">Udeafhuset</a></li>
            <%end if %>

           <h3 class="menuh3"><%=tsa_txt_448 %></h3>

                  <li><a href="<%=toSubVerPath15 %>projektgrupper.asp"><%=tsa_txt_481 %></a></li>
                  <li><a href="<%=toSubVerPath15 %>akt_gruppe.asp?menu=job&func=favorit"><%=tsa_txt_482 %></a></li>
                  <li><a href="<%=toSubVerPath15 %>ulev_gruppe.asp?menu=job&func=favorit"><%=tsa_txt_483 %></a></li>
                  <li><a href="<%=toSubVerPath15 %>fomr.asp"><%=tsa_txt_460 %></a></li>
                  <li><a href="<%=toSubVerPath15 %>milepale_typer.asp"><%=tsa_txt_484 %></a></li>
                                
                  <li><a href="<%=toSubVerPath15 %>stfolder_gruppe.asp?ketype=e"><%=tsa_txt_485 %></a></li>

                  <li><a href="<%=toSubVerPath15 %>job_movecustomer_multiple.asp">Flyt job til ny kunde</a></li>
                   <li><a href="<%=toSubVerPath15 %>akt_movejob_multiple.asp">Flyt aktivitet/timer til nyt job</a></li> 

                   <li><a href="<%=toSubVerPath15 %>selectboxoptions.asp" target="_blank">Selectbox Options</a></li>
                         
                    
           
           <%end if %>

           <h3 class="menuh3"><%=tsa_txt_486 %></h3>

                    <li><a href="<%=toSubVerPath15 %>kundetyper.asp?ketype=e"><%=tsa_txt_487 %></a></li>

                     <%
        
			if cint(crmOnOff) = 1 then
                %>
          

           
                        <li><a href="<%=toSubVerPath15 %>crmstatus.asp?menu=tok&ketype=e"><%=tsa_txt_488 %></a></li>
                            <li><a href="<%=toSubVerPath15 %>crmemne.asp?menu=tok&ketype=e"><%=tsa_txt_489 %></a></li>
                                <li><a href="<%=toSubVerPath15 %>crmkontaktform.asp?menu=tok&ketype=e"><%=tsa_txt_490 %></a></li>
                                   
                                             

           <%end if %>

           <%if lto = "gd" then


           else     
           %>
           <h3 class="menuh3"><%=tsa_txt_492 %></h3>
           
		
		<li><a href="<%=toSubVerPath15 %>momskoder.asp" target="_blank"><%=tsa_txt_493 %></a></li>
		<li><a href="<%=toSubVerPath14 %>nogletalskoder.asp" target="_blank"><%=tsa_txt_494 %></a></li>
		<li><a href="<%=toSubVerPath14 %>momsafsluttet.asp" target="_blank"><%=tsa_txt_495 %></a></li>

               <li><a href="<%=toSubVerPath15 %>travel_diet_tariff.asp">Rejse og Diæt Tariffer</a></li>

              <%
            '**** INDLaeS / IMPORTER  funktioner     
            if (lto = "oko" OR lto = "intranet - local") AND level = 1 then %>

            <li><a href="../timereg_net/importer_job.aspx?lto=<%=lto%>&mid=<%=session("mid")%>&editor=<%=session("user") %>" target="_blank">NAV-importer sag/job</a></li>
            <li><a href="../timereg_net/importer_akt.aspx?lto=<%=lto%>&mid=<%=session("mid")%>&editor=<%=session("user") %>" target="_blank">NAV-importer sagslinje/aktivitet</a></li>

           <%end if


             '**** INDLaeS / IMPORTER  funktioner     
            if (lto = "wilke" OR lto = "intranet - local") AND level = 1 then %>

            <li><a href="../timereg_net/importer_job.aspx?lto=<%=lto%>&mid=<%=session("mid")%>&editor=<%=session("user") %>" target="_blank">C5-importer Job (sagslinjer & aktiviteter)</a></li>
           
           <%end if

            '**** INDLaeS / IMPORTER  funktioner  Primavera   
            if (lto = "cflow" OR lto = "intranet - local") AND level = 1 then %>

            <li><a href="../timereg_net/importer_akt.aspx?lto=<%=lto%>&mid=<%=session("mid")%>&editor=<%=session("user") %>&importtype=cflow_prima" target="_blank">Primavera importer proj. & aktiviteter</a></li>
            
           <%end if


            '**** INDLaeS / IMPORTER  funktioner     
            if (lto = "dencker" OR lto = "dencker_test" OR lto = "intranet - local") AND level = 1 then %>
             <li><a href="../timereg_net/importer_job.aspx?lto=<%=lto%>&mid=<%=session("mid")%>&editor=<%=session("user") %>&importtype=d1" target="_blank">Monitor-importer job (salgsordre)</a></li>
            <li><a href="../timereg_net/importer_job_monitor.aspx?lto=<%=lto%>&mid=<%=session("mid")%>&editor=<%=session("user") %>" target="_blank">Monitor-importer aktiviteter</a></li>
           
           <%end if


             '**** INDLaeS / IMPORTER  funktioner     
            if (lto = "tia" OR lto = "intranet - local") AND level = 1 then %>
            
             <li><a href="../timereg_net/importer_job.aspx?lto=<%=lto%>&mid=<%=session("mid")%>&editor=<%=session("user") %>&importtype=t1" target="_blank">NAV-importer job/proj.</a></li>
            <li><a href="../timereg_net/importer_akt.aspx?lto=<%=lto%>&mid=<%=session("mid")%>&editor=<%=session("user") %>&importtype=t2" target="_blank">NAV-importer tasks/aktiviteter</a></li>
             <li><a href="../timereg_net/importer_med.aspx?lto=<%=lto%>&mid=<%=session("mid")%>&editor=<%=session("user") %>&importtype=m3" target="_blank">NAV-importer medarbejdere</a></li>
           
           <%end if


               '**** INDLaeS / IMPORTER  funktioner     
            if (lto = "epi2017") AND level = 1 then %>
            
            <li><a href="../timereg_net/importer_timer.aspx?lto=<%=lto%>&mid=<%=session("mid")%>&editor=<%=session("user") %>" target="_blank">CSV Import timer</a></li>
            <li><a href="../timereg_net/importer_omk.aspx?lto=<%=lto%>&mid=<%=session("mid")%>&editor=<%=session("user") %>&importtype=exp1" target="_blank">CSV  Import omkostninger</a></li>
            <li><a href="../timereg_net/importer_med.aspx?lto=<%=lto%>&mid=<%=session("mid")%>&editor=<%=session("user") %>&importtype=t2" target="_blank">CSV Import medarbejdere BETA</a></li>

           <%end if

             '**** INDLaeS / IMPORTER  funktioner     
            if (lto = "epi2017" OR lto = "intranet - local") AND level = 1 then %>

            <li><a href="../timereg_net/importer_job_salgbudget_epi.aspx?lto=<%=lto%>&mid=<%=session("mid")%>&editor=<%=session("user") %>&importtype=episb" target="_blank">CSV Import Salgsbudget BETA</a></li>
          
           
           <%end if
                
            end if ' lto gd
            %>
	        


           <h3 class="menuh3"><%=tsa_txt_471 %></h3>

             <li><a href='<%=toSubVerPath15 %>bgrupper.asp'><%=tsa_txt_496 %></a></li>
            <li><a href='<%=toSubVerPath15 %>medarbtyper.asp'><%=tsa_txt_497 %></a></li>
            <li><a href="<%=toSubVerPath15 %>medarbtyper_grp.asp"><%=tsa_txt_498 %></a></li>

           <li><a href='<%=toSubVerPath15 %>ferietildel.asp'><%=menu_txt_012 %></a></li>

          


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


            <%end if 'adminmenu %>

        </ul>   



       <!-- ERP -->
       <ul class="menupkt_n2" id="ul_menu-slider-7" style="display:none; visibility:visible;">
                <h3 class="menuh3"><%=tsa_txt_505 %>

                    <%
                      call timesimon_fn() %>

                    <%if cint(timesimon) = 1 then %>
                    <br />& Budget Simulering.
                    <%end if %>

                </h3>
            
                <%select case lto 
                case "bf", "ddc"
                 case else %>
	        <li><a href='<%=toSubVerPath14 %>erp_tilfakturering.asp'><%=tsa_txt_506 %></a></li>
            <%end select %>


             <%select case lto 
                case "ddc"
                 case else %>

           <li><a href='<%=toSubVerPath14 %>erp_opr_faktura_fs.asp?reset=0'><%=tsa_txt_507 %></a></li>
           <li><a href='<%=toSubVerPath14 %>erp_fakhist.asp'><%=tsa_txt_508 %></a></li>
           <%end select %>
           
             <%select case lto 
            case "bf", "ddc"

                 call timesimon_fn() %>

                    <%if cint(timesimon) = 1 then %>
                    <li><a href="<%=toSubVerPath15 %>timbudgetsim.asp">Timebudget simulering (overblik)</a></li>
                   <li><a href="<%=toSubVerPath15 %>timbudgetsim.asp?func=forecast&jobid=-1" target="_blank">Timebudget simulering (forecast pr. medarb.)</a></li>
           
                    <%end if


            case else %>
                       <li><a href='<%=toSubVerPath14 %>erp_fakturaer_find.asp'><%=tsa_txt_509 %></a></li>

                       <h3 class="menuh3"><%=tsa_txt_510 %></h3>
                       <li><a href="<%=toSubVerPath14 %>erp_job_afstem.asp?menu=erp&show=joblog_afstem"><%=tsa_txt_511 %></a></li>
                       <li><a href='<%=toSubVerPath14 %>erp_serviceaft_saldo.asp?menu=erp'><%=tsa_txt_512 %></a></li>

                       <h3 class="menuh3"><%=tsa_txt_513 %></h3>

                       <li><a href="<%=toSubVerPath15 %>kontoplan.asp?nolist=1"><%=tsa_txt_514 %></a></li>

                       <li><a href="<%=toSubVerPath15 %>kontoplan.asp?menu=kon&func=opret"><%=tsa_txt_515 %></a></li>
           

		            <li><a href="<%=toSubVerPath15 %>posteringer.asp?id=0"><%=tsa_txt_516 %></a></li>
	                <li><a href="<%=toSubVerPath15 %>posteringer.asp?id=0&kontonr=-1&kid=0"><%=tsa_txt_517 %></a></li>
		            <li><a href="<%=toSubVerPath15 %>posteringer.asp?id=0&kontonr=-2&kid=0"><%=tsa_txt_518 %></a></li>
                       <li><a href="<%=toSubVerPath14 %>resultatop.asp"><%=tsa_txt_519 %></a></li>
		
		
                       <h3 class="menuh3"><%=tsa_txt_520 %></h3>
                       <li><a href="<%=toSubVerPath14 %>budget_aar_dato.asp"><%=tsa_txt_521 %></a></li>

                    <%if level <= 2 OR level = 6 then  %>
                    <li><a href="<%=toSubVerPath14 %>budget_bruttonetto.asp"><%=replace(tsa_txt_522, "|", "&") %></a></li>
                    <li><a href="<%=toSubVerPath14 %>budget_medarb.asp?menu=erp"><%=tsa_txt_523 %></a></li>
            
                       <%call timesimon_fn() %>

                        <%if cint(timesimon) = 1 then %>
                        <li><a href="<%=toSubVerPath15 %>timbudgetsim.asp">Timebudget simulering (overblik)</a></li>
                       <li><a href="<%=toSubVerPath15 %>timbudgetsim.asp?func=forecast&jobid=-1" target="_blank">Timebudget simulering (forecast pr. medarb.)</a></li>
           
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
 
        <ul><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
            <br /><br /><br /><br /><br /><br /><br /><br /><br /><li>...</li>&nbsp;</ul>

         

        </nav>
        <!--<nav class="menu-slider menu-slider-left" id="menu-slider-1">-->
        
        <!--
        <nav>
            <h3 class="menuh3">Tidsregistrering</h3>
            <summary>I denne sektion har vi samlet alt vedr. Timeregistrering. Du har her mulighed for at registrere timer eller fae et overblik over dine nuvaerende registreringer</summary>
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


  
