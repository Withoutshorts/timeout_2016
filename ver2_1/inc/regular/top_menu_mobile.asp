

<% function mobile_header %>


        <%select case thisfile
        case "ugeseddel_2011.asp", "favorit_mob", "logindhist_2018", "joblist", "jobs", "monitor"
        relpathTT = "../timetag_web/"
        relpathTO = ""
        case else 
        relpathTT = ""
        relpathTO = "../to_2015/"
        end select 
            
            
            
        ddDato = year(now) &"/"& month(now) &"/"& day(now)
        ddDato_ugedag = day(now) &"/"& month(now) &"/"& year(now)
        ddDato_ugedag_w = datepart("w", ddDato_ugedag, 2, 2)
            
        varTjDatoUS_man_tt = dateAdd("d", -(ddDato_ugedag_w-1), ddDato_ugedag)
            
        call TimeOutVersion()

        %>


<div id="header">

    <header class="navbar navbar-inverse" role="banner" style="margin-top:-30px;">
    

    <div class="container" style="background-color:transparent;">

    <a class="navbar-brand navbar-brand-img" style="padding-top:20px;">
        Timeout Mobile
    </a>

      <div class="navbar-header">
        <button class="navbar-toggle" type="button" data-toggle="collapse" data-target=".navbar-collapse">
          <span class="sr-only">Toggle navigation</span>
          <i class="fa fa-bars"></i>
        </button>
      </div> <!-- /.navbar-header -->
       

        <nav class="collapse navbar-collapse" role="navigation">

            <%if lto <> "hidalgo" AND lto <> "kongeaa" then %>
            <ul class="nav navbar-nav navbar-right mainnav-menu">

                <%if cint(mt_mobil_visstopur) <> 1 then %>
                
                        <%select case lto 
                         case "miele", "nt" 
                            
                         case else
                            
                            showTreg = 0
                            select case lto
                            case "cflow"
                            'call medariprogrpFn(session("mid"))
                            'if instr(medariprogrpTxt, "#14#") <> 0 OR instr(medariprogrpTxt, "#16#") <> 0 OR instr(medariprogrpTxt, "#3#") <> 0 OR instr(medariprogrpTxt, "#19#") <> 0 then 
                            'showUge_fav = 1
                            'end if
                            'case else
                            showTreg = 0
                            case else
                            showTreg = 1
                            end select

                            if cint(showTreg) = 1 then
                            %>
                            <li style="border-bottom:1px #000000 solid;">
                            <a href="<%=relpathTT%>timetag_web.asp"><%=ttw_txt_015 %> </a>
                            </li>
                            <%end if %>
                            
                            <%'mobile_menu_timereg_on_of 
                            showfav = 0    
                            select case lto
                            case "cflow"
                            call medariprogrpFn(session("mid"))
                                if instr(medariprogrpTxt, "#14#") <> 0 OR instr(medariprogrpTxt, "#16#") <> 0 OR instr(medariprogrpTxt, "#3#") <> 0 OR instr(medariprogrpTxt, "#19#") <> 0 OR level = 1 then 
                                showfav = 1
                                end if
                            case else
                            showfav = 1  
                            end select



                                call vis_favorit_fn()
                                %>
                            
                                <% if cint(vis_favorit) = 1 AND cint(showfav) = 1 then %>
                                <li style="border-bottom:1px #000000 solid;">
                               <!-- <a href="<%=relpathTO%>favorit.asp?FM_medid=<%=session("mid") %>">Favoritliste</a> -->
                                    <a href="<%=relpathTO%>favorit.asp"><%=favorit_txt_001 %></a>
                                </li>
                                <%end if %>

                           

                        <%end select



                                select case lto
                                case "miele"
                                         %>
                                  <li style="border-bottom:1px #000000 solid;">
                                    <a href="<%=relpathTT%>checkin.asp"><%=ttw_txt_015 %></a>
                                  </li>

                                <%
                                 case "hidalgo"
                                         %>
                                  <li style="border-bottom:1px #000000 solid;">
                                    <a href="<%=relpathTT%>checkin.asp">Stemple ind/ud</a>
                                  </li>

                                  <li style="border-bottom:1px #000000 solid;">
                                      <a href="<%=relpathTT%>timetag_web_kpi.asp"><%=ttw_txt_017 %></a>
                                    </li>

                                <%
                                case "cflow"
                                 if instr(medariprogrpTxt, "#14#") <> 0 OR instr(medariprogrpTxt, "#16#") <> 0 OR instr(medariprogrpTxt, "#3#") <> 0 OR instr(medariprogrpTxt, "#19#") <> 0 then 
                                %>
                                <li style="border-bottom:1px #000000 solid;">
                                <a href="<%=relpathTO%>ugeseddel_2011.asp?usemrn=<%=session("mid")%>&varTjDatoUS_man=<%=varTjDatoUS_man_tt %>"><%=tsa_txt_337 %></a>
                                </li> 
                                <%end if %>

                                <li style="border-bottom:1px #000000 solid;">
                                <a href="<%=relpathTO%>logindhist_2011.asp?usemrn=<%=session("mid")%>&varTjDatoUS_man=<%=varTjDatoUS_man_tt %>"><%=tsa_txt_340 %></a>
                                </li> 

                                <%
                                case "nt"
                    
                                case else%>

                                    <li style="border-bottom:1px #000000 solid;">
                                      <a href="<%=relpathTO%>ugeseddel_2011.asp?usemrn=<%=session("mid")%>&varTjDatoUS_man=<%=varTjDatoUS_man_tt %>"><%=ttw_txt_016 %></a>
                                    </li>

                                    <li style="border-bottom:1px #000000 solid;">
                                      <a href="<%=relpathTT%>timetag_web_kpi.asp"><%=ttw_txt_017 %></a>
                                    </li>

                         <%end select%>


                  
                        <%if lto = "hestia" OR lto = "tbg" OR lto = "mpt" OR lto = "dencker" then %>
                        <li style="border-bottom:1px #000000 solid;">
                        <a href="<%=relpathTT%>mat_web.asp">Materialer</a>
                        </li>                
                        <%end if %>

                        
                        <%if lto = "lm" then %>
                         <li style="border-bottom:1px #000000 solid;">
                            <a href="<%=relpathTT%>expence.asp"><%=expence_txt_080 %></a>
                        </li>
                        <%else %>
                        <li style="border-bottom:1px #000000 solid;">
                            <a href="<%=relpathTT%>expence.asp"><%=medarb_txt_136 %></a>
                        </li>
                        <%end if %>

                        <%if lto = "mpt" then %>
                        <li style="border-bottom:1px #000000 solid;">
                            <a href="<%=relpathTO%>job_list.asp">Projekter</a>
                        </li>

                        <li style="border-bottom:1px #000000 solid;">
                            <a href="<%=relpathTO%>materialer.asp">Lager</a>
                        </li>
                        <%end if %>

                <%end if %>
                
                       
                       
                       
                  <li><a href="https://play.google.com/store/apps/details?id=com.teamviewer.quicksupport.market"><%=tsa_txt_433 %> Android</a></li>
                <li><a href="https://itunes.apple.com/us/app/teamviewer-quicksupport/id661649585"><%=tsa_txt_433 %> IOS</a></li>
                  <li style="border-bottom:1px #000000 solid; background-color:red;">
                <a href="../sesaba.asp"><%=ttw_txt_027 %></a>
                </li>               
              
            </ul>
            <%else %>

            <ul class="nav navbar-nav navbar-right mainnav-menu">
                <li style="border-bottom:1px #000000 solid;">
                    <a href="<%=relpathTO%>monitor.asp?func=startside">Stempling</a>
                </li>

                <li style="border-bottom:1px #000000 solid;">
                    <a href="<%=relpathTT%>timetag_web_kpi.asp"><%=ttw_txt_017 %></a>
                 </li>
            </ul>

            <%end if %>

        </nav>

    </div> <!-- /.container -->

  </header>
</div>


<%end function %>