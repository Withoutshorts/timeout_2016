

<% function mobile_header %>


        <%select case thisfile
        case "ugeseddel_2011.asp"
        relpathTT = "../timetag_web/"
        relpathTO = ""
        case else 
        relpathTT = ""
        relpathTO = "../to_2015/"
        end select %>


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

            <ul class="nav navbar-nav navbar-right mainnav-menu">

                <li style="border-bottom:1px #000000 solid;">
                <a href="<%=relpathTT%>timetag_web.asp"><%=ttw_txt_015 %></a>
                </li>

                <li style="border-bottom:1px #000000 solid;">
                <a href="<%=relpathTO%>ugeseddel_2011.asp?usemrn=<%=session("mid")%>&varTjDatoUS_man=<%=varTjDatoUS_man_tt %>"><%=ttw_txt_016 %></a>
                </li>    

                <li style="border-bottom:1px #000000 solid;">
                <a href="<%=relpathTT%>timetag_web_kpi.asp"><%=ttw_txt_017 %></a>
                </li>         
              
            </ul>

        </nav>

    </div> <!-- /.container -->

  </header>
</div>


<%end function %>