<!DOCTYPE html>

<% function mobile_header %>


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

                <li>
                <a href="timetag_web.asp">Tidsregistrering</a>
                </li>

                <li>
                <a href="../to_2015/ugeseddel_2011.asp?usemrn=<%=session("mid")%>&varTjDatoUS_man=<%=varTjDatoUS_man_tt %>">Ugeseddel</a>
                </li>    

                <li>
                <a href="timetag_web_kpi.asp">Nøgetal</a>
                </li>           
              
            </ul>

        </nav>

    </div> <!-- /.container -->

  </header>
</div>


<%end function %>