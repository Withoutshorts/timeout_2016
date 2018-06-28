

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->


<%
    visning = request("vising")
    'visning 1 = hold pause
    'visning 2 = afslut pause
    ' visning eller = scane kort

    select case func
    case "login"



    end select

%>


<script src="js/frokost_jav.js" type="text/javascript"></script>

<div class="wrapper">
    <div class="content">


        <div class="container">
            <div class="portlet">
                <div class="portlet">
                    <div class="portlet-body">

                        <div class="row">
                            <div class="col-lg-12" style="text-align:center;">
                                <a href="#"><img src="img/Dencker_logo.png" width="100" /></a>
                            </div>
                        </div>

                        <br /><br />

                        <div class="row">
                            <div class="col-lg-5"></div>
                            <div class="col-lg-2"><input type="text" id="medarbId"  class="form-control input-small" /></div>

                            <div class="col-lg-2">
                                <span style="font-size:125%; visibility:hidden;" id="pauseCheck" class="fa fa-check"></span>
                            </div>
                            
                        </div> 

                        <br />

                        <div class="row">
                            <h4 class="col-lg-12" style="text-align:center">Scan dit kort</h4>
                        </div>

                    </div>
                </div>
            </div>
        </div>

    </div>
</div>

    

<!--#include file="../inc/regular/footer_inc.asp"-->