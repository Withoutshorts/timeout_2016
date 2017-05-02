

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<style>

    .btn 
    {
        font-size: 150%;
        width:100px;
    }

</style>

<%
    if len(session("user")) = 0 then
	%>
	
	<%
	errortype = 5
	call showError(errortype)
       response.End
	end if 

    medid = session("user")
    func = request("func") 
%>


<div class="wrapper">
    <div class="content">

        <div class="container">
            <div class="portlet">
                <div class="portlet-body">

                    <div class="row">
                        <div class="col-lg-12" style="text-align:center;>
                            <a href="#"><img src="img/Dencker_logo.png" width="100" /></a>
                        </div>
                    </div>
                    
                    <%                        
                        select case func
                        case "startside" 
                    %>

                    <div class="row" style="margin-top:10%">
                        <div class="col-lg-12" style="text-align:center"><a href="komme_gaa_chip.asp?func=kom" class="btn btn-default" style="width:200px;"><b>Scan kort</b></a></div>
                    </div>

                    <%
                        case "kom"
                        response.Write medid 
                    %>


                    <div class="row" style="margin-top:10%">
                        <div class="col-lg-1"></div>
                        <div class="col-lg-2"><a href="komme_gaa_chip.asp?func=checkind" class="btn btn-success"><b>Normal</b></a></div>
                        <div class="col-lg-1"></div>            
                        <div class="col-lg-2"><a type="submit" class="btn btn-success"><b>Extern</b></a></div>
                        <div class="col-lg-1"></div>            
                        <div class="col-lg-2"><a type="submit" class="btn btn-success"><b>U. løn</b></a></div>
                        <div class="col-lg-1"></div>            
                        <div class="col-lg-2"><a type="submit" class="btn btn-success"><b>Tilkald</b></a></div>
                    </div>

                    
                    <%case "gaa" %>

                    <div class="row" style="margin-top:10%">
                        <div class="col-lg-12" style="text-align:center"><a type="submit" class="btn btn-danger"><b>Gå</b></a></div>
                    </div>


                    <%case "checkind" 
                        
                       
                      
                        
                    %>
                    <div class="row" style="margin-top:10%">
                        <h4 class="col-lg-12" style="text-align:center"><%response.Write medid & "<br>" & Date %></h4>
                    </div>

                    <%
                        end select
                    %>

                </div>
            </div>
        </div>

    </div>
</div>
    

<!--#include file="../inc/regular/footer_inc.asp"-->