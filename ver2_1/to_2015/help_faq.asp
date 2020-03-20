

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<%
    
if len(session("user")) = 0 then
	
errortype = 5
call showError(errortype)
response.end
end if
     
%>

<%call menu_2014 %>

<div class="wrapper">
    <div class="content">

        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u>Hjælp / Help / FAQ</u></h3>
                <div class="portlet-body">
                    <div class="row">
                        <h4 class="col-lg-3"><%if lto = "sduuas" OR lto = "sducei" then %>Timeout<%end if %> Kontakt information</h4>
                        <div class="col-lg-1"></div>
                        <h4 class="col-lg-3"><%if lto = "sduuas" OR lto = "sducei" then %>Timeout<%end if %> Contact information</h4>
                        <%if lto = "sduuas" OR lto = "sducei" then %>
                        <div class="col-lg-1"></div>
                        <h4 class="col-lg-2">UAS/CEI Contact information</h4>
                        <%end if %>
                    </div>
                    <div class="row">
                        <h5 class="col-lg-4">Mandag - Fredag: 09:00 - 16:00</h5>
                        <h5 class="col-lg-4">Monday - Friday: 09:00 - 16:00</h5>
                        <%if lto = "sduuas" OR lto = "sducei" then %>
                        <h5 class="col-lg-4">Email: TimeOut@mmmi.sdu.dk</h5>
                        <%end if %>
                    </div>
                    <div class="row">
                        <h5 class="col-lg-4">Telefon: 25 36 55 00</h5>
                        <h5 class="col-lg-4">Phone: +45 25 36 55 00</h5>
                    </div>
                    <div class="row">
                        <h5 class="col-lg-4">Email: Support@outzource.dk</h5>
                        <h5 class="col-lg-4">Email: Support@outzource.dk</h5>
                    </div>

                    <br /><br /><br />
                    <div class="panel-group accordion-panel" id="accordion-paneled">
                        <div class="panel panel-default">
                            <div class="panel-heading"><h4 class="panel-title"><a class="accordion-toggle" data-toggle="collapse" data-target="#collapseOne">Danske vejledninger</a></h4></div> 
                            <div id="collapseOne" class="panel-collapse collapse">
                                <div class="panel-body">

                                    <table class="table">
                                        <tr style="border-top:hidden">
                                            <td><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/pdf/help/TimeOut - FAQ.pdf" target="_blank">TimeOut - FAQ.pdf</a></td>
                                        </tr>
                                        <tr style="border-top:hidden">
                                            <td><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/pdf/help/TimeOut Guide - Afstemning.pdf" target="_blank">TimeOut Guide - Afstemning.pdf</a></td>
                                            <td><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/pdf/help/TimeOut Guide - Dashboard.pdf" target="_blank">TimeOut Guide - Dashboard.pdf</a></td>
                                            <td><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/pdf/help/TimeOut Guide - Diæter.pdf" target="_blank">TimeOut Guide - Diæter.pdf</a></td>
                                            <td><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/pdf/help/TimeOut Guide - Favoritliste.pdf" target="_blank">TimeOut Guide - Favoritliste.pdf</a></td>                                           
                                        </tr>
                                        <tr style="border-top:hidden">
                                            <td><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/pdf/help/TimeOut Guide - Grandtotal.pdf" target="_blank">TimeOut Guide - Grandtotal.pdf</a></td>
                                            <td><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/pdf/help/TimeOut Guide - HR Liste og Matr oversigt.pdf" target="_blank">TimeOut Guide - HR LIste og Matr oversigt.pdf</a></td>
                                            <td><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/pdf/help/TimeOut Guide - Joblog.pdf" target="_blank">TimeOut Guide - Joblog.pdf</a></td>
                                            <td><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/pdf/help/TimeOut Guide - Joboverblik.pdf" target="_blank">TimeOut Guide - Joboverblik.pdf</a></td>
                                        </tr>
                                        <tr style="border-top:hidden">
                                            <td><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/pdf/help/TimeOut Guide - Kunde.pdf" target="_blank">TimeOut Guide - Kunde.pdf</a></td>
                                            <td><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/pdf/help/TimeOut Guide - Login.pdf" target="_blank">TimeOut Guide - Login.pdf</a></td>
                                            <td><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/pdf/help/TimeOut Guide - Matarialer.pdf" target="_blank">TimeOut Guide - Materialer.pdf</a></td>
                                            <td><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/pdf/help/TimeOut Guide - Medarbejder.pdf" target="_blank">TimeOut Guide - Medarbejder.pdf</a></td>
                                        </tr>
                                        <tr style="border-top:hidden">
                                            <td><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/pdf/help/TimeOut Guide - Medarbejdertype_og-opdatering.pdf" target="_blank">TimeOut Guide - Medarbejdertype.pdf</a></td>
                                            <td><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/pdf/help/TimeOut Guide - Mobil.pdf" target="_blank">TimeOut Guide - Mobil.pdf</a></td>
                                            <td><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/pdf/help/TimeOut Guide - Opret projekt inkl forecast.pdf" target="_blank">TimeOut Guide - Projekt inkl forecast.pdf</a></td>
                                            <td><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/pdf/help/TimeOut Guide - Tidsregistrering.pdf" target="_blank">TimeOut Guide - Tidsregistreing.pdf</a></td>
                                        </tr>
                                        <tr style="border-top:hidden">
                                            <td><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/pdf/help/TimeOut Guide - Ugeseddel.pdf" target="_blank">TimeOut Guide - Ugeseddel.pdf</a></td>
                                            <td><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/pdf/help/Timeout _flyt_aktivitet_timer_funktion.pdf" target="_blank">TimeOut Guide - Flyt aktivitet timer</a></td>
                                        </tr>
                                    </table>

                                </div>
                            </div>
                        </div>
                        
                        <div class="panel panel-default">
                            <div class="panel-heading"><h4 class="panel-title"><a class="accordion-toggle" data-toggle="collapse" data-target="#collapse2">English guides</a></h4></div> 
                            <div id="collapse2" class="panel-collapse collapse">
                                <div class="panel-body">
                                    <table class="table">
                                        <tr style="border-top:hidden">
                                            <td><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/pdf/help/TimeOut_FAQ.pdf" target="_blank">TimeOut - FAQ.pdf</a></td>
                                        </tr>
                                        <tr style="border-top:hidden">
                                            <td><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/pdf/help/TimeOut Guide_CreateProject.pdf" target="_blank">TimeOut Guide - Create Project.pdf</a></td>
                                            <td><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/pdf/help/TimeOut Guide_Grandtotal.pdf" target="_blank">TimeOut Guide - Grandtotal.pdf</a></td>
                                            <td><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/pdf/help/TimeOut Guide_Joblog.pdf" target="_blank">TimeOut Guide - Joblog.pdf</a></td>
                                            <td><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/pdf/help/TimeOut Guide_Mobile.pdf" target="_blank">TimeOut Guide - Mobile.pdf</a></td>                                          
                                        </tr>
                                        <tr style="border-top:hidden">
                                            <td><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/pdf/help/TimeOut Guide_Weeklynote.pdf" target="_blank">TimeOut Guide - Weekly note.pdf</a></td>     
                                        </tr>
                                    </table>                                   
                                </div>
                            </div>
                        </div> 

                        <div class="panel panel-default">
                            <div class="panel-heading"><h4 class="panel-title"><a class="accordion-toggle" data-toggle="collapse" data-target="#collapse3"><%=lto %> guides</a></h4></div> 
                            <div id="collapse3" class="panel-collapse collapse">
                                <div class="panel-body">

                                    <div class="row">
                                        <div class="col-lg-6">
                                            <table class="table">
                                                <%
                                                select case lto 
                                                
                                                case "tia"    
                                                %>
                                                <tr style="border-top:hidden">
                                                    <td><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/pdf/help/tia/TimeOutFAQ.PPTX" target="_blank">Timeout guide.pptx</a></td>
                                                </tr>
                                                <%
                                                case "sduuas", "sducei"
                                                %>
                                                
                                                <tr style="border-top:hidden">                                                    
                                                    <td><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/pdf/help/sduuas/TimeOut-philosophy-and-guidance-ver3.pdf" target="_blank">Timeout guide.pdf</a></td>
                                                </tr>

                                                 <tr style="border-top:hidden">                                                    
                                                    <td><a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/interreg/Excel_add_in/16-07-11-Timeseddel-Interreg-5A - English.xlsm" target="_blank">Interreg Excel</a></td>
                                                </tr>

                                                <%case else %>
                                                <tr style="border-top:hidden">
                                                    <td>Ingen filer fundet</td>
                                                </tr>
                                                <%end select %>
                                            </table>
                                        </div>

                                        <!-- Administrator kommentar -->
                                       <!-- <div class="col-lg-6">
                                            <table class="table">
                                                <tr style="border-top:hidden">
                                                    <td><textarea rows="4" class="form-control input-small" placeholder="Administrator kommentar" readonly></textarea></td>
                                                </tr>
                                            </table>
                                        </div> -->
                                    </div>
                                                          
                                </div>
                            </div>
                        </div>

                    </div>

                </div>
            </div>
        </div>

    </div>
</div>   

<!--#include file="../inc/regular/footer_inc.asp"-->