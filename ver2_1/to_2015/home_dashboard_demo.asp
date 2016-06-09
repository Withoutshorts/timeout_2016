
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->



<style>
    .container{
    width:80%;
}



</style>


<%

thisfile = "home_dashboard_demo.asp"    
    
call menu_2014() %> 
<div class="wrapper">
    <div class="content">

        <div class="container">
            <div class="portlet"> 
                <h2 class="portlet-title" style="text-align:center"><u>Velkommen til TimeOut Demo</u></h2>
                <div class="portlet-body">

                    <div class="row">
                        <div class="col-lg-2">&nbsp</div>
                        <div class="col-lg-8">
                            <h3 class="lead" align="center">
                                I vores demo har du mulighed for at få et godt indblik i TimeOut. <br /> 
                                For den bedste oplevelse anbefaler vi, at du går gennem nedenstående <b>trin-model.</b><br /><br />

                                Herefter kan du gå på opdagelse i de øvrige funktioner under hvert menu punkt.

                              

                            </h3>
                        </div>
                           <div class="col-lg-3-pad-t20">
                            <div class="panel-group accordion-panel" id="accordion-paneled">

                              <div class="panel">

                                <div class="panel-heading">
                                  <h4 class="panel-title">
                                    <a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion-paneled" href="#collapseOne">
                                        Kontakt
                                    </a>
                                  </h4>
                                </div> <!-- /.panel-heading -->

                                <div id="collapseOne" class="panel-collapse collapse">
                                  <div class="panel-body">
                                    <a href="mailto:jda@outzource.dk?Subject=Henvendelse%20til%20Outzource" target="_top">Send en Mail</a>
                                      <br /><br />
                                    <p>Telefon: 25 36 55 00</p>
                                  </div> <!-- /.panel-body -->
                                </div> <!-- /.panel-collapse -->

                              </div> <!-- /.panel --></div>
                        </div>
                    </div>
                   
                   

                   
               

                  

                    <div class="row">
                        <div class="col-lg-1">&nbsp</div>
                        <div class="col-lg-1">&nbsp</div>
                        <div class="col-lg-2">
                            <table class="table">
                                <tr>
                                   <th style="border:hidden;"><img width="147" height="64" src="img/home_db_demo/curvedarrow.jpg" /></th>                        
                                </tr>
                            </table>
                        </div>
                        <div class="col-lg-2">
                            <table class="table">
                                <tr>
                                   <th style="border:hidden;"><img width="147" height="64" src="img/home_db_demo/curvedarrow.jpg" /></th>                        
                                </tr>
                            </table>
                        </div>
                        <div class="col-lg-2">
                            <table class="table">
                                <tr>
                                   <th style="border:hidden;"><img width="147" height="64"  src="img/home_db_demo/curvedarrow.jpg" /></th>                        
                                </tr>
                            </table>
                        </div>
                        <div class="col-lg-2">
                            <table class="table">
                                <tr>
                                   <th style="border:hidden;"><img width="147" height="64"  src="img/home_db_demo/curvedarrow.jpg" /></th>                        
                                </tr>
                            </table>
                        </div>
                    </div>


                   <div class="row">
                       <div class="col-lg-1">&nbsp</div>
                      
                       <div class="col-lg-2">
                           <table class="table" style="width:100%;">
                               <tr>
                                   <th><h5>Trin 1</h5><span style="font-size:10.2em; color:#CCCCCC;" class="glyph icon-user"></span><!--<img src="img/home_db_demo/medarb2.png" width="180" height="180">-->
                                  </th>
                               </tr>
                               <tr>
                                  <td colspan="1"><a href="medarb.asp?menu=medarb&func=opret"><h5 class="lead" style="font-size:1.2em;">Opret medarbejder</h5></a>
                                      <div id="accordion-help" class="panel-group accordion-small">
                     
                                     
                                             <a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion-help" href="#medinfo" style="color:black;"><p>Læs mere....</p></a>
                                          <div id="medinfo" class="alledage panel-collapse collapse"> 
                         
                                          <p>
                                             Angiv stamdata, brugernavn og password. E-mail benyttes til notifikation samt godkendelser.<br /><br />
                                              Medarbejdertype er lig normtid pr. uge. Projektgruppen er den afdeling medarbejderen tilhører.

                                    </p></div></div>
                                  </td>
                               </tr>
                           </table>
                           
                       </div> 

                       
                       <div class="col-lg-2">
                           <table class="table" style="width:100%;"">
                               <tr>
                                   <th><h5>Trin 2</h5><span style="font-size:10.2em; color:yellowgreen;" class="glyph icon-building"></span> <!--<img src="img/home_db_demo/oprkund2.PNG" width="180" height="180">--></th>
                               </tr>
                               <tr>
                                   <td colspan="1"><a href="kunder.asp?menu=&func=opret&ketype=k&medarb=0"><h5 class="lead" style="font-size:1.2em;">Opret kunde</h5></a>
                                       <div id="accordion-help" class="panel-group accordion-small">
                     
                                       
                                             <a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion-help" href="#kundinfo" style="color:black;"><p>Læs mere....</p></a>
                                          <div id="kundinfo" class="alledage panel-collapse collapse"> 
                         
                                          <p>
                                            Udover stamdata er CVR og evt. EAN
                                            vigtig ifm. fakturering.<br /><br />
                                            Kontaktpersoner kan også oprettes men er 
                                            ikke en nødvendighed for at bruge TimeOut	


                                    </p></div></div>
                                   
                               </tr>
                           </table>
                           
                       </div> 

                       
                       <div class="col-lg-2">
                           <table class="table" style="width:100%;"">
                               <tr>
                                   <th><h5>Trin 3</h5><span style="font-size:10.2em;" class="glyph icon-file-cabinet"></span><!--<img src="img/home_db_demo/projekt3.png" width="180" height="180">--></th>
                               </tr>
                               <tr>
                                   <td colspan="1"><a href="../timereg/jobs.asp?func=opret&id=0"><h5 class="lead" style="font-size:1.2em;">Opret projekt</h5></a>
                                       <div id="accordion-help" class="panel-group accordion-small">
                     
                                      
                                             <a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion-help" href="#proinfo" style="color:black;"><p>Læs mere....</p></a>
                                          <div id="proinfo" class="alledage panel-collapse collapse"> 
                         
                                          <p>
                                              Angiv stamdata, brugernavn og password. E-mail benyttes til notifikation samt godkendelser.<br /><br />
                                              Medarbejdertype er lig normtid pr. uge. Projektgruppen er den afdeling medarbejderen tilhører.

                                    </p></div></div>
                                   </td>
                               </tr>
                           </table>
                           
                       </div> 

                       
                       <div class="col-lg-2">
                           <table class="table" style="width:100%;"">
                               <tr>
                                   <th><h5>Trin 4</h5><span style="font-size:10.2em; color:#5582d2;" class="glyph icon-stopwatch"></span><!--<img src="img/home_db_demo/timereg.png" width="180" height="180">--></th>
                               </tr>
                               <tr>
                                   <td colspan="1"><a href="../timereg/timereg_akt_2006.asp"><h5 class="lead" style="font-size:1.2em;">Timeregistrering</h5></a>
                                       <div id="accordion-help" class="panel-group accordion-small">
                     
                                        
                                             <a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion-help" href="#reginfo" style="color:#000000;"><p>Læs mere....</p></a>
                                          <div id="reginfo" class="alledage panel-collapse collapse"> 
                         
                                          <p>
                                              
                                            På Timeregistrerings-siden vælges projekt<br /><br />
                                            Timer angives i 100-tal. <br />
                                            Fx skrives 2½ time således: 2,5<br /><br />
                                            Kommentar og udlæg angives ved at klikke på "+" tegnet.

                                                  

                                    </p></div></div>
                                   </td>
                               </tr>
                           </table>
                           
                       </div> 

                       
                       <div class="col-lg-2">
                           <table class="table" style="width:100%;"">
                               <tr>
                                   <th><h5>Trin 5</h5><span style="font-size:10.2em; color:darkred;" class="glyph icon-chart-graph2"></span> <!--<img src="img/home_db_demo/rapstaimg.png" width="180" height="180">--></th>
                               </tr>
                               <tr>
                                   <td colspan="1"><a href="../timereg/joblog_timetotaler.asp"><h5 class="lead" style="font-size:1.2em;">Rapporter</h5></a>
                                       <div id="accordion-help" class="panel-group accordion-small">
                     
                                       
                                             <a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion-help" href="#rapinfo" style="color:black;"><p>Læs mere....</p></a>
                                          <div id="rapinfo" class="alledage panel-collapse collapse"> 
                         
                                          <p>
                                             Timeout indeholder flere forskellige typer rapporter og dashboards. <br /><br />
                                             Der kan slices og opdeles efter kunder, projekter, medarbejdere, afdelinger forretningsområder m.m.
                                              </p>

                                    </div></div>
                                   </td>
                               </tr>
                           </table>
                           
                       </div> 
                </div>

              

                <div class="row">
                    <div class="col-lg-2"><h6 class="lead">Vi kan også:</h6></div>
                </div>
              
                <div class="row">
                    <div class="col-lg-2"><a href="img/home_db_demo/TimeTag_TimeOut Mobile.pdf" target="_blank">Timeout Mobile og timetag</a></div>
                </div>
                <div class="row">
                    <div class="col-lg-2"><a href="img/home_db_demo/RessourceStyring.pdf" target="_blank">Ressource Styring</a></div>
                </div>
                <div class="row">
                    <div class="col-lg-2"><a href="img/home_db_demo/Rapporter_Statistikker.pdf" target="_blank">Statistikker & Rapporter</a></div>

                      <!--<div class="col-lg-4">
                                Kontakt os gerne for spørgsmål eller specifikke krav som du måtte have.<br /><br />

                                God fornøjelse.<br />
                                Hilsen Outzource
                          </div>-->

                </div>


            </div>
            </div>
        </div>
        
    </div>
</div>

<!--#include file="../inc/regular/footer_inc.asp"-->