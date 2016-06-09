
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


<%call menu_2014() %> 
<div class="wrapper">
    <div class="content">

        <div class="container">
            <div class="portlet"> 
                <h2 class="portlet-title" style="text-align:center"><u><br />Velkommen til TimeOut Demo</u></h2>
                <div class="portlet-body">

                    <div class="row">
                        <div class="col-lg-2">&nbsp</div>
                        <div class="col-lg-9">
                            <h3 class="lead">
                                I vores demo har du mulighed for at få et godt indblik i TimeOut.  
                                For den bedste oplevelse anbefaler vi at gå du gennem nedenstående trin-model.
                            </h3>
                        </div>
                    </div>
                    <div class="row">
                        <div class="col-lg-4">&nbsp</div>
                        <div class="col-lg-8">
                            <h3 class="lead">
                                Herefter kan du gå på opdagelse i de øvrige funktioner under hvert menu punkt.
                            </h3>
                        </div>
                    </div>
                    <div class="row">
                        <div class="col-lg-3">&nbsp</div>
                        <div class="col-lg-8">
                            <h3 class="lead">
                                &nbsp &nbsp &nbsp Kontakt os gerne for spørgsmål eller specifikke krav som du måtte have.

                                God fornøjelse.
                                Hilsen Outzource
                            </h3>
                        </div>
                    </div>

                    <div class="row">
                        <div class="col-lg-10">&nbsp</div>
                        <div class="col-lg-2-pad-t20">
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
                                    <a href="mailto:someone@example.com?Subject=Hello%20again" target="_top">Send en Mail</a>
                                      <br /><br />
                                    <p>Telefon: 25 36 55 00</p>
                                  </div> <!-- /.panel-body -->
                                </div> <!-- /.panel-collapse -->

                              </div> <!-- /.panel --></div>
                        </div>
                    </div>
                   
               
                   <!--img style="position:absolute; top:320px; left:450px;" width="10%" src="img/curvedarrow.jpg" />
                   <img style="position:absolute; top:320px; left:700px;" width="10%" src="img/curvedarrow.jpg" />
                   <img style="position:absolute; top:320px; left:950px;" width="10%" src="img/curvedarrow.jpg" />
                   <img style="position:absolute; top:320px; left:1200px;" width="10%" src="img/curvedarrow.jpg" /-->

                   <br /><br /> <br /> 


                    <div class="row">
                        <div class="col-lg-1">&nbsp</div>
                        <div class="col-lg-1">&nbsp</div>
                        <div class="col-lg-2">
                            <table class="table">
                                <tr>
                                   <th style="border:hidden;"><img width="90%" src="img/curvedarrow.jpg" /></th>                        
                                </tr>
                            </table>
                        </div>
                        <div class="col-lg-2">
                            <table class="table">
                                <tr>
                                   <th style="border:hidden;"><img width="90%" src="img/curvedarrow.jpg" /></th>                        
                                </tr>
                            </table>
                        </div>
                        <div class="col-lg-2">
                            <table class="table">
                                <tr>
                                   <th style="border:hidden;"><img width="90%" src="img/curvedarrow.jpg" /></th>                        
                                </tr>
                            </table>
                        </div>
                        <div class="col-lg-2">
                            <table class="table">
                                <tr>
                                   <th style="border:hidden;"><img width="90%" src="img/curvedarrow.jpg" /></th>                        
                                </tr>
                            </table>
                        </div>
                    </div>


                   <div class="row">
                       <div class="col-lg-1">&nbsp</div>
                      
                       <div class="col-lg-2">
                           <table class="table table-bordered" style="width:100%;">
                               <tr>
                                   <th><img src="img/oprmbr.PNG" width="85%" height="175"></th>
                               </tr>
                               <tr>
                                  <th colspan="1"><a href="#"><h5>Trin 1</h5><h5 class="lead">Opret medarbejder</h5></a>
                                      <div id="accordion-help" class="panel-group accordion-small">
                     
                                        <div class="panel open">
                                          <div class="panel-heading">
                                             <a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion-help" href="#medinfo" style="color:black;"><p>Læs mere....</p></a>
                                          <div id="medinfo" class="alledage panel-collapse collapse"> 
                         
                                          <div class="panel-body">
                                             <p> Angiv stamdata, brugernavn og password.
                                              E-mail benyttes til notifikation samt godkendelser.
                                              Medarbejdertype er lig normtid pr. uge.
                                              Projektgruppen er den afdeling medarbejderen tilhører.</p>

                                    </div></div></div></div></div>
                                  </th>
                               </tr>
                           </table>
                           
                       </div> 

                       
                       <div class="col-lg-2">
                           <table class="table table-bordered" style="width:100%;"">
                               <tr>
                                   <th><img src="img/oprkund.PNG" width="100%" height="175"></th>
                               </tr>
                               <tr>
                                   <th colspan="1"><a href="#"><h5>Trin 2</h5><h5 class="lead">Opret kunde</h5></a>
                                       <div id="accordion-help" class="panel-group accordion-small">
                     
                                        <div class="panel open">
                                          <div class="panel-heading">
                                             <a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion-help" href="#kundinfo" style="color:black;"><p>Læs mere....</p></a>
                                          <div id="kundinfo" class="alledage panel-collapse collapse"> 
                         
                                          <div class="panel-body">
                                              Angiv stamdata, brugernavn og password.
                                              E-mail benyttes til notifikation samt godkendelser.
                                              Medarbejdertype er lig normtid pr. uge.
                                              Projektgruppen er den afdeling medarbejderen tilhører.

                                    </div></div></div></div></div>
                                   
                               </tr>
                           </table>
                           
                       </div> 

                       
                       <div class="col-lg-2">
                           <table class="table table-bordered" style="width:100%;"">
                               <tr>
                                   <th><img src="img/projektimg.jpg" width="100%" height="175"></th>
                               </tr>
                               <tr>
                                   <th colspan="1"><a href="#"><h5>Trin 3</h5><h5 class="lead">&nbsp Opret projekt</h5></a>
                                       <div id="accordion-help" class="panel-group accordion-small">
                     
                                        <div class="panel open">
                                          <div class="panel-heading">
                                             <a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion-help" href="#proinfo" style="color:black;"><p>Læs mere....</p></a>
                                          <div id="proinfo" class="alledage panel-collapse collapse"> 
                         
                                          <div class="panel-body">
                                              Angiv stamdata, brugernavn og password.
                                              E-mail benyttes til notifikation samt godkendelser.
                                              Medarbejdertype er lig normtid pr. uge.
                                              Projektgruppen er den afdeling medarbejderen tilhører.

                                    </div></div></div></div></div>
                                   </th>
                               </tr>
                           </table>
                           
                       </div> 

                       
                       <div class="col-lg-2">
                           <table class="table table-bordered" style="width:100%;"">
                               <tr>
                                   <th><img src="img/tidsreg.png" width="100%" height="175"></th>
                               </tr>
                               <tr>
                                   <th colspan="1"><a href="#"><h5>Trin 4</h5><h5 class="lead">Registrering af tid m.m.</h5></a>
                                       <div id="accordion-help" class="panel-group accordion-small">
                     
                                        <div class="panel open">
                                          <div class="panel-heading">
                                             <a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion-help" href="#reginfo" style="color:black;"><p>Læs mere....</p></a>
                                          <div id="reginfo" class="alledage panel-collapse collapse"> 
                         
                                          <div class="panel-body">
                                              Angiv stamdata, brugernavn og password.
                                              E-mail benyttes til notifikation samt godkendelser.
                                              Medarbejdertype er lig normtid pr. uge.
                                              Projektgruppen er den afdeling medarbejderen tilhører.

                                    </div></div></div></div></div>
                                   </th>
                               </tr>
                           </table>
                           
                       </div> 

                       
                       <div class="col-lg-2">
                           <table class="table table-bordered" style="width:100%;"">
                               <tr>
                                   <th><img src="img/rapstaimg.png" width="100%" height="175"></th>
                               </tr>
                               <tr>
                                   <th colspan="1"><a href="#"><h5>Trin 5</h5><h5 class="lead">Rapporter og statistikker</h5></a>
                                       <div id="accordion-help" class="panel-group accordion-small">
                     
                                        <div class="panel open">
                                          <div class="panel-heading">
                                             <a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion-help" href="#rapinfo" style="color:black;"><p>Læs mere....</p></a>
                                          <div id="rapinfo" class="alledage panel-collapse collapse"> 
                         
                                          <div class="panel-body">
                                              Angiv stamdata, brugernavn og password.
                                              E-mail benyttes til notifikation samt godkendelser.
                                              Medarbejdertype er lig normtid pr. uge.
                                              Projektgruppen er den afdeling medarbejderen tilhører.

                                    </div></div></div></div></div>
                                   </th>
                               </tr>
                           </table>
                           
                       </div> 
                </div>

              

                <div class="row">
                    <div class="col-lg-2"><h6 class="lead">Læs mere om:</h6></div>
                </div>
                <div class="row">
                    <div class="col-lg-5"><a href="#">En medarbejders version af timeout</a></div>
                </div>
                <div class="row">
                    <div class="col-lg-5"><a href="#">Timeout Mobile og timetag</a></div>
                </div>
                <div class="row">
                    <div class="col-lg-5"><a href="#">Ressource Styring</a></div>
                </div>
                <div class="row">
                    <div class="col-lg-5"><a href="#">Statistikker & Rapporter</a></div>
                </div>


            </div>
            </div>
        </div>
        
    </div>
</div>

<!--#include file="../inc/regular/footer_inc.asp"-->