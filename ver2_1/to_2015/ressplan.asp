


<%response.buffer = true 
Session.LCID = 1030
%>
			        

<!--#include file="../inc/connection/conn_db_inc.asp"-->


<%'** JQUERY START ************************* %>

<%'** JQUERY END ************************* %>
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->

<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<%
    if len(session("user")) = 0 then
	%>
	
	<%
	errortype = 5
	call showError(errortype)
       response.End
	end if


    func = request("func")                    
%>

<style>
    .hover_img a { position:relative; color:black; }
    .hover_img a span { position:absolute; display:none; z-index:90; border: 2px solid green; }
    .hover_img a:hover span { display:block; width:200px; height:200px;}

    table th 
    {
        font-size:110%
    }

    #block_container
{
    text-align:left;
}
#bloc1, #bloc2
{
    display:inline;
}



</style>



<div class="wrapper">
    <div class="content">   
        
        <div class="container" style="width:1600px;">
            <div class="portlet">
                <h3 class="portlet-title"><u>Ressource Planlægning</u></h3>
                <div class="portlet-body">


                    <div id="basicModal" class="modal fade" style="margin-top:100px">
                        
                    <div class="modal-dialog">
                        

                      <div class="modal-content">                                                          
                        <div class="modal-body">
                            <br />

                            <div class="row">
                                <div class="col-lg-2">Start:</div>
                                <div class="col-lg-4">
                                    <div class='input-group date' id='datepicker_stdato'>
                                        <input type="text" class="form-control input-small" name="" value="" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon input-small">
                                            <span class="fa fa-calendar"></span>
                                        </span>
                                    </div>
                                </div>
                                <div class="col-lg-2">Fra:</div>
                                <div class="col-lg-4"><input type="text" class="form-control input-small" placeholder="00:00" /></div>
                            </div>
                            <div class="row">
                                <div class="col-lg-2">Slut:</div>
                                <div class="col-lg-4">
                                    <div class='input-group date' id='datepicker_stdato'>
                                        <input type="text" class="form-control input-small" name="" value="" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon input-small">
                                            <span class="fa fa-calendar"></span>
                                        </span>
                                    </div>
                                </div>
                                <div class="col-lg-2">Til:</div>
                                <div class="col-lg-4"><input type="text" class="form-control input-small" placeholder="00:00" /></div>
                            </div>

                            <div class="row">
                                <div class="col-lg-2">Gentagelse:</div>
                                <div class="col-lg-4">
                                    <select class="form-control input-small" id="test" onchange="func2()">
                                        <option value="dag">Daglig</option>
                                        <option value="uge">Ugentlig</option>
                                        <option value="maaned">Hver Måned</option>
                                        <option value="maaned">Hver 2. Måned</option>
                                        <option value="maaned">Hver 3. Måned</option>
                                        <option value="maaned">Hver 6. Måned</option>
                                        <option value="maaned" onclick="selmaaned.call">Hver 12. Månded</option>
                                    </select>
                                </div>
                                <div class="col-lg-2">Vigtig:</div>
                                <div class="col-lg-4"><input type="checkbox" value="0" /></div>
                            </div>
                            <div id="showmonth" style="display:none">
                                <div class="row">
                                    <div class="col-lg-2">Gentag på:</div>
                                    <div class="col-lg-2"><input type="radio" name="monthdet" value="1" /> Ugedag</div>
                                    <div class="col-lg-2"><input type="radio" name="monthdet" value="2" /> Dato</div>

                                </div>
                               
                            </div>
                            <br />
                            <div class="row">
                                <div class="col-lg-2">Medarbejder:</div>
                                <div class="col-lg-4">
                                    <select class="form-control input-small">
                                        <option value="1">Hans</option>
                                        <option value="2">Per</option>
                                        <option value="3">Søren Karlsen</option>
                                    </select>
                                </div>
                                <div class="col-lg-2">Kunde:</div>
                                <div class="col-lg-4"><input type="text" value="Outzource Aps" class="form-control input-small" readonly/></div>
                            </div>
                            <div class="row">
                                <div class="col-lg-2">Type:</div>
                                <div class="col-lg-4">
                                    <select class="form-control input-small">
                                        <option valse="1">Faktuerbar</option>
                                        <option valse="2">Ikke Faktuerbar</option>
                                        <option valse="3">Ferie</option>
                                        <option valse="4">Sygdom</option>
                                    </select>
                                </div>
                                <div class="col-lg-2">Projekt:</div>
                                <div class="col-lg-4"><input type="text" value="Design" class="form-control input-small" readonly/></div>
                            </div>
                            <div class="row">
                                <div class="col-lg-2">Overskrift:</div>
                                <div class="col-lg-4">
                                    <select class="form-control input-small">
                                        <option value="">Ingen</option>
                                        <option value="">Aktivitet</option>
                                        <option value="">Projekt</option>
                                        <option value="">Kunde</option>                                     
                                    </select>
                                </div>
                                <div class="col-lg-2">Aktivitet:</div>
                                <div class="col-lg-4"><input type="text" value="Produkion" class="form-control input-small" readonly/></div>
                            </div>

                          <!--  <br />
                            <div class="row">
                                <div class="col-lg-2">Medarbejder:</div>
                                <div class="col-lg-4">
                                    <select class="form-control input-small">
                                        <option value="1">Hans</option>
                                        <option value="2">Per</option>
                                        <option value="3">Søren Karlsen</option>
                                    </select>
                                </div>
                                
                                <div class="col-lg-2">Kunde:</div>
                                <div class="col-lg-4"><input type="text" value="Outzource Aps" class="form-control input-small" readonly/></div>

                            </div>
                            
                            <div class="row">
                                <div class="col-lg-2">Dato:</div>
                                <div class="col-lg-4">
                                    <div class='input-group date' id='datepicker_stdato'>
                                    <input type="text" class="form-control input-small" name="" value="" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon input-small">
                                            <span class="fa fa-calendar"></span>
                                        </span>
                                    </div>
                                </div>
                                

                                <div class="col-lg-2">Projekt:</div>
                                <div class="col-lg-4"><input type="text" value="Design" class="form-control input-small" readonly/></div>
                            </div>

                            <div class="row">
                                <div class="col-lg-2">Tid:</div>
                                <div class="col-lg-4">
                                    <div class="bootstrap-timepicker">
                                        <input type="text" id="timepicker1" class="form-control input-small" placeholder="10:00 - 11:00">
                                    </div>
                                </div>

                                <div class="col-lg-2">Aktivitet:</div>
                                <div class="col-lg-4"><input type="text" value="Produkion" class="form-control input-small" readonly/></div>
                            </div>
                            <br />
                            <div class="row">
                                <div class="col-lg-2">Kategoriser:</div>
                                <div class="col-lg-4">
                                    <select class="form-control input-small">
                                        <option valse="1">Grå</option>
                                        <option valse="2">Blå</option>
                                        <option valse="3">Gul</option>
                                        <option valse="4">Grøn</option>
                                        <option valse="5">Rød</option>
                                    </select>
                                </div>

                                <div class="col-lg-2">Overskrift:</div>
                                <div class="col-lg-4">
                                    <select class="form-control input-small">
                                        <option value="">Ingen</option>
                                        <option value="">Aktivitet</option>
                                        <option value="">Projekt</option>
                                        <option value="">Kunde</option>                                     
                                    </select>
                                </div>
                            </div> -->
                                
                                <script type="text/javascript">
                                    function func2(){
                                        if (document.getElementById('test').value == "maaned")
                                        {
                                            document.getElementById("showmonth").style.display = "inherit";
                                        } else {
                                            document.getElementById("showmonth").style.display = "none";
                                        }
                                    }
                                </script>

                           <!-- <div class="row">
                                <div class="col-lg-2">Husk:</div>
                                <div class="col-lg-4"><input type="checkbox" value="0" /></div>
                                
                                <div class="col-lg-2">Gentagelse:</div>
                                <div class="col-lg-4">
                                    <select class="form-control input-small" id="test" onchange="func2()">
                                        <option value="dag">Daglig</option>
                                        <option value="uge">Ugentlig</option>
                                        <option value="maaned">Hver Måned</option>
                                        <option value="maaned">Hver 2. Måned</option>
                                        <option value="maaned">Hver 3. Måned</option>
                                        <option value="maaned">Hver 6. Måned</option>
                                        <option value="maaned" onclick="selmaaned.call">Hver 12. Månded</option>
                                    </select>
                                </div>
                            </div>

                            <div id="showmonth" style="display:none">
                                <div class="row">
                                    <div class="col-lg-6">&nbsp</div>
                                    <div class="col-lg-2">På ugedag:</div>
                                    <div class="col-lg-4"><input type="radio" name="monthdet" value="1" /></div>
                                </div>
                                <div class="row">
                                    <div class="col-lg-6">&nbsp</div>
                                    <div class="col-lg-2">På dato:</div>
                                    <div class="col-lg-4"><input type="radio" name="monthdet" value="2" /></div>
                                </div>
                            </div> -->
                        
            
                        </div> 

                        <div class="modal-footer" style="text-align:left">
                          <button type="button" class="btn btn-default">Opdater</button>
                        </div> 
                      </div> 
                    </div>
                  </div>


                     <!--   <div class="well">
                            <div class="row">
                                <div class="col-lg-2"><b>Visninger:</b></div>                                           
                        
                                <div class="col-lg-2">
                                    <select class="form-control input-small">
                                        <option value="0">Alle</option>
                                        <option value="1">Vis projekt tid</option>
                                        <option value="2">Vis fravær</option>
                                    </select>
                                </div>

                                <div class="col-lg-5">&nbsp</div>
                            
                                <div class="col-lg-1">
                                    <form action="ressplan.asp?func=day" method="post">
                                        <button type="submit" class="btn btn-secondary btn-sm pull-right"><b>Dag</b></button>
                                    </form>
                                </div>

                                <div class="col-lg-1">
                                    <form action="ressplan.asp?func=week" method="post">
                                        <button type="submit" class="btn btn-secondary btn-sm pull-right"><b>Uge</b></button>
                                    </form>
                                </div>

                                <div class="col-lg-1">
                                    <form action="ressplan.asp?func=maened" method="post">
                                        <button type="submit" class="btn btn-secondary btn-sm pull-right"><b>Måned</b></button>
                                    </form>
                                </div>
                                                                                                                       
                          </div>
                        </div> -->


                   <%select case func %>   
                    
                    <%
                        case "day" 
                    %>
                    <!-----------------------------------------------  Dag --------------------------------------------------------------------------------------->
                    <table class="table dataTable table-bordered ui-datatable">
                        <tr>
                            <th style="background-color:#f5f5f5; vertical-align:middle; border-top:hidden;">
                                <input type="text" class="form-control input-small" placeholder="Søg kunde/job" value="">
                            </th>
                            <th colspan="70" style="text-align:center; border-top:hidden; background-color:#f5f5f5; font-size:175%"><a href="#"><</a> &nbsp Mandag 01-01-16 &nbsp <a href="#">></a> 
                                    <div class="pull-right">
                                        <a href="ressplan.asp?func=day" class="btn btn-secondary btn-sm"><b>Dag</b></a>
                                        <a href="ressplan.asp?func=week" class="btn btn-default btn-sm"><b>Uge</b></a>
                                        <a href="ressplan.asp?" class="btn btn-default btn-sm"><b>Måned</b></a>
                                    </div>
                            </th>
                        </tr>

                        <tr>
                            <th style="background-color:#f5f5f5">
                                <select class="form-control input-small">
                                      <option value="0">Vis type</option>
                                      <option value="1">Projekt tid</option>
                                      <option value="2">Fravær</option>
                                      <option value="3">Vigtige</option>
                                </select>
                            </th>
                            <th colspan="69" style="background-color:#f5f5f5; border-top:hidden"></th>
                        </tr>

                        <tr>
                            <th style="background-color:#f5f5f5">
                                <select class="form-control input-small">
                                      <option value="0">Vis gruppe</option>
                                      <option value="1">Projekt tid</option>
                                      <option value="2">Fravær</option>
                                </select>
                            </th>
                            <th style="text-align:center; background-color:#f5f5f5; width:5%">06:00</th>                           
                            <th style="text-align:center; background-color:#f5f5f5; width:5%">07:00</th>
                            <th style="text-align:center; background-color:#f5f5f5; width:5%">08:00</th>
                            <th style="text-align:center; background-color:#f5f5f5; width:5%">09:00</th>
                            <th style="text-align:center; background-color:#f5f5f5; width:5%">10:00</th>
                            <th style="text-align:center; background-color:#f5f5f5; width:5%">11:00</th>
                            <th style="text-align:center; background-color:#f5f5f5; width:5%">12:00</th>
                            <th style="text-align:center; background-color:#f5f5f5; width:5%">13:00</th>
                            <th style="text-align:center; background-color:#f5f5f5; width:5%">14:00</th>
                            <th style="text-align:center; background-color:#f5f5f5; width:5%">15:00</th>
                            <th style="text-align:center; background-color:#f5f5f5; width:5%">16:00</th>
                            <th style="text-align:center; background-color:#f5f5f5; width:5%">17:00</th>
                            <th style="text-align:center; background-color:#f5f5f5; width:5%">18:00</th>
                            <th style="text-align:center; background-color:#f5f5f5; width:5%">19:00</th>
                            <th style="text-align:center; background-color:#f5f5f5; width:5%">20:00</th>
                            <th style="text-align:center; background-color:#f5f5f5; width:5%">21:00</th>
                            <th style="text-align:center; background-color:#f5f5f5; width:5%">22:00</th>
                        </tr>
                        
                        <!--
                                    'sort = Request("sort")
	                        ekspTxt = ""
	
		
		                    '** Søge kriterier **
		                    if cint(visikkemedarbejdere) <> 1 then
		
			                    if usekri = 1 then '**SøgeKri
			                    sqlsearchKri = " (mnavn LIKE '"& thiskri &"%' OR init LIKE '"& thiskri &"%' OR (mnr LIKE '"& thiskri &"%' AND mnr <> '0'))" 
			                    else
			                    sqlsearchKri = " (mid <> 0)"
			                    end if
			
		                    else
		                    sqlsearchKri = " (mid = 0)"
		                    end if

                            if cint(vispasogluk) = 1 then
                            vispasoglukKri = " AND mansat <> -1"
                            else
                            vispasoglukKri = " AND mansat = 1"
                            end if

                                if level <> 1 then 
                                strSQLAdminRights = " AND mid = " & session("mid") 
                                end if
                                -->
                



                           <!-- strSQL = "SELECT mnavn, mid, email, lastlogin, init, mnr, mansat, brugergruppe, medarbejdertype, mt.type AS mtypenavn, b.navn AS brugergruppenavn FROM medarbejdere AS m "_
                            &"LEFT JOIN medarbejdertyper AS mt ON (mt.id = m.medarbejdertype) "_
                            &"LEFT JOIN brugergrupper AS b ON (b.id = m.brugergruppe) WHERE "& sqlsearchKri &" "& vispasoglukKri &" "& strSQLAdminRights &" ORDER BY mnavn LIMIT 4000" 
        
                            'response.write strSQL
                            'response.flush

                            oRec.open strSQL, oConn, 3
        
                            while not oRec.EOF
                    
                                if len(trim(request("jq_visalle"))) <> 0 then
                                jq_vispasluk = request("jq_visalle")
                                else
                                jq_vispasluk = 0
                                end if

                                if jq_vispasluk <> "1" then
                                visAlleSQLval = " AND mansat = 1 "
                                else
                                visAlleSQLval = " AND mansat <> -1 "
                                end if

                                select case oRec("mansat")
                                case 3
                                    mStatus = "Passiv"
                                case 1
                                    mStatus = "Aktiv"
                                case 2
                                    mStatus = "De-aktiv"
                                end select
                             
                                mtypenavn = oRec("mtypenavn")
                                mBrugergruppe = oRec("brugergruppenavn")
                        -->
                        <!-- Da kalenderen går fra 6-22, er der 17 timer på en dag, de 17 timer er 100%
                             Hvis man har arbejdet 8,5 er kallenderen altså kun 50% færdig, selv om man måske kun skal arbejde 8,5 time 
                            Så vi skal udregne hvor mange timer man har brugt på hver opgave og fidne procenten ud fra de 17 timer -->
                        <%
                            starttid = 11
                            sluttid = 12.25

                            aktptimer = sluttid - starttid

                            aktiprocent = aktptimer / 17 * 100

                        %>
                       <!-- <%=aktptimer %> <br />
                        <%=aktiprocent %> -->
                        <tr>
                            <td style="width:10%; vertical-align:middle"><!--=left(oRec("mnavn"), 20)--><div style="font-size:110%">Oliver Storm</div></td>

                            <td style="letter-spacing:-4px;">
                                
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:25%; height:40px; background-color:#fa573c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;"></div>
                                        <div style="text-align:right">             
                                            <span style="margin-right:12px; font-size:110%; display:none; color:#404040" class="fa fa-exclamation-circle"></span>
                                            <span style="margin-right:7px; font-size:110%; display:none; color:#404040" class="fa fa-refresh"></span>
                                        </div>   
                                    </div>
                                </a>
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:25%; height:40px; background-color:#fa573c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;"></div>
                                        <div style="text-align:right">             
                                            <span style="margin-right:12px; font-size:110%; display:none; color:#404040" class="fa fa-exclamation-circle"></span>
                                            <span style="margin-right:7px; font-size:110%; display:none; color:#404040" class="fa fa-refresh"></span>
                                        </div>     
                                    </div>
                                </a>
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:25%; height:40px; background-color:#fa573c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;"></div>
                                        <div style="text-align:right">             
                                            <span style="margin-right:12px; font-size:110%; display:none; color:#404040" class="fa fa-exclamation-circle"></span>
                                            <span style="margin-right:7px; font-size:110%; display:none; color:#404040" class="fa fa-refresh"></span>
                                        </div>    
                                    </div>
                                </a>
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:25%; height:40px; background-color:#fa573c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;"></div>
                                        <div style="text-align:right">             
                                            <span style="margin-right:12px; font-size:110%; display:none; color:#404040" class="fa fa-exclamation-circle"></span>
                                            <span style="margin-right:7px; font-size:110%; display:none; color:#404040" class="fa fa-refresh"></span>
                                        </div>    
                                    </div>
                                </a>
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#b3dc6c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;">Sup..</div>
                                        <div style="text-align:right">             
                                            <span style="margin-right:12px; font-size:110%; display:normal; color:#404040" class="fa fa-exclamation-circle"></span>
                                            <span style="margin-right:7px; font-size:110%; display:normal; color:#404040" class="fa fa-refresh"></span>
                                        </div>        
                                    </div>
                                </a>

                            </td>
                            <td style="letter-spacing:-4px;">
                                
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#fa573c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;"></div>     
                                    </div>
                                </a> <!-- if small --->
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:50%; height:40px; background-color:#b3dc6c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;">S..</div>
                                        <div style="text-align:left">             
                                            <span style="margin-right:9px; font-size:110%; visibility:inherit; color:#404040" class="fa fa-exclamation-circle"></span>
                                            <span style="margin-right:5px; font-size:110%; visibility:inherit; color:#404040" class="fa fa-refresh"></span>
                                        </div>     
                                    </div>
                                </a> <!-- end if --->
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:50%; height:40px; background-color:#b3dc6c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;">S..</div>
                                        <div style="text-align:left">             
                                            <span style="margin-right:9px; font-size:110%; visibility:hidden; color:#404040" class="fa fa-exclamation-circle"></span>
                                            <span style="margin-right:5px; font-size:110%; visibility:hidden; color:#404040" class="fa fa-refresh"></span>
                                        </div>     
                                    </div>
                                </a>

                            </td>
                            <td style="letter-spacing:-4px;">
                                
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#fa573c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;"></div>                                             
                                    </div>
                                </a>
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#b3dc6c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;">Sup..</div>  
                                        <div style="text-align:right">             
                                            <span style="font-size:110%; margin-right:12px; display:normal; color:#404040" class="fa fa-exclamation-circle"></span>
                                            <span style="font-size:110%; margin-right:7px; display:normal; color:#404040" class="fa fa-refresh"></span>
                                        </div>    
                                    </div>
                                </a>

                            </td>
                            <td style="letter-spacing:-4px;">
                                
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#fa573c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;"></div>     
                                    </div>
                                </a>
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#b3dc6c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;">Sup..</div>     
                                    </div>
                                </a>

                            </td>
                            <td style="letter-spacing:-4px;">
                                
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#fa573c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;"></div>     
                                    </div>
                                </a>
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#b3dc6c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;">Sup..</div>     
                                    </div>
                                </a>

                            </td>
                            <td style="letter-spacing:-4px;">
                                
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#fa573c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;"></div>     
                                    </div>
                                </a>
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#b3dc6c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;">Sup..</div>     
                                    </div>
                                </a>

                            </td>
                            <td style="letter-spacing:-4px;">
                                
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#fa573c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;"></div>     
                                    </div>
                                </a>
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#b3dc6c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;">Sup..</div>     
                                    </div>
                                </a>

                            </td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td></td>
                                           
                        </tr>

                        <tr>
                            <td style="width:10%; vertical-align:middle"><!--=left(oRec("mnavn"), 20)--><div style="font-size:110%">Jesper Aaberg</div></td>

                            <td style="letter-spacing:-4px;">
                                
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#498be7; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;">Pro..</div>
                                        <div style="text-align:right">             
                                            <span style="margin-right:12px; font-size:110%; display:normal; color:#404040" class="fa fa-exclamation-circle"></span>
                                            <span style="margin-right:7px; font-size:110%; display:normal; color:#404040" class="fa fa-refresh"></span>
                                        </div>      
                                    </div>
                                </a>
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#498be7; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;">Sup..</div>
                                        <div style="text-align:right">             
                                            <span style="margin-right:12px; font-size:110%; display:normal; color:#404040" class="fa fa-exclamation-circle"></span>
                                            <span style="margin-right:7px; font-size:110%; display:normal; color:#404040" class="fa fa-refresh"></span>
                                        </div>      
                                    </div>
                                </a>

                            </td>
                            <td style="letter-spacing:-4px;">
                                
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#498be7; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;">Pro..</div>     
                                    </div>
                                </a>
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#498be7; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;">Sup..</div>
                                        <div style="text-align:right">             
                                            <span style="margin-right:12px; font-size:110%; display:normal; color:#404040" class="fa fa-exclamation-circle"></span>
                                            <span style="margin-right:7px; font-size:110%; display:normal; color:#404040" class="fa fa-refresh"></span>
                                        </div>      
                                    </div>
                                </a>

                            </td>
                            <td style="letter-spacing:-4px;">
                                
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;">Pro..</div>     
                                    </div>
                                </a>
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;">Sup..</div>     
                                    </div>
                                </a>

                            </td>
                            <td style="letter-spacing:-4px;">
                                
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:50%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;">E..</div>     
                                    </div>
                                </a>
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:50%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;">S..</div>     
                                    </div>
                                </a>

                            </td>
                            <td style="letter-spacing:-4px;">
                                
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;"></div>     
                                    </div>
                                </a>
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;">Sup..</div>     
                                    </div>
                                </a>

                            </td>
                            <td style="letter-spacing:-4px;">
                                
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:100%;"></div>     
                                    </div>
                                </a>
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;">Sup..</div>     
                                    </div>
                                </a>

                            </td>
                            <td style="letter-spacing:-4px;">
                                
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;"></div>     
                                    </div>
                                </a>
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;">Sup..</div>     
                                    </div>
                                </a>

                            </td>
                            <td style="letter-spacing:-4px;">
                                
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#b3dc6c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:100%;"></div>     
                                    </div>
                                </a>
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#b3dc6c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:100%;"></div>     
                                    </div>
                                </a>

                            </td>
                            <td style="letter-spacing:-4px;">
                                
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#b3dc6c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;">Ou..</div>     
                                    </div>
                                </a>
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#b3dc6c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%;">Su..</div>     
                                    </div>
                                </a>

                            </td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td></td>
                                           
                        </tr>
                        
                        <!--
                            oRec.movenext
                            wend
                      
                            oRec.close 
                        -->
                    </table>

              <!--       <br /><br /><br /><br />

                    <div class="row">
                        <div class="col-lg-2" style="background-color:#0086b3;"><b>Fakturerbar tid</b></div>
                    </div>
                     <div class="row">
                        <div class="col-lg-2" style="background-color:#e6b800;"><b>Ikke fakturerbar tid</b></div>
                    </div>
                    <br />
                    <div class="row">
                        <div class="col-lg-2" style="background-color:#00b33c;"><b>Ferie</b></div>
                    </div>
                    <div class="row">
                        <div class="col-lg-2" style="background-color:#e60000;"><b>Sygdom</b></div>
                    </div> -->


                             
                   <%
                       case "week"
                   %>

                    
              </div> <!-- /.form-group -->
                                       
                    <!----------------------------------------------- Uge -------------------------------------------------------------------------------------->
                    
                    
                    <table class="table dataTable table-bordered ui-datatable test">
                       
                        <tr>
                            <th style="background-color:#f5f5f5; vertical-align:middle; border-top:hidden;"s>
                                <select class="form-control input-small">
                                    <option value="0">Sorter efter</option>
                                    <option value="1">Kunde/Job</option>
                                    <option value="2">Gruppe</option>
                                    <option value="3">Medarbejder</option>
                                </select>
                            </th>
                            <th colspan="29" style="text-align:center; border-top:hidden; background-color:#f5f5f5; font-size:175%"><a href="#"><</a> &nbsp Uge 1 - Januar 2016 &nbsp <a href="#">></a> 
                                <div class="pull-right">
                                    <a href="ressplan.asp?func=day" class="btn btn-default btn-sm"><b>Dag</b></a>
                                    <a href="ressplan.asp?func=week" class="btn btn-secondary btn-sm"><b>Uge</b></a>
                                    <a href="ressplan.asp?" class="btn btn-default btn-sm"><b>Måned</b></a>
                                </div>
                                
                            </th>
                        </tr>

                        <tr>
                            <th style="background-color:#f5f5f5">
                                <select class="form-control input-small">
                                      <option value="0">Vis type</option>
                                      <option value="1">Projekt tid</option>
                                      <option value="2">Fravær</option>
                                    <option value="3">Vigtige</option>
                                </select>
                            </th>
                            <th style="text-align:center; background-color:#f5f5f5; width:13%"><a href="ressplan.asp?func=day" style="text-decoration:none; color:inherit">Mandag</a></th>
                            <th style="text-align:center; background-color:#f5f5f5; width:13%">Tirsdag</th>
                            <th style="text-align:center; background-color:#f5f5f5; width:13%">Onsdag</th>
                            <th style="text-align:center; background-color:#f5f5f5; width:13%">Torsdag</th>
                            <th style="text-align:center; background-color:#f5f5f5; width:13%">Fredag</th>
                            <th style="text-align:center; background-color:lightgray; width:13%">Lørdag</th>
                            <th style="text-align:center; background-color:lightgray; width:13%">Søndag</th>
                        </tr>

                        <tr>
                            <th style="background-color:#f5f5f5">
                                <select class="form-control input-small">
                                      <option value="0">Vis gruppe</option>
                                      <option value="1">Design 2014()</option>
                                      <option value="2">Timetag</option>
                                </select>
                            </th>
                            <th style="text-align:center; background-color:#f5f5f5">1</th>
                            <th style="text-align:center; background-color:#f5f5f5">2</th>
                            <th style="text-align:center; background-color:#f5f5f5">3</th>
                            <th style="text-align:center; background-color:#f5f5f5">4</th>
                            <th style="text-align:center; background-color:#f5f5f5">5</th>
                            <th style="text-align:center; background-color:lightgray">6</th>
                            <th style="text-align:center; background-color:lightgray;">7</th>
                        </tr>
                        
                       

                        <!--
                                    'sort = Request("sort")
	                        ekspTxt = ""
	
		
		                    '** Søge kriterier **
		                    if cint(visikkemedarbejdere) <> 1 then
		
			                    if usekri = 1 then '**SøgeKri
			                    sqlsearchKri = " (mnavn LIKE '"& thiskri &"%' OR init LIKE '"& thiskri &"%' OR (mnr LIKE '"& thiskri &"%' AND mnr <> '0'))" 
			                    else
			                    sqlsearchKri = " (mid <> 0)"
			                    end if
			
		                    else
		                    sqlsearchKri = " (mid = 0)"
		                    end if

                            if cint(vispasogluk) = 1 then
                            vispasoglukKri = " AND mansat <> -1"
                            else
                            vispasoglukKri = " AND mansat = 1"
                            end if

                                if level <> 1 then 
                                strSQLAdminRights = " AND mid = " & session("mid") 
                                end if
                                %>
                



                            strSQL = "SELECT mnavn, mid, email, lastlogin, init, mnr, mansat, brugergruppe, medarbejdertype, mt.type AS mtypenavn, b.navn AS brugergruppenavn FROM medarbejdere AS m "_
                            &"LEFT JOIN medarbejdertyper AS mt ON (mt.id = m.medarbejdertype) "_
                            &"LEFT JOIN brugergrupper AS b ON (b.id = m.brugergruppe) WHERE "& sqlsearchKri &" "& vispasoglukKri &" "& strSQLAdminRights &" ORDER BY mnavn LIMIT 4000" 
        
                            'response.write strSQL
                            'response.flush

                            oRec.open strSQL, oConn, 3
        
                            while not oRec.EOF
                    
                                if len(trim(request("jq_visalle"))) <> 0 then
                                jq_vispasluk = request("jq_visalle")
                                else
                                jq_vispasluk = 0
                                end if

                                if jq_vispasluk <> "1" then
                                visAlleSQLval = " AND mansat = 1 "
                                else
                                visAlleSQLval = " AND mansat <> -1 "
                                end if

                                select case oRec("mansat")
                                case 3
                                    mStatus = "Passiv"
                                case 1
                                    mStatus = "Aktiv"
                                case 2
                                    mStatus = "De-aktiv"
                                end select

                                mtypenavn = oRec("mtypenavn")
                                mBrugergruppe = oRec("brugergruppenavn")
                        -->
                        
                        <tr>
                            <td style="width:10%; vertical-align:middle;"><!-- =left(oRec("mnavn"), 20) --><div style="font-size:110%">Oliver Storm</div></td>
                            
                            <td style="letter-spacing:-4px;">
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:50%; height:40px; background-color:#498be7; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Outzour... </div>
                                        <div style="text-align:right">             
                                            <span style="margin-right:12px; font-size:100%; visibility:inherit; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="margin-right:7px; font-size:100%; visibility:inherit; color:#595959" class="fa fa-refresh"></span>
                                        </div>
                                    </div>
                                </a>

                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:50%; height:40px; background-color:#b3dc6c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Kia</div>     
                                        <div style="text-align:right">             
                                            <span style="margin-right:12px; font-size:110%; visibility:hidden; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="margin-right:7px; font-size:110%; visibility:hidden; color:#595959" class="fa fa-refresh"></span>
                                        </div>
                                    </div>
                                </a>

                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#b3dc6c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">VW Aps</div>     
                                        <div style="text-align:right">             
                                            <span style="margin-right:12px; font-size:110%; visibility:hidden; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="margin-right:7px; font-size:110%; visibility:hidden; color:#595959" class="fa fa-refresh"></span>
                                         </div>
                                    </div>
                                </a>
                                

                            </td>
                                
                            <td style="letter-spacing:-4px;">
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:50%; height:40px; background-color:#b3dc6c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Kia</div>   
                                        <div style="text-align:right">             
                                            <span style="margin-right:12px; font-size:110%; display:none; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="margin-right:7px; font-size:110%; display:none; color:#595959" class="fa fa-refresh"></span>
                                        </div>
                                    </div>                                   
                                </a>

                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:50%; height:40px; background-color:#b3dc6c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Kia</div>     
                                        <div style="text-align:right">             
                                            <span style="margin-right:12px; font-size:110%; display:none; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="margin-right:7px; font-size:110%; display:none; color:#595959" class="fa fa-refresh"></span>
                                        </div>
                                    </div>
                                </a>
                            </td>

                            <td style="letter-spacing:-4px;">
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:25%; height:40px; background-color:#498be7; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Ou..</div>     
                                        <div style="text-align:right">             
                                            <span style="margin-right:12px; font-size:110%; visibility:hidden; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="margin-right:7px; font-size:110%; visibility:hidden; color:#595959" class="fa fa-refresh"></span>
                                        </div>
                                    </div>
                                </a>

                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:50%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Ford</div>     
                                        <div style="text-align:right">             
                                            <span style="margin-right:12px; font-size:110%; visibility:hidden; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="margin-right:7px; font-size:110%; visibility:hidden; color:#595959" class="fa fa-refresh"></span>
                                        </div>
                                    </div>
                                </a>

                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:25%; height:40px; background-color:#fa573c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">BM..</div>     
                                        <div style="text-align:right">             
                                            <span style="margin-right:12px; font-size:110%; visibility:inherit; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="margin-right:7px; font-size:110%; visibility:inherit; color:#595959" class="fa fa-refresh"></span>
                                        </div>
                                    </div>
                                </a>
                            </td>

                            <td></td>

                            <td style="letter-spacing:-4px;">
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:25%; height:40px; background-color:#498be7; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Ou..</div>     
                                        <div style="text-align:right">             
                                            <span style="margin-right:12px; font-size:110%; visibility:hidden; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="margin-right:7px; font-size:110%; visibility:hidden; color:#595959" class="fa fa-refresh"></span>
                                        </div>
                                    </div>
                                </a>

                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:50%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Ford</div>     
                                        <div style="text-align:right">             
                                            <span style="margin-right:12px; font-size:110%; visibility:hidden; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="margin-right:7px; font-size:110%; visibility:hidden; color:#595959" class="fa fa-refresh"></span>
                                        </div>
                                    </div>
                                </a>

                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:25%; height:40px; background-color:#fa573c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">BM..</div>     
                                        <div style="text-align:right">             
                                            <span style="margin-right:12px; font-size:110%; visibility:inherit; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="margin-right:7px; font-size:110%; visibility:inherit; color:#595959" class="fa fa-refresh"></span>
                                        </div>
                                    </div>
                                </a>
                            </td>

                            <td></td>

                            <td></td>



                            
                        </tr>
                            <tr>
                                <td style="width:10%; vertical-align:middle;"><!-- =left(oRec("mnavn"), 20) --><div style="font-size:110%">Jesper Aaberg</div></td>
                                                                                                                
                                <td style="letter-spacing:-4px;">
                                
                                    <a data-toggle="modal" href="#basicModal" style="color:black;">
                                        <div style="width:100%; height:40px; background-color:#fa573c; display:inline-block; border:1px solid; border-color:#fff">
                                            <div style="letter-spacing:0px; font-size:110%;">&nbsp</div>     
                                            <div style="text-align:right">             
                                                <span style="margin-right:12px; font-size:110%; display:none; color:#595959" class="fa fa-exclamation-circle"></span>
                                                <span style="margin-right:7px; font-size:110%; display:none; color:#595959" class="fa fa-refresh"></span>
                                            </div>
                                        </div>
                                    </a>

                                </td>

                                <td style="letter-spacing:-4px;">
                                
                                    <a data-toggle="modal" href="#basicModal" style="color:black;">
                                        <div style="width:100%; height:40px; background-color:#fa573c; display:inline-block; border:1px solid; border-color:#fff">
                                            <div style="letter-spacing:0px; font-size:110%;">&nbsp</div>     
                                            <div style="text-align:right">             
                                                <span style="margin-right:12px; font-size:110%; display:none; color:#595959" class="fa fa-exclamation-circle"></span>
                                                <span style="margin-right:7px; font-size:110%; display:none; color:#595959" class="fa fa-refresh"></span>
                                            </div>
                                        </div>
                                    </a>

                                </td>

                                <td style="letter-spacing:-4px;">
                                
                                    <a data-toggle="modal" href="#basicModal" style="color:black;">
                                        <div style="width:100%; height:40px; background-color:#fa573c; display:inline-block; border:1px solid; border-color:#fff">
                                            <div style="letter-spacing:0px; font-size:110%;">&nbsp</div>     
                                            <div style="text-align:right">             
                                                <span style="margin-right:12px; font-size:110%; display:none; color:#595959" class="fa fa-exclamation-circle"></span>
                                                <span style="margin-right:7px; font-size:110%; display:none; color:#595959" class="fa fa-refresh"></span>
                                            </div>
                                        </div>
                                    </a>

                                </td>

                                <td style="letter-spacing:-4px;">
                                
                                    <a data-toggle="modal" href="#basicModal" style="color:black;">
                                        <div style="width:100%; height:40px; background-color:#fa573c; display:inline-block; border:1px solid; border-color:#fff">
                                            <div style="letter-spacing:0px; font-size:110%;">&nbsp</div>     
                                            <div style="text-align:right">             
                                                <span style="margin-right:12px; font-size:110%; display:none; color:#595959" class="fa fa-exclamation-circle"></span>
                                                <span style="margin-right:7px; font-size:110%; display:none; color:#595959" class="fa fa-refresh"></span>
                                            </div>
                                        </div>
                                    </a>

                                </td>

                                <td style="letter-spacing:-4px;">
                                
                                    <a data-toggle="modal" href="#basicModal" style="color:black;">
                                        <div style="width:100%; height:40px; background-color:#fa573c; display:inline-block; border:1px solid; border-color:#fff">
                                            <div style="letter-spacing:0px; font-size:110%;">&nbsp</div>     
                                            <div style="text-align:right">             
                                                <span style="margin-right:12px; font-size:110%; display:none; color:#595959" class="fa fa-exclamation-circle"></span>
                                                <span style="margin-right:7px; font-size:110%; display:none; color:#595959" class="fa fa-refresh"></span>
                                            </div>
                                        </div>
                                    </a>

                                </td>

                                <td></td>
                                <td></td>

                                                                                   
                            </tr>

                            <tr>
                                <td style="width:10%; vertical-align:middle;"><!-- =left(oRec("mnavn"), 20) --><div style="font-size:110%">Søren K</div></td>
                                                                                                                
                                <td style="letter-spacing:-4px;">
                                
                                    <a data-toggle="modal" href="#basicModal" style="color:black;">
                                        <div style="width:100%; height:40px; background-color:#498be7; display:inline-block; border:1px solid; border-color:#fff">
                                            <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Produktion</div>   
                                            <div style="text-align:right">             
                                                <span style="margin-right:12px; font-size:110%; display:none; color:#595959" class="fa fa-exclamation-circle"></span>
                                                <span style="margin-right:7px; font-size:110%; display:none; color:#595959" class="fa fa-refresh"></span>
                                            </div>
                                        </div>
                                    </a>

                                </td>

                                <td style="letter-spacing:-4px;">
                                
                                    <a data-toggle="modal" href="#basicModal" style="color:black;">
                                        <div style="width:100%; height:40px; background-color:#498be7; display:inline-block; border:1px solid; border-color:#fff">
                                            <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Produktion</div>  
                                            <div style="text-align:right">             
                                                <span style="margin-right:12px; font-size:110%; display:none; color:#595959" class="fa fa-exclamation-circle"></span>
                                                <span style="margin-right:7px; font-size:110%; display:none; color:#595959" class="fa fa-refresh"></span>
                                            </div>
                                        </div>
                                    </a>

                                </td>

                                <td style="letter-spacing:-4px;">
                                
                                    <a data-toggle="modal" href="#basicModal" style="color:black;">
                                        <div style="width:100%; height:40px; background-color:#498be7; display:inline-block; border:1px solid; border-color:#fff">
                                            <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Produktion</div>    
                                            <div style="text-align:right">             
                                                <span style="margin-right:12px; font-size:110%; display:none; color:#595959" class="fa fa-exclamation-circle"></span>
                                                <span style="margin-right:7px; font-size:110%; display:none; color:#595959" class="fa fa-refresh"></span>
                                            </div>
                                        </div>
                                    </a>
                                    <a data-toggle="modal" href="#basicModal" style="color:black;">
                                        <div style="width:100%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                            <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Support</div>     
                                            <div style="text-align:right">             
                                                <span style="margin-right:12px; font-size:110%; display:normal; color:#595959" class="fa fa-exclamation-circle"></span>
                                                <span style="margin-right:7px; font-size:110%; display:normal; color:#595959" class="fa fa-refresh"></span>
                                            </div>
                                        </div>
                                    </a>
                                    <a data-toggle="modal" href="#basicModal" style="color:black;">
                                        <div style="width:100%; height:40px; background-color:#498be7; display:inline-block; border:1px solid; border-color:#fff">
                                            <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">EPI</div>   
                                            <div style="text-align:right">             
                                                <span style="margin-right:12px; font-size:110%; display:none; color:#595959" class="fa fa-exclamation-circle"></span>
                                                <span style="margin-right:7px; font-size:110%; display:normal; color:#595959" class="fa fa-refresh"></span>
                                            </div>
                                        </div>
                                    </a>

                                </td>

                                <td style="letter-spacing:-4px;">
                                
                                    <a data-toggle="modal" href="#basicModal" style="color:black;">
                                        <div style="width:100%; height:40px; background-color:#498be7; display:inline-block; border:1px solid; border-color:#fff">
                                            <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Produktion</div>    
                                            <div style="text-align:right">             
                                                <span style="margin-right:12px; font-size:110%; display:normal; color:#595959" class="fa fa-exclamation-circle"></span>
                                                <span style="margin-right:7px; font-size:110%; display:none; color:#595959" class="fa fa-refresh"></span>
                                            </div>
                                        </div>
                                    </a>

                                </td>

                                <td style="letter-spacing:-4px;">
                                
                                    <a data-toggle="modal" href="#basicModal" style="color:black;">
                                        <div style="width:100%; height:40px; background-color:#498be7; display:inline-block; border:1px solid; border-color:#fff">
                                            <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Produktion</div>     
                                            <div style="text-align:right">             
                                                <span style="margin-right:12px; font-size:110%; display:normal; color:#595959" class="fa fa-exclamation-circle"></span>
                                                <span style="margin-right:7px; font-size:110%; display:normal; color:#595959" class="fa fa-refresh"></span>
                                            </div>
                                        </div>
                                    </a>

                                </td>

                                <td></td>
                                <td></td>

                                                                                   
                            </tr>

                            

                            
                        
                        
                        

                        <!--
                            oRec.movenext
                            wend
                      
                            oRec.close 
                        -->

                    </table>
                    
                   <!-- <br /><br /><br /><br />

                    <div class="row">
                        <div class="col-lg-2" style="background-color:#498be7;"><b>Fakturerbar tid</b></div>
                    </div>
                     <div class="row">
                        <div class="col-lg-2" style="background-color:#fbe983;"><b>Ikke fakturerbar tid</b></div>
                    </div>
                    <br />
                    <div class="row">
                        <div class="col-lg-2" style="background-color:#b3dc6c;"><b>Ferie</b></div>
                    </div>
                    <div class="row">
                        <div class="col-lg-2" style="background-color:#fa573c;"><b>Sygdom</b></div>
                    </div> -->






                  <%   
                      case else 
                  %>
                    <!------------------------------------------------ Måned ------------------------------------------------------------------------------------>


                   <!--<table>
                        <tr>
                            <th>test</th>
                        </tr> -->

                        <%
	                    strSQL = "SELECT id, navn, opengp, orgvir FROM projektgrupper WHERE id > 1 ORDER BY navn"
	                    oRec.open strSQL, oConn, 3
	
	                    x = 0
	                    n = 0
	                    while not oRec.EOF


                            	x = 0
				                call antalMediPgrp(oRec("id"))
				
				                'if antalMediPgrpX > 0 then
				                'x = antalMediPgrpX
				                'Antal = x
				                'else
				                'x = 0
				                'Antal = x
				                'end if
	
	                            antalMediPgrpX = 0

                                call fTeamleder(session("mid"), oRec("id"))

                        if cint(lastid) = cint(oRec("id")) then
                       trBgCol = "#FFFFE1" 
                       else
                       trBgCol = ""
                       end if%>

                        <!--<tr>
                            <td> <a href="ressplan.asp?id=<%=oRec("id") %>"><%=oRec("navn") %></a></td>
                        </tr> -->
                        <%

                        x = 0
	                    n = n + 1
                        oRec.movenext
                        wend
                        oRec.close
                        %>
                   <!-- </table> -->

                   

                    <table class="table dataTable table-striped table-bordered ui-datatable">
                       <!-- <thead> -->
                            <tr> 
                                <th style="border:hidden; background-color:#f5f5f5">&nbsp</th>                                                            
                                <th colspan="72" style="text-align:center; border-top:hidden; background-color:#f5f5f5; font-size:175%"><a href="#"><</a> &nbsp Januar 2016 &nbsp <a href="#">></a> 
                                    <div class="pull-right">
                                        <a href="ressplan.asp?func=day" class="btn btn-default btn-sm"><b>Dag</b></a>
                                        <a href="ressplan.asp?func=week" class="btn btn-default btn-sm"><b>Uge</b></a>
                                        <a href="ressplan.asp?" class="btn btn-secondary btn-sm"><b>Måned</b></a>
                                    </div>
                                </th>
                            </tr>

                            

                           <!-- <tr>
                                <th style="border-right-color:black; border-top:hidden">&nbsp</th>
                                <th colspan="14" style="text-align:center; border-right-color:black; border-top-color:black;">Uge 1</th>
                                <th colspan="14" style="text-align:center;">Uge 2</th>
                                <th colspan="14" style="text-align:center;">Uge 3</th>
                                <th colspan="14" style="text-align:center;">Uge 4</th>
                                <th colspan="14" style="text-align:center;">Uge 5</th>
                            </tr> -->

                       <!-- </thead> -->
                      <!--  <tbody>  -->
                            
                            <tr>
                                <th style="background-color:#f5f5f5; vertical-align:middle; width:5%">
                                    <input type="text" class="form-control input-small" placeholder="Søg kunde/job" value="">
                                </th>
                                
                                <th colspan="7" style="text-align:center; background-color:#f5f5f5"><a href="ressplan.asp?func=week" style="color:dimgray">Uge 1</a></th>
                                <th colspan="7" style="text-align:center; background-color:#f5f5f5">Uge 2</th>
                                <th colspan="7" style="text-align:center; background-color:#f5f5f5">Uge 3</th>
                                <th colspan="7" style="text-align:center; background-color:#f5f5f5">Uge 4</th>
                                <th colspan="7" style="text-align:center; background-color:#f5f5f5">Uge 5</th>
                            </tr>                     
                            
                            <tr>
                                <th style="width:3%; background-color:#f5f5f5">
                                    <select class="form-control input-small">
                                      <option value="0">Vis type</option>
                                      <option value="1">Projekt tid</option>
                                      <option value="2">Fravær</option>
                                        <option value="3">Vigtige</option>
                                    </select>
                                </th>
                                <th style="width:2%; background-color:#f5f5f5; text-align:center"><a href="ressplan.asp?func=day" style="color:dimgray">Man</a></th>
                                <th style="width:2%; text-align:center">Tir</th>
                                <th style="width:2%; background-color:#f5f5f5; text-align:center"><a href="ressplan.asp?func=day" style="color:dimgray">Ons</a></th>
                                <th style="width:2%; background-color:#f5f5f5; text-align:center">Tor</th>
                                <th style="width:2%; background-color:#f5f5f5; text-align:center">Fre</th>
                                <th style="width:2%; background-color:lightgray; text-align:center">Lør</th>
                                <th style="width:2%; background-color:lightgray; text-align:center">Søn</th>

                                <th style="width:2%; background-color:#f5f5f5; text-align:center">Man</th>
                                <th style="width:2%; background-color:#f5f5f5; text-align:center">Tir</th>
                                <th style="width:2%; background-color:#f5f5f5; text-align:center">Ons</th>
                                <th style="width:2%; background-color:#f5f5f5; text-align:center">Tor</th>
                                <th style="width:2%; background-color:#f5f5f5; text-align:center">Fre</th>
                                <th style="width:2%; background-color:lightgray; text-align:center">Lør</th>
                                <th style="width:2%; background-color:lightgray; text-align:center">Søn</th>

                                <th style="width:2%; background-color:#f5f5f5; text-align:center">Man</th>
                                <th style="width:2%; background-color:#f5f5f5; text-align:center">Tir</th>
                                <th style="width:2%; background-color:#f5f5f5; text-align:center">Ons</th>
                                <th style="width:2%; background-color:#f5f5f5; text-align:center">Tor</th>
                                <th style="width:2%; background-color:#f5f5f5; text-align:center">Fre</th>
                                <th style="width:2%; background-color:lightgray; text-align:center">Lør</th>
                                <th style="width:2%; background-color:lightgray; text-align:center">Søn</th>

                                <th style="width:2%; background-color:#f5f5f5; text-align:center">Man</th>
                                <th style="width:2%; background-color:#f5f5f5; text-align:center">Tir</th>
                                <th style="width:2%; background-color:#f5f5f5; text-align:center">Ons</th>
                                <th style="width:2%; background-color:#f5f5f5; text-align:center">Tor</th>
                                <th style="width:2%; background-color:#f5f5f5; text-align:center">Fre</th>
                                <th style="width:2%; background-color:lightgray; text-align:center">Lør</th>
                                <th style="width:2%; background-color:lightgray; text-align:center">Søn</th>

                                <th style="width:2%; background-color:#f5f5f5; text-align:center">Man</th>
                                <th style="width:2%; background-color:#f5f5f5; text-align:center">Tir</th>
                                <th style="width:2%; background-color:#f5f5f5; text-align:center">Ons</th>

                                
                            </tr>
                            <tr>
                                <th style="background-color:#f5f5f5">
                                    <select class="form-control input-small">
                                      <option value="0">Vælg gruppe</option>
                                      <option value="1">Design 2014()</option>
                                      <option value="2">Timetag</option>
                                    </select>
                                </th>
                                <th style="width:1%; text-align:center; background-color:#f5f5f5;">1</th>
                                <th style="width:1%; text-align:center; background-color:#f5f5f5;">2</th>
                                <th style="width:1%; text-align:center; background-color:#f5f5f5;">3</th>
                                <th style="width:1%; text-align:center; background-color:#f5f5f5;">4</th>
                                <th style="width:1%; text-align:center; background-color:#f5f5f5;">5</th>
                                <th style="width:1%; text-align:center; background-color:lightgray;">6</th>
                                <th style="width:1%; text-align:center; background-color:lightgray;">7</th>

                                <th style="width:1%; text-align:center; background-color:#f5f5f5;">8</th>
                                <th style="width:1%; text-align:center; background-color:#f5f5f5;">9</th>
                                <th style="width:1%; text-align:center; background-color:#f5f5f5;">10</th>
                                <th style="width:1%; text-align:center; background-color:#f5f5f5;">11</th>
                                <th style="width:1%; text-align:center; background-color:#f5f5f5;">12</th>
                                <th style="width:1%; text-align:center; background-color:lightgray;">13</th>
                                <th style="width:1%; text-align:center; background-color:lightgray;">14</th>

                                <th style="width:1%; text-align:center; background-color:#f5f5f5;">15</th>
                                <th style="width:1%; text-align:center; background-color:#f5f5f5;">16</th>
                                <th style="width:1%; text-align:center; background-color:#f5f5f5;">17</th>
                                <th style="width:1%; text-align:center; background-color:#f5f5f5;">18</th>
                                <th style="width:1%; text-align:center; background-color:#f5f5f5;">19</th>
                                <th style="width:1%; text-align:center; background-color:lightgray;">20</th>
                                <th style="width:1%; text-align:center; background-color:lightgray;">21</th>

                                <th style="width:1%; text-align:center; background-color:#f5f5f5;">22</th>
                                <th style="width:1%; text-align:center; background-color:#f5f5f5;">23</th>
                                <th style="width:1%; text-align:center; background-color:#f5f5f5;">24</th>
                                <th style="width:1%; text-align:center; background-color:#f5f5f5;">25</th>
                                <th style="width:1%; text-align:center; background-color:#f5f5f5;">26</th>
                                <th style="width:1%; text-align:center; background-color:lightgray;">27</th>
                                <th style="width:1%; text-align:center; background-color:lightgray;">28</th>

                                <th style="width:1%; text-align:center; background-color:#f5f5f5;">29</th>
                                <th style="width:1%; text-align:center; background-color:#f5f5f5;">30</th>
                                <th style="width:1%; text-align:center; background-color:#f5f5f5;">31</th>

                            </tr> 
                            <!--
                                    'sort = Request("sort")
	                        ekspTxt = ""
	
		
		                    '** Søge kriterier **
		                    if cint(visikkemedarbejdere) <> 1 then
		
			                    if usekri = 1 then '**SøgeKri
			                    sqlsearchKri = " (mnavn LIKE '"& thiskri &"%' OR init LIKE '"& thiskri &"%' OR (mnr LIKE '"& thiskri &"%' AND mnr <> '0'))" 
			                    else
			                    sqlsearchKri = " (mid <> 0)"
			                    end if
			
		                    else
		                    sqlsearchKri = " (mid = 0)"
		                    end if

                            if cint(vispasogluk) = 1 then
                            vispasoglukKri = " AND mansat <> -1"
                            else
                            vispasoglukKri = " AND mansat = 1"
                            end if

                                if level <> 1 then 
                                strSQLAdminRights = " AND mid = " & session("mid") 
                                end if
                                -->
                



                            <!--strSQL = "SELECT mnavn, mid, email, lastlogin, init, mnr, mansat, brugergruppe, medarbejdertype, mt.type AS mtypenavn, b.navn AS brugergruppenavn FROM medarbejdere AS m "_
                            &"LEFT JOIN medarbejdertyper AS mt ON (mt.id = m.medarbejdertype) "_
                            &"LEFT JOIN brugergrupper AS b ON (b.id = m.brugergruppe) WHERE "& sqlsearchKri &" "& vispasoglukKri &" "& strSQLAdminRights &" ORDER BY mnavn LIMIT 4000" 
        
                            'response.write strSQL
                            'response.flush

                            oRec.open strSQL, oConn, 3
        
                            while not oRec.EOF
                    
                                if len(trim(request("jq_visalle"))) <> 0 then
                                jq_vispasluk = request("jq_visalle")
                                else
                                jq_vispasluk = 0
                                end if

                                if jq_vispasluk <> "1" then
                                visAlleSQLval = " AND mansat = 1 "
                                else
                                visAlleSQLval = " AND mansat <> -1 "
                                end if

                                select case oRec("mansat")
                                case 3
                                    mStatus = "Passiv"
                                case 1
                                    mStatus = "Aktiv"
                                case 2
                                    mStatus = "De-aktiv"
                                end select

                                mtypenavn = oRec("mtypenavn")
                                mBrugergruppe = oRec("brugergruppenavn")
                                -->

                            <tr>
                                <td style="width:10%; font-size:110%; vertical-align:middle"><!--=left(oRec("mnavn"), 20)--><div style="font-size:110%">Oliver Storm</div></td>
                                

                                <td style="letter-spacing:-4px;">
                                
                                    <a data-toggle="modal" href="#basicModal" style="color:black;">
                                        <div style="width:100%; height:40px; background-color:#b3dc6c; display:inline-block; border:1px solid; border-color:#fff">
                                            <div style="letter-spacing:0px; font-size:110%;">O..</div>
                                                
                                        </div>
                                    </a>
                                    <a data-toggle="modal" href="#basicModal" style="color:black;">
                                        <div style="width:100%; height:40px; background-color:#b3dc6c; display:inline-block; border:1px solid; border-color:#fff">
                                            <div style="letter-spacing:0px; font-size:110%;">P..</div>     
                                        </div>
                                    </a>

                                </td>

                                <td style="letter-spacing:-4px;">
                                
                                    <a data-toggle="modal" href="#basicModal" style="color:black;">
                                        <div style="width:50%; height:40px; background-color:#b3dc6c; display:inline-block; border:1px solid; border-color:#fff">
                                            <div style="letter-spacing:0px; font-size:110%;"></div>     
                                        </div>
                                    </a>
                                    <a data-toggle="modal" href="#basicModal" style="color:black;">
                                        <div style="width:50%; height:40px; background-color:#498be7; display:inline-block; border:1px solid; border-color:#fff">
                                            <div style="letter-spacing:0px; font-size:110%;"></div>     
                                        </div>
                                    </a>
                                    <a data-toggle="modal" href="#basicModal" style="color:black;">
                                        <div style="width:50%; height:40px; background-color:#b3dc6c; display:inline-block; border:1px solid; border-color:#fff">
                                            <div style="letter-spacing:0px; font-size:110%;"></div>     
                                        </div>
                                    </a>
                                    <a data-toggle="modal" href="#basicModal" style="color:black;">
                                        <div style="width:50%; height:40px; background-color:#498be7; display:inline-block; border:1px solid; border-color:#fff">
                                            <div style="letter-spacing:0px; font-size:110%;"></div>     
                                        </div>
                                    </a>

                                </td>
                                <td style="letter-spacing:-4px;">
                                
                                    <a data-toggle="modal" href="#basicModal" style="color:black;">
                                        <div style="width:100%; height:40px; background-color:#fa573c; display:inline-block; border:1px solid; border-color:#fff">
                                            <div style="letter-spacing:0px; font-size:110%;">K..</div>     
                                        </div>
                                    </a>
                                    <a data-toggle="modal" href="#basicModal" style="color:black;">
                                        <div style="width:100%; height:40px; background-color:#fa573c; display:inline-block; border:1px solid; border-color:#fff">
                                            <div style="letter-spacing:0px; font-size:110%;">S..</div>                                                
                                        </div>
                                    </a>

                                </td>
                                <td style="letter-spacing:-4px;">
                                
                                    <a data-toggle="modal" href="#basicModal" style="color:black;">
                                        <div style="width:100%; height:40px; background-color:#b3dc6c; display:inline-block; border:1px solid; border-color:#fff">
                                            <div style="letter-spacing:0px; font-size:110%;">O..</div>     
                                        </div>
                                    </a>

                                </td>
                                <td style="letter-spacing:-4px;">
                                
                                    <a data-toggle="modal" href="#basicModal" style="color:black;">
                                        <div style="width:100%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                            <div style="letter-spacing:0px; font-size:110%;">E..</div>     
                                        </div>
                                    </a>

                                </td>
                                <td style="letter-spacing:-4px;">
                                
                                    <a data-toggle="modal" href="#basicModal" style="color:black;">
                                        <div style="width:100%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                            <div style="letter-spacing:0px; font-size:110%;">E..</div>     
                                        </div>
                                    </a>

                                </td>
                                <td style="letter-spacing:-4px;">
                                
                                    <a data-toggle="modal" href="#basicModal" style="color:black;">
                                        <div style="width:100%; height:40px; background-color:#498be7; display:inline-block; border:1px solid; border-color:#fff">
                                            <div style="letter-spacing:0px; font-size:110%;">E..</div>     
                                        </div>
                                    </a>

                                </td>
                                <td>&nbsp</td>
                                <td>&nbsp</td>
                                <td>&nbsp</td>
                                <td>&nbsp</td>
                                <td>&nbsp</td>
                                <td>&nbsp</td>
                                <td>&nbsp</td>
                                <td>&nbsp</td>
                                <td>&nbsp</td>
                                <td>&nbsp</td>
                                <td>&nbsp</td>
                                <td>&nbsp</td>
                                <td>&nbsp</td>
                                <td>&nbsp</td>
                                <td>&nbsp</td>
                                <td>&nbsp</td>
                                <td>&nbsp</td>
                                <td>&nbsp</td>
                                <td>&nbsp</td>
                                <td>&nbsp</td>
                                <td>&nbsp</td>
                                <td>&nbsp</td>
                                <td>&nbsp</td>
                                <td>&nbsp</td>
                               
                            </tr>
                            
                         
                                                                 
                      
                    </table>

                     <br /><br /><br /><br />

                    
                </div>
            </div>
        </div>
        <%end select %>
    </div>
</div>





<!--#include file="../inc/regular/footer_inc.asp"-->