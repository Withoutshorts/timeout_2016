


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
    


    table th 
    {
        font-size:110%
    }

   


</style>

<!----------------------------------------
------------------------------------------   
------------------------------------------  
------------------------------------------    
    
    
      
    
        Color codes: 

        invoiceble: 9fc6e7

        non-invoiceble: fbe983

        Absence: fa573c

        Holiday: b3dc6c

-------------------------------------------    
-------------------------------------------    
-------------------------------------------    
-------------------------------------------->


<div class="wrapper">
    <div class="content">   
        
        <div class="container" style="width:1600px;">
            <div class="portlet">
                <h3 class="portlet-title"><u>Ressource Planner</u></h3>
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
                                <div class="col-lg-2">From:</div>
                                <div class="col-lg-4"><input type="text" class="form-control input-small" placeholder="00:00" /></div>
                            </div>
                            <div class="row">
                                <div class="col-lg-2">End:</div>
                                <div class="col-lg-4">
                                    <div class='input-group date' id='datepicker_stdato'>
                                        <input type="text" class="form-control input-small" name="" value="" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon input-small">
                                            <span class="fa fa-calendar"></span>
                                        </span>
                                    </div>
                                </div>
                                <div class="col-lg-2">To:</div>
                                <div class="col-lg-4"><input type="text" class="form-control input-small" placeholder="00:00" /></div>
                            </div>

                            <div class="row">
                                <div class="col-lg-2">Recurrence:</div>
                                <div class="col-lg-4">
                                    <select class="form-control input-small" id="test" onchange="func2()">
                                        <option value="dag">Daily</option>
                                        <option value="uge">Weekly</option>
                                        <option value="maaned">Every Month</option>
                                        <option value="maaned">Every 2. Month</option>
                                        <option value="maaned">Every 3. Month</option>
                                        <option value="maaned">Every 6. Month</option>
                                        <option value="maaned" onclick="selmaaned.call">Every Year</option>
                                    </select>
                                </div>
                                <div class="col-lg-2">Important:</div>
                                <div class="col-lg-4"><input type="checkbox" value="0" /></div>
                            </div>
                            <div>
                                <div class="row">
                                    <div class="col-lg-2">End after:</div>
                                    <div class="col-lg-2"><input type="number" name="antalre" value="99" class="form-control input-small" /></div>
                                    <div class="col-lg-2">Recurrences</div>
                                </div>                               
                            </div>
                            <br />
                            <div class="row">
                                <div class="col-lg-2">Employee:</div>
                                <div class="col-lg-4">
                                    <select class="form-control input-small">
                                        <option value="1">Hans</option>
                                        <option value="2">Per</option>
                                        <option value="3">Søren Karlsen</option>
                                    </select>
                                </div>
                                <div class="col-lg-2">Custommer:</div>
                                <div class="col-lg-4"><input type="text" value="Outzource Aps" class="form-control input-small" readonly/></div>
                            </div>
                            <div class="row">
                                <div class="col-lg-2">Heading:</div>
                                <div class="col-lg-4">
                                    <select class="form-control input-small">
                                        <option value="">Activity</option>
                                        <option value="">Project</option>
                                        <option value="">Custommer</option>                                     
                                    </select>
                                </div>

                                <div class="col-lg-2">Project:</div>
                                <div class="col-lg-4"><input type="text" value="Design" class="form-control input-small" readonly/></div>
                            </div>
                            <div class="row">
                                <div class="col-lg-2">Type:</div>
                                <div class="col-lg-4">
                                    <select class="form-control input-small" readonly>
                                        <option valse="1">Invoiceable</option>
                                        <option valse="2">None Invoiceable</option>
                                        <option valse="3">Holdiday</option>
                                        <option valse="4">Illness</option>
                                    </select>
                                </div>
                                
                                <div class="col-lg-2">activity:</div>
                                <div class="col-lg-4"><input type="text" value="Produkion" class="form-control input-small" readonly/></div>
                            </div>

                          
                                
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

                           
                        
            
                        </div> 

                        <div class="modal-footer" style="text-align:left">
                          <button type="button" class="btn btn-default">Opdater</button>
                        </div> 
                      </div> 
                    </div>
                  </div>


                     


                   <%select case func %>   
                    
                    <%
                        case "day" 
                    %>
                    <!-----------------------------------------------  Dag --------------------------------------------------------------------------------------->
                    <table class="table dataTable table-bordered ui-datatable">
                        <tr>
                            <th style="background-color:#f5f5f5; vertical-align:middle; border-top:hidden;">
                                <select class="form-control input-small">
                                    <option value="0">Sort by</option>
                                    <option value="1">Customer</option>
                                    <option value="2">Job</option>
                                    <option value="3">Group</option>
                                    <option value="4">Employees</option>
                                </select>
                            </th>
                            <th colspan="70" style="text-align:center; border-top:hidden; background-color:#f5f5f5; font-size:175%"><a href="#"><</a> &nbsp Monday 01-01-16 &nbsp <a href="#">></a> 
                                    <div class="pull-right">
                                        <a href="ressplan_2017.asp?func=day" class="btn btn-secondary btn-sm"><b>Day</b></a>
                                        <a href="ressplan_2017.asp?func=week" class="btn btn-default btn-sm"><b>Week</b></a>
                                        <a href="ressplan_2017.asp?" class="btn btn-default btn-sm"><b>Month</b></a>
                                    </div>
                            </th>
                        </tr>

                        <tr>
                            <th style="background-color:#f5f5f5">
                                <select class="form-control input-small">
                                      <option value="0">View type</option>
                                      <option value="1">Project time</option>
                                      <option value="2">Absence</option>
                                    <option value="3">Important</option>
                                </select>
                            </th>
                            <th colspan="69" style="background-color:#f5f5f5; border-top:hidden"></th>
                        </tr>

                        <tr>
                            <th style="background-color:#f5f5f5">
                                <select class="form-control input-small">
                                      <!--<option value="0">All</option> all employees should be selected at first by standard  -->                                                                                 
                                      <option value="1">Oliver Storm</option>
                                      <option value="2">Jesper, Aaberg</option>
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
                        
                        
                        <tr>
                            <td style="width:10%; vertical-align:middle"><!--=left(oRec("mnavn"), 20)--><div style="font-size:110%">Oliver Storm</div></td>


                            <td style="letter-spacing:-4px;">
                                
                                <!------ the div tag with a witdh of 25% should look like this ------>
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:25%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; text-align:center;">W</div>
                                           
                                    </div>
                                </a>


                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:25%; height:40px; background-color:#9fc6e7; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; text-align:center;">H</div>
                                            
                                    </div>
                                </a>
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:25%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; text-align:center">I</div>
                                            
                                    </div>
                                </a>
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:25%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; text-align:center;">G</div>
                                           
                                    </div>
                                </a>

                                <!------ the div tag with a witdh of 100% should look like this ------>
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#9fc6e7; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Support</div>
                                        <div style="text-align:right">             
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:5px; display:none; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:3px; display:normal; color:#595959" class="fa fa-refresh"></span>
                                        </div>      
                                    </div>
                                </a>

                            </td>
                            <td style="letter-spacing:-4px;">
                                
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">E</div>  
                                        <div style="text-align:right">             
                                            <span style="margin-right:12px; font-size:100%; visibility:hidden; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="margin-right:7px; font-size:100%; visibility:hidden; color:#595959" class="fa fa-refresh"></span>
                                        </div>   
                                    </div>
                                </a>
                                
                                 
                                <!------ the div tag with a witdh of 75% should look like this ------>
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:75%; height:40px; background-color:#9fc6e7; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">WW..</div>
                                        <div style="text-align:right">             
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:5px; display:normal; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:3px; display:normal; color:#595959" class="fa fa-refresh"></span>
                                        </div>      
                                    </div>
                                </a>



                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#9fc6e7; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Suppo..</div>
                                        <div style="text-align:right">             
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:5px; display:normal; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:3px; display:normal; color:#595959" class="fa fa-refresh"></span>
                                        </div>     
                                    </div> 
                                </a>

                            </td>
                            <td style="letter-spacing:-4px;">
                                
                                <!------ the div tag with a witdh of 33% should look like this ------>
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:33%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; text-align:center;">S</div>
                                        <div style="text-align:right">             
                                            <span style="margin-right:12px; font-size:100%; visibility:hidden; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="margin-right:7px; font-size:100%; visibility:hidden; color:#595959" class="fa fa-refresh"></span>
                                        </div>     
                                    </div>
                                </a>
                                
                                 
                                
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:33%; height:40px; background-color:#9fc6e7; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; text-align:center;">S</div>
                                        <div style="text-align:right">             
                                            <span style="margin-right:12px; font-size:100%; visibility:hidden; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="margin-right:7px; font-size:100%; visibility:hidden; color:#595959" class="fa fa-refresh"></span>
                                        </div>    
                                    </div>
                                </a>



                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:33%; height:40px; background-color:#9fc6e7; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; text-align:center;">P</div>
                                        <div style="text-align:right">             
                                            <span style="margin-right:12px; font-size:100%; visibility:hidden; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="margin-right:7px; font-size:100%; visibility:hidden; color:#595959" class="fa fa-refresh"></span>
                                        </div>    
                                    </div> 
                                </a>

                            </td>
                           
                            
                            
                            
                            <td style="letter-spacing:-4px;">
                                <!------ a div tag with a witdh of 50% should look like this ------>
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:50%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">S..</div>
                                        <div style="text-align:center">             
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:2px; visibility:inherit; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="font-size:100%; letter-spacing:0px; visibility:inherit; color:#595959" class="fa fa-refresh"></span>
                                        </div>     
                                    </div>
                                </a>
                                
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:50%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">S..</div>
                                        <div style="text-align:center">             
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:2px; visibility:inherit; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="font-size:100%; letter-spacing:0px; visibility:inherit; color:#595959" class="fa fa-refresh"></span>
                                        </div>     
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
                            <th style="background-color:#f5f5f5; vertical-align:middle; border-top:hidden;">
                                <select class="form-control input-small">
                                    <option value="0">Sort by</option>
                                    <option value="1">Customer</option>
                                    <option value="2">Job</option>
                                    <option value="3">Group</option>
                                    <option value="4">Employees</option>
                                </select>
                            </th>
                            <th colspan="29" style="text-align:center; border-top:hidden; background-color:#f5f5f5; font-size:175%"><a href="#"><</a> &nbsp Uge 1 - Januar 2016 &nbsp <a href="#">></a> 
                                <div class="pull-right">
                                    <a href="ressplan_2017.asp?func=day" class="btn btn-default btn-sm"><b>Day</b></a>
                                    <a href="ressplan_2017.asp?func=week" class="btn btn-secondary btn-sm"><b>Week</b></a>
                                    <a href="ressplan_2017.asp?" class="btn btn-default btn-sm"><b>Month</b></a>
                                </div>
                                
                            </th>
                        </tr>

                        <tr>
                            <th style="background-color:#f5f5f5">
                                <select class="form-control input-small">
                                      <option value="0">View type</option>
                                      <option value="1">Project time</option>
                                      <option value="2">Absence</option>
                                    <option value="3">Important</option>
                                </select>
                            </th>
                            <th style="text-align:center; background-color:#f5f5f5; width:13%"><a href="ressplan.asp?func=day" style="text-decoration:none; color:inherit">Monday</a></th>
                            <th style="text-align:center; background-color:#f5f5f5; width:13%">Torsday</th>
                            <th style="text-align:center; background-color:#f5f5f5; width:13%">Wednesday</th>
                            <th style="text-align:center; background-color:#f5f5f5; width:13%">Thursday</th>
                            <th style="text-align:center; background-color:#f5f5f5; width:13%">Friday</th>
                            <th style="text-align:center; background-color:lightgray; width:13%">Saturday</th>
                            <th style="text-align:center; background-color:lightgray; width:13%">Sunday</th>
                        </tr>

                        <tr>
                            <th style="background-color:#f5f5f5">
                                <select class="form-control input-small">
                                      <!--<option value="0">All</option> all employees should be select at first by standard  --> 
                                      <option value="1">Oliver Storm, Jesper Aaberg</option> <!-- imagine you just picked more than one employee   -->
                                      <option value="1">Oliver Storm</option>
                                      <option value="2">Jesper, Aaberg</option>
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
                        
                  
                        
                        <tr>
                            <td style="width:10%; vertical-align:middle;"><!-- =left(oRec("mnavn"), 20) --><div style="font-size:110%">Oliver Storm</div></td>
                            
                            <td style="letter-spacing:-4px;">
                                <!-------- this is how a 50% width box should look like ------> 
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:50%; height:40px; background-color:#9fc6e7; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Outzour... </div>
                                        <div style="text-align:right">             
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:5px; display:normal; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:3px; display:normal; color:#595959" class="fa fa-refresh"></span>
                                        </div>
                                    </div>
                                </a>

                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:50%; height:40px; background-color:#9fc6e7; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Kia</div>     
                                        <div style="text-align:right">             
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:5px; display:normal; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:3px; display:normal; color:#595959" class="fa fa-refresh"></span>
                                        </div>
                                    </div>
                                </a>

                                <!-------- this is how a 100% width box should look like ------> 
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#9fc6e7; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Support</div>
                                        <div style="text-align:right">             
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:5px; display:normal; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:3px; display:normal; color:#595959" class="fa fa-refresh"></span>
                                        </div>      
                                    </div>
                                </a>
                                

                            </td>
                                
                            
                            
                            <td style="letter-spacing:-4px;">
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:25%; height:40px; background-color:#9fc6e7; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Ou..</div>     
                                        <div style="text-align:right">             
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:5px; visibility:inherit; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:3px; visibility:inherit; color:#595959" class="fa fa-refresh"></span>
                                        </div>
                                    </div>
                                </a>

                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:50%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Ford</div>     
                                        <div style="text-align:right">             
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:5px; visibility:inherit; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:3px; visibility:inherit; color:#595959" class="fa fa-refresh"></span>
                                        </div>
                                    </div>
                                </a>

                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:25%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">BM..</div>     
                                        <div style="text-align:right">             
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:5px; visibility:hidden; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:3px; visibility:hidden; color:#595959" class="fa fa-refresh"></span>
                                        </div>
                                    </div>
                                </a>
                            </td>


                            <td style="letter-spacing:-4px;">
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:25%; height:40px; background-color:#9fc6e7; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Ou..</div>     
                                        <div style="text-align:right">             
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:5px; visibility:inherit; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:3px; visibility:inherit; color:#595959" class="fa fa-refresh"></span>
                                        </div>
                                    </div>
                                </a>

                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:25%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Ford</div>     
                                        <div style="text-align:right">             
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:5px; visibility:inherit; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:3px; visibility:inherit; color:#595959" class="fa fa-refresh"></span>
                                        </div>
                                    </div>
                                </a>

                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:25%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">BM..</div>     
                                        <div style="text-align:right">             
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:5px; visibility:hidden; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:3px; visibility:hidden; color:#595959" class="fa fa-refresh"></span>
                                        </div>
                                    </div>
                                </a>

                                <!-------- this is how a 25% width box should look like (probably) ------>
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:25%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">BM..</div>     
                                        <div style="text-align:right">             
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:5px; visibility:inherit; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:3px; visibility:inherit; color:#595959" class="fa fa-refresh"></span>
                                        </div>
                                    </div>
                                </a>
                            </td>

                            <td style="letter-spacing:-4px;">
                                
                               
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#b3dc6c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Holiday</div>
                                        <div style="text-align:right">             
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:5px; display:none; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:3px; display:none; color:#595959" class="fa fa-refresh"></span>
                                        </div>      
                                    </div>
                                </a>
                                

                            </td>

                            <td style="letter-spacing:-4px;">
                                
                               
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#b3dc6c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Holiday</div>
                                        <div style="text-align:right">             
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:5px; display:none; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:3px; display:none; color:#595959" class="fa fa-refresh"></span>
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
                                    <div style="width:50%; height:40px; background-color:#9fc6e7; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Outzour... </div>
                                        <div style="text-align:right">             
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:5px; display:normal; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:3px; display:normal; color:#595959" class="fa fa-refresh"></span>
                                        </div>
                                    </div>
                                </a>

                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:50%; height:40px; background-color:#9fc6e7; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Kia</div>     
                                        <div style="text-align:right">             
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:5px; display:normal; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:3px; display:normal; color:#595959" class="fa fa-refresh"></span>
                                        </div>
                                    </div>
                                </a>

                                
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#9fc6e7; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Support</div>
                                        <div style="text-align:right">             
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:5px; display:normal; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:3px; display:normal; color:#595959" class="fa fa-refresh"></span>
                                        </div>      
                                    </div>
                                </a>
                                

                            </td>
                                
                            
                            
                            <td style="letter-spacing:-4px;">
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:25%; height:40px; background-color:#9fc6e7; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Ou..</div>     
                                        <div style="text-align:right">             
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:5px; visibility:inherit; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:3px; visibility:inherit; color:#595959" class="fa fa-refresh"></span>
                                        </div>
                                    </div>
                                </a>

                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:50%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Ford</div>     
                                        <div style="text-align:right">             
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:5px; visibility:inherit; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:3px; visibility:inherit; color:#595959" class="fa fa-refresh"></span>
                                        </div>
                                    </div>
                                </a>

                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:25%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">BM..</div>     
                                        <div style="text-align:right">             
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:5px; visibility:hidden; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:3px; visibility:hidden; color:#595959" class="fa fa-refresh"></span>
                                        </div>
                                    </div>
                                </a>
                            </td>


                            <td style="letter-spacing:-4px;">
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:25%; height:40px; background-color:#9fc6e7; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Ou..</div>     
                                        <div style="text-align:right">             
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:5px; visibility:inherit; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:3px; visibility:inherit; color:#595959" class="fa fa-refresh"></span>
                                        </div>
                                    </div>
                                </a>

                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:25%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Ford</div>     
                                        <div style="text-align:right">             
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:5px; visibility:inherit; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:3px; visibility:inherit; color:#595959" class="fa fa-refresh"></span>
                                        </div>
                                    </div>
                                </a>

                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:25%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">BM..</div>     
                                        <div style="text-align:right">             
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:5px; visibility:hidden; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:3px; visibility:hidden; color:#595959" class="fa fa-refresh"></span>
                                        </div>
                                    </div>
                                </a>

                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:25%; height:40px; background-color:#fbe983; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">BM..</div>     
                                        <div style="text-align:right">             
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:5px; visibility:inherit; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:3px; visibility:inherit; color:#595959" class="fa fa-refresh"></span>
                                        </div>
                                    </div>
                                </a>
                            </td>

                            <td style="letter-spacing:-4px;">
                                
                               
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#fa573c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Illness</div>
                                        <div style="text-align:right">             
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:5px; display:none; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:3px; display:none; color:#595959" class="fa fa-refresh"></span>
                                        </div>      
                                    </div>
                                </a>
                                

                            </td>

                            <td style="letter-spacing:-4px;">
                                
                               
                                <a data-toggle="modal" href="#basicModal" style="color:black;">
                                    <div style="width:100%; height:40px; background-color:#fa573c; display:inline-block; border:1px solid; border-color:#fff">
                                        <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">Illness</div>
                                        <div style="text-align:right">             
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:5px; display:none; color:#595959" class="fa fa-exclamation-circle"></span>
                                            <span style="font-size:100%; letter-spacing:0px; margin-right:3px; display:none; color:#595959" class="fa fa-refresh"></span>
                                        </div>      
                                    </div>
                                </a>
                                

                            </td>

                            <td></td>
                            <td></td>
                            


                            
                        </tr>

                            

                            

                            

                            
                        
                        
                        


                    </table>
                    
                  


                  <%   
                      case else 
                  %>
                    <!------------------------------------------------ Måned ------------------------------------------------------------------------------------>


                    <table class="table dataTable table-striped table-bordered ui-datatable">
                       <!-- <thead> -->
                            <tr> 
                                <th style="background-color:#f5f5f5; border:hidden"></th>                                                            
                                <th colspan="72" style="text-align:center; border-top:hidden; background-color:#f5f5f5; font-size:175%"><a href="#"><</a> &nbsp January 2016 &nbsp <a href="#">></a> 
                                    <div class="pull-right">
                                        <a href="ressplan_2017.asp?func=day" class="btn btn-default btn-sm"><b>day</b></a>
                                        <a href="ressplan_2017.asp?func=week" class="btn btn-default btn-sm"><b>Week</b></a>
                                        <a href="ressplan_2017.asp?" class="btn btn-secondary btn-sm"><b>Month</b></a>
                                    </div>
                                </th>
                            </tr>

                            


                            
                            <tr>
                                <th style="background-color:#f5f5f5; vertical-align:middle;">
                                <select class="form-control input-small">
                                    <option value="0">Sort by</option>
                                    <option value="1">Customer</option>
                                    <option value="2">Job</option>
                                    <option value="3">Group</option>
                                    <option value="4">Employees</option>
                                </select>
                            </th> 
                               
                                <th colspan="7" style="text-align:center; background-color:#f5f5f5"><a href="ressplan.asp?func=week" style="color:dimgray">Week 1</a></th>
                                <th colspan="7" style="text-align:center; background-color:#f5f5f5">Week 2</th>
                                <th colspan="7" style="text-align:center; background-color:#f5f5f5">Week 3</th>
                                <th colspan="7" style="text-align:center; background-color:#f5f5f5">Week 4</th>
                                <th colspan="7" style="text-align:center; background-color:#f5f5f5">Week 5</th>
                            </tr>                     
                            
                            <tr>
                                <th style="background-color:#f5f5f5">
                                    <select class="form-control input-small">
                                          <option value="0">View type</option>
                                          <option value="1">Project time</option>
                                          <option value="2">Absence</option>
                                        <option value="3">Important</option>
                                    </select>
                                </th>
                                <th style="width:3%; background-color:#f5f5f5; text-align:center"><a href="ressplan.asp?func=day" style="color:dimgray">Mon</a></th>
                                <th style="width:3%; background-color:#f5f5f5;text-align:center">Tue</th>
                                <th style="width:3%; background-color:#f5f5f5; text-align:center">Wed</th>
                                <th style="width:3%; background-color:#f5f5f5; text-align:center">Thu</th>
                                <th style="width:3%; background-color:#f5f5f5; text-align:center">Fri</th>
                                <th style="width:3%; background-color:lightgray; text-align:center">Sat</th>
                                <th style="width:3%; background-color:lightgray; text-align:center">Sun</th>

                                <th style="width:3%; background-color:#f5f5f5; text-align:center"><a href="ressplan.asp?func=day" style="color:dimgray">Mon</a></th>
                                <th style="width:3%; background-color:#f5f5f5;text-align:center">Tue</th>
                                <th style="width:3%; background-color:#f5f5f5; text-align:center">Wed</th>
                                <th style="width:3%; background-color:#f5f5f5; text-align:center">Thu</th>
                                <th style="width:3%; background-color:#f5f5f5; text-align:center">Fri</th>
                                <th style="width:3%; background-color:lightgray; text-align:center">Sat</th>
                                <th style="width:3%; background-color:lightgray; text-align:center">Sun</th>

                                <th style="width:3%; background-color:#f5f5f5; text-align:center"><a href="ressplan.asp?func=day" style="color:dimgray">Mon</a></th>
                                <th style="width:3%; background-color:#f5f5f5;text-align:center">Tue</th>
                                <th style="width:3%; background-color:#f5f5f5; text-align:center">Wed</th>
                                <th style="width:3%; background-color:#f5f5f5; text-align:center">Thu</th>
                                <th style="width:3%; background-color:#f5f5f5; text-align:center">Fri</th>
                                <th style="width:3%; background-color:lightgray; text-align:center">Sat</th>
                                <th style="width:3%; background-color:lightgray; text-align:center">Sun</th>

                                <th style="width:3%; background-color:#f5f5f5; text-align:center"><a href="ressplan.asp?func=day" style="color:dimgray">Mon</a></th>
                                <th style="width:3%; background-color:#f5f5f5;text-align:center">Tue</th>
                                <th style="width:3%; background-color:#f5f5f5; text-align:center">Wed</th>
                                <th style="width:3%; background-color:#f5f5f5; text-align:center">Thu</th>
                                <th style="width:3%; background-color:#f5f5f5; text-align:center">Fri</th>
                                <th style="width:3%; background-color:lightgray; text-align:center">Sat</th>
                                <th style="width:3%; background-color:lightgray; text-align:center">Sun</th>

                                <th style="width:3%; background-color:#f5f5f5; text-align:center"><a href="ressplan.asp?func=day" style="color:dimgray">Mon</a></th>
                                <th style="width:3%; background-color:#f5f5f5;text-align:center">Tue</th>
                                <th style="width:3%; background-color:#f5f5f5; text-align:center">Wed</th>                                
                            </tr>
                            <tr>
                                <th style="background-color:#f5f5f5">
                                    <select class="form-control input-small">
                                          <!-- <option value="0">All</option> -->
                                          <option value="1">Oliver Storm</option>
                                          <option value="2">Jesper Aaberg</option>
                                    </select>
                                </th>
                                <th style="text-align:center; background-color:#f5f5f5;">1</th>
                                <th style="text-align:center; background-color:#f5f5f5;">2</th>
                                <th style="text-align:center; background-color:#f5f5f5;">3</th>
                                <th style="text-align:center; background-color:#f5f5f5;">4</th>
                                <th style="text-align:center; background-color:#f5f5f5;">5</th>
                                <th style="text-align:center; background-color:lightgray;">6</th>
                                <th style="text-align:center; background-color:lightgray;">7</th>

                                <th style="text-align:center; background-color:#f5f5f5;">8</th>
                                <th style="text-align:center; background-color:#f5f5f5;">9</th>
                                <th style="text-align:center; background-color:#f5f5f5;">10</th>
                                <th style="text-align:center; background-color:#f5f5f5;">11</th>
                                <th style="text-align:center; background-color:#f5f5f5;">12</th>
                                <th style="text-align:center; background-color:lightgray;">13</th>
                                <th style="text-align:center; background-color:lightgray;">14</th>

                                <th style="text-align:center; background-color:#f5f5f5;">15</th>
                                <th style="text-align:center; background-color:#f5f5f5;">16</th>
                                <th style="text-align:center; background-color:#f5f5f5;">17</th>
                                <th style="text-align:center; background-color:#f5f5f5;">18</th>
                                <th style="text-align:center; background-color:#f5f5f5;">19</th>
                                <th style="text-align:center; background-color:lightgray;">20</th>
                                <th style="text-align:center; background-color:lightgray;">21</th>

                                <th style="text-align:center; background-color:#f5f5f5;">22</th>
                                <th style="text-align:center; background-color:#f5f5f5;">23</th>
                                <th style="text-align:center; background-color:#f5f5f5;">24</th>
                                <th style="text-align:center; background-color:#f5f5f5;">25</th>
                                <th style="text-align:center; background-color:#f5f5f5;">26</th>
                                <th style="text-align:center; background-color:lightgray;">27</th>
                                <th style="text-align:center; background-color:lightgray;">28</th>

                                <th style="text-align:center; background-color:#f5f5f5;">29</th>
                                <th style="text-align:center; background-color:#f5f5f5;">30</th>
                                <th style="text-align:center; background-color:#f5f5f5;">31</th>

                            </tr> 
                          

                            <tr>
                                <td style="width:10%; font-size:110%; vertical-align:middle"><!--=left(oRec("mnavn"), 20)--><div style="font-size:110%">Oliver Storm</div></td>
                                

                                <td style="letter-spacing:-4px;">
                                    
                                    <!------------- this is how a box of 100% witdh should look like ------------->
                                    <a data-toggle="modal" href="#basicModal" style="color:black;">
                                        <div style="width:100%; height:40px; background-color:#9fc6e7; display:inline-block; border:1px solid; border-color:#fff">
                                            <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">O..</div>
                                                
                                        </div>
                                    </a>
                                    <a data-toggle="modal" href="#basicModal" style="color:black;">
                                        <div style="width:100%; height:40px; background-color:#9fc6e7; display:inline-block; border:1px solid; border-color:#fff">
                                            <div style="letter-spacing:0px; font-size:110%; margin-left:2px;">P..</div>     
                                        </div>
                                    </a>

                                </td>

                                <td style="letter-spacing:-4px;">
                                    <!------------- this is how a box of 50% witdh should look like ------------->
                                    <a data-toggle="modal" href="#basicModal" style="color:black;">
                                        <div style="width:50%; height:40px; background-color:#9fc6e7; display:inline-block; border:1px solid; border-color:#fff">
                                            <div style="letter-spacing:0px; font-size:110%;"></div>     
                                        </div>
                                    </a>
                                    <a data-toggle="modal" href="#basicModal" style="color:black;">
                                        <div style="width:50%; height:40px; background-color:#9fc6e7; display:inline-block; border:1px solid; border-color:#fff">
                                            <div style="letter-spacing:0px; font-size:110%;"></div>     
                                        </div>
                                    </a>
                                    <a data-toggle="modal" href="#basicModal" style="color:black;">
                                        <div style="width:50%; height:40px; background-color:#9fc6e7; display:inline-block; border:1px solid; border-color:#fff">
                                            <div style="letter-spacing:0px; font-size:110%;"></div>     
                                        </div>
                                    </a>
                                    <a data-toggle="modal" href="#basicModal" style="color:black;">
                                        <div style="width:50%; height:40px; background-color:#9fc6e7; display:inline-block; border:1px solid; border-color:#fff">
                                            <div style="letter-spacing:0px; font-size:110%;"></div>     
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