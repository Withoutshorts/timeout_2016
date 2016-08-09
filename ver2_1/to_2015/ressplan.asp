


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
    .hover_img a span { position:absolute; display:none; z-index:90; border: 2px solid white; }
    .hover_img a:hover span { display:block; width:200px;}

    table th 
    {
        font-size:125%
    }

</style>

<div class="wrapper">
    <div class="content">   
        
        <div style="width:1600px;" class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u>Ressource Planlægning</u></h3>
                <div class="portlet-body">
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
                    <table class="table dataTable table-striped table-bordered ui-datatable">
                        <tr>
                            <th style="background-color:#f5f5f5; vertical-align:middle">
                                <input type="text" class="form-control input-small" placeholder="Søg kunde/job" value="">
                            </th>
                            <th colspan="70" style="text-align:center; border-top:hidden; background-color:#f5f5f5; font-size:175%"><a href="#"><</a> &nbsp Mandag 01-01-16 &nbsp <a href="#">></a> 
                                    <div class="pull-right">
                                        <a href="ressplan.asp?func=day" class="btn btn-secondary btn-sm"><b>Dag</b></a>
                                        <a href="ressplan.asp?func=week" class="btn btn-secondary btn-sm"><b>Uge</b></a>
                                        <a href="ressplan.asp?" class="btn btn-secondary btn-sm"><b>Måned</b></a>
                                    </div>
                            </th>
                        </tr>

                        <tr>
                            <th style="background-color:#f5f5f5">
                                <select class="form-control input-small">
                                      <option value="0">Vis type</option>
                                      <option value="1">Projekt tid</option>
                                      <option value="2">Fravær</option>
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
                            <th colspan="4" style="text-align:center; border-right-color:darkgray; background-color:#f5f5f5;">06:00</th>
                            <th colspan="4" style="text-align:center; border-right-color:darkgray; background-color:#f5f5f5;">07:00</th>
                            <th colspan="4" style="text-align:center; border-right-color:darkgray; background-color:#f5f5f5;">08:00</th>
                            <th colspan="4" style="text-align:center; border-right-color:darkgray; background-color:#f5f5f5;">09:00</th>
                            <th colspan="4" style="text-align:center; border-right-color:darkgray; background-color:#f5f5f5;">10:00</th>
                            <th colspan="4" style="text-align:center; border-right-color:darkgray; background-color:#f5f5f5;">11:00</th>
                            <th colspan="4" style="text-align:center; border-right-color:darkgray; background-color:#f5f5f5;">12:00</th>
                            <th colspan="4" style="text-align:center; border-right-color:darkgray; background-color:#f5f5f5;">13:00</th>
                            <th colspan="4" style="text-align:center; border-right-color:darkgray; background-color:#f5f5f5;">14:00</th>
                            <th colspan="4" style="text-align:center; border-right-color:darkgray; background-color:#f5f5f5">15:00</th>
                            <th colspan="4" style="text-align:center; border-right-color:darkgray; background-color:#f5f5f5;">16:00</th>
                            <th colspan="4" style="text-align:center; border-right-color:darkgray; background-color:#f5f5f5;">17:00</th>
                            <th colspan="4" style="text-align:center; border-right-color:darkgray; background-color:#f5f5f5;">18:00</th>
                            <th colspan="4" style="text-align:center; border-right-color:darkgray; background-color:#f5f5f5;">19:00</th>
                            <th colspan="4" style="text-align:center; border-right-color:darkgray; background-color:#f5f5f5;">20:00</th>
                            <th colspan="4" style="text-align:center; border-right-color:darkgray; background-color:#f5f5f5;">21:00</th>
                            <th colspan="4" style="text-align:center; border-right-color:darkgray; background-color:#f5f5f5;">22:00</th>
                        </tr>
                        <%
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
                



                            <%strSQL = "SELECT mnavn, mid, email, lastlogin, init, mnr, mansat, brugergruppe, medarbejdertype, mt.type AS mtypenavn, b.navn AS brugergruppenavn FROM medarbejdere AS m "_
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
                        %>
                        <tr>
                            <td style="width:10%;"><%=left(oRec("mnavn"), 20) %></td>
                            
                            <td></td>
                            <td></td>
                            <td style="background-color:#0086b3; border-right:hidden" class="hover_img">
                                <a href="#" class="fa fa-edit"><span style="background-color:cornflowerblue;"></span></a>
                            </td>
                            <td style="background-color:#e6b800; border-right-color:darkgrey"></td>
                            <td style="background-color:#e6b800; border-right:hidden"></td>
                            <td></td>
                            <td style="background-color:#00b33c"></td>
                            <td style="background-color:#e60000; border-right-color:darkgrey"></td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td style="border-right-color:darkgrey;"></td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td style="border-right-color:darkgrey;"></td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td style="border-right-color:darkgrey;"></td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td style="border-right-color:darkgrey;"></td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td style="border-right-color:darkgrey;"></td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td style="border-right-color:darkgrey;"></td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td style="border-right-color:darkgrey;"></td>
                            <td style="background-color:greenyellow"></td>
                            <td></td>
                            <td></td>
                            <td style="border-right-color:darkgrey;"></td>
                            <td style="background-color:#e60000"></td>
                            <td style="background-color:#e60000"></td>
                            <td></td>
                            <td style="border-right-color:darkgrey;"></td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td style="border-right-color:darkgrey;"></td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td style="border-right-color:darkgrey;"></td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td style="border-right-color:darkgrey;"></td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td style="border-right-color:darkgrey;"></td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td style="border-right-color:darkgrey;"></td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td></td>

                        </tr>
                        <%
                            oRec.movenext
                            wend
                      
                            oRec.close 
                        %>
                    </table>

                     <br /><br /><br /><br />

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
                    </div>


                             
                   <%
                       case "week"
                   %>

                    <!----------------------------------------------- Uge -------------------------------------------------------------------------------------->

                    <table class="table dataTable table-striped table-bordered ui-datatable">
                        <tr>
                            <th style="background-color:#f5f5f5; vertical-align:middle"">
                                <input type="text" class="form-control input-small" placeholder="Søg kunde/job" value="">
                            </th>
                            <th colspan="29" style="text-align:center; border-top:hidden; background-color:#f5f5f5; font-size:175%"><a href="#"><</a> &nbsp Uge 1 - Januar 2016 &nbsp <a href="#">></a> 
                                <div class="pull-right">
                                    <a href="ressplan.asp?func=day" class="btn btn-secondary btn-sm"><b>Dag</b></a>
                                    <a href="ressplan.asp?func=week" class="btn btn-secondary btn-sm"><b>Uge</b></a>
                                    <a href="ressplan.asp?" class="btn btn-secondary btn-sm"><b>Måned</b></a>
                                </div>
                                
                            </th>
                        </tr>

                        <tr>
                            <th style="border-right-color:black; background-color:#f5f5f5">
                                <select class="form-control input-small">
                                      <option value="0">Vis type</option>
                                      <option value="1">Projekt tid</option>
                                      <option value="2">Fravær</option>
                                </select>
                            </th>
                            <th colspan="4" style="text-align:center; border-right-color:black; border-top-color:black; background-color:#f5f5f5;"><a href="ressplan.asp?func=day" style="color:dimgray">Mandag</a></th>
                            <th colspan="4" style="text-align:center; border-right-color:darkgray; background-color:#f5f5f5;">Tirsdag</th>
                            <th colspan="4" style="text-align:center; border-right-color:darkgray; background-color:#f5f5f5;">Onsdag</th>
                            <th colspan="4" style="text-align:center; border-right-color:darkgray; background-color:#f5f5f5;">Torsdag</th>
                            <th colspan="4" style="text-align:center; border-right-color:darkgray; background-color:#f5f5f5;">Fredag</th>
                            <th colspan="4" style="text-align:center; border-right-color:darkgray; background-color:lightgray;">Lørdag</th>
                            <th colspan="4" style="text-align:center; border-right-color:darkgray; background-color:lightgray;">Søndag</th>
                        </tr>

                        <tr>
                            <th style="border-right-color:black; background-color:#f5f5f5">
                                <select class="form-control input-small">
                                      <option value="0">Vælg gruppe</option>
                                      <option value="1">Design 2014()</option>
                                      <option value="2">Timetag</option>
                                </select>
                            </th>
                            <th colspan="4" style="text-align:center; border-right-color:black; background-color:#f5f5f5">1</th>
                            <th colspan="4" style="text-align:center; border-right-color:darkgray; background-color:#f5f5f5">2</th>
                            <th colspan="4" style="text-align:center; border-right-color:darkgray; background-color:#f5f5f5">3</th>
                            <th colspan="4" style="text-align:center; border-right-color:darkgray; background-color:#f5f5f5">4</th>
                            <th colspan="4" style="text-align:center; border-right-color:darkgray; background-color:#f5f5f5">5</th>
                            <th colspan="4" style="text-align:center; border-right-color:darkgray; background-color:lightgray">6</th>
                            <th colspan="4" style="text-align:center; border-right-color:darkgray; background-color:lightgray;">7</th>
                        </tr>

                       

                        <%
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
                



                            <%strSQL = "SELECT mnavn, mid, email, lastlogin, init, mnr, mansat, brugergruppe, medarbejdertype, mt.type AS mtypenavn, b.navn AS brugergruppenavn FROM medarbejdere AS m "_
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
                        %>

                        <tr>
                            <td style="width:10%; border-right-color:black;"><%=left(oRec("mnavn"), 20) %></td>
                            <td style="background-color:#e60000; border-right:hidden;" class="hover_img">
                                <a href="projektgrupper.asp" class="fa fa-edit"><span style="background-color:dimgray;"><h5>Voice(59) <br />
                                Rengøring (Toilet)<br />
                                10:00-11:00                                                                                    
                                </h5></span></a>
                            </td>
                            <td style="background-color:#e60000"></td>
                            <td style="background-color:#0086b3; border-right:hidden"></td>
                            <td style="border-right-color:black; background-color:#e60000;"></td>

                            <td></td>
                            <td></td>
                            <td></td>
                            <td style="border-right-color:darkgray"></td>

                            <td colspan="2"></td>                      
                            <td style="background-color:#e6b800"></td>
                            <td></td>
                            <td style="background-color:#00b33c;"></td> 

                            <td style="background-color:#00b33c; border-right:hidden"></td>
                            <td style="background-color:##e60000"></td>
                            <td style="border-right-color:darkgray"></td>

                            <td></td>
                            <td></td>
                            <td></td>
                            <td style="border-right-color:darkgray"></td>

                            <td style="background-color:lightgray; border-right:hidden"></td>
                            <td style="background-color:lightgray; border-right:hidden"></td>
                            <td style="background-color:lightgray; border-right:hidden"></td>
                            <td style="background-color:lightgray; border-right-color:darkgray"></td>

                            <td style="background-color:lightgray; border-right:hidden"></td>
                            <td style="background-color:lightgray; border-right:hidden"></td>
                            <td style="background-color:lightgray; border-right:hidden"></td>
                            <td style="background-color:lightgray; border-right-color:darkgray"></td>

                            
                        </tr>
                        <%
                            oRec.movenext
                            wend
                      
                            oRec.close 
                        %>

                    </table>

                    <br /><br /><br /><br />

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
                    </div>






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
                                        <a href="ressplan.asp?func=day" class="btn btn-secondary btn-sm"><b>Dag</b></a>
                                        <a href="ressplan.asp?func=week" class="btn btn-secondary btn-sm"><b>Uge</b></a>
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
                                <th style="background-color:#f5f5f5; vertical-align:middle"">
                                    <input type="text" class="form-control input-small" placeholder="Søg kunde/job" value="">
                                </th>
                                
                                <th colspan="14" style="text-align:center; background-color:#f5f5f5"><a href="ressplan.asp?func=week" style="color:dimgray">Uge 1</a></th>
                                <th colspan="14" style="text-align:center; background-color:#f5f5f5">Uge 2</th>
                                <th colspan="14" style="text-align:center; background-color:#f5f5f5">Uge 3</th>
                                <th colspan="14" style="text-align:center; background-color:#f5f5f5">Uge 4</th>
                                <th colspan="14" style="text-align:center; background-color:#f5f5f5">Uge 5</th>
                            </tr>                     
                            
                            <tr>
                                <th style="width:3%; border-right-color:black; background-color:#f5f5f5">
                                    <select class="form-control input-small">
                                      <option value="0">Vis type</option>
                                      <option value="1">Projekt tid</option>
                                      <option value="2">Fravær</option>
                                    </select>
                                </th>
                                <th style="width:3%; border-right-color:black; border-top-color:black; background-color:#f5f5f5" colspan="2"><a href="ressplan.asp?func=day" style="color:dimgray">Man</a></th>
                                <th style="width:3%; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">Tir</th>
                                <th style="width:3%; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">Ons</th>
                                <th style="width:3%; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">Tor</th>
                                <th style="width:3%; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">Fre</th>
                                <th style="width:3%; background-color:lightgray; border-right-color:darkgray;" colspan="2">Lør</th>
                                <th style="width:3%; background-color:lightgray; border-right-color:darkgray;" colspan="2">Søn</th>

                                <th style="width:3%; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">Man</th>
                                <th style="width:3%; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">Tir</th>
                                <th style="width:3%; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">Ons</th>
                                <th style="width:3%; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">Tor</th>
                                <th style="width:3%; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">Fre</th>
                                <th style="width:3%; background-color:lightgray; border-right-color:darkgray;" colspan="2">Lør</th>
                                <th style="width:3%; background-color:lightgray; border-right-color:darkgray;" colspan="2">Søn</th>

                                <th style="width:3%; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">Man</th>
                                <th style="width:3%; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">Tir</th>
                                <th style="width:3%; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">Ons</th>
                                <th style="width:3%; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">Tor</th>
                                <th style="width:3%; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">Fre</th>
                                <th style="width:3%; background-color:lightgray; border-right-color:darkgray;" colspan="2">Lør</th>
                                <th style="width:3%; background-color:lightgray; border-right-color:darkgray;" colspan="2">Søn</th>

                                <th style="width:3%; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">Man</th>
                                <th style="width:3%; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">Tir</th>
                                <th style="width:3%; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">Ons</th>
                                <th style="width:3%; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">Tor</th>
                                <th style="width:3%; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">Fre</th>
                                <th style="width:3%; background-color:lightgray; border-right-color:darkgray;" colspan="2">Lør</th>
                                <th style="width:3%; background-color:lightgray; border-right-color:darkgray;" colspan="2">Søn</th>

                                <th style="width:3%; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">Man</th>
                                <th style="width:3%; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">Tir</th>
                                <th style="width:3%; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">Ons</th>

                                
                            </tr>
                            <tr>
                                <th style="border-right-color:black; background-color:#f5f5f5">
                                    <select class="form-control input-small">
                                      <option value="0">Vælg gruppe</option>
                                      <option value="1">Design 2014()</option>
                                      <option value="2">Timetag</option>
                                    </select>
                                </th>
                                <th style="width:3%; text-align:center; border-right-color:black; background-color:#f5f5f5" colspan="2">1</th>
                                <th style="width:3%; text-align:center; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">2</th>
                                <th style="width:3%; text-align:center; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">3</th>
                                <th style="width:3%; text-align:center; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">4</th>
                                <th style="width:3%; text-align:center; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">5</th>
                                <th style="width:3%; text-align:center; background-color:lightgray; border-right-color:darkgray;" colspan="2">6</th>
                                <th style="width:3%; text-align:center; background-color:lightgray; border-right-color:darkgray;" colspan="2">7</th>

                                <th style="width:3%; text-align:center; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">8</th>
                                <th style="width:3%; text-align:center; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">9</th>
                                <th style="width:3%; text-align:center; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">10</th>
                                <th style="width:3%; text-align:center; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">11</th>
                                <th style="width:3%; text-align:center; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">12</th>
                                <th style="width:3%; text-align:center; background-color:lightgray; border-right-color:darkgray;" colspan="2">13</th>
                                <th style="width:3%; text-align:center; background-color:lightgray; border-right-color:darkgray;" colspan="2">14</th>

                                <th style="width:3%; text-align:center; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">15</th>
                                <th style="width:3%; text-align:center; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">16</th>
                                <th style="width:3%; text-align:center; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">17</th>
                                <th style="width:3%; text-align:center; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">18</th>
                                <th style="width:3%; text-align:center; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">19</th>
                                <th style="width:3%; text-align:center; background-color:lightgray; border-right-color:darkgray;" colspan="2">20</th>
                                <th style="width:3%; text-align:center; background-color:lightgray; border-right-color:darkgray;" colspan="2">21</th>

                                <th style="width:3%; text-align:center; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">22</th>
                                <th style="width:3%; text-align:center; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">23</th>
                                <th style="width:3%; text-align:center; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">24</th>
                                <th style="width:3%; text-align:center; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">25</th>
                                <th style="width:3%; text-align:center; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">26</th>
                                <th style="width:3%; text-align:center; background-color:lightgray; border-right-color:darkgray;" colspan="2">27</th>
                                <th style="width:3%; text-align:center; background-color:lightgray; border-right-color:darkgray;" colspan="2">28</th>

                                <th style="width:3%; text-align:center; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">29</th>
                                <th style="width:3%; text-align:center; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">30</th>
                                <th style="width:3%; text-align:center; background-color:#f5f5f5; border-right-color:darkgray;" colspan="2">31</th>

                            </tr> 
                            <%
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
                



                            <%strSQL = "SELECT mnavn, mid, email, lastlogin, init, mnr, mansat, brugergruppe, medarbejdertype, mt.type AS mtypenavn, b.navn AS brugergruppenavn FROM medarbejdere AS m "_
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
                                %>

                            <tr>
                                <td style="width:10%; border-right-color:black;"><%=left(oRec("mnavn"), 20) %></td>
                                <td style="background-color:#0086b3; border-right:hidden"></td>
                                <td style="background-color:#e6b800; border-right-color:black;"></td>
                                <td style="background-color:#e6b800; border-right:hidden"></td>
                                <td style="background-color:#e6b800; border-right-color:darkgray;"></td>
                                <td style="background-color:#00b33c; border-right:hidden"></td>
                                <td style="background-color:#00b33c; border-right-color:gray;"></td>
                                <td style="background-color:#e6b800; border-right:hidden"></td>
                                <td style="background-color:#00b33c; border-right-color:darkgray;"></td>
                                <td style="background-color:#00b33c; border-right:hidden"></td>
                                <td style="background-color:#e60000; border-right-color:darkgray;"></td>
                                <td style="background-color:#e60000; border-right:hidden"></td>
                                <td style="background-color:#e60000; border-right-color:darkgray;"></td>
                                <td style="background-color:#e60000; border-right:hidden"></td>
                                <td style="background-color:#e60000; border-right-color:darkgray;"></td>

                                <td style="border-right:hidden;"></td>
                                <td style="border-right-color:darkgray;"></td>
                                <td style="border-right:hidden;"></td>
                                <td style="border-right-color:darkgray;"></td>
                                <td style="border-right:hidden;"></td>
                                <td style="border-right-color:darkgray;"></td>
                                <td style="border-right:hidden;"></td>
                                <td style="border-right-color:darkgray;"></td>
                                <td style="border-right:hidden;"></td>
                                <td style="border-right-color:darkgray;"></td>
                                <td style="border-right:hidden; background-color:lightgray"></td>
                                <td style="background-color:lightgray; border-right-color:darkgray;"></td>
                                <td style="border-right:hidden; background-color:lightgray"></td>
                                <td style="background-color:lightgray; border-right-color:darkgray;"></td>

                                <td style="border-right:hidden;"></td>
                                <td style="border-right-color:darkgray;"></td>
                                <td style="border-right:hidden;"></td>
                                <td style="border-right-color:darkgray;"></td>
                                <td style="border-right:hidden;"></td>
                                <td style="border-right-color:darkgray;"></td>
                                <td style="border-right:hidden;"></td>
                                <td style="border-right-color:darkgray;"></td>
                                <td style="border-right:hidden;"></td>
                                <td style="border-right-color:darkgray;"></td>
                                <td style="border-right:hidden; background-color:lightgray"></td>
                                <td style="background-color:lightgray; border-right-color:darkgray"></td>
                                <td style="border-right:hidden; background-color:lightgray"></td>
                                <td style="background-color:lightgray; border-right-color:darkgray"></td>

                                <td style="border-right:hidden;"></td>
                                <td style="border-right-color:darkgray;"></td>
                                <td style="border-right:hidden;"></td>
                                <td style="border-right-color:darkgray;"></td>
                                <td style="border-right:hidden;"></td>
                                <td style="border-right-color:darkgray;"></td>
                                <td style="border-right:hidden;"></td>
                                <td style="border-right-color:darkgray;"></td>
                                <td style="border-right:hidden;"></td>
                                <td style="border-right-color:darkgray;"></td>
                                <td style="border-right:hidden; background-color:lightgray"></td>
                                <td style="background-color:lightgray; border-right-color:darkgray"></td>
                                <td style="border-right:hidden; background-color:lightgray"></td>
                                <td style="background-color:lightgray; border-right-color:darkgray"></td>
                                
                                <td style="border-right:hidden;"></td>
                                <td style="border-right-color:darkgray;"></td>
                                <td style="border-right:hidden;"></td>
                                <td style="border-right-color:darkgray;"></td>
                                <td style="border-right:hidden;"></td>
                                <td style="border-right-color:darkgray;"></td>
                                
                               


                               

                            </tr>
                            <%
                                oRec.movenext
                                wend
                      
                                oRec.close 
                            %> 

                          <!--  <tr>
                                <td style="border-bottom:hidden; border-right:hidden;"></td>
                                <td colspan="14" style="border-bottom:hidden; border-right:hidden; border-top-color:black;"></td>
                                <td colspan="14" style="border-bottom:hidden; border-right:hidden;"></td>
                                <td colspan="14" style="border-bottom:hidden; border-right:hidden;"></td>
                                <td colspan="14" style="border-bottom:hidden; border-right:hidden;"></td>
                                <td colspan="14" style="border-bottom:hidden; border-right:hidden;"></td>
                            </tr> -->
                                                                 
                       <!-- </tbody> -->
                    </table>

                     <br /><br /><br /><br />

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
                    </div>

                </div>
            </div>
        </div>
        <%end select %>
    </div>
</div>





<!--#include file="../inc/regular/footer_inc.asp"-->