<%response.buffer = true%>

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/job_func.asp"-->
<!--#include file = "../timereg/CuteEditor_Files/include_CuteEditor.asp" -->
<!--#include file="../inc/regular/topmenu_inc.asp"--> 
<!--#include file="inc/convertDate.asp"-->
<!--#include file="inc/timbudgetsim_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->


<!--------------- Github 1 ----------------->

<%call menu_2014 %>

<% 
    if len(session("user")) = 0 then
	%>
	
	<%
	errortype = 5
	call showError(errortype)
       response.End
	end if


    func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if







    Select case func 
    case "opret", "red"

       
    strSQL = "SELECT id, jobnavn, jobnr, kkundenavn, jobknr FROM job, kunder WHERE id = " & id &" AND kunder.Kid = jobknr"

        'Response.Write strSQL
	    'Response.flush
	
	    oRec.open strSQL, oConn, 3
	
	    if not oRec.EOF then
	    strNavn = oRec("jobnavn")
	    strjobnr = oRec("jobnr")
	    strKnavn = oRec("kkundenavn")
	    strKnr = oRec("jobknr")

        end if


        oRec.close

        dbfunc = "dbred"
%>










<div class="wrapper">
    <div class="content">



    
 <!------------------------------- Sideindhold------------------------------------->
 

        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u>Projekt oprettelse</u></h3>
            </div>

            <div class="portlet-body">
                            	           
                <div class="panel-group accordion-panel" id="accordion-paneled">
                    <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapse3">
                                Kunde info
                            </a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapse3" class="panel-collapse collapse in">
                            <div class="panel-body">
                            
                                
                                  <div class="row">
                                  
                                      <div class="col-lg-2">Kunde:</div>
                                      <div class="col-lg-3">
                                          <select class="form-control input-small" name="FM_kunde" id="FM_kunde" onchange="submit();">
		                                    <option value="0">Alle <%=writethis%></option>
		                                    <%
				                                    strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE ketype <> 'e' ORDER BY Kkundenavn"
				                                    oRec.open strSQL, oConn, 3
				                                    while not oRec.EOF
				
				                                    if cint(fmkunde) = cint(oRec("Kid")) then
				                                    isSelected = "SELECTED"
				                                    else
				                                    isSelected = ""
				                                    end if
				                                    %>
				                                    <option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%>&nbsp;(<%=oRec("Kkundenr")%>)</option>
				                                    <%
				                                    oRec.movenext
				                                    wend
				                                    oRec.close
				                                    %>
		                                </select>
                                      </div>
                              
                                 
                            
                                    <div class="col-lg-1">Kontaktper.:</div>
                                    <div class="col-lg-2">
                                        <%
		                                strSQLkpers = "SELECT k.navn, k.id AS kid FROM kontaktpers k WHERE kundeid = "& strKundeId &" ORDER BY k.navn"
		                                'Response.Write strSQLkpers
		                                %>
		                                <select name="FM_opr_kpers" id="FM_opr_kpers" class="form-control input-small">
		                                    <option value="0">Ingen</option>
	
		                                    <%
		
		                                    'oRec.open strSQLkpers, oConn, 3 
		                                    'while not oRec.EOF 
		
		
		                                    'if cint(intkundekpers) = cint(oRec("kid")) then
		                                    'ts3 = "SELECTED"
		                                    'else
		                                    'ts3 = ""
		                                    'end if
		                                    %>
		                                        <!---- job_inc ---->
		                                    <%
		                                    'oRec.movenext
		                                    'wend
		                                    'oRec.close 
		                                    %>
		                                </select>
                                    </div>
                   
                            
                                    <div class="col-lg-1">Ders ref.</div>
                                    <div class="col-lg-2"><input class="form-control input-small" type="text" name="FM_exch" value="" placeholder="Ref"/></div>
                                </div>

                                <br /> <br />

                       

                                <div class="row">
                                                                      
                                        <div class="col-lg-2">Projektnavn:</div>
                                        <div class="col-lg-3"><input type="text" name="FM_navn" id="FM_navn" value="<%=strNavn%>" class="form-control input-small"></div>

                                        

                                        <div class="col-lg-1">Projekt nr:</div>
                                        <div class="col-lg-2"><input class="form-control input-small" type="text" name="FM_exch" value="" placeholder="eks. 1234"/></div>

                                        
                                        <div class="col-lg-1">Status:</div>
                                        <div class="col-lg-2">
                                            <select class="form-control input-small">
                                                <option value="1">Aktiv</option>
                                                <option value="2">Til fakturering</option>
                                                <option value="3">Lukket</option>
                                                <option value="4">Gennemsyn</option>
                                                <option value="5">Tilbud</option>
                                            </select>
                                        </div>
                                </div>

                                <div class="row">
                                    
                                    <div class="col-lg-2">Ansvarlig:</div>
                                    <div class="col-lg-3">
                                            <select name="FM_jobans_<%=ja %>" id="FM_jobans_<%=ja %>" >
						                        <option value="0">Ingen</option>
							                        <!----- Jobs linje 5361 --->
						                        </select>
                                    </div>

                                        

                                        <div class="col-lg-1">Start dato:</div>
                                        <div class="col-lg-1">
                                            <select class="form-control input-small">

                                            </select>
                                        </div>
                                        <div class="col-lg-1">
                                                <select class="form-control input-small">

                                                </select>
                                            </div>

                                        

                                        <div class="col-lg-1">Slut dato:</div>
                                        <div class="col-lg-2"><input class="form-control input-small" type="number" name="FM_exch" value="" placeholder="eks. 11-12-2016"/></div>

                                </div>

                                <div class="row">
                                    <div class="col-lg-2">Projektgruppe:</div>
                                        <div class="col-lg-3">
                                            <select class="form-control input-small">
                                                <option value="1">Rengøring</option>
                                                <option value="2">Udvikler</option>
                                                <option value="3">Kørsel</option>
                                                <option value="4">Opvask</option>
                                                <option value="5">Madlavning</option>
                                            </select>
                                        </div>
                                    
                                    <%if avansproobr = 1 then %>
                                    <div class="col-lg-6">Beskrivelse: <br />
                                    <textarea id="TextArea1" name="FM_betbet" cols="70" rows="3" class="form-control input-small"></textarea>   
		               
                                    </div>
                                    <%end if %>
                                </div>

                                <div class="row">
                                    <div class="col-lg-2"><a href="#">Tilføj grp.</a></div>
                                </div>


                                <%avansproobr = 0 %>
                                <%if avansproobr = 1 then%>
                                <!--Skal kun ske hvis avanc. er slået til i kontrol panel -->
                                <div class="row">
                                        <div class="col-lg-1">&nbsp</div>
                                        <div class="col-lg-1">Push:</div>
                                        <div class="col-lg-2"><input type="checkbox" name="pushbtn" value="push"></div>
                                </div>
                                <!-- avnas. slut -->
                                <%end if %>

                            </div> <!-- /.panel-body -->
                        </div> <!-- /.panel-collapse -->
                    </div> <!-- /.panel -->
                   </div>


                
                

                    <div class="panel-group accordion-panel" id="accordion-paneled">
                    <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapse4">
                                Aktiviteter
                            </a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapse4" class="panel-collapse collapse in">
                            <div class="panel-body">
                            
                             <script src="js/tableedittest.js" type="text/javascript"></script>
                              <table class="table dataTable table-striped table-bordered table-hover ui-datatable editabletable">
                                  <thead>
                                      <tr>
                                          <th style="width:20%">Aktivitet</th>
                                          <th style="width:10%">Status</th>
                                          <th style="width:9%">Type</th>
                                          <th style="width:5%">Timer</th>
                                          <th style="width:5%">Pris</th>
                                          <th style="width:5%">Start dato</th>
                                          <th style="width:5%">Slut dato</th>
                                          <%if avansproobr = 1 then%>
                                          <th style="width:5%">Start tidspunkt</th>
                                          <th style="width:5%">Slut tidspunkt</th>
                                          <%end if %>
                                          <th style="width:2%">Res</th>
                                          <th style="width:2%">Slet</th>
                                      </tr>
                                  </thead>

                                  <tbody>
                                      <tr>
                                          <td><a href="#">aa</a></td>
                                          <td><div contenteditable>hej </div></td>
                                          <td>fakt. bar</td>
                                          <td>6</td>
                                          <td>122 kr</td>
                                          <td>09-07</td>
                                          <td>12-07</td>
                                          <%if avansproobr = 1 then%>
                                          <td>08:30</td>
                                          <td>10:30</td>
                                          <%end if %>
                                          <td><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></td>
                                          <td><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></td>
                                      </tr>

                                      <tr>
                                          <td><a href="#">aa</a></td>
                                          <td><div contenteditable>hej </div></td>
                                          <td>fakt. bar</td>
                                          <td>6</td>
                                          <td>122 kr</td>
                                          <td>09-07</td>
                                          <td>12-07</td>
                                          <%if avansproobr = 1 then%>
                                          <td>08:30</td>
                                          <td>10:30</td>
                                          <%end if %>
                                          <td><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></td>
                                          <td><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></td>
                                      </tr>
                                  </tbody>
                              </table>
                             <div class="row">
                                 <div class="col-lg-3"><a href="#">Tilføj aktiitet</a></div>
                             </div>
                            <div class="row">
                                 <div class="col-lg-3"><a href="#">Tilføj stam akt. gruppe</a></div>
                             </div>

                            </div> <!-- /.panel-body -->
                        </div> <!-- /.panel-collapse -->
                    </div>
   
             
                    
                   </div>

                <div class="panel-group accordion-panel" id="accordion-paneled">
                    <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapse5">
                                Avancerede
                            </a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapse5" class="panel-collapse collapse in">
                            <div class="panel-body">
                                <div class="row">
                                    <div class="col-lg-2">Job medansvarlig:</div>
                                    <div class="col-lg-3">
                                        <select class="form-control input-small">
                                            <option value="1">Henrik</option>
                                            <option value="2">Søren</option>
                                            <option value="3">jesper</option>
                                        </select>
                                    </div>
                                    <div class="col-lg-1"><input class="form-control input-small" type="number" name="FM_exch" value="100 %" style="text-align:right"/></div>
                                    <div class="col-lg-1"><a href="#"><span class="glyphicon glyphicon-plus"></span></a></div>
                                </div>
                                <br />
                                <div class="row">                                                           
                                    <div class="col-lg-2">Jobejer:</div>
                                    <div class="col-lg-3">
                                        <select class="form-control input-small">
                                            <option value="1">Henrik</option>
                                            <option value="2">Søren</option>
                                            <option value="3">jesper</option>
                                        </select>
                                    </div>
                                </div>
                             
                                <br />

                                <%if func = "red" then %>
                                <div class="row">
                                    <div class="col-lg-7"><input type="checkbox" name="FM_opdaterprojektgrupper" id="FM_opdaterprojektgrupper" value="1" <%=syncAktProjGrpCHK %>> <b>Overfør</b> <!--(synkroniser) valgte projektgrupper, 
                                til <b>aktiviteterne</b> på dette job.--> </div>
                                </div>
                                <%end if %>

                                 <%
                           if lto <> "execon" then%>
                                <div class="row">
                                    <div class="col-lg-7">
								        <input type="checkbox" name="FM_gemsomdefault" id="FM_gemsomdefault" value="1"><b> Skift <!--standard--> forvalgt projektgruppe</b> <a data-toggle="modal" href="#styledModalSstGrp20"><span class="fa fa-info-circle"></span></a> <!--til den gruppe der her vælges som projektgruppe 1.
								        <span style="color:#999999;">Gemmes som cookie i 30 dage.</span> -->
                                    </div>
                                </div>
                                <div id="styledModalSstGrp22" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
                                        <div class="modal-dialog">
                                            <div class="modal-content" style="border:none !important;padding:0;">
                                              <div class="modal-header">
                                                <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                                  <h5 class="modal-title"><%=tsa_txt_medarb_069 %></h5>
                                                </div>
                                                <div class="modal-body">
                                                <%=tsa_txt_medarb_104 %>

                                                </div>
                                            </div>
                                        </div>
                                 </div>
                                <%end if %>
								
                               
                              

                                <%if func <> "red" then
                                 
                                    if lto = "jm" OR lto = "synergi1" OR lto = "micmatic" then 'OR lto = "lyng" OR lto = "glad" then
                                    forvalgCHK = "CHECKED"
                                    else
                                    forvalgCHK = ""
                                    end if
                                 
                                 else
                                 
                                 forvalgCHK = ""

                                 end if %>

                                <br /><br />

                               <div id="accordion-help" class="panel-group accordion-simple">
                                <div class="panel">
                                
                                      <a data-toggle="collapse" data-parent="#accordion-help" href="#faq-general-1">Forretingsområder:</a>
                                      <!-- <br /> <span>Forretningsområder bruges bl.a. til at se tidsforbrug på tværs af kunder og job, og til at se hvilke slags opgaver man bruger sin tid på. 
                                       <br />
                                       Alle aktiviteter på dette job tæller altid med i de forretningsområder der er valgt på jobbet. Specifikke forretningsområder kan vælges på den enkelte aktivitet.</span>
                                       -->
                                  <div id="faq-general-1" class="panel-collapse collapse">
                                    <div class="panel-body">
                                         <%
                                         ' uTxt = "Forretningsområder bruges bl.a. til at se tidsforbrug på tværs af kunder og job, og til at se hvilke slags opgaver man bruger sin tid på."
                                         ' uWdt = 300
								
								        'call infoUnisport(uWdt, uTxt) 
                                
                                    
                                            call fomr_mandatory_fn()

                                            div_tild_forr_Pos = "relative"
                                            div_tild_forr_Lft = "0px"
                                            div_tild_forr_Top = "0px"
                                            div_tild_forr_z = "0"

                                            if cint(fomr_mandatoryOn) = 1 then
                                            div_tild_forr_VZB = "visible"
                                            div_tild_forr_DSP = ""

                                                if func = "opret" AND step = "2" then
                                                div_tild_forr_bdr = "10"
                                                div_tild_forr_pd = "20"
                                                else
                                                div_tild_forr_bdr = "0"
                                                div_tild_forr_pd = "20"
                                                end if

                                        
                                                if func = "opret" AND step = "2" then
                                                div_tild_forr_Pos = "absolute"
                                                div_tild_forr_Lft = "600px"
                                                div_tild_forr_Top = "500px"
                                                div_tild_forr_z = "400000"
                                                end if

                                            else
                                            div_tild_forr_VZB = "hidden"
                                            div_tild_forr_DSP = "none"

                                                div_tild_forr_bdr = "0"
                                                div_tild_forr_pd = "0"

                                            end if
                                            %>

                                            
                                            <%

                                            '** Finder kundetype, til forvalgte forrretningsområder '***
                                            thisKtype_segment = 0
                                            if func = "opret" AND step = 2 AND strKundeId <> "" then 


                                                strSQLktyp = "SELECT ktype FROM kunder WHERE kid = " & strKundeId
                                                oRec5.open strSQLktyp, oConn, 3
                                                if not oRec5.EOF then

                                                thisKtype_segment = oRec5("ktype")

                                                end if
                                                oRec5.close


                                            end if

                                
                                
                                            'strSQLf = "SELECT id, navn FROM fomr WHERE id <> 0 ORDER BY navn"
                                                strSQLf = "SELECT f.navn AS fnavn, f.id, f.konto, kp.kontonr AS kkontonr, kp.navn AS kontonavn, fomr_segment FROM fomr AS f "_
                                                &" LEFT JOIN kontoplan AS kp ON (kp.id = f.konto) WHERE f.id <> 0 AND f.jobok = 1 ORDER BY f.navn"


                                                'if session("mid") = 1 then
                                                'response.write strSQLf
                                                'response.flush
                                                'end if

                                                %>

                                        
                                        <div class="row">
                                            <div class="col-lg-5"><b>Konto:</b></div>
                                        </div>
                                            <div class="row">
                                                <div class="col-lg-12">    
                                                   
                                            <select name="FM_fomr" id="FM_fomr" multiple="multiple" size="5" class="form-control input-small">
                                            <option value="0">Ingen valgt</option>
                                    
                                                <%
                                                fa = 0
                                                strchkbox = ""
                                                oRec.open strSQLf, oConn, 3
                                                while not oRec.EOF

                                                    if func = "opret" AND step = 2 then '*** Opret (forvalgt)

                                                    if instr(oRec("fomr_segment"), "#"& thisKtype_segment &"#") <> 0 then
                                                    fSel = "SELECTED"
                                                    else
                                                    fSel = ""
                                                    end if


                                                    else '** Rediger Forretningsområder
                                    
                                                    if instr(strFomr_rel, "#"&oRec("id")&"#") <> 0 then
                                                    fSel = "SELECTED"
                                                    else
                                                    fSel = ""
                                                    end if

                                                    end if



                                                if oRec("konto") <> 0 then
                                                kontonrVal = " ("& left(oRec("kontonavn"), 10) &" "& oRec("kkontonr") &")"

                                                if cint(fomr_konto) = cint(oRec("id")) then
                                                fokontoCHK = "CHECKED"
                                                else
                                                fokontoCHK = ""
                                                end if

                                                strchkbox = strchkbox & "<input type='radio' class='FM_fomr_konto' id='FM_fomr_konto_"& oRec("id") &"' name='FM_fomr_konto' value="& oRec("id") &" "& fokontoCHK &"> " & left(oRec("fnavn"), 20) &" "& kontonrVal &"<br>"
                                                else
                                                kontonrVal = ""
                                                end if 
                                    
                                                %>
                                                <option value="<%=oRec("id")%>" <%=fSel %>><%=oRec("fnavn") &" "& kontonrVal %></option>
                                                <%
                                                    fa = fa + 1

                                    
                                                oRec.movenext
                                                wend
                                                oRec.close
                                                %>
                                            </select>
                                                 </div>
                                            </div>  
                                        <br />
                                       <!-- <div class="row">
                                            <div class="col-lg-12"><b>Forvalgt konto på faktura / ERP system</b><br />
                                            Vælg herunder blandt de forretningsområder der har tilknyttet en omsætningskonto, og hvor fakturaer på dette job skal posteres på denne konto:<br />  
                                            <%=strchkbox %></div>
                                        </div>   -->
                                        
                                        <%if func <> "red" then

                                        select case lto
                                        case "hestia", "intranet - local"
                                        fomr_sync_CHK = "CHECKED"
                                        case else
                                        fomr_sync_CHK = ""
                                        end select

                                        else
                                        fomr_sync_CHK = ""
                                        end if

                                        %>                        

                                        </div> <!-- /.panel-body -->
                                    </div> <!-- /.panel-collapse -->
                                </div>
   
             
                    
                                </div>

                                <br />

                                <div id="accordion-help" class="panel-group accordion-simple">
                                <div class="panel">
                                
                                      <a data-toggle="collapse" data-parent="#accordion-help" href="#faq-general-2">Avanceret indstillinger:</a>
                                       <!--<br /> <span>Tildel bla. prioitet, faktura-indstillinger, pre-konditioner, kundeadgang mm.</span>
                                        -->
                                  <div id="faq-general-2" class="panel-collapse collapse in">
                                    <div class="panel-body">
                                                     
                                        <div class="row">
                                            <div class="col-lg-2"><b>Prioitet:</b> <a data-toggle="modal" href="#styledModalSstGrp23"><span class="fa fa-info-circle"></span></a></div>
                                            <div class="col-lg-3">&nbsp</div>
                                            <div class="col-lg-4"><b>Pre-konditioner opfyldt:</b> <a data-toggle="modal" href="#styledModalSstGrp24"><span class="fa fa-info-circle"></span></a> <!--<br /><span style="font-size:90%; font-weight:lighter">Underleverandør klar, materialer indkøbt mm.</span>--></div>
                                        </div>
                                        <div id="styledModalSstGrp23" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
                                            <div class="modal-dialog">
                                                <div class="modal-content" style="border:none !important;padding:0;">
                                                  <div class="modal-header">
                                                    <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                                      <h5 class="modal-title"></h5>
                                                    </div>
                                                    <div class="modal-body">
                                                         Prioiteter på nuværende aktive job ligger mellem: <b><%=lowestRiskval%> - <%=highestRiskval%></b><br /><br />
									                     <b>-1 = Internt job</b> vises ikke under fakturering og igangværende job.<br />
                                                         <b>-2 = HR job</b> vises i HR mode på timereg. siden<br />
                                                         <b>-3 = Internt job</b> men der skal kunne laves ressouceforecast på dette job. 
									                     <br /><br />
                                                         -1 / -2 / -3 medfører enkel visning af aktivitetslinjer på timereg. siden, dvs. der bliver ikke vist tidsforbrug, start- og slut -datoer mv. på aktiviteterne. <br /><br />&nbsp;
									
                                                    </div>
                                                </div>
                                            </div>
                                        </div>  
                                        <div id="styledModalSstGrp24" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
                                            <div class="modal-dialog">
                                                <div class="modal-content" style="border:none !important;padding:0;">
                                                  <div class="modal-header">
                                                    <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                                      <h5 class="modal-title"></h5>
                                                    </div>
                                                    <div class="modal-body">                                                
                                                        Underleverandør klar, materialer indkøbt mm.
                                                    </div>
                                                </div>
                                            </div>
                                        </div> 
                                        
                                            <%strSQLr = "SELECT risiko FROM job WHERE risiko > -1 AND jobstatus = 1 ORDER BY risiko DESC"
									 
									     highestRiskval = 1
									     oRec4.open strSQLr, oConn, 3
									     if not oRec4.EOF then
									     highestRiskval = oRec4("risiko")
									     end if
									     oRec4.close
									 
									     %>
									 
									     <%strSQLr = "SELECT risiko FROM job WHERE risiko > -1 AND jobstatus  = 1 ORDER BY risiko"
									 
									     lowestRiskval = 0
									     oRec4.open strSQLr, oConn, 3
									     if not oRec4.EOF then
									     lowestRiskval = oRec4("risiko")
									     end if
									     oRec4.close
									 
									     %>
                                        <div class="row">
                                            <div class="col-lg-1"> <input id="prio" name="prio" type="text" value="<%=intprio %>" class="form-control input-small" /></div>
                                           <!-- <div class="col-lg-3">Prioiteter på nuværende aktive job ligger mellem: <b><%=lowestRiskval%> - <%=highestRiskval%></b></div> -->
                                            <div class="col-lg-4">&nbsp</div>
                                            <div class="col-lg-3">
                                                <select name="FM_preconditions_met" id="Select1" size="1" class="form-control input-small">
                                                <option value="0" <%=preconditions_met_SEL0 %>>Ikke angivet</option>
                                                <option value="1" <%=preconditions_met_SEL1 %>>Ja</option>
                                                <option value="2" <%=preconditions_met_SEL2 %>>Nej - afvent</option>
                                                </select>
                                            </div>
                                        </div>
                                        <div class="row">
                                            <div class="col-lg-12">
									     <!---<b>1 = Internt job</b>--> <!--vises ikke under fakturering og igangværende job.-->
                                         <!---<b><!---2 = HR job</b> --><!--vises i HR mode på timereg. siden-->
                                         <!---<b><!-- 3 = Internt job</b>--> <!--men der skal kunne laves ressouceforecast på dette job.--> 
									    
                                         <!---1 / -2 / -3 medfører enkel visning af aktivitetslinjer på timereg. siden, dvs. der bliver ikke vist tidsforbrug, start- og slut -datoer mv. på aktiviteterne. <br /><br />&nbsp;-->
									    </div>
                                        </div>  
                                        


                                        <%
                                        preconditions_met_SEL0 = "SELECTED"
                                        preconditions_met_SEL1 = ""
                                        preconditions_met_SEL2 = ""

                                        select case cint(preconditions_met)
                                        case 0
                                        preconditions_met_SEL0 = "SELECTED"
                                        case 1
                                        preconditions_met_SEL1 = "SELECTED"
                                        case 2
                                        preconditions_met_SEL2 = "SELECTED"
                                        case else
                                        preconditions_met_SEL0 = "SELECTED"
                                        end select
                                        %>

                                       
                                        
                                          
                                        <br />
                                        

                                        <div class="row">
                                            <div class="col-lg-4"><b>Fakturaindstillinger:</b><!--<br /><span style="font-size:11px; font-weight:lighter;">(Nedarves fra kunde ved joboprettelse)</span>--></div>
                                            <div class="col-lg-1">&nbsp</div>
                                            <div class="col-lg-3"><b>Skal job være åben for kunde?</b></div>

                                        </div>

                                        <div class="row">
                                            <div class="col-lg-1">Valuta:</div>
                                            <div class="col-lg-3"><select name="FM_valuta" class="form-control input-small">
							                <%strSQL = "SELECT id, valuta, valutakode, kurs FROM valutaer WHERE id <> 0 ORDER BY id " 
							                oRec.open strSQL, oConn, 3
							                while not oRec.EOF 
							                 if oRec("id") = cint(valuta) then
							                 vSEL = "SELECTED"
							                 else
							                 vSEL = ""
							                 end if%>
							                <option value="<%=oRec("id") %>" <%=vSEL %>><%=oRec("valuta") %> | <%=oRec("valutakode") %> | kurs: <%=oRec("kurs")%></option>
							                <%
							                oRec.movenext
							                wend
							                oRec.close%>
							
							                </select>

                                            </div>
                                            <div class="col-lg-1">&nbsp</div>
                                            <div class="col-lg-3"><input type="checkbox" name="FM_kundese" id="FM_kundese" value="1" <%=kundechk%>>&nbsp;Gør job tilgængeligt for kontakt.</div>
                                        </div>

                                        <div class="row">
                                            <div class="col-lg-1">Moms:</div>
                                            <div class="col-lg-3"><%strSQLmoms = "SELECT id, moms FROM fak_moms WHERE id <> 0 ORDER BY id " %>
                                            <select name="FM_jfak_moms" class="form-control input-small">

                                                <%oRec6.open strSQLmoms, oConn, 3
                                                while not oRec6.EOF 

                                                    if cint(jfak_moms) = cint(oRec6("id")) then
                                                    fakmomsSeL = "SELECTED"
                                                    else
                                                    fakmomsSeL = ""
                                                    end if

                                                %><option value="<%=oRec6("id") %>" <%=fakmomsSeL %>><%=oRec6("moms") %>%</option><%
                  
                                                oRec6.movenext
                                                wend 
                                                oRec6.close%>

                                            </select>

                                            </div>
                                            <div class="col-lg-1">&nbsp</div>
                                            <div class="col-lg-7"><b>Når job åbnes for kontakt:</b> <a data-toggle="modal" href="#styledModalSstGrp24"><span class="fa fa-info-circle"></span></a> <!--hvornår skal registrerede timer så være tilgængelige? --></div>
                                        </div>
                                        <div id="styledModalSstGrp25" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
                                            <div class="modal-dialog">
                                                <div class="modal-content" style="border:none !important;padding:0;">
                                                  <div class="modal-header">
                                                    <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                                      <h5 class="modal-title"></h5>
                                                    </div>
                                                    <div class="modal-body">                                                
                                                        Hvis job åbnes for kontakt, hvornår skal registrerede timer så være tilgængelige?
                                                    </div>
                                                </div>
                                            </div>
                                        </div>

                                        <div class="row">
                                            <div class="col-lg-1">Sprog:</div>
                                            <%strSQLsprog = "SELECT id, navn FROM fak_sprog WHERE id <> 0 ORDER BY id " %>
                                            <div class="col-lg-3"><%strSQLmoms = "SELECT id, moms FROM fak_moms WHERE id <> 0 ORDER BY id " %>
                                            <select name="FM_jfak_sprog" class="form-control input-small">

                                                <%oRec6.open strSQLsprog, oConn, 3
                                                while not oRec6.EOF 

                                                    if cint(jfak_sprog) = cint(oRec6("id")) then
                                                    faksprogSeL = "SELECTED"
                                                    else
                                                    faksprogSeL = ""
                                                    end if

                                                %><option value="<%=oRec6("id") %>" <%=faksprogSeL %>><%=oRec6("navn") %></option><%
                  
                                                oRec6.movenext
                                                wend 
                                                oRec6.close%>

                                            </select>

                                            </div>
                                            <div class="col-lg-1">&nbsp</div>
                                            <div class="col-lg-4"><input type="radio" name="FM_kundese_hv" value="0" <%=hvchk1%>> Offentliggør timer, så snart de er indtastet.</div>
                                        </div>
                                        <div class="row">
                                            <div class="col-lg-5">&nbsp</div>
                                            <div class="col-lg-6"><input type="radio" name="FM_kundese_hv" value="1" <%=hvchk2%>> Offentliggør først timer når jobbet er lukket. (afsluttet/godkendt)</div>
                                        </div>

                                        <br /><br />
                                        <%if func = "opret" AND step = 2 OR func = "red" then %>


                                       <!-- <div class="row">
                                            <h5 class="col-lg-4">Skal job være åben for kunde?</h5>
                                        </div>

                                        <div class="row">
                                            <div class="col-lg-6"><input type="checkbox" name="FM_kundese" id="FM_kundese" value="1" <%=kundechk%>>&nbsp;<b>Gør job tilgængeligt for kontakt.</b><br>
		                                    <!--Hvis tilgængelig for kontakt tilvælges, udsendes der en mail til kontakt-stamdata emailadressen, med link til kontakt loginside.--></div>
                                        <!-- </div> -->

                                        <%if func = "opret" then
		                                hvchk1 = "checked"
		                                hvchk2 = ""
		                                else
			                                if cint(intkundeok) = 2 then
			                                hvchk1 = ""
			                                hvchk2 = "checked"
			                                else
			                                hvchk1 = "checked"
			                                hvchk2 = ""
			                                end if
		                                end if%>
                                        <br />
                                        <!--<div class="row">
                                            <div class="col-lg-5">
                                                <b>Hvis job åbnes for kontakt, hvornår skal registrerede timer så være tilgængelige?</b><br>
		                                        <input type="radio" name="FM_kundese_hv" value="0" <%=hvchk1%>>Offentliggør timer, så snart de er indtastet.<br>
		                                        <input type="radio" name="FM_kundese_hv" value="1" <%=hvchk2%>>Offentliggør først timer når jobbet er lukket. (afsluttet/godkendt)
                                            </div>
                                        </div> -->

                                        <%end if %>
                                        </div> <!-- /.panel-body -->
                                    </div> <!-- /.panel-collapse -->
                                </div>
   
             
                    
                               </div>




                                                                 

                            </div> <!-- /.panel-body -->
                        </div> <!-- /.panel-collapse -->
                    </div>
                   </div>

                   <div class="panel-group accordion-panel" id="accordion-paneled">
                    <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapse7">
                                Økonomi
                            </a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapse7" class="panel-collapse collapse in">
                            <div class="panel-body">
                            
                                <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                                    <thead>
                                        <tr style="font-size:75%">
                                            <th style="border-right:hidden">Bruttoomsætning<br /><span style="font-size:75%">Nettooms. + Salgsomk.</span></th>
                                            <th colspan="4" style="text-align:right; border-right:hidden">= 5100</th>
                                            <th>&nbsp</th>
                                        </tr>

                                        <tr style="font-size:75%">
                                            <th>Nettoomkostning, timer <br /><span style="font-size:75%">Oms. før salgsomk.	</span></th>
                                            <th style="width:10%">Timer</th>
                                            <th style="width:10%">Timepris</th>
                                            <th style="width:10%">Faktor</th>
                                            <th style="width:10%">Beløb</th>
                                            <th>Salgs<br />timepris</th>
                                        </tr>
                                    </thead>

                                    <tbody>
                                        <tr style="font-size:75%">
                                            <td>Gns. timepris / kostpris: 0,00 / 305,56</td>
                                            <td><input class="form-control input-small" type="number" name="FM_exch" value="23"/></td>
                                            <td><input class="form-control input-small" type="number" name="FM_exch" value="31"/></td>
                                            <td><input class="form-control input-small" type="number" name="FM_exch" value="0"/></td>
                                            <td><input class="form-control input-small" type="number" name="FM_exch" value=""/></td>
                                            <td><input class="form-control input-small" type="number" name="FM_exch" value="82,00"/></td>
                                        </tr>
                                    </tbody>
                                  
                                </table>

                                <br /><br />

                                <div id="accordion-help" class="panel-group accordion-simple">
                                <div class="panel">
                                
                                      <a data-toggle="collapse" data-parent="#accordion-help" href="#fordeltimbu">Fordel timebudget på finansår:</a>
                                       
                                  <div id="fordeltimbu" class="panel-collapse collapse">
                                    <div class="panel-body">

                                        <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                                            <thead>
                                                <tr style="font-size:75%">
                                                    <th style="width:40%">Budget FY</th>
                                                    <td style="width:10%; text-align:right">Timer</td>
                                                    <td style="width:10%; text-align:right">Budgetår (FY)</td>
                                                </tr>
                                            </thead>
                                            <tbody>
                                                <tr style="font-size:75%">
                                                    <td>År 0</td>
                                                    <td><input class="form-control input-small" type="number" name="FM_exch" value="23" style="text-align:right"/></td>
                                                    <td><input class="form-control input-small" type="number" name="FM_exch" value="2016" style="text-align:right"/></td>
                                                </tr>
                                                <tr style="font-size:75%">
                                                    <td>År 1</td>
                                                    <td><input class="form-control input-small" type="number" name="FM_exch" value="23" style="text-align:right"/></td>
                                                    <td><input class="form-control input-small" type="number" name="FM_exch" value="2017" style="text-align:right"/></td>
                                                </tr>
                                                <tr style="font-size:75%">
                                                    <td>År 2</td>
                                                    <td><input class="form-control input-small" type="number" name="FM_exch" value="23" style="text-align:right"/></td>
                                                    <td><input class="form-control input-small" type="number" name="FM_exch" value="2018" style="text-align:right"/></td>
                                                </tr>
                                                <tr style="font-size:75%">
                                                    <td>År 3</td>
                                                    <td><input class="form-control input-small" type="number" name="FM_exch" value="23" style="text-align:right"/></td>
                                                    <td><input class="form-control input-small" type="number" name="FM_exch" value="2019" style="text-align:right"/></td>
                                                </tr>
                                                <tr style="font-size:75%">
                                                    <td>År 4</td>
                                                    <td><input class="form-control input-small" type="number" name="FM_exch" value="23" style="text-align:right"/></td>
                                                    <td><input class="form-control input-small" type="number" name="FM_exch" value="2020" style="text-align:right"/></td>
                                                </tr>
                                                
                                            </tbody>
                                        </table>

                                        
                                             
                                    </div> <!-- /.panel-body -->
                                    </div> <!-- /.panel-collapse -->
                                </div>                              
                                </div>

                                
                                <div class="row">                      
                                    <div class="col-lg-2"><b>Jobtype</b></div>
                                    <div class="col-lg-2"><input type="checkbox" name="#" value="#"> <b>Åbn</b> <a data-toggle="modal" href="#styledModalSstGrp26"><span class="fa fa-info-circle"></span></a> <!--for manuel indtastning og beregning af Brutto- og Netto -omsætning.--></div>
                                </div>
                                <div id="styledModalSstGrp26" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
                                            <div class="modal-dialog">
                                                <div class="modal-content" style="border:none !important;padding:0;">
                                                  <div class="modal-header">
                                                    <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                                      <h5 class="modal-title"></h5>
                                                    </div>
                                                    <div class="modal-body">                                                
                                                        Åbn for manuel indtastning og beregning af Brutto- og Netto -omsætning.
                                                    </div>
                                                </div>
                                            </div>
                                        </div>                 
                                <div class="row">
                                    <div class="col-lg-2"><input type="checkbox" name="#" value="#"> Fastpris <a data-toggle="modal" href="#styledModalSstGrp27"><span class="fa fa-info-circle"></span></a> <!--(bruttoomsætning benyttes ved fakturering)--> </div>
                                </div>
                                 <div id="styledModalSstGrp27" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
                                    <div class="modal-dialog">
                                        <div class="modal-content" style="border:none !important;padding:0;">
                                            <div class="modal-header">
                                            <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                                <h5 class="modal-title"></h5>
                                            </div>
                                            <div class="modal-body">                                                
                                               <b>Fastpris</b> (bruttoomsætning benyttes ved fakturering)
                                            </div>
                                        </div>
                                    </div>
                                </div>   
                                <div class="row">
                                    <div class="col-lg-2"><input type="checkbox" name="#" value="#"> Lbn. timer <a data-toggle="modal" href="#styledModalSstGrp28"><span class="fa fa-info-circle"></span></a> <!-- (timeforbrug på hver enkelt aktivitet * medarb. timepris benyttes ved fakturering) --> </div>
                                </div>
                                 <div id="styledModalSstGrp28" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
                                    <div class="modal-dialog">
                                        <div class="modal-content" style="border:none !important;padding:0;">
                                            <div class="modal-header">
                                            <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                                <h5 class="modal-title"></h5>
                                            </div>
                                            <div class="modal-body">                                                
                                               <b>Lbn. timer</b> (timeforbrug på hver enkelt aktivitet * medarb. timepris benyttes ved fakturering)
                                            </div>
                                        </div>
                                    </div>
                                </div>   

                            </div> <!-- /.panel-body -->
                        </div> <!-- /.panel-collapse -->
                    </div> <!-- /.panel -->
                   </div>



                    <div class="panel-group accordion-panel" id="accordion-paneled">
                    <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#mater">
                                Materialer
                            </a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="mater" class="panel-collapse collapse in">
                            <div class="panel-body">
                            
                              <div class="row">
                                  <div class="col-lg-2"><b>Salgsomkostninger</b></div>
                              </div>
                              <div class="row">
                                  <div class="col-lg-2">Antal linjer:</div>
                                  <div class="col-lg-1">
                                      <select class="form-control input-small">
                                            <option value="1">1</option>
                                            <option value="2">2</option>
                                            <option value="3">3</option>
                                            <option value="1">4</option>
                                            <option value="2">5</option>
                                            <option value="3">6</option>
                                            <option value="1">7</option>
                                            <option value="2">8</option>
                                            <option value="3">9</option>
                                            <option value="3">10</option>
                                      </select>
                                  </div>
                              </div>

                              <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                                  <thead>
                                      <tr style="font-size:75%">
                                          <th style="width:40%">* udgift/salgsomkost.</th>
                                          <th style="width:11%">Stk.</th>
                                          <th style="width:11%">Stk. pris</th>
                                          <th style="width:11%">Indkøbspris</th>
                                          <th style="width:11%">Faktor</th>
                                          <th style="width:11%">Salgspris</th>
                                          <th>&nbsp</th>
                                      </tr>
                                  </thead>
                                  <tbody>
                                      <tr style="font-size:75%">
                                          <td>&nbsp</td>
                                          <td><input class="form-control input-small" type="number" name="FM_exch" value="" placeholder="0,00"/></td>
                                          <td><input class="form-control input-small" type="number" name="FM_exch" value="" placeholder="0,00"/></td>
                                          <td><input class="form-control input-small" type="number" name="FM_exch" value="" placeholder="0,00"/></td>
                                          <td><input class="form-control input-small" type="number" name="FM_exch" value="" placeholder="0,00"/></td>
                                          <td><input class="form-control input-small" type="number" name="FM_exch" value="" placeholder="0,00"/></td>
                                          <td><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></td>

                                      </tr>
                                  </tbody>
                              </table>

                                <div class="row">
                                    <div class="col-lg-12"><a href="#">Tilføj linje</a>
                                        <a data-toggle="modal" href="#styledModalSstGrp20"><span class="fa fa-info-circle"></span></a>
                                    </div>
                                

                                <div id="styledModalSstGrp20" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
                                        <div class="modal-dialog">
                                            <div class="modal-content" style="border:none !important;padding:0;">
                                              <div class="modal-header">
                                                <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                                  <h5 class="modal-title"><%=tsa_txt_medarb_069 %></h5>
                                                </div>
                                                <div class="modal-body">
                                                <%=tsa_txt_medarb_104 %>

                                                </div>
                                            </div>
                                        </div>
                                 </div></div>
                                
                            </div> <!-- /.panel-body -->
                        </div> <!-- /.panel-collapse -->
                    </div> <!-- /.panel -->
                   </div>

               <br /><br /><br /><br /><br /><br /><br />
            </div>
           

        <%end select %>

        </div>
    </div>

    

<!--#include file="../inc/regular/footer_inc.asp"-->