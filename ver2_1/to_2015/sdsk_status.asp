

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/sdsk_func.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->


<%
    sub crmaktstatheader
%>
    <div class="row">
        <div class="col-lg-2"><img src="../ill/logo_bg.gif" width="163" height="53" alt="" border="0"></div>
        <div class="col-lg-2"><b>Timeout Kontrolpanel - Incident Status</b><br />
            Tilføj, fjern eller rediger Incident Status.
        </div>
    </div>


   <!-- <table border=0 cellpadding=0 cellspacing=0 width="450">
	<tr>
	<td valign="top" width="163"><img src="../ill/logo_bg.gif" width="163" height="53" alt="" border="0"></td>
	<td valign="bottom"><b>Timeout Kontrolpanel - Incident Status</b><br>
	Tilføj, fjern eller rediger Incident Status.</td>
	</tr>
	</table><br> -->
<%
    end sub


    if len(session("user")) = 0 then
	%>
	
	<%
	errortype = 5
	call showError(errortype)
	response.End
    end if


    call menu_2014 %> 
        
    <div class="wrapper">
        <div class="content">

    <%    
    func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	thisfile = "sdsk_kontrolpanel"
	kview = "j"        
        
    select case func
	case "slet"

    'call sdsktopmenu()

    '*** Her spørges om det er ok at der slettes en medarbejder ***
	    oskrift = "Incident - Status" 
        slttxta = "Du er ved at <b>SLETTE</b> en status. Er dette korrekt?"
        slttxtb = "" 
        slturl = "sdsk_status.asp?menu=tok&func=sletok&id="& id

        call sletcnf_2015(oskrift, slttxta, slttxtb, slturl)


    case "sletok"
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM sdsk_status WHERE id = "& id &"")
	Response.redirect "sdsk_status.asp?menu=tok&shokselector=1"


    case "dbopr", "dbred"
	'*** Her indsættes en ny type i db ****
		if len(request("FM_navn")) = 0 then
		%>
		    
		<%
		errortype = 8
		call showError(errortype)
		
		else
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function
		
		strNavn = SQLBless(request("FM_navn"))
		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
		strFarve = "#FFFFFF" 'request("FM_farve")
		progrp = request("FM_progrp")
		
		if len(request("FM_luk")) <> 0 then
		intLuk = request("FM_luk")
		else
		intLuk = 0
		end if
		
		if len(request("FM_vispahovedliste")) <> 0 then
		vispahovedliste = request("FM_vispahovedliste")
		else
		vispahovedliste = 0
		end if
		
		if func = "dbopr" then
		oConn.execute("INSERT INTO sdsk_status (navn, editor, dato, farve, vispahovedliste, progrp, luk) VALUES "_
		&" ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', '"& strFarve &"', "& vispahovedliste &", "& progrp &", "& intLuk &")")
		else
		oConn.execute("UPDATE sdsk_status SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"', farve = '"& strFarve &"', "_
		&" vispahovedliste = "& vispahovedliste &", progrp = "& progrp &", luk = "& intLuk &" WHERE id = "&id&"")
		end if
		
		Response.redirect "sdsk_status.asp?menu=tok&shokselector=1"
		end if
	
	case "rsptid"
	
	rsptid_a = request("FM_rsptid_a")
	rsptid_b = request("FM_rsptid_b")
	
	
	strSQL = "UPDATE sdsk_status SET rsptid_a = 0 WHERE id <> 0" 
	oConn.execute(strSQL)
	strSQL = "UPDATE sdsk_status SET rsptid_a = 1 WHERE id = " & rsptid_a
	oConn.execute(strSQL)
	strSQL = "UPDATE sdsk_status SET rsptid_b = 0 WHERE id <> 0" 
	oConn.execute(strSQL)
	strSQL = "UPDATE sdsk_status SET rsptid_b = 1 WHERE id = " & rsptid_b
	oConn.execute(strSQL)
	
	Response.redirect "sdsk_status.asp"
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	progrp = 0
	
	else
	strSQL = "SELECT navn, editor, dato, farve, vispahovedliste, progrp, luk FROM sdsk_status WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	strFarve = oRec("farve")
	vispahovedliste = oRec("vispahovedliste")
	progrp = oRec("progrp")
	intLuk = oRec("luk")
	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "opdaterpil" 
	end if



    %>
        <div id="styledhelpbox_1" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
            <div class="modal-dialog">
                <div class="modal-content" style="border:none !important;padding:0;">
                    <div class="modal-header">
                    <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                        <h5 class="modal-title">Statusgruppe</h5>
                    </div>
                    <div class="modal-body">
                        Denne statusgruppe bliver vist med <s>overstreget skrift på incident listen.</s><br />
		                Derfor bør denne status hedde "Lukket" ell. "Afsluttet og godkendt" etc.
                    </div>
                </div>
            </div>
        </div>

        <div id="styledhelpbox_2" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
            <div class="modal-dialog">
                <div class="modal-content" style="border:none !important;padding:0;">
                    <div class="modal-header">
                    <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                        <h5 class="modal-title">Status</h5>
                    </div>
                    <div class="modal-body">
                        Denne status skal medtages i forvalgte. (Vises på Incidentlisten ved åbning.)
                    </div>
                </div>
            </div>
        </div>

        <div id="styledhelpbox_3" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
            <div class="modal-dialog">
                <div class="modal-content" style="border:none !important;padding:0;">
                    <div class="modal-header">
                    <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                        <h5 class="modal-title">Projektgruppe</h5>
                    </div>
                    <div class="modal-body">
                        Adviser denne projektgruppe<br>
		                når der skiftes til denne status.
                    </div>
                </div>
            </div>
        </div>

        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u>Incident Status - <%=varbroedkrumme%></u></h3>
                <div class="portlet-body">

                    <%
                        call sdsktopmenu() 
                    %>

                    <form action="sdsk_status.asp?menu=tok&func=<%=dbfunc%>" method="post">
                    <input type="hidden" name="id" value="<%=id%>">
                    <div class="row">
                        <div class="col-lg-10">&nbsp</div>
                        <div class="col-lg-2 pad-b10">
                            <button type="submit" class="btn btn-success btn-sm pull-right"><b>Opdatér</b></button>
                        </div>
                    </div>

                    <div class="well well-white">

                        <div class="row">
                            <div class="col-lg-12">
                                <h4 class="panel-title-well">Stamdata</h4>
                            </div>
                        </div>

                        <div class="row">
                            <div class="col-lg-1">&nbsp</div>
                            <div class="col-lg-2">Navn: 
                             <%if id = 1 OR id = 4 then%>
                             <a data-toggle="modal" href="#styledhelpbox_1"><span class="fa fa-info-circle"></span></a>
                             <%end if %>
                            </div>

                            <div class="col-lg-3"><input type="text" name="FM_navn" value="<%=strNavn%>" class="form-control input-small"></div>
                        

                            <%if cint(vispahovedliste) = 1 then
		                    vispahovedlisteCHK = "CHECKED"
		                    else
		                    vispahovedlisteCHK = ""
		                    end if%>
                            <div class="col-lg-4"><input type="checkbox" name="FM_vispahovedliste" id="FM_vispahovedliste" value="1" <%=vispahovedlisteCHK%>> Denne status skal medtages i forvalgte. <a data-toggle="modal" href="#styledhelpbox_2"><span class="fa fa-info-circle"></span></a></div>
                        </div>

                        <div class="row">

                            <div class="col-lg-1">&nbsp</div>
                            <div class="col-lg-2">Projektgruppe: <a data-toggle="modal" href="#styledhelpbox_3"><span class="fa fa-info-circle"></span></a></div>
                            <div class="col-lg-3">
                                <select name="FM_progrp" id="FM_progrp" class="form-control input-small">
		                            <option value="0">Ingen</option>
		                            <%
		                            strSQL = "SELECT id, navn FROM projektgrupper WHERE id <> 0 ORDER BY navn"
		                            oRec.open strSQL, oConn, 3
		                            while not oRec.EOF
		
			                            if cint(progrp) = oRec("id") then
			                            pSel = "SELECTED"
			                            else
			                            pSel = ""
			                            end if%>
		                            <option value="<%=oRec("id")%>" <%=pSel%>><%=oRec("navn")%></option>
		                            <%
		                            oRec.movenext
		                            wend
		
		                            oRec.close
		                            %>
		                        </select>
                            </div>
                            <div class="col-lg-4">
                                <%if cint(intLuk) = 1 then
		                        lukCHK = "CHECKED"
		                        else
		                        lukCHK = ""
		                        end if%>
                                <input type="checkbox" name="FM_luk" id="FM_luk" value="1" <%=lukCHK%>> Luk aktivitet når der skiftes til denne status.
                            </div>

                        </div>

                    </div>



                    <br /><br />
                    <div class="row">
                        <div class="col-lg-2">
                            <a href="Javascript:history.back()" class="btn btn-default btn-sm"><b>Tilbage</b></a>
                        </div>
                    </div>


                    <%if dbfunc = "dbred" then%>
                    <br /><br /><br /><br />
                    <div class="row">
                        <div class="col-lg-12">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></div>
                    </div>
                    <%end if %>

                </form>
                </div>
                
                <!--------------- Liste --------------------->
                <%case else %>

                <div id="styledhelpbox_4" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
                    <div class="modal-dialog">
                        <div class="modal-content" style="border:none !important;padding:0;">
                            <div class="modal-header">
                            <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                <h5 class="modal-title">Side info</h5>
                            </div>
                            <div class="modal-body">
                                Respons tider bruges til at holde øje med hvor lang der går fra en incident er oprettet til den skifter til den valgte status.
			                    <br><br>Når en responstid er gemt på en incident kan denne ikke overskrives, 
			                    men hvis oprettelses tiden på det pågældende incident fornyes bliver respons tiderne A og B nulstillet.
                            </div>
                        </div>
                    </div>
                </div>


                <div class="container">
                    <div class="portlet">
                        <h3 class="portlet-title"><u>Incident Status</u></h3>
                        <div class="portlet-body">
                            
                            <%
                                call sdsktopmenu() 
                            %>


                            <div class="row">
                                <div class="col-lg-10">&nbsp</div>
                                <div class="col-lg-2 pad-b10">
                                    <a href="sdsk_status.asp?menu=tok&func=opret" class="btn btn-success btn-sm pull-right"><b>Opret ny Status</b></a>
                                </div>
                            </div>
                            

                            <form action="sdsk_status.asp?func=rsptid" method="post">
                            <table class="table dataTable table-striped table-bordered table-hover ui-datatable">

                                <thead>
                                    <tr>
                                        <th>Id</th>
                                        <th>Status navn</th>
                                        <th>Slet?</th>
                                        <th>Responstid A</th>
                                        <th>Responstid B</td>
		                                <th>Luk aktivitet?</th>
                                    </tr>
                                </thead>

                                <tbody>

                                    <%
	                                'sort = Request("sort")
	                                'if sort = "navn" then
	                                'strSQL = "SELECT id, navn, rsptid_a, rsptid_b FROM sdsk_status ORDER BY navn"
	                                'else
	                                strSQL = "SELECT id, navn, rsptid_a, rsptid_b, luk FROM sdsk_status ORDER BY rsptid_a DESC, rsptid_b DESC"
	                                'end if
	
	                                oRec.open strSQL, oConn, 3
	                                while not oRec.EOF 
	                                %>
                                    <tr>
                                        <td><%=oRec("id")%></td>
                                        <td><a href="sdsk_status.asp?menu=tok&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%> </a></td>
                                        <td align="center">
		                                <%if oRec("id") > 6 then %>
		                                <a href="sdsk_status.asp?menu=tok&func=slet&id=<%=oRec("id")%>"><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></a>
		                                <%else %>
		                                sys.grp.
		                                <%end if %>
		                                </td>
                                        <td align=center>
		                                <%if cint(oRec("rsptid_a"))  = 1 then
		                                rsptid_a_CHK = "CHECKED"
		                                else
		                                rsptid_a_CHK = ""
		                                end if%>
		
		                                <input type="radio" name="FM_rsptid_a" value="<%=oRec("id")%>" <%=rsptid_a_CHK%>>
                                        </td>
                                        <td align=center>
		                                <%if cint(oRec("rsptid_b"))  = 1 then
		                                rsptid_b_CHK = "CHECKED"
		                                else
		                                rsptid_b_CHK = ""
		                                end if%>
		
		                                <input type="radio" name="FM_rsptid_b" value="<%=oRec("id")%>" <%=rsptid_b_CHK%>></td>
		                                <td>
		                                <%
		                                if oRec("luk") = 1 then
		                                Response.write "Lukker akt."
		                                else
		                                Response.write "&nbsp;"
		                                end if
		                                %></td>
                                    </tr>
                                    <%
                                        x = 0
	                                    oRec.movenext
	                                    wend 
                                    %>
                                </tbody>

                            </table>

                            <div class="row">
                                <div class="col-lg-2">
                                    
                                </div>
                                <div class="col-lg-8">&nbsp</div>
                                <div class="col-lg-2 pad-b10">
                                    <button type="submit" class="btn btn-secondary btn-sm pull-right"><b>Opdater Rsp. tider</b></button>
                                </div>
                            </div>

                            <br /><br />
                            <div class="row">
                                <div class="col-lg-4">
                                    Side info: <a data-toggle="modal" href="#styledhelpbox_4"><span class="fa fa-info-circle"></span></a>
                                </div>
                            </div>

                            </form>

                            

                        </div>
                    </div>
                </div>
            </div>
        </div>


    

    
<%end select %>

</div></div>

<!--#include file="../inc/regular/footer_inc.asp"-->