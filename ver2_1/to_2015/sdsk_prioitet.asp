

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/sdsk_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->


<% 
    sub crmaktstatheader
%>
    <div class="row">
        <div class="col-lg-2"><img src="../ill/logo_bg.gif" width="163" height="53" alt="" border="0"></div>
        <div class="col-lg-2"><b>Timeout Kontrolpanel - Incident prioitet</b><br />
            Tilføj, fjern eller rediger Incident prioitet.
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
	respinse.end
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
	
	if len(request("lastedit")) <> 0 then
	lastedit = request("lastedit")
	else
	lastedit = 0
	end if
	
	select case func
	case "slet"

	'call  sdsktopmenu()

    '*** Her spørges om det er ok at der slettes en medarbejder ***

	    oskrift = "Incident - Type" 
        slttxta = "Du er ved at <b>SLETTE</b> en Type. Er dette korrekt?"
        slttxtb = "" 
        slturl = "sdsk_prioitet.asp?menu=tok&func=sletok&id="& id

        call sletcnf_2015(oskrift, slttxta, slttxtb, slturl)
	

	case "sletok"
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM sdsk_prioitet WHERE id = "& id &"")
	Response.redirect "sdsk_prioitet.asp?menu=tok&shokselector=1"

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
		
		function SQLBless2(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless2 = tmp
		end function
		
		strNavn = SQLBless(request("FM_navn"))
		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
		
		rsptid = SQLBless2(request("FM_rsptid"))
		
		dblfaktor = SQLBless2(request("FM_faktor"))
		if len(request("FM_kunweekend")) <> 0 then
		intkunweekend = 1
		else
		intkunweekend = 0
		end if
		
		if len(request("FM_prio_grp")) <> 0 then
		prio_grp = request("FM_prio_grp")
		else
		prio_grp = 1
		end if
		
		
		
		
		intadvisering = SQLBless2(request("FM_advisering"))
		
		if func = "dbopr" then
		oConn.execute("INSERT INTO sdsk_prioitet (navn, editor, dato, responstid, faktor, kunweekend, advisering, priogrp, type) "_
		&" VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', "& rsptid &", "_
		&" "& dblfaktor &", "& intkunweekend &", "& intadvisering &", "& prio_grp &")")
				
				
				strSQL = "SELECT id FROM sdsk_prioitet ORDER by id DESC LIMIT 0, 1"
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				lastedit = oRec("id")
				oRec.movenext
				wend
				
				oRec.close
				
		else
		oConn.execute("UPDATE sdsk_prioitet SET navn ='"& strNavn &"', editor = '" &strEditor &"', "_
		&" dato = '" & strDato &"', responstid = '"& rsptid &"', faktor = "& dblfaktor &", "_
		&" kunweekend = "& intkunweekend &", advisering = "& intadvisering &", priogrp = "& prio_grp &" WHERE id = "&id&"")
		
		lastedit = id
		
		end if
		
		
		
		Response.redirect "sdsk_prioitet.asp?menu=tok&shokselector=1&lastedit="&lastedit&"&FM_prio_grp="&prio_grp
		end if
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	
	strNavn = ""
	strTimepris = ""
	varSubVal = "opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	rsptid = 0
	intkunweekend = 1
	
	else
	strSQL = "SELECT navn, editor, dato, responstid, faktor, kunweekend, advisering, type FROM sdsk_prioitet WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	rsptid = oRec("responstid")
	dblfaktor = oRec("faktor")
	intkunweekend = oRec("kunweekend")
	dblAdvi = oRec("advisering")
	
	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "opdaterpil" 
	end if
	
	if len(request("prio_grp")) <> 0 then
	prio_grp = request("prio_grp")
	else
	prio_grp = 1
	end if
	
	
	if intkunweekend = 1 then
	kwChecked = "CHECKED"
	else
	kwChecked = ""
	end if

	%>

        <div id="styledhelpbox_1" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
            <div class="modal-dialog">
                <div class="modal-content" style="border:none !important;padding:0;">
                    <div class="modal-header">
                    <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                        <h5 class="modal-title">Responder</h5>
                    </div>
                    <div class="modal-body">
                        Responder kun indenfor normal arbejds / åbningstid? <br />
                        Arbejdstid sættes i kontrolpanelet.
                    </div>
                </div>
            </div>
        </div>

        <div id="styledhelpbox_2" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
            <div class="modal-dialog">
                <div class="modal-content" style="border:none !important;padding:0;">
                    <div class="modal-header">
                    <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                        <h5 class="modal-title">Advisering</h5>
                    </div>
                    <div class="modal-body">
                        Efter antal dage med inaktivitet.
                    </div>
                </div>
            </div>
        </div>

        <div id="styledhelpbox_3" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
            <div class="modal-dialog">
                <div class="modal-content" style="border:none !important;padding:0;">
                    <div class="modal-header">
                    <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                        <h5 class="modal-title">Faktor</h5>
                    </div>
                    <div class="modal-body">
                        Bruges til at bestemme hvor mange klip der<br> skal tages fra en aftale, for hver time der registreres.
                    </div>
                </div>
            </div>
        </div>


        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u>Incident Type - <%=varbroedkrumme%></u></h3>
                <div class="portlet-body">
               
                    <%
	                call  sdsktopmenu()
	                %>

                    <form action="sdsk_prioitet.asp?menu=tok&func=<%=dbfunc%>" method="post">
	                <input type="hidden" name="FM_prio_grp" value="<%=prio_grp%>">
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
                            <div class="col-lg-1"></div>
                            <div class="col-lg-2">Navn:</div>
                            <div class="col-lg-3"><input type="text" name="FM_navn" value="<%=strNavn%>" class="form-control input-small"></div>
                            <div class="col-lg-2">Responstid: (0 = ingen)</div>
                            <div class="col-lg-1"><input type="text" name="FM_rsptid" value="<%=rsptid%>" size="5" class="form-control input-small"></div>
                            <div class="col-lg-1">timer</div>
                        </div>

                        <div class="row">
                            <div class="col-lg-1"></div>
                            <div class="col-lg-2">Responder indenfor normal arbejdestid  <a data-toggle="modal" href="#styledhelpbox_1"><span class="fa fa-info-circle"></span></a></div>
                            <div class="col-lg-3"><input type="checkbox" name="FM_kunweekend" value="1" <%=kwChecked%>></div>
                            <div class="col-lg-2">Advisering efter: <a data-toggle="modal" href="#styledhelpbox_2"><span class="fa fa-info-circle"></span></a></div>
                            <div class="col-lg-1"><input type="text" name="FM_advisering" value="<%=dblAdvi%>" class="form-control input-small"></div>
                            <div class="col-lg-1">dage</div>
                        </div>

                        <div class="row">
                            <div class="col-lg-1"></div>
                            <div class="col-lg-2">Faktor: <a data-toggle="modal" href="#styledhelpbox_3"><span class="fa fa-info-circle"></span></a></div>
                            <div class="col-lg-1"><input type="text" name="FM_faktor" value="<%=dblFaktor%>" class="form-control input-small"></div>                       
                        </div>

                    </div>
                    </form>

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
                    
                </div>


            </div>
        </div>

        <%case else 

            if len(request("FM_prio_grp")) <> 0 then
	        prio_grp = request("FM_prio_grp")
	        else
	        prio_grp = 1
	        end if    
        %>

        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u>Incident Type</u></h3>
                <div class="portlet-body">
                    <%
	                call  sdsktopmenu()
	                %>

                    <div class="row">
                        <div class="col-lg-10">&nbsp</div>
                        <div class="col-lg-2 pad-b10">
                            <a href="sdsk_prioitet.asp?menu=tok&func=opret&prio_grp=<%=prio_grp%>" class="btn btn-success btn-sm pull-right"><b>Opret ny Type</b></a>
                        </div>
                    </div>

                    <form method="post" action="sdsk_prioitet.asp">
                        <div class="well">
                            <div class="row">
                                <div class="col-lg-2">Aftalegruppe:</div>
                                <div class="col-lg-3">
                                    <select name="FM_prio_grp" id="FM_prio_grp" class="form-control input-small">
	                                <%
	                                strSQL = "SELECT id, navn FROM sdsk_prio_grp WHERE id <> 0"
	                                oRec.open strSQL, oConn, 3
	                                while not oRec.EOF
		
		
		                                if cint(prio_grp) = oRec("id") then
		                                pgSEL = "SELECTED"
		                                else
		                                pgSel = ""
		                                end if%>
	                                <option value="<%=oRec("id")%>" <%=pgSel%>><%=oRec("navn")%></option>
	                                <%
	                                oRec.movenext
	                                wend
	
	                                oRec.close
	                                %>
	                                </select>
                                </div>
                                <div class="col-lg-7 pull-right">
                                    <button type="submit" class="btn btn-secondary btn-sm pull-right"><b>Vælg</b></button>
                                </div>

                            </div>
                        </div>
                    </form>

                    <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                        <thead>
                            <tr>
                                <th>Id</th>
                                <th>Type (Antal inci.)</th>
                                <th><b>Faktor</b></th>
		                        <th><b>Responstid (Timer)</b></th>
                                <th></th>
                            </tr>
                        </thead>

                        <tbody>

                            <%
	                        strSQL = "SELECT p.id, p.navn, p.responstid, p.faktor, p.priogrp, COUNT(s.id) AS antal FROM sdsk_prioitet p "_
	                        &" LEFT JOIN sdsk s ON (s.prioitet = p.id) WHERE p.priogrp = "& prio_grp &" GROUP BY p.id ORDER BY p.navn"
	
	                        oRec.open strSQL, oConn, 3
	                        while not oRec.EOF 
	
	                        if cint(oRec("id")) = cint(lastedit) then
	                        bgCol = "#ffff99"
	                        else
	                        bgCol = "#ffffff"
	                        end if%>


                            <tr>
                                <td><%=oRec("id")%></td>
		                        <td><a href="sdsk_prioitet.asp?menu=tok&func=red&id=<%=oRec("id")%>&prio_grp=<%=oRec("priogrp")%>"><%=oRec("navn")%> </a>&nbsp;&nbsp;(<%=oRec("antal")%>)</td>
		                        <td><%=oRec("faktor")%></td>
		                        <td><%=oRec("responstid")%></td>
		                        <td>
		                        <%if cint(oRec("antal")) = 0 then%>
		                        <a href="sdsk_prioitet.asp?menu=tok&func=slet&id=<%=oRec("id")%>"><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></a>
		                        <%end if%>&nbsp;</td>
                            </tr>
                            <%
	                        x = 0
	                        oRec.movenext
	                        wend
	                        %>

                        </tbody>
                    </table>

                    

                </div>
            </div>
        </div>


        <%end select %>
    </div>
</div>

<!--#include file="../inc/regular/footer_inc.asp"-->