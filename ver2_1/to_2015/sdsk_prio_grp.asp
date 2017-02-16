

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
        <div class="col-lg-4"><b>Timeout Kontrolpanel - Aftalegrupper.</b><br />
            Tilføj, fjern eller rediger Aftalegrupper.
        </div>
    </div>
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
    '*** Her spørges om det er ok at der slettes en medarbejder ***

	oskrift = "Aftalegruppe" 
    slttxta = "Du er ved at <b>SLETTE</b> en Aftalegruppe. Er dette korrekt?"
    slttxtb = "" 
    slturl = "sdsk_prio_grp.asp?menu=tok&func=sletok&id="& id

    call sletcnf_2015(oskrift, slttxta, slttxtb, slturl)

	case "sletok"
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM sdsk_prio_grp WHERE id = "& id &"")
	oConn.execute("DELETE FROM sdsk_prioitet WHERE priogrp = "& id &"")
	Response.redirect "sdsk_prio_grp.asp?menu=tok&shokselector=1" 

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
		if len(request("FM_gemtider")) <> 0 then
		gemtider = request("FM_gemtider")
		else
		gemtider = 0
		end if
		
		if func = "dbopr" then
		oConn.execute("INSERT INTO sdsk_prio_grp (navn, editor, dato, gemtider) VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', "& gemtider &")")
		else
		oConn.execute("UPDATE sdsk_prio_grp SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"', gemtider = "& gemtider &" WHERE id = "&id&"")
		end if
		
		Response.redirect "sdsk_prio_grp.asp?menu=tok&shokselector=1"
		end if
		
	case "kopier"
		
		
		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
		
		strSQL = "SELECT navn, id, gemtider FROM sdsk_prio_grp WHERE id = "& id
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then
						
						
						'*** Kopierer grp ***
						strSQLkopi = "INSERT INTO sdsk_prio_grp (editor, dato, navn, gemtider) VALUES "_
						&" ('"& strEditor &"', '"& strDato &"', '"& oRec("navn") &" (kopi)', "& oRec("gemtider") &")"
						oConn.execute(strSQLkopi)
						
						strSQL3 = "SELECT id FROM sdsk_prio_grp WHERE id <> 0 ORDER BY id DESC"
						oRec3.open strSQL3, oConn, 3
						if not oRec3.EOF then
						
						nygrpId = oRec3("id")
						
						end if
						oRec3.close
						
						'** Henter prioiteter ***
						strSQL2 = "SELECT navn, responstid, faktor, kunweekend, advisering, type FROM sdsk_prioitet WHERE priogrp =" & oRec("id")
						oRec2.open strSQL2, oConn, 3
						
						while not oRec2.EOF 
						
						strNavn = oRec2("navn")
						rsptid = replace(oRec2("responstid"), ",", ".")
						dblfaktor = replace(oRec2("faktor"), ",", ".")
						intkunweekend = oRec2("kunweekend")
						dblAdvi = replace(oRec2("advisering"), ",", ".")
						intType = oRec2("type")
						
							
								strSQLkopi2 = "INSERT INTO sdsk_prioitet ( editor, dato, navn, responstid, faktor, "_
								&" kunweekend, advisering, priogrp, type) "_
								&" VALUES ('"& strEditor &"', '"& strDato &"', '"& strNavn &"', "& rsptid &", "& dblfaktor &", "_
								&" "& intkunweekend &", "& dblAdvi &", "& nygrpId &", "& intType &")"
								oConn.execute(strSQLkopi2)
						
						
						oRec2.movenext
						wend
						oRec2.close
					
					
		end if
		
		oRec.close	
		Response.redirect "sdsk_prio_grp.asp?menu=tok&shokselector=1"
		
		
		
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	gemtider = 1
	
	else
	strSQL = "SELECT navn, editor, dato, gemtider FROM sdsk_prio_grp WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	gemtider = oRec("gemtider")
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
                        <h5 class="modal-title">Skjul responstider</h5>
                    </div>
                    <div class="modal-body">
                        Skjul responstider for kontakter, der er medlem af denne aftalegruppe, på Incidents listen.
                    </div>
                </div>
            </div>
        </div>


        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u>Aftalegrupper - <%=varbroedkrumme%></u></h3>
                <div class="portlet-body">
                    <%
	                call  sdsktopmenu()
	                %>

                    <form action="sdsk_prio_grp.asp?menu=tok&func=<%=dbfunc%>" method="post">
                    <div class="row">
                        <div class="col-lg-10">&nbsp</div>
                        <div class="col-lg-2 pad-b10">
                            <button type="submit" class="btn btn-success btn-sm pull-right"><b>Opdatér</b></button>
                        </div>
                    </div> 
                    <input type="hidden" name="id" value="<%=id%>">
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
                            <div class="col-lg-2">
                                <%
	                            if cint(gemtider) = 1 then
	                            gemTiderCHK = "CHECKED"
	                            else
	                            gemTiderCHK = ""
	                            end if
	                            %>
                                <input type="checkbox" name="FM_gemtider" id="FM_gemtider" value="1" <%=gemTiderCHK%>>
                                Skjul responstider <a data-toggle="modal" href="#styledhelpbox_1"><span class="fa fa-info-circle"></span></a>
                            </div>
                        </div>

                    </div>
                    </form>

                    <br /><br />
                    <div class="row">
                        <div class="col-lg-2">
                            <a href="Javascript:history.back()" class="btn btn-default btn-sm"><b>Tilbage</b></a>
                        </div>
                    </div>



                    <%if func = "red" then %>
                    <br /><br /><br /><br /><div style="font-weight: lighter;">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></div>
                    <%end if %>


                  

                </div>
            </div>
        </div>

        <%case else %>

            <div class="container">
                <div class="portlet">
                    <h3 class="portlet-title"><u>Aftalegrupper</u></h3>
                    <div class="portlet-body">

                        <%
	                        call  sdsktopmenu()
	                    %>

                        <div class="row">
                            <div class="col-lg-10">&nbsp</div>
                            <div class="col-lg-2 pad-b10">
                                <a href="sdsk_prio_grp.asp?menu=tok&func=opret"" class="btn btn-success btn-sm pull-right"><b>Opret ny Aftalegruppe</b></a>
                            </div>
                        </div>

                        <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                            <thead>
                                <tr>
                                    <th>Id</th>
                                    <th><b>Aftalegruppe</b> (Antal kontakter)</th>
                                    <th></th>
                                    <th></th>
                                    <th></th>
                                </tr>
                            </thead>

                            <tbody>

                                <%
	                            strSQL = "SELECT s.id, s.navn, COUNT(k.kid) AS antal FROM sdsk_prio_grp s "_
	                            &" LEFT JOIN kunder k ON (k.sdskpriogrp = s.id) WHERE s.id <> 0 GROUP BY s.id ORDER BY s.navn"
	
	                            'Response.write strSQL
	                            'Response.flush
	
	                            oRec.open strSQL, oConn, 3
	                            while not oRec.EOF 
	                            %>
                                <tr>
                                    <td><%=oRec("id")%></td>
                                    <td><a href="sdsk_prio_grp.asp?menu=tok&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%> </a>
		                            <%if oRec("id") = 1 then%>
		                            &nbsp;- Standard gruppe. 
		                            <%end if%>
		
		                            &nbsp;&nbsp;<b>(<%=oRec("antal")%>)</b>
		                            </td>
                                    <td><a href="sdsk_prioitet.asp?FM_prio_grp=<%=oRec("id")%>" class=rmenu>Se grp.&nbsp;<img src="../ill/soeg-knap.gif" width="16" height="16" alt="" border="0"></a></td>
		                            <td><a href="sdsk_prio_grp.asp?func=kopier&id=<%=oRec("id")%>" class=rmenu><img src="../ill/ac0038-16_kopier.gif" width="20" height="16" alt="" border="0"></a></td>
                                    <td>
		                            <%if oRec("id") <> 1 AND cint(oRec("antal")) = 0 then%>
		                            <a href="sdsk_prio_grp.asp?menu=tok&func=slet&id=<%=oRec("id")%>"><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></a>
                                    <%end if %>
                                    </td>
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
                                <a href="Javascript:window.close()" class="btn btn-default btn-sm"><b>Luk vindue</b></a>
                            </div>
                        </div>

                    </div>
                </div>
            </div>

















            
       <%end select %>

    </div>
</div> 
        
    

<!--#include file="../inc/regular/footer_inc.asp"-->