

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
        <div class="col-lg-2"><b>Timeout Kontrolpanel - Prioitetsklasser</b><br />
            Tilføj, fjern eller rediger Prioitetsklasser.
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
    thisfile = "sdsk_kontrolpanel"
	kview = "j"
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	select case func
	case "slet"

	'call sdsktopmenu()

    '*** Her spørges om det er ok at der slettes en medarbejder ***

	    oskrift = "Prioitet" 
        slttxta = "Du er ved at <b>SLETTE</b> en Prioitet. Er dette korrekt?"
        slttxtb = "" 
        slturl = "sdsk_prio_typ.asp?menu=tok&func=sletok&id="& id

        call sletcnf_2015(oskrift, slttxta, slttxtb, slturl)
	
	case "sletok"
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM sdsk_prio_typ WHERE id = "& id &"")
	oConn.execute("DELETE FROM sdsk_prioitet WHERE priogrp = "& id &"")
	Response.redirect "sdsk_prio_typ.asp?menu=tok&shokselector=1"
    
    case "dbopr", "dbred"
	'*** Her indsættes en ny type i db ****
		if len(request("FM_navn")) = 0 then
		%>
		<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
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
		
		if func = "dbopr" then
		oConn.execute("INSERT INTO sdsk_prio_typ (navn, editor, dato) VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"')")
		else
		oConn.execute("UPDATE sdsk_prio_typ SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"' WHERE id = "&id&"")
		end if
		
		Response.redirect "sdsk_prio_typ.asp?menu=tok&shokselector=1"
		end if
		
	case "kopier"
		
		
		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
		
		strSQL = "SELECT navn, id FROM sdsk_prio_typ WHERE id = "& id
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then
						
						
						'*** Kopierer grp ***
						strSQLkopi = "INSERT INTO sdsk_prio_typ (editor, dato, navn) VALUES "_
						&" ('"& strEditor &"', '"& strDato &"', '"& oRec("navn") &" (kopi)')"
						oConn.execute(strSQLkopi)
						
						strSQL3 = "SELECT id FROM sdsk_prio_typ WHERE id <> 0 ORDER BY id DESC"
						oRec3.open strSQL3, oConn, 3
						if not oRec3.EOF then
						
						nygrpId = oRec3("id")
						
						end if
						oRec3.close
						
						'** Henter prioiteter ***
						strSQL2 = "SELECT navn, responstid, faktor, kunweekend, advisering FROM sdsk_prioitet WHERE priogrp =" & oRec("id")
						oRec2.open strSQL2, oConn, 3
						
						while not oRec2.EOF 
						
						strNavn = oRec2("navn")
						rsptid = replace(oRec2("responstid"), ",", ".")
						dblfaktor = replace(oRec2("faktor"), ",", ".")
						intkunweekend = oRec2("kunweekend")
						dblAdvi = replace(oRec2("advisering"), ",", ".")
						
							
								strSQLkopi2 = "INSERT INTO sdsk_prioitet ( editor, dato, navn, responstid, faktor, "_
								&" kunweekend, advisering, priogrp) "_
								&" VALUES ('"& strEditor &"', '"& strDato &"', '"& strNavn &"', "& rsptid &", "& dblfaktor &", "_
								&" "& intkunweekend &", "& dblAdvi &", "& nygrpId &")"
								oConn.execute(strSQLkopi2)
						
						
						oRec2.movenext
						wend
						oRec2.close
					
					
		end if
		
		oRec.close	
		Response.redirect "sdsk_prio_typ.asp?menu=tok&shokselector=1"
		
		
		
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	
	else
	strSQL = "SELECT navn, editor, dato FROM sdsk_prio_typ WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "opdaterpil" 
	end if
	%> 

        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u>Prioitet  - <%=varbroedkrumme%></u></h3>

                <div class="portlet-body">
                    <%
	                    call  sdsktopmenu()
	                %>
                    <form action="sdsk_prio_typ.asp?menu=tok&func=<%=dbfunc%>" method="post">
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
                        </div>
                    </div>
                    </form>

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

        <%case else %>

        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u>Prioiteter</u></h3>
                <div class="portlet-body">

                    <%
	                    call  sdsktopmenu()
	                %>

                    <div class="row">
                        <div class="col-lg-10">&nbsp</div>
                        <div class="col-lg-2 pad-b10">
                            <a href="sdsk_prio_typ.asp?menu=tok&func=opret " class="btn btn-success btn-sm pull-right"><b>Opret ny Prioitet</b></a>
                        </div>
                    </div>

                    <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                        <thead>
                            <tr>
                                <th>Id</th>
                                <th><b>Prioitet</b> (Antal Incidents)</th>
                                <th></th>
                            </tr>
                        </thead>
                        <tbody>
                            <%
	                        strSQL = "SELECT s.id, s.navn, COUNT(p.id) AS antal FROM sdsk_prio_typ s "_
	                        &" LEFT JOIN sdsk_prioitet p ON (p.type = s.id) WHERE s.id <> 0 GROUP BY s.id ORDER BY s.navn"
	
	                        'Response.write strSQL
	                        'Response.flush
	
	                        oRec.open strSQL, oConn, 3
	                        while not oRec.EOF 
	                        %>
                            <tr>
                                <td><%=oRec("id")%></td>
                                <td><a href="sdsk_prio_typ.asp?menu=tok&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%> </a>
		                        <%if oRec("id") = 1 then%>
		                        &nbsp;- Standard Priotetsklasse. 
		                        <%end if%>
		
		                        &nbsp;&nbsp;<b>(<%=oRec("antal")%>)</b>
		                        </td>
                                <td>
		                        <%if oRec("id") <> 1 AND cint(oRec("antal")) = 0 then%>
		                        <a href="sdsk_prio_typ.asp?menu=tok&func=slet&id=<%=oRec("id")%>"><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></a>
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

                    

                </div>
            </div>
        </div>


        <%end select %>
    </div>
</div>

 
        
    

<!--#include file="../inc/regular/footer_inc.asp"-->