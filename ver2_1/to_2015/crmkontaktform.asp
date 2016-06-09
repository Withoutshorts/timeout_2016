

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--include file="inc/budget_firapport_inc.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<%call menu_2014 %>
<div class="wrapper">
    <div class="content">


<%
    

if len(session("user")) = 0 then
	
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
	
	select case func
	case "slet"
	'*** Her spørges om det er ok at der slettes en medarbejder ***
        oskrift = "CRM - Kontaktform" 
        slttxta = "Du er ved at <b>SLETTE</b> en kontaktform. Er dette korrekt?"
        slttxtb = "" 
        slturl = "crmkontaktform.asp?func=sletok&id="& id

        call sletcnf_2015(oskrift, slttxta, slttxtb, slturl)


    case "sletok"
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM crmkontaktform WHERE id = "& id &"")
	Response.redirect "crmkontaktform.asp?menu=tok&shokselector=1"
	
	case "dbopr", "dbred"
	'*** Her indsættes en ny type i db ****
		if len(request("FM_navn")) = 0 then
		
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
		strIkon = request("FM_ikon")
		
		if func = "dbopr" then
		oConn.execute("INSERT INTO crmkontaktform (navn, editor, dato, ikon) VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', '"& strIkon &"')")
		else
		oConn.execute("UPDATE crmkontaktform SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"', ikon = '"& strIkon &"' WHERE id = "&id&"")
		end if
		
		Response.redirect "crmkontaktform.asp?menu=tok&shokselector=1"
		end if
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "Opret" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	
	else
	strSQL = "SELECT navn, editor, dato, ikon FROM crmkontaktform WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	strIkon = oRec("ikon")
	end if
	oRec.close
	
	select case strIkon
	case "crm_ikon_brev.gif"
	iKontext = "Brev"
	case "crm_ikon_email.gif"
	iKontext = "Email"
	case "crm_ikon_tlf.gif"
	iKontext = "Telefon"
	case else
	iKontext  = "Ingen"
	end select
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "Opdater" 
	end if
	%>

    <!--Rediger kontaktform-->
        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u>CRM - Kontaktform</u></h3>
                <form action="crmkontaktform.asp?menu=tok&func=<%=dbfunc%>" method="post">
	             <input type="hidden" name="id" value="<%=id%>">
    
                    <div class="row">
                    <div class="col-lg-10">&nbsp</div>
                    <div class="col-lg-2 pad-b10">
                    <button type="submit" class="btn btn-success btn-sm pull-right"><b>Opdatér</b></button>
                    </div>
                    </div>

                <div class="portlet-body">
                    <div class="well well-white">
                        <div class="row">
                        <div class="col-lg-1">&nbsp</div>
                        <div class="col-lg-2">Navn:</div>
                        <div class="col-lg-4"><input type="text" class="form-control input-small" name="FM_navn" value="<%=strNavn%>" size="30" ></div>
                        </div>  
                    
                        <div class="row">
                            <div class="col-lg-1">&nbsp</div>
                            <div class="col-lg-2">Ikon:</div>
                            <div class="col-lg-4">
                                <select class="form-control input-small" name="FM_ikon">
		                        <%if func = "red" then%>
		                        <option value="<%=strIkon%>" selected><%=iKontext%></option>
		                        <%end if%>
	                                <option value="crm_ikon_brev.gif">Brev</option>
	                                <option value="crm_ikon_email.gif">Email</option>
	                                <option value="crm_ikon_tlf.gif">Telefon</option>
	                                <option value="crm_ikon_mode.gif">Møde</option>
	                                <option value="blank.gif">Ingen</option>
	                        </select>
                            </div>
                        </div>
                                 
                    </div>

                     <%if dbfunc = "dbred" then%>
                    <br /><br /><br /><div style="font-weight: lighter;">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></div>
                    <%end if %>

                </div>
                   
                </form>
            </div>
        </div>


<%case else %>



<!--Liste-->
        <script src="js/crmstatus_liste.js" type="text/javascript"></script>
    <div class="container">
        <div class="portlet">
            <div class="portlet-body">
                <h3 class="portlet-title"><u>CRM - Kontaktform</u></h3>
                <form action="crmkontaktform.asp?menu=tok&func=opret" method="post">
                    <section>
                             <div class="row">
                                 <div class="col-lg-10">&nbsp;</div>
                                 <div class="col-lg-2">
                                <button class="btn btn-sm btn-success pull-right"><b>Opret ny</b></button><br />&nbsp;
                                </div>
                            </div>
                    </section>
                 </form>

                <table id="crm_liste" class="table dataTable table-striped table-bordered table-hover ui-datatable">
                    <thead>
                        <tr>
                            <th style="width: 10%">id</th>
                            <th style="width: 40%">Kontaktform</th>
                            <th style="width: 5%">Slet</th>
                        </tr>
                    </thead>
                    <tbody>
                        <%
	                    strSQL = "SELECT id, navn FROM crmkontaktform ORDER BY navn"
	                    
	                    oRec.open strSQL, oConn, 3
	                    while not oRec.EOF 
	                    %>
                        <tr>
                            <td><%=oRec("id")%></td>
                            <td><a href="crmkontaktform.asp?menu=tok&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%> </a></td>
                            <td><a href="crmkontaktform.asp?menu=tok&func=slet&id=<%=oRec("id")%>"><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></td>
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