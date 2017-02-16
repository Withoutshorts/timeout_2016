

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
        <div class="col-lg-2"><b>Timeout Kontrolpanel - Incident Kategori</b><br />
            Tilføj, fjern eller rediger Incident Kategori.
        </div>
    </div>
<%
end sub


    if len(session("user")) = 0 then
    %>
    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
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

    thisfile = "sdsk_kontrolpanel"
	kview = "j"

    call menu_2014 %> 
        
<div class="wrapper">
    <div class="content">


<%
select case func 
case "slet"

    'call sdsktopmenu()

    '*** Her spørges om det er ok at der slettes en medarbejder ***
	    oskrift = "Incident - Kategori" 
        slttxta = "Du er ved at <b>SLETTE</b> en Incident Kategori. Er dette korrekt?"
        slttxtb = "" 
        slturl = "sdsk_typer.asp?menu=tok&func=sletok&id="& id

        call sletcnf_2015(oskrift, slttxta, slttxtb, slturl)

case "sletok"
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM sdsk_typer WHERE id = "& id &"")
	Response.redirect "sdsk_typer.asp?menu=tok&shokselector=1"



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
		oConn.execute("INSERT INTO sdsk_typer (navn, editor, dato) VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"')")
		else
		oConn.execute("UPDATE sdsk_typer SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"' WHERE id = "&id&"")
		end if
		
		Response.redirect "sdsk_typer.asp?menu=tok&shokselector=1"
		end if
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	
	else
	strSQL = "SELECT navn, editor, dato FROM sdsk_typer WHERE id=" & id
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
                <h3 class="portlet-title"><u>Incident Kategori - <%=varbroedkrumme%></u></h3>
                <div class="portlet-body">
                    <%
	                    call  sdsktopmenu()
	                %>
                    <form action="sdsk_typer.asp?menu=tok&func=<%=dbfunc%>" method="post">
                        <input type="hidden" name="id" value="<%=id%>">

                        <div class="row">
                            <div class="col-lg-10">&nbsp</div>
                            <div class="col-lg-2 pad-b10">
                                <button type="submit" class="btn btn-success btn-sm pull-right"><b>Opdatér</b></button>
                            </div>
                        </div>

                        <div class="well well-white">
                            <div class="row">
                            <div class="col-lg-1">&nbsp</div>
                            <div class="col-lg-2">Navn:</div>
                            <div class="col-lg-3"><input type="text" name="FM_navn" value="<%=strNavn%>" class="form-control input-small" /></div>
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

        <%case else %>

        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u>Incident kategorier</u></h3>
                <div class="portlet-body">

                    <%
	                call  sdsktopmenu()
	                %>

                    <div class="row">
                        <div class="col-lg-10">&nbsp</div>
                        <div class="col-lg-2 pad-b10">
                            <a href="sdsk_typer.asp?menu=tok&func=opret" class="btn btn-success btn-sm pull-right"><b>Opret ny Incident kategori</b></a>
                        </div>
                    </div>

                    <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                        <thead>
                            <tr>
                                <th>Id</th>
                                <th>Kategori</th>
                                <th></th>
                            </tr>
                        </thead>
                        <tbody>

                            <%
	                        'sort = Request("sort")
	                        'if sort = "navn" then
	                        strSQL = "SELECT id, navn FROM sdsk_typer ORDER BY navn"
	                        'else
	                        'strSQL = "SELECT id, navn FROM sdsk_typer ORDER BY id"
	                        'end if
	
	                        oRec.open strSQL, oConn, 3
	                        while not oRec.EOF 
	                        %>

                            <tr>
                                <td><%=oRec("id")%></td>
		                        <td><a href="sdsk_typer.asp?menu=tok&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%> </a></td>
		                        <td><a href="sdsk_typer.asp?menu=tok&func=slet&id=<%=oRec("id")%>"><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></a></td>
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