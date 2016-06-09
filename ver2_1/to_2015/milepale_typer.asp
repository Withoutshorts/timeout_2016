<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->


<%call menu_2014 %>
<div class="wrapper">
    <div class="content">


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
	
	select case func
	case "slet"
	'*** Her spørges om det er ok at der slettes  ***

        oskrift = "Milepæl-type" 
        slttxta = "Du er ved at <b>SLETTE</b> en milepæl type. Er dette korrekt?"
        slttxtb = "" 
        slturl = "milepale_typer.asp?func=sletok&id="& id

        call sletcnf_2015(oskrift, slttxta, slttxtb, slturl)



	case "sletok"
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM milepale_typer WHERE id = "& id &"")
	Response.redirect "milepale_typer.asp?menu=tok&shokselector=1"

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

        if len(trim(request("FM_mpt_fak"))) <> 0 then
        mpt_fak = request("FM_mpt_fak")
        else
        mpt_fak = 0
        end if
		
		if func = "dbopr" then
		oConn.execute("INSERT INTO milepale_typer (navn, editor, dato, ikon, mpt_fak) VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', '"& strIkon &"', "& mpt_fak &")")
		else
		oConn.execute("UPDATE milepale_typer SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"', ikon = '"& strIkon &"', mpt_fak = "& mpt_fak &" WHERE id = "&id&"")
		end if
		
		Response.redirect "milepale_typer.asp?menu=tok&shokselector=1"
		end if
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "Opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	strIkon = "gron"
	
	else
	strSQL = "SELECT navn, editor, dato, ikon, mpt_fak FROM milepale_typer WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	strIkon = oRec("ikon")
    mpt_fak = oRec("mpt_fak")
	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "Opdaterpil" 
	end if
	%>
	 <!--Rediger-->

    <div class="container">
    <div class="portlet">
            <h3 class="portlet-title"><u>Milepæle-typer</u></h3>
            <form action="milepale_typer.asp?menu=tok&func=<%=dbfunc%>" method="post">
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
                    <div class="col-lg-2">Milepælnavn:</div>
                    <div class="col-lg-2"><input type="text" class="form-control input-small" name="FM_navn" value="<%=strNavn%>"></div>
                </div>

                <div class="row">
                    <div class="col-lg-1">&nbsp</div>
                     <%if cint(mpt_fak) = 1 then
                    mpt_fakCHK = "CHECKED"
                    else
                    mpt_fakCHK = ""
                    end if%>
                   <div class="col-lg-4"><input id="Checkbox1" name="FM_mpt_fak" value="1" type="checkbox" <%=mpt_fakCHK%> />&nbsp Vis denne type under ERP - "Til fakturering"</div>
                </div>

                </div>

            </div>
        <%if dbfunc = "dbred" then%>
        <br /><br /><br /><div style="font-weight: lighter;">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></div>
        <%end if %>
        </form>

    </div>            
    </div>


<%case else %>
<!--Liste-->
<script src="js/milepaleliste.js" type="text/javascript"></script>
    
        <div class="container">
        <div class="portlet">
            <h3 class="portlet-title"><u>Milepæle-typer</u></h3>
            <form action="milepale_typer.asp?menu=tok&func=opret" method="post">
            <section>
                         <div class="row">
                             <div class="col-lg-10">&nbsp;</div>
                             <div class="col-lg-2">
                                <button class="btn btn-sm btn-success pull-right"><b>Opret ny</b></button><br />&nbsp;
                            </div>
                 </div>
              </section>
             </form>

            <div class="portlet-body">

                <table id="milepaleliste" class="table dataTable table-striped table-bordered table-hover ui-datatable">
                    <thead>
                        <tr>
                            <th>id</th>
                            <th>Emne</th>
                            <th>Vis under "Til fakturering"</th>
                            <th style="width: 5%">Slet</th>
                        </tr>
                    </thead>
                    <tbody>
                        <%
	                    sort = Request("sort")
	                    'if sort = "navn" then
	                    strSQL = "SELECT id, navn, mpt_fak FROM milepale_typer ORDER BY navn"
	                    'else
	                    'strSQL = "SELECT id, navn, mpt_fak FROM milepale_typer ORDER BY id"
	                    'end if
	
                        c = 1

	                    oRec.open strSQL, oConn, 3
	                    while not oRec.EOF 

                        select case right(c, 1)
                        case 0,2,4,6,8
                        bgt = "#EFF3FF"
                        case else
                        bgt = "#FFFFFF"
                        end select

	                    %>
                        <tr>
                            <td><%=oRec("id")%></td>
                            <td><a href="milepale_typer.asp?menu=tok&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%> </a></td>
                            <td>
                            <%if cint(oRec("mpt_fak")) = 1 then
                            %>
                            Ja
                            <%
                            end if %>
                            </td>
                            <td><a href="milepale_typer.asp?menu=tok&func=slet&id=<%=oRec("id")%>"><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></a></td>
                        </tr>
                        <%
	                    x = 0
                        c = c + 1
	                    oRec.movenext
	                    wend
	                    %>	
                    </tbody>
                    <tfoot>
                        <tr>
                            <td>id</td>
                            <td>Emne</td>
                            <td>Vis under "Til fakturering"</td>
                            <td>Slet</td>
                        </tr>
                    </tfoot>
                </table>

            </div>
        </div>
      </div>











<%end select %>
    </div>
</div>

<!--#include file="../inc/regular/footer_inc.asp"-->