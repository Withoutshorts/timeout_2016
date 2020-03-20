

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<script src="js/medarbtyper_grp_jav.js" type="text/javascript"></script>

<div class="wrapper">
<div class="content">

<%
    if len(session("user")) = 0 then
        errortype = 5
        call showError(errortype)
        response.End
    end if
	
    call menu_2014()
	
    '** Faste filter kri ***'
    thisfile = "medarbtyper_grp"
	
    func = request("func")
    
    if func = "dbopr" then
    id = 0
    else
    id = request("id")
    end if



    select case func 
    case "slet"
    '*** Her spørges om det er ok at der slettes ***
	    oskrift = medarb_txt_137
        slttxta = medarb_txt_005 &" <b>"&medarb_txt_004&"</b> "&medarb_txt_138&""
        slttxtb = "" 
        slturl = "medarbtyper_grp.asp?menu=tok&func=sletok&id="& id

        call sletcnf_2015(oskrift, slttxta, slttxtb, slturl)

    case "sletok"
        oConn.execute("DELETE FROM medarbtyper_grp WHERE id = "& id &"")

	    Response.redirect "medarbtyper_grp.asp?menu=tok&shokselector=1"

    case "dbopr", "dbred"

        if len(request("FM_navn")) = 0 then
            errortype = 8
		    call showError(errortype) 
            response.End
        end if

        function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function

        strNavn = SQLBless(request("FM_navn"))
		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)

        if len(trim(request("FM_opencalc"))) <> 0 then
        opencalc = 0 '** Lukket MAKS 1
                
                   '*** Nulstiller
                   oConn.execute("UPDATE medarbtyper_grp SET opencalc = 1 WHERE id <> 0")
        else
        opencalc = 1 '** Åben (default)
        end if

        if func = "dbopr" then
		oConn.execute("INSERT INTO medarbtyper_grp (navn, editor, dato, opencalc) VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', "& opencalc &")")
		else
		oConn.execute("UPDATE medarbtyper_grp SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"', opencalc = "& opencalc &" WHERE id = "&id&"")
		end if

        Response.redirect "medarbtyper_grp.asp?menu=tok&shokselector=1"

    case "opret", "red"
    '*** Her indlæses form til rediger/oprettelse af ny type ***
        if func = "opret" then
	    strNavn = ""
	    strTimepris = ""
	    varSubVal = "opretpil" 
	    varbroedkrumme = "Opret ny"
	    dbfunc = "dbopr"
	    opencalc = 1

	    else
	    strSQL = "SELECT navn, editor, dato, opencalc FROM medarbtyper_grp WHERE id=" & id
	    oRec.open strSQL,oConn, 3
	
	    if not oRec.EOF then
	    strNavn = oRec("navn")
	    strDato = oRec("dato")
	    strEditor = oRec("editor")
	    opencalc = oRec("opencalc")
        end if
	    oRec.close
	
        if opencalc = 0 then
        opencalcCHK = "CHECKED"
        else
        opencalcCHK = ""
        end if

	    dbfunc = "dbred"
	    varbroedkrumme = "Rediger"
	    varSubVal = "opdaterpil" 
	    end if
%>


        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u><%=medarb_txt_137 %></u></h3>
                <div class="portlet-body">
                    <form action="medarbtyper_grp.asp?menu=tok&func=<%=dbfunc%>" method="post">
                        <input type="hidden" name="id" value="<%=id%>">

                        <div class="row">
                            <div class="col-lg-12 pad-b5"><button type="submit" class="btn btn-success btn-sm pull-right"><b><%=medarb_txt_020 %></b></button></div>
                        </div>

                        <div class="well">
                            <h4 class="panel-title-well"><%=medarb_txt_021 %></h4>

                            <div class="row">
                                <div class="col-lg-3"><%=medarb_txt_022 %>:</div>
                                <div class="col-lg-3"><%=medarb_txt_139 %>:</div>
                            </div>
                            <div class="row">
                                <div class="col-lg-3"><input type="text" class="form-control input-small" name="FM_navn" value="<%=strNavn%>" class="form-control input-small"></div>
                                <div class="col-lg-3"><input type="CHECKBOX" name="FM_opencalc" value="1" <%=opencalcCHK %>> (<%=medarb_txt_142 %>)</div>
                            </div>

                        </div>
                    </form>

                    <%if dbfunc = "dbred" then%> 

                        <br /><br /><br /><br />

		                <div class="row"> 
                            <div class="col-lg-12"><%=medarb_txt_094 %> <b><%=strDato%></b> <%=medarb_txt_095 %> <b><%=strEditor%></b></div>
		                </div>               
	                <%end if%>

                </div>


            </div>
        </div>

    <%case else 'Liste %>
    
        
        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u><%=medarb_txt_137 %></u></h3>

                <div class="portlet-body">
                        <div class="row">
                            <div class="col-lg-12"><a href="medarbtyper_grp.asp?menu=tok&func=opret" class="btn btn-success btn-sm pull-right"><b><%=medarb_txt_122 %></b></a></div>
                        </div>

                    <table id="liste" class="table dataTable table-striped table-bordered table-hover ui-datatable">
                        <thead>
                            <tr>
                                <th><%=medarb_txt_143 %></th>
                                <th><%=medarb_txt_144 %></th>
                                <th><%=medarb_txt_140 %></th>
                                <th><%=medarb_txt_141 %></th>
                                <th></th>
                            </tr>
                        </thead>

                        <tbody>
                            <%
                                strSQL = "SELECT mt.id AS id, navn, COUNT(m.id) AS antal, opencalc FROM medarbtyper_grp AS mt "

	                            strSQL = strSQL & " LEFT JOIN medarbejdertyper AS m ON (m.mgruppe = mt.id) GROUP BY mt.id "

                                if sort = "navn" then
	                            strSQL = strSQL & "ORDER BY navn"
	                            else
	                            strSQL = strSQL & "ORDER BY m.id"
	                            end if

                                'Response.Write strSQL
                                'Response.flush

                                oRec.open strSQL, oConn, 3
                                while not oRec.EOF
                            %>
                                <tr>
                                    <td><%=oRec("id")%></td>
                                    <td><a href="medarbtyper_grp.asp?menu=tok&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%> </a></td>
                                    <td><%=oRec("antal")%></td>
                                    <td>
                                        <%
                                        if cint(oRec("opencalc")) = 1 then
                                            response.Write medarb_txt_007
                                        else
                                            response.Write medarb_txt_008
                                        end if
                                        %>
                                    </td>
                                    <td style="text-align:center;">
                                        <%if cint(oRec("antal")) = 0 AND oRec("id") <> 1 then %>
                                            <a href="medarbtyper_grp.asp?menu=tok&func=slet&id=<%=oRec("id")%>"><span style="color:darkred;" class="fa fa-times"></span></a>
                                        <%end if %>
                                    </td>
                                </tr>
                            <%
                                oRec.movenext
                                wend
                                oRec.close
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