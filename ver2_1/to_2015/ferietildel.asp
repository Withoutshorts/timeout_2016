

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<%call menu_2014 %>

<div class="wrapper">
    <div class="content">


        <%

        if len(session("user")) = 0 then
	    errortype = 5
	    call showError(errortype)
        response.End
	    end if
        

        if len(trim(request("FM_medarb"))) <> 0 then
            FM_medarb = request("FM_medarb")
        else
            FM_medarb = 1
        end if

        if len(trim(request("grp_tpID"))) <> 0 then
        grp_tpID = request("grp_tpID")
        else
        grp_tpID = 0
        end if

        select case FM_medarb
            case 1

                chbMedarb1 = "checked"
                SELstatus1 = ""
                chbMedarb2 = ""
                SELstatus2 = "DISABLED"            
                if grp_tpID <> 0 then
                grp_typIDSQL = " WHERE medarbejdertype = "& grp_tpID
                else
                grp_typIDSQL = ""
                end if

            case 2

                chbMedarb1 = ""
                SELstatus1 = "DISABLED"
                chbMedarb2 = "checked"
                SELstatus2 = ""
                if grp_tpID <> 0 then
                grp_typIDSQL = " WHERE ProjektgruppeId = "& grp_tpID
                else
                grp_typIDSQL = ""
                end if

        end select


        
        

        'select case grp_tpID
         

        %>




        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u>Ferie tildeling</u></h3>

                <div class="portlet-body">

                    <form action="ferietildel.asp" method="post">
                    <div class="well">
                        <div class="row">
                            <div class="col-lg-12">
                                <h4 class="panel-title-well">Søgefilter</h4>
                            </div>
                        </div>
                        <div class="row">
                            <div class="col-lg-1">Ferie år:</div>
                            <div class="col-lg-2"></div>
                            <div class="col-lg-3">Medarbejdertype <input type="radio" name="FM_medarb" value="1" <%=chbMedarb1 %> /></div>
                            <div class="col-lg-3">Projektgrupper <input type="radio" name="FM_medarb" value="2" <%=chbMedarb2 %> /></div>                            
                        </div>
                        <div class="row">
                            <div class="col-lg-2">
                                <select class="form-control input-small" name="ferieaar">
                                    <option value="1">01-05-2017 - 04-30-2018</option>
                                    <option value="2">01-05-2018 - 04-30-2019</option>
                                    <option value="3">01-05-2019 - 04-30-2020</option>
                                    <option value="4">01-05-2020 - 04-30-2021</option>
                                </select>
                            </div>
                            <div class="col-lg-1"></div>
                            <div class="col-lg-3">
                                <select class="form-control input-small" name="grp_tpID" <%=SELstatus1 %>>
                                    <option value="0">Alle</option>

                                    <%
                                        if FM_medarb = 1 then

                                        strSQL = "SELECT id, type FROM medarbejdertyper WHERE id <> 1 ORDER BY type"
                                        oRec.open strSQL, oConn, 3
                                        while not oRec.EOF
                                        
                                        if cint(grp_tpID) = cint(oRec("id")) then
                                        SELOP = "SELECTED"
                                        else
                                        SELOP = ""
                                        end if

                                        %><option value="<%=oRec("id") %>" <%=SELOP %>><%=oRec("type") %></option><%

                                        oRec.movenext
                                        wend
                                        oRec.close
                                        

                                        end if
                                    %>

                                </select>
                            </div>
                            <div class="col-lg-3">
                                <select class="form-control input-small" name="grp_tpID" <%=SELstatus2 %>>

                                    <%
                                        if FM_medarb = 2 then

                                            strSQL = "SELECT id, navn FROM projektgrupper WHERE id <> 1 ORDER BY navn"
                                            oRec.open strSQL, oConn, 3
                                            while not oRec.EOF
                                        
                                            if cint(grp_tpID) = cint(oRec("id")) then
                                            SELOP = "SELECTED"
                                            else
                                            SELOP = ""
                                            end if

                                            %><option value="<%=oRec("id") %>" <%=SELOP %>><%=oRec("navn") %></option><%

                                            oRec.movenext
                                            wend
                                            oRec.close
                                        
                                        else

                                            %><option>Alle-gruppen (alle medarbejdere)</option><%

                                        end if
                                    %>

                                </select>
                            </div>

                            <div class="col-lg-2"><button type="submit" class="btn btn-secondary btn-sm pull-right"><b>Vis</b></button></div>
                        </div>
                       
                    </div>
                    </form>

                    <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                        <thead>
                            <tr>
                                <th>Navn</th>
                                <th>Ansat d.</th>
                                <th>Norm</th>
                                <th>Ferie saldo <br /><span style="font-size:9px;">1.5.2017 - 30.4.18</span></th>
                                <th>Optjent ferie <br /><span style="font-size:9px;">1.1.2017 - 31.12.2017</span></th>
                                <th>Korrigering overført</th>
                                <th>Ferie saldo <br /><span style="font-size:9px;">Næste år</span></th>
                            </tr>
                        </thead>

                        <tbody>

                            <%
                                'ferieafholdt = 0
                                'select case FM_medarb
                                'case 1
                                strSQL = "SELECT Mid, Mnavn, ansatdato, medarbejdertype FROM medarbejdere ORDER BY mnavn"
                                'case 2
                                'strSQL = "SELECT Mid, Mnavn, ansatdato FROM medarbejdere LEFT JOIN progrupperelationer pr ON (pr.MedarbejderId = Mid)"
                                'end select

                                strSQL = strSQL & grp_typIDSQL

                                oRec.open strSQL 
                                while not oRec.EOF

                                %>
                                    <tr>
                                        <td><%=oRec("Mnavn") %></td>
                                        <td><%=oRec("ansatdato") %></td>
                                        <td><input type="text" class="form-control input-small" style="width:50px;" /></td>
                                        <td><input type="text" class="form-control input-small" style="width:50px;" /></td>
                                        <td><input type="text" class="form-control input-small" style="width:50px;" /></td>
                                        <td><input type="text" class="form-control input-small" style="width:50px;" /></td>
                                        <td><input type="text" class="form-control input-small" style="width:50px;" /></td>
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



        













    </div>
</div>

<!--#include file="../inc/regular/footer_inc.asp"-->