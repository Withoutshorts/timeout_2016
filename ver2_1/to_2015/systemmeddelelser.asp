

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->


<div class="wrapper">
<div class="content">

    <%
    if len(session("user")) = 0 then
	errortype = 5
	call showError(errortype)
    response.End
    end if

    if lto <> "outz" then
        %>
        <div class="container">
            <div class="portlet-body">
                <div class="row" style="text-align:center;">
                    <div class="col-lg-12">
                        <h2 style="color:darkred">Access Denied!</h2>
                        <br />
                        <h4 style="color:darkred">Du har ikke adgang til denne side</h4>
                   </div>      
                </div>
            </div>
        </div>
        <%
        response.End
    end if

    call menu_2014

    func = request("func")
    select case func 
        case "dbopr"
            'Opretter hvis man har tastet ny
            newOverskrift = request("FM_newoverskrift") 
            newBesked = request("FM_newbesked")

            if newBesked <> "" then
            
                if len(trim(request("FM_newfra"))) <> 0 AND isdate(request("FM_newfra")) then
                    newfra = year(request("FM_newfra")) &"-"& month(request("FM_newfra")) &"-"& day(request("FM_newfra"))
                end if

                if len(trim(request("FM_newtil"))) <> 0 AND isdate(request("FM_newfra")) then
                    newtil = year(request("FM_newtil")) &"-"& month(request("FM_newtil")) &"-"& day(request("FM_newtil"))
                end if

                editor = session("mid")
                editdate = year(now) &"-"& month(now) &"-"& day(now)

                strConnect_admin = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154;Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_admin;" 
                Set oConn_admin = Server.CreateObject("ADODB.Connection")
                Set oRec_admin = Server.CreateObject ("ADODB.Recordset")

                oConn_admin.open strConnect_admin

                'response.Write "< <br><br><br><br><br> ---------- DATOER " & newfra & " --------- " & newtil

                oConn_admin.Execute("INSERT INTO systemmeddelelser (editor, editdate, overskrift, besked, visfra, vistil) VALUES ('"& editor &"', '"& editdate &"', '"& newOverskrift &"', '"& newBesked &"', '"& newfra &"', '"& newtil &"')")
     

                oConn_admin.close
            end if

            response.Redirect "systemmeddelelser.asp"

        case "sletok"
            id = request("id")    

            if len(trim(request("id"))) <> 0 then

                strConnect_admin = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154;Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_admin;" 
                Set oConn_admin = Server.CreateObject("ADODB.Connection")
                Set oRec_admin = Server.CreateObject ("ADODB.Recordset")

                oConn_admin.open strConnect_admin

                oConn_admin.Execute("DELETE FROM systemmeddelelser WHERE id = "& id)

                oConn_admin.close
            end if

            response.Redirect "systemmeddelelser.asp"

    end select
    %>
        
    <script src="js/datepicker.js" type="text/javascript"></script>
    <div class="container">
        <div class="portlet">

            <h3 class="portlet-title"><u>System meddelelser</u></h3>

            <div class="portlet-body">

                <form action="systemmeddelelser.asp?func=dbopr" method="post">

                    <div class="row">
                        <div class="col-lg-12">
                            <button type="submit" class="btn btn-success btn-sm pull-right"><b>Opdatér</b></button>
                        </div>
                    </div>

                    <div class="row">
                        <div class="col-lg-12">
                            <table class="table table-striped table-bordered">
                                <thead>
                                    <tr>
                                        <th>Overskrift</th>
                                        <th>Besked</th>
                                        <th>Vis fra</th>
                                        <th>Vis til</th>
                                        <th></th>
                                    </tr>
                                </thead>

                                <tbody>

                                    <tr>
                                        <td><input name="FM_newoverskrift" type="text" class="form-control input-small" /></td>
                                        <td><input name="FM_newbesked" type="text" class="form-control input-small" /></td>
                                        <td>
                                            <div class='input-group date'>
                                                  <input name="FM_newfra" type="text" autocomplete="off" class="form-control input-small" value="" placeholder="dd-mm-yyyy" />
                                                    <span class="input-group-addon input-small">
                                                    <span class="fa fa-calendar">
                                                    </span>
                                                </span>
                                          </div>
                                        </td>
                                        <td>
                                            <div class='input-group date'>
                                                  <input name="FM_newtil" type="text" autocomplete="off" class="form-control input-small" value="" placeholder="dd-mm-yyyy" />
                                                    <span class="input-group-addon input-small">
                                                    <span class="fa fa-calendar">
                                                    </span>
                                                </span>
                                          </div>
                                        </td>
                                        <td></td>
                                    </tr>

                                    <%
                                    strConnect_admin = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154;Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_admin;" 
                                    Set oConn_admin = Server.CreateObject("ADODB.Connection")
                                    Set oRec_admin = Server.CreateObject ("ADODB.Recordset")

                                    oConn_admin.open strConnect_admin


                                    strSQL = "SELECT id, overskrift, besked, visfra, vistil FROM systemmeddelelser"
                                    oRec_admin.open strSQL, oConn_admin, 3
                                    while not oRec_admin.EOF


                                    %>
                                    <tr>
                                        <td><%=oRec_admin("overskrift") %></td>
                                        <td><%=oRec_admin("besked") %></td>
                                        <td><%=oRec_admin("visfra") %></td>
                                        <td><%=oRec_admin("vistil") %></td>
                                        <td style="text-align:center; width:5px;"><a href="systemmeddelelser.asp?func=sletok&id=<%=oRec_admin("id") %>"><span style="color:darkred;" class="fa fa-times"></span></a></td>
                                    </tr>
                                    <%
                                    oRec_admin.movenext
                                    wend
                                    oRec_admin.close

                                    oConn_admin.close
                                    %>
                                </tbody>
                            </table>
                        </div>
                    </div>
                </form>

            </div>
        </div>
    </div>



</div>
</div>


<!--#include file="../inc/regular/footer_inc.asp"-->