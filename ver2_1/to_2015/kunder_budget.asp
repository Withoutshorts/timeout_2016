

<!--#include file="../inc/connection/conn_db_inc.asp"-->

<%
'** Jquery section 
if Request.Form("AjaxUpdateField") = "true" then
    Select Case Request.Form("control")

        case "FN_findkunder"
        
            sogtxt = request("sogtxt")

            strOPT = ""
            strSQL = "SELECT kid, kkundenavn, kkundenr FROM kunder WHERE kkundenavn LIKE '%"&sogtxt&"%'"
            'response.Write strSQL
            'response.end

            k = 0
            oRec.open strSQL, oConn, 3
            while not oRec.EOF

                strOPT = strOPT & "<option value='"&oRec("kid")&"'>"&oRec("kkundenavn")&" ("&oRec("kkundenr")&")</option>"

            k = k + 1
            oRec.movenext
            wend
            oRec.close

            if cint(k) = 0 then
                strOPT = "<option DISABLED>"&kunder_txt_120&"</option>"
            end if

            response.Write strOPT


        case "FN_getkundenavn"
            
            kundeSEL = request("kundeSEL")      
    
            strkundenavn = ""

            strSQL = "SELECT kkundenavn, kkundenr FROM kunder WHERE kid = "& kundeSEL
            oRec.open strSQL, oConn, 3
            if not oRec.EOF then
                strkundenavn = oRec("kkundenavn") & " (" & oRec("kkundenr") & ")"    
            end if
            oRec.close

            response.Write strkundenavn 

        case "FN_getkundedata"

            startaar = request("startaar")
            slutaar = request("slutaar")

            kundeSEL = request("kundeSEL") 
            kundeid = 0

            strSQL = "SELECT kid, kkundenavn FROM kunder WHERE kid = "& kundeSEL            
            
            oRec.open strSQL, oConn, 3
            if not oRec.EOF then
                strkundenavn = oRec("kkundenavn")
                kundeid = oRec("kid")
            end if
            oRec.close

            strTD = "<tr>"
            strTD = strTD & "<td><input name='FM_kundeid' type='hidden' value='"&kundeid&"' /> "&strkundenavn&" <input type='hidden' id='kundefindes_"&kundeid&"' value='1' /></td>"
            a = startaar
            for a = startaar TO slutaar

                sales_goal = 0
                budgetid = 0
                strSQL = "SELECT id, salesgoal FROM budget_kunder WHERE kundeid = "& kundeSEL &" AND date_year = "& a
                'response.Write "<tr><td>"&strSQL&"</td></tr>"
                'response.End
                oRec.open strSQL, oConn, 3
                if not oRec.EOF then
                    sales_goal = oRec("salesgoal")
                    budgetid = oRec("id")
                end if
                oRec.close

                strTD = strTD & "<td style='text-align:right;'><input type='text' name='FM_salesgoal_"&kundeid&"_"&a&"' id='FM_salesgoal_"&kundeid&"' value='"&sales_goal&"' style='display:none; text-align:right' class='form-control input-small inputfield' /><span class='txtfield'>"&formatnumber(sales_goal, 0)&"</span></td>"

            next
            strTD = strTD & "</tr>"

            response.Write strTD

    end select
Response.end
end if
%>



<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../timereg/inc/isint_func.asp"-->

<div class="wrapper">
<div class="content">

<%
    
    if len(session("user")) = 0 then
	errortype = 5
	call showError(errortype)
	response.end
	end if

    call menu_2014()

    if len(trim(request("FM_startaar"))) <> 0 then
        startaar = request("FM_startaar")
        response.cookies("2015")("kundebudget_startaar") = startaar
    else

        if request.cookies("2015")("kundebudget_startaar") <> "" then
            startaar = request.cookies("2015")("kundebudget_startaar")
        else
            startaar = (year(now)) - 1
        end if
    end if

    if len(trim(request("FM_slutaar"))) then
        slutaar = request("FM_slutaar")
        response.cookies("2015")("kundebudget_slutaar") = slutaar
    else
        if request.cookies("2015")("kundebudget_slutaar") <> "" then
            slutaar = request.cookies("2015")("kundebudget_slutaar")
        else
            slutaar = (year(now)) + 6
        end if
    end if

    arrdiff = slutaar - startaar

    arrdiff = arrdiff + 1

    'response.write "---------------------------------------------    startaar" & startaar & "   slutaar" & slutaar
    'response.Write "---- herher arrdiff " & arrdiff


    func = request("func")


    select case func
    
    case "save"

        kundeids = split(request("FM_kundeid"), ",")
        for i = 0 TO UBOUND (kundeids)
            
            kundeids(i) = replace(kundeids(i), " ", "")

            yr = startaar
            for yr = startaar TO slutaar
            
                if len(trim(request("FM_salesgoal_"&kundeids(i)&"_"&yr))) <> 0 then
                    salesgoal = request("FM_salesgoal_"&kundeids(i)&"_"&yr)

                    salesgoal = replace(salesgoal, ".", "")
                    salesgoal = replace(salesgoal, ",", ".")

                    budget_salesgoal = salesgoal

                    isInt = 0
			        call erDetInt(salesgoal) 
			        if isInt > 0 then

                                  
                            errortype = 225
				            call showError(errortype)
			
			                response.End
                
                    end if
                else
                    salesgoal = 0
                end if

                findes = 0
                strSQL = "SELECT id FROM budget_kunder WHERE kundeid = "& kundeids(i) & " AND date_year = "& yr
                oRec.open strSQL, oConn, 3
                if not oRec.EOF then
                    findes = 1
                end if
                oRec.close

                'response.Write "<br> --------------------------------------- KUNDE " & kundeids(i) & " ÅR " & a & " SALES GOAL " & salesgoal & " BUDGET FINDES " & findes & " her " & "FM_salesgoal_"&kundeids(i)&"_"&yr

                if cint(findes) = 1 then
                    strExecute = "UPDATE budget_kunder SET salesgoal = "& salesgoal & " WHERE kundeid = "& kundeids(i) &" AND date_year = "& yr
                else
                    strExecute = "INSERT INTO budget_kunder SET kundeid = "& kundeids(i) & ", date_year ="& yr & ", salesgoal = "& salesgoal
                end if

                'response.Write "<br> -------------- " & strExecute

                oConn.execute(strExecute)

            next           


        next

        response.Redirect "kunder_budget.asp"


    case else ' Liste 
%>

    <script src="js/kunder_budget_jav1.js" type="text/javascript"></script>    

    <input type="hidden" id="startaar" value="<%=startaar %>" />
    <input type="hidden" id="slutaar" value="<%=slutaar %>" />

    <div class="container">
        <div class="portlet">

            <h3 class="portlet-title"><u><%=kunder_txt_119 %></u></h3>

            <div class="portlet-body">

                <form action="kunder_budget.asp" method="post">
                    <div class="well">

                        <div class="row">
                            <div class="col-lg-12">
                                <h4 class="panel-title-well"><%=kunder_txt_082 %></h4>
                            </div>
                        </div>

                        <div class="row">
                            <div class="col-lg-5"><%=kunder_txt_121 %>:</div>
                            <div class="col-lg-2"></div>
                            <div class="col-lg-2"><%=kunder_txt_122 %>:</div>
                            <div class="col-lg-2"><%=kunder_txt_123 %>:</div>
                           <!-- <div class="col-lg-4">Eller vælg kunder her:</div> -->
                        </div>

                        <div class="row">
                            <div class="col-lg-5">
                                <input type="hidden" id="valgt_kunde" />
                                <input type="text" id="sog_kunde" class="form-control input-small" placeholder="<%=kunder_txt_124 %>" autocomplete="off" /> 

                                <select class="form-control input-small" id="kunde_DD" style="display:none;" size="10">
                                    <option value="0"><%=kunder_txt_120 %></option>
                                </select>
                            </div>
                            <div class="col-lg-1"><a class="btn btn-secondary btn-sm" id="kunde_ADD"><b><%=kunder_txt_125 %></b></a></div>

                            <div class="col-lg-1"></div>

                            <div class="col-lg-2">
                                <select class="form-control input-small" name="FM_startaar">
                                    <%
                                    i = (year(now)) - 3
                                    for i = (year(now)) - 3 TO (year(now)) + 6
                                        if cint(i) = cint(startaar) then
                                            SEL = "SELECTED"
                                        else
                                            SEL = ""
                                        end if
                                        response.Write "<option value='"&i&"' "&SEL&" >"&i&"</option>"
                                    next
                                    %>
                                </select>
                            </div>

                            <div class="col-lg-2">
                                <select class="form-control input-small" name="FM_slutaar">
                                    <%
                                    i = (year(now)) - 3
                                    for i = (year(now)) - 3 TO (year(now)) + 6

                                        if cint(i) = cint(slutaar) then
                                            SEL = "SELECTED"
                                        else
                                            SEL = ""
                                        end if

                                        response.Write "<option value='"&i&"' "&SEL&" >"&i&"</option>"
                                    next
                                    %>
                                </select>
                            </div>

                            <div class="col-lg-1">
                                <button type="submit" class="btn btn-secondary btn-sm pull-right"><b><%=kunder_txt_126 %></b></button>
                            </div>

                        </div>     
                    
                        <div class="row">
                            <div class="col-lg-12"><span id="errormessage" style="display:none; color:darkred;"><%=kunder_txt_127 %></span></div>
                        </div>

                    </div>

                </form>


                <!---------- Listen ----------->
                <form action="kunder_budget.asp?func=save" method="post">

                    <input type="hidden" name="FM_startaar" value="<%=startaar %>" />
                    <input type="hidden" name="FM_slutaar" value="<%=slutaar %>" />

                    <div class="row">
                        <div class="col-lg-2"><span style="font-size:150%; cursor:pointer;" id="editbudget" class="fa fa-pencil-square-o"></span></div>
                        <div class="col-lg-10"><button type="submit" class="btn btn-success btn-sm pull-right"><b><%=medarb_txt_020 %></b></button></div>
                    </div>

                    <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                        <thead>
                            <tr>
                                <th><%=kunder_txt_128 %></th>
                                <%

                                dim totalaar
                                redim totalaar(3000)

                                i = startaar                                 
                                for i = startaar TO slutaar %>
                                <th style="width:10%;"><%=i %></th>
                                <%next %>
                            </tr>
                        </thead>

                        <tbody id="list_body">
                            <%
                                strSQL = "SELECT id, kundeid, k.kkundenavn as kundenavn FROM budget_kunder LEFT JOIN kunder k ON (kundeid = k.kid) GROUP BY kundeid"
                                'response.Write "her " & strSQL 
                                oRec.open strSQL, oConn, 3
                                while not oRec.EOF
                            %>
                                <tr>
                                    <td><%=oRec("kundenavn") %> <input type='hidden' id='kundefindes_<%=oRec("kundeid") %>' value='1' /></td>

                                        <%
                                        salesgoal = 0
                                        i = startaar 
                                        for i = startaar TO slutaar

                                            strSQLBudget = "SELECT salesgoal, date_year FROM budget_kunder WHERE kundeid = "& oRec("kundeid") & " AND date_year = "&i
                                            'response.Write "<br> herher " & strSQLBudget 
                                            oRec2.open strSQLBudget, oConn, 3
                                            if not oRec2.EOF then

                                                salesgoal = oRec2("salesgoal")

                                            end if
                                            oRec2.close

                                            totalaar(i) = totalaar(i) + salesgoal

                                            response.Write "<td style=text-align:right;'><input type='hidden' name='FM_kundeid' value='"&oRec("kundeid")&"' /> <input type='text' name='FM_salesgoal_"&oRec("kundeid")&"_"&i&"' value='"&salesgoal&"' class='form-control input-small inputfield' style='display:none;' /> <span class='txtfield'>"&formatnumber(salesgoal,0)&"</span></td>"
                                        next
                                        %>
                                </tr>
                            <%
                                oRec.movenext
                                wend
                                oRec.close
                            %>
                        </tbody>

                        <tfoot>
                            <tr>
                                <th>Total</th>
                                
                                <%
                                t = startaar
                                for t = startaar TO slutaar %>
                                    
                                    <th style="text-align:right;"><%=formatnumber(totalaar(t), 0) %></th>

                                <%next %>
                            </tr>
                        </tfoot>

                    </table>
                </form>




            </div>
        </div>
    </div>


    <%end select %>











</div>
</div>


    

<!--#include file="../inc/regular/footer_inc.asp"-->