

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

        func = request("func")
	    if func = "dbopr" then
	    id = 0
	    else
	    id = request("id")
	    end if

        call menu_2014

        select case func
            case "slet"
                '*** Her spørges om det er ok at der slettes ***
	        oskrift = "Selectbox Options" 
            slttxta = momskoder_txt_001 & " <b>" & momskoder_txt_002 & "</b> a selectbox option?"
            slttxtb = "" 
            slturl = "selectboxoptions.asp?menu=kon&func=sletok&id="& id

            call sletcnf_2015(oskrift, slttxta, slttxtb, slturl)

            case "sletok"
	        '*** Her slettes ***
	        oConn.execute("DELETE FROM selectboxoptions WHERE id = "& id &"")
	        Response.redirect "selectboxoptions.asp?menu=kon"


            case "dbopr", "dbred"

                if len(request("FM_navn")) = 0 then
                    errortype = 47
		            call showError(errortype)
                    response.End
                end if

             
                
                function SQLBless(s)
				dim tmp
				tmp = s
				tmp = replace(tmp, "'", "''")
				SQLBless = tmp
				end function
				
				function SQLBless2(s2)
				dim tmp2
				tmp2 = s2
				tmp2 = replace(tmp2, ",", ".")
				SQLBless2 = tmp2
				end function


                strNavn = SQLBless(request("FM_navn"))
				strEditor = session("user")
				strDato = year(now)&"/"&month(now)&"/"&day(now)
				selectboxlist = SQLBless2(request("FM_selectboxlist"))
				
				if func = "dbopr" then
				oConn.execute("INSERT INTO selectboxoptions (navn, editor, dato, selectboxlist) VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', "& selectboxlist &")")
				else
				oConn.execute("UPDATE selectboxoptions SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"', selectboxlist = "& selectboxlist &" WHERE id = "&id&"")
				end if
				
				Response.redirect "selectboxoptions.asp?menu=kon"

                case "opret", "red"
	                '*** Her indlæses form til rediger/oprettelse af ny type ***
	                if func = "opret" then
	                strNavn = ""
	                varSubTxt = momskoder_txt_011 
	                dbfunc = "dbopr"
	
	                else
	                strSQL = "SELECT navn, editor, dato, selectboxlist FROM selectboxoptions WHERE id=" & id
	                oRec.open strSQL,oConn, 3
	
	                if not oRec.EOF then
	                strNavn = oRec("navn")
	                strDato = oRec("dato")
	                strEditor = oRec("editor")
	                selectboxlist = oRec("selectboxlist")
	                end if
	                oRec.close
	
	                dbfunc = "dbred"
	                varSubTxt = momskoder_txt_010
	                end if

                    %>

                    <div class="container">
                        <div class="portlet">
                            <h3 class="portlet-title"><u>Selectbox Options</u></h3>

                            <div class="portlet-body">
                                <form action="selectboxoptions.asp?menu=tok&func=<%=dbfunc%>" method="post">
                                    <input type="hidden" name="id" value="<%=id%>">

                                    <div class="row">
                                        <div class="col-lg-12 pad-b5">
                                            <button type="submit" class="btn btn-success btn-sm pull-right"><b><%=varSubTxt %></b></button>
                                        </div>
                                    </div>

                                    <div class="well well-white">
                                         <div class="row">
                                            <div class="col-lg-12">
                                                <h4 class="panel-title-well"><%=medarb_txt_021 %></h4>
                                            </div>
                                        </div>

                                        <div class="row">
                                            <div class="col-lg-1"></div>
                                            <div class="col-lg-2"><%=momskoder_txt_005 %>:</div>
                                            <div class="col-lg-3"><input type="text" class="form-control input-small" name="FM_navn" value="<%=strNavn%>"></div>
                                        </div>

                                        <div class="row">
                                            <div class="col-lg-1"></div>
                                            <div class="col-lg-2">Apply to Selectbox:</div>
                                            <div class="col-lg-3">
                                                <select class="form-control input-small" name="FM_selectboxlist">
		                                            <%select case cint(selectboxlist)
		                                            case 0
		                                            sel0 = "SELECTED"
		                                            sel1 = ""
		                                            sel2 = ""
		                                            sel3 = ""
		                                            case 1
		                                            sel0 = ""
		                                            sel1 = "SELECTED"
                                                    sel2 = ""
		                                            sel3 = ""
                                                    case 2
		                                            sel0 = ""
                                                    sel1 = ""
                                                    sel2 = "SELECTED"
		                                            sel3 = ""
                                                    case 3
		                                            sel0 = ""
                                                    sel1 = ""
                                                    sel2 = ""
		                                            sel3 = "SELECTED"
		                                          
		                                            end select
		                                            %>
		                                            <option value="0" <%=sel0%>>Data Vendors</option>
		                                            <option value="1" <%=sel1%>>Categories of processing</option>
                                                    <option value="2" <%=sel2%>>Ordinary personal data</option>
                                                    <option value="3" <%=sel3%>>Sensitive personal data</option>
                                                    
		                                            
		                                        </select>
                                            </div>
                                        </div>


                                    </div>
                                </form>
                            </div>

                        </div>
                    </div>

                <%case else 'List %>

                <div class="container">
                    <div class="portlet">
                        <h3 class="portlet-title"><u>Selectbox Options</u></h3>

                        <div class="portlet-body">

                            <div class="row">
                                <div class="col-lg-12">
                                    <a href="selectboxoptions.asp?func=opret"><span class="btn btn-sm btn-success pull-right"><b><%=momskoder_txt_009 %> +</b></span></a>
                                </div>
                            </div>

                            <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                                <thead>
                                    <tr>
                                        <th style="width:55%;"><%=momskoder_txt_005 %></th>
                                        <th style="width:15%;">Page</th>
                                        <th style="width:20%;">Selectboxlist</th>
                                        <th style="text-align:center;"><%=momskoder_txt_007 %></th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <%
	                                strSQL = "SELECT id, navn, selectboxlist FROM selectboxoptions ORDER BY selectboxlist, navn"
	                                x = 0
	                                oRec.open strSQL, oConn, 3
	                                while not oRec.EOF 
	                                %>
                                    <tr>
                                        <td><a href="selectboxoptions.asp?menu=kon&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%> </a></td>
                                        <td>Job</td>
                                        <td><%
                                            
                                            selectboxlistName = ""
                                            select case oRec("selectboxlist") 
                                            case 0
                                               selectboxlistName = "Data Vendors"
		                                    case 1
                                               selectboxlistName = "Categories of processing"
                                            case 2
                                                selectboxlistName = "Ordinary personal data"
                                            case 3
                                                selectboxlistName = "Sensitive personal data"
                                            case 4
                                                selectboxlistName = "X-Old"
                                            end select
                                            %><%=selectboxlistName %></td>
                                        <td style="text-align:center;"><a href="selectboxoptions.asp?menu=kon&func=slet&id=<%=oRec("id")%>"><span style="color:darkred;" class="fa fa-times"></span></a></td>
                                    </tr>
                                    <%
	                                x = x + 1
	                                oRec.movenext
	                                wend
                                    oRec.close
	                                %>

                                    <%if x = 0 then %>
                                    <tr>
                                        <td colspan="2" style="text-align:center;"><%=momskoder_txt_008 %></td>
                                    </tr>
                                    <%end if %>
                                </tbody>
                            </table>
                        </div>

                    </div>
                </div>


                <%
        end select
    %>
        

</div>
</div>

<!--#include file="../inc/regular/footer_inc.asp"-->