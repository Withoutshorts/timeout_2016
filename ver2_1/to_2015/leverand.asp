

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<div class="wrapper">
    <div class="content">
        <%
            if len(session("user")) = 0 then
	        %>
	        <!--#include file="../inc/regular/header_inc.asp"-->
	        <%
	        errortype = 5
	        call showError(errortype)
	        response.End
            end if
	
            call menu_2014

	        func = request("func")
	        if func = "dbopr" then
	        id = 0
	        else
	        id = request("id")
	        end if
	
	        rdir = request("rdir")


            select case func
            case "slet"
                '*** Her spørges om det er ok at der slettes en medarbejder ***
	            oskrift = leverand_txt_001
                slttxta = leverand_txt_002 & " <b>"&leverand_txt_003&"</b> "& leverand_txt_004
                slttxtb = "" 
                slturl = "leverand.asp?menu=mat&func=sletok&id="&id

                call sletcnf_2015(oskrift, slttxta, slttxtb, slturl)
            case "sletok"
	            '*** Her slettes en medarbejder ***
	            oConn.execute("DELETE FROM leverand WHERE id = "& id &"")
	            Response.redirect "leverand.asp?menu=mat&shokselector=1"
            case "dbopr", "dbred"
	            '*** Her indsættes eller redigeres en type i db ****

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

                strAdr = SQLBless(request("FM_adr"))
		        strPostnr = request("FM_postnr")
		        strCity = request("FM_city")
		        strLand = request("FM_land")
		        strEmail = request("FM_email")
		        strTlf = request("FM_tlf")
		        strFax = request("FM_fax")
		        strLevnr = request("FM_lnr")
		        strBesk = request("FM_besk")

                intType = request("FM_type")
		
		        strKpers1 = SQLBless(request("FM_kpers1"))
		        strKperstlf1 = request("FM_kperstlf1")
		        strKpersemail1 = request("FM_kpersemail1")
		
		        strKpers2 = SQLBless(request("FM_kpers2"))
		        strKperstlf2 = request("FM_kperstlf2")
		        strKpersemail2 = request("FM_kpersemail2")

                if func = "dbopr" then
		            strSQL = "INSERT INTO leverand (editor, dato, navn, "_
		            &" adresse, postnr, city, land, email, tlf, fax, besk, levnr, type, "_
		            &" kpers1, kperstlf1, kpersemail1, kpers2, kperstlf2, kpersemail2 "_
		            &" ) VALUES "_
		            &" ('"& strEditor &"', '"& strDato &"', '"& strNavn &"', "_
		            &" '"&strAdr&"', '"&strPostnr&"', '"&strCity&"', '"&strLand&"', '"&strEmail&"', "_
		            &" '"&strTlf&"', '"&strFax&"', '"&strBesk&"', "_
		            &" '"&strLevnr&"', "&intType&", '"&strKpers1&"', '"&strKperstlf1&"', "_
		            &" '"&strKpersemail1&"', '"&strKpers2 &"', "_
		            &" '"&strKperstlf2 &"', '"&strKpersemail2 &"')"		
		        else		
		            strSQL = "UPDATE leverand SET editor = '"& strEditor &"', dato = '"& strDato &"', navn = '"& strNavn &"', "_
		            &" adresse = '"&strAdr&"', postnr = '"&strPostnr&"', city = '"&strCity&"', land = '"&strLand&"', "_
		            &" email = '"&strEmail&"', tlf = '"&strTlf&"' , fax = '"&strFax&"', besk = '"&strBesk&"', levnr = '"&strLevnr&"', type = "&intType&", "_
		            &" kpers1 = '"&strKpers1&"', kperstlf1 = '"&strKperstlf1&"', kpersemail1 = '"&strKpersemail1&"', "_
		            &" kpers2 = '"&strKpers2 &"', kperstlf2 = '"&strKperstlf2 &"', kpersemail2 = '"&strKpersemail2 &"'"_
		            &" WHERE id = "& id
		        end if
		
		        oConn.execute(strSQL)

                if rdir = "mat" then
		            Response.redirect "materialer.asp?menu=mat"
		        else
		            Response.redirect "leverand.asp?menu=mat&shokselector=1"
		        end if




                case "opret", "red"
	                '*** Her indlæses form til rediger/oprettelse af ny type ***

                    if func = "opret" then
                        strNavn = ""
	                    strTimepris = ""
	                    varSubVal = leverand_txt_005
	                    varbroedkrumme = leverand_txt_024
	                    dbfunc = "dbopr"
                    else
                        strSQL = "SELECT editor, dato, navn, "_
	                    &" adresse, postnr, city, land, email, tlf, fax, besk, levnr, type, "_
	                    &" kpers1, kperstlf1, kpersemail1, kpers2, kperstlf2, kpersemail2 "_
	                    &" FROM leverand WHERE id=" & id
	                    oRec.open strSQL,oConn, 3

                        if not oRec.EOF then
		                    strNavn = oRec("navn")
		                    strDato = oRec("dato")
		                    strEditor = oRec("editor")
		
		                    strAdr = oRec("adresse")
		                    strPostnr = oRec("postnr")
		                    strCity = oRec("city")
		                    strLand = oRec("land")
		                    strEmail = oRec("email")
		                    strTlf = oRec("tlf")
		                    strFax = oRec("fax")
		                    strLevnr = oRec("levnr")
		                    strBesk = oRec("besk")
		
		                    intType = oRec("type")
		
		                    strKpers1 = oRec("kpers1")
		                    strKperstlf1 = oRec("kperstlf1")
		                    strKpersemail1 = oRec("kpersemail1")
		
		                    strKpers2 = oRec("kpers2")
		                    strKperstlf2 = oRec("kperstlf2")
		                    strKpersemail2 = oRec("kpersemail2")	
	                    end if
	                    oRec.close

                        dbfunc = "dbred"
	                    varbroedkrumme = leverand_txt_006
	                    varSubVal = leverand_txt_007

                    end if
                    
        %>

                    <div class="container">
                        <div class="portlet">
                            <h3 class="portlet-title"><u><%=leverand_txt_001 %> <span style="font-size:9px;"> - <%=varbroedkrumme%></span></u></h3>
                            <div class="portlet-body">
                                <form action="leverand.asp?menu=mat&func=<%=dbfunc%>" method="post">
                                    <input type="hidden" name="id" value="<%=id%>">
	                                <input type="hidden" name="rdir" value="<%=rdir%>">
                                    <div class="well well-white">
                                        <h4 class="panel-title-well">Stamdata</h4>

                                        <div class="row">
                                            <div class="col-lg-1"></div>
                                            <div class="col-lg-3"><%=leverand_txt_009 %>:</div>
                                            <div class="col-lg-3"><%=leverand_txt_010 %>:</div>
                                            <div class="col-lg-2"><%=leverand_txt_011 %>:</div>
                                            <div class="col-lg-2"><%=leverand_txt_012 %>:</div>
                                        </div>
                                        <div class="row">
                                            <div class="col-lg-1"></div>
                                            <div class="col-lg-3"><input type="text" name="FM_navn" value="<%=strNavn%>" class="form-control input-small" ></div>
                                            <div class="col-lg-3"><input type="text" name="FM_email" value="<%=strEmail%>" class="form-control input-small" ></div>
                                            <div class="col-lg-2"><input type="text" name="FM_lnr" value="<%=strLevnr%>" class="form-control input-small" ></div>
                                            <div class="col-lg-2">
                                                <select name="FM_type" id="FM_type" class="form-control input-small">
		                                            <%select case intType 
		                                            case 1
		                                            sel1 = "SELECTED"
		                                            sel2 = ""
		                                            sel3 = ""
		                                            case 2
		                                            sel1 = ""
		                                            sel2 = "SELECTED"
		                                            sel3 = ""
		                                            case 3
		                                            sel1 = ""
		                                            sel2 = ""
		                                            sel3 = "SELECTED"
		                                            case else
		                                            sel1 = "SELECTED"
		                                            sel2 = ""
		                                            sel3 = ""
		                                            end select%>
		                                            <option value="1" <%=sel1%>><%=leverand_txt_013 %></option>
		                                            <option value="2" <%=sel2%>><%=leverand_txt_014 %></option>
		                                            <option value="3" <%=sel3%>><%=leverand_txt_013 %> & <%=leverand_txt_014 %></option>
		                                        </select>
                                            </div>
                                        </div>
                                        
                                        <div class="row">
                                            <div class="col-lg-1"></div>
                                            <div class="col-lg-3"><%=leverand_txt_015 %>:</div>
                                            <div class="col-lg-1"><%=leverand_txt_016 %>:</div>
                                            <div class="col-lg-2"><%=leverand_txt_017 %>:</div>
                                            <div class="col-lg-2"><%=leverand_txt_018 %>:</div>
                                        </div>
                                        <div class="row">
                                            <div class="col-lg-1"></div>
                                            <div class="col-lg-3"><input type="text" name="FM_adr" value="<%=strAdr%>" class="form-control input-small" ></div>
                                            <div class="col-lg-1"><input type="text" name="FM_postnr" value="<%=strPostnr%>" class="form-control input-small" ></div>
                                            <div class="col-lg-2"><input type="text" name="FM_city" value="<%=strCity%>" class="form-control input-small" ></div>
                                            <div class="col-lg-2">
                                                <select name="FM_land" class="form-control input-small">
		                                            <%if func = "red" then%>
		                                            <option checked><%=strLand%></option>
		                                            <%else%>
		                                            <option checked>Danmark</option>
		                                            <%end if%>
		                                            <!--#include file="../timereg/inc/inc_option_land.asp"-->
		                                        </select>
                                            </div>
                                        </div>

                                        <div class="row">
                                            <div class="col-lg-1"></div>
                                            <div class="col-lg-3"><%=leverand_txt_019 %>.:</div>
                                            <div class="col-lg-3"><%=leverand_txt_020 %>:</div>
                                        </div>
                                        <div class="row">
                                            <div class="col-lg-1"></div>
                                            <div class="col-lg-3"><input type="text" name="FM_tlf" value="<%=strTlf%>" class="form-control input-small" ></div>
                                            <div class="col-lg-3"><input type="text" name="FM_fax" value="<%=strFax%>" class="form-control input-small" ></div>
                                        </div>

                                        <br />

                                        <div class="row">
                                            <div class="col-lg-1"></div>
                                            <div class="col-lg-2"><b><%=leverand_txt_021 & " 1:" %></b></div>
                                        </div>
                                        <div class="row">
                                            <div class="col-lg-1"></div>
                                            <div class="col-lg-3"><%=leverand_txt_009 %></div>
                                            <div class="col-lg-3"><%=leverand_txt_010 %></div>
                                            <div class="col-lg-2"><%=leverand_txt_019 %>.:</div>
                                        </div>
                                        <div class="row">
                                            <div class="col-lg-1"></div>
                                            <div class="col-lg-3"><input type="text" name="FM_kpers1" value="<%=strKpers1%>" class="form-control input-small" ></div>
                                            <div class="col-lg-3"><input type="text" name="FM_kpersemail1" value="<%=strKpersemail1%>" class="form-control input-small" ></div>
                                            <div class="col-lg-2"><input type="text" name="FM_kperstlf1" value="<%=strKperstlf1%>" class="form-control input-small" ></div>
                                        </div>

                                        <div class="row">
                                            <div class="col-lg-1"></div>
                                            <div class="col-lg-2"><b><%=leverand_txt_021 & " 2:" %></b></div>
                                        </div>
                                        <div class="row">
                                            <div class="col-lg-1"></div>
                                            <div class="col-lg-3"><%=leverand_txt_009 %>:</div>
                                            <div class="col-lg-3"><%=leverand_txt_010 %>:</div>
                                            <div class="col-lg-2"><%=leverand_txt_019 %>.:</div>
                                        </div>
                                        <div class="row">
                                            <div class="col-lg-1"></div>
                                            <div class="col-lg-3"><input type="text" name="FM_kpers2" value="<%=strKpers2%>" class="form-control input-small" ></div>
                                            <div class="col-lg-3"><input type="text" name="FM_kpersemail2" value="<%=strKpersemail2%>" class="form-control input-small" ></div>
                                            <div class="col-lg-2"><input type="text" name="FM_kperstlf2" value="<%=strKperstlf2%>" class="form-control input-small" ></div>
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="col-lg-12 pad-b10">
                                            <button type="submit" class="btn btn-success btn-sm pull-right"><b><%=varSubVal %></b></button>
                                        </div>
                                    </div>

                                </form>


                                <%if dbfunc = "dbred" then %>
                                    <br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
                                    <div class="row">
                                        <div class="col-lg-12" style="font-weight: lighter;"><%=leverand_txt_022 %> <b><%=strDato%></b> <%=leverand_txt_023 %> <b><%=strEditor%></b></div>
                                    </div>
                                <%end if %>

                            </div>
                        </div>
                    </div>



                <%case else %>
                    
                    <div class="container">
                        <div class="portlet">
                            <h3 class="portlet-title"><u><%=leverand_txt_001 %> <span style="font-size:9px;"></span></u></h3>
                            <div class="portlet-body">
                                <div class="row">
                                    <div class="col-lg-12"><a href="leverand.asp?menu=mat&func=opret" class="btn btn-success btn-sm pull-right"><b><%=leverand_txt_024 %></b></a></div>
                                </div>

                                <table id="orderTable" class="table dataTable table-striped table-bordered table-hover ui-datatable">
                                    <thead>
                                        <tr>
                                            <th><%=leverand_txt_025 %></th>
                                            <th><%=leverand_txt_009 %></th>
                                            <th><%=leverand_txt_013 %> / <%=leverand_txt_014 %></th>
                                            <th><%=leverand_txt_026 %></th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <%
                                            strSQL = "SELECT id, navn, type, levnr FROM leverand WHERE type = 1 OR type = 2 OR type = 3 ORDER BY navn"


                                            'Response.write strSQL
	                                        'Response.flush
	                                        oRec.open strSQL, oConn, 3
	                                        while not oRec.EOF 
	
	                                        antalordrer2 = 0


                                        %>
                                            <tr>
                                                <td><%=oRec("levnr")%></td>
                                                <td><a href="leverand.asp?menu=mat&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%> </a></td>
                                                <td>
                                                    <%select case oRec("type")
		                                            case 1
		                                            strType = leverand_txt_013
		                                            case 3
		                                            strType = leverand_txt_013 &" & "& leverand_txt_014
		                                            case else
		                                            strType = leverand_txt_014
		                                            end select
		
		                                            Response.write strType%>
                                                </td>
                                                <td style="text-align:center;"><a href="leverand.asp?menu=mat&func=slet&id=<%=oRec("id")%>"><span style="color:darkred;" class="fa fa-times"></span></a></td>
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

                    <script type="text/javascript">
                        $("#orderTable").DataTable({
        
                        });
                    </script>



        <%end select %>
    </div>
</div>

    

<!--#include file="../inc/regular/footer_inc.asp"-->