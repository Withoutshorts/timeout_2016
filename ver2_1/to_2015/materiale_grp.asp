

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

    call menu_2014()

    select case func

    case "slet"
	'*** Her sp�rges om det er ok at der slettes en medarbejder ***
	    oskrift = matgrp_txt_001 
        slttxta = matgrp_txt_002 & " <b>"&matgrp_txt_003&"</b> " & matgrp_txt_004
        slttxtb = "" 
        slturl = "materiale_grp.asp?menu=mat&func=sletok&id="&id

        call sletcnf_2015(oskrift, slttxta, slttxtb, slturl)


    case "sletok"
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM materiale_grp WHERE id = "& id &"")
	Response.redirect "materiale_grp.asp?menu=mat&shokselector=1"


    case "dbopr", "dbred"

        if len(trim(request("FM_navn"))) = 0 then		
		    errortype = 104
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
		strGrpNr = SQLBless(request("FM_nr"))
        medarbansv = request("FM_medarbansv")

        matgrp_konto = request("FM_matgrp_konto")

        %><!--#include file="../timereg/inc/isint_func.asp"--><%

        call erDetInt(request("FM_av")) 
		if isInt > 0 OR len(trim(request("FM_av"))) = 0 then
				
			%>
			<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
			<%
				
			errortype = 105
			call showError(errortype)
			
		isInt = 0
			
		Response.end
		end if

        strAv = replace(formatnumber(request("FM_av"), 0), ",",".")
        kundeid = request("FM_kunde")

        if len(trim(request("FM_fragtpris"))) <> 0 then
            fragtpris = request("FM_fragtpris")
            fragtpris = replace(fragtpris, ".", "")
            fragtpris = replace(fragtpris, ",", ".")
        else
            fragtpris = 0
        end if


        if func = "dbopr" then
		
		oConn.execute("INSERT INTO materiale_grp (navn, editor, dato, nummer, av, mgkundeid, medarbansv, matgrp_konto, fragtpris) VALUES "_
		&" ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', '"& strGrpNr &"', "& strAv &", "& kundeid &", "& medarbansv &", "& matgrp_konto &", "& fragtpris &")")
		
		else
		
		oConn.execute("UPDATE materiale_grp SET navn ='"& strNavn &"', "_
		&" nummer = '"& strGrpNr &"', editor = '" &strEditor &"', dato = '" & strDato &"', av = "& strAv &", mgkundeid = "& kundeid &", medarbansv = "& medarbansv &", matgrp_konto = "& matgrp_konto &", fragtpris = "& fragtpris &" WHERE id = "&id&"")
		
		    
		    if request("FM_opdater_av") = "1" then
	        strSQLupd = "UPDATE materialer SET salgspris = (indkobspris * (("& strAv &"/100) + 1)) WHERE matgrp = " & id
	        
	        oConn.execute(strSQLupd)
	        end if
		    
		
		end if


        Response.redirect "materiale_grp.asp?menu=mat&shokselector=1"



     case "opret", "red"

	    if func = "opret" then
	        strNavn = ""
	        varSubVal = matgrp_txt_005
	        varbroedkrumme = matgrp_txt_006
	        dbfunc = "dbopr"
	        strGrpNr = ""
	        sidst_redigeret = ""
	        strAv = 0
            medarbansv = 0
	        matgrp_konto = 0
            fragtpris = 0
	    else
	        strSQL = "SELECT navn, editor, dato, nummer, av, mgkundeid, medarbansv, matgrp_konto, fragtpris FROM materiale_grp WHERE id=" & id
	        oRec.open strSQL,oConn, 3
	
	        if not oRec.EOF then
	
	        strNavn = oRec("navn")
	        strDato = oRec("dato")
	        strEditor = oRec("editor")
	        strGrpNr = oRec("nummer")
	        strAv = oRec("av")
            kid = oRec("mgkundeid")
            medarbansv = oRec("medarbansv")
            matgrp_konto = oRec("matgrp_konto")
	
            if len(oRec("fragtpris")) <> 0 then
            fragtpris = oRec("fragtpris")
            else
            fragtpris = 0
            end if

	        end if
	        oRec.close
	
	        sidst_redigeret = "Sidst opdateret den <b>"&formatdatetime(strDato, 1)&"</b> af <b>"&strEditor&"</b>"
	        dbfunc = "dbred"
	        varbroedkrumme = matgrp_txt_007
	        varSubVal = matgrp_txt_008
	    end if

%>


        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u><%=matgrp_txt_001 %></u></h3>


                <div class="portlet-body">

                    <form action="materiale_grp.asp?menu=mat&func=<%=dbfunc%>" method="post">
                        <input type="hidden" name="id" value="<%=id%>">

                        <div class="row">
                                <div class="col-lg-12 pad-b10">
                                    <button type="submit" class="btn btn-success btn-sm pull-right"><b><%=medarb_txt_020 %></b></button>
                                </div>
                        </div>

                        <div class="well well-white">
                            <h4 class="panel-title-well"><%=medarb_txt_021 %></h4>

                            <div class="row">
                                <div class="col-lg-1"></div>
                                <div class="col-lg-3"><%=matgrp_txt_010 %>:</div>

                                <div class="col-lg-3"><%=matgrp_txt_011 %>:</div>
                                <div class="col-lg-3"><%=matgrp_txt_012 %>:</div>
                            </div>

                            <div class="row">
                                <div class="col-lg-1"></div>
                                <div class="col-lg-3"><input type="text" name="FM_navn" value="<%=strNavn%>" class="form-control input-small" ></div>
                                <div class="col-lg-3"><input type="text" name="FM_nr" value="<%=strGrpNr%>" class="form-control input-small" ></div>
                                <div class="col-lg-3">
                                    <select name="FM_kunde" id="FM_kunde" size="1" class="form-control input-small">
                                        <option value="0"><%=matgrp_txt_013 %></option>
		                                <%
		
		
				                        strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE ketype <> 'e' ORDER BY Kkundenavn"
				                        oRec.open strSQL, oConn, 3
				                        while not oRec.EOF
				
				                        if cint(kid) = cint(oRec("Kid")) then
				                        isSelected = "SELECTED"
				                        else
				                        isSelected = ""
				                        end if
				                        %>
				                        <option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%> (<%=oRec("Kkundenr") %>)</option>
				                        <%
				                        oRec.movenext
				                        wend
				                        oRec.close
				                        %>				
		                            </select>
                                </div>
                            </div>

                           

                            <div class="row">
                                <div class="col-lg-1"></div>
                                <div class="col-lg-3"><%=matgrp_txt_014 %> %</div>           
                                
                                <div class="col-lg-3"><%=matgrp_txt_015 %>:</div>
                                <div class="col-lg-3"><%=matgrp_txt_016 %></div>
                            </div>
                            <div class="row">
                                <div class="col-lg-1"></div>
                                <div class="col-lg-3"><input type="text" name="FM_av" value="<%=strAv%>" class="form-control input-small" style="display:inline-block; width:50%" > <span style="display:inline-block;"> % (<%=matgrp_txt_017 %>)</span></div>     
                                
                                <div class="col-lg-3">    
                                    <select name="FM_medarbansv" class="form-control input-small">
                                        <option value="0"><%=matgrp_txt_018 %>..</option><!-- onchange="submit(); -->
					                    <%
					                    strSQL = "SELECT Mid, Mnavn, init FROM medarbejdere WHERE mansat = 1 GROUP BY mid ORDER BY Mnavn"
					                    oRec.open strSQL, oConn, 3
					                    while not oRec.EOF 
					
					                    if cint(oRec("Mid")) = cint(medarbansv) then
					                    rchk = "SELECTED"
					                    else
					                    rchk = ""
					                    end if
                    
                                            if len(trim(oRec("init"))) <> 0 then
                                            medTxt = oRec("mnavn") & " ["& oRec("init") &"]"
                                            else
                                            medTxt = oRec("mnavn") 
                                            end if
                        
                                        %>
					                    <option value="<%=oRec("Mid")%>" <%=rchk%>><%=medTxt%></option>
					                    <%
					
					
					                    oRec.movenext
					                    wend
					                    oRec.close
				                    %>
                                    </select>
                                </div>

                                <div class="col-lg-3">
                                    <select name="FM_matgrp_konto" class="form-control input-small">
                                        <option value="0"><%=matgrp_txt_019 %>..</option>
                                        <%
                                            call licKid()

                                            '*** Kontoplan
                                            if licensindehaverKid <> 0 then
                                            licensindehaverKid = licensindehaverKid
                                            else
                                            licensindehaverKid = 0
                                            end if
        
                                             strSQLKontoliste = "SELECT navn, kontonr, id FROM kontoplan "_
                                             &" WHERE kid = "& licensindehaverKid &" ORDER BY navn"

                                            oRec.open strSQLkontoliste, oConn, 3
					                        while not oRec.EOF 
					
					                        if cint(matgrp_konto) = cint(oRec("id")) then
					                        rchk = "SELECTED"
					                        else
					                        rchk = ""
					                        end if

                                             %>
					                        <option value="<%=oRec("id")%>" <%=rchk%>><%=oRec("navn") & " ("&oRec("kontonr")%>)</option>
					                        <%

                                            oRec.movenext
					                        wend
					                        oRec.close
                                        %>
                                    </select>
                                </div>
                            </div>                          

                            <div class="row">
                                <div class="col-lg-1"></div>
                                <div class="col-lg-3"><input id="FM_opdater_av" name="FM_opdater_av" type="checkbox" value="1" /><%=" " & matgrp_txt_020 %></div>
                            </div>

                            <%if lto = "nt" then %>
                            <div class="row">
                                <div class="col-lg-1"></div>
                                <div class="col-lg-2"><br />Freight PC:</div>
                            </div>

                            <div class="row">
                                <div class="col-lg-1"></div>
                                <div class="col-lg-2"><input type="text" name="FM_fragtpris" class="form-control input-small" value="<%=fragtpris %>" /></div>
                                <div class="col-lg-1">&nbsp;</div>
                            </div>
                            <%else %>
                                <input type="hidden" name="FM_fragtpris" class="form-control input-small" value="0" />
                            <%end if %>

                            
                        </div>

                    </form>

                    <%if func = "red" then %>
                    <%
                        select case month(strDato)
                        case 1
                            monthTxt = matgrp_txt_030
                        case 2
                            monthTxt = matgrp_txt_031
                        case 3
                            monthTxt = matgrp_txt_032
                        case 4
                            monthTxt = matgrp_txt_033
                        case 5
                            monthTxt = matgrp_txt_034
                        case 6
                            monthTxt = matgrp_txt_035
                        case 7
                            monthTxt = matgrp_txt_036
                        case 8
                            monthTxt = matgrp_txt_037
                        case 9
                            monthTxt = matgrp_txt_038
                        case 10
                            monthTxt = matgrp_txt_039
                        case 11
                            monthTxt = matgrp_txt_040
                        case 12
                            monthTxt = matgrp_txt_041
                        end select
                    %>
                    <br /><br /> <br /><br /> <br /><br /> <br /><br /> <br /><br /> <br /><br />
                    <div class="row">
                        <div class="col-lg-12"><%=matgrp_txt_021 %> <b><%=day(strDato) &". "& monthTxt &" "& year(strDato) %></b> <%=matgrp_txt_022 %> <b> <%=strEditor %> </b></div>
                    </div>
                    <%end if %>

                </div>

            </div>
        </div>

        <%case else %>


        <script src="js/matgrp_liste.js" type="text/javascript"></script>
        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u><%=matgrp_txt_023 %></u></h3>

                <div class="portlet-body">

                    <div class="row">
                        <div class="col-lg-12"><a href="materiale_grp.asp?menu=mat&func=opret" class="btn btn-success btn-sm pull-right"><b><%=matgrp_txt_005 %></b></a></div>
                    </div>

                    <table id="matgpTable" class="table dataTable table-striped table-bordered table-hover ui-datatable">
                        <thead>
                            <tr>
                                <th><%=matgrp_txt_024 %> %</th>
                                <th><%=matgrp_txt_025 %></th>
                                <th><%=matgrp_txt_012 %></th>
                                <th><%=matgrp_txt_016 %></th>
                                <th><%=matgrp_txt_026 %> <br /> <%=matgrp_txt_027 %></th>
                                <th><%=matgrp_txt_028 %>

                                    <%if lto = "nt" then %>
                                    <br />Freight
                                    <%end if %>
                                </th>
                                <th><%=matgrp_txt_029 %></th>
                            </tr>
                        </thead>

                        <tbody>
                            <%
                                strSQL = "SELECT mg.id, mg.navn, mg.nummer, mg.av, mg.mgkundeid, k.kkundenavn, kkundenr, medarbansv, m.mnavn, init, matgrp_konto, kp.kontonr, kp.navn AS kontonavn, fragtpris FROM materiale_grp mg "_
                                &" LEFT JOIN kunder AS k ON (k.kid = mg.mgkundeid) "_
                                &" LEFT JOIN kontoplan AS kp ON (kp.id = mg.matgrp_konto) "_
                                &" LEFT JOIN medarbejdere AS m ON (m.mid = mg.medarbansv) WHERE mg.id > 0"

                                oRec.open strSQL, oConn, 3
	                            while not oRec.EOF 
                            %>
                                <tr>
                                    <td><a href="materiale_grp.asp?menu=mat&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%></a>&nbsp;&nbsp;<%=oRec("av") %> %</td>
                                    <td><%=oRec("nummer")%></td>
                                    <td>
                                        <%if oRec("mgkundeid") <> 0 then %>
                                        <%=oRec("kkundenavn") & " ("& oRec("kkundenr") &")" %>
                                        <%else %>
                                        &nbsp;
                                        <%end if %>
                                    </td>
                                    <td>
                                        <%if oRec("matgrp_konto") <> 0 then %>
                                        <%=oRec("kontonavn") & " ("& oRec("kontonr") &")" %>
                                        <%else %>
                                        &nbsp;
                                        <%end if %>
                                    </td>
                                    <td>
                                        <% if len(trim(oRec("init"))) <> 0 then
                                                    medTxt = oRec("mnavn") & " ["& oRec("init") &"]"
                                                    else
                                                    medTxt = oRec("mnavn") 
                                                    end if %>

                                        <%=medTxt %>&nbsp;
                                    </td>
                                    <td>
		                                <%
		                                antal = 0
		                                strSQL2 = "SELECT count(m.id) AS antal FROM materialer m "_
	                                    &" WHERE (m.matgrp = "& oRec("id") &") GROUP BY m.matgrp"
	                                    oRec2.open strSQL2, oConn, 3
                                        if not oRec2.EOF then
                                        antal = oRec2("antal")
                                        end if
                                        oRec2.close
	
		                                %>
		                                <%=antal%>

                                        <%if lto = "nt" then %>
                                        <br />Freight: <%=oRec("fragtpris") %>
                                        <%end if %>

                                    </td>
                                    <td>
		                                <%if cdbl(antal) = 0 then %>
		                                <a href="materiale_grp.asp?menu=mat&func=slet&id=<%=oRec("id")%>" class="red"><span style="color:darkred;" class="fa fa-times"></span></a>
		                                <%else %>
		                                &nbsp;
		                                <%end if %>
                                    </td>
                                </tr>
                            <%
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