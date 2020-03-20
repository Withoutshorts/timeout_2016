

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
	response.end
    end if
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if

    if request("print") <> "j" then
    call menu_2014()
    end if

    if len(request("FM_soeg")) <> 0 then 
	thiskri = request("FM_soeg")
	useKri = 1
	else
	thiskri = ""
	useKri = 0
	end if
	
	if len(request("FM_kd")) <> 0 then
	filterKD = request("FM_kd")
	else
	filterKD = 0
	end if
	
	if len(request("FM_tp")) <> 0 then
	filterTP = request("FM_tp")
	else
	filterTP = 0
	end if
	
	
	if len(request("FM_status")) <> 0 then
	filterStatus = request("FM_status")
	else
	filterStatus = 1
	end if
	
	
	if len(request("FM_nogletal")) <> 0 THEN 
	filterNogletal = request("FM_nogletal")
	else
	filterNogletal = 0
	end if



    select case func

        case "slet"
            '*** Her spørges om det er ok at der slettes en medarbejder ***
	        oskrift = kontoplan_txt_001 
            slttxta = kontoplan_txt_002 & "<b>"&kontoplan_txt_003&"</b> " & kontoplan_txt_004 
            slttxtb = "" 
            slturl = "kontoplan.asp?menu=kon&func=sletok&id="&id

        call sletcnf_2015(oskrift, slttxta, slttxtb, slturl)

        case "sletok"
	        '*** Her slettes en konto ***
	
	        kontonr = -3
	        strSQL = "SELECT kontonr FROM kontoplan WHERE id = "& id
	        oRec.open strSQL, oConn, 3
	        if not oRec.EOF then
	        kontonr = oRec("kontonr")
	        end if
	        oRec.close
	
	        '*** Sletter posteringer ***
	        oConn.execute("DELETE FROM posteringer WHERE kontonr = "& kontonr &" OR modkontonr = "& kontonr &"")
	
	        '**** Sletter konto ****
	        oConn.execute("DELETE FROM kontoplan WHERE id = "& id &"")
	        Response.redirect "kontoplan.asp?menu=erp"

        case "dbopr", "dbred"

            showmenu = request("showmenu")

            if len(request("FM_navn")) = 0 OR len(request("FM_kontonr")) = 0 then
                if showmenu = "no" then
		        %>
		        
		        <%
		        useleftdiv = "j"
		        else
		        %>
		        
		
		        <%
		        end if
		
		        errortype = 46
		        call showError(errortype)        
                response.End
            end if

            %>
			<!--#include file="../timereg/inc/isint_func.asp"-->
			<%

            call erDetInt(request("FM_kontonr"))
            if isInt > 0 OR instr(request("FM_kontonr"), ".") <> 0 then

                if showmenu = "no" then
				%>
				
				<%
				useleftdiv = "j"
				else
				%>
				
				
				<%
				end if

                errortype = 62
			    call showError(errortype)	
                response.End
			    isInt = 0
            end if


            intkontonr = request("FM_kontonr")

            '**** Findes kontonr allerede ?? **
		    errKontonrFindes = 0
		    strSQL = "SELECT kontonr FROM kontoplan WHERE kontonr = "& intkontonr & " AND id <> "& id
		    oRec.open strSQL, oConn, 3 
		    while not oRec.EOF 
							
		    errKontonrFindes = 1
							
		    oRec.movenext
		    wend
		    oRec.close 
				
		    if errKontonrFindes = 1 then                   
                if showmenu = "no" then
				%>
				
				<%
				useleftdiv = "j"
				else
				%>
				
							
				<%
				end if

		        errortype = 46
		        call showError(errortype)
                response.end
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

            kundeid = request("FM_kundeid")
			intMomskode = request("FM_momskode")
			intKeycode = request("FM_keycode")
			intDebkre = request("FM_debkre")
			intStatus = request("FM_status")
			intType = request("FM_type")

            if func = "dbopr" then
			oConn.execute("INSERT INTO kontoplan (navn, editor, dato, kontonr, kid, "_
			&" momskode, debitkredit, status, keycode, type) "_
			&" VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', "_
			&" '"& intkontonr &"', '"& kundeid &"', "& intMomskode &", "&intDebkre&", "_
			&" "&intStatus&", "&intKeycode&", "& intType &")")
			else
			oConn.execute("UPDATE kontoplan SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"', kontonr = '"& intkontonr &"', "_
			&" momskode = "& intMomskode &", kid = '"& kundeid &"', debitkredit = "& intDebkre &", "_
			&" status = "&intStatus&", keycode = "&intKeycode&", type = "& intType &" WHERE id = "&id&"")
			end if

            Response.redirect "kontoplan.asp?menu=erp&id="&id




            case "opret", "red"
	        '*** Her indlæses form til rediger/oprettelse af ny type ***
	        if func = "opret" then
	        strNavn = ""
	        intMomskode = 1
	        varSubVal = "Opretpil" 
	        varbroedkrumme = kontoplan_txt_005
	        dbfunc = "dbopr"
	        intdebkre = 2
	        intType = 1
	        intStatus = 1
	        intNogletalskode = 0
	
	        else
	        strSQL = "SELECT navn, editor, dato, id, kid, kontonr, momskode, keycode, debitkredit, status, type FROM kontoplan WHERE id=" & id
	        oRec.open strSQL,oConn, 3
	
	        if not oRec.EOF then
	        strNavn = oRec("navn")
	        strDato = oRec("dato")
	        strEditor = oRec("editor")
	        intKid = oRec("kid")
	        intdebkre = oRec("debitkredit")
	        intStatus = oRec("status")
	        intMomskode = oRec("momskode")
	        varKontonr = oRec("kontonr")
	        intNogletalskode = oRec("keycode")
	        intType = oRec("type")
	        end if
	        oRec.close
	
	        dbfunc = "dbred"
	        varbroedkrumme = kontoplan_txt_006
	        varSubVal = "Opdaterpil" 
	        end if
	
	        if len(request("showmenu")) <> 0 then
	        showmenu = "no"
	        leftdiv = 20
	        topdiv = 80
	        else
	        leftdiv = 90
	        topdiv = 102
	        showmenu = ""
	        end if 
	        
    %>

    <div class="container">
        <div class="portlet">
            <h3 class="portlet-title"><u><%=varbroedkrumme %></u></h3>

            <div class="portlet-body">

                <form action="kontoplan.asp?menu=kon&func=<%=dbfunc%>&showmenu=<%=showmenu%>" method="post">
                    <input type="hidden" name="id" value="<%=id%>">

                    <div class="row">
                        <div class="col-lg-12 pad-b5">
                        <button type="submit" class="btn btn-success btn-sm pull-right"><b><%=medarb_txt_020 %></b></button>
                        </div>
                    </div>

                    <div class="well well-white">
                        <h4 class="panel-title-well"><%=medarb_txt_021 %></h4>

                        <div class="row">
                            <div class="col-lg-1"></div>
                            <div class="col-lg-3"><%=kontoplan_txt_007 %>:</div>
                            <div class="col-lg-3"><%=kontoplan_txt_008 %>:</div>
                            <div class="col-lg-3"><%=kontoplan_txt_009 %>:</div>
                        </div>

                        <div class="row">
                            <div class="col-lg-1"></div>
                            <div class="col-lg-3">
                                <%if func = "red" then %>
                                <input type="text" value="<%=varKontonr%>" class="form-control input-small" readonly>
                                <input type="hidden" name="FM_kontonr" value="<%=varKontonr%>">
                                <%else %>
                                <input type="text" name="FM_kontonr" value="<%=varKontonr%>" size="30" class="form-control input-small">
                                <%end if %>
                            </div>
                            <div class="col-lg-3"><input type="text" name="FM_navn" value="<%=strNavn%>" class="form-control input-small"></div>
                            <div class="col-lg-3">
                                <select name="FM_kundeid" class="form-control input-small">		
		                            <%
		                            if showmenu = "no" then '*** Fra Kontakt opr.
			                            if len(request("kundeid")) <> 0 then
			                            thisKid = request("kundeid")
			                            else
			                            thisKid = 0
			                            end if
		                            sqlWhere = " WHERE kid = "& thisKid
		                            else
		                            sqlWhere = "" 
		                            end if
		
		                            strSQL = "SELECT kkundenavn, kkundenr, kid, useasfak FROM kunder "& sqlWhere &" ORDER BY kkundenavn"
		                            oRec.open strSQL, oConn, 3 
		                            while not oRec.EOF 
		
		
		
			                            if func <> "red" then
				                            if oRec("useasfak") = 1 then
				                            selth = "SELECTED"
				                            else
				                            selth = ""
				                            end if
			                            else
				                            if oRec("kid") = cint(intKid) then
				                            selth = "SELECTED"
				                            else
				                            selth = ""
				                            end if
			                            end if
		
		
		
		                            %>
		                            <option value="<%=oRec("kid")%>" <%=selth%>><%=oRec("kkundenavn")%> (<%=oRec("kkundenr")%>)</option>
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
                            <div class="col-lg-3"><%=kontoplan_txt_010 %>:</div>
                            <div class="col-lg-3"><%=kontoplan_txt_011 %>:</div>
                            <div class="col-lg-3"><%=kontoplan_txt_012 %>:</div>
                        </div>

                        <div class="row">
                            <div class="col-lg-1"></div>
                            <div class="col-lg-3">
                                <select name="FM_debkre" class="form-control input-small">
		                        <%
		                        select case intdebkre 
		                        case 1 
		                        selth1 = "SELECTED"
		                        selth2 = ""
		                        selth3 = ""
		                        selth4 = ""
		                        selth5 = ""
		                        selth6 = ""
		                        case 2
		                        selth1 = ""
		                        selth2 = "SELECTED"
		                        selth3 = ""
		                        selth4 = ""
		                        selth5 = ""
		                        selth6 = ""
		                        case 3
		                        selth1 = ""
		                        selth2 = ""
		                        selth3 = "SELECTED"
		                        selth4 = ""
		                        selth5 = ""
		                        selth6 = ""
		                        case 4
		                        selth1 = ""
		                        selth2 = ""
		                        selth3 = ""
		                        selth4 = "SELECTED"
		                        selth5 = ""
		                        selth6 = ""
		                        case 5
		                        selth1 = ""
		                        selth2 = ""
		                        selth3 = ""
		                        selth4 = ""
		                        selth5 = "SELECTED"
		                        selth6 = ""
		                        case 6
		                        selth1 = ""
		                        selth2 = ""
		                        selth3 = ""
		                        selth4 = ""
		                        selth5 = ""
		                        selth6 = "SELECTED"
		                        end select
		                        %>
		                        <option value="1" <%=selth1%>><%=kontoplan_txt_013 %></option>
		                        <option value="2" <%=selth2%>><%=kontoplan_txt_014 %></option>
		                        <option value="3" <%=selth3%>><%=kontoplan_txt_015 %></option>
		                        <option value="4" <%=selth4%>><%=kontoplan_txt_016 %></option>
		                        <option value="5" <%=selth5%>><%=kontoplan_txt_017 %></option>
		                        <option value="6" <%=selth6%>><%=kontoplan_txt_018 %></option>
		
		                        </select>
                            </div>
                            <div class="col-lg-3">
                                <select name="FM_type" id="FM_type" class="form-control input-small">
		                            <%select case cint(intType)
		                            case 1
		                            selth1 = "SELECTED"
		                            selth2 = ""
		                            selth3 = ""
		                            case 2
		                            selth1 = ""
		                            selth2 = "SELECTED"
		                            selth3 = ""
		                            case else
		                            selth1 = ""
		                            selth2 = ""
		                            selth3 = "SELECTED"
		                            end select
		                            %>
		                            <option value="1" <%=selth1%>><%=kontoplan_txt_019 %></option>
		                            <option value="2" <%=selth2%>><%=kontoplan_txt_020 %></option>
		                            <option value="3" <%=selth3%>><%=kontoplan_txt_018 %></option>
		                            </select>
                            </div>
                            <div class="col-lg-3">
                                <select name="FM_momskode" id="FM_momskode" class="form-control input-small">
		                            <%
		                            'if func = "red" then
		                            'strSQL = "SELECT navn, id, kvotient FROM momskoder WHERE id="&intMomskode
		                            'else
		                            strSQL = "SELECT navn, id, kvotient FROM momskoder ORDER BY navn"
		                            'end if
		
		                            oRec.open strSQL, oConn, 3 
		                            while not oRec.EOF 
		
		                            if intMomskode = oRec("id") then
		                            momSel = "SELECTED"
		                            else
		                            momSel = ""
		                            end if
		                            %>
		                            <option value="<%=oRec("id")%>" <%=momSel%>><%=oRec("navn")%> (<%=oRec("kvotient")%>)</option>
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
                            <div class="col-lg-3"><%=kontoplan_txt_021 %>:</div>
                            <div class="col-lg-3"><%=kontoplan_txt_020 %></div>
                        </div>
                        <div class="row">
                            <div class="col-lg-1"></div>
                            <div class="col-lg-3">
                                <select name="FM_keycode" class="form-control input-small">
		                            <option value="0"><%=kontoplan_txt_018 %></option>
		                            <%
		                            strSQL = "SELECT navn, id, ap_type FROM nogletalskoder ORDER BY navn"
		                            oRec.open strSQL, oConn, 3 
		                            while not oRec.EOF 
		                            if intNogletalskode = oRec("id") then
		                            keySel = "SELECTED"
		                            else
		                            keySel = ""
		                            end if
		
		                            select case oRec("ap_type")
		                            case 1
		                            ap_type = "Aktiver" 
		                            case 2
		                            ap_type = "Passiver"
		                            case else
		                            ap_type = "Ingen"
		                            end select
		
		
		                            %>
		                            <option value="<%=oRec("id")%>" <%=keySel %>><%=oRec("navn")%> (<%=ap_type%>)</option>
		                            <%
		                            oRec.movenext
		                            wend
		                            oRec.close
		                            %>
		                        </select>

                            </div>
                            <div class="col-lg-3">
                                <select name="FM_status" class="form-control input-small">

                                    <%
	                                    select case intStatus
	                                    case 1
	                                    stSEL1 = "SELECTED"
	                                    stSEL2 = ""
	                                    case 2
	                                    stSEL1 = ""
	                                    stSEL2 = "SELECTED"
	                                    case else
	                                    stSEL1 = "SELECTED"
	                                    stSEL2 = ""
	                                    end select
	                                %>

		                            <option value="1" <%=stSEL1 %>><%=kontoplan_txt_024 %></option>
		                            <option value="2" <%=stSEL2 %>><%=kontoplan_txt_025 %></option>
		                        </select>
                            </div>
                        </div>

                    </div>
                </form>

                <%if dbfunc = "dbred" then
                    
                    
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


    <%case else 
        
        print = request("print")
        
    %>


    <%
        if len(trim(request("sort"))) <> 0 then
	    sort = request("sort")
	            
	            if sort = 2 then
	            sortCHK2 = "CHECKED"
	            sortCHK1 = ""
	            sortOrderKri = "kontoplan.kontonr, kkundenavn"
	            else
	            sortCHK2 = ""
	            sortCHK1 = "CHECKED"
	            sortOrderKri = "kkundenavn"
	            end if
	   
	    else
	    sort = 2
	    sortCHK2 = "CHECKED"
	    sortCHK1 = ""
	    sortOrderKri = "kontoplan.kontonr, kkundenavn"
	    end if

        'if len(request("sort")) <> 0 then
            
            if len(request("visikkenul")) <> 0 then
            visikkenul = 1
            visikkenulCHK = "CHECKED"
            else
            visikkenul = 0
            visikkenulCHK = ""
            end if
        'else
            'if len(request.cookies("erp")("viskunkontimpos")) <> 0 then
                'if request.cookies("erp")("viskunkontimpos") = 1 then
                'visikkenul = 1
                'visikkenulCHK = "CHECKED"
                'else
                'visikkenul = 0
                'visikkenulCHK = ""
                'end if
             'else
                'visikkenul = 0
                'visikkenulCHK = ""
             'end if
       'end if 
       
       response.cookies("erp")("viskunkontimpos") = visikkenul
        
       expTxt = ""
	
        if len(trim(request("nolist"))) <> 0 then
        nolist = 1
        else
	    nolist = 0
        end if


        if len(trim(request("FM_startdato"))) <> 0 then
            strStartDato = request("FM_startdato")
            strStartDatoSQL = year(strStartDato) &"-"& month(strStartDato) &"-"& day(strStartDato)

            response.Cookies("kontoplan")("FM_startdato") = strStartDato
        else
            if request.Cookies("kontoplan")("FM_startdato") <> "" then
                strStartDato = day(request.Cookies("kontoplan")("FM_startdato")) &"-"& month(request.Cookies("kontoplan")("FM_startdato")) &"-"& year(request.Cookies("kontoplan")("FM_startdato"))                
            else
                strStartDatoAdd = DateAdd("m", -1, now)
                strStartDato = day(strStartDatoAdd) &"-"& month(strStartDatoAdd) &"-"& year(strStartDatoAdd)
                
            end if
            strStartDatoSQL = year(strStartDato) &"-"& month(strStartDato) &"-"& day(strStartDato) 
        end if

        if len(trim(request("FM_slutdato"))) <> 0 then
            strSlutDato = request("FM_slutdato")
            strSlutDatoSQL = year(strSlutDato) &"-"& month(strSlutDato) &"-"& day(strSlutDato) 
            
            response.Cookies("kontoplan")("FM_slutdato") = strSlutDato

        else
            if request.Cookies("kontoplan")("FM_slutdato") <> "" then
                strSlutDato = day(request.Cookies("kontoplan")("FM_slutdato")) &"-"& month(request.Cookies("kontoplan")("FM_slutdato")) &"-"& year(request.Cookies("kontoplan")("FM_slutdato"))
            else
                strSlutDato = day(now) &"-"& month(now) &"-"& year(now)                
            end if

            strSlutDatoSQL = year(strSlutDato) &"-"& month(strSlutDato) &"-"& day(strSlutDato) 
        end if

        'response.Write "<br> HER HER ST" & strStartDatoSQL
        'response.Write "<br> HER HER SL" & strSlutDatoSQL

    %>
    <script src="js/kontoplan_jav2.js" type="text/javascript"></script> 
    <div class="container">
        <div class="portlet">
            <h3 class="portlet-title"><u><%=kontoplan_txt_001 %></u></h3>
            
            <div class="portlet-body">
                
                <%if print <> "j" then %>

                <div class="row">
                    <div class="col-lg-12 pad-b5">
                        <a href="kontoplan.asp?menu=kon&func=opret" class="btn btn-sm btn-success pull-right"><b><%=kontoplan_txt_005 %></b></a>
                    </div>
                </div>

                <form action="kontoplan.asp?menu=kon" method="POST">
                <div class="well">
                    <h4 class="panel-title-well"><%=kontoplan_txt_026 %></h4>

                    <div class="row">
                        <div class="col-lg-3"><b><%=kontoplan_txt_027 %></b> (<%=kontoplan_txt_028 %>)</div>

                        <div class="col-lg-2"><%=kontoplan_txt_029 %>:</div>
                        <div class="col-lg-2"><%=kontoplan_txt_010 %>:</div>
                        <div class="col-lg-2"><%=kontoplan_txt_020 %>:</div>
                        <div class="col-lg-2"><%=kontoplan_txt_030 %></div>
                    </div>

                    <div class="row">
                        <div class="col-lg-3"><input type="text" name="FM_soeg" id="FM_soeg" value="<%=thiskri%>" class="form-control input-small"></div>
                        <div class="col-lg-2">
                            <select name="FM_tp" id="FM_tp" class="form-control input-small">
			                    <%
			                    select case filterTP 
			                    case 0
			                    sel0 = "SELECTED"
			                    sel1 = ""
			                    sel2 = ""
			                    case 1
			                    sel0 = ""
			                    sel1 = "SELECTED"
			                    sel2 = ""
			                    case 2
			                    sel0 = ""
			                    sel1 = ""
			                    sel2 = "SELECTED"
			                    case else
			                    sel0 = "SELECTED"
			                    sel1 = ""
			                    sel2 = ""
			                    end select
			                    %>
			                    <option value="0" <%=sel0%>><%=kontoplan_txt_031 %></option>
			                    <option value="2" <%=sel2%>><%=kontoplan_txt_019 %></option>
			                    <option value="1" <%=sel1%>><%=kontoplan_txt_020 %></option>
			                </select>
                        </div>
                        <div class="col-lg-2">
                            <select name="FM_kd" id="FM_kd" class="form-control input-small">
			                    <%
			                    select case filterKD 
			                    case 0
			                    sel0 = "SELECTED"
			                    sel1 = ""
			                    sel2 = ""
			                    case 1
			                    sel0 = ""
			                    sel1 = "SELECTED"
			                    sel2 = ""
			                    case 2
			                    sel0 = ""
			                    sel1 = ""
			                    sel2 = "SELECTED"
			                    case else
			                    sel0 = "SELECTED"
			                    sel1 = ""
			                    sel2 = ""
			                    end select
			                    %>
			                    <option value="0" <%=sel0%>><%=kontoplan_txt_031 %></option>
			                    <option value="2" <%=sel2%>><%=kontoplan_txt_032 %></option>
			                    <option value="1" <%=sel1%>><%=kontoplan_txt_033 %></option>
			                </select>
                        </div>
                        <div class="col-lg-2">
                            <select name="FM_status" id="FM_status" class="form-control input-small">
			                    <%
			                    select case filterStatus 
			                    case 0
			                    sel0 = "SELECTED"
			                    sel1 = ""
			                    sel2 = ""
			                    case 1
			                    sel0 = ""
			                    sel1 = "SELECTED"
			                    sel2 = ""
			                    case 2
			                    sel0 = ""
			                    sel1 = ""
			                    sel2 = "SELECTED"
			                    case else
			                    sel0 = "SELECTED"
			                    sel1 = ""
			                    sel2 = ""
			                    end select
			                    %>
			                    <option value="0" <%=sel0%>><%=kontoplan_txt_031 %></option>
			                    <option value="1" <%=sel1%>><%=kontoplan_txt_047 %></option>
			                    <option value="2" <%=sel2%>><%=kontoplan_txt_048 %></option>
			                </select>
                        </div>
                        <div class="col-lg-2">
                            <select name="FM_nogletal" id="FM_nogletal" class="form-control input-small">
			                    <option value="0" <%=sel0%>><%=kontoplan_txt_031 %></option>
			                    <%
			
			                    strSQLnt = "SELECT id, navn, ap_type FROM nogletalskoder WHERE id <> 0 ORDER BY navn"
			                    oRec.open strSQLnt, oConn, 3
			                    while not oRec.EOF
			
			                    if cint(filterNogletal) = oRec("id") then
			                    selNt = "SELECTED"
			                    else
			                    selNt = ""
			                    end if 
			
			                    %>
			                    <option value="<%=oRec("id") %>" <%=selNt%>><%=oRec("navn") %></option>
			
			                    <%
			                    oRec.movenext
			                    wend
			                    oRec.close
			
			
			                    %>
		
			                </select>
                        </div>
                    </div>

                    <br />

                    <div class="row">
                        <div class="col-lg-2"><%=kontoplan_txt_034 %>:</div>
                    </div>

                    <div class="row">
                        <div class="col-lg-3">
                            <table>
                                <tr>
                                    <td>
                                        <div class='input-group date' id='datepicker_stdato'>
                                        <input type="text" class="form-control input-small" name="FM_startdato" value="<%=strStartDato %>" placeholder="dd-mm-yyyy" />
                                            <span class="input-group-addon input-small">
                                                <span class="fa fa-calendar">
                                                </span>
                                            </span>
                                        </div>
                                    </td>
                                    <td style="padding-left:10px; padding-right:10px;">til</td>
                                    <td>
                                        <div class='input-group date' id='datepicker_stdato'>
                                        <input type="text" class="form-control input-small" name="FM_slutdato" value="<%=strSlutDato %>" placeholder="dd-mm-yyyy" />
                                            <span class="input-group-addon input-small">
                                                <span class="fa fa-calendar">
                                                </span>
                                            </span>
                                        </div>
                                    </td>
                                </tr>                           
                            </table>
                        </div>
                        <div class="col-lg-4"><input id="visikkenul" name="visikkenul" type="checkbox" <%= visikkenulCHK %> /><%=" " & kontoplan_txt_035 %></div>

                        <div class="col-lg-4"><button type="submit" class="btn btn-secondary btn-sm pull-right"><b><%=kontoplan_txt_036 %></b></button></div>

                    </div>
                    
                </div>
                </form>

                <%end if %>

                <% 
                if cint(nolist) = 1 then
                    response.end
	            end if %>

                <%
                    if print = "j" then                          
                    trColor = "#f7f7f7"
                    tableclass = "" 
                    tableid = ""
                    else
                    trColor = ""
                    tableclass = "table-striped"
                    tableid = "kontoplan_table"
                    end if   
                %>
                
                <table id="<%=tableid %>" class="table dataTable table-bordered table-hover ui-datatable <%=tableclass %>">
                    <thead>
                        <tr>
                            <th><%=kontoplan_txt_007 %></th>
                            <th><%=kontoplan_txt_008 %></th>
                            <th><%=kontoplan_txt_009 %></th>
                            <th><%=kontoplan_txt_010 %></th>
                            <th><%=kontoplan_txt_012 %></th>
                            <th><%=kontoplan_txt_011 %></th>
                            <th><%=kontoplan_txt_021 %></th>

                            <%if print <> "j" then%>
	                        <th style="text-align:center;"><%=kontoplan_txt_020 %></th>
	                        <%else%>
	                        <!-- <th></th> -->
	                        <%end if%>

                            <th style="text-align:right;"><%=kontoplan_txt_037 %></th>

                            <%if print <> "j" then %>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <%end if %>
		                       
                            <%
                            expTxt = expTxt & kontoplan_txt_007&";"&kontoplan_txt_008&";"&kontoplan_txt_009&";"&kontoplan_txt_014&" / "&kontoplan_txt_013&";"&kontoplan_txt_012&";"&kontoplan_txt_011&";"&kontoplan_txt_021&";"&kontoplan_txt_020&";"&kontoplan_txt_037&";"
                            expTxt = expTxt & kontoplan_txt_038&";"&kontoplan_txt_049&";"&kontoplan_txt_039&";"&kontoplan_txt_040&";"&kontoplan_txt_041&";"

                            'expTxt = expTxt &"Kontonr;Kontonavn;Kontoindehaver;Debitor/kreditor;Momskode;Type;Nøgletalskode;Status;Saldo"
	                        'expTxt = expTxt &";Dato;Bilagsnr;Tekst;Debit;Kredit"
                            %>

                        </tr>
                    </thead>


                    <tbody>

                        <%
                            if useKri <> 0 then
	                        useSQLKri = " WHERE (kontoplan.navn LIKE '"&thiskri&"%' OR kontoplan.kontonr = '"&thiskri&"' OR kkundenavn LIKE '"&thiskri&"%')" 
	                        else
	                        useSQLKri = " WHERE kontoplan.id > 0 "
	                        end if
	
	                        select case filterKD
	                        case 0
	                        dkkri = " AND debitkredit <> 99 "
	                        case 1
	                        dkkri = " AND debitkredit = 1 "
	                        case 2
	                        dkkri = " AND debitkredit = 2 "
	                        case else
	                        dkkri = " AND debitkredit <> 99 "
	                        end select
	
	                        select case filterTP
	                        case 0
	                        typkri = " AND type <> 99 "
	                        case 1
	                        typkri = " AND type = 2 "
	                        case 2
	                        typkri = " AND type = 1 "
	                        case else
	                        typkri = " AND type <> 99 "
	                        end select
	
	                        select case filterStatus
	                        case 0
	                        statusKri = " AND kontoplan.status <> -1 "
	                        case 1
	                        statusKri = " AND kontoplan.status = " & filterStatus
	                        case 2
	                        statusKri = " AND kontoplan.status = " & filterStatus
	                        end select
	
	                        select case filterNogletal
	                        case 0
	                        ntKri = " AND keycode <> -1"
	                        case else
	                        ntKri = " AND keycode = " & filterNogletal
	                        end select
	
	                        if visikkenul <> 1 then
	                        grpby = "kontoplan.kontonr"
	                        else
	                        grpby = "p.kontonr"
	                        end if
	
	                        imoms = 0
	                        umoms = 0
	                        saldo = 0
	
	                        if print <> "j" then
	                        btop = 0
	                        else
	                        btop = 1
	                        end if

                            	
	                        periodefilter = " posteringsdato BETWEEN '"& strStartDatoSQL &"' AND '"& strSlutDatoSQL &"'"
	

	                        strSQL = "SELECT kontoplan.id AS id, kontoplan.navn AS navn, kontoplan.kontonr, kontoplan.kid AS kid, "_
	                        &" kontoplan.type, kkundenavn, kkundenr, kunder.kid AS kunderid, momskode, "_
	                        &" kontoplan.status, debitkredit, momskoder.navn AS momskodenavn, "_
	                        &" nogletalskoder.navn AS nogletalkodenavn, sum(p.nettobeloeb) AS saldo "_
	                        &" FROM kontoplan LEFT JOIN kunder ON (kunder.kid = kontoplan.kid) "_
	                        &" LEFT JOIN momskoder ON (momskoder.id = momskode) "_
	                        &" LEFT JOIN nogletalskoder ON (nogletalskoder.id = keycode) "_
	                        &" LEFT JOIN posteringer p ON (p.kontonr = kontoplan.kontonr AND "& periodefilter &" AND p.status = 1)"_
                            &  useSQLKri &" "& dkkri &" "& typkri &" "& statusKri &" "& ntKri &" GROUP BY "& grpby &" ORDER BY "& sortOrderKri 
	
	                        'Response.Write strSQL
	                        'Response.flush
	
	                        oRec.open strSQL, oConn, 3
	                        while not oRec.EOF 
	
	                        if print = "j" then
	                        bgcol = "#8caae6" '"#FFFF99"
	                        else
	
		                        if cint(id) = oRec("id") then
		                        bgcol = "#ffff99"
		                        else
			
			                        'if oRec("debitkredit") = 5 then 'Overskrift
		                            'bgcol = "#cccccc"
		                            'else
		                            bgcol = "#FFFFFF"
		                            'end if
			
		
		                        end if
	                        end if
	
                        %>

                        <tr style="background-color:<%=trColor%>;">
                            <td><%=oRec("kontonr")%></td>

                            <td>
		                    <%if print <> "j" then%>
		                    <a href="posteringer.asp?menu=erp&id=<%=oRec("id")%>&kontonr=<%=oRec("kontonr")%>&kid=<%=oRec("kunderid")%>&FM_startdato=<%=strStartDato %>&FM_slutdato=<%=strSlutDato %>" class=vmenu><%=oRec("navn")%> </a>
		                    <%else%>
		                    <%=oRec("navn")%>
		                    <%end if%></td>

                            <%expTxt = expTxt &"xx99123sy#z"&oRec("kontonr")&";"&oRec("navn")
		
		                    if oRec("debitkredit") <> 5 then
		                    expTxt = expTxt &";"&oRec("kkundenavn") 
		                    else
		                    expTxt = expTxt &";"
		                    end if
		                    %>

                            <td><i>
		                    <%if oRec("debitkredit") <> 5 then %>
		                    <%=oRec("kkundenavn")%>
		                    <%end if%>
                            </i></td>

                            <td><%		
		                        dbst = ""		
		                        'if oRec("type") = 2 then 'status
			
			                        if oRec("debitkredit") = 1 then
			                        dbst = "Kreditor"
			                        end if
			
			                        if oRec("debitkredit") = 2 then
			                        dbst = "Debitor"
			                        end if
			
	                            'else
		    
		                            if oRec("debitkredit") = 3 then
		                            dbst = "Samlekonto for oms."
		                            end if
		    
		                            if oRec("debitkredit") = 4 then
		                            dbst = "Samlekonto for udg."
		                            end if
		    
		                            if oRec("debitkredit") = 5 then
		                            dbst = "Overskrift"
		                            end if
		    
		                            if oRec("debitkredit") = 6 then
		                            dbst = ""
		                            end if
		    
		                         'end if%>
		
		                        <%=dbst&"&nbsp;"%>		
		                    </td>

                            <td>
		                        <%if oRec("debitkredit") <> 5 then %>
		                        <%
		                        Response.write oRec("momskodenavn")
		                        %>
		                        <%end if %>
		                    </td>

                            <td> <%
		                        tpe = ""
		
		                        select case oRec("type")
		                        case 1
		                        tpe = "Drift"
		                        case 2
		                        tpe = "Status"
		                        case else
		                        tpe = ""
		                        end select%>
		
		                        <%if oRec("debitkredit") <> 5 then %>
		                        <b><%=tpe %></b>
		                        <%end if %>
                            </td>

                            <td><%=oRec("nogletalkodenavn") %></td>

                            <%		
		                    sta = ""
		                    if print <> "j" then%>
			
                                <%if oRec("debitkredit") <> 5 then %>
			                    <td>
			                    <%if oRec("status") = 1 then
			                    sta = "<font color=LimeGreen>Åben</font>"
			                    staexp = "åben"
			                    else
			                    sta = "<font color=darkred>Lukket</font>"
			                    staexp = "lukket"
			                    end if%>
			
			                    <%=sta %></td>
                                <%else %>
                                <th>&nbsp</th>
		                        <%end if %>

		                    <%else%>
		                   <!-- <td>&nbsp;</td> -->
		                    <%end if%>

                            <td style="text-align:right;">
		                    <%
		
		                    saldo = oRec("saldo")
		
		                    if len(saldo) <> 0 OR saldo > 0 then
		                    saldo = saldo
		                    else
		                    saldo = 0
		                    end if
		
		
		                    totalSaldo = totalSaldo + (saldo)
		
		                    if oRec("debitkredit") <> 5 then
		                    expTxt = expTxt &";"&dbst&";"&oRec("momskodenavn")&";"&tpe&";"&oRec("nogletalkodenavn")&";"&staexp&";"&saldo 
		                    else
		                    expTxt = expTxt &";"&dbst&";;;;;"
		                    end if
		                    %>
		
	                        <%if oRec("debitkredit") <> 5 then %>
		                    <%=formatnumber(saldo,2) %>
		                    <%else %>
		                    &nbsp;
		                    <%end if%>
		                    </td>

                            <%if print <> "j" then%>
		                    <td style="text-align:center;"><a href="kontoplan.asp?menu=erp&func=red&id=<%=oRec("id")%>"><img src="../ill/blyant.gif" width="12" height="11" alt="Rediger konto" border="0"></a></td>
		                    <td style="text-align:center;">
		                        <%if level = 1 then 
		        
		        
		                       posFindes = 0
		                       '*** Tjekker om der findes posteringer på konto ***'
		                       strSQLpos = "SELECT p.kontonr FROM posteringer p WHERE p.kontonr = "& oRec("kontonr") & " GROUP BY p.kontonr" 
		                       oRec2.open strSQLpos, oConn, 3
		       
		                       if not oRec2.EOF then
		       
		                       posFindes = 1
		       
		                       end if
		                       oRec2.close
    		        
		                        if posFindes <> 1 then%>
		                        <a href="kontoplan.asp?menu=erp&func=slet&id=<%=oRec("id")%>"><span style="color:darkred;" class="fa fa-times"></span></a>
		                        <%else %>
		                        &nbsp;
		                        <%end if %>
		
		                        <%else %>
		                        &nbsp;
		                        <%end if %>
		                    </td>
		
		                    <%else%>
		                   <!-- <td>&nbsp;</td>
		                    <td>&nbsp;</td> -->
		                    <%end if%>

                        </tr>
                        <%
	                    x = 0
						
						
						if print = "j" then
						
						%>
						
						<tr>
                            <td colspan="100"><b><%=kontoplan_txt_042 %></b></td>
						</tr>
							
							
							<tr>
								<td><b><%=kontoplan_txt_038 %></b></td>
								<td><b><%=kontoplan_txt_049 %></b></td>
								<td><b><%=kontoplan_txt_039 %></b></td>
								<td><b><%=kontoplan_txt_050 %></b></td>
								<td><b><%=kontoplan_txt_043 %></b></td>
								<td><b><%=kontoplan_txt_040 %></b></td>
								<td><b><%=kontoplan_txt_041 %></b></td>

                                <td colspan="10">&nbsp</td>

							</tr>
							
							
						<%
						
						    end if


                            '**** Ved eksporter/print hentes posteringer ****
						    kassekladeorkonto = "(p.kontonr = "& oRec("kontonr") &" AND p.status = 1)"
						    momskri = " AND (kontoplan.kontonr = p.kontonr) "
						    useSQLKri = " p.id > 0 "
								
						    strSQL3 = "SELECT p.id, p.bilagstype, p.kontonr, "_
						    &" p.modkontonr, p.bilagsnr, p.beloeb, p.nettobeloeb, "_
						    &" p.moms, p.tekst, posteringsdato, p.status, "_
						    &" p.oprid, debitkredit, kontoplan.navn AS kontonavn FROM posteringer p, kontoplan "_
						    &" WHERE "& kassekladeorkonto &" AND "& periodefilter &" "& momskri &" AND "& useSQLKri &" ORDER BY posteringsdato DESC, id DESC"
						
						    'Response.Write strSQL3
						    'Response.flush
						    p = 0
						    oRec3.open strSQL3, oConn, 3
						    while not oRec3.EOF 
						
								
						    if print = "j" then%>

                                <tr>
									
						            <td><%=oRec3("posteringsdato")%></td>
						            <td><%=oRec3("bilagsnr")%></td>
						            <td><%=oRec3("tekst")%></td>
									
						            <td><%=oRec3("kontonr")%></td>
									
						            <td><%=oRec3("modkontonr")%></td>
									
						            <%if oRec3("beloeb") >= 0 then %>
						            <td><%=formatnumber(oRec3("beloeb"), 2)%></td>
						            <td>&nbsp;</td>
                                    <%else %>
                                        <td>&nbsp;</td>
                                        <td><%=formatnumber(oRec3("beloeb"), 2)%></td>
						            <%end if %>

                                    <td colspan="10">&nbsp</td>
									
								</tr>

                            <%end if 
                                
                                expTxt = expTxt &"xx99123sy#z"&oRec("kontonr")&";;;;;;;;;"&oRec3("posteringsdato")&";"&oRec3("bilagsnr")&";"&oRec3("tekst")
								
								
								
								if oRec3("beloeb") >= 0 then 
								expTxt = expTxt & ";"&formatnumber(oRec3("beloeb"), 2)&";;"
								else
								expTxt = expTxt & ";;"&formatnumber(oRec3("beloeb"), 2)&";"
								end if
								
								p = p + 1
								oRec3.movenext
								wend
								oRec3.close
								%>

                                <%if print = "j" then %>
                                <tr>
                                    <td colspan="100">&nbsp</td>
                                </tr>
                                <%end if %>


                        <%
                            oRec.movenext
	                        wend
	                        oRec.close

                        %>
                     
                    </tbody>

                </table>

                <div class="row">
                    <div class="col-lg-12"><%=kontoplan_txt_051 %>: <b><%=formatnumber(totalSaldo,2) %></b></div>
                </div>

                <%if print <> "j" then %>
                <div class="row">
                    <div class="col-lg-12">
                        <b><%=kontoplan_txt_044 %></b>
                    </div>
                </div>

                <form action="../timereg/kontoplan_eksport.asp" target="_blank" method="post" name=theForm2 onsubmit="BreakItUp2()">
                    <input type="hidden" name="datointerval" id="datointerval" value="<%=formatdatetime(strStartDatoSQL, 1) & " - " & formatdatetime(strSlutDatoSQL, 1)%>">
			        <input type="hidden" name="txt1" id="txt1" value="">
			        <input type="hidden" name="BigTextArea" id="BigTextArea" value="<%=expTxt%>">
			        <input type="hidden" name="txt20" id="txt20" value="">

                    <div class="row">
                        <div class="col-lg-3"><input id="Submit5" type="submit" value="<%=kontoplan_txt_045 %>" class="btn btn-sm" /></div>
                    </div>

                </form>

                <form action="kontoplan.asp?menu=kon&print=j&FM_soeg=<%=thiskri%>&FM_kd=<%=filterKD%>&FM_tp=<%=filterTP%>&visikkenul=<%=visikkenul%>&FM_nogletal=<%=filterNogletal%>&FM_status=<%=filterStatus %>" target="_blank" method="post">
                    <div class="row">
                        <div class="col-lg-3"><input id="Submit6" type="submit" value="<%=kontoplan_txt_046 %>" class="btn btn-sm" /></div>
                    </div>
                </form>

                <%else

                    Response.Write("<script language=""JavaScript"">window.print();</script>")

                %>

                <%end if %>

            </div>
        </div>
    </div>

    <%end select %>

</div>
</div>

<!--#include file="../inc/regular/footer_inc.asp"-->