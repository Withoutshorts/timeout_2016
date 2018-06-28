

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/mat_func.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<script src="js/materialer_jav2.js" type="text/javascript"></script>

<div class="wrapper">
<div class="content">

    <style>
        /* The Modal (background) */
    .modal {
        display: none; /* Hidden by default */
        position: fixed; /* Stay in place */
        z-index: 1; /* Sit on top */
        padding-top: 100px; /* Location of the box */
        left: 0;
        top: 0;
        width: 100%; /* Full width */
        height: 100%; /* Full height */
        overflow: auto; /* Enable scroll if needed */
        background-color: rgb(0,0,0); /* Fallback color */
        background-color: rgba(0,0,0,0.4); /* Black w/ opacity */
    }

    /* Modal Content */
    .modal-content {
        background-color: #fefefe;
        margin: auto;
        padding: 20px;
        border: 1px solid #888;
        width: 300px;
        height: 350px;
    }

    .picmodal:hover,
    .picmodal:focus {
    text-decoration: none;
    cursor: pointer;

    }
    </style>

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
		if len(request("id")) <> 0 then
		id = request("id")
		else
		id = 0
		end if
	end if
	
	menu = request("menu")
	
	session.lcid = 1030 'DK
	

	select case func
	case "flytlager"
	
	    nytLager = request("FM_nymgrp")
	
	    if len(request("FM_antal")) <> 0 then
	    antal = replace(request("FM_antal"), ",", ".")
	    else
	    antal = 0
	    end if

        call erDetInt(antal)
	    if isInt > 0 then

	    errortype = 76
	    call showError(errortype)
	    isInt = 0
	    response.end
	    end if

        response.Write "<br> id " & id & " antal " & antal & " nytlager " & nytLager
        call flytlager(id, antal, nytLager, 0)
	
	    response.redirect "materialer.asp?menu="&menu&"&id="&id&"&FM_sog="&request("FM_sog")



    case "slet"
	'*** Her spørges om det er ok at der slettes et materiale ***
        oskrift = "Materialer" 
        'slttxta = "Du er ved at <b>SLETTE</b> en kunde. Er dette korrekt?<br><br>Du vil samtidig <b><u>slette alle</u></b> tilhørende job, aktiviteter, ressourceforecast og fakturaer på denne kunde."
         slttxta = "Du er ved at" &" "& "<b>" & "Slette" &"</b> "& "et materiale. Er dette korrekt?"
 
        slttxtb = "" 
        slturl = "materialer.asp?menu=mat&func=sletok&id="&id&"&menu="&menu

        call sletcnf_2015(oskrift, slttxta, slttxtb, slturl)

    case "sletok"
	'*** Her slettes et materiale ***
	oConn.execute("DELETE FROM materialer WHERE id = "& id &"")
	Response.redirect "materialer.asp?menu=mat&shokselector=1"

    case "dbopr", "dbred"
	'*** Her indsættes en ny type i db ****
		if len(request("FM_navn")) = 0 then
		%>
		
		<%
		errortype = 8
		call showError(errortype)
		
		else
		
			%>
			<!--#include file="../timereg/inc/isint_func.asp"-->
			<%
			call erDetInt(request("FM_nr"))
			if isInt > 0 then
			sortorder = 1000
			else
			sortorder = request("FM_nr")
			end if
			
			isInt = 0
			
	        if len(trim(request("FM_nr"))) = 0 OR request("FM_nr") = "0"  then
		    %>
		   
		    <%
		    errortype = 73
		    call showError(errortype)
		    
		
		else
			
			
			call erDetInt(request("FM_antal"))
			if isInt > 0 then
			%>
			
			
			<%
			errortype = 54
			call showError(errortype)
			
			isInt = 0
		
		else
			
		
			call erDetInt(request("FM_minlager"))
			if isInt > 0 then
			%>
			
			
			<%
			errortype = 75
			call showError(errortype)
			
			isInt = 0
		
		else
			
			
			call erDetInt(request("FM_indkobspris"))
			if isInt > 0 then
			%>
			
			
			<%
			errortype = 55
			call showError(errortype)
			
			isInt = 0
		
		else
			
			call erDetInt(request("FM_salgspris"))
			if isInt > 0 then
			%>
			
			
			<%
			errortype = 56
			call showError(errortype)
			
			isInt = 0
		
		else
			
			call erDetInt(request("FM_ptid"))
			if isInt > 0 then
			%>
			
			
			<%
			errortype = 57
			call showError(errortype)
			
			isInt = 0
		
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
		
        strVarenr = SQLBless(request("FM_nr"))
		
		if len(request("FM_antal")) <> 0 then
		intAntal = replace(request("FM_antal"), ",", ".")
		else
		intAntal = 0
		end if
		
		if len(request("FM_indkobspris")) <> 0 then
		dblIndkobspris = replace(request("FM_indkobspris"), ",", ".")
		else
		dblIndkobspris = 0
		end if
		
		if len(request("FM_salgspris")) <> 0 then
		dblSalgspris = replace(request("FM_salgspris"), ",", ".")
		else
		dblSalgspris = 0
		end if
		
		'dtStart_dag = request("FM_start_dag")
		'dtStart_mrd = request("FM_start_mrd")
		'dtStart_aar = request("FM_start_aar")
		
		'sqlDato = dtStart_aar & "/" & dtStart_mrd & "/" & dtStart_dag  

        sqlYear = year(request("restDate"))
        sqlMonth = month(request("restDate"))
        sqlDay = day(request("restDate"))
        sqlRestDato = sqlYear & "/" & sqlMonth & "/" & sqlDay  

        if isDate(sqlRestDato) = true then
            sqlRestDato = sqlRestDato
        else
            sqlRestDato = year(now) & "/" & month(now) & "/" & day(now)
        end if
		
		intMatGrp = request("FM_mgrp")
		intPtid = request("FM_ptid")
		
        if len(trim(request("FM_pic_navn"))) <> 0 then
		intPic = request("FM_pic")
        else
        intPic = 1 'blank gif
        end if


                'response.write "len(trim(request(FM_pic_navn))): " & len(trim(request("FM_pic_navn")))
                'response.end
		
		if len(request("FM_enhed")) <> 0 then
		strEnhed = request("FM_enhed")
		else
		strEnhed = "Stk."
		end if
		
		if len(request("FM_ptid_arr")) <> 0 then
		ptid_arr = request("FM_ptid_arr")
		else
		ptid_arr = 0
		end if
		
		intLev_a = request("FM_lev_a")
		intLev_b = request("FM_lev_b")
		intSer_a = request("FM_ser_a")
		
		if len(trim(request("FM_minlager"))) <> 0 then
		intMinLager = replace(request("FM_minlager"), ",", ".")
		else
		intMinLager = 0
		end if
		
		strLokal = SQLBless(request("FM_lokal"))
		
		strBetegn = SQLBless(request("FM_betegn"))
		
		varenrFindes = 0
		strSQL = "SELECT id FROM materialer WHERE matgrp = "& intMatGrp &" AND varenr = '"& strVarenr &"' AND id <> "& id
		oRec.open strSQL, oConn, 3
		while not oRec.EOF
		varenrFindes = 1
		oRec.movenext
		wend
		
		oRec.close
		
		
		if varenrFindes = 1 then
		
			%>
			<!--#include file="../inc/regular/header_inc.asp"-->
			
			<%
			errortype = 74
			call showError(errortype)
		
		
		else
		
		if func = "dbopr" then
		
		strSQL = "INSERT INTO materialer "_
		&" (navn, editor, dato, varenr, antal, arrestordredato, salgspris, "_
		&" indkobspris, matgrp, ptid, ptid_arr, enhed, pic, leva, levb, sera, minlager, "_
		&" sortorder, betegnelse, lokation) VALUES "_
		&" ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', "_
		&" '"& strVarenr &"', "& intAntal &", '"& sqlRestDato &"', "& dblSalgspris &", "_
		&" "& dblIndkobspris &", "& intMatGrp &", "_
		&" "& intPtid &", "& ptid_arr &", '"& strEnhed &"', "& intPic &", "_
		&" "& intLev_a &", "& intLev_b &", "& intSer_a &", "& intMinLager &", "_
		&" "& sortorder &", '"& strBetegn &"', '"& strLokal &"')"
		
		'Response.write strSQL
		'Response.flush
		
		oConn.execute(strSQL)
		
		
		strSQL2 = "SELECT id FROM materialer WHERE id <> 0 ORDER BY id DESC"
		oRec2.open strSQL2, oConn, 3 
		if not oRec2.EOF then
			lastid = oRec2("id")
		end if
		oRec2.close 
		
		
		else
		
		strSQL = "UPDATE materialer SET navn ='"& strNavn &"', "_
		&" varenr = '"& strVarenr &"', editor = '" &strEditor &"', dato = '" & strDato &"', "_
		&" antal = "& intAntal &", arrestordredato =  '"& sqlRestDato &"', "_
		&" salgspris = "& dblSalgspris &", indkobspris = "& dblIndkobspris &", "_
		&" matgrp = "& intMatGrp &", ptid = "& intPtid &", ptid_arr = "& ptid_arr &", "_
		&" enhed = '"& strEnhed &"', pic = "& intPic &", "_
		&" leva = "& intLev_a &", levb = "& intLev_b &", sera = "& intSer_a &", "_
		&" minlager = "& intMinLager &", sortorder = "& sortorder &", "_
		&" betegnelse = '"& strBetegn &"', lokation = '"& strLokal &"'"_
		&" WHERE id = "&id
		
		lastid = id
		'Response.write strSQL
		'Response.flush
		
		oConn.execute(strSQL)
		
		strSQL = "UPDATE materiale_forbrug SET matnavn = '"& strNavn &"', matvarenr = '"& strVarenr &"' WHERE matid = " & id
		oConn.execute(strSQL)
		
		end if
		
	
		Response.redirect "materialer.asp?menu="&menu&"&shokselector=1&soglev="&request("soglev")&"&FM_mgrp="&request("FM_mgrp")&"&FM_sog="&request("FM_sog")&"&id="&lastid
		
		
		end if 'varenr findes i forvejen i gruppe.
		end if 'Navn
		end if 'Varenr
		end if 'salgsspris
		end if 'indkbspris
		end if 'minlager
		end if 'antal
		end if 'Produktionstid
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	varSubVal = "Opret" 
	varbroedkrumme = "Opret nyt materiale (vare)"
	dbfunc = "dbopr"
	sidst_redigeret = ""
	
	'strDag = day(now)
	'strMrd = month(now)
	'strAar = year(now)
	
	matGrpSel = request("FM_mgrp")
	

	intPtid = 0
	ptid_arr = 0
	
	strPic = 0
	intlev_a = 0
	intlev_b = 0
	intSer_a = 0
	intMinLager = 0
	strLokal = ""

    StrTdato = year(now) & "/" & month(now) & "/" & day(now)
	
	else
	
	strSQL = "SELECT navn, editor, dato, varenr, antal, "_
	& "ptid_arr, arrestordredato, salgspris, "_
	&" indkobspris, matgrp, ptid, enhed, pic, leva, levb, sera, minlager, betegnelse, lokation FROM materialer WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	strVarenr = oRec("varenr")
	dblKobsPris = oRec("indkobspris")
	dblSalgsPris = oRec("salgspris")
	intAntal = oRec("antal")
	StrTdato = oRec("arrestordredato")
	matGrpSel = oRec("matgrp")
	intPtid = oRec("ptid")
	ptid_arr = oRec("ptid_arr")
	strEnhed = oRec("enhed")
	strPic = oRec("pic")
	intlev_a = oRec("leva")
	intlev_b = oRec("levb")
	intSer_a = oRec("sera")
	intMinLager = oRec("minlager")
	strBetegn = oRec("betegnelse")
	strLokal = oRec("lokation")
	
	end if
	oRec.close
	
	sidst_redigeret = "Sidst opdateret den <b>"&formatdatetime(strDato, 1)&"</b> af <b>"&strEditor&"</b>"
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "Opdater" 
	end if
	
	
	strPicNavn = "../ill/blank.gif"
	
	if ptid_arr = 0 then
	chkArr_ptid1 = ""
	else
	chkArr_ptid1 = "CHECKED"
	end if
	%>
	
	<%if menu <> "0" then
	
	topD = "102"
	lDiv = "90"
	
	 %>
	 
	
    <!--
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">

	<%'call tsamainmenu(11)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	'if showonejob <> 1 then
	'	call mattopmenu()
	'end if
	%>
	</div>
        -->
	<%
        
        call menu_2014()
        else 
	
	topD = "20"
	lDiv = "20"
	
	%>
	 
	<%end if %>
	
	 <!--#include file="../timereg/inc/dato2.asp"-->         

    <!-------------------------------Sideindhold------------------------------------->
    <div class="container">
        <div class="portlet">
            <h3 class="portlet-title"><u>Materialer - <%=varbroedkrumme%> rest <%=StrTdato %></u></h3>

            <div class="portlet-body">


                <form action="materialer.asp?menu=<%=menu%>&func=<%=dbfunc%>" method="post">
                    <input type="hidden" name="id" value="<%=id%>">
	                <input type="hidden" name="FM_sog" value="<%=request("FM_sog")%>">

                    <div class="row">
                    <div class="col-lg-10">&nbsp</div>
                    <div class="col-lg-2 pad-b10">
                    <button type="submit" class="btn btn-success btn-sm pull-right"><b><%=medarb_txt_020 %></b></button>
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
                            <div class="col-lg-2">Navn:&nbsp<span style="color:red;">*</span></div>
                            <div class="col-lg-3"><input type="text" name="FM_navn" value="<%=strNavn%>" class="form-control input-small"></div>

                            <div class="col-lg-2">Materiale / vare nr.&nbsp<span style="color:red;">*</span>:</div>
                            <div class="col-lg-3"><input type="text" name="FM_nr" value="<%=strVarenr%>" class="form-control input-small" ></div>
                        </div>

                        <div class="row">
                            <div class="col-lg-1"></div>
                            <div class="col-lg-2">Lokation:</div>
                            <div class="col-lg-3"><input type="text" name="FM_lokal" value="<%=strLokal%>" maxlength="20" class="form-control input-small" ></div> 

                            <div class="col-lg-2">Beskrivelse/note</div>
                            <div class="col-lg-3"><textarea name="FM_betegn" id="FM_betegn" cols="35" rows="4" class="form-control input-small"><%=strBetegn%></textarea></div>
                        </div>

                    </div>


                    <!-- Indstillinger -->
                    <div class="panel-group accordion-panel" id="accordion-paneled">
                        <!-- PersonData -->
                        <div class="panel panel-default">
                            <div class="panel-heading"><h4 class="panel-title"><a class="accordion-toggle" data-toggle="collapse" data-target="#collapseOne">Indstillinger </a></h4></div>
                           
                            <div id="collapseOne" class="panel-collapse collapse">
                                <div class="panel-body">

                                
                                    <%
		                                strSQL = "select id, filnavn FROM filer WHERE id = " & strPic
		                                oRec.open strSQL, oConn, 3 
		                                if not oRec.EOF then
			                                strPicNavn = "../inc/upload/"&lto&"/"&oRec("filnavn")
			                                strPicNavn_only = oRec("filnavn")
		                                end if
		                                oRec.close 
		                            %>
                                    <div class="row">
                                        <div class="col-lg-1"></div>

                                        <div class="col-lg-2"><a href="#" onclick="Javascript:window.open('../timereg/upload.asp?FM_filtype=6&func=opret&type=mat&id=<%=id%>&kundeid=0&jobid=0&FM_folderid=-10&FM_adg_alle=1', '', 'width=450,height=675,resizable=yes,scrollbars=yes')" class=vmenu>Billede upload</a></div>
                                       
                                        <div class="col-lg-2">
                                            <input type="text" class="form-control input-small" name="FM_pic_navn" id="FM_pic_navn" value="<%=strPicNavn_only%>" readonly><br>
		                                    <input type="hidden" name="FM_pic" id="FM_pic" value="<%=strPic %>">                                       
                                        </div>

                                        <%if strPic <> 1 then %>
                                        <div class="col-lg-1"><span id="modal_1" class="fa fa-image picmodal" style="font-size:125%"></span></div>
                                       
                                        <div id="myModal_1" class="modal">
                                            <div class="modal-content">
                                                <img src="<%=strPicNavn %>" alt='' border='0' width="100%" height="100%">
                                            </div>
                                        </div>

                                         <%end if %>

                                        <div class="col-lg-2">Materialegruppe (avance%):</div>
                                        <div class="col-lg-3">
                                            <select name="FM_mgrp" id="FM_mgrp" class="form-control input-small">
		                                        <option value="0">Ingen</option>
		                                        <%
		                                        strSQL = "SELECT id, navn, av FROM materiale_grp ORDER BY navn"
		                                        oRec.open strSQL, oConn, 3 
		                                        while not oRec.EOF 
		
		                                        if cint(matGrpSel) = oRec("id") then
		                                        selThis = "SELECTED"
		                                        else
		                                        selThis = ""
		                                        end if%>
		                                        <option value="<%=oRec("id")%>" <%=selThis%>><%=oRec("navn")%> (<%=oRec("av") %>%)</option>
		                                        <%
		                                        oRec.movenext
		                                        wend
		                                        oRec.close %>
		                                    </select>
                                        </div>
                                    </div>

                                    <div class="row">
                                        <div class="col-lg-1"></div>
                                        <div class="col-lg-2">Antal på lager (hel-tal):</div>
                                        <div class="col-lg-1"><input type="text" name="FM_antal" value="<%=intAntal%>" class="form-control input-small" ></div>

                                        <div class="col-lg-2"></div>

                                        <div class="col-lg-2">Minimuslager:</div>
                                        <div class="col-lg-1"><input type="text" name="FM_minlager" value="<%=intMinLager%>" class="form-control input-small"></div>
                                        <div class="col-lg-2">(hel-tal, -1 = Lagerstatus udsendes ikke)</div>
                                    </div>

                                    <div class="row">
                                        <div class="col-lg-1"></div>
                                        <div class="col-lg-2">Enhed:</div>
                                        <div class="col-lg-1"><input type="text" name="FM_enhed" value="<%=strEnhed%>" class="form-control input-small" /></div>
                                        <div class="col-lg-2">stk (standard), m, kg, etc..</div>
                                    </div>

                                    <div class="row">
                                        <div class="col-lg-1"></div>
                                        <div class="col-lg-2">Indkøbspris</div>
                                        <div class="col-lg-1"><input type="text" name="FM_indkobspris" value="<%=dblKobsPris%>" class="form-control input-small" ></div>
                                        <div class="col-lg-2">Kr.</div>

                                        <div class="col-lg-2">Salgspris:</div>
                                        <div class="col-lg-2"><input type="text" name="FM_salgspris" value="<%=dblSalgsPris%>" class="form-control input-small" /></div>
                                        <div class="col-lg-2">kr.</div>
                                    </div>

                                    <div class="row">
                                        <div class="col-lg-1"></div>
                                        <div class="col-lg-2"><input type="checkbox" name="FM_ptid_arr" value="1" <%=chkArr_ptid1%>> Brug Restordre dato:</div>
                                        <div class="col-lg-3">
                                            <div class='input-group date' id='datepicker_stdato'> 
                                            <input type="text" class="form-control input-small" name="restDate" value="<%=StrTdato %>" placeholder="dd-mm-yyyy" />
                                            <span class="input-group-addon input-small">
                                                    <span class="fa fa-calendar">
                                                    </span>
                                                </span>
                                            </div>
                                        </div>

                                        <div class="col-lg-2">Produktionstid:</div>
                                        <div class="col-lg-1"><input type="text" name="FM_ptid" value="<%=intPtid%>" class="form-control input-small"  ></div>
                                        <div class="col-lg-2">(Antal hele dage efter restordredato eller dagsdato)</div>
                                    </div>

                                    <div class="row">
                                        <div class="col-lg-1"></div>

                                        <div class="col-lg-2">Leverandør A:</div>
                                        <div class="col-lg-3">
                                            <select name="FM_lev_a" id="FM_lev_a" class="form-control input-small">
			                                    <option value="0" SELECTED>Ingen</option>
			                                    <%strSQL = "SELECT id, navn, levnr FROM leverand WHERE id <> 0 AND type = 1 OR type = 3 ORDER BY navn"
			                                    oRec.open strSQL, oConn, 3 
			                                    while not oRec.EOF 
			
			                                    if cint(intLev_a) = oRec("id") then
			                                    selLevA = "SELECTED"
			                                    else
			                                    selLevA = ""
			                                    end if
			                                    %>
			                                    <option value="<%=oRec("id")%>" <%=selLevA%>><%=oRec("navn")%> (<%=oRec("levnr") %>)</option>
			                                    <%
			                                    oRec.movenext
			                                    wend
			                                    oRec.close %>
			                                </select>
                                        </div>

                                        <div class="col-lg-2">Leverandør B</div>
                                        <div class="col-lg-3">
                                            <select name="FM_lev_b" id="FM_lev_b" class="form-control input-small">
			                                    <option value="0" SELECTED>Ingen</option>
			                                    <%strSQL = "SELECT id, navn, levnr FROM leverand WHERE id <> 0 AND type = 1 OR type = 3 ORDER BY navn"
			                                    oRec.open strSQL, oConn, 3 
			                                    while not oRec.EOF 
			
			                                    if cint(intLev_b) = oRec("id") then
			                                    selLevB = "SELECTED"
			                                    else
			                                    selLevB = ""
			                                    end if
			                                    %>
			                                    <option value="<%=oRec("id")%>" <%=selLevB%>><%=oRec("navn")%> (<%=oRec("levnr") %>)</option>
			                                    <%
			                                    oRec.movenext
			                                    wend
			                                    oRec.close %>
			                                </select>
                                        </div>
                                    </div>

                                    <div class="row">
                                        <div class="col-lg-1"></div>
                                        <div class="col-lg-2">Servicepartner</div>
                                        <div class="col-lg-3">
                                            <select name="FM_ser_a" id="FM_ser_a" class="form-control input-small">
			                                    <option value="0" SELECTED>Ingen</option>
			                                    <%strSQL = "SELECT id, navn, levnr FROM leverand WHERE id <> 0 AND type = 2 OR type = 3 ORDER BY navn"
			                                    oRec.open strSQL, oConn, 3 
			                                    while not oRec.EOF 
			                                    if cint(intSer_a) = oRec("id") then
			                                    selSerA = "SELECTED"
			                                    else
			                                    selSerA = ""
			                                    end if
			                                    %>
			                                    <option value="<%=oRec("id")%>" <%=selSerA%>><%=oRec("navn")%> (<%=oRec("levnr") %>)</option>
			                                    <%
			                                    oRec.movenext
			                                    wend
			                                    oRec.close %>
			                                </select>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>

                </form>

            </div>

 
        </div>
    </div>




    <%case else



    if menu <> "0" then
	
	topD = "102"
	lDiv = "90"
	
	 %>
	 
	
<!--
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
		<%'call tsamainmenu(11)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	'if showonejob <> 1 then
	'	call mattopmenu()
	'end if
	%>
	</div>
    -->

	<%
        
        call menu_2014()
                
    else 
	
	level = session("rettigheder")
	topD = "20"
	lDiv = "20"
	
	%>
	
	
	<%end if %>
	
	<script language="javascript">
	<!--
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	
	//function visgrp(){
	//grp = document.getElementById("FM_mgrp").value
	//memu = document.getElementById("menuuse").value
	////limitnr = document.getElementById("FM_limitvarenr").value
	//window.location.href = "materialer.asp?menu="+memu+"&FM_mgrp="+grp
	//}	
	
	function rensfelt(felt){
	document.getElementById(felt).value = ""
	}
	
	

	
	function BreakItUp2()
	{
	  //Set the limit for field size.
	  var FormLimit = 102399
	
	  //Get the value of the large input object.
	  var TempVar = new String
	  TempVar = document.theForm2.BigTextArea.value
	
	  //If the length of the object is greater than the limit, break it
	  //into multiple objects.
	  if (TempVar.length > FormLimit)
	  {
	    document.theForm2.BigTextArea.value = TempVar.substr(0, FormLimit)
	    TempVar = TempVar.substr(FormLimit)
	
	    while (TempVar.length > 0)
	    {
	      var objTEXTAREA = document.createElement("TEXTAREA")
	      objTEXTAREA.name = "BigTextArea"
	      objTEXTAREA.value = TempVar.substr(0, FormLimit)
	      document.theForm2.appendChild(objTEXTAREA)
	
	      TempVar = TempVar.substr(FormLimit)
	    }
	  }
	}
	
	//-->
	</script>
	<%
	        
	        
	        if len(request("FM_mgrp")) <> 0 then
        	    
        	    matGrpSel = request("FM_mgrp")
		        Response.cookies("mat")("matgrp") = matGrpSel
		       
		    else
        		
		        if len(Request.cookies("mat")("matgrp")) <> 0 then
		        matGrpSel = Request.cookies("mat")("matgrp")
		        else
		        matGrpSel = 0
		        end if
        		
	        end if
	        
	        if cint(matGrpSel) <> 0 then
            matgrpSQLKri = " m.matgrp = "& matGrpSel & " AND " 
            odrbyMG = ""
            else
            matgrpSQLKri = " m.matgrp <> -1 AND " 
            odrbyMG = "m.matgrp, "
            end if
	        
	        
	
	        stLimit = 0
	        stLimitfundet = 0
	        
	        
	           if len(trim(request("FM_mgrp"))) <> 0 then
        				
				       stLimitKri = trim(request("FM_limitvarenr"))
				       response.cookies("mat")("stvarenr") = stLimitKri
        		else
				       if request.cookies("mat")("stvarenr") <> "0" then
				       stLimitKri = request.cookies("mat")("stvarenr")
        			   else
        			   stLimitKri = ""
        			   end if	
		        end if
	        
	                    if stLimitKri <> "" then 
	                    
	                    strSQL = "SELECT m.varenr FROM materialer m "_
				        &" WHERE m.varenr >= '"& stLimitKri &"'"
        				
				        oRec.open strSQL, oConn, 3
				        if not oRec.EOF then
				        stLimit = oRec("varenr") 
				        stLimitfundet = 1
				       
				        end if
				        oRec.close
				        
				       
				       end if
				        
				        if cint(stLimitfundet) = 1 then
				        useStLimitSQLkri = " AND m.varenr >= '"& stLimit & "'"
				        odrbySL = "m.varenr, "
				        else
				        useStLimitSQLkri = ""
				        odrbySL = ""
				        end if
	            
	            
	        
	        
	         
        	
        	
	        call erDetInt(request("FM_limit"))
	        if len(trim(request("FM_limit"))) <> 0 then
        	        
	                if isInt > 0 OR request("FM_limit") > 1000 then
	                useLimit = 1000
	                else
	                useLimit = request("FM_limit")
	                end if
        	        
	                Response.Cookies("mat")("mat_uselimit") = useLimit
	                
	        else
        	
	            if Request.Cookies("mat")("mat_uselimit") <> "" then
	            useLimit = Request.Cookies("mat")("mat_uselimit")
	            else
	            useLimit = 100
	            end if
        	    
	        end if 
	        isInt = 0
	
	
            if len(request("FM_mgrp")) <> 0 then
			sogVal = trim(request("FM_sog"))
			response.cookies("mat")("sogval") = sogVal
			odrbyMN = "m.navn"
			else
			   
			    if request.cookies("mat")("sogval") <> "0" then
			    sogVal = request.cookies("mat")("sogval")
			    odrbyMN = "m.navn"
			    else
			    sogVal = ""
			    odrbyMN = ""
			    end if
			    
			end if
			
			orderBYSQL = odrbySL & odrbyMG & odrbyMN 
			
			'*** fra Print **
			'if instr(sogVal, "99ogprocent99") <> 0 then
			'sogVal = replace(sogVal, "99ogprocent99", "%")
			'else
			'sogVal = sogVal
			'end if
			
			'if instr(sogVal, "%") <> 0 then
			'sogValPrint = replace(sogVal, "%", "99ogprocent99")
			'else
			'sogValPrint = sogVal
			'end if
			
			if len(request("FM_minlager")) <> 0 THEN
			minlager = 1
			minLagerCHK = "CHECKED"
			else
			minlager = 0
			minLagerCHK = ""
			end if
			
			
			 if len(request("soglev")) <> 0 then 
                sogKrilev = request("soglev")
                Response.cookies("mat")("soglev") = sogKrilev
                
                if request("soglev") <> "0" then
                sogLevSQLkri = " (m.leva = " & sogKrilev & " OR m.levb = " & sogKrilev &") AND "
                else
                sogLevSQLkri = " m.leva <> -1 AND "
                end if
                
                else
                    
                    if request.cookies("mat")("soglev") <> "0" AND request.cookies("mat")("soglev") <> "" then
                    sogKrilev = request.cookies("mat")("soglev")
                    sogLevSQLkri = " (m.leva = " & sogKrilev & " OR m.levb = " & sogKrilev &") AND "
                    else
                    sogKrilev = 0
                    sogLevSQLkri = " m.leva <> -1 AND "
                    end if
                    
                    
            end if
            
            
            
            Response.Cookies("mat").expires = date + 45
            
            
            
           
			
	
	%>
	<!-------------------------------Sideindhold------------------------------------->
    <div class="container">
        <div class="portlet">
            <h3 class="portlet-title"><u>Materialer/lager</u></h3>

            <div class="portlet-body">

                <% 
                nLnk = "materialer.asp?menu="&menu&"&func=opret&soglev="&sogKrilev&"&FM_mgrp="&matGrpSel&"&FM_sog="&sogVal
                %>

                <%if level <= 2 OR level = 6 then %>
                    <form action="<%=nLnk %>" method="post" target="_blank"> 
                        <div class="row">
                            <div class="col-lg-10">&nbsp;</div>
                            <div class="col-lg-2">
                                <button class="btn btn-sm btn-success pull-right"><b>Opret materiale +</b></button><br />&nbsp
                            </div>
                        </div>
                    </form>    
                <%end if %>

                <form action="materialer.asp?menu=mat" method="post">
                    <div class="well">
                        <h4 class="panel-title-well"><%=medarb_txt_098 %></h4>

                        <div class="row">
                            <div class="col-lg-7">Søg på navn eller varenr. Søg på flere materialer ved at opdele søgningen med ";"</div>
                        </div>
                        <div class="row">
                            <div class="col-lg-3"><input type="text" name="FM_sog" id="FM_sog" value="<%=sogVal%>" class="form-control input-small" placeholder="% Wildcard"></div>
                        </div>
                        <br />

                        <div class="row">
                            <div class="col-lg-3">Materialegruppe:</div>
                            <div class="col-lg-3">Leverandør:</div>
                            <div class="col-lg-2">Vis fra Mat./vare nr.:</div>
                            <div class="col-lg-2">og de næste</div>
                        </div>
                        <div class="row">
                            <div class="col-lg-3">
                                <select name="FM_mgrp" id="FM_mgrp" class="form-control input-small">
		                            <option value="0">Alle</option>
		                            <%
		                            strSQL = "SELECT id, navn, av FROM materiale_grp ORDER BY navn"
		                            oRec.open strSQL, oConn, 3 
		                            while not oRec.EOF 
		
		                            if cint(matGrpSel) = oRec("id") then
		                            selThis = "SELECTED"
		                            else
		                            selThis = ""
		                            end if%>
		                            <option value="<%=oRec("id")%>" <%=selThis%>><%=oRec("navn")%>
		                            <%if level <= 2 OR level = 6 then %>
		                            &nbsp;(<%=oRec("av") %>%)
		                            <%end if %>
		                            </option>
		                            <%
		                            oRec.movenext
		                            wend
		                            oRec.close %>
		                        </select>
                            </div>

                            <div class="col-lg-3">
                                <select id="soglev" name="soglev" class="form-control input-small">
                                    <option value="0" SELECTED>Alle</option>
                                    <% 
	     
                                           strSQL3 = "SELECT id, navn FROM leverand WHERE id <> 0 ORDER BY navn"
           
                                           oRec3.open strSQL3, oConn, 3 
                                           while not oRec3.EOF 
                   
                                                if cint(sogKrilev) = oRec3("id") then
                                                slevSEL = "SELECTED"
                                                else
                                                slevSEL = ""
                                                end if
                   
                                           %>
                                           <option value="<%=oRec3("id") %>" <%=slevSEL %>><%=oRec3("navn") %></option>
                                           <%
                                           oRec3.movenext
                                           wend
                                           oRec3.close
                                    %> 
                                </select>  
                            </div>

                            <div class="col-lg-2"><input type="text" name="FM_limitvarenr" id="FM_limitvarenr" value="<%=stLimitKri%>" class="form-control input-small" /></div>
                            <div class="col-lg-2"><input id="FM_limit" name="FM_limit" type="text" size="3" value="<%=useLimit%>" class="form-control input-small" /></div>
                            <div class="col-gl-1">(maks 100)</div>
                        
                        </div>

                        <div class="row">
                            <div class="col-lg-4"><input id="FM_minlager" name="FM_minlager" type="checkbox" <%=minLagerCHK %> /> Vis kun materialer hvor antal på lager er mindre end minimuslager.</div>
                        </div>

                        <div class="row">
                            <div class="col-lg-12">
                                <button id="Submit1" class="btn btn-sm btn-secondary pull-right"><b>Vis materialer >></b></button>
                            </div>
                        </div>

                    </div>
                </form>

                <% 
		        if level <= 2 OR level = 6 then 
		        cspan = 13
		        else
		        cspan = 11
		        end if
		
		
	            tTop = 50
	            tLeft = 0
	            tWdth = 1184
	
	
	            'call tableDiv(tTop,tLeft,tWdth)
		
		        %>

                <table class="table dataTable table-striped table-bordered table-hover ui-datatable">

                    <thead>
                        <tr>

                            <th>Gruppe
                                <%if level <= 2 OR level = 6 then %>
		                        <br />(Ava. %)
		                        <%end if %>
                            </th>
                            <th>Lokation</th>
                            <th>Billede</th>
                            <th>Navn <br /> Betegnelse / Note</th>
                            <th>Varenr.</th>
                            <th>Opd. dato</th>
		                    <th>På lager</b> (ordre)</th>
		                    <th>Min. lager</th>
		                    <%if level <= 2 OR level = 6 then %>
		                    <th>Indkøbs-pris</th>
		                    <th>Salgs-pris</th>
		                    <%end if%>
		                    <th>Pri. lev.</th>
		
		                    <th>Slet?</th>
		                    <th colspan="4">Flyt antal til.. (ny gruppe)</th>                            
                        </tr>

                        <%
                            strExport = "Gruppe;Lokation;Varenr;Navn;Betegnelse;Antal på lager;Enhed;Min. Lager"
	
	                        if level = 1 then
	                        strExport = strExport &";Inkøbspris;Salgspris"
	                        end if
                        %>
                    </thead>

                    <tbody>
                        <%
                             sogKri = " m.navn <> 'xffgd123!fsd999'"
	                        if len(sogVal) <> 0 then
		
		                        sogKri = ""
		                        sogeKriUse = split(sogVal, ";")
		                        for b = 0 to UBOUND(sogeKriUse)
                                sogKri = sogKri &" m.navn LIKE '%"& sogeKriUse(b) &"%' OR m.varenr = '%"& sogeKriUse(b) &"%' OR m.betegnelse LIKE '"& sogeKriUse(b) &"%' OR "
		                        next
		
		                        lensogKri = len(sogKri)
		                        leftsogKri = left(sogKri, lensogKri - 3)
		                        sogKri = leftsogKri
		
	                        end if
	
	
	                        if cint(minlager) = 1 then
	                        lagerSQLkri = " (m.antal < minlager) AND "
	                        else
	                        lagerSQLkri = " (m.antal <> 999999999) AND "
	                        end if
	
	                        '**** Main SQL ***'
	                        strSQL = "SELECT m.id, m.navn, m.varenr, m.antal, m.dato, m.betegnelse, "_
	                        &" l.navn AS levnavn, l.id AS lid, mg.navn AS gnavn, m.enhed, leva, minlager, "_
	                        &" mg.id AS grpid, m.indkobspris, m.salgspris, SUM(mol.antal) AS restordre, m.lokation, mg.av, pic FROM materialer m "_
	                        &" LEFT JOIN materiale_grp mg ON (mg.id = matgrp) "_
	                        &" LEFT JOIN leverand l ON (l.id = leva) "_
	                        &" LEFT JOIN materiale_ordrer_linier mol ON (mol.varenr = m.varenr AND mol.gruppe = mg.navn AND mol.status = 0) "_
	                        &" WHERE "& lagerSQLkri &""& matgrpSQLKri &""& sogLevSQLkri &" ("& sogKri &") "& useStLimitSQLkri &""_
	                        &" GROUP BY m.id ORDER BY "& orderBYSQL &" LIMIT "& useLimit
	
	                        'Response.Write strSQL
	                        'Response.flush
	
	                        x = 0
	
	                        strMatids = "0"
	
	                        oRec.open strSQL, oConn, 3
	                        while not oRec.EOF 
	
	                        strMatids = strMatids &","& oRec("id")
	
	                        if cdbl(id) = oRec("id") then
	                        bgthis = "#ffff99"
	                        else
	                            select case right(x, 1)
	                            case 0,2,4,6,8
	                            bgthis = "#eff3ff"
	                            case else
	                            bgthis = "#ffffff"
	                            end select
	                        end if
                        %>

                        <tr bgcolor="<%=bgthis%>">

                            <td><b><%=oRec("gnavn")%></b>
                            <%if (level <= 2 OR level = 6) AND len(trim(oRec("av")) <> 0) then %> 
                            &nbsp;(<%=oRec("av") %>%)
                            <%end if %></td>

                            <td><%=oRec("lokation") %></td>

                            <td style="text-align:center">
                                <%

                                strPicNavn = ""
                                    strPicNavn_only = ""
		                        strSQL = "select id, filnavn FROM filer WHERE id = " & oRec("pic")
		                        oRec5.open strSQL, oConn, 3 
		                        if not oRec5.EOF then
			                        strPicNavn = "../inc/upload/"&lto&"/"&oRec5("filnavn")
			                        strPicNavn_only = oRec5("filnavn")
		                        end if
		                        oRec5.close 

                               if cint(oRec("pic")) <> 1 then %>
                                <!--<%=strPicNavn_only %>-->

                                <div class="col-lg-1"><span id="modal_<%=oRec("id") %>" class="fa fa-image picmodal" style="font-size:125%"></span></div>

                                <div id="myModal_<%=oRec("id") %>" class="modal">
                                    <div class="modal-content">
                                        <img src="<%=strPicNavn %>" alt='' border='0' width="100%" height="100%">
                                    </div>
                                </div>

                                <%else %>
                                &nbsp;
	                            <%end if %>
                            </td>

                            <td>
		                        <%if level <= 2 OR level = 6 then %>
		                        <a href="materialer.asp?menu=<%=menu%>&func=red&id=<%=oRec("id")%>&soglev=<%=sogKrilev%>&FM_mgrp=<%=matGrpSel%>&FM_sog=<%=sogVal%>" class=vmenu><%=oRec("navn")%> </a>
		                        <%else %>
		                        <b><%=oRec("navn")%> </b>
		                        <%end if %>
		
		                        <%if len(trim(oRec("betegnelse"))) <> 0 then %>
		                        <br /><i>
		                            <%if len(oRec("betegnelse")) > 60 then %>
		                            <%=left(oRec("betegnelse"), 60) & ".."%>
		                            <%else %>
		                            <%=oRec("betegnelse") %>
		                            <%end if %></i>
		                        <%end if %>
		                    </td>

                            <td><b><%=oRec("varenr")%></b></td>

                            <td align=right style="padding-right:5px; white-space:nowrap;">
		                    <%if len(trim(oRec("dato"))) <> 0 then %>
 		                    <%=formatdatetime(oRec("dato"), 2)%>
 		                    <% end if%></td>

                            <td><b><%=oRec("antal")%>&nbsp;<%=oRec("enhed")%></b>
		                        <%if oRec("antal") < oRec("minlager") then
		                        %>&nbsp;<font color="darkred"><b>!</b></font><%
		                        end if%>
		
		                        <%if oRec("restordre") > 0 then %>
		                        (<%=oRec("restordre")%>)
		
		                        <%end if %>
		                    </td>

                            <td><%=oRec("minlager")%>&nbsp;<%=oRec("enhed")%></td>

                            <%if level <= 2 OR level = 6 then %>
                
                            <%if lto = "dencker" OR lto = "xintranet - local" then
                            dicim = 3
                            else
                            dicim = 2
                            end if %>

		                    <td><%=formatnumber(oRec("indkobspris"),dicim) %></td>
		                    <td><%=formatnumber(oRec("salgspris"),dicim) %></td>
		                    <%end if%>

                            <%if len(oRec("levnavn")) > 30 then 
		                    levnavn = left(oRec("levnavn"), 30) & ".."
		                    else
		                    levnavn = oRec("levnavn") 
		                    end if%>
		
		                    <td>
		                    <a href="../timereg/leverand.asp?menu=<%=menu%>&func=red&id=<%=oRec("lid")%>&rdir=mat" class=vmenu>
		                    <%=left(levnavn, 10)%></a>
		                    </td>

                            <td style="text-align:center">
		                    <%if level <= 2 OR level = 6 then %>
		                    <a href="materialer.asp?menu=<%=menu%>&func=slet&id=<%=oRec("id")%>&soglev=<%=sogKrilev%>&FM_mgrp=<%=matGrpSel%>&FM_sog=<%=sogVal%>"><span style="color:darkred;" class="fa fa-times"></span></a>
		                    <%end if%>
		                    &nbsp;
		                    </td>

                            <form method=post action="materialer.asp?menu=<%=menu%>&func=flytlager&id=<%=oRec("id")%>">
		                    <input type="hidden" name="FM_sog" id="<%=oRec("id")%>" value="<%=sogValPrint%>"> 

		                    <td style="border-right:none"><input type="text" style="width:30px;" name="FM_antal" id="FM_antal" value="" class="form-control input-small"></td> 
                            <td style="border-right:none">til:</td>
                            <td style="border-right:none">
                                <select name="FM_nymgrp" id="FM_nymgrp" style="width:125px;" class="form-control input-small">
		                            <%
		                            strSQL2 = "SELECT id, navn, av FROM materiale_grp ORDER BY navn"
		                            oRec2.open strSQL2, oConn, 3 
		                            while not oRec2.EOF 
		
		                            if cint(oRec2("id")) = oRec("grpid") then
		                            selThis2 = "SELECTED"
		                            else
		                            selThis2 = ""
		                            end if%>
		                            <option value="<%=oRec2("id")%>" <%=selThis2%>><%=oRec2("navn")%>
		                            <%if level <= 2 OR level = 6 then %> 
		                            &nbsp;(<%=oRec2("av") %>%)
		                            <%end if %>
		                            </option>
		                            <%
		                            oRec2.movenext
		                            wend
		                            oRec2.close %>
		                        </select>
                            </td>
		                    <td><input type="image" src="../ill/soeg-knap.gif" style="padding-top:5px;"></td>
                            </form>

                        </tr>

                        <%
	
	                    if len(oRec("navn")) <> 0 then
	                    navn = replace(oRec("navn"), Chr(34), "''")
	                    else
	                    navn = ""
	                    end if
	
	                    if len(oRec("betegnelse")) <> 0 then
	                    betegnelse = replace(oRec("betegnelse"), Chr(34), "''")
	                    else
	                    betegnelse = ""
	                    end if
	
	                    strExport = strExport &"xx99123sy#z"&oRec("gnavn")&";"&oRec("lokation")&";"&oRec("varenr")&";"& navn &";"& betegnelse &";"&oRec("antal")&";"&oRec("enhed")&";"&oRec("minlager")
	
	                    if level = 1 then
	                    strExport = strExport &";"&formatnumber(oRec("indkobspris"),decim)&";"&formatnumber(oRec("salgspris"),decim)
	                    end if
	
	
	                    bestilantal = (oRec("antal") - oRec("minlager")) 
	                    if bestilantal < 0 then
	                    bestilantal = -(bestilantal)
	                    else
	                    bestilantal = 0
	                    end if
	
	                    strAntalFM = strAntalFM &"<input id='FM_antal_b_"&oRec("id")&"' name='FM_antal_b_"&oRec("id")&"' type=hidden value='"&bestilantal&"' />" 
	
	
	                    x = x + 1
	                    oRec.movenext
	                    wend
	
	                    if x = 0 then%>
                        <tr>
                            <td colspan="20" style="text-align:center">Der blev ikke fundet nogen materialer der matcher de valgte søgekriterier.</td>
                        </tr>
                        <%end if %>

                    </tbody>


                </table>

                <br /><br />

                <form action="materialer_ordrer.asp?func=opr" method="post" target="_blank"> 
                    <div class="row">
                        <div class="col-lg-10">&nbsp;</div>
                        <div class="col-lg-2">
                            <button class="btn btn-sm btn-success pull-right"><b>Opret mat. ordre +</b></button><br />&nbsp
                        </div>
                    </div>
                </form>

                <div class="row">
                    <div class="col-lg-2"><b>Funktioner</b></div>
                </div>
                <div class="row">
                    <div class="col-lg-2">
                        <form action="../timereg/materialer_eksport.asp" method="post" name=theForm2 onsubmit="BreakItUp2()"> <!--  -->
			                <input type="hidden" name="txt1" id="txt1" value="">
			                <input type="hidden" name="BigTextArea" id="BigTextArea" value="<%=strExport%>">
			                <input type="hidden" name="txt20" id="txt20" value="">

                            <input type="submit" value="Eksport til csv." class="btn btn-sm" /><br />
                        </form>
                    </div>
                </div>

            </div>
        </div>
    </div>






<%end select %>



<script type="text/javascript">


$(".picmodal").click(function() {

    var modalid = this.id
    var idlngt = modalid.length
    var idtrim = modalid.slice(6, idlngt)

    //var modalidtxt = $("#myModal_" + idtrim);
    var modal = document.getElementById('myModal_' + idtrim);

    modal.style.display = "block";

    window.onclick = function (event) {
        if (event.target == modal) {
            modal.style.display = "none";
        }
    }

});


</script>


</div>
</div>


<!--#include file="../inc/regular/footer_inc.asp"-->