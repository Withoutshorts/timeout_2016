

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->





<script type="text/javascript" src="js/datepicker.js"></script>
<script type="text/javascript" src="js/udlaeg_jav7.js"></script>
<script src="../timereg/inc/header_lysblaa_stat_20170808.js" type="text/javascript"></script>

<%if request("media") <> "print" then %>
<div class="wrapper">
    <div class="content">
<%end if %>

        <%
            if len(session("user")) = 0 then
	
	        errortype = 5
	        call showError(errortype)
            response.end
            end if

            thisfile = "expenceweb"

            func = request("func")     
            media = request("media")
            
            if len(trim(request("mid"))) <> 0 then
	        usemrn = request("mid")

            else

	            if len(request.cookies("timereg_2006")("usemrn")) <> 0 then
		        usemrn = request.cookies("timereg_2006")("usemrn")
		        else
		        usemrn = session("mid")
		        end if
	
            end if


            'usemrn = session("mid")

            if len(trim(request("FM_progrp"))) <> 0 then
	        progrp = request("FM_progrp")
	        else
            progrp = 0
	        end if

            if len(request("FM_medarb")) <> 0 OR func = "export" then
	
	            if left(request("FM_medarb"), 1) = "0" then 'ikke l�ngere mulig efer jq v�lg alle automatisk
	        
                    if media <> "print" then
	                thisMiduse = request("FM_medarb_hidden")
    	            else
    	            thisMiduse = request("FM_medarb")
    	            end if
    	
    	        intMids = split(thisMiduse, ", ")
	            else
	            thisMiduse = request("FM_medarb")
	            intMids = split(thisMiduse, ", ")
	            end if
	
	        else
	    
	            thisMiduse = session("mid") 
	            intMids = split(thisMiduse, ", ")
	   
	        end if


            if len(trim(request("FM_vis"))) <> 0 then
                visning = request("FM_vis")
                response.Cookies("expence")("visning") = visning
            else
                if request.Cookies("expence")("visning") <> "" then
                    visning = request.Cookies("expence")("visning")
                else
                    visning = 0
                end if
            end if

            visSEL0 = ""
            visSEL1 = ""
            visSEL2 = ""
            visSEL3 = ""
            sqlVisWhere = ""
            select case cint(visning)
                case 0
                    visSEL0 = "SELECTED"
                    sqlVisWhere = ""
                case 1
                    visSEL1 = "SELECTED"
                    sqlVisWhere = "AND godkendt = 1"
                case 2
                    visSEL2 = "SELECTED"
                    sqlVisWhere = "AND godkendt = 2"
                case 3
                    visSEL3 = "SELECTED"
                    sqlVisWhere = "AND godkendt = 0"
            end select


            if len(trim(request("FM_faktbar"))) <> 0 then
                faktbar = request("FM_faktbar")
                response.Cookies("expence")("FM_faktbar") = faktbar
            else
                if request.Cookies("expence")("FM_faktbar") <> "" then
                    faktbar = request.Cookies("expence")("FM_faktbar")
                else
                    faktbar = 0
                end if
            end if
            
            faktSELAlle = ""
            faktSEL0 = ""
            faktSEL1 = ""
            faktSEL5 = ""
            sqlFaktWhere = ""
            select case cint(faktbar)
                case -1
                    faktSELAlle = "SELECTED"
                    sqlFaktWhere = ""
                case 0
                    faktSEL0 = "SELECTED"
                case 1
                    faktSEL1 = "SELECTED"
                case 2
                    faktSEL2 = "SELECTED"
                case 5
                    faktSEL5 = "SELECTED"
            end select

            if cint(faktbar) <> -1 then
            sqlFaktWhere = "AND mf.intkode = "& faktbar
            end if

            if media <> "eksport" AND media <> "print" then
            call menu_2014
            end if


            call positiv_aktivering_akt_fn()






            select case func

                case "slet"
                    '*** Her sp�rges om der skal slettes ***
	                    oskrift = "Udl�g" 
                        slttxta = "Du er ved at <b>SLETTE</b> et udl�g. Er dette korrekt?"
                        slttxtb = "" 
                        slturl = "materialerudlaeg.asp?func=sletok&matid="& request("matid")

                        call sletcnf_2015(oskrift, slttxta, slttxtb, slturl)
            
                case "sletok"

                    matid = request("matid")
                    matnavn = request("matnavn")

                    strSQLdel = "DELETE FROM materiale_forbrug WHERE id = " & matid
                    oConn.execute(strSQLdel) 

                    matnavn = replace(matnavn, "'", "")

	                call insertDelhist("mat", matid, 0, matnavn, session("mid"), session("user"))

                    response.Redirect "materialerudlaeg.asp"
            
            case "aproveExp"
                
            expenceid = split(request("expenceid"),", ")

            for i = 0 to UBOUND(expenceid)

                if len(trim(request("godkendmat_" & expenceid(i)))) <> 0 then
                    godkendtVal = request("godkendmat_" & expenceid(i))
                else
                    godkendtVal = 0
                end if

                afregnetPreVal = request("afregnetPreVal_" & expenceid(i))

                if len(trim(request("afregnMat_" & expenceid(i)))) <> 0 then
                afregnetNewVal = request("afregnMat_" & expenceid(i))
                else
                afregnetNewVal = 0
                end if
                
                if afregnetNewVal = 1 then
                    godkendtVal = 1 'Ved afrenget bliver den godkendt automatisk
                end if

                afregnetDate = year(now) &"-"& month(now) &"-"& day(now)

                preGKVal = request("godkendtPreVal_" & expenceid(i))

                'response.Write "<br> preval " & afregnetPreVal & " newval " & afregnetNewVal & " pregodkend " & preGKVal & " newgodkend " & godkendtVal

                if cint(godkendtVal) <> 0 AND cint(godkendtVal) <> cint(preGKVal) OR cint(afregnetNewVal) <> cint(afregnetPreVal) then
                    'response.Write "<br> godkendt Value " & godkendtVal 

                    strSQL = "UPDATE materiale_forbrug SET godkendt = "& godkendtVal & ", gkaf = '"& session("user") & "', gkdato = '"& year(now) &"-"& month(now) &"-"& day(now) &"', afregnet = "& afregnetNewVal &", afdate = '"& afregnetDate &"' WHERE id = "& expenceid(i) 
                    'response.Write "<br>" & strSQL
                    oConn.execute(strSQL)                    
                end if 

            next

            response.Redirect "materialerudlaeg.asp?FM_sog_jobnavn="&request("FM_sog_jobnavn")&"&FM_personlig="&request("FM_personlig")&"&FM_fromdateSearch="&request("FM_fromdateSearch")&"&FM_medarb="&thisMiduse&"&FM_medarb_hidden="&thisMiduse


            case "dbred", "dbopr"

            jobid = request("FM_jobid")
            aktid = request("FM_aktivitetid")
            medid = request("mid")
            aftid = 0
            matId = 0
            intAntal = 1 'Ved udlaeg er antal altid 1
            strDato = year(now)&"/"&month(now)&"/"&day(now)
            strEditor = session("user")
            regdato = year(now)&"/"&month(now)&"/"&day(now)
            bilagsnr = "" 
            navn = request("FM_navn")
            FM_pris = request("FM_belob")
            salgspris = 0
            varenr = 0
            valuta = request("FM_valuta")
            gruppe = request("FM_gruppe")
            personlig = request("FM_personlig")
            intkode = request("FM_faktbar")
            betegnelse = ""
            opretlager = 0
            matregid = request("matregid")
            matava = 0
            otf = 1
            matreg_opdaterpris = 0


            mat_func = func
            mattype = 0
            call indlaes_mat(matregid, otf, medid, jobid, aktid, aftid, matId, strEditor, strDato, intAntal, regdato, valuta, intkode, personlig, bilagsnr, FM_pris, salgspris, navn, gruppe, varenr, opretlager, betegnelse, mat_func, matreg_opdaterpris, matava, mattype)

            response.Redirect "materialerudlaeg.asp?FM_sog_txt="& jobid

            case "opr", "red"


                if cint(vis_passive) = 1 then
                chkPassive = "CHECKED"
                else
                chkPassive = ""
                end if               

                id = request("id")

                if func = "red" then
	    
                    dbfunc = "dbred"

	                oSkrkift = expence_txt_058
                    fileLink = ""
                    strSQL = "SELECT mf.id as matregid, usrid, matnavn, forbrugsdato, matkobspris, mf.valuta, matgrp, personlig, intkode, filnavn, jobid, aktid, j.jobnavn as jobnavn, a.navn as aktnavn, basic_valuta, basic_kurs, basic_belob, mf.dato as kobsdato FROM materiale_forbrug as mf LEFT JOIN job as j ON (jobid = j.id) LEFT JOIN aktiviteter as a ON (mf.aktid = a.id) WHERE mf.id ="& id
                    'response.Write strSQL
                    oRec.open strSQL, oConn, 3
                    if not oRec.EOF then

                        matregid = oRec("matregid")
                        usemrn = oRec("usrid")
                        matnavn = oRec("matnavn")
                        belob = oRec("matkobspris")
                        valuta = oRec("valuta")
                        matgruppe = oRec("matgrp")
                        matpersonlig = oRec("personlig")
                        intkode = oRec("intkode")
                        jobid = oRec("jobid")
                        jobnavn = oRec("jobnavn")
                        aktid = oRec("aktid")  
                        aktnavn = oRec("aktnavn")
                        forbrugsdato = oRec("forbrugsdato")

                        if oRec("filnavn") <> "" then 
                            fileLink = "../inc/upload/"&lto&"/"&oRec("filnavn")
                        end if
                        
                        basic_valuta = oRec("basic_valuta")

                        if basic_valuta = "NA" then
                            basic_valuta = "DKK"
                        end if

                        basic_kurs = oRec("basic_kurs")
                        basic_belob = oRec("basic_belob")

                        kobYear = year(oRec("kobsdato"))                   
                        kobMonth = month(oRec("kobsdato"))
                        if cint(kobMonth) < 10 then
                            kobMonth = "0" & kobMonth
                        end if
                        kobDay = day(oRec("kobsdato"))
                        if kobDay < 10 then
                        kobDay = "0" & kobDay
                        end if                    

                        kobsdato = kobYear &"-"& kobMonth &"-"& kobDay
                    

                    end if
                    oRec.close

	            else 'opr
	    

                    dbfunc = "dbopr"
                    oSkrkift = expence_txt_059
                    matregid = 0
                    usemrn = 0
                    forbrugsdato = day(now) &"-"& month(now) &"-"& year(now)
                    matnavn = ""
                    belob = 0
                    valuta = basisValId
                    matgruppe = 0
                    
                    select case lto
	                case "immenso", "epi", "epi_no", "epi_sta", "epi_ab", "alfanordic", "intranet - local"
	                matpersonlig = 1
	                case else
	                matpersonlig = 0
	                end select

                    jobid = 0
                    jobnavn = ""
                    aktid = 0
                    aktnavn = ""

                    select case lto 
                    case "synergi1", "intranet - local", "epi", "epi_as", "epi_se", "epi_os", "epi_uk", "alfanordic"
                    intKode = 1 'Intern
                    case else 
	                intKode = 2 'Ekstern = viderefakturering
                    end select

                    basic_valuta = "DKK"
                    basic_kurs = 0
                    basic_belob = 0

                    kobYear = year(now)                   
                    kobMonth = month(now)
                    if cint(kobMonth) < 10 then
                        kobMonth = "0" & kobMonth
                    end if
                    kobDay = day(now)
                    if kobDay < 10 then
                    kobDay = "0" & kobDay
                    end if

                    kobsdato = kobYear &"-"& kobMonth &"-"& kobDay
                

	            end if



            %>

                <style>
                    table, tr, td
                    {
                        color:black;
                        padding:0 15px 10px 0px;
                    }
                </style>

                <div class="container">
                    <div class="portlet">
                        <h3 class="portlet-title"><u><%=expence_txt_001 %> - <%=oSkrkift %></u></h3>
                        <div class="portlet-body">
                            
                            <div class="well well-white">
                                <h4 class="panel-title-well"><!--<%=medarb_txt_021 %></h4>-->
                              
                            <br />&nbsp;
                                <div class="row" id="errorMessageRow" style="display:none;">
                                    <div class="col-lg-12"><span id="errorMessage" style="color:darkred;"></span></div>
                                </div>

                                <div class="row">
                                    <div class="col-lg-1"></div>
                                  <!--  <form action="materialerudlaeg.asp?func=<%=dbfunc %>" method="post"> -->

                                        <input type="hidden" id="jq_jobidc" name="" value="-1"/>
                                        <input type="hidden" id="jq_aktidc" name="" value="-1"/>
                                        <input type="hidden" id="varTjDatoUS_man" value="<%=year(now)&"-"&month(now)&"-"&day(now) %>" />
                                        <input type="hidden" id="lto" name="" value="<%=lto %>"/>
                                        <input type="hidden" id="FM_pa" name="FM_pa" value="<%=pa_aktlist %>"/>
                                        <input type="hidden" name="matregid" id="matregid" value="<%=matregid %>" />
                                        <input type="hidden" id="func" value="<%=func %>" />                                        

                                        <div class="col-lg-6">
                                            <table style="display:inline-table;">

                                                <tr>
                                                    <td style="vertical-align:top;"><%=expence_txt_002 %> <span style="color:red;">*</span></td>                                        
                                                    <td style="width:100%"><input type="text" id="FM_job" name="FM_job" placeholder="<%=tsa_txt_066 %>/<%=tsa_txt_236 %>" value="<%=jobnavn %>" class="FM_job form-control input-small" autocomplete="off" />

                                                        <select id="dv_job" size="10" class="form-control input-small chbox_job" style="visibility:hidden; display:none; width:100%;">
                                                            <option><%=expence_txt_003 %>..</option>
                                                        </select>

                                                         <input type="hidden" id="FM_jobid" name="FM_jobid" value="<%=jobid %>"/>
                                                    </td>
                                                </tr>

                                                <tr>
                                                    <td style="vertical-align:top;"><%=expence_txt_004 %></td>
                                                    <td><input style="font-size:<%=inputFont%>; height:<%=inputHeight%>" type="text" id="FM_akt_0" name="activity" placeholder="<%=tsa_txt_068%>" value="<%=aktnavn %>" class="FM_akt form-control input-small" autocomplete="off" />

                                                        <select id="dv_akt_0" class="form-control input-small chbox_akt" size="10" style="visibility:hidden; display:none;"> <!-- chbox_akt -->
                                                            <option><%=expence_txt_003 %>..</option>
                                                        </select>

                                                        <input type="hidden" id="FM_aktid_0" name="FM_aktivitetid" value="<%=aktid %>"/>
                                                    </td>
                                                </tr>

                                                <tr>
                                                    <td><%=expence_txt_005 %>
                                                        <!--
                                                        <br /> 
                                                       <span style="font-size:9px; display:none;">Passive <input type="checkbox" value="1" id="FM_vis_passive" name="FM_vis_passive" <%=chkPassive %> /></span> 
                                                        -->
                                                    </td>
                                                    <td id="td_selmedarb"><%call selmedarbOptions_2018 %></td>
                                                </tr>

                                                <tr>
                                                    <td>Forbrugsdato</td>
                                                    <td>
                                                        <div class='input-group date'>
                                                            <input type="text" class="form-control input-small" name="FM_forbrugsdato" id="FM_forbrugsdato" value="<%=forbrugsdato %>" placeholder="dd-mm-yyyy" />
                                                            <span class="input-group-addon input-small">
                                                                <span class="fa fa-calendar">
                                                                </span>
                                                            </span>
                                                        </div>
                                                    </td>
                                                </tr>

                                                <tr>
                                                    <td><%=expence_txt_006 %></td>
                                                    <td><input type="text" name="FM_navn" id="FM_navn" class="form-control input-small" value="<%=matnavn %>" /></td>
                                                </tr>

                                                <tr>
                                                    <td><%=expence_txt_007 %></td>
                                                    <td><input type="text" name="FM_belob" id="FM_belob" class="form-control input-small" value="<%=belob %>" style="width:75%" /></td>

                                                    <input type="hidden" id="FM_basic_valuta" value="<%=basic_valuta %>" />
                                                    <input type="hidden" id="FM_basic_kurs" value="<%=basic_kurs %>" />
                                                    <input type="hidden" id="FM_basic_belob" value="<%=basic_belob %>" />
                                                    <input type="hidden" id="FM_kobsdato" value="<%=kobsdato%>" />

                                                </tr>

                                                <tr>
                                                    <td><%=expence_txt_008 %></td>
                                                    <td>
                                                        <select name="FM_valuta" id="FM_valuta" class="form-control input-small" style="width:75%">
                                                        <!--<option value="0"><=tsa_txt_229 %></option>-->
		                                                    <%
		                                                    strSQL3 = "SELECT id, valutakode, grundvaluta FROM valutaer ORDER BY valutakode"
    		
    		
		                                                    oRec3.open strSQL3, oConn, 3 
		                                                    while not oRec3.EOF 
    		
		                                                    if cint(valuta) = oRec3("id") then
		                                                    valGrpCHK = "SELECTED"
		                                                    else
		                                                    valGrpCHK = ""
		                                                    end if
		    
		   
		                                                    %>
		                                                    <option value="<%=oRec3("id")%>" <%=valGrpCHK %> data-valutakode="<%=oRec3("valutakode") %>" ><%=oRec3("valutakode")%></option>
		                                                    <%
		                                                    oRec3.movenext
		                                                    wend
		                                                    oRec3.close %>
                                                        </select>
                                                    </td>
                                                </tr>

                                                <tr>
                                                    <td><%=expence_txt_009 %></td>
                                                    <td>
                                                        <select name="FM_gruppe" id="FM_gruppe" class="form-control input-small" style="width:75%">
                                                        <option value="0"><%=expence_txt_010 %></option>
                                                        <%
                                                            strSQL2 = "SELECT id, navn, av, nummer FROM materiale_grp ORDER BY navn"
                                                            oRec2.open strSQL2, oConn, 3
                                                            while not oRec2.EOF

                                                            if cint(matgruppe) = oRec2("id") then
                                                            matgrpSEL = "SELECTED"
                                                            else
                                                            matgrpSEL = ""
                                                            end if

                                                            response.Write "<option value='"&oRec2("id")&"' "& matgrpSEL &">"&oRec2("navn")&"</option>"

                                                            oRec2.movenext
                                                            wend
                                                            oRec2.close
                                                        %>

                                                    </select>
                                                    </td>
                                                </tr>

                                                <tr>
                                                    <td><%=expence_txt_011 %></td>
                                                    <td>
                                                        <select name="FM_personlig" id="FM_personlig" class="form-control input-small" style="width:75%">
                                                            <%
                                                            personSELCHB = ""
                                                            firmaSELCHB = ""
                                                            SELMCard = ""
                                                            SELworkplus = ""
                                                            SELbrobizz = ""
                                                            SELRejse = ""

                                                            select case matpersonlig
                                                            case 1
                                                                personSELCHB = "SELECTED"
                                                            case 0
                                                                firmaSELCHB = "SELECTED"
                                                            case 20
                                                                SELMCard = "SELECTED" 
                                                            case 21
                                                                SELworkplus = "SELECTED" 
                                                            case 22
                                                                SELbrobizz = "SELECTED" 
                                                            case 23
                                                                SELRejse = "SELECTED" 

                                                            end select
                                                            %>

                                                            <option value="1" <%=personSELCHB %>><%=expence_txt_011 %></option>
                                                            <option value="0" <%=firmaSELCHB %>><%=expence_txt_012 %></option>

                                                            <%if lto = "lm" then %>

                                                            <option value="20" <%=SELMCard %>> M.Card</option>
                                                            <option value="21" <%=SELworkplus %>>Workplus</option>
                                                            <option value="22" <%=SELbrobizz %>>BroBizz</option>
                                                            <option value="23" <%=SELRejse %>>Rejsekort</option>

                                                            <%end if %>
                                                        </select>
                                                    </td>
                                                </tr>

                                                <tr>
                                                    <td><%=expence_txt_013 %></td>
                                                    <td>
                                                        <select name="FM_faktbar" id="FM_faktbar" class="form-control input-small" style="width:75%">
                                                            <%
                                                                intkodeSEL0 = ""
                                                                intkodeSEL1 = ""
                                                                intkodeSEL2 = ""
                                                                intkodeSEL5 = ""

                                                                select case intkode
                                                                case 0
                                                                intkodeSEL0 = "SELECTED"
                                                                case 1
                                                                intkodeSEL1 = "SELECTED"
                                                                case 2
                                                                intkodeSEL2 = "SELECTED"
                                                                case 5
                                                                intkodeSEL5 = "SELECTED"
                                                                end select

                                                            %>

                                                            <%if lto <> "execon" AND lto <> "immenso" AND lto <> "unik" then %>
		                                                    <option value="0" <%=intkodeSEL0 %>><%=tsa_txt_235 %></option><!-- ingen -->
		                                                    <%end if%>
		    
		                                                    <option value="1" <%=intkodeSEL1 %>><%=tsa_txt_232 %></option><!-- intern -->
		                                                    <option value="2" <%=intkodeSEL2 %>><%=tsa_txt_233 %> <%if lto = "nt" then %> 100 % <%end if %></option><!-- ekstern -->

                                                            <%if lto = "nt" then %>
                                                            <option value="5" <%=intkodeSEL5 %>><%=tsa_txt_233 %> 50 %</option>
                                                            <%end if %>

                                                        </select>
                                                    </td>
                                                </tr>

                                            </table>
                                        </div>
                                  <!--  </form> -->                                    

                                     <%
                                        matId = id

                                        if func = "opr" then
                                            strSQL = "SELECT id FROM materiale_forbrug"
                                            oRec.open strSQL, oConn, 3 
                                            while not oRec.EOF
                                            lastid = oRec("id")
                                            oRec.movenext
                                            wend
                                            oRec.close


                                            matId = lastid + 1

                                        end if
                                    %>

                                     <form ENCTYPE="multipart/form-data" id="imageUpload" action="../timereg/upload_bin.asp?matUpload=1&matId=<%=matId %>&thisfile=<%=thisfile %>&FM_sog_jobnavn=<%=request("FM_sog_jobnavn")%>&FM_personlig=<%=request("FM_personlig")%>&FM_fromdateSearch=<%=request("FM_fromdateSearch")%>&FM_medarbejder=<%=thisMidUse %>" method="post">
                                         
                                        <div class="col-lg-5">
                                            <table>
                                                <tr>
                                                    <td>
                                                        <div class="fileinput fileinput-new" data-provides="fileinput">
                                                            <div class="fileinput-preview thumbnail" style="width: 250px; height: 300px;">
                                                                <img id="imageholder" src="<%=fileLink %>" alt='' border='0' style="max-width:100%">                                                   
                                                            </div>
                                                        </div>                                                 
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td style="text-align:center;">
                                                        <label class="btn btn-default btn-sm"><b><INPUT id="image-file" NAME="fileupload1" TYPE="file" style="display:none;" onchange="readURL(this);"> <%=expence_txt_014 %> </b></label>
                                                    </td>
                                                </tr>
                                            </table>
                                        </div>

                                         <script type="text/javascript">
                                             function readURL(input) {
                                                   // alert("show pic")

                                                    if (input.files && input.files[0]) {
                                                        var reader = new FileReader();

                                                        reader.onload = function (e) {
                                                            $('#imageholder')
                                                                .attr('src', e.target.result);
                                                        };

                                                        reader.readAsDataURL(input.files[0]);
                                                    }
                                                }
                                         </script>

                                    </form>


                                </div>
                            </div>

                                 <div class="row">
                                    <div class="col-lg-12">
                                        <button type="submit" id="matsub" class="btn btn-success btn-sm pull-right"><b><%=medarb_txt_020 %></b></button>
                                    </div>
                                </div>

                            


                          <!--  <div class="well well-white" style="display:none;">
                                <h4 class="panel-title-well"><%=medarb_txt_021 %></h4>

                                <div class="row">

                                    <div class="col-lg-1"></div>

                                    <div class="col-lg-1"><%=expence_txt_006 %></div>
                                    <div class="col-lg-3"><input type="text" name="FM_navn" class="form-control input-small" value="<%=matnavn %>" /></div>

                                    <div class="col-lg-3"><%=expence_txt_005 %> 
                                        <%if level <= 2 OR level = 6 then %>
                                        &nbsp <input type="checkbox" value="1" id="FM_vis_passive" name="FM_vis_passive" <%=chkPassive %> /> <%=expence_txt_015 %>
                                        <%end if %>
                                    </div>
                                    <div class="col-lg-3" id="">
                                        <%call selmedarbOptions_2018 %>
                                    </div>
                                </div>

                                <br />

                                <div class="row">
                                    <div class="col-lg-1"></div>
                                    <div class="col-lg-1"><%=expence_txt_016 %>:</div>
                                    <div class="col-lg-2"><input type="text" class="form-control input-small" value="<%=matantal %>" /></div>

                                    <div class="col-lg-1"></div>

                                    <div class="col-lg-3"><%=expence_txt_017 %></div>
                                    <div class="col-lg-2">
                                        <div class='input-group date'>
                                            <input type="text" class="form-control input-small" name="FM_forbrugsdato" value="<%=FM_forbrugsdato %>" placeholder="dd-mm-yyyy" />
                                            <span class="input-group-addon input-small">
                                                <span class="fa fa-calendar">
                                                </span>
                                            </span>
                                        </div>
                                    </div>

                                </div>

                                <div class="row">
                                    <div class="col-lg-1"></div>
                                    <div class="col-lg-1"><%=expence_txt_018 %>:</div>   
                                    <div class="col-lg-2"><input type="text" class="form-control input-small" /></div>

                                    <div class="col-lg-1"></div>
                                    <div class="col-lg-3"><%=expence_txt_008 %></div>
                                    <div class="col-lg-2">
                                        <select class="form-control input-small">
                                            <option>DKK</option>
                                        </select>
                                    </div>

                                </div>

                                <div class="row">
                                    <div class="col-lg-1"></div>
                                    <div class="col-lg-1"><%=expence_txt_019 %>:</div>     
                                    <div class="col-lg-2"><input type="text" class="form-control input-small" /></div>

                                    <div class="col-lg-1"></div>

                                    <div class="col-lg-3"><%=expence_txt_020 %>:</div>
                                    <div class="col-lg-2"><input type="text" class="form-control input-small" /></div>

                                </div>

                                <div class="row">
                                    <div class="col-lg-1"></div>
                                    <div class="col-lg-1"><%=expence_txt_021 %>: <br /></div>
                                    <div class="col-lg-1"><input type="text" class="form-control input-small" /></div>
                                    <div class="col-lg-2"><input type="checkbox" id="matreg_opretlager_0" name="opretlager" value="1" /> <%=expence_txt_022 %></div>
       
                                    <div class="col-lg-3"><%=expence_txt_023 %></div>
                                    <div class="col-lg-2">
                                        <select class="form-control input-small">
                                            <option><%=expence_txt_010 %></option>
                                        </select>
                                    </div>
                                </div>

                                <div class="row">
                                    <div class="col-lg-1"></div>
                                    <div class="col-lg-1"><%=expence_txt_024 %>:</div>
                                    <div class="col-lg-2">
                                        <select class="form-control input-small">
                                            <option><%=expence_txt_025 %></option>
                                        </select>
                                    </div>

                                    <div class="col-lg-1"></div>

                                    <div class="col-lg-3"><%=expence_txt_026 %></div>
                                    <div class="col-lg-2"><input type="text" class="form-control input-small" /></div>

                                </div>
                                

                            </div> -->

                        </div>


                    </div>
                </div>





        <%
        case else 


            if level = 1 then
                approveActive = ""
                approveActiveInt = 1
                afregnetActive = ""
                afregnetActiveInt = 1
           else
                afregnetActive = "DISABLED"
                approveActiveInt = 0
                afregnetActiveInt = 0
                call fTeamleder(session("mid"), 0, 1)
                if strPrgids = "ingen" then
                    approveActive = "DISABLED"
                else
                    approveActive = ""
                    approveActiveInt = 1
                end if
            end if 'level = 1





            'jobnr_sog = request("FM_sog_jobnr") 
            'jobnavnsog = request("FM_sog_jobnavn")

            sogKri = ""
            if len(trim(request("FM_sog_txt"))) <> 0 then
                jobsogtxt = request("FM_sog_txt")
                sogKri = "AND (jobnavn LIKE '%"& jobsogtxt &"%' OR j.jobnr like '%"& jobsogtxt &"%')"
            else
                jobsogtxt = ""
            end if  
           
            if len(trim(request("FM_fromdateSearch"))) <> 0 then
                FM_fromdateSearch = request("FM_fromdateSearch")
                response.Cookies("to_2015")("fromdateSearch") = FM_fromdateSearch
            else
                if request.Cookies("to_2015")("fromdateSearch") <> "" then
                    FM_fromdateSearch = request.Cookies("to_2015")("fromdateSearch")
                else 
                    FM_fromdateSearch = "01-"& month(now) &"-"& year(now)
                    response.Cookies("to_2015")("fromdateSearch") = FM_fromdateSearch
                end if               
            end if

            FM_fromdateSearchSQL = year(FM_fromdateSearch) &"-"& month(FM_fromdateSearch) &"-"& day(FM_fromdateSearch)

            sogKri = sogKri & " AND forbrugsdato >= '"& FM_fromdateSearchSQL &"'"


            if len(trim(request("FM_personlig"))) <> 0 then
                FM_personlig = request("FM_personlig")
            else
                FM_personlig = "-1"
            end if

            if FM_personlig <> "-1" then
                sogKri = sogKri & " AND personlig = "& FM_personlig
            end if

            'if thisMiduse <> "0" then

                'FM_medarbejder = split(thisMiduse,", ")
                'strFomr_rel = request("FM_medarb")
                'strFomr_rel = replace(strFomr_rel, "X234", "#")
                antalM = 0

                for m = 0 to UBOUND(intMids)
                    if m = 0 then
                        sogKri = sogKri & " AND (mf.usrid = "& intMids(m)
                    else
                        sogKri = sogKri & " OR mf.usrid = "& intMids(m)
                    end if
                    antalM = antalM + 1
                next

                if antalM > 0 then
                    sogKri = sogKri & ")"
                end if

                'fomrIDs = replace(strFomr_rel, "#", ",")
                'fomrIDs = replace(fomrIDs, ",,", ",")
               

            'else

            '    strFomr_rel = "#0#"
            '    FM_medarbejder = 0

            'end if


            'response.Write "<br> <br><br><br><br>her her" & fomrIDs
            'response.Write sogKri
            strSogeKriterier = "FM_sog_txt="&request("FM_sog_txt")&"&FM_personlig="&request("FM_personlig")&"&FM_medarb="&thisMiduse&"&FM_medarb_hidden="& thisMiduse &"&FM_fromdateSearch="&request("FM_fromdateSearch")

        %>


        
       


      <!-- <script type="text/javascript" src="js/udlaeg_liste_jav.js"></script> -->

        <!--
        <style>

            /* The Modal (background) */
            .modal {
                display: none; /* Hidden by default */
                position: fixed; /* Stay in place */
                z-index: 1000; /* Sit on top */
                padding-top: 150px; /* Location of the box */
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
            }

            .picmodal:hover,
            .picmodal:focus {
            text-decoration: none;
            cursor: pointer;
            }
   
        </style>
        -->

        <%
            if media = "eksport" then
                displayExp = "none"
            end if

            conWidth = "width:65%"
            if media = "print" then
                conWidth = "width:100%;"
            end if
        %>

        <div class="container" style="display:<%=displayExp%>; <%=conWidth%>">
            <div class="portlet">
                <h3 class="portlet-title"><u><%=expence_txt_001 %></u></h3>
                <div class="portlet-body">

                
                    <%if media <> "print" then %>
                    <form action="materialerudlaeg.asp?func=opr&<%=strSogeKriterier %>" method="post">
                            <input name="FM_medarb" type="hidden" value="<%=thisMidUse %>" />
                            <input name="FM_medarb_hidden" type="hidden" value="<%=thisMidUse %>" />
                        <div class="row">
                            <div class="col-lg-10">&nbsp</div>
                            <div class="col-lg-2 pad-b10">
                                <button type="submit" class="btn btn-success btn-sm pull-right"><b><%=expence_txt_027 %></b></button>
                            </div>
                        </div>
                    </form>
                    <%end if %>

                    <%if media <> "print" then %>
                    <div class="well">
                        <form action="materialerudlaeg.asp" method="post">
                            
                            <h4 class="panel-title-well"><%=expence_txt_028 %></h4>

                            <div class="row">
                                <div class="col-lg-3"><%=expence_txt_029 %>:</div>
                                <div class="col-lg-2"><%=expence_txt_030 %>:</div>
                                <div class="col-lg-2"><%=expence_txt_031 %>:</div>                   
                                <div class="col-lg-2"><%=expence_txt_013 %></div>
                                <div class="col-lg-2"><%=expence_txt_032 %>:</div>
                            </div>

                            <div class="row">                                
                                <div class="col-lg-3"><input type="text" class="form-control input-small" name="FM_sog_txt" value="<%=jobsogtxt %>" /></div>

                                <div class="col-lg-2">

                                    <select class="form-control input-small" name="FM_personlig">

                                        <%
                                            SELall = ""
                                            SELpersonlig = ""
                                            SELfirma = ""

                                            SELMCard = ""
                                            SELworkplus = ""
                                            SELbrobizz = ""
                                            SELRejse = ""

                                            select case FM_personlig
                                            case "-1"
                                            SELall = "SELECTED"
                                            case 1
                                            SELpersonlig = "SELECTED" 
                                            case 0
                                            SELfirma = "SELECTED"

                                            case 20
                                                SELMCard = "SELECTED" 
                                            case 21
                                                SELworkplus = "SELECTED" 
                                            case 22
                                                SELbrobizz = "SELECTED" 
                                            case 23
                                                SELRejse = "SELECTED" 

                                            end select
                                        %>

                                        <option value="-1" <%=SELall %>"><%=expence_txt_033 %></option>
                                        <option value="1" <%=SELpersonlig %>><%=expence_txt_034 %></option>
                                        <option value="0" <%=SELfirma %>><%=expence_txt_012 %></option>

                                        <%if lto = "lm" then %>
                                            <option value="20" <%=SELMCard %>> M.Card</option>
                                            <option value="21" <%=SELworkplus %>>Workplus</option>
                                            <option value="22" <%=SELbrobizz %>>BroBizz</option>
                                            <option value="23" <%=SELRejse %>>Rejsekort</option>
                                        <%end if %>
                                    </select>
                                </div>

                                <div class="col-lg-2">
                                    <div class='input-group date'>
                                        <input type="text" class="form-control input-small" name="FM_fromdateSearch" value="<%=FM_fromdateSearch %>" placeholder="dd-mm-yyyy" autocomplete="off" />
                                        <span class="input-group-addon input-small">
                                            <span class="fa fa-calendar">
                                            </span>
                                        </span>
                                    </div>
                                </div>

                                <div class="col-lg-2">
                                    <select name="FM_faktbar" class="form-control input-small">
                                        <%if lto = "nt" then %>
                                            <option value="-1" <%=faktSELAlle %> ><%=expence_txt_033 %></option>
                                            <option value="1" <%=faktSEL1 %>><%=expence_txt_070 %></option>
                                            <option value="2" <%=faktSEL2 %>><%=expence_txt_013 %> 100 %</option>
                                            <option value="5" <%=faktSEL5 %>><%=expence_txt_013 %> 50 %</option>
                                        <%else %>
                                            <option value="-1" <%=faktSELAlle %> ><%=expence_txt_033 %></option>
                                            <option value="1" <%=faktSEL1 %> ><%=expence_txt_070 %></option>
                                            <option value="2" <%=faktSEL2 %> ><%=expence_txt_013 %></option>
                                        <%end if %>
                                    </select>
                                </div>

                                <div class="col-lg-2">
                                    <select name="FM_vis" class="form-control input-small">
                                        <option value="0" <%=visSEL0 %>><%=expence_txt_033 %></option>
                                        <option value="1" <%=visSEL1 %>><%=expence_txt_035 %></option>
                                        <option value="2" <%=visSEL2 %>><%=expence_txt_036 %></option>
                                        <option value="3" <%=visSEL3 %>><%=expence_txt_037 %></option>
                                    </select>
                                </div>

                            </div>

                            <div class="row">
                                <div class="col-lg-3"><%=expence_txt_005 %>:

                                </div>
                            </div>

                            <div class="row">
                             
                                <script src="../timereg/inc/header_lysblaa_stat_20170808.js" type="text/javascript"></script>
                                <%
                                   'if session("mid") = 1 then
                                   'Response.Write "<hr><br><br><br><br><br><br><br><br><br><br>"
                                   'Response.Write "HER: " & progrp & " level: "& level & " usemrn: " & usemrn
                                   'end if
                                  
                                    %>
                                <%call progrpmedarb_2018 %>

                            </div>

                            <div class="row">
                                <div class="col-lg-12"><button type="submit" class="btn btn-secondary btn-sm pull-right"><b><%=expence_txt_038 %></b></button></div>
                            </div>

                        </form>
                    </div>
                    <%end if %>
                    

                    <form action="materialerudlaeg.asp?func=aproveExp&<%=strSogeKriterier %>" method="post">

                         <%if media <> "print" then %>
                        <div class="row">
                            <div class="col-lg-12">
                                <button type="submit" class="btn btn-success btn-sm pull-right"><b><%=medarb_txt_020 %></b></button>
                            </div>
                        </div>
                        <%end if %>

                        <table id="expencetable" class="table dataTable table-striped table-bordered table-hover ui-datatable">

                            <thead>
                                <tr>
                                    <th><%=expence_txt_002 %> <br /> <%=expence_txt_004 %></th>
                                    <th><%=expence_txt_006 %></th>
                                    <th><%=expence_txt_039 %></th>
                                    <th><%=expence_txt_005 %></th>
                                    <th><%=expence_txt_040 %></th>
                                    <th><%=expence_txt_007 & " " %><%=expence_txt_008 %></th>

                                   <!-- <th>Basis valuta</th> expence_txt_041  -->
                                    <th><%=expence_txt_007 &" "& basisValISO%><br />
                                       
                                    </th>


                                    <th><%=expence_txt_042 %></th>

                                    <%if lto <> "lm" then %>
                                    <th><%=expence_txt_013 %></th>
                                    <%end if %>

                                    <%if media <> "print" then %><th><%=expence_txt_043 %></th><%end if %>
                                    <th style="text-align:center;">
                                        <%if media <> "print" then %>
                                        <%=expence_txt_044 %> <br /><input type="radio" name="godkenAfvisAll" id="godkendAll" <%=approveActive %> />
                                        <%else %>
                                        <%=expence_txt_045 %>
                                        <%end if %>
                                    </th>                                
                                    <th style="text-align:center;">
                                        <%if media <> "print" then %>
                                        <%=expence_txt_046 %> <br /><input type="radio" id="afvisAll" name="godkenAfvisAll" <%=approveActive %> />
                                        <%else %>
                                        <%=expence_txt_047 %>
                                        <%end if %>
                                    </th>                                    
                                    <th style="text-align:center;">
                                        <%if media <> "print" then %>
                                        <%=expence_txt_048 %> <br /> <input type="checkbox" id="afregnAlle" <%=afregnetActive %> />
                                        <%else %>
                                        <%=expence_txt_048 %>
                                        <%end if %>
                                    </th>

                                    <%if media <> "print" then %>
                                    <th style="text-align:center;"><%=expence_txt_066 %></th>
                                    <%end if %>
                                    
                                </tr>
                            </thead>

                            <tbody>

                                <%
                                
                                ekspTxt = ""

                                strSQL = "SELECT mf.id as uploadid, matid, matnavn, intkode, matantal, matkobspris, matvarenr, forbrugsdato, jobid, usrid, matgrp, personlig, filnavn, mf.valuta, godkendt, gkaf, gkdato, afregnet, afdate, mf.aktid as expaktid, a.navn as aktnavn, "_ 
                                & "j.jobnavn as jobnavn, j.id as projektid, "_ 
                                & "m.mnavn as medarbnavn, m.init as medarbinit, "_ 
                                & "mg.navn as matgrpnavn, basic_valuta, basic_kurs, basic_belob, k.kkundenavn as kundenavn "_
                                & "FROM materiale_forbrug as mf "_
                                & "LEFT JOIN job as j ON (j.id = jobid) "_
                                &" LEFT JOIN medarbejdere as m ON (m.mid = usrid) "_
                                &" LEFT JOIN aktiviteter as a ON (a.id = mf.aktid) "_
                                &" LEFT JOIN kunder as k ON (k.kid = j.jobknr) "_
                                &" LEFT JOIN materiale_grp as mg ON (mg.id = matgrp) WHERE matid = 0 "& sqlVisWhere &" "& sogKri & " AND mattype = 1 "& sqlFaktWhere & " ORDER BY mf.id DESC"
                             
                               'if session("mid") = 1 then     
                               'response.Write "<br> Sql" & strSQL
                               'end if
                                    
                                antal = 0
                                oRec.open strSQL, oConn, 3

                                while not oRec.EOF
                                'response.Write "<br> " & oRec("matnavn") & " fil " & oRec("filnavn")
                                afregnetSEL = ""
                                godkendtSEL = ""
                                godkendtExpt = ""
                                afvistExpt = ""
                                afvistSEL = ""
                                afregnetExpt = ""
                                personExpTxt = ""
                                faktbartxt = ""
                                aktnavneks = ""

                                if oRec("afregnet") <> 0 then
                                    afregnetSEL = "CHECKED"
                                    afregnetExpt = expence_txt_049
                                end if
                                
                                select case oRec("godkendt")
                                case 1
                                    godkendtSEL = "CHECKED"
                                    godkendtExpt = expence_txt_049
                                    afvistSEL = ""
                                case 2
                                    afvistSEL = "CHECKED"
                                    afvistExpt = expence_txt_049
                                    godkendtSEL = ""
                                end select

                                thisApproveactiveInt = approveActiveInt
                                if cint(oRec("afregnet")) <> 0 then
                                    thisApproveactiveInt = 0
                                end if

                                %>
                                <input type="hidden" value="<%=oRec("uploadid") %>" name="expenceid" />
                                <input type="hidden" value="<%=oRec("godkendt") %>" name="godkendtPreVal_<%=oRec("uploadid") %>" />
                                <input type="hidden" value="<%=oRec("afregnet") %>" name="afregnetPreVal_<%=oRec("uploadid") %>" />

                                <tr>
                                    <td>
                                        <%if media <> "print" then %>
                                        <span style="font-size:9px;"><%=oRec("kundenavn") %></span><br />
                                        <a href="../timereg/jobs.asp?func=red&id=<%=oRec("projektid") %>&jobid=<%=oRec("projektid") %>"><%=oRec("jobnavn") %></a>
                                        <%else %>
                                        <%=oRec("jobnavn") %>
                                        <%end if %>

                                        <%
                                        if oRec("aktnavn") <> "" then
                                            aktnavneks = " / " & oRec("aktnavn")
                                            response.Write "<span style='font-size:9px;'>(" & oRec("aktnavn") & ")</span>"
                                        end if
                                        %>
                                    </td>

                                    <td>
                                        <%if media <> "print" AND cint(oRec("afregnet")) <> 1 AND level = 1 then %>
                                                <a href="materialerudlaeg.asp?func=red&id=<%=oRec("uploadid") %>&<%=strSogeKriterier %>"><%=oRec("matnavn") %> </a>
                                        <%else %>
                                            <%=oRec("matnavn") %>
                                        <%end if %>

                                        <%if cdbl(oRec("matantal")) > 1 then %>
                                        <span style="font-size:9px;">(<%=oRec("matantal") %>)</span>
                                        <%end if %>
                                    </td>
                               
                                    <td><%=oRec("forbrugsdato") %></td>
                                    <td><%=oRec("medarbnavn") %> (<%=oRec("medarbinit") %>)</td>

                                    <td>
                                        <%
                                            if len(trim(oRec("matgrpnavn"))) <> 0 then
                                                response.Write oRec("matgrpnavn")
                                            else
                                                response.Write "-"
                                            end if
                                        %>
                                    </td>

                                    <td style="text-align:right;">
                                        <%
                                        valutakode = ""
                                        strSQLValuta = "SELECT valutakode FROM valutaer WHERE id = "& oRec("valuta")
                                        oRec2.open strSQLValuta, oConn, 3
                                        if not oRec.EOF then
                                            valutakode = oRec2("valutakode")
                                        end if
                                        oRec2.close
                                        %>
                                         
                                        <%
                                        prisIalt = oRec("matkobspris") * oRec("matantal")
                                        response.Write formatnumber(prisIalt, 2) & " " & valutakode %>
                                    </td>

                                   <!-- <td><%=oRec("basic_valuta") %></td> -->

                                    <td style="text-align:right;"><%=formatnumber((oRec("basic_belob") * oRec("matantal")), 2) & " " & oRec("basic_valuta") %></td>

                                    <td style="text-align:center;">

                                        <%
                                        select case cint(oRec("personlig")) 
                                            case 0
                                                personExpTxt = expence_txt_050
                                            case 1
                                                personExpTxt = expence_txt_049
                                            case 20
                                                personExpTxt = "M.Card"
                                            case 21
                                                personExpTxt = "Workplus"
                                            case 22
                                                personExpTxt = "BroBizz"
                                            case 23
                                                personExpTxt = "Rejsekort"
                                        end select
                                        %>

                                        <%=personExpTxt %>
                                    </td>

                                    <%if lto <> "lm" then %>
                                    <td>
                                        <%
                                            select case cint(oRec("intkode"))
                                                case 0
                                                    faktbartxt = expence_txt_071
                                                case 1
                                                    faktbartxt = expence_txt_070
                                                case 2
                                                    if lto = "nt" then
                                                        faktbartxt = expence_txt_013 & " 100 %"
                                                    else
                                                        faktbartxt = expence_txt_071
                                                    end if
                                                case 5
                                                    faktbartxt = expence_txt_013 & " 50 %"
                                            end select

                                            response.Write faktbartxt
                                        %>
                                    </td>
                                    <%end if %>

                                     <%if media <> "print" then %>
                                    <td style="text-align:center;">
                                    
                                        <%if oRec("filnavn") <> "" then %>
                                            <span id="modal_<%=oRec("uploadid") %>" class="fa fa-image picmodal" style="cursor:pointer;"></span>
                                            <div id="myModal_<%=oRec("uploadid") %>" class="modal" style="padding-top:12%; background-color: rgb(0,0,0); background-color: rgba(0,0,0,0.4);"> 
                                                <img id="image_<%=oRec("uploadid") %>" data-id="<%=oRec("uploadid") %>" src="../inc/upload/<%=lto %>/<%=oRec("filnavn") %>" alt='' border='0' style="width:45%; cursor:pointer;" class="modal-content rotateImage" />
                                                <br /><br /><br /><br /><br /><br /><br />

                                              <!--  <b><span class="btn btn-default btn-sm rotateImage" id="<%=oRec("uploadid")%>"><span style="font-size:200%;" class="fa fa-refresh"></span></span></b> -->
                                            </div>

                                        <%else %>
                                        -
                                        <%end if %>

                                    </td>
                                    <%end if %>

                                    <td style="text-align:center;">
                                        <%if media <> "print" then %>
                                            <%if thisApproveactiveInt = 1 then %>                                            
                                            <input type="radio" name="godkendmat_<%=oRec("uploadid") %>" <%=godkendtSEL %> value="1" class="godkendCHB" />
                                            <%else %>
                                            <input type="radio" disabled <%=godkendtSEL %> />
                                            <input type="hidden" name="godkendmat_<%=oRec("uploadid") %>" value="<%=oRec("godkendt") %>" />
                                            <%end if %>
                                        <%else %>
                                            
                                            <%if cint(oRec("godkendt")) = 1 then %>
                                            <%=expence_txt_049 %>
                                            <%else %>
                                            <%=expence_txt_050 %>
                                            <%end if %>
                                        <%end if %>

                                        <%if oRec("gkdato") <> "" AND cint(oRec("godkendt")) = 1 then %>
                                        <span style="font-size:9px;"><br /> (<%=oRec("gkdato") %>)</span>
                                        <%end if %>
                                    </td>

                                    <td style="text-align:center;">
                                        <%if media <> "print" then %>
                                            <%if thisApproveactiveInt = 1 then %>  
                                            <input type="radio" name="godkendmat_<%=oRec("uploadid") %>" <%=afvistSEL %> value="2" class="afvisCHB" />
                                            <%else %>
                                            <input type="radio" disabled <%=afvistSEL %> />
                                            <%end if %>
                                        <%else %>
                                            <%if cint(oRec("godkendt")) = 2 then %>
                                            <%=expence_txt_049 %>
                                            <%else %>
                                            <%=expence_txt_050 %>
                                            <%end if %>
                                        <%end if %>

                                        <%if oRec("gkdato") <> "" AND cint(oRec("godkendt")) = 2 then %>
                                        <span style="font-size:9px;"><br /> (<%=oRec("gkdato") %>)</span>
                                        <%end if %>
                                    </td>
                                    
                                    <td style="text-align:center;">
                                        <%if media <> "print" then %>
                                            <%if afregnetActiveInt = 1 then %>
                                            <input type="checkbox" class="afregnMat" name="afregnMat_<%=oRec("uploadid") %>" <%=afregnetSEL %> value="1" />
                                            <%else %>
                                            <input type="checkbox" disabled <%=afregnetSEL %> />
                                            <input type="hidden" name="afregnMat_<%=oRec("uploadid") %>" value="<%=oRec("afregnet") %>" />
                                            <%end if %>
                                        <%else %>
                                           <%if cint(oRec("afregnet")) <> 0 then %>
                                            <%=expence_txt_049 %>
                                            <%else %>
                                            <%=expence_txt_050 %>
                                            <%end if %>
                                        <%end if %>

                                        <%if oRec("afdate") <> "" AND cint(oRec("afregnet")) <> 0 then %>
                                        <span style="font-size:9px;"><br /> (<%=oRec("afdate") %>) </span>
                                        <%end if %>
                                    </td>

                                    <%if media <> "print" then %>
                                    <td style="text-align:center; vertical-align:middle;">
                                        <%if cint(oRec("afregnet")) = 0 then %>
                                        <a href="materialerudlaeg.asp?func=slet&matid=<%=oRec("uploadid") %>"><span style="color:darkred;" class="fa fa-times"></span></a>
                                        <%end if %>
                                    </td>
                                    <%end if %>

                                </tr>

                                <%
                                ekspTxt = ekspTxt & oRec("jobnavn") & aktnavneks &";"& oRec("matnavn") &";"& oRec("forbrugsdato") &";"& oRec("medarbnavn") & oRec("medarbinit") &";"& oRec("matgrpnavn") &";"& formatnumber(oRec("matkobspris"),2) &";"& valutakode &";"& formatnumber(oRec("basic_belob"), 2) &";"& oRec("basic_valuta") &";"& personExpTxt &";"& faktbartxt &";"& godkendtExpt &";"& afvistExpt &";"& afregnetExpt &"; xx99123sy#z"
                                antal = antal + 1
                                oRec.movenext
                                wend
                                oRec.close
                                %>

                                <%if antal = 0 then %>
                                <tr>
                                    <td colspan="20" style="text-align:center;"><%=expence_txt_051 %></td>
                                </tr>
                                <%end if %>
                            </tbody>

                        </table>

                         <%if media <> "print" then %>
                        <div class="row">
                            <div class="col-lg-12">
                                <button type="submit" class="btn btn-success btn-sm pull-right"><b><%=medarb_txt_020 %></b></button>
                            </div>
                        </div>
                        <%end if %>

                    </form>

                <%if media <> "print" then %>
                <div class="row">
                    <div class="col-lg-12"><b><%=expence_txt_052 %></b></div>
                </div>
                <form action="materialerudlaeg.asp?media=eksport&<%=strSogeKriterier %>" method="Post" target="_blank">
                  
                    <div class="row">
                     <div class="col-lg-12 pad-r30">                            
                     <input style="width:100px;" id="Submit5" type="submit" value="<%=expence_txt_053 %>" class="btn btn-sm" /><br />
                    <!--Eksporter viste kunder og kontaktpersoner som .csv fil-->                        
                    </div>
                    </div>
                </form>

                <form action="materialerudlaeg.asp?media=print" method="post" target="_blank">

                    <input type="hidden" name="FM_medarb" value="<%=request("FM_medarb") %>" />
                    <input type="hidden" name="FM_sog_txt" value="<%=request("FM_sog_txt") %>" /> 
                    <input type="hidden" name="FM_personlig" value="<%=request("FM_personlig") %>" />
                    <input type="hidden" name="FM_fromdateSearch" value="<%=request("FM_fromdateSearch") %>" />
                    <input type="hidden" name="FM_vis" value="<%=request("FM_vis") %>" />

                    <div class="row">
                        <div class="col-lg-12 pad-r30">
                            <input style="width:100px;" type="submit" value="<%=expence_txt_056 %>" class="btn btn-sm" />
                        </div>
                    </div>
                </form>
                <%end if %>

                </div>
            </div>
        </div>


        <%

        if media = "print" then
            Response.Write("<script language=""JavaScript"">window.print();</script>")                  
        end if





'************************************************************************************************************************************************
'******************* Eksport **************************' 

                if media = "eksport" then

                    call TimeOutVersion()
    
    
                    ekspTxt = ekspTxt 'request("FM_ekspTxt")
	                ekspTxt = replace(ekspTxt, "xx99123sy#z", vbcrlf)
	
	
	
	                filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	                filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
				                Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				                if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\to_2015\materialerudlaeg.asp" then
					                Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\expence_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					                Set objNewFile = nothing
					                Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\expence_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				                else
					                Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\expence_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					                Set objNewFile = nothing
					                Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\expence_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				                end if
				
				
				
				                file = "expence_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
				
				
				                '**** Eksport fil, kolonne overskrifter ***
				
			
				               ' strOskrifter = "Job / akt.; Navn; Dato; Medarbejder; type; Bel�b; Valuta; Basis bel�b; Basis valuta; Personlig; Godkendt; Afvist; Afregnet"
				                strOskrifter = expence_txt_054 &";"& expence_txt_006 &";"& expence_txt_039 &";"& expence_txt_005 &";"& expence_txt_040 &";"& expence_txt_007 &";"& expence_txt_008 &";"& expence_txt_041 &";"& expence_txt_057 &";"& expence_txt_011 &";"& expence_txt_013 &";"& expence_txt_045 &";"& expence_txt_047 &";"& expence_txt_048 &";"
				
				
				                objF.WriteLine(strOskrifter & chr(013))
				                objF.WriteLine(ekspTxt)
				                objF.close
				
				                %>
				                <div style="position:absolute; top:-100px; left:200px; width:300px; padding:20px;">
	                            <table border=0 cellspacing=1 cellpadding=0 width="100%">
	                            <tr><td valign=top bgcolor="#ffffff" style="padding:5px;">
	                            <img src="../ill/outzource_logo_200.gif" />
	                            </td>
	                            </tr>
	                            <tr>
	                            <td valign=top bgcolor="#ffffff" style="padding:5px 5px 5px 15px;">
                
                             
	                            <a href="../inc/log/data/<%=file%>" target="_blank" ><%=expence_txt_055 %> >></a>
	                            </td></tr>
	                            </table>
                                </div>
	            
	          
	            
	                            <%
                
                
                                Response.end
	                            'Response.redirect "../inc/log/data/"& file &""	

                end if  'export
 '************************************************************************************************************************************************             
          %>


        <%end select %>





<%if request("media") <> "print" then %>
    </div>
</div>
<%end if %>

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


    

<!--#include file="../inc/regular/footer_inc.asp"-->