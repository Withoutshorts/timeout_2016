

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<% 
'section for ajax calls
 if Request.Form("AjaxUpdateField") = "true" then
        Select Case Request.Form("control")
       
        case "FN_gemjobc"


         case "FN_vis_pas_medarb"

                usemrn = request("matreg_medid")
                vis_passive = request("vispassive_medarb")
                jqHTML = "1"

                
                call selmedarbOptions_2018

                '*** ÆØÅ **'
                call jq_format(strMedarbOptionsHTML)
                strMedarbOptionsHTML = jq_formatTxt

        response.write strMedarbOptionsHTML
        response.end


        case "FN_sogjobogkunde"

        call selectJobogKunde_jq()


        case "FN_sogakt"
              
        call selectAkt_jq


        case "indlasmat"
            
            jobid = request("jobid")
            aktid = request("aktid")
            medid = request("mid")
            aftid = 0
            matId = 0
            intAntal = 1 'Ved udlaeg er antal altid 1
            strDato = year(now)&"/"&month(now)&"/"&day(now)
            strEditor = session("user")
            regdato = year(now)&"/"&month(now)&"/"&day(now)
            bilagsnr = "" 
            navn = request("navn")
            FM_pris = request("FM_pris")
            FM_pris = replace(FM_pris, ".", ",")
            salgspris = 0
            varenr = 0
            valuta = request("valuta")
            gruppe = request("gruppe")
            personlig = request("personlig")
            intkode = request("intkode")
            betegnelse = ""
            opretlager = 0
            matregid = request("matregid")
            matava = 0
            otf = 1
            matreg_opdaterpris = 0

            mat_func = request("mat_func")

            unikId = request("unikId")

            call indlaes_mat(matregid, otf, medid, jobid, aktid, aftid, matId, strEditor, strDato, intAntal, regdato, valuta, intkode, personlig, bilagsnr, FM_pris, salgspris, navn, gruppe, varenr, opretlager, betegnelse, mat_func, matreg_opdaterpris, matava, unikId)

            
    end select
    response.End
    end if

%>



<script type="text/javascript" src="js/udlaeg_jav2.js"></script>

<div class="wrapper">
    <div class="content">

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

            call menu_2014
            call positiv_aktivering_akt_fn()

            select case func
            
            
            case "aproveExp"
                
            expenceid = split(request("expenceid"),", ")

            for i = 0 to UBOUND(expenceid)

                if len(trim(request("godkendmat_" & expenceid(i)))) <> 0 then
                    godkendtVal = request("godkendmat_" & expenceid(i))
                else
                    godkendtVal = 0
                end if

                afregnetPreVal = request("afregnetPreVal_" & expenceid(i))

                if len(trim(request("afregnMat_" & expenceid(i)))) then
                afregnetNewVal = 1
                else
                afregnetNewVal = 0
                end if

                preGKVal = request("godkendtPreVal_" & expenceid(i))

                'response.Write "<br> preval " & afregnetPreVal & " newval " & afregnetNewVal & " pregodkend " & preGKVal & " newgodkend " & godkendtVal

                if cint(godkendtVal) <> 0 AND cint(godkendtVal) <> cint(preGKVal) OR cint(afregnetNewVal) <> cint(afregnetPreVal) then
                    'response.Write "<br> godkendt Value " & godkendtVal 

                    strSQL = "UPDATE materiale_forbrug SET godkendt = "& godkendtVal & ", gkaf = '"& session("user") & "', gkdato = '"& year(now) &"-"& month(now) &"-"& day(now) &"', afregnet = "& afregnetNewVal &" WHERE id = "& expenceid(i) 
                    'response.Write "<br>" & strSQL
                    oConn.execute(strSQL)                    
                end if 

            next

            response.Redirect "materialerudlaeg.asp?FM_sog_jobnavn="&request("FM_sog_jobnavn")&"&FM_personlig="&request("FM_personlig")&"&FM_fromdateSearch="&request("FM_fromdateSearch")&"&FM_medarbejder="&request("FM_medarbejder")


            case "opr", "red"


                if cint(vis_passive) = 1 then
                chkPassive = "CHECKED"
                else
                chkPassive = ""
                end if               

                id = request("id")

                if func = "red" then
	    
                    dbfunc = "dbred"

	                oSkrkift = "Rediger"
                    fileLink = ""
                    strSQL = "SELECT mf.id as matregid, usrid, matnavn, matkobspris, mf.valuta, matgrp, personlig, intkode, filnavn, jobid, aktid, j.jobnavn as jobnavn, a.navn as aktnavn, unikid FROM materiale_forbrug as mf LEFT JOIN job as j ON (jobid = j.id) LEFT JOIN aktiviteter as a ON (mf.aktid = a.id) WHERE mf.id ="& id
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

                        if oRec("filnavn") <> "" then 
                            fileLink = "../inc/upload/"&lto&"/"&oRec("filnavn")
                        end if
                        
                        unikId = oRec("unikid")
                    

                    end if
                    oRec.close

	            else 'opr
	    

                    dbfunc = "dbopr"
                    oSkrkift = "Opret"
                    matregid = 0
                    usemrn = 0
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

                    uniksessionid = Session.SessionID & formatdatetime(now, 3)
                    uniksessionid = replace(uniksessionid, ":", "")
                    unikId = uniksessionid

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
                        <h3 class="portlet-title"><u>Udlæg - <%=oSkrkift %></u></h3>
                        <div class="portlet-body">
                            
                            <div class="well well-white">
                                <h4 class="panel-title-well"><%=medarb_txt_021 %></h4>
                           
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
                                        <input type="hidden" id="unikId" value="<%=unikId %>" />

                                        <div class="col-lg-6">
                                            <table style="display:inline-table;">

                                                <tr>
                                                    <td style="vertical-align:top;">Job <span style="color:red;">*</span></td>                                        
                                                    <td style="width:100%"><input type="text" id="FM_job" name="FM_job" placeholder="<%=tsa_txt_066 %>/<%=tsa_txt_236 %>" value="<%=jobnavn %>" class="FM_job form-control input-small" autocomplete="off" />

                                                        <select id="dv_job" size="10" class="form-control input-small chbox_job" style="visibility:hidden; display:none; width:100%;">
                                                            <option>Henter..</option>
                                                        </select>

                                                         <input type="hidden" id="FM_jobid" name="FM_jobid" value="<%=jobid %>"/>
                                                    </td>
                                                </tr>

                                                <tr>
                                                    <td style="vertical-align:top;">Aktivitet</td>
                                                    <td><input style="font-size:<%=inputFont%>; height:<%=inputHeight%>" type="text" id="FM_akt_0" name="activity" placeholder="<%=tsa_txt_068%>" value="<%=aktnavn %>" class="FM_akt form-control input-small" autocomplete="off" />

                                                        <select id="dv_akt_0" class="form-control input-small chbox_akt" size="10" style="visibility:hidden; display:none;"> <!-- chbox_akt -->
                                                            <option>Henter..</option>
                                                        </select>

                                                        <input type="hidden" id="FM_aktid_0" name="FM_aktivitetid" value="<%=aktid %>"/>
                                                    </td>
                                                </tr>

                                                <tr>
                                                    <td>Medarbejder <br /> 
                                                       <span style="font-size:9px; display:none;">Passive <input type="checkbox" value="1" id="FM_vis_passive" name="FM_vis_passive" <%=chkPassive %> /></span> 
                                                    </td>
                                                    <td id="td_selmedarb"><%call selmedarbOptions_2018 %></td>
                                                </tr>


                                                <tr>
                                                    <td>Navn</td>
                                                    <td><input type="text" name="FM_navn" id="FM_navn" class="form-control input-small" value="<%=matnavn %>" /></td>
                                                </tr>

                                                <tr>
                                                    <td>Beløb</td>
                                                    <td><input type="text" name="FM_belob" id="FM_belob" class="form-control input-small" value="<%=belob %>" style="width:75%" /></td>
                                                </tr>

                                                <tr>
                                                    <td>Valuta</td>
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
		                                                    <option value="<%=oRec3("id")%>" <%=valGrpCHK %>><%=oRec3("valutakode")%></option>
		                                                    <%
		                                                    oRec3.movenext
		                                                    wend
		                                                    oRec3.close %>
                                                        </select>
                                                    </td>
                                                </tr>

                                                <tr>
                                                    <td>Gruppe</td>
                                                    <td>
                                                        <select name="FM_gruppe" id="FM_gruppe" class="form-control input-small" style="width:75%">
                                                        <option value="0">Ingen</option>
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
                                                    <td>Personlig</td>
                                                    <td>
                                                        <select name="FM_personlig" id="FM_personlig" class="form-control input-small" style="width:75%">
                                                            <%
                                                            personSELCHB = ""
                                                            firmaSELCHB = ""

                                                            select case matpersonlig
                                                            case 1
                                                            personSELCHB = "SELECTED"
                                                            case 0
                                                            firmaSELCHB = "SELECTED"
                                                            end select
                                                            %>

                                                            <option value="1" <%=personSELCHB %>>Personlig</option>
                                                            <option value="0" <%=firmaSELCHB %>>Firma betalt</option>
                                                        </select>
                                                    </td>
                                                </tr>

                                                <tr>
                                                    <td>Faktbar</td>
                                                    <td>
                                                        <select name="FM_faktbar" id="FM_faktbar" class="form-control input-small" style="width:75%">
                                                            <%
                                                                intkodeSEL0 = ""
                                                                intkodeSEL1 = ""
                                                                intkodeSEL2 = ""

                                                                select case intkode
                                                                case 0
                                                                intkodeSEL0 = "SELECTED"
                                                                case 1
                                                                intkodeSEL1 = "SELECTED"
                                                                case 2
                                                                intkodeSEL2 = "SELECTED"
                                                                end select
                                                            %>

                                                            <%if lto <> "execon" AND lto <> "immenso" AND lto <> "unik" then %>
		                                                    <option value="0" <%=intkodeSEL0 %>><%=tsa_txt_235 %></option><!-- ingen -->
		                                                    <%end if%>
		    
		                                                    <option value="1" <%=intkodeSEL1 %>><%=tsa_txt_232 %></option><!-- intern -->
		                                                    <option value="2" <%=intkodeSEL2 %>><%=tsa_txt_233 %></option><!-- ekstern -->
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

                                     <form ENCTYPE="multipart/form-data" id="imageUpload" action="../timereg/upload_bin.asp?matUpload=1&unikId=<%=unikId %>&thisfile=<%=thisfile %>&<%="FM_sog_jobnavn="&request("FM_sog_jobnavn")&"&FM_personlig="&request("FM_personlig")&"&FM_fromdateSearch="&request("FM_fromdateSearch")&"&FM_medarbejder="&request("FM_medarbejder") %>" method="post">
                                         
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
                                                        <label class="btn btn-default btn-sm"><b><INPUT id="image-file" NAME="fileupload1" TYPE="file" style="display:none;" onchange="readURL(this);"> Vælg billede </b></label>
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

                            


                            <div class="well well-white" style="display:none;">
                                <h4 class="panel-title-well"><%=medarb_txt_021 %></h4>

                                <div class="row">

                                    <div class="col-lg-1"></div>

                                    <div class="col-lg-1">Navn</div>
                                    <div class="col-lg-3"><input type="text" name="FM_navn" class="form-control input-small" value="<%=matnavn %>" /></div>

                                    <div class="col-lg-3">Medarbejder 
                                        <%if level <= 2 OR level = 6 then %>
                                        &nbsp <input type="checkbox" value="1" id="FM_vis_passive" name="FM_vis_passive" <%=chkPassive %> /> Vis Passive
                                        <%end if %>
                                    </div>
                                    <div class="col-lg-3" id="">
                                        <%call selmedarbOptions_2018 %>
                                    </div>
                                </div>

                                <br />

                                <div class="row">
                                    <div class="col-lg-1"></div>
                                    <div class="col-lg-1">Antal:</div>
                                    <div class="col-lg-2"><input type="text" class="form-control input-small" value="<%=matantal %>" /></div>

                                    <div class="col-lg-1"></div>

                                    <div class="col-lg-3">Forbrugsdato</div>
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
                                    <div class="col-lg-1">Indkøbspris:</div>   
                                    <div class="col-lg-2"><input type="text" class="form-control input-small" /></div>

                                    <div class="col-lg-1"></div>
                                    <div class="col-lg-3">Valuta</div>
                                    <div class="col-lg-2">
                                        <select class="form-control input-small">
                                            <option>DKK</option>
                                        </select>
                                    </div>

                                </div>

                                <div class="row">
                                    <div class="col-lg-1"></div>
                                    <div class="col-lg-1">Salgspris:</div>     
                                    <div class="col-lg-2"><input type="text" class="form-control input-small" /></div>

                                    <div class="col-lg-1"></div>

                                    <div class="col-lg-3">Enhed:</div>
                                    <div class="col-lg-2"><input type="text" class="form-control input-small" /></div>

                                </div>

                                <div class="row">
                                    <div class="col-lg-1"></div>
                                    <div class="col-lg-1">Varenr: <br /></div>
                                    <div class="col-lg-1"><input type="text" class="form-control input-small" /></div>
                                    <div class="col-lg-2"><input type="checkbox" id="matreg_opretlager_0" name="opretlager" value="1" /> På lager</div>
       
                                    <div class="col-lg-3">Gruppe (avance %)</div>
                                    <div class="col-lg-2">
                                        <select class="form-control input-small">
                                            <option>Ingen</option>
                                        </select>
                                    </div>
                                </div>

                                <div class="row">
                                    <div class="col-lg-1"></div>
                                    <div class="col-lg-1">Intern kode:</div>
                                    <div class="col-lg-2">
                                        <select class="form-control input-small">
                                            <option>Ekstern</option>
                                        </select>
                                    </div>

                                    <div class="col-lg-1"></div>

                                    <div class="col-lg-3">Billagsnr</div>
                                    <div class="col-lg-2"><input type="text" class="form-control input-small" /></div>

                                </div>
                                

                            </div>

                        </div>


                    </div>
                </div>





        <%
        case else 

            'jobnr_sog = request("FM_sog_jobnr") 
            'jobnavnsog = request("FM_sog_jobnavn")
            if len(trim(request("FM_sog_txt"))) <> 0 then
                jobsogtxt = request("FM_sog_txt")
                sogKri = "AND (jobnavn LIKE '%"& jobsogtxt &"%' OR j.jobnr like '%"& jobsogtxt &"%')"
            else
                jobsogtxt = ""
            end if  
           
            if len(trim(request("FM_fromdateSearch"))) <> 0 then
                FM_fromdateSearch = request("FM_fromdateSearch")
            else
                FM_fromdateSearch = "01-"& month(now) &"-"& year(now)
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

            if request("FM_medarbejder") <> "0" then

                FM_medarbejder = split(request("FM_medarbejder"),", ")
                strFomr_rel = request("FM_medarbejder")
                strFomr_rel = replace(strFomr_rel, "X234", "#")
                antalM = 0

                for m = 0 to UBOUND(FM_medarbejder)
                    if m = 0 then
                        sogKri = sogKri & " AND (mf.usrid = "& FM_medarbejder(m)
                    else
                        sogKri = sogKri & " OR mf.usrid = "& FM_medarbejder(m)
                    end if
                    antalM = antalM + 1
                next

                if antalM > 0 then
                    sogKri = sogKri & ")"
                end if

                fomrIDs = replace(strFomr_rel, "#", ",")
                fomrIDs = replace(fomrIDs, ",,", ",")
                'fomrIDs = left(fomrIDs, len(fomrIDs)-1)

            else

                strFomr_rel = "#0#"
                FM_medarbejder = 0

            end if


            'response.Write "<br> <br><br><br><br>her her" & fomrIDs
            'response.Write sogKri
            strSogeKriterier = "FM_sog_txt="&request("FM_sog_txt")&"&FM_personlig="&request("FM_personlig")&"&FM_fromdateSearch="&request("FM_fromdateSearch")&"&FM_medarbejder="&fomrIDs

        %>

      <!-- <script type="text/javascript" src="js/udlaeg_liste_jav.js"></script> -->

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

        <%
            if media = "eksport" then
                displayExp = "none"
            end if
        %>

        <div class="container" style="display:<%=displayExp%>;">
            <div class="portlet">
                <h3 class="portlet-title"><u>Udlæg</u></h3>
                <div class="portlet-body">

                    <input name="FM_medarbejder" type="hidden" value="<%=fomrIDs %>" />

                    <form action="materialerudlaeg.asp?func=opr&<%=strSogeKriterier %>" method="post">
                        <div class="row">
                            <div class="col-lg-10">&nbsp</div>
                            <div class="col-lg-2 pad-b10">
                                <button type="submit" class="btn btn-success btn-sm pull-right"><b>Opret Udlæg</b></button>
                            </div>
                        </div>
                    </form>


                    <div class="well">
                        <form action="materialerudlaeg.asp" method="post">
                            <h4 class="panel-title-well">Søgefilter</h4>

                            <div class="row">
                                <div class="col-lg-3">Søg jobnavn eller jobnr.:</div>
                                <div class="col-lg-2">Personlig / Firmabetalt</div>
                                <div class="col-lg-2">Vis fra</div>
                            </div>

                            <div class="row">                                
                                <div class="col-lg-3"><input type="text" class="form-control input-small" name="FM_sog_txt" value="<%=jobsogtxt %>" /></div>

                                <div class="col-lg-2">

                                    <select class="form-control input-small" name="FM_personlig">

                                        <%
                                            SELall = ""
                                            SELpersonlig = ""
                                            SELfirma = ""

                                            select case FM_personlig
                                            case "-1"
                                            SELall = "SELECTED"
                                            case 1
                                            SELpersonlig = "SELECTED" 
                                            case 0
                                            SELfirma = "SELECTED"
                                            end select
                                        %>

                                        <option value="-1" <%=SELall %>">Alle</option>
                                        <option value="1" <%=SELpersonlig %>>Personlige</option>
                                        <option value="0" <%=SELfirma %>>Firma betalt</option>
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

                            </div>

                            <div class="row">
                                <div class="col-lg-3">Medarbejder:</div>
                            </div>

                            <div class="row">
                                <div class="col-lg-3">
                                    <select class="form-control input-small" name="FM_medarbejder" multiple size="7">
                                        <option value="0">Alle</option>
                                        <%
                                            strSQL = "SELECT mid, mnavn, init FROM medarbejdere WHERE mansat = 1"
                                            oRec.open strSQL, oConn, 3
                                            while not oRec.EOF

                                                if oRec("init") <> "" then
                                                medarbTxt = oRec("mnavn") &" ("& oRec("init") &")"
                                                else
                                                medarbTxt = oRec("mnavn")
                                                end if
                                                
                                                mSEL = ""
                                                 
                                                if request("FM_medarbejder") <> "0" then
                                                    for m = 0 to UBOUND(FM_medarbejder)

                                                        if mSEL = "" then
                                                            if cint(oRec("mid")) = cint(FM_medarbejder(m)) then
                                                                mSEL = "SELECTED"
                                                            end if
                                                        end if

                                                    next
                                                end if

                                                response.Write "<option value='"&oRec("mid")&"'"& mSEL &">"& medarbTxt &" </option>" 
                                            oRec.movenext
                                            wend
                                            oRec.close
                                        %>
                                    </select>
                                </div>
                            </div>

                            <div class="row">
                                <div class="col-lg-12"><button type="submit" class="btn btn-secondary btn-sm pull-right"><b>Vis</b></button></div>
                            </div>

                        </form>
                    </div>

                    <form action="materialerudlaeg.asp?func=aproveExp&<%=strSogeKriterier %>" method="post">

                        <div class="row">
                            <div class="col-lg-12">
                                <button type="submit" class="btn btn-success btn-sm pull-right"><b><%=medarb_txt_020 %></b></button>
                            </div>
                        </div>

                        <table id="expencetable" class="table dataTable table-striped table-bordered table-hover ui-datatable">

                            <thead>
                                <tr>
                                    <th>Job <br /> Aktivitet</th>
                                    <th>Navn</th>
                                    <th>Dato</th>
                                    <th>Medarbejder</th>
                                    <th>Type</th>
                                    <th>Beløb</th>
                                    <th>Personlig udlæg</th>
                                    <th>Billede</th>
                                    <th style="text-align:center;">Godkend <br /><input type="radio" name="godkenAfvisAll" id="godkendAll" /></th>                                
                                    <th style="text-align:center;">Afvis <br /><input type="radio" id="afvisAll" name="godkenAfvisAll" /></th>
                                    <th style="text-align:center;">Afregnet <br /> <input type="checkbox" id="afregnAlle" /></th>
                                </tr>
                            </thead>

                            <tbody>

                                <%
                                
                                ekspTxt = ""

                                strSQL = "SELECT mf.id as uploadid, matid, matnavn, matkobspris, matvarenr, forbrugsdato, jobid, usrid, matgrp, personlig, filnavn, mf.valuta, godkendt, gkaf, gkdato, afregnet, mf.aktid as mataktid, "_ 
                                & "j.jobnavn as jobnavn, j.id as projektid, "_ 
                                & "m.mnavn as medarbnavn, m.init as medarbinit, "_ 
                                & "mg.navn as matgrpnavn, a.navn as aktnavn "_
                                & "FROM materiale_forbrug as mf LEFT JOIN job as j ON (j.id = jobid) LEFT JOIN medarbejdere as m ON (m.mid = usrid) LEFT JOIN materiale_grp as mg ON (mg.id = matgrp) LEFT JOIN aktiviteter as a ON (mf.aktid = a.id) WHERE matid = 0 "& sogKri
                                'response.Write "<br> Sql" & strSQL
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

                                if oRec("afregnet") <> 0 then
                                    afregnetSEL = "CHECKED"
                                    afregnetExpt = "JA"
                                end if
                                
                                select case oRec("godkendt")
                                case 1
                                    godkendtSEL = "CHECKED"
                                    godkendtExpt = "JA"
                                    afvistSEL = ""
                                case 2
                                    afvistSEL = "CHECKED"
                                    afvistExpt = "JA"
                                    godkendtSEL = ""
                                end select


                                %>
                                <input type="hidden" value="<%=oRec("uploadid") %>" name="expenceid" />
                                <input type="hidden" value="<%=oRec("godkendt") %>" name="godkendtPreVal_<%=oRec("uploadid") %>" />
                                <input type="hidden" value="<%=oRec("afregnet") %>" name="afregnetPreVal_<%=oRec("uploadid") %>" />

                                <tr>
                                    <td><a href="jobs.asp?func=red&jobid=<%=oRec("projektid") %>"><%=oRec("jobnavn") %></a>
                                        <%if oRec("mataktid") <> 0 then %>
                                        <br /><span style="font-size:9px;"><%=oRec("aktnavn") %></span>
                                        <%end if %>
                                    </td>

                                    <td>
                                        <%if oRec("afregnet") = 0 then %>
                                            <%if level = 1 then %>
                                                <a href="materialerudlaeg.asp?func=red&id=<%=oRec("uploadid") %>&<%=strSogeKriterier %>"><%=oRec("matnavn") %> </a>
                                            <%else %>
                                                <%=oRec("matnavn") %>
                                            <%end if %>
                                        <%else %>
                                            <a href="materialerudlaeg.asp?func=red&id=<%=oRec("uploadid") %>&<%=strSogeKriterier %>"><%=oRec("matnavn") %> </a>
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

                                    <td>
                                        <%
                                        valutakode = ""
                                        strSQLValuta = "SELECT valutakode FROM valutaer WHERE id = "& oRec("valuta")
                                        oRec2.open strSQLValuta, oConn, 3
                                        if not oRec.EOF then
                                            valutakode = oRec2("valutakode")
                                        end if
                                        oRec2.close
                                        %>

                                        <%=formatnumber(oRec("matkobspris"),2) & " " & valutakode %>
                                    </td>

                                    <td>
                                        <%if cint(oRec("personlig")) <> 0 then 
                                            personExpTxt = "Ja"
                                        %>
                                        Ja
                                        <%else 
                                            personExpTxt = "Nej"
                                        %>
                                        Nej
                                        <%end if %>
                                    </td>

                                    <td style="text-align:center;">
                                    
                                        <%if oRec("filnavn") <> "" then %>
                                            <span id="modal_<%=oRec("uploadid") %>" class="fa fa-image picmodal"></span>
                                            <div id="myModal_<%=oRec("uploadid") %>" class="modal">
                                                <img src="../inc/upload/<%=lto %>/<%=oRec("filnavn") %>" alt='' border='0' style="width:20%;" class="modal-content">
                                            </div>

                                        <%else %>
                                        -
                                        <%end if %>

                                    </td>

                                    <td style="text-align:center;"><input type="radio" name="godkendmat_<%=oRec("uploadid") %>" <%=godkendtSEL %> value="1" class="godkendCHB" /></td>
                                    <td style="text-align:center;"><input type="radio" name="godkendmat_<%=oRec("uploadid") %>" <%=afvistSEL %> value="2" class="afvisCHB" /></td> 
                                    <td style="text-align:center;"><input type="checkbox" class="afregnMat" name="afregnMat_<%=oRec("uploadid") %>" <%=afregnetSEL %> /></td>
                                </tr>

                                <%
                                ekspTxt = ekspTxt & oRec("jobnavn") &";"& oRec("matnavn") &";"& oRec("forbrugsdato") &";"& oRec("medarbnavn") & oRec("medarbinit") &";"& oRec("matgrpnavn") &";"& formatnumber(oRec("matkobspris"),2) &" "& valutakode &";"& personExpTxt &";"& godkendtExpt &";"& afvistExpt &";"& afregnetExpt &"; xx99123sy#z"
                                oRec.movenext
                                wend
                                oRec.close
                                %>

                            </tbody>

                        </table>

                        <div class="row">
                            <div class="col-lg-12">
                                <button type="submit" class="btn btn-success btn-sm pull-right"><b><%=medarb_txt_020 %></b></button>
                            </div>
                        </div>

                    </form>

                <div class="row">
                    <div class="col-lg-12"><b>Funktioner</b></div>
                </div>
                <form action="materialerudlaeg.asp?media=eksport&<%=strSogeKriterier %>" method="Post" target="_blank">
                  
                    <div class="row">
                     <div class="col-lg-12 pad-r30">                        
                     <input id="Submit5" type="submit" value="Eksport til csv." class="btn btn-sm" /><br />
                    <!--Eksporter viste kunder og kontaktpersoner som .csv fil-->                        
                    </div>
                    </div>
                </form>

                </div>
            </div>
        </div>


        <%

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
				
			
				                strOskrifter = "Job / akt.; Navn; Dato; Medarbejder; type; Beløb; Personlig; Godkendt; Afvist; Afregnet"
				
				
				
				
				                objF.WriteLine(strOskrifter & chr(013))
				                objF.WriteLine(ekspTxt)
				                objF.close
				
				                %>
				                <div style="position:absolute; top:100px; left:200px; width:300px;">
	                            <table border=0 cellspacing=1 cellpadding=0 width="100%">
	                            <tr><td valign=top bgcolor="#ffffff" style="padding:5px;">
	                            <img src="../ill/outzource_logo_200.gif" />
	                            </td>
	                            </tr>
	                            <tr>
	                            <td valign=top bgcolor="#ffffff" style="padding:5px 5px 5px 15px;">
                
                             
	                            <a href="../inc/log/data/<%=file%>" target="_blank" >Din CSV. fil er klar >></a>
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






    </div>
</div>


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