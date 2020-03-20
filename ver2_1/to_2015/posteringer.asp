

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../timereg/inc/konto_inc.asp"-->
<!--#include file="../inc/regular/erp_func.asp"-->
<!--#include file="../timereg/inc/functions_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<%call menu_2014 %>

<div class="wrapepr">
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

        kundeid = request("kid")

        if len(request("kontonr")) <> 0 then
	     kontonr = request("kontonr")
	        
	        
	        %>
			<!--#include file="../timereg/inc/isint_func.asp"-->
			<%
			call erDetInt(kontonr)
			if isInt > 0 then

			errortype = 95
			call showError(errortype)
			
			isInt = 0
			Response.End
			end if
	              
	    else
	        kontonr = 0
	    end if

        if len(request("FM_soeg")) <> 0 then 
	    thiskri = request("FM_soeg")
	    useKri = 1
	    else
	    thiskri = ""
	    useKri = 0
	    end if

        call finddatoer()

        select case func


        case "bog"
	        oConn.execute("UPDATE posteringer SET status = 1 WHERE status = 0")
	        Response.redirect "posteringer.asp?menu=kon&id=0"

        case "slet"
            '*** Her spørges om det er ok at der slettes ***
	        oskrift = posteringer_txt_001 
            slttxta = posteringer_txt_002 & " <b>"&posteringer_txt_003&"</b> "& posteringer_txt_004
            slttxtb = "" 
            slturl = "posteringer.asp?menu=kon&func=sletok&id="&id&"&kontonr="&kontonr&"&oprid="&request("oprid")

            call sletcnf_2015(oskrift, slttxta, slttxtb, slturl)

        case "sletok"
	        '*** Her slettes ***
	        oConn.execute("DELETE FROM posteringer WHERE oprid = "& request("oprid") &"")
	        if kontonr <> 0 then
	        Response.redirect "posteringer.asp?menu=kon&id="&id&"&kontonr="&kontonr
	        else
	        Response.redirect "posteringer.asp?menu=kon&id=0"
	        end if

        case "dbopr", "dbred"

            response.Write "totaltotaltotaltotaltotaltotaltotaltotaltotaltotaltotaltotaltotaltotaltotaltotal " & len(request("FM_total"))
            if len(request("FM_total")) = 0 then	
		        useleftdiv = "t"
		        errortype = 50
		        call showError(errortype)
                response.End	
		    else			
			    call erDetInt(request("FM_total"))
			    if isInt > 0 then

			    useleftdiv = "t"
			    errortype = 49
			    call showError(errortype)
                response.End
                end if
			    isInt = 0
            end if

            if func = "dbred" then
		    opridOld = request("oprid")
		    strSQLDel = "DELETE FROM posteringer WHERE oprid = "& opridOld
		    oConn.execute(strSQLDel)
		    end if

            intkontonr = request("FM_kontonr_sel")
		    modkontonr = request("FM_modkonto_sel")
				
		    response.cookies("erp")("kontonr_1") = intkontonr
		    response.cookies("erp")("kontonr_2") = modkontonr

            function SQLBless(s)
		    dim tmp
		    tmp = s
		    tmp = replace(tmp, "'", "''")
		    SQLBless = tmp
		    end function
		
		    function SQLBless2(s)
		    dim tmp2
		    tmp2 = s
		    tmp2 = replace(tmp2, ",", ".")
		    SQLBless2 = tmp2
		    end function
		
		    function SQLBless3(s)
		    dim tmp3
		    tmp3 = s
		    tmp3 = replace(tmp3, ".", ",")
		    SQLBless3 = tmp3
		    end function

            strEditor = session("user")
		    strDato = year(now)&"/"&month(now)&"/"&day(now)


            vorref = request("FM_ref")
		    varBilag = request("FM_bilagsnr")
		    strTekst = SQLBless(request("FM_tekst"))

            if len(trim(request("FM_posdato"))) <> 0 then
		        posteringsdato = year(request("FM_posdato"))&"/"&month(request("FM_posdato"))&"/"&day(request("FM_posdato"))
            else
                posteringsdato = year(now)&"/"&month(now)&"/"&day(now)
            end if
		
		    kundeid = request("FM_kundeid")
		
		    intTotal = request("FM_total")
		
		    debkre = "pos"
		
		    rdirKontonr = request("rdirkontonr")
		
		    if len(request("status")) <> 0 AND request("status") <> 0 then
		    intStatus = 1
		    else
		    intStatus = 0
		    end if

            '*** Beregner moms Debit Konto ***'
	        call beregnmoms(debkre, intTotal, intkontonr)
	    
	        intNetto = SQLBless2(intNetto)
		    intMoms = SQLBless2(intMoms)
		    intTotal = SQLBless2(intTotal)
		
		
	        '**** Postering debit konto ****'
		    'call opretpos("1", func, modkontonr, intkontonr,intTotal_konto,intTotal_modkonto,intMoms,strEditor,varBilag,strTekst,posteringsdato,intStatus,vorref)
		    'response.Write "<br> tekst " & strTekst
		    oprid = 0
		    call opretPosteringSingle(oprid,"1", func, intkontonr, modkontonr, intNetto, intNetto, intMoms, strEditor, varBilag, strTekst, posteringsdato, intStatus, vorref)
		
		
		
		    '*** Beregner moms Kredit konto (Modkonto) ***'
	        call beregnmoms(debkre, intTotal, modkontonr)
	    
	   
		
		    intNetto = SQLBless2(-intNetto)
		    intMoms = SQLBless2(-intMoms)
		    intTotal = SQLBless2(-intTotal) 
		
		
		
	        '**** Postering kredit konto ****'
		    'call opretpos("1", func, modkontonr, intkontonr,intTotal_konto,intTotal_modkonto,intMoms,strEditor,varBilag,strTekst,posteringsdato,intStatus,vorref)

		
		    call opretPosteringSingle(oprid, "1", func, modkontonr, intkontonr, intNetto, intNetto, intMoms, strEditor, varBilag, strTekst, posteringsdato, intStatus, vorref)

            response.Redirect "posteringer.asp?menu=kon&kid="&kundeid&"&kontonr="&rdirKontonr&"&id="&id
            %>

          <!--  <script type="text/javascript">
		        window.opener.location.href = "posteringer.asp?menu=kon&kid=<%=kundeid%>&kontonr=<%=rdirKontonr%>&id=<%=id%>"
		        window.close();
		    </script> -->

		<%
            case "opret", "red"
	            '*** Her indlæses form til rediger/oprettelse af ny type ***
	            if func = "opret" then
	                strNavn = ""
	                intMomskode = 1
	                varSubVal = "Opretpil" 
	                varbroedkrumme = posteringer_txt_005
	                dbfunc = "dbopr"
	                intDag = day(now)
	                intMd = month(now)
	                intAar = year(now) 
	                intRef = session("Mid")
	                intBelob_opr = 0
	                oprid = 0
	
	                varBilagsnr = 0 
	                foundone = 0
	                strSQL = "SELECT bilagsnr FROM posteringer WHERE bilagsnr IS NOT NULL ORDER BY id DESC"
	                oRec.open strSQL, oConn, 3 
	                while not oRec.EOF AND foundone = 0
		
        
                        if IsNull(oRec("bilagsnr")) <> true AND len(trim(oRec("bilagsnr"))) <> "" then

                        call erDetInt(oRec("bilagsnr"))
		
		                if isInt > 0 then
		                varBilagsnr = varBilagsnr
		                else
		                varBilagsnr = oRec("bilagsnr") + 1  'Sidst oprettede 
		                foundone = 1
		                end  if

                        else

                        varBilagsnr = 0

                        end if
	
	                isInt = 0	
	                oRec.movenext
	                wend 
	                oRec.close
	
	                'if len(request("kontonr")) <> 0 AND request("kontonr") <> 0 then
	                'varKontonr = request("kontonr")
	                'else
	                varKontonr = request.cookies("erp")("kontonr_1")
	                'end if
	
	                varModkontonr = request.cookies("erp")("kontonr_2")
	
	                    if len(request("status")) <> 0 AND request("status") <> 0 then
		                intStatus = 1
		                else
		                intStatus = 0
		                end if
	
	            else
	                strSQL = "SELECT posteringer.dato AS dato, posteringer.editor AS editor, id, bilagstype, kontonr, modkontonr, bilagsnr, beloeb, nettobeloeb, moms, tekst, posteringsdato, status, mid, mnavn, mnr, oprid FROM posteringer LEFT JOIN medarbejdere ON (mid = att) WHERE posteringer.id = "& id &" ORDER BY posteringsdato"
	                oRec.open strSQL, oConn, 3
	
	                if not oRec.EOF then
	                varBilagsnr = oRec("bilagsnr")
	                intBeloeb = oRec("beloeb")
	                if intBeloeb < 0 then
	                intBelob_opr = intBeloeb
	                lenintBeloeb = len(intBeloeb)
	                rightintBeloeb = right(intBeloeb, (lenintBeloeb-1))
	                intBeloeb = rightintBeloeb
	                else
	                intBelob_opr = intBeloeb
	                intBeloeb = intBeloeb
	                end if
	
	                intNetto = oRec("nettobeloeb")
	                intMoms = oRec("moms")
	                strTekst = oRec("tekst")
	                posdato = oRec("posteringsdato")
	                intRef = oRec("mid")
	                strMnavn = oRec("mnavn") 
	                intMedmnr = oRec("mnr") 
	
	                intDag = day(posdato)
	                intMd = month(posdato)
	                intAar = year(posdato) 
	
	                strDato = oRec("dato")
	                strEditor = oRec("editor")
	                intStatus = oRec("status")
	
	                varKontonr = oRec("kontonr")
	                varModkontonr = oRec("modkontonr")
	
	                oprid = oRec("oprid")
	
	                end if
	                oRec.close
	
	                dbfunc = "dbred"
	                varbroedkrumme = posteringer_txt_040
	                varSubVal = "Opdaterpil" 
	            end if
        %>

        <script>
	        function fjerndot(){
	        document.all["FM_total"].value = document.all["FM_total"].value.replace(".",",")
	        }
	
	
	        function rensdag(){
	        document.getElementById("FM_dag").value = ""
	        }
	
	        function rensmd(){
	        document.getElementById("FM_md").value = ""
	        }
	
	        function rensaar(){
	        document.getElementById("FM_aar").value = ""
	        }
	
	        function bodyonload(){
	        document.getElementById("FM_total").focus();
	        }
	
	        function valgmodkont(){
	        document.getElementById("FM_modkonto").value = document.getElementById("FM_modkonto_sel").value
	        }
	
	        function valgkont(){
	        document.getElementById("FM_kontonr").value = document.getElementById("FM_kontonr_sel").value
	        //alert(document.getElementById("FM_kontonr_sel").selectedIndex)
	        }
	
	        function valgkont_sel(){
	
            var newval = document.getElementById("FM_kontonr").value
            var x=document.getElementById("FM_kontonr_sel")
            for(i=0; i<=250; i++)
                {
                    if(x.options[i].value == newval)
                    {
                        x.selectedIndex = i;
                        break;
                    }
                }
	        }
	
	        function valgmodkont_sel(){
            var newval = document.getElementById("FM_modkonto").value
	        var x=document.getElementById("FM_modkonto_sel")
            for(i=0; i<=250; i++)
                {
                    if(x.options[i].value == newval)
                    {
                        x.selectedIndex = i;
                        break;
                     }
                }
            }
	
	</script>


    <div class="container">
        <div class="portlet">
            <h3 class="portlet-title"><u><%=posteringer_txt_001 %> - <%=varbroedkrumme %> </u></h3>

            <div class="portlet-body">

                <form action="posteringer.asp?menu=kon&func=<%=dbfunc%>&status=<%=intStatus%>&rdirkontonr=<%=kontonr%>&oprid=<%=oprid%>" method="post">

                    <input type="hidden" name="id" value="<%=id%>">

                    <div class="well well-white">
                        <h4 class="panel-title-well"><%=posteringer_txt_007 %></h4>

                        <div class="row">
                            <div class="col-lg-1"></div>
                            <div class="col-lg-3"><%=posteringer_txt_008 %>:</div>
                            <div class="col-lg-2"><%=posteringer_txt_009 %>:</div>
                            <div class="col-lg-2"><%=posteringer_txt_010 %>:</div>
                            <div class="col-lg-3"><%=posteringer_txt_011 %></div>
                        </div>

                        <div class="row">

                            <div class="col-lg-1"></div>
                            <div class="col-lg-3">
                                <div class='input-group date' id='datepicker_stdato'>
                                    <input type="text" class="form-control input-small" name="FM_posdato" value="<%=day(now)&"-"&month(now)&"-"&year(now) %>" placeholder="dd-mm-yyyy" />
                                    <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                                </div>
                            </div>

                            <div class="col-lg-2"><input style="display:inline-block;" type="text" name="FM_total" id="FM_total" onkeyup="fjerndot();" value="<%=intBeloeb%>" class="form-control input-small"></div>

                            <div class="col-lg-2"><input type="text" name="FM_bilagsnr" value="<%=varBilagsnr%>" class="form-control input-small"></div>

                            <div class="col-lg-3">
                                <select name="FM_ref" class="form-control input-small">
		                            <%
		                            strSQL = "SELECT Mnavn, Mnr, Mid FROM medarbejdere ORDER BY mnavn"
		                            oRec.open strSQL, oConn, 3 
		                            while not oRec.EOF 
		
		                            if oRec("mid") = cint(intRef) then
		                            selth = "SELECTED"
		                            else
		                            selth = ""
		                            end if
		                            %>
		                            <option value="<%=oRec("mid")%>" <%=selth%>><%=oRec("mnavn")%> (<%=oRec("mnr")%>)</option>
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
                            <div class="col-lg-5"><%=posteringer_txt_012 %>:</div>
                            <div class="col-lg-4"><%=posteringer_txt_013 %>:</div>
                        </div>

                        <div class="row">
                            <div class="col-lg-1"></div>
                            <div class="col-lg-4">
                                <select name="FM_kontonr_sel" id="FM_kontonr_sel" class="form-control input-small" onchange="valgkont()";>
		                            <%	
                                        
		                            strSQL = "SELECT kontonr, kontoplan.navn AS kontonavn, kontoplan.id, kid, debitkredit, "_
		                            &" keycode, type, nk.navn, m.navn AS momskode FROM kontoplan "_
		                            &" LEFT JOIN nogletalskoder nk ON (nk.id = keycode) "_
		                            &" LEFT JOIN momskoder m ON (m.id = kontoplan.momskode) "_
		                            &" WHERE status = 1 ORDER BY kontonr, kontoplan.navn"
		
		                            'Response.Write strSQL
		                            'Response.Flush

		                            oRec.open strSQL, oConn, 3 
		                            while not oRec.EOF 
		
		                            if cstr(varKontonr) = cstr(oRec("kontonr")) then
		                            selkon = "SELECTED"
		                            else
		                            selkon = ""
		                            end if
		
		                            'if oRec("debitkredit") = 1 then
		                            'debitkredit = "Kre"
		                            'else
		                            'debitkredit = "Deb"
		                            'end if
		
		                            if oRec("type") = 1 then
		                            ktype = "Drift"
		                            else
		                            ktype = "Status"
		                            end if
		
		                            %>
		                            <option value="<%=oRec("kontonr")%>" <%=selkon%>>(<%=oRec("kontonr")%>)&nbsp;<%=oRec("kontonavn")%> - <%=ktype %> - <%=oRec("momskode") %></option>
		                            <%
		                            oRec.movenext
		                            Wend 
		                            oRec.close%>
		                        </select>
                            </div> 
                            <div class="col-lg-1"><input id="FM_kontonr" name="FM_modkontonr" value="<%=varKontonr%>" class="form-control input-small" type="text" onkeyup="valgkont_sel()" /></div>
                            <div class="col-lg-4">
                                <select name="FM_modkonto_sel" id="FM_modkonto_sel" class="form-control input-small" onchange="valgmodkont()">
		                            <%

                                    strSQL = "SELECT kontonr, kontoplan.navn AS kontonavn, kontoplan.id, kid, debitkredit, "_
		                            &" keycode, type, nk.navn, m.navn AS momskode FROM kontoplan "_
		                            &" LEFT JOIN nogletalskoder nk ON (nk.id = keycode) "_
		                            &" LEFT JOIN momskoder m ON (m.id = kontoplan.momskode) "_
		                            &" WHERE status = 1 ORDER BY kontonr, kontoplan.navn"

		                            oRec.open strSQL, oConn, 3 
		                            while not oRec.EOF 
		
		
		                            if cstr(varModKontonr) = cstr(oRec("kontonr")) then
		                            selkon = "SELECTED"
		                            else
		                            selkon = ""
		                            end if
		
		                            if oRec("debitkredit") = 1 then
		                            debitkredit = "Kre"
		                            else
		                            debitkredit = "Deb"
		                            end if
		
		                            if oRec("type") = 1 then
		                            ktype = "Drift"
		                            else
		                            ktype = "Status"
		                            end if
		
		                            %>
		                            <option value="<%=oRec("kontonr")%>" <%=selkon%>>(<%=oRec("kontonr")%>)&nbsp;<%=oRec("kontonavn")%> - <%=ktype %> - <%=oRec("momskode") %></option>
		                            <%
		                            oRec.movenext
		                            Wend 
		                            oRec.close%>
		                        </select>
                            </div>
                            <div class="col-lg-1"><input id="FM_modkonto" name="FM_modkonto" value="<%=varModKontonr%>" class="form-control input-small" type="text" onkeyup="valgmodkont_sel()" /></div>

                        </div>

                        <br /><br />

                        <div class="row">
                            <div class="col-lg-1"></div>
                            <div class="col-lg-3"><%=posteringer_txt_014 %>:</div>
                        </div>

                        <div class="row">
                            <div class="col-lg-1"></div>
                            <div class="col-lg-5">
                                <textarea id="FM_tekst" name="FM_tekst" class="form-control input-small" rows="3"><%=strTekst%></textarea>
                            </div>
                        </div>

                    </div>

                    <div class="row">
                        <div class="col-lg-12">
                            <button type="submit" class="btn btn-success btn-sm pull-right"><b><%=medarb_txt_020 %></b></button>
                        </div>
                    </div>


                </form>

                <%if func = "red" then %>

                <br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
                <div class="row">
                    <div class="col-lg-12"><%=posteringer_txt_015 %> <b><%=strDato%></b> <%=posteringer_txt_016 %> <b><%=strEditor%></b></div>
                 </div>

                <%end if %>

            </div>
        </div>
    </div>


    <%case else %>
    
    <div class="container">
        <div class="portlet">

            <%
                if kontonr = 0 then
	                bgthis = "#cccccc"
	                txt = posteringer_txt_020 & ":"
	                txt2 = posteringer_txt_019
	                ext = "_gray"
	                cspan = 12
	            else
		            select case kontonr
		            case "-1" 
		            bgthis = "pink"
		            txt = posteringer_txt_021&";"
		            txt2 = posteringer_txt_022 
		            ext = ""
		            cspan = 10
		            case "-2"
		            bgthis = "pink"
		            txt = posteringer_txt_021 & ":"
		            txt2 = posteringer_txt_023 
		            ext = ""
		            cspan = 10
		            case else
		            bgthis = "#5582d2"
		            txt = posteringer_txt_021 & ":"
		
		
		            knavn = ""
		            '*** Henter navn på konto ***'
		            strSQLknavn = "SELECT navn FROM kontoplan WHERE kontonr = " & kontonr
		            oRec.open strSQLknavn, oConn, 3
		            if not oRec.EOF then
		
		            knavn = oRec("navn")
		
		            end if
		            oRec.close
		
		
		            txt2 = posteringer_txt_024 & ": " & kontonr & " - "& knavn 
		            ext = ""
		            cspan = 12
		            end select
	            end if
	
	            if kontonr <> 0 then
	                intStatus = 1
	                else
	                intStatus = 0
	            end if
                    
                oprLink = "posteringer.asp?menu=kon&func=opret&kid="&kundeid&"&kontonr="&kontonr&"&status="&intStatus


                if len(trim(request("FM_startdato"))) <> 0 then
                    strStartDato = request("FM_startdato")
                    strStartDatoSQL = year(strStartDato) &"-"& month(strStartDato) &"-"& day(strStartDato) 
                    response.cookies("erp")("posStartDato") = strStartDatoSQL
                else
                    if request.Cookies("erp")("posStartDato") <> "" then
                        strStartDato = day(request.Cookies("erp")("posStartDato")) &"-"& month(request.Cookies("erp")("posStartDato")) &"-"& year(request.Cookies("erp")("posStartDato"))
                        strStartDatoSQL = year(strStartDato)&"-"&month(strStartDato)&"-"&day(strStartDato)
                    else
                        strStartDatoAdd = dateadd("m", -1, now)                                
                        strStartDato = day(strStartDatoAdd)&"-"&month(strStartDatoAdd)&"-"&year(strStartDatoAdd)
                        strStartDatoSQL = year(strStartDato)&"-"&month(strStartDato)&"-"&day(strStartDato)
                    end if
                end if

                if len(trim(request("FM_slutdato"))) <> 0 then
                strSlutDato = request("FM_slutdato")
                strSlutDatoSQL = year(strSlutDato) &"-"& month(strSlutDato) &"-"& day(strSlutDato)
                'response.cookies("erp")("posSlutDato") = strSlutDatoSQL
                else
                    if request.Cookies("erp")("posSlutDato") <> "" then
                        strSlutDato = day(request.Cookies("erp")("posSlutDato")) &"-"& month(request.Cookies("erp")("posSlutDato")) &"-"& year(request.Cookies("erp")("posSlutDato"))
                        strSlutDatoSQL = year(strSlutDato)&"-"&month(strSlutDato)&"-"&day(strSlutDato)
                    else
                        strSlutDato = day(now)&"-"&month(now)&"-"&year(now)
                        strSlutDatoSQL = year(strSlutDato)&"-"&month(strSlutDato)&"-"&day(strSlutDato)
                    end if
                end if
                   
                'response.Write "<br> her " & strStartDatoSQL & " - " & strSlutDatoSQL

	            %>

            <h3 class="portlet-title"><u><%=txt2 %></u></h3>

            <div class="portlet-body">
              
                <%if kontonr = 0 then %>
                <div class="row pad-b5">
                    <div class="col-lg-12"><a href="<%=oprLink %>" class="btn btn-success btn-sm pull-right"><b><%=posteringer_txt_025 %></b></a></div>
                </div>
                <%end if %>

                <form action="posteringer.asp?menu=kon&kid=<%=kundeid%>" method="post">
                    <div class="well">
                        <h4 class="panel-title-well"><%=posteringer_txt_026 %></h4>

                        <div class="row">
                            <div class="col-lg-3"><%=posteringer_txt_027 %>:</div>
                            <div class="col-lg-3"><%=posteringer_txt_028 %>:</div>
                            <div class="col-lg-2"><%=posteringer_txt_029 %>:</div>
                            <div class="col-lg-2"><%=posteringer_txt_030 %>:</div>
                        </div>

                        <div class="row">
                            <div class="col-lg-3"><input type="text" name="kontonr" id="kontonr" value="<%=kontonr%>" class="form-control input-small"></div>
                            <div class="col-lg-3"><input type="text" name="FM_soeg" id="FM_soeg" value="<%=thiskri%>" class="form-control input-small"></div>
                            <div class="col-lg-2">
                                <div class='input-group date' id='datepicker_stdato'>
                                    <input type="text" class="form-control input-small" name="FM_startdato" value="<%=strStartDato %>" placeholder="dd-mm-yyyy" />
                                    <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                                </div>
                            </div>
                            <div class="col-lg-2">
                                <div class='input-group date' id='datepicker_stdato'>
                                    <input type="text" class="form-control input-small" name="FM_slutdato" value="<%=strSlutDato %>" placeholder="dd-mm-yyyy" />
                                    <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                                </div>
                            </div>

                            <div class="col-lg-2"><button type="submit" class="btn btn-secondary btn-sm pull-right"><b><%=posteringer_txt_031 %></b></button></div>

                        </div>
                    </div>
                </form>

                <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                    <thead>
                        <tr>
                            <th><%=posteringer_txt_008 %></th>
                            <th><%=posteringer_txt_010 %></th>
                            <th><%=posteringer_txt_014 %></th>
                            <th><%=posteringer_txt_032 %></th>
                            <th><%=posteringer_txt_033 %></th>
                            <th><%=posteringer_txt_034 %></th>

                            <%if kontonr >= "0" then %>
	                        <th><%=posteringer_txt_035 %></th>
	                        <th><%=posteringer_txt_036 %></th>
	                        <%end if %>

                            <%if kontonr <= "0" then %>
                            <th><%=posteringer_txt_037 %></th>
	                        <%else %>
	                        <th></th>
                            <%end if %>
                            <th><%=posteringer_txt_038 %></th>	
	                        <th><%=posteringer_txt_018 %></th>
                            <th>&nbsp</th>
                        </tr>
                    </thead>

                    <tbody>

                        <%
	                    if useKri <> 0 then
	                    useSQLKri = " posteringer.bilagsnr = "& thiskri  
	                    else
	                    useSQLKri = " posteringer.id > 0 "
	                    end if
	
	                    'if usePeriodefilter = "j" then
	                    periodefilter = " AND posteringsdato BETWEEN '"& strStartDatoSQL &"' AND '"& strSlutDatoSQL &"'"
	                    'else
	                    'periodefilter = ""
	                    'end if
	
	                    intNetto_tot = 0
	                    intMoms_tot = 0
	                    intTotal_belob = 0
	
	                    select case kontonr
		                    case "-1", "-2"
		
			                    if kontonr = "-2" then
			                    kassekladeorkonto = "(posteringer.kontonr > 0 AND posteringer.status = 1 AND posteringer.moms < 0)"
			                    momskri = " k.kontonr = posteringer.kontonr" 'AND k.debitkredit = 2 AND k.type = 2
			                    momskri2 = " k2.kontonr = posteringer.modkontonr "
			                    end if
			
			                    if kontonr = "-1" then 'AND kontonr < 0 then
			                    kassekladeorkonto = "(posteringer.kontonr > 0 AND posteringer.status = 1 AND posteringer.moms > 0)"
			                    momskri = " k.kontonr = posteringer.kontonr"  ' AND k.debitkredit = 1 AND k.type = 1
			                    momskri2 = " k2.kontonr = posteringer.modkontonr "
			                    end if
			
			                    useSQLKri = useSQLKri &" AND k.type <> -1"
			
		                    case 0
			                    kassekladeorkonto = "(posteringer.kontonr > 0 AND posteringer.status = 0)"
			                    momskri = " k.kontonr = posteringer.kontonr "	
			                    momskri2 = " k2.kontonr = posteringer.modkontonr "	
		                    case else
			
			                    kassekladeorkonto = "(posteringer.kontonr = "& kontonr &" AND posteringer.status = 1)"
			                    momskri = " k.kontonr = posteringer.kontonr "
			                    momskri2 = " k2.kontonr = posteringer.modkontonr "
			
			
	                    end select 
	
	
	
	                    strSQL = "SELECT posteringer.id, posteringer.bilagstype, posteringer.kontonr, "_
	                    &" posteringer.modkontonr, posteringer.bilagsnr, posteringer.beloeb, posteringer.nettobeloeb, "_
	                    &" posteringer.moms, posteringer.tekst, posteringsdato, posteringer.status, mid, mnavn, mnr, "_
	                    &" posteringer.oprid, k.debitkredit AS debitkredit, k.navn AS kontonavn, k2.debitkredit AS debitkredit2, k2.navn AS modkontonavn FROM posteringer "_
	                    &" LEFT JOIN kontoplan k ON ("& momskri &")  "_
	                    &" LEFT JOIN kontoplan k2 ON ("& momskri2 &")  "_
	                    &" LEFT JOIN medarbejdere ON (mid = att) WHERE "& kassekladeorkonto &" "& periodefilter &" AND "& useSQLKri &" GROUP BY posteringer.id ORDER BY posteringsdato, id"
	
	                    'Response.write strSQL
	                    'Response.Flush
	
	                    oRec.open strSQL, oConn, 3
	                    c = 0
	                    debitTot = 0
	                    kreditTot = 0
	
	
	
	                    while not oRec.EOF 
	
	                    if cdbl(id) = oRec("id") then
	                    bgcol = "#ffffe1"
	                    else
	    
	                        select case right(c, 1)
	                        case 0,2,4,6,8
	                        bgcol = "#EFF3FF"
	                        case else
    	                    bgcol = "#ffffff"
	                        end select
    	
	
	
	                    end if%>

                        <tr>
                            <td><%=oRec("posteringsdato")%></td>

                            <%if kontonr = 0 AND oRec("nettobeloeb") > 0 then%>
		                    <td>
		                    <!--<a href="posteringer.asp?menu=kon&func=red&id=<%=oRec("id")%>&kontonr=<%=kontonr%>">-->
		                        <a href="posteringer.asp?menu=kon&func=red&id=<%=oRec("id")%>&kontonr=<%=kontonr%>" target="_blank"><%=oRec("bilagsnr")%></a>
		                    </td>	    
		                    <%else%>
		                    <td><%=oRec("bilagsnr")%></td>
		                    <%end if%>

                            <td><%=oRec("tekst")%></td>
                            <td><%=oRec("kontonr")%> - <%=oRec("kontonavn") %></td>
                            <td><%=oRec("modkontonr")%> - <%=oRec("modkontonavn") %></td>
                            <td><%=oRec("mnavn")%> (<%=oRec("mnr")%>)</td>

                            <%select case kontonr
		                        case "-1", "-2"%>
		                        <td style="text-align:right;"><%=formatnumber(oRec("moms"), 2)%></td>
		                        <%
		                        SaldoTot = saldoTot + (oRec("moms"))
		                        %>
		                        <td style="text-align:right;"><%=formatnumber(SaldoTot, 2) %></td>
                                <%
                                SaldoGrandTot = SaldoTot
		                        case else%>
		    
		                            <%if oRec("nettobeloeb") > 0 then %>
		                            <td style="text-align:right;">
		                            <%=formatnumber(oRec("nettobeloeb"), 2)%></td>
		                            <td style="text-align:right;">
                                    &nbsp;</td>
                                    <%
                                    debitTot = debitTot + oRec("nettobeloeb")
                                    else %>
                                    <td>&nbsp</td>
		                            <td style="text-align:right;">
                                    <%=formatnumber(oRec("nettobeloeb"), 2)%></td>
                                    <%
                                    kreditTot = kreditTot + oRec("nettobeloeb")
                                    end if %>
            
           
		                                        <%if kontonr = "0" then %>
						                         <td style="text-align:right;"><%=formatnumber(oRec("moms"), 2)%></td>
		                                        <%else %>
                                                 <td>&nbsp</td>
						                        <%end if%>
		    
		    
		                            <%
		                            momsTot = momsTot + (oRec("moms"))
		                            if kontonr = "0" then
		                            SaldoTot = debitTot + (kreditTot) + (momsTot)
		                            else
		                            SaldoTot = debitTot + (kreditTot)
		                            end if 
		    
		    
		                            %>
		                            <td style="text-align:right;"><%=formatnumber(SaldoTot, 2) %></td>
                                <%
		                        SaldoGrandTot = SaldoTot
		
		                    end select%>

                            <td style="text-align:right;"><%select case oRec("status") 
		                        case 0
		                        'Response.write "<font color=#cccccc>Kladde</font>"
                                Response.Write "Kladde"
		                        case 1
		                        Response.write "Bogført"
		                        case else
		                        Response.Write "Kladde"
		                        end select%>
                            </td>

                            <%if kontonr = 0 then%>
		                    <td><a href="posteringer.asp?menu=kon&func=slet&id=<%=oRec("id")%>&kontonr=<%=kontonr%>&oprid=<%=oRec("oprid")%>"><span style="color:darkred;" class="fa fa-times"></span></a></td>
		                    <%else%>
		                    <td>&nbsp;</td>
		                    <%end if%>

                        </tr>


                        <%
                            x = 0
	                        c = c + 1
	                        oRec.movenext
	                        wend
                        %>

                        <tr>
	                        <td colspan=6 style="text-align:right;"><b><%=txt%></b></td>
			
			                <%if kontonr = "-1" OR kontonr = "-2" then%>
			                <td>&nbsp</td>
			                    <td style="text-align:right;"><b><%=formatnumber(SaldoGrandTot, 2)%></b></td>
			                <%else%>							
						        <td style="text-align:right;"><b><%=formatnumber(debitTot, 2)%></b></td>
						        <td style="text-align:right;"><b><%=formatnumber(kreditTot, 2)%></b></td>
						        <%if kontonr = "0" then %>
						        <td style="text-align:right;"><b><%=formatnumber(momsTot, 2)%></b></td>
						        <%else %>
                                <td>&nbsp</td>
						    <%end if%>

					        <td style="text-align:right;"><b><%=formatnumber(SaldoGrandTot, 2)%></b></td>
						
			                <%end if%>
			
	                        <td colspan=3 align=right>&nbsp;
	                        <%if cdbl(formatnumber(SaldoGrandTot, 2)) = 0 AND kontonr = 0 then%>
	                        <a href="posteringer.asp?menu=kon&func=bog" class=vmenu><%=posteringer_txt_039 %> >></a>
	                        <%end if%></td>
	                    </tr>

                    </tbody>


                </table>


            </div>
        </div>
    </div>

    
    <%end select %>


</div>
</div>



<script type="text/javascript">

    $(document).ready(function () {

        $('.date').datepicker({

        });

});

</script>



<!--#include file="../inc/regular/footer_inc.asp"-->