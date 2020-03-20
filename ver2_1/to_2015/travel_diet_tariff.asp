<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->


<div class="wrapper">
    <div class="content">
        

<%
if len(session("user")) = 0 then

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
	
	select case func
	case "slet"
	'*** Her spørges om det er ok at der slettes  ***

        oskrift = "Rejse-Diæt" 
        slttxta = "Du er ved at <b>SLETTE</b> en Rejse-Diæt tarif. Er dette korrekt?"
        slttxtb = "" 
        slturl = "travel_diet_tariff.asp?func=sletok&id="& id

        call sletcnf_2015(oskrift, slttxta, slttxtb, slturl)



	case "sletok"
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM travel_diet_tariff WHERE id = "& id &"")
	Response.redirect "travel_diet_tariff.asp?menu=tok&shokselector=1"

    case "dbopr", "dbred"
	'*** Her indsættes en ny type i db ****

        useleftdiv = "to_2015"    

		if len(request("tdf_diet_name")) = 0 then
		
		errortype = 8
		call showError(errortype)
		
		else


        if len(trim(request("tdf_diet_dayprice"))) = 0 OR len(trim(request("tdf_diet_dayprice_half"))) = 0 OR _
        len(trim(request("tdf_diet_morgenamount"))) = 0 OR len(trim(request("tdf_diet_middagamount"))) = 0 OR _ 
        len(trim(request("tdf_diet_aftenamount"))) = 0 then

        errortype = 230
		call showError(errortype)

        Response.end
        end if

        tdf_diet_dayprice = replace(request("tdf_diet_dayprice"), ",", ".")
        tdf_diet_dayprice_half = replace(request("tdf_diet_dayprice_half"), ",", ".")
        tdf_diet_morgenamount = replace(request("tdf_diet_morgenamount"), ",", ".")
        tdf_diet_middagamount = replace(request("tdf_diet_middagamount"), ",", ".")
        tdf_diet_aftenamount = replace(request("tdf_diet_aftenamount"), ",", ".")

        tdf_diet_name = replace(request("tdf_diet_name"), "'", "")

        if len(trim(request("tdf_diet_current"))) <> 0 then
        tdf_diet_current = 1

                strSQlu = "UPDATE travel_diet_tariff SET tdf_diet_current = 0 WHERE id <> 0"
                oConn.execute(strSQLu)

        else
        tdf_diet_current = 0
        end if
		
		tdf_diet_editor = session("user")
		tdf_diet_dato = year(now)&"/"&month(now)&"/"&day(now)
		
        


     
		
		if func = "dbopr" then

        strSQL = "INSERT INTO travel_diet_tariff (tdf_diet_editor, tdf_diet_dato, tdf_diet_name, tdf_diet_current, tdf_diet_dayprice, tdf_diet_dayprice_half, tdf_diet_morgenamount, tdf_diet_middagamount, tdf_diet_aftenamount) "_
        &" VALUES ('"& tdf_diet_editor &"', '"& tdf_diet_dato &"', '"& tdf_diet_name &"', "& tdf_diet_current &", "& tdf_diet_dayprice &", "& tdf_diet_dayprice_half &", "& tdf_diet_morgenamount &", "& tdf_diet_middagamount &", "& tdf_diet_aftenamount &")"

        'Response.Write strSQL
        'response.flush

		oConn.execute(strSQL)


		else
	
            oConn.execute("UPDATE travel_diet_tariff SET tdf_diet_name ='"& tdf_diet_name &"', tdf_diet_editor = '"& tdf_diet_editor &"', tdf_diet_dato = '" & tdf_diet_dato &"', "_
            &" tdf_diet_current = "& tdf_diet_current &", tdf_diet_dayprice = "& tdf_diet_dayprice &", tdf_diet_dayprice_half = "& tdf_diet_dayprice_half &", tdf_diet_morgenamount = "& tdf_diet_morgenamount &", "_
            &" tdf_diet_middagamount = "& tdf_diet_middagamount &", tdf_diet_aftenamount = "& tdf_diet_aftenamount &" WHERE id = "&id&"")

		end if
		
		Response.redirect "travel_diet_tariff.asp"
		end if
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "Opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	strIkon = "gron"
	
	else



	        strSQL_tdf = "SELECT tdf_diet_editor, tdf_diet_dato, tdf_diet_name, tdf_diet_current, tdf_diet_dayprice, tdf_diet_dayprice_half, tdf_diet_morgenamount, tdf_diet_middagamount, tdf_diet_aftenamount FROM travel_diet_tariff WHERE id = " & id
            tdf_fundet = 0
            oRec6.open strSQL_tdf, oConn, 3
            if not oRec6.EOF then

                tdf_fundet = 1
                tdf_diet_dayprice = formatnumber(oRec6("tdf_diet_dayprice"), 2)
                tdf_diet_dayprice_half = formatnumber(oRec6("tdf_diet_dayprice_half"), 2)
                tdf_diet_morgenamount = formatnumber(oRec6("tdf_diet_morgenamount"), 2)
                tdf_diet_middagamount = formatnumber(oRec6("tdf_diet_middagamount"), 2)
                tdf_diet_aftenamount = formatnumber(oRec6("tdf_diet_aftenamount"), 2)

                tdf_diet_name = oRec6("tdf_diet_name")
                tdf_diet_current = oRec6("tdf_diet_current")

                tdf_diet_editor = oRec6("tdf_diet_editor")
                tdf_diet_dato = oRec6("tdf_diet_dato")

            end if
            oRec6.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "Opdaterpil" 
	end if
	%>
	 <!--Rediger-->

    <div class="container">
    <div class="portlet">
            <h3 class="portlet-title"><u>Rejse - Diæt tariffer</u></h3>
            <form action="travel_diet_tariff.asp?menu=tok&func=<%=dbfunc%>" method="post">
	        <input type="hidden" name="id" value="<%=id%>">
                <div class="row">
                    <div class="col-lg-10">&nbsp</div>
                    <div class="col-lg-2 pad-b10">
                    <button type="submit" class="btn btn-success btn-sm pull-right"><b>Opdatér</b></button>
                    </div>
                </div>

            <div class="portlet-body">
                <div class="well well-white">
               
                    
                 <div class="row">
                    <div class="col-lg-1">&nbsp</div>
                    <div class="col-lg-2">Navn: <font color=red size=2>*</font></div>
                    <div class="col-lg-2"><input type="text" class="form-control input-small" name="tdf_diet_name" value="<%=tdf_diet_name%>"></div>
                </div>

                      <div class="row">
                    <div class="col-lg-1">&nbsp</div>
                    <div class="col-lg-2">Dagspris: <font color=red size=2>*</font></div>
                    <div class="col-lg-2"><input type="text" class="form-control input-small" name="tdf_diet_dayprice" value="<%=tdf_diet_dayprice%>"></div>
                </div>

                      <div class="row">
                    <div class="col-lg-1">&nbsp</div>
                    <div class="col-lg-2">½ Dagspris: <font color=red size=2>*</font></div>
                    <div class="col-lg-2"><input type="text" class="form-control input-small" name="tdf_diet_dayprice_half" value="<%=tdf_diet_dayprice_half%>"></div>
                </div>

                      <div class="row">
                    <div class="col-lg-1">&nbsp</div>
                    <div class="col-lg-2">Beløb morgen: <font color=red size=2>*</font></div>
                    <div class="col-lg-2"><input type="text" class="form-control input-small" name="tdf_diet_morgenamount" value="<%=tdf_diet_morgenamount%>"></div>
                </div>

                      <div class="row">
                    <div class="col-lg-1">&nbsp</div>
                    <div class="col-lg-2">Beløb middag: <font color=red size=2>*</font></div>
                    <div class="col-lg-2"><input type="text" class="form-control input-small" name="tdf_diet_middagamount" value="<%=tdf_diet_middagamount%>"></div>
                </div>

                      <div class="row">
                    <div class="col-lg-1">&nbsp</div>
                    <div class="col-lg-2">Beløb aften: <font color=red size=2>*</font></div>
                    <div class="col-lg-2"><input type="text" class="form-control input-small" name="tdf_diet_aftenamount" value="<%=tdf_diet_aftenamount%>"></div>
                </div>



                <div class="row">
                    <div class="col-lg-1">&nbsp</div>
                     <%if cint(tdf_diet_current) = 1 then
                    tdf_diet_currentCHK = "CHECKED"
                    else
                    tdf_diet_currentCHK = ""
                    end if%>
                   <div class="col-lg-4"><input id="Checkbox1" name="tdf_diet_current" value="1" type="checkbox" <%=tdf_diet_currentCHK%> />&nbsp Aktuel tarif</div>
                </div>

                </div>

            </div>
        <%if dbfunc = "dbred" then%>
        <br /><br /><br /><div style="font-weight: lighter;">Sidst opdateret den <b><%=tdf_diet_dato%></b> af <b><%=tdf_diet_editor%></b></div>
        <%end if %>
        </form>

    </div>            
    </div>


<%case else %>
<!--Liste-->

    
        <div class="container">
        <div class="portlet">
            <h3 class="portlet-title"><u>Rejse - Diæt tariffer</u></h3>
            <form action="travel_diet_tariff.asp?menu=tok&func=opret" method="post">
            <section>
                         <div class="row">
                             <div class="col-lg-10">&nbsp;</div>
                             <div class="col-lg-2">
                                <button class="btn btn-sm btn-success pull-right"><b>Opret ny</b></button><br />&nbsp;
                            </div>
                 </div>
              </section>
             </form>

            <div class="portlet-body">

                <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                    <thead>
                        <tr>
                            <th>Id</th>
                            <th>Navn</th>
                            <th>Aktuel</th>
                            <th style="width: 5%">Slet</th>
                        </tr>
                    </thead>
                    <tbody>
                        <%
	                    
	                    strSQL = "SELECT id, tdf_diet_name, tdf_diet_current FROM travel_diet_tariff ORDER BY tdf_diet_name"
	                 
	
                        c = 1

	                    oRec.open strSQL, oConn, 3
	                    while not oRec.EOF 

                        select case right(c, 1)
                        case 0,2,4,6,8
                        bgt = "#EFF3FF"
                        case else
                        bgt = "#FFFFFF"
                        end select

	                    %>
                        <tr>
                            <td><%=oRec("id")%></td>
                            <td><a href="travel_diet_tariff.asp?menu=tok&func=red&id=<%=oRec("id")%>"><%=oRec("tdf_diet_name")%> </a></td>
                            <td>
                            <%if cint(oRec("tdf_diet_current")) = 1 then
                            %>
                            Ja
                            <%
                            end if %>
                            </td>
                            <td><a href="travel_diet_tariff.asp?menu=tok&func=slet&id=<%=oRec("id")%>"><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></a></td>
                        </tr>
                        <%
	                    x = 0
                        c = c + 1
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