


<%response.buffer = true 
Session.LCID = 1030
%>
			        

<!--#include file="../inc/connection/conn_db_inc.asp"-->


<%'** JQUERY START ************************* %>

<%'** JQUERY END ************************* %>
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../timereg/inc/isint_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->



<%
if len(session("user")) = 0 then
	errortype = 5
	call showError(errortype)
        response.End
	end if
	

	
	'** Faste filter kri ***'
	thisfile = "aktiviteter"
	
	func = request("func")
	if func = "opret" then
	id = 0
	else
	id = request("id")
	end if
	
    if len(trim(id)) = 0 then
    id = 0
    end if

      
%>
    
    <%call menu_2014 %>
    <div id="wrapper">
        <div class="content">

    <%
  
    select case func 
    case "slet"

        oskrift = fomr_txt_001
        slttxta = fomr_txt_002 & " <b>"&fomr_txt_002&"</b> " & fomr_txt_004
        slttxtb = "" 
        slturl = "aktiviteter.asp?func=sletok&id="& id

        call sletcnf_2015(oskrift, slttxta, slttxtb, slturl)



	case "sletok"
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM aktiviteter WHERE id = "& id &"")

	Response.redirect "aktiviteter.asp"
%>


<%

    case "dbopr", "dbred"
	'*** Her indsættes en ny type i db ****
		if len(request("FM_navn")) = 0 then
		
		errortype = 8
		call showError(errortype)
		
		else

		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function
		
		aktNavn = SQLBless(request("FM_navn"))
        akt_fomr = SQLBless(request("FM_fomr"))
        aktfase =  SQLBless(request("FM_fase"))
        aktDesc = SQLBless(request("FM_desc"))

		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)

    


		if func = "dbopr" then
		oConn.execute("INSERT INTO aktiviteter (navn, editor, dato, job, fase, beskrivelse) VALUES "_
        &" ('"& aktNavn &"', '"& strEditor &"', '"& strDato &"', 4, '"& aktfase &"', '"& aktDesc &"')")
		
                aktidThis = 0
                strSQLa = "SELECT id FROM aktiviteter WHERE id <> 0 ORDER BY id DESC LIMIT 1"
                oRec.open strSQLa, oConn, 3
                if not oRec.EOF then

                aktidThis = oRec("id")

                end if
                oRec.close
                
        else
		oConn.execute("UPDATE aktiviteter SET navn ='"& aktNavn &"', editor = '" &strEditor &"', "_
        &" dato = '" & strDato &"', fase = '"& aktfase &"', beskrivelse = '"& aktDesc &"' WHERE id = "& id &"")
		end if

        
		
        if func = "dbred" then
        oConn.execute("DELETE FROM fomr_rel WHERE for_aktid = "& id)
        else
        id = aktidThis
        end if

        arrFomr = split(akt_fomr, ", ")
        for f = 0 TO UBOUND(arrFomr)

            oConn.execute("INSERT INTO fomr_rel (for_fomr, for_jobid, for_aktid, for_faktor) VALUES "_
            &" ("& arrFomr(f) &", 0, "& id &", 0)")

        next


		Response.redirect "aktiviteter.asp"
		end if



    case "red", "opret"
    '*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""

    strDato = ""
	strEditor = ""
    aktDesc = ""

    form = "" 'oRec("business_unit")
    strFase = ""

	
	dbfunc = "dbopr"
  

	else
	
    
    strSQL = "SELECT navn, editor, dato, fase, beskrivelse FROM aktiviteter WHERE id= " & id
    oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	aktNavn = oRec("navn")
	strDato = formatdatetime(oRec("dato"), 2)
	strEditor = oRec("editor")
    aktDesc = oRec("beskrivelse")

    form = "" 'oRec("business_unit")
    strFase = oRec("fase")

   

	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "Opdater" 
	end if
    
%>



<!--Aktiviteter redigering-->
<div class="container">
    <div class="portlet">
        <h3 class="portlet-title">
          <u>Activities</u>
        </h3>

        <form action="aktiviteter.asp?menu=tok&func=<%=dbfunc%>" method="post">
             <input type="hidden" name="id" value="<%=id%>">
        <div class="row">
            <div class="col-lg-10">&nbsp</div>
            <div class="col-lg-2 pad-b10">
            <button type="submit" class="btn btn-success btn-sm pull-right"><b><%=fomr_txt_005 %></b></button>
            </div>
        </div>

        <div class="portlet-body">
            
            <section>
              <div class="well well-white">  
           
            <div class="row">
                <div class="col-lg-1">&nbsp</div>
                <div class="col-lg-2">Theme:&nbsp<span style="color:red;">*</span></div>
                <div class="col-lg-3">

                     <%strSQLf = "SELECT for_fomr, for_faktor, f.navn AS fomrnavn FROM fomr_rel "_
                     &" LEFT JOIN fomr f ON (f.id = for_fomr)"_
                     &" WHERE for_aktid = " & id 

                         'Response.write strSQLf 
                         'Response.flush
                    formisWrt = " AND f.id <> 0"

                    %><select multiple name="FM_fomr" class="form-control input-small">
                         <%
                     oRec.open strSQLf, oConn, 3
                     while not oRec.EOF 

                             %><option value="<%=oRec("for_fomr") %>" selected><%=oRec("fomrnavn") %></option><%

                            formisWrt = formisWrt &  " AND f.id <> " & oRec("for_fomr")
                     oRec.movenext
                     Wend
                     oRec.close

                              
                         
                    strSQLf = "SELECT f.id, f.navn AS fomrnavn FROM fomr f "_
                    &" WHERE aktok = 1 "& formisWrt &" ORDER BY f.navn" 

                         'Response.write strSQLf 
                         'Response.flush
                    
                     oRec.open strSQLf, oConn, 3
                     while not oRec.EOF 

                             %><option value="<%=oRec("id") %>"><%=oRec("fomrnavn") %></option><%


                     oRec.movenext
                     Wend
                     oRec.close

                     %></select>

                </div>
            </div>
             <div class="row">
                <div class="col-lg-1">&nbsp</div>
                <div class="col-lg-2">Activity:&nbsp<span style="color:red;">*</span></div>
                <div class="col-lg-3"><input type="text" name="FM_navn" class="form-control input-small" value="<%=aktNavn %>" placeholder="Activity"/></div>
             
            </div>
                   <div class="row">
                <div class="col-lg-1">&nbsp</div>
                <div class="col-lg-2">Description:&nbsp<span style="color:red;">*</span></div>
                <div class="col-lg-6"><input type="text" name="FM_desc" class="form-control input-small" value="<%=aktDesc %>" placeholder="Description"/></div>
             
            </div>

           <div class="row">
                <div class="col-lg-1">&nbsp</div>
                <div class="col-lg-2">Intensity:</div>
                <div class="col-lg-3">

                    <%
                        
                        strIntensitetSEL0 = ""
                        strIntensitetSEL1 = ""
                        strIntensitetSEL2 = ""
                        strIntensitetSEL3 = ""

                        
                        select case strFase 
                        case "Ingen"
                        strIntensitetSEL0 = "SELECTED"
                        case "Lav intensitet"
                        strIntensitetSEL1 = "SELECTED"
                        case "Middel intensitet"
                        strIntensitetSEL2 = "SELECTED"
                        case "Høj intensitet"
                        strIntensitetSEL3 = "SELECTED"
                        end select
                        %>
                                        <select name="FM_fase" class="form-control input-small">
                                            <option value="Ingen" <%=strIntensitetSEL0 %>>Ingen</option>
                                            <option value="Lav intensitet" <%=strIntensitetSEL1 %>>Lav intensitet</option>
                                            <option value="Middel intensitet" <%=strIntensitetSEL2 %>>Middel intensitet</option>
                                            <option value="Høj intensitet" <%=strIntensitetSEL3 %>>Høj intensitet</option>
                                        </select>
                  
            </div>



           
           
               
            </section>
            <%if dbfunc = "dbred" then%> 
             <br /><br /><br /><div style="font-weight: lighter;"><%=fomr_txt_018 %> <b><%=strDato%></b> <%=fomr_txt_019 %> <b><%=strEditor%></b></div>
            <%end if  %>    
                </div>
             </form>
            </div>
        </div>
	
	
	





	<%
	

case else 
%>



<!--Aktiviteter liste-->
<div class="container">
    <div class="portlet">
        <h3 class="portlet-title">
            <u>Activities</u>
        </h3>
                
        <form action="aktiviteter.asp?menu=tok&func=opret" method="post">
          <section>
                         <div class="row">
                             <div class="col-lg-10">&nbsp;</div>
                             <div class="col-lg-2">
                          <button class="btn btn-sm btn-success pull-right"><b><%=fomr_txt_040 %> +</b></button><br />&nbsp;
                 <!--<a href="kunder.asp?menu=<%=menu%>&func=opret&ketype=<%=ketype%>&FM_soeg=<%=thiskri%>&medarb=<%=medarb%>" class="alt">Opret ny +</a>-->
                                     </div>
                 </div>
              </section>
         </form>

        <div class="portet-body">
            <table id="example" class="table dataTable table-striped table-bordered table-hover ui-datatable"> 
                <thead>
                    <tr>
                        <th>Theme</th>
                        <th>Activity</th>
                        <th>Description</th>
                        <th>Intensity</th>
                      
                        <th style="width: 5%"><%=fomr_txt_045 %></th>
                    </tr>
                </thead>
                <tbody>
                    <%
                        strSQL = "SELECT a.id, a.navn, fase, beskrivelse, "_
                        &" for_fomr, for_faktor, f.navn AS fomrnavn "_
                        &" FROM aktiviteter a "_
                        &" LEFT JOIN fomr_rel fl ON (fl.for_aktid = a.id)"_ 
                        &" LEFT JOIN fomr f ON (f.id = fl.for_fomr)"_ 
                        &" WHERE job = 4 GROUP BY a.id ORDER BY f.navn"

                        'Response.write strSQL
                        'response.flush

	                    oRec.open strSQL,oConn, 3
                        while not oRec.EOF 
                    %>
                    <tr>
                        <td>
                            <%'strSQLf = "SELECT for_fomr, for_faktor, f.navn AS fomrnavn FROM fomr_rel "_
                             '&" LEFT JOIN fomr f ON (f.id = for_fomr)"_
                             '&" WHERE for_aktid = " & oRec("id") 

                             'oRec2.open strSQLf, oConn, 3
                             'while not oRec2.EOF 

                                     %><%=oRec("fomrnavn") %><br /><%

                              
                             'oRec2.movenext
                             'Wend
                             'oRec2.close %>

                        </td>
                        <td><a href="aktiviteter.asp?func=red&id=<%=oRec("id") %>"><%=oRec("navn") %></a>
                        </td>
                        <td><%=oRec("beskrivelse") %></td>
                        <td><%=oRec("fase") %></td>
                        <td><a href="aktiviteter.asp?func=slet&id=<%=oRec("id")%>"><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></td>
                    </tr>
                    <%
					oRec.movenext
					wend
					oRec.close
					%>
                </tbody>
            </table>
        </div>
    </div>
</div>



<%end select %>
</div> <!--Content-->
</div> <!--Wrapper-->

<!--#include file="../inc/regular/footer_inc.asp"-->
