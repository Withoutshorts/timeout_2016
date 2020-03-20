


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
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
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
        slturl = "fomr.asp?func=sletok&id="& id

        'call sletcnf_2015(oskrift, slttxta, slttxtb, slturl)



	case "sletok"
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM fomr WHERE id = "& id &"")
    oConn.execute("DELETE FROM fomr_rel WHERE for_fomr = "& id &"")

	Response.redirect "fomr.asp?menu=tok&shokselector=1"
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

		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)

    

		if func = "dbopr" then
		oConn.execute("INSERT INTO aktiviteter (navn, editor, dato, jobok, aktok, konto, business_unit, business_area_label, fomr_segment) VALUES "_
        &" ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', "& jobok &", "& aktok &", "& konto &", '"& business_unit &"', '"& business_area_label &"', '"& fomr_segment &"')")
		else
		oConn.execute("UPDATE aktiviteter SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"', jobok = "& jobok &", aktok = "& aktok &", "_
        &" konto = "& konto &", business_unit = '"& business_unit &"', business_area_label = '"& business_area_label &"', fomr_segment = '"& fomr_segment &"' WHERE id = "& id &"")
		end if
		
		Response.redirect "fomr.asp?menu=tok&shokselector=1"
		end if



    case "red", "opret"
    '*** Her indlæses form til rediger/oprettelse af ny type ***
	
    
    if func = "opret" then
	strNavn = ""
	
	dbfunc = "dbopr"
  
   
	else
	strSQL = "SELECT navn, editor, dato, fase FROM aktivieter WHERE id= " & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	aktNavn = oRec("navn")
	strDato = formatdatetime(oRec("dato"), 2)
	strEditor = oRec("editor")
    strFase = oRec("fase")


    form = "" 'oRec("business_unit")
  

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
                <div class="col-lg-2"><%=fomr_txt_007 %>:&nbsp<span style="color:red;">*</span></div>
                <div class="col-lg-3"><input type="text" name="FM_navn" id="FM_navn" class="form-control input-small" value="<%=strNavn %>"/></div>
             
            </div>
            <div class="row">
                <div class="col-lg-1">&nbsp</div>
                <div class="col-lg-2">Fomr:</div>
                <div class="col-lg-3"><input type="text" name="FM_fomr" id="FM_fomr" class="form-control input-small" value="<%=fomr %>" /></div>
            </div>
                   <div class="row">
                <div class="col-lg-1">&nbsp</div>
                <div class="col-lg-2">Fase:</div>
                <div class="col-lg-3"><input type="text" name="FM_fase" id="FM_fase" class="form-control input-small" value="<%=strFase %>" /></div>
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
            <u><%=fomr_txt_001 %></u>
        </h3>
                
        <form action="fomr.asp?menu=tok&func=opret" method="post">
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
                        <th><%=fomr_txt_041 %></th>
                        <th><%=fomr_txt_006 %></th>
                        <th><%=fomr_txt_042 %></th>
                        <%if cint(fomr_account) <> 1 then %>
                        <th><%=fomr_txt_014 %></th>
                        <%end if %>
                        <th><%=fomr_txt_043 %></th>
                        <th><%=fomr_txt_044 %></th>
                        <th style="width: 5%"><%=fomr_txt_045 %></th>
                    </tr>
                </thead>
                <tbody>
                    <%
                        strSQL = "SELECT navn FROM aktiviteter WHERE job = 4"
	                    oRec.open strSQL,oConn, 3
                        while not oRec.EOF 
                    %>
                    <tr>
                     
                        <td><a href="aktivieter.asp?func=red&id=<%=oRec("id") %>"><%=oRec("navn") %></a>
                           
                        <td><a href=">aktivieter.asp?menu=tok&func=slet&id=<%=oRec("id")%>"><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></td>
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
