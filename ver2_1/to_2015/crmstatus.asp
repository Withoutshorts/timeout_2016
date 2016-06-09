

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<%call menu_2014 %>

    <div id="wrapper">
        <div class="content">

    
<% 
    if len(session("user")) = 0 then
	%>
	
	<%
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
%>	 

<%
	select case func
	case "slet"
	'*** Her spørges om det er ok at der slettes en medarbejder ***
	    oskrift = "Aktions - statustyper" 
        slttxta = "Du er ved at <b>SLETTE</b> en status. Er dette korrekt?"
        slttxtb = "" 
        slturl = "crmstatus.asp?func=sletok&id="& id

        call sletcnf_2015(oskrift, slttxta, slttxtb, slturl)



    case "sletok"
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM crmstatus WHERE id = "& id &"")
	Response.redirect "crmstatus.asp?menu=tok&shokselector=1"



    case "dbopr", "dbred"
	'*** Her indsættes en ny type i db ****
		if len(request("FM_navn")) = 0 then
		%>
		    
		<%
		errortype = 8
		call showError(errortype)
		
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
		
		if func = "dbopr" then
		oConn.execute("INSERT INTO crmstatus (navn, editor, dato) VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"')")
		else
		oConn.execute("UPDATE crmstatus SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"' WHERE id = "&id&"")
		end if
		
		Response.redirect "crmstatus.asp?menu=tok&shokselector=1"
		end if
	
	case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	
	else
	strSQL = "SELECT navn, editor, dato FROM crmstatus WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "opdaterpil" 
	end if
	%>

<!--Redigere-->
<div class="container">
    <div class="portlet">
        <h3 class="portlet-title"><u>Aktions - statustyper <span style="font-size:60%"><%=varbroedkrumme%></span></u></h3>
        <div class="portlet-body">
        
        <form action="crmstatus.asp?func=<%=dbfunc%>" method="post">
	    <input type="hidden" name="id" value="<%=id%>">
    
            <div class="row">
                <div class="col-lg-10">&nbsp</div>
                <div class="col-lg-2 pad-b10">
                    <button type="submit" class="btn btn-success btn-sm pull-right"><b>Opdatér</b></button>
                </div>
            </div>

        
                <div class="well well-white">
            <div class="row">
                <div class="col-lg-1">&nbsp</div>
                <div class ="col-lg-1">Navn:</div>
                <div class="col-lg-4"><input name="FM_navn" type="text" class="form-control input-small" value="<%=strNavn %>" placeholder="Navn"/> </div>
            </div>
            </div>

            <%if func = "red" then %>
            <br /><br /><br /><div style="font-weight: lighter;">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></div>
           <%end if %>
        </form>
        </div>
    </div>
</div>

<%
case else    
%>

<script src="js/crmstatus_liste.js" type="text/javascript"></script>

<!--Liste-->
<div class="container">
    <div class="portlet">
        <h3 class="portlet-title"><u>Aktions - statustyper</u></h3>
        <div class="portlet-body">
            
            <form action="crmstatus.asp?menu=medarber&func=opret" method="post">
                    
            <div class="row">
                <div class="col-lg-10">&nbsp;</div>
                <div class="col-lg-2">
                    <button class="btn btn-sm btn-success pull-right"><b>Opret ny</b></button><br />&nbsp;
                </div>
            </div>
                    
             </form>

        
          
               <table id="crm_liste" class="table dataTable table-striped table-bordered table-hover ui-datatable">                  
                   <thead>
                       <tr>
                           <th style="width: 10%">id</th>
                           <th style="width: 40%">Status</th>
                           <th style="width: 5%">Slet</th>
                       </tr>
                   </thead>     
                   <tbody>
                       <%
                       strSQL = "SELECT id, navn FROM crmstatus ORDER BY navn"
	                    oRec.open strSQL, oConn, 3
	                    while not oRec.EOF
                        %>
                       <tr>
                           <td><%=oRec("id") %></td>
                           <td><a href="crmstatus.asp?func=red&id=<%=oRec("id") %>"><%=oRec("navn") %></a></td>
                           <td><a href="crmstatus.asp?menu=tok&func=slet&id=<%=oRec("id")%>"><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></a></td>
                       </tr>
                      <%oRec.movenext
                          wend
                          oRec.close %>
                   </tbody>  
               </table>

            
        
        </div>
    </div>
</div>

<%end select  %>

</div>
</div>

<!--#include file="../inc/regular/footer_inc.asp"-->