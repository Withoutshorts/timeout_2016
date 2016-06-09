
<%response.buffer = true 
Session.LCID = 1030
%>
			        

<!--#include file="../inc/connection/conn_db_inc.asp"-->


<%'** JQUERY START ************************* %>

<%'** JQUERY END ************************* %>
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->



<%'** ER SESSION UDLØBET  ****
    if len(session("user")) = 0 then
	
	errortype = 5
	call showError(errortype)
    response.end
    end if %>

<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->





<%
   	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if  
    
    
     if media <> "print" then    
        call menu_2014
    end if
     
%>    

    
        <div id="wrapper">
            <div class="content">
        

    <%
    select case func
			
    case "slet"
    %>
    <!--slet sidens indhold-->
    <div class="container">
        <div class="porlet">
            
            <h3 align="center" class="portlet-title">
               Kundetyper
            </h3>
            
            <div class="portlet-body">
                <div align="center">Du er ved at <b>SLETTE</b> en kundetype. Er dette korrekt?
                </div><br />
                <div align="center"><b><a href="kundetyper.asp?menu=tok&func=sletok&id=<%=id%>">Ja</a></b>&nbsp&nbsp&nbsp&nbsp<b><a href="kundetyper.asp?">Nej</a></b>
                </div>
                <br /><br />
                </div>
            </div>
        </div>
    
    <%
    case "sletok"
        oConn.execute("DELETE FROM kundetyper WHERE id = "& id &"")
	    Response.redirect "kundetyper.asp?"
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
		
		strNavn = SQLBless(request("FM_navn"))
		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)

        ktlabel = SQLBless(request("FM_ktlabel"))
		
		intRabat = request("FM_rabat")
		
		if func = "dbopr" then
		oConn.execute("INSERT INTO kundetyper (navn, editor, dato, rabat, ktlabel) VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', "& intRabat &", '"& ktlabel &"')")
		else
		oConn.execute("UPDATE kundetyper SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"', rabat = "& intRabat &", ktlabel = '"& ktlabel &"' WHERE id = "&id&"")
		end if
		
		Response.redirect "kundetyper.asp?menu=tok&shokselector=1"
		end if


    case "opret", "red"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "Opret" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"
	intRabat = 0
	strktlabel = ""


	else
	strSQL = "SELECT navn, editor, rabat, dato, ktlabel FROM kundetyper WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	
	strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	intRabat = oRec("rabat")
    strktlabel = oRec("ktlabel")
	
	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "Opdater" 
	end if
	%>


<!--Redigering-->
<div class="container">
      <div class="portlet">
        <h3 class="portlet-title">
          <u>Kundetype/segment - <%=varbroedkrumme %></u>
        </h3>

          <form action="kundetyper.asp?func=<%=dbfunc%>" method="post">
	        <input type="hidden" name="id" value="<%=id%>">
            
              <section>
                  <div class="well well-white">
            <div class="row">
            <div class="col-lg-10">&nbsp</div>
            <div class="col-lg-2 pad-b10 pad-r30">
            <button type="submit" class="btn btn-success btn-sm pull-right"><b>Opdatér</b></button>
            </div>
            </div>

         
          <div class="row">
              <div class="col-lg-1 pad-t10">&nbsp</div>
              <div class="col-lg-2 pad-t10">Navn:</div>
              <div class="col-lg-3 pad-t10"><input class="form-control input-small" type="text" name="FM_navn" value="<%=strNavn%>" /></div>
              <div class="col-lg-1 pad-t10">Label:</div>
              <div class="col-lg-2 pad-t10"><input class="form-control input-small" type="text" name="FM_ktlabel" value="<%=strktlabel%>" /></div>
          </div>
          <div class="row">
              <div class="col-lg-1">&nbsp</div>
              <div class="col-lg-2">Rabat:</div>
              <div class="col-lg-1">
                  <%
        rSel0 = ""
        rSel10 = ""
        rSel15 = ""
        rSel25 = ""
        rSel30 = ""
        rSel40 = ""
        rSel50 = ""
        rSel60 = ""
        rSel75 = ""

		'**** husk også at rette på faktura oprettelse ***
		select case intRabat
		case 0
		rSel0 = "SELECTED"
        case 10
        rSel10 = "SELECTED"
        case 15
        rSel15 = "SELECTED"
        case 20
        rSel20 = "SELECTED"
        case 25
        rSel25 = "SELECTED"
        case 30
        rSel30 = "SELECTED"
        case 35
        rSel35 = "SELECTED"
        case 40
        rSel40 = "SELECTED"
        case 50
        rSel50 = "SELECTED"
        case 60
        rSel60 = "SELECTED"
        case 75
        rSel75 = "SELECTED"
		end select
		%>

            <select class="form-control input-small" id="FM_rabat" name="FM_rabat">
                <option value=0  <%=rSel0%>>0%</option>
                <option value=10 <%=rSel10%>>10%</option>
                <option value=15 <%=rSel15%>>15%</option>
                <option value=20 <%=rSel20%>>20%</option>
                <option value=25 <%=rSel25%>>25%</option>
                <option value=30 <%=rSel30%>>30%</option>
                <option value=35 <%=rSel35%>>35%</option>
                <option value=40 <%=rSel40%>>40%</option>
                <option value=50 <%=rSel50%>>50%</option>
                <option value=60 <%=rSel60%>>60%</option>
                <option value=75 <%=rSel75%>>75%</option>
             </select>

              </div>
          </div>
        </div>
        </section>
              <br /><br /><br />&nbsp;
              <%if func = "red" then %>
              <div style="font-weight: lighter;">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></div>
              <%end if %>
        </form>
       </div>
</div>


<% 
case else    
%>

<script src="js/kundetyper.js" type="text/javascript"></script>

<div class="container">
    <div class="porlet">
        <h3 class="portlet-title"><u>Kundetyper/segmenter</u></h3>
        
        <form action="kundetyper.asp?menu=kundetyper&func=opret" method="post">
                <section>
                         <div class="row">
                             <div class="col-lg-10">&nbsp;</div>
                             <div class="col-lg-2">
                            <button class="btn btn-sm btn-success pull-right"><b>Opret ny</b></button><br />&nbsp;
                            </div>
                        </div>
                </section>
         </form>

        <div class="porlet-body">
          
           <table id="kundetyper" class="table dataTable table-striped table-bordered table-hover ui-datatable">                 
               <thead>
                   <tr>
                       <th style="width: 5%">Id</th>
                       <th>Kundetype</th>
                       <th style="width: 25%">Label</th>
                       <th style="width: 15%">Rabat %</th>
                       <th style="width: 5%">Slet</th>
                   </tr>
                    </thead>
               <tbody>
    <%
	strSQL = "SELECT id, navn, rabat, ktlabel FROM kundetyper ORDER BY navn"


	oRec.open strSQL, oConn, 3
    while not oRec.EOF
	%>
              
               
                   <tr>
                       <td><%=oRec("id")%></td>
                       <td><a href="kundetyper.asp?menu=tok&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%> </a></td>
                       <td><%=oRec("ktlabel")%></td>
                       <td><%=oRec("rabat")%>%</td>
                       <td><a href="kundetyper.asp?menu=tok&func=slet&id=<%=oRec("id")%>"><span style="color:darkred;" class="fa fa-times"></span></a></td>
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

<%end select  %>
</div> 
</div>

<!--#include file="../inc/regular/footer_inc.asp"-->