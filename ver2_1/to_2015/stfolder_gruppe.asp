
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
<script src="js/stfolder_liste.js" type="text/javascript"></script>

<%call menu_2014%>

<div class="wrapper">
    <div class="content">

    <%

	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	select case func
	case "slet", "sletfolder"
		
		
		gruppeid = request("gruppeid") 
		gruppenavn = request("gruppenavn")	
		
		if func = "sletfolder" then
		folderellergruppe = 1
		sletfunc = "sletfolderok"
		txt = "folder"
		else
		folderellergruppe = 0
		sletfunc = "sletok"
		txt = "standardfolder gruppe"
		end if
	
	'*** Her spørges om det er ok at der slettes en medarbejder ***
	%>
	<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	    <div class="container" style="width:500px;">
        <div class="porlet">        
            <h3 class="portlet-title">
               <u>Standardfolder grupper</u>
            </h3>
            
            <div class="portlet-body">
                <div align="center">Du er ved at <b>SLETTE</b> en <b><%=txt%></b>. Er dette korrekt?
                </div><br />
                <div align="center"><b><a class="btn btn-primary btn-sm" role="button" href="stfolder_gruppe.asp?menu=job&func=<%=sletfunc%>&id=<%=id%>&gruppeid=<%=gruppeid%>&gruppenavn=<%=gruppenavn%>"">&nbsp Ja &nbsp</a></b>&nbsp&nbsp&nbsp&nbsp<a class="btn btn-default btn-sm" role="button" href="stfolder_gruppe.asp?"><b>Nej</b></a>
                </div>
                <br /><br />
                </div>
            </div>
            </div>
	<%
	'*** Her slettes ***
	case "sletok", "sletfolderok"
	
	if func = "sletfolderok" then
	
	gruppeid = request("gruppeid") 
	gruppenavn = request("gruppenavn")	
	
	oConn.execute("DELETE FROM foldere WHERE id = "& id &"")
	Response.redirect "stfolder_gruppe.asp?menu=job&func=visgruppe&gruppeid="&gruppeid&"&gruppenavn="&gruppenavn&"&id=0"
			
	else
	
	oConn.execute("DELETE FROM folder_grupper WHERE id = "& id &"")
	Response.redirect "stfolder_gruppe.asp?menu=job"
	end if
	
	case "dbopr", "dbred"
	'*** Her indsættes en ny type i db ****
		if len(request("FM_navn")) = 0 then
		%>
		<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
		
		<%
		errortype = 8
		useleftdiv = "c"
		call showError(errortype)
		
		else
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function
		
		if len(request("folderellergruppe")) <> 0 then
		folderellergruppe = request("folderellergruppe")
		else
		folderellergruppe = 0
		end if
		
					if len(request("FM_kundese")) <> 0 then
					kundese = request("FM_kundese")
					else
					kundese = 0
					end if
			
					gruppeid = request("gruppeid") 
					gruppenavn = request("gruppenavn")	
		
		strNavn = SQLBless(request("FM_navn"))
		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
		
		if func = "dbopr" then
				if folderellergruppe = 1 then
					oConn.execute("INSERT INTO folder_grupper (navn, editor, dato) VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"')")
					
					'**** Finder id ****
					strSQL = "SELECT id FROM folder_grupper WHERE id <> 0 ORDER BY id DESC"
					oRec.open strSQL, oConn, 3
					if not oRec.EOF then
					
					id = oRec("id")
					
					end if
					oRec.close
				
				else
					strSQL = "INSERT INTO foldere (kundese, navn, editor, dato, stfoldergruppe) VALUES ("& kundese &", '"& strNavn &"', '"& strEditor &"', '"& strDato &"', "& gruppeid &")"
					oConn.execute(strSQL)
					
					
					'**** Finder id ****
					strSQL = "SELECT id FROM foldere WHERE id <> 0 ORDER BY id DESC"
					oRec.open strSQL, oConn, 3
					if not oRec.EOF then
					
					id = oRec("id")
					
					end if
					oRec.close
					
				end if
		else
				if folderellergruppe = 1 then
					oConn.execute("UPDATE folder_grupper SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"' WHERE id = "&id&"")
				else
					oConn.execute("UPDATE foldere SET kundese = "& kundese &", navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"' WHERE id = "&id&"")
				end if
		end if
		
			if folderellergruppe = 1 then
			Response.redirect "stfolder_gruppe.asp?menu=job&shokselector=1&id="&id
			else
			Response.redirect "stfolder_gruppe.asp?menu=job&func=visgruppe&gruppeid="&gruppeid&"&gruppenavn="&gruppenavn&"&id="&id
			end if
		end if
	
	
	
	
	case "opret", "red", "opretfolder", "redfolder"
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" OR func = "opretfolder" then
		
		strNavn = ""
		strTimepris = ""
		varSubVal = "Opret" 
		kundese = 0
		dbfunc = "dbopr"
		
	else
		
		if func = "redfolder" then
			
		strSQL = "SELECT kundese, navn, editor, dato FROM foldere WHERE id=" & id
		oRec.open strSQL,oConn, 3
		if not oRec.EOF then
		strNavn = oRec("navn")
		strDato = oRec("dato")
		strEditor = oRec("editor")
		kundese = oRec("kundese")
		end if
		oRec.close
		
		else
		
		
		strSQL = "SELECT navn, editor, dato FROM folder_grupper WHERE id=" & id
		oRec.open strSQL,oConn, 3
		if not oRec.EOF then
		strNavn = oRec("navn")
		strDato = oRec("dato")
		strEditor = oRec("editor")
		end if
		oRec.close
		
		end if
		
		dbfunc = "dbred"
		varSubVal = "Opdater" 

	end if
	
	
	if func = "opret" OR func = "red" then
		folderellergruppe = 1
			if func = "opret" then
			varbroedkrumme = " grupper -- Opret ny"
			else
			varbroedkrumme = " grupper -- Rediger"
			end if
		txt1 = "Gruppe:"
		formfelt2 = ""
	else
		folderellergruppe = 0
			
			if func = "opretfolder" then
			varbroedkrumme = " -- Opret ny"
			else
			varbroedkrumme = " -- Rediger"
			end if
		
		txt1 = "Folder:"
		
		if kundese = 1 then
		kundeseCHK = "CHECKED"
		else
		kundeseCHK = ""
		end if
		
		formfelt2 = "<br><input type='checkbox' name='FM_kundese' id='FM_kundese' value='1' "& kundeseCHK &"> Folder tilgængelig for eksterne kontakter."
		gruppeid = request("gruppeid")
	end if
	%>
	<!--REd-->
            <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u>Standardfolder grupper -- Rediger</u></h3>
                <form action="stfolder_gruppe.asp?menu=crm&func=<%=dbfunc%>&folderellergruppe=<%=folderellergruppe%>&gruppeid=<%=gruppeid%>&gruppenavn=<%=gruppenavn%>" method="post">
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
                        <div class="col-lg-2"><%=txt1%></div>
                        <div class="col-lg-2"><input type="text" class="form-control input-small" name="FM_navn"  value="<%=strNavn%>" /></div>
                    </div>
                   </div>
                    <%if dbfunc = "dbred" then%>
                    <br /><br /><br /><div style="font-weight: lighter;">Sidst opdateret den <b><%=strDato%></b> af <b><%=strEditor%></b></div>
                    <%end if %>
                </div>
                </form>
            </div>
            </div>

    <%case "visgruppe"
	
        gruppenavn = request("gruppenavn")
        gruppeid = request("gruppeid")
	
	
	
    %> 
        <!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
            <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u>Foldere i gruppen: <%=gruppenavn%></u></h3>
                <form action="stfolder_gruppe.asp?menu=job&func=opretfolder&gruppeid=<%=gruppeid%>&gruppenavn=<%=gruppenavn%>" method="post">
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
                    <table id="stfolderliste" class="table dataTable table-striped table-bordered table-hover ui-datatable">
                        <thead>
                            <tr>
                                <th>Id</th>
		                        <th>Folder</th>
		                        <th>Ekstern adgang?</th>
		                        <th>Slet</th>
                            </tr>
                        </thead>
                        <tbody>
                            <%
	                        strSQL = "SELECT id, navn, kundese FROM foldere WHERE stfoldergruppe = "& gruppeid &" ORDER BY navn"
	                        'Response.write strSQL 
	                        'Response.flush
	                        c = 0
	                        oRec.open strSQL, oConn, 3
	                        while not oRec.EOF 
	                        %>
                            <%
	                        if cint(id) = oRec("id") then
		                        bgc = "#ffff99"
	                        else
		                        select case right(c, 1)
		                        case 0, 2, 4, 6, 8
		                        bgc = "#ffffff"
		                        case else
		                        bgc = "#EFF3FF"
		                        end select
	
	                        end if
	                        %>
                            <tr>   
                               <td><%=oRec("id")%></td>
                               <td height="20"><a href="stfolder_gruppe.asp?menu=job&func=redfolder&gruppeid=<%=gruppeid%>&gruppenavn=<%=gruppenavn%>&id=<%=oRec("id")%>"><span class="fa fa-folder-o"></span>&nbsp;<%=oRec("navn")%></a></td>
                               <td><%if oRec("kundese") = 1 then%>
		                        <b><i>V</i></b>
		                        <%else%>
		                        &nbsp;
		                        <%end if%>
		                        </td>
                                <td><a href="stfolder_gruppe.asp?menu=job&func=sletfolder&gruppeid=<%=gruppeid%>&gruppenavn=<%=gruppenavn%>&id=<%=oRec("id")%>"><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></td>
                                </tr>
                                <%
	                        c = c + 1
	                        x = 0
	                        oRec.movenext
	                        wend
	                        %>
                        </tbody>
                    </table>
                </div>
            </div>
            </div>

            <!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--Liste-->
<%case else %>

            <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u>Standardfolder grupper</u></h3>
                <form action="stfolder_gruppe.asp?menu=job&func=opret"" method="post">
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
                    <table id="stfolderliste" class="table dataTable table-striped table-bordered table-hover ui-datatable">
                        <thead>
                            <tr>
                                <th>id</th>
                                <th>Gruppe</th>
                                <th>Foldere</th>
                                <th>Slet</th>
                            </tr>
                        </thead>
                       
                        <tbody>
                             <%
	                            strSQL = "SELECT id, navn FROM folder_grupper ORDER BY navn"
	                            c = 0
	                            oRec.open strSQL, oConn, 3
	                            while not oRec.EOF 
	                            %>
                                <%
	                            if cint(id) = oRec("id") then
		                            bgc = "#ffff99"
	                            else
		                            select case right(c, 1)
		                            case 0, 2, 4, 6, 8
		                            bgc = "#ffffff"
		                            case else
		                            bgc = "#EFF3FF"
		                            end select
	                            end if
	                            %>
                                <%
                                    intAntal = 0
					                '** Henter aktiviteter i den aktuelle gruppe ****
					                strSQL2 = "SELECT count(id) AS antal FROM foldere WHERE stfoldergruppe = "& oRec("id") 
					                oRec2.open strSQL2, oConn, 3
					                if not oRec2.EOF then
					                intAntal = oRec2("antal")
					                end if
					                oRec2.close 
					            %>
                            <tr>                               
                                <td><%=oRec("id")%></td>
                                <%if oRec("id") = 2 then%>
		                        <td><%=oRec("navn")%></td>
		                        <%else%>
		                        <td><a href="stfolder_gruppe.asp?menu=job&func=red&id=<%=oRec("id")%>"><%=oRec("navn")%></a></td>
		                        <%end if%>

                                <td>&nbsp;(<%=intAntal%>)&nbsp;<a href='stfolder_gruppe.asp?menu=job&func=visgruppe&gruppeid=<%=oRec("id")%>&gruppenavn=<%=oRec("navn")%>' class='vmenu' target="_blank"><span class="fa fa-external-link"></span></a></td>
                               
                                <%if oRec("id") <= 2 then%>
		                        <td>&nbsp</td>
		                        <%else%>
                                <%if cint(intAntal) = 1 then %>
			                    <td>&nbsp</td>   
                                <%else %>
                                <td><a href="stfolder_gruppe.asp?menu=job&func=slet&id=<%=oRec("id")%>"><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></a></td>
                                <%end if %>
                                <%end if %>
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