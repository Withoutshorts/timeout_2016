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
<!--include file="inc/budget_firapport_inc.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->





<%
    func = request("func")

	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if

    media = request("media")


    if media <> "eksport" then

    if media <> "print" then    
        call menu_2014
    end if

    %>

        <div id="wrapper">
        <div class="content">

    <%end if 'media



    select case func
    case "slet"
        '***Her spørges om det er ok at slette en medarbejder***

        oskrift = progrp_txt_049 
        slttxta = progrp_txt_001 &" "& progrp_txt_002 &" "& progrp_txt_003

        slttxtb = "" 
        slturl = "projektgrupper.asp?menu=tok&func=sletok&id="&id

        call sletcnf_2015(oskrift, slttxta, slttxtb, slturl)


    case "sletok" 
    '*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM projektgrupper WHERE id = "& id &"")
	Response.redirect "projektgrupper.asp?menu=job"
                

   
    case "opdprorel"
	
	z = 0
	x = Request("FM_antal_x")
	y = Request("FM_antal_y")
	
	'Response.write "x " & x & " y " & y & "<br>"
	strProjektgruppeId = request("FM_projektgruppeId")
	
	'** Finder jobid på job hvor denne projktgruppe har adgang til ***'
    guidjobids = ""
    guideasyids = ""
	call projktgrpPaaJobids(strProjektgruppeId)
	
	call positiv_aktivering_akt_fn() 'wwf
                         
	
		            
	'oConn.execute("DELETE FROM progrupperelationer WHERE projektgruppeId = "& strProjektgruppeId &"")
	
	
	For z = 0 to y - 1
	'Response.write z & "<br>"
	'Response.write "strMid: " & request("FM_Mid_"&z&"")& " , Optages: "& request("FM_medlem_"&z&"") & "<br>"
	'Response.flush
		
		strMid = request("FM_Mid_"&z&"")
		    
		    
		   if len(trim(strMid)) <> 0 then 
		    
		    fandtes = 0
		    strSQLfindes = "SELECT projektgruppeId, MedarbejderId"_
		    &" FROM progrupperelationer WHERE projektgruppeId = "& strProjektgruppeId & " AND MedarbejderId = "&strMid
		    
		    oRec4.open strSQLfindes, oConn, 3
		    if not oRec4.EOF then
		    
		    fandtes = 1
		    
		    end if
		    oRec4.close  
		    
		    'Response.Write "fandtes:" & fandtes
		    
		    '** Team leder **'
		    if request("FM_teamleder_"&z&"") = 1 then
    		teamleder = 1
    		else
    		teamleder = 0
    		end if
		    
            '** Notificer **'
		    if request("FM_notificer_"&z&"") = 1 then
    		notificer = 1
    		else
    		notificer = 0
    		end if


		    '** fandtes relation i forvejen ??***'
		    if cint(fandtes) = 1 then
		    
		        '*** Hvis ja, skal den slettes? ***'
		        if request("FM_medlem_"&z&"") <> "1" then
    		      oConn.execute("DELETE FROM progrupperelationer WHERE projektgruppeId = "& strProjektgruppeId &" AND MedarbejderId = "& strMid &"")
                  'Response.Write " -- DELETE FROM progrupperelationer WHERE projektgruppeId = "& strProjektgruppeId &"" & "<br

                  '*** Sletter fra timereg usejob og tilføjer igen **'
                  del = 1 '1 slet fra timereg_usejob / 0: tilføj kun
                  call tilfojtilTU(strMid, del)
                
                else
                 
                 oConn.execute("UPDATE progrupperelationer SET teamleder = "& teamleder &", notificer = "& notificer &" WHERE projektgruppeId = "& strProjektgruppeId &" AND MedarbejderId = "& strMid &"")  

                 end if
		    
		    else
		    
		         '*** skal den oprettes? ***
		         if request("FM_medlem_"&z&"") = 1 then
    		        
    		        
    		        
		            oConn.execute("INSERT INTO progrupperelationer (projektgruppeId, MedarbejderId, teamleder, notificer) VALUES ("& strProjektgruppeId &", "& strMid &", "& teamleder &", "& notificer &")")
		            
		            
		            '*** sætter aktiv i guiden aktiv job
                    'forvalgt = 0 'off / 1 = on
                    'Response.Write guideasyids
                    'Response.end
		            'call setGuidenUsejob(strMid, guidjobids, 1, guideasyids, forvalgt,2)
                    del = 0 '1 slet fra timereg_usejob / 0: tilføj kun
                    
                    if cint(positiv_aktivering_akt_val) <> 1 then
                    call tilfojtilTU(strMid, del)
		            end if
		            
		            
		         end if
		    
		    
		    
		    end if
		
		end if 'strmid <> 0
	
	Next
	
	               
	'Response.end
	
	Response.redirect "projektgrupper.asp?menu=job&lastid="&strProjektgruppeId

    %>
    
    <% case "med" 
        
     orgvir = 0   
     
        
        if media <> "eksport" then
         %>
        <!--Projektgruppe medlemmer-->
        <script src="js/projektgrupper_medl.js" type="text/javascript"></script>

        <div class="container">
            <div class="portlet">
            <h3 class="portlet-title"><u><%=progrp_txt_004 %></u></h3>

            <form action="projektgrupper.asp?menu=job&func=opdprorel" method="post">
	        <input type="hidden" name="FM_projektgruppeId" value="<%=id%>">
        

            <div class="row">
            <div class="col-lg-10">&nbsp</div>
            <div class="col-lg-2"><button type="submit" class="btn btn-success btn-sm pull-right"><b><%=progrp_txt_005 %></b></button></div>
            </div>
            <%

        end if


            Dim MedarbId()

            if id <> 0 then
            progrpSQL = " id = "& id
            else
            progrpSQL = " (id <> 0 AND id <> 10) "
            end if


	        strSQL3 = "SELECT id, navn, orgvir FROM projektgrupper WHERE "& progrpSQL & " ORDER BY navn"

            'response.Write "strSQL3: " & strSQL3
            'Response.flush

	        oRec3.open strSQL3, oConn, 3
	        while not oRec3.EOF 

	        proGruppeNavn = oRec3("navn")
            orgvir = oRec3("orgvir")

	      


                '***************************************
                if media <> "eksport" then
            
                    medIAndreOrgProjGrpStr = "#0#"
                    if cint(orgvir) = 1 then

                    '*** Tjekker Om Medarb er med i andre projgrp
                    strSQLmedIAndreOrgProjGrp = "SELECT p.id, m.mid FROM projektgrupper AS p "_
                    &" LEFT JOIN progrupperelationer AS pr ON (pr.ProjektgruppeId = p.id) "_
                    &" LEFT JOIN medarbejdere AS m ON (m.mid = pr.medarbejderid) "_
                    &" WHERE orgvir = 1 AND p.id <> "& id & " AND mid IS NOT NULL "
              
                        'response.write strSQLmedIAndreOrgProjGrp

                        oRec.open strSQLmedIAndreOrgProjGrp, oConn, 3
                        while not oRec.EOF  
             

                            medIAndreOrgProjGrpStr = medIAndreOrgProjGrpStr & ",#"& oRec("mid") &"#"
	                

                        oRec.movenext
                        wend 
                        oRec.close


                    end if

                end if 'media
                '***************************************

         if media <> "eksport" then
	        %>        

        <div class="portlet-body">
            <section>
            <p><%=progrp_txt_006 &" " %> <b><%=proGruppeNavn%></b></p>
           
            <!--Liste med medlemmer--> 
            <table class="table dataTable table-striped table-bordered table-hover ui-datatable">                  
               <thead>
                   <tr>
                       <th style="width: 45%"><%=progrp_txt_007 %></th>
                       <th style="width: 20%"><%=progrp_txt_008 %></th>
                       <th style="width: 5%"><%=progrp_txt_009 %></th>
                       <th style="width: 5%"><%=progrp_txt_010 %></th>
                       <th style="width: 5%"><%=progrp_txt_011 %></th>
                       <th style="width: 5%"><%=progrp_txt_012 %></th>
                       <th style="width: 10%"><%=progrp_txt_013 %></th>                       
                   </tr>
            </thead>
                    <% 

            end if 'media

                          

                            strSQL = "SELECT Mnavn, Mid, MedarbejderId, teamleder, init, mnr, mansat, notificer, lastlogin, ansatdato FROM progrupperelationer, medarbejdere"_
                            &" WHERE Mid = MedarbejderId AND ProjektgruppeId = "& oRec3("id") &" ORDER BY Mnavn"
	                        oRec.open strSQL, oConn, 3
	                        x = 0
	                        Redim MedarbId(x)
	                        while not oRec.EOF 
	
	               
	
	                        Redim preserve MedarbId(x)
	                        MedarbId(x) = oRec("Mid")

                            call mstatus_lastlogin
                                                   


                        if media <> "eksport" then
                          %>  
               
                       <tbody>
                           <tr>
                               <td><a href="medarb.asp?func=red&id=<%=oRec("Mid") %>"><%=oRec("Mnavn")%>
                                   <% 
                                    if len(trim(oRec("init"))) <> 0 then
                                    %>
                                    &nbsp;[<%=oRec("init") %>]
                                    <%end if    
                                    %>
                                   </a></td>

                               <td>
                                     <%=mstatus %>

                               </td>

                               <td><input type="checkbox" name="FM_medlem_<%=x%>" value="1" CHECKED>
		                       <input type="hidden" name="FM_Mid_<%=x%>" value="<%=oRec("Mid")%>"></td>
                               <%if cint(oRec("teamleder")) =  1 then
		                        tmCHK = "CHECKED"
		                        else
		                        tmCHK = ""
		                        end if %>
		
                                <%if cint(oRec("notificer")) =  1 then
		                        ntCHK = "CHECKED"
		                        else
		                        ntCHK = ""
		                        end if 
                               %>
                               <%if level = 1 then %>
		                        <td><input type="checkbox" name="FM_teamleder_<%=x%>" value="1" <%=tmCHK %>></td>
                                <td><input type="checkbox" name="FM_notificer_<%=x%>" value="1" <%=ntCHK %>></td>
		                        <%else%>
                               <td><input type="checkbox" name="" value="1" DISABLED <%=tmCHK %>>
                                <input id="Hidden1" name="FM_teamleder_<%=x%>" value="<%=oRec("teamleder") %>" type="hidden" /></td>
                                <td><input type="checkbox" name="" value="1" DISABLED <%=ntCHK %>>
                                <input id="Hidden1" name="FM_notificer_<%=x%>" value="<%=oRec("notificer") %>" type="hidden" /></td>
		                        <%end if %>

                               
                               <td style="white-space:nowrap;"><%=oRec("ansatdato") %></td>
                               <td style="white-space:nowrap;"><%=lastLoginDateFm %></td>
                           </tr>
                           <%
                           else

                               ekspTxt = ekspTxt & proGruppeNavn &";"& oRec("mnavn") &";"& oRec("init") &";"& mstatus &";"& oRec("teamleder") &";"& oRec("notificer") &";"& oRec("ansatdato") &";"& lastLoginDateFm &";xx99123sy#z"

                           end if 'media

                            x = x + 1
                           oRec.movenext
                            wend
                           oRec.close



                  oRec3.movenext
                  wend
	              oRec3.close




                 if media <> "eksport" then
                   %>
               </tbody>     
           </table>
                <%
             if media <> "eksport" AND x <> 0 then %>
            <%=progrp_txt_014&" " %> <%=x%><br /><br />
            <%end if %>

            <!--Ikke medlemmer liste-->
            <div class="pad-t10"><p><%=progrp_txt_015 %></p></div>
            <table id="tb_progrp_ikkemed" class="table dataTable table-striped table-bordered table-hover ui-datatable">                  
               <thead>
                   <tr>
                       <th style="width: 45%"><%=progrp_txt_007 %></th>
                       <th style="width: 20%"><%=progrp_txt_008 %></th>
                       <th style="width: 10%;"><input type="checkbox" value="0" id="FM_chk_all_add"  /> <%=progrp_txt_009&" " %> ?</th>
                       <th style="width: 5%"><%=progrp_txt_010 %> ?</th>
                       <th style="width: 5%"><%=progrp_txt_011 %></th>
                       <th style="width: 5%"><%=progrp_txt_012 %></th>
                       <th style="width: 10%"><%=progrp_txt_013 %></th>
                   </tr>
                </thead>
                   <%

                    '*** Henter ikke-medlemmer
	                Dim T
	                T = 0
	                For T = 0 to x - 1
	                strExclude = strExclude & "Mid <> "&MedarbId(T)&" AND "
	                Next
	
	                antChar = len(strExclude)
	                if antChar > 0 then
	                LantChar = left(strExclude, (antChar -5)) 
	                strExcludeFinal = "WHERE " & LantChar
	
	                strSQL = "SELECT Mnavn, Mid, init, mnr, mansat, lastlogin, ansatdato FROM medarbejdere "& strExcludeFinal &" AND (mansat <> '2') ORDER BY Mnavn"
	                else
	                strSQL = "SELECT Mnavn, Mid, init, mnr, mansat, lastlogin, ansatdato FROM medarbejdere WHERE mansat <> '2' ORDER BY Mnavn"
	                end if

                    'response.write "strSQL: " & strSQL & "<br>" 


	                oRec.open strSQL, oConn, 3

                    y = x
	                r = oRec.recordcount

	                while not oRec.EOF


                    call mstatus_lastlogin
                    %>
             
                    <tbody>
                   <tr>
                       <td><a href="medarb.asp?func=red&id=<%=oRec("Mid") %>"><%=oRec("Mnavn")%>
                           <% 
                            if len(trim(oRec("init"))) <> 0 then
                            %>
                            &nbsp;[<%=oRec("init") %>]
                            <%end if    
                            %>
                           </a></td>

                        <td>
                             <%=mstatus %>

                       </td>

                    

                       
                           
                      <%
                       '** Må ikke være med i mere end 1 org. projekgruppe ***'   
                       if cint(orgvir) = 0 OR instr(medIAndreOrgProjGrpStr, ",#"& oRec("mid") &"#") = 0 then %>
                               
                           <td><input type="checkbox" name="FM_medlem_<%=y%>" class="chk_medlem_clas" value="1">
                          
		                   <input type="hidden" name="FM_Mid_<%=y%>" value="<%=oRec("Mid")%>"></td>
                           <%if level = 1 then %>
		                    <td><input type="checkbox" name="FM_teamleder_<%=y%>" value="1"></td>
                            <td><input type="checkbox" name="FM_notificer_<%=y%>" value="1"></td>
		                    <%else %>
		                    <td><input type="checkbox" name="" value="0" DISABLED>
                            <input id="Text1" name="FM_teamleder_<%=y%>" value="0" type="hidden" /></td>

                            <td><input type="checkbox" name="" value="0" DISABLED>
                            <input id="Text1" name="FM_notificer_<%=y%>" value="0" type="hidden" /></td>
		                    <%end if %>
                           
                             <td style="white-space:nowrap;"><%=oRec("ansatdato") %></td>
                             <td style="white-space:nowrap;"><%=lastLoginDateFm %></td>

                       <%else %>
                       <td colspan="5" style="font-size:9px;"><%=progrp_txt_016 %></td>

                       <%end if %>

                   </tr>
                   <%
                   y = y + 1
                   oRec.movenext
                   wend
                   oRec.close
                   %>
               </tbody>     
             </table>
            </section>

                    <div style="margin-top:15px; margin-bottom:15px;">
                    <button type="submit" class="btn btn-success btn-sm pull-right"><b><%=progrp_txt_005 %></b></button>
                     
         
                    <div class="clearfix"></div>
                    </div>

               </div>
                    <input type="hidden" name="FM_antal_y" value="<%=y%>">
                    </form>
                </div>

      
    
            <% 
            end if 'media

            if media <> "eksport" then %>
            
                    <br /><br />

                      <section>
                            <div class="row">
                                 <div class="col-lg-12">
                                    <b><%=progrp_txt_017 %></b>
                                    </div>
                                </div>
                                <form action="projektgrupper.asp?media=eksport&func=med&id=<%=id%>" method="Post" target="_blank">
                  
                                <div class="row">
                                 <div class="col-lg-12 pad-r30">
                         
                                <input id="Submit5" type="submit" value="<%=progrp_txt_018 %>" class="btn btn-sm" /><br />
                         
                                     </div>


                            </div>
                            </form>
                
                        </section>    

                        <br /><br />



            <%end if


            '************************************************************************************************************************************************
            '******************* Eksport **************************' 
            if media = "eksport" then

                    call TimeOutVersion()
    
    
                    ekspTxt = ekspTxt 'request("FM_ekspTxt")
	                ekspTxt = replace(ekspTxt, "xx99123sy#z", vbcrlf)
	
	
                    'response.write ekspTxt
                    'response.Flush 

	
	                filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	                filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
				                Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				                if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\to_2015\projektgrupper.asp" then
					                Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\projektgrupperexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					                Set objNewFile = nothing
					                Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\projektgrupperexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				                else
					                Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\projektgrupperexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					                Set objNewFile = nothing
					                Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\projektgrupperexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				                end if
				
				
				
				                file = "projektgrupperexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
				
				
				                '**** Eksport fil, kolonne overskrifter ***
	                            strOskrifter = "Projektgruppe; Medarbejder; Init; Teamleder; Notificer; Status; Ansatdato; Sidst logget ind;"
				
				
				
				
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
                
                             
	                            <a href="../inc/log/data/<%=file%>" target="_blank" ><%=progrp_txt_026 %> >></a>
	                            </td></tr>
	                            </table>
                                </div>
	            
	          
	            
	                            <%
                
                
                                Response.end
	                            'Response.redirect "../inc/log/data/"& file &""	
				



                end if  
                '************************************************************************************************************************************************ %>

              </div>


<%
	
	case "dbopr", "dbred"
	'*** Her indsættes en ny type i db ****
		if len(request("FM_navn")) = 0 then
		%>
		<!--#include file="../inc/regular/header_inc.asp"-->
	
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
		
		if len(trim(request("FM_opengp"))) <> 0 then
		intopengp = 1
		else
		intopengp = 0
		end if

        if len(trim(request("FM_orgvir"))) <> 0 then
		orgvir = request("FM_orgvir")
		else
		orgvir = 0
		end if


        
		
		if func = "dbopr" then
		oConn.execute("INSERT INTO projektgrupper (navn, editor, dato, opengp, orgvir) VALUES ('"& strNavn &"', '"& strEditor &"', '"& strDato &"', "& intopengp &", "& orgvir &")")
		
			strSQL = "SELECT id FROM projektgrupper WHERE id <> 0 ORDER BY id DESC"
			oRec.open strSQL, oConn, 3
			if not oRec.EOF then
			
			lastid = oRec("id")
			
			end if
			
			oRec.close
		
		else
		
		oConn.execute("UPDATE projektgrupper SET navn ='"& strNavn &"', editor = '" &strEditor &"', dato = '" & strDato &"', opengp = "& intopengp &", orgvir = "& orgvir &" WHERE id = "&id&"")
		lastid = id
		
		end if


        if func = "dbopr" AND request("FM_tilfoj_m") = "1" then
        '*** tilføjer den der opretter gruppen som teamleder ***
        strSQLt = "INSERT INTO progrupperelationer (projektgruppeid, medarbejderid, teamleder) VALUES ("& lastId &", "& session("mid") &", 1)"
        oConn.execute(strSQLt)
        end if

		
		Response.redirect "projektgrupper.asp?menu=job&lastid="&lastid
		end if


 
    case "red", "opret"
    '*** Her indlæses form til rediger/oprettelse af ny type ***
	if func = "opret" then
	strNavn = ""
	strTimepris = ""
	varSubVal = "opretpil" 
	varbroedkrumme = "Opret ny"
	dbfunc = "dbopr"

     intOrgVir = 0
	
	else
	strSQL = "SELECT navn, editor, dato, opengp, orgvir FROM projektgrupper WHERE id=" & id
	oRec.open strSQL,oConn, 3
	
	if not oRec.EOF then
	
    strNavn = oRec("navn")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	intOpenGp = oRec("opengp")
    intOrgVir = oRec("orgvir")

	end if
	oRec.close
	
	dbfunc = "dbred"
	varbroedkrumme = "Rediger"
	varSubVal = "opdaterpil" 
	end if
	
 %>   






<!--redigere-->
<div class="container">
      <div class="portlet">
        <h3 class="portlet-title">
          <u><%=progrp_txt_027 %></u>
        </h3>

        <form action="projektgrupper.asp?menu=job&func=<%=dbfunc%>" method="post">
	    <input type="hidden" name="id" value="<%=id%>">

        <div class="row">
        <div class="col-lg-10">&nbsp</div>
        <div class="col-lg-2 pad-b10"><button type="submit" class="btn btn-success btn-sm pull-right"><b><%=progrp_txt_005 %></b></button></div>
        
        
        </div>



          <div class="portlet-body">
         <section>
             <div class="well well-white">
                    <div class="row">
                        <div class="col-lg-12">
                            <h4 class="panel-title-well"><%=progrp_txt_028 %></h4>
                        </div>
                    </div>

             <div class="row">
                 <div class="col-lg-1 pad-t10">&nbsp</div>
                 <div class="col-lg-2 pad-t10"><%=progrp_txt_007 %>:&nbsp<span style="color:red;">*</span></div>
                 <div class="col-lg-3 pad-t10">   
                  <input name="FM_navn" type="text" class="form-control input-small" value="<%=strNavn %>" placeholder="<%=progrp_txt_007 %>"/>
                 </div>
             </div>
                 
                 
                <%select case cint(intOrgVir) 
                case "1"
                intOrgVirCHK1 = "CHECKED"
                intOrgVirCHK0 = ""
                intOrgVirCHK2 = ""
                case "2"
	            intOrgVirCHK1 = ""
                intOrgVirCHK0 = ""
	            intOrgVirCHK2 = "CHECKED"
                case else
	            intOrgVirCHK1 = ""
                intOrgVirCHK0 = "CHECKED"
                intOrgVirCHK2 = ""
	            end select %>

                 <div class="row">
                     <div class="col-lg-1">&nbsp</div>
                     <div class="col-lg-4 pad-t20"><input name="FM_orgvir" value="1" type="radio" <%=intOrgVirCHK1 %>/>&nbsp;<b><%=progrp_txt_029&" " %></b> <%=progrp_txt_032 %>. <br />
                         <input name="FM_orgvir" value="0" type="radio" <%=intOrgVirCHK0 %>/>&nbsp;<b><%=progrp_txt_031&" " %></b> <%=progrp_txt_032 %>.<br />
                         <input name="FM_orgvir" value="2" type="radio" <%=intOrgVirCHK2 %>/>&nbsp;<b>HR</b> gruppe.<br />
                         <br />
                         <div style="background-color:#CCCCCC; padding:10px;"><%=progrp_txt_033&" " %> <u><%=progrp_txt_034 %></u> <%=" "& progrp_txt_035 %></div>
                     </div>
                 </div>


                 <%if intOpenGp = "1" then
	            intOpenGpCHK = "CHECKED"
	            else
	            intOpenGpCHK = ""
	            end if %>

                  <div class="row">
                     <div class="col-lg-1">&nbsp</div>
                     <div class="col-lg-4 pad-t20"><input id="Checkbox1" name="FM_opengp" value="1" type="checkbox" <%=intOpenGpCHK %>/>&nbsp;<%=progrp_txt_036 %></div>
                 </div>

                 <% if func = "opret" then %>
                 <div class="row">
                 <div class="col-lg-1">&nbsp</div>
                 <div class="col-lg-4"><input id="Checkbox2" name="FM_tilfoj_m" value="1" type="checkbox"  CHECKED/>&nbsp;<%=progrp_txt_037 %></div>
                 </div>
                 <% end if %>

           </div>
         </section>
         <%if func = "red" then %><br /><br /><br /><div style="font-weight: lighter;"><%=progrp_txt_038&" " %> <b><%=strDato%></b> <%=progrp_txt_039&" " %> <b><%=strEditor%></b></div><%end if %>
        </div>
        </form>
       </div>

</div>
<%
case else    


        if len(trim(request("lastid"))) <> 0 then
        lastid = request("lastid")
        else
        lastid = 0
        end if

%>



<script src="js/projektgrupper2.js" type="text/javascript"></script>
<input type="hidden" id="sogtekst" value="<%=progrp_txt_050 %>" />
<div class="container">
      <div class="portlet">
        <h3 class="portlet-title">
          <u><%=progrp_txt_049 %></u>
        </h3>

          <form action="projektgrupper.asp?menu=job&func=opret&id=0" method="post">
              <input type="hidden" name="lto" id="lto" value="<%=lto%>">
            <section>
              
                    
                         <div class="row">
                             <div class="col-lg-10">&nbsp;</div>
                             <div class="col-lg-2">
                          <button class="btn btn-sm btn-success pull-right"><b><%=progrp_txt_041 %> +</b></button><br />&nbsp;
                            </div>
                        </div>
            </section>
         </form>

          <div class="portlet-body">

              <table id="progrp_list" class="table dataTable table-striped table-bordered table-hover ui-datatable">                 
               <thead>
                   <tr>
                       <th style="width: 5%"><%=progrp_txt_042 %></th>
                       <th style="width: 55%"><%=progrp_txt_043 %></th>
                       <th style="width: 20%"><%=progrp_txt_044 %></th>
                       <th style="width: 8%"><%=progrp_txt_045 %></th>
                       <th style="width: 7%"><%=progrp_txt_046 %></th>
                       <th style="width: 5%"><%=progrp_txt_047 %></th>
                   </tr>
               </thead>
               <tbody>
                   <%
	                    strSQL = "SELECT id, navn, opengp, orgvir FROM projektgrupper WHERE id > 1 ORDER BY navn"
	                    oRec.open strSQL, oConn, 3
	
	                   
	                    n = 0
	                    while not oRec.EOF


                            
				                
                            
				                
	                         
                                call fTeamleder(session("mid"), oRec("id"))

                        if cint(lastid) = cint(oRec("id")) then
                       trBgCol = "#FFFFE1" 
                       else
                       trBgCol = ""
                       end if
                       
                      
                       antalMediPgrpX = 0
                       call antalMediPgrp(oRec("id"))
				     
                       
                       %>

                        <tr style="background-color:<%=trBgCol%>;">

                  
                       <td><%=oRec("id")%></td>
                       <td>
                           <%if oRec("id") <> "10" AND (oRec("opengp") = 1 OR level = 1 OR erTeamleder = 1) then%>
                           <a href="projektgrupper.asp?func=red&id=<%=oRec("id") %>"><%=oRec("navn") %></a>
                           <%else %>
                           <%=oRec("navn") %>
                           <%end if %>
                       </td>
                       <td>
                            <%if oRec("id") <> "10" AND (oRec("opengp") = 1 OR level = 1 OR erTeamleder = 1) then%>
                           <a href="projektgrupper.asp?menu=job&func=med&id=<%=oRec("id")%>"><%=progrp_txt_051 %> (<%=antalMediPgrpX%>) </a>
                           <%else %>
                           <%=progrp_txt_051 %> (<%=antalMediPgrpX %>)
                           <%end if %>

                       </td>
                            <td>
                                <%if cint(oRec("orgvir")) = 1 then 'ORGANI
                                 orgVir = progrp_txt_052   
                                 else
                                 orgVir = progrp_txt_053   
                                 end if %>

                                <%=orgVir %>

                            </td>
                       <td>
                        <%if oRec("opengp") = 0 then %>
	                    <%=progrp_txt_054 %>
	                    <%else %>
	                    <%=progrp_txt_055 %>
	                    <%end if %>
                       </td>
                       <td style="text-align:center;">
                           <%if oRec("id") <> "10" AND (level = 1 OR erTeamleder = 1) then%>
                           <a href="projektgrupper.asp?menu=tok&func=slet&id=<%=oRec("id")%>"><span style="color:darkred;" class="fa fa-times"></span></a>
                           <%else %>
                           &nbsp;
                           <%end if %>
                       </td>

                   </tr>
                <%

                    
	                n = n + 1
                    oRec.movenext
                    wend
                    oRec.close
                %>
               </tbody>    
                 <tfoot>
                      <tr>
                       <th><%=progrp_txt_042 %></th>
                       <th><%=progrp_txt_043 %></th>
                       <th><%=progrp_txt_044 %></th>
                       <th><%=progrp_txt_045 %></th>
                       <th><%=progrp_txt_046 %></th>
                       <th><%=progrp_txt_047 %></th>
                   </tr>

                 </tfoot>
                   
           </table>

          </div>

      </div>


      <br /><br />

          <section>
                <div class="row">
                     <div class="col-lg-12">
                        <b><%=progrp_txt_017 %></b>
                        </div>
                    </div>
                    <form action="projektgrupper.asp?media=eksport&func=med&id=0" method="Post" target="_blank">
                  
                    <div class="row">
                     <div class="col-lg-12 pad-r30">
                         
                    <input id="Submit5" type="submit" value="<%=progrp_txt_048 %>" class="btn btn-sm" /><br />
                         
                         </div>


                </div>
                </form>
                
            </section>   
              <br /><br />&nbsp;

</div>


<%end select  %>

</div>
</div>

<!--#include file="../inc/regular/footer_inc.asp"-->