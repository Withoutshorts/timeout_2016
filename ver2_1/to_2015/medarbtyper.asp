<%response.buffer = true 
Session.LCID = 1030


    'if session("mid") = 1 AND request("func") = "opret" then
    'Response.write "<br><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;HER"
    'Response.end
    'end if
%>

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<%'** JQUERY START ************************* %>

<%'** JQUERY END ************************* %>


<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->

<%
 '** ER SESSION UDL�BET  ****
    if len(session("user")) = 0 then
    errortype = 5
	call showError(errortype)
    response.end
    end if

 media = request("media")
 %>










<%
    
    
        


    if media <> "print" AND media <> "eksport" then    
        call menu_2014
    end if
    

     if media <> "eksport" then     %>


<div class="wrapper">
    <div class="content">
        

<% end if

'*** S�tter lokal dato/kr format. Skal inds�ttes efter kalender.
	Session.LCID = 1030
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if

   


	select case func
	case "slet"
	'*** Her sp�rges om det er ok at der slettes en medarbejder ***
	%>
	<!--Slet sidens indhold-->
    <div class="container" style="width:500px;">
    <div class="porlet">
            
            <h3 class="portlet-title">
               <%=medarbtyp_txt_001 %>
            </h3>
            
                <div class="portlet-body">
                    <div style="text-align:center;"><%=medarbtyp_txt_002 %> <b><%=medarbtyp_txt_003 %></b> <%=medarbtyp_txt_004 %>
                    </div><br />
                  
                        <div style="text-align:center;"><a class="btn btn-primary btn-sm" role="button" href="medarbtyper.asp?menu=tok&func=sletok&id=<%=id%>">&nbsp <%=medarbtyp_txt_049 %> &nbsp</a>&nbsp&nbsp&nbsp&nbsp<a class="btn btn-default btn-sm" role="button" href="Javascript:history.back()"><%=medarbtyp_txt_005 %></a>
                    </div>
                    <br /><br />
                 </div>
    </div>
    </div>


<%
	case "sletok"
	'*** Her slettes en medarbejder ***
	oConn.execute("DELETE FROM medarbejdertyper WHERE id = "& id &"")
	Response.redirect "medarbtyper.asp?menu=medarb"


	
    case "med"
                       
            ekspTxt = ""

        

            if level = 1 then


              

                       
                
                              mtypgrpLps = 0


                                if id <> 0 then
                                  mtypgrpSQL = " id = "& id
                                else
                                  mtypgrpSQL = " id <> 0"
                                end if

                                
                                strSQLtypegrupper = "SELECT id, type, timepris, timepris_a2, timepris_a3, timepris_a4, timepris_a5, kostpris, normtimer_man, normtimer_tir, normtimer_ons, normtimer_tor, normtimer_fre, normtimer_lor, normtimer_son FROM medarbejdertyper mt WHERE "& mtypgrpSQL &" ORDER BY type"
	
                                'Response.write strSQLtypegrupper
                                'Response.flush
    
                                oRec3.open strSQLtypegrupper, oConn, 3
                                while not oRec3.EOF 
	                            mTypegruppeNavn = oRec3("type")
	                            


                              if media <> "eksport" AND mtypgrpLps = 0 then  %>

                            <!--Medlemmer-->
                            <div class="container">
                            <div class="portlet">
                                <h3 class="portlet-title"><u><%=medarbtyp_txt_006 %></u></h3>

                                <div><%=medarbtyp_txt_007 %> <b><%=mTypegruppeNavn%></b></div><br />
                                <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                                    <thead>
                                        <tr>
                                            <th><%=medarbtyp_txt_008 %></th>
                                            <th><%=medarbtyp_txt_009 %></th>
                                            <th><%=medarbtyp_txt_010 %></th>
                                            <!--<th>Email</th>-->
                                            <th><%=medarbtyp_txt_011 %></th>
                                            <th><%=medarbtyp_txt_012 %></th>
                                            <th><%=medarbtyp_txt_013 %></th>
                                            <!--<th>Opsagtdato</th>-->
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <%
                                    end if 'media


                        
                            mTypeSQlkri = " medarbejdertype = "& oRec3("id") 
                           

	                        strSQL = "SELECT mnavn, mid, mansat, init, mnr, ansatdato, opsagtdato, medarbejdertype, lastlogin, email "_
                            &" FROM medarbejdere "_
                            &" WHERE "& mTypeSQlkri &" AND mansat <> 2 GROUP BY mid ORDER BY medarbejdertype, mansat, mnavn"
	                        oRec.open strSQL, oConn, 3
	

                            'response.write strSQL & "<br>"
                            'response.flush

                            x = 0
	                        while not oRec.EOF 



                                                     call mstatus_lastlogin


                                                        
                                                     lastTregDato = ""   
                                                     strSQLlastTreg = "SELECT tdato FROM timer WHERE tmnr = "& oRec("mid") & " ORDER BY tdato DESC LIMIT 1"
                                                     oRec4.open strSQLlastTreg, oConn, 3
                                                     if not oRec4.EOF then
                                                        lastTregDato = oRec4("tdato")
                                                     end if
                                                     oRec4.close 
                                                     
                                                    


                            if media <> "eksport" then

                                    select case right(x, 1) 'if oRec("mansat") = 3 then
                                    case 0,2,4,6,8
                                    bgthis = "#EFF3FF"
                                    case else
                                    bgthis = "#FFFFFF"
                                    end select

	                                        %>
                                            <tr>
                                                <td><a href="medarb.asp?menu=medarb&func=red&id=<%=oRec("Mid")%>"><%=oRec("mnavn")%> [<%=oRec("init") %>]</a></td>
                                                <td><%=oRec("init") %></td>
                                                <td><%=mstatus %></td>
                                                <!--<td><%=oRec("email") %></td>-->
                                                <td><%=oRec("ansatdato") %></td>
                                                <!--<td>
                                                    <%if cDate(oRec("opsagtdato")) <> "1-1-2044" then %>
                                                    <%=oRec("opsagtdato") %>
                                                    <%opsagtdatoTxt = oRec("opsagtdato")
                                                    else  
                                                    opsagtdatoTxt = ""  
                                                    %>
                                                    <%end if %>

                                                </td>
                                                -->
                                                <td><%=lastLoginDateFm %></td>
                                                    <td><%=lastTregDato%></td>
                                            </tr>
                                            <%

                                    else
                                                '";"& opsagtdatoTxt
                                                ekspTxt = ekspTxt & mTypegruppeNavn & ";" & oRec("mnavn") & ";" & oRec("init") & ";"& mstatus &";" & oRec("ansatdato") &";"& lastLoginDateFm &";"& lastTregDato &";"
                                                ekspTxt = ekspTxt & oRec3("timepris") &";"& oRec3("timepris_a2") &";"& oRec3("timepris_a3") &";"& oRec3("timepris_a4") &";"& oRec3("timepris_a5") &";"& oRec3("kostpris") &";"
                                                ekspTxt = ekspTxt & oRec3("normtimer_man") &";"& oRec3("normtimer_tir") &";"& oRec3("normtimer_ons") &";"& oRec3("normtimer_tor") &";"& oRec3("normtimer_fre") &";"& oRec3("normtimer_lor") &";"& oRec3("normtimer_son") &";xx99123sy#z"
                                    end if


                                    x = x + 1
                                    oRec.movenext
                                    wend 
                                    oRec.close


                            oRec3.movenext
                            wend
	                        oRec3.close
                       
                                        
                    if media <> "eksport" AND x <> 0 then                    
                    %>
                    </tbody>
                </table>
              <%=medarbtyp_txt_014 %>: <%=x%>
            </div>

                    <br /><br />

                      <section>
                            <div class="row">
                                 <div class="col-lg-12">
                                    <b><%=medarbtyp_txt_015 %></b>
                                    </div>
                                </div>
                                <form action="medarbtyper.asp?media=eksport&func=med&id=<%=id%>" method="Post" target="_blank">
                  
                                <div class="row">
                                 <div class="col-lg-12 pad-r30">
                         
                                <input id="Submit5" type="submit" value="<%=medarbtyp_txt_016 %>" class="btn btn-sm" /><br />
                         
                                     </div>


                            </div>
                            </form>
                
                        </section>    

                        <br /><br />


            <%else 


                 
        
            
            '************************************************************************************************************************************************
            '******************* Eksport **************************' 

                    call TimeOutVersion()
    
    
                    ekspTxt = ekspTxt 'request("FM_ekspTxt")
	                ekspTxt = replace(ekspTxt, "xx99123sy#z", vbcrlf)
	
	
	
	                filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	                filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
				                Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				                if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\to_2015\medarbtyper.asp" then
					                Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\medarbtyperexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					                Set objNewFile = nothing
					                Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\medarbtyperexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				                else
					                Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\medarbtyperexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					                Set objNewFile = nothing
					                Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\medarbtyperexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				                end if
				
				
				
				                file = "medarbtyperexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
				
				
				                '**** Eksport fil, kolonne overskrifter ***
	                            'strOskrifter = "Medarbejdetype; Medarbejder; Init; Status; Ansatdato; Sidst logget ind; Seneste tidsreg.;Timepris;Timepris2;Timepris3;Timepris4;Timepris5;Kostpris;Normtimer Man.;Normtimer Tir.;Normtimer Ons.;Normtimer Tor.;Normtimer Fre.;Normtimer L�r.;Normtimer S�n.;"
				                strOskrifter = medarbtyp_txt_001&"; "& medarbtyp_txt_109&"; "& medarbtyp_txt_117&"; "& medarbtyp_txt_010&"; "& medarbtyp_txt_011&"; "& medarbtyp_txt_012&"; "& medarbtyp_txt_013&"; "& medarbtyp_txt_018&";"& medarbtyp_txt_018&"2;"& medarbtyp_txt_018&"3;"& medarbtyp_txt_018&"4;"& medarbtyp_txt_018&"5;"& medarbtyp_txt_019&";"& medarbtyp_txt_118&";"& medarbtyp_txt_119&";"& medarbtyp_txt_120&";"& medarbtyp_txt_121&";"& medarbtyp_txt_122&";"& medarbtyp_txt_123&";"& medarbtyp_txt_124&";"
				
				
				
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
                
                             
	                            <a href="../inc/log/data/<%=file%>" target="_blank" ><%=medarbtyp_txt_017 %> >></a>
	                            </td></tr>
	                            </table>
                                </div>
	            
	          
	            
	                            <%
                
                
                                Response.end
	                            'Response.redirect "../inc/log/data/"& file &""	
				



                end if  
                '************************************************************************************************************************************************
              
                
               else%> 
    	            <div><%=medarbtyp_txt_028 %></div>      
                    <%end if %>

	
    
      

<%

case "dbopr", "dbred"
	%>
	<!--#include file="../timereg/inc/isint_func.asp"-->
	<%


    useleftdiv = "to_2015"

	'*** Her inds�ttes en ny type i db ****
		if len(request("FM_navn")) = 0 then
		
		errortype = 8
		call showError(errortype)
		
		else
		
				function SQLBlessDOT(s)
				dim tmpdot
				tmpdot = s
				tmpdot = replace(tmpdot, ".", ",")
				SQLBlessDOT = tmpdot
				end function
				
				
				call erDetInt(SQLBlessDOT(request("FM_Timepris")))
				if isInt > 0 then
				
				errortype = 39
				call showError(errortype)
				isInt = 0
						
				else
					
				call erDetInt(SQLBlessDOT(request("FM_Kostpris")))
				if isInt > 0 then
				
				errortype = 39
				call showError(errortype)
				isInt = 0
						
				else
					
				
					if len(request("FM_Kostpris")) = 0 OR len(request("FM_timepris")) = 0 then
					
					errortype = 39
					call showError(errortype)
					
					else
						
						isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_son")))
						int1 = isInt 
						
						isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_man")))
						int2 = isInt
						
						isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_tir")))
						int3 = isInt
						
						isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_ons")))
						int4 = isInt
						
						isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_tor")))
						int5 = isInt
						
						isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_fre")))
						int6 = isInt
						
						isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_lor")))
						int7 = isInt
						
						isInt = 0
						call erDetInt(request("FM_Kostpris")) 
						int8 = isInt
						
						isInt = 0
						call erDetInt(request("FM_timepris"))
						int9 = isInt
						
						'isInt = 0
						'call erDetInt(request("FM_timepris1"))
						'int10 = isInt
						
						isInt = 0
						call erDetInt(request("FM_timepris2"))
						int11 = isInt
						
						isInt = 0
						call erDetInt(request("FM_timepris3"))
						int12 = isInt
						
						isInt = 0
						call erDetInt(request("FM_timepris4"))
						int13 = isInt
						
						isInt = 0
						call erDetInt(request("FM_timepris5"))
						int14 = isInt

                        isInt = 0
						call erDetInt(request("FM_kostpristarif_A"))
						int15 = isInt

                        sInt = 0
						call erDetInt(request("FM_kostpristarif_B"))
						int16 = isInt

                        sInt = 0
						call erDetInt(request("FM_kostpristarif_C"))
						int17 = isInt

                        sInt = 0
						call erDetInt(request("FM_kostpristarif_D"))
						int18 = isInt

                        sInt = 0
						call erDetInt(request("FM_norm_month"))
						int19 = isInt
                        

                        isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_son_lige")))
						int20 = isInt 

                        isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_man_lige")))
						int21 = isInt

                        isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_tir_lige")))
						int22 = isInt

                        isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_ons_lige")))
						int23 = isInt

                        isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_tor_lige")))
						int24 = isInt

                        isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_fre_lige")))
						int25 = isInt

                        isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_lor_lige")))
						int26 = isInt


                        isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_son_ulige2")))
						int27 = isInt 

                        isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_man_ulige2")))
						int28 = isInt

                        isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_tir_ulige2")))
						int29 = isInt

                        isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_ons_ulige2")))
						int30 = isInt

                        isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_tor_ulige2")))
						int31 = isInt

                        isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_fre_ulige2")))
						int32 = isInt

                        isInt = 0
						call erDetInt(SQLBlessDOT(request("FM_norm_lor_ulige2")))
						int33 = isInt

                   

                        '** Til�fj FM_norm_man_modetid_ulige2

						
						if int1 > 0 OR int2 > 0 OR int3 > 0 OR int4 > 0 OR int5 > 0 OR int6 > 0 OR int7 > 0 OR int8 > 0 OR int9 > 0 OR int11 > 0 OR int12 > 0 OR int13 > 0 OR int14 > 0 _
                        OR int15 > 0 OR int16 > 0 OR int17 > 0 OR int18 > 0 _
                        OR int19 > 0 OR int20 > 0 OR int21 > 0 OR int22 > 0 OR int23 > 0 OR int24 > 0 OR int25 > 0 OR int26 > 0 OR int27 > 0 OR int28 > 0 OR int29 > 0 OR int30 OR int31 > 0 OR int32 > 0 OR int33 > 0 then
						
						errortype = 39
						call showError(errortype)
						isInt = 0
					
						else
							
							
							
							function SQLBless(s)
							dim tmp
							tmp = s
							tmp = replace(tmp, "'", "''")
							SQLBless = tmp
							end function
							
							function SQLBlessDOT2(s)
							dim tmpdot
							tmpdot = s
							tmpdot = replace(tmpdot, ",", ".")
							SQLBlessDOT2 = tmpdot
							end function
							
							strNavn = SQLBless(request("FM_navn"))
							strTimepris = SQLBlessDOT2(request("FM_Timepris"))
							strTimepris1 = strTimepris 'SQLBlessDOT2(request("FM_timepris1"))
							strTimepris2 = SQLBlessDOT2(request("FM_timepris2"))
							strTimepris3 = SQLBlessDOT2(request("FM_timepris3"))
							strTimepris4 = SQLBlessDOT2(request("FM_timepris4"))
							strTimepris5 = SQLBlessDOT2(request("FM_timepris5"))
							
							
							if len(strTimepris) <> 0 then
							strTimepris = strTimepris
							else
							strTimepris = 0
							end if
							
							if len(strTimepris1) <> 0 then
							strTimepris1 = strTimepris1
							else
							strTimepris1 = 0
							end if
							
							if len(strTimepris2) <> 0 then
							strTimepris2 = strTimepris2
							else
							strTimepris2 = 0
							end if
							
							if len(strTimepris3) <> 0 then
							strTimepris3 = strTimepris3
							else
							strTimepris3 = 0
							end if
							
							if len(strTimepris4) <> 0 then
							strTimepris4 = strTimepris4
							else
							strTimepris4 = 0
							end if
							
							if len(strTimepris5) <> 0 then
							strTimepris5 = strTimepris5
							else
							strTimepris5 = 0
							end if
							
							
							dubKostpris = replace(request("FM_kostpris"), ".", "")
                            dubKostpris = replace(dubKostpris, ",", ".")


                            kostpristarif_A = replace(request("FM_kostpristarif_A"), ".", "")
                            kostpristarif_A = replace(kostpristarif_A, ",", ".")

                            kostpristarif_B = replace(request("FM_kostpristarif_B"), ".", "")
                            kostpristarif_B = replace(kostpristarif_B, ",", ".")

                            kostpristarif_C = replace(request("FM_kostpristarif_C"), ".", "")
                            kostpristarif_C = replace(kostpristarif_C, ",", ".")

                            kostpristarif_D = replace(request("FM_kostpristarif_D"), ".", "")
                            kostpristarif_D = replace(kostpristarif_D, ",", ".")

                            
							normtimer_son = SQLBlessDOT2(request("FM_norm_son"))
							normtimer_man = SQLBlessDOT2(request("FM_norm_man"))
							normtimer_tir =	SQLBlessDOT2(request("FM_norm_tir"))
							normtimer_ons = SQLBlessDOT2(request("FM_norm_ons"))
							normtimer_tor = SQLBlessDOT2(request("FM_norm_tor"))
							normtimer_fre = SQLBlessDOT2(request("FM_norm_fre"))
							normtimer_lor = SQLBlessDOT2(request("FM_norm_lor"))

                            normtimer_son_modetid = request("FM_norm_son_modetid")
							normtimer_man_modetid = request("FM_norm_man_modetid")
							normtimer_tir_modetid = request("FM_norm_tir_modetid")
							normtimer_ons_modetid = request("FM_norm_ons_modetid")
							normtimer_tor_modetid = request("FM_norm_tor_modetid")
							normtimer_fre_modetid = request("FM_norm_fre_modetid")
							normtimer_lor_modetid = request("FM_norm_lor_modetid")

                            normtimer_rul = request("FM_normtimer_rul")

                           

                           '*** lige uger
                            normtimer_son_lige = SQLBlessDOT2(request("FM_norm_son_lige"))
							normtimer_man_lige = SQLBlessDOT2(request("FM_norm_man_lige"))
							normtimer_tir_lige = SQLBlessDOT2(request("FM_norm_tir_lige"))
							normtimer_ons_lige = SQLBlessDOT2(request("FM_norm_ons_lige"))
							normtimer_tor_lige = SQLBlessDOT2(request("FM_norm_tor_lige"))
							normtimer_fre_lige = SQLBlessDOT2(request("FM_norm_fre_lige"))
							normtimer_lor_lige = SQLBlessDOT2(request("FM_norm_lor_lige"))

                            if len(trim(normtimer_son_lige)) = 0 then
                            normtimer_son_lige = 0
                            end if

                            if len(trim(normtimer_man_lige)) = 0 then
                            normtimer_man_lige = 0
                            end if

                            if len(trim(normtimer_tir_lige)) = 0 then
                            normtimer_tir_lige = 0
                            end if

                            if len(trim(normtimer_ons_lige)) = 0 then
                            normtimer_ons_lige = 0
                            end if

                            if len(trim(normtimer_tor_lige)) = 0 then
                            normtimer_tor_lige = 0
                            end if

                            if len(trim(normtimer_fre_lige)) = 0 then
                            normtimer_fre_lige = 0
                            end if

                            if len(trim(normtimer_lor_lige)) = 0 then
                            normtimer_lor_lige = 0
                            end if


                            normtimer_son_modetid_lige = request("FM_norm_son_modetid_lige")
							normtimer_man_modetid_lige = request("FM_norm_man_modetid_lige")
							normtimer_tir_modetid_lige = request("FM_norm_tir_modetid_lige")
							normtimer_ons_modetid_lige = request("FM_norm_ons_modetid_lige")
							normtimer_tor_modetid_lige = request("FM_norm_tor_modetid_lige")
							normtimer_fre_modetid_lige = request("FM_norm_fre_modetid_lige")
							normtimer_lor_modetid_lige = request("FM_norm_lor_modetid_lige")


                           '** Rul uge 3
                            normtimer_son_ulige2 = SQLBlessDOT2(request("FM_norm_son_ulige2"))
							normtimer_man_ulige2 = SQLBlessDOT2(request("FM_norm_man_ulige2"))
							normtimer_tir_ulige2 =	SQLBlessDOT2(request("FM_norm_tir_ulige2"))
							normtimer_ons_ulige2 = SQLBlessDOT2(request("FM_norm_ons_ulige2"))
							normtimer_tor_ulige2 = SQLBlessDOT2(request("FM_norm_tor_ulige2"))
							normtimer_fre_ulige2 = SQLBlessDOT2(request("FM_norm_fre_ulige2"))
							normtimer_lor_ulige2 = SQLBlessDOT2(request("FM_norm_lor_ulige2"))

                            normtimer_son_modetid_ulige2 = request("FM_norm_son_modetid_ulige2")
							normtimer_man_modetid_ulige2 = request("FM_norm_man_modetid_ulige2")
							normtimer_tir_modetid_ulige2 = request("FM_norm_tir_modetid_ulige2")
							normtimer_ons_modetid_ulige2 = request("FM_norm_ons_modetid_ulige2")
							normtimer_tor_modetid_ulige2 = request("FM_norm_tor_modetid_ulige2")
							normtimer_fre_modetid_ulige2 = request("FM_norm_fre_modetid_ulige2")
							normtimer_lor_modetid_ulige2 = request("FM_norm_lor_modetid_ulige2")

                           
                            

                            

                            maanedsnorm = SQLBlessDOT2(request("FM_norm_month"))
						    if len(maanedsnorm) <> 0 then
							maanedsnorm = maanedsnorm
							else
							maanedsnorm = 0
							end if
        
							
							if len(dubKostpris) <> 0 then
							dubKostpris = dubKostpris
							else
							dubKostpris = 0
							end if

                            if len(kostpristarif_A) <> 0 then
							kostpristarif_A = kostpristarif_A
							else
							kostpristarif_A = 0
							end if
							
                            if len(kostpristarif_B) <> 0 then
							kostpristarif_B = kostpristarif_B
							else
							kostpristarif_B = 0
							end if

                            if len(kostpristarif_C) <> 0 then
							kostpristarif_C = kostpristarif_C
							else
							kostpristarif_C = 0
							end if

                            if len(kostpristarif_D) <> 0 then
							kostpristarif_D = kostpristarif_D
							else
							kostpristarif_D = 0
							end if
							 
							
							if len(normtimer_man) <> 0 then
							normtimer_man = normtimer_man
							else
							normtimer_man = 0
							end if
							
							if len(normtimer_tir) <> 0 then
							normtimer_tir = normtimer_tir
							else
							normtimer_tir = 0
							end if
							
							if len(normtimer_ons) <> 0 then
							normtimer_ons = normtimer_ons
							else
							normtimer_ons = 0
							end if
							
							if len(normtimer_tor) <> 0 then
							normtimer_tor = normtimer_tor
							else
							normtimer_tor = 0
							end if
							
							if len(normtimer_fre) <> 0 then
							normtimer_fre = normtimer_fre
							else
							normtimer_fre = 0
							end if
							
							if len(normtimer_lor) <> 0 then
							normtimer_lor = normtimer_lor
							else
							normtimer_lor = 0
							end if
							
							if len(normtimer_son) <> 0 then
							normtimer_son = normtimer_son
							else
							normtimer_son = 0
							end if
							
							tp0_valuta = request("FM_valuta_0")
							tp1_valuta = tp0_valuta
							tp2_valuta = request("FM_valuta_2")
							tp3_valuta = request("FM_valuta_3")
							tp4_valuta = request("FM_valuta_4")
							tp5_valuta = request("FM_valuta_5")

                            kp1_valuta = request("FM_valuta_6")
							

                            if len(trim(request("feriesats"))) <> 0 then
                            feriesats = request("feriesats")
                            feriesats = replace(feriesats, ",", ".")
                            else
                            feriesats = 0
                            end if

                            
                            if len(trim(request("FM_maxflex"))) <> 0 then
                            maxflex = request("FM_maxflex")
                            maxflex = replace(maxflex, ",", ".")
                            else
                            maxflex = 0
                            end if

                            

                            mtsortorder = request("FM_sortorder")
                            sostergp = request("FM_soster") 

                            if sostergp > 0 then
                                    
                                    findHgrpSQL = "SELECT mgruppe FROM medarbejdertyper WHERE id = " & sostergp
                                    
                                    'Response.write findHgrpSQL
                                    'Response.end 
                                    oRec5.open findHgrpSQL, oConn, 3
                                    if not oRec5.EOF then
                                    
                                    mgruppe = oRec5("mgruppe")        

                                    end if
                                    oRec5.close 
                            else
                                mgruppe = request("FM_gruppe")
                                if cint(mgruppe) = -1 then 'OR cint(mgruppe) = 0
                                mgruppe = 1
                                end if
							end if


                            if len(trim(request("FM_afslutugekri_proc"))) <> 0 then 
                            afslutugekri_proc = request("FM_afslutugekri_proc")
                            else
                            afslutugekri_proc = 0
                            end if

                            afslutugekri_proc = replace(afslutugekri_proc, ",", ".")

                            if len(trim(request("FM_afslutugekri"))) <> 0 then
                            afslutugekri = request("FM_afslutugekri")
                            else
                            afslutugekri = 0
                            end if
							
                            if len(trim(request("FM_mt_mobil_visstopur"))) <> 0 then
                            mt_mobil_visstopur = 1
                            else
                            mt_mobil_visstopur = 0
                            end if


							if request("FM_Kostpris") < 0 OR strTimepris < 0 OR strTimepris1 < 0 OR strTimepris2 < 0 OR strTimepris3 < 0 OR strTimepris4 < 0 OR strTimepris5 < 0 then
							
							errortype = 39
							call showError(errortype)
							
							else
		

                if len(trim(request("FM_noflex"))) <> 0 then
                noflex = 1
                else
                noflex = 0
                end if

                if len(trim(request("FM_mtype_passiv"))) <> 0 then
                mtype_passiv = 1
                else
                mtype_passiv = 0
                end if
                                
				
				strEditor = session("user")
				strDato = session("dato")
				

                                 'Response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;func:" & func
                                 'Response.end

				        if func = "dbopr" then
				        oConn.execute("INSERT INTO medarbejdertyper (type, timepris, editor, dato, kostpris, normtimer_son, normtimer_man, normtimer_tir, normtimer_ons, normtimer_tor, normtimer_fre, normtimer_lor, "_
				        &" timepris_a1, timepris_a2, timepris_a3, timepris_a4, timepris_a5, "_
				        &" tp0_valuta, tp1_valuta, tp2_valuta, tp3_valuta, tp4_valuta, tp5_valuta, sostergp, mtsortorder, mgruppe, afslutugekri, afslutugekri_proc, noflex, "_
                        &" kostpristarif_A, kostpristarif_B, kostpristarif_C, kostpristarif_D, kp1_valuta, mt_mobil_visstopur, feriesats, mtype_passiv, mtype_maxflex, maanedsnorm, normtimer_rul,"_
                        &" normtimer_man_modetid, normtimer_tir_modetid, normtimer_ons_modetid, normtimer_tor_modetid, normtimer_fre_modetid, normtimer_lor_modetid, normtimer_son_modetid,"_
                        &" normtimer_man_lige, normtimer_tir_lige, normtimer_ons_lige, normtimer_tor_lige, normtimer_fre_lige, normtimer_lor_lige, normtimer_son_lige,"_
                        &" normtimer_man_modetid_lige, normtimer_tir_modetid_lige, normtimer_ons_modetid_lige, normtimer_tor_modetid_lige, normtimer_fre_modetid_lige, normtimer_lor_modetid_lige, normtimer_son_modetid_lige,"_
                        &" normtimer_man_ulige2, normtimer_tir_ulige2, normtimer_ons_ulige2, normtimer_tor_ulige2, normtimer_fre_ulige2, normtimer_lor_ulige2, normtimer_son_ulige2,"_
                        &" normtimer_man_modetid_ulige2, normtimer_tir_modetid_ulige2, normtimer_ons_modetid_ulige2, normtimer_tor_modetid_ulige2, normtimer_fre_modetid_ulige2, normtimer_lor_modetid_ulige2, normtimer_son_modetid_ulige2"_
                        &") VALUES"_
				        &" ('"& strNavn &"', "& strTimepris &", '"& strEditor &"', '"& strDato &"', "& dubKostpris &", "_
				        &" "& normtimer_son &", "& normtimer_man &", "& normtimer_tir &", "& normtimer_ons &", "_
				        &" "& normtimer_tor &", "& normtimer_fre &", "& normtimer_lor &", "& strTimepris1 &", "_
				        &" "& strTimepris2 &", "& strTimepris3 &", "& strTimepris4 &", "& strTimepris5 &", "_
				        &" "& tp0_valuta &","& tp1_valuta &","& tp2_valuta &","& tp3_valuta &","& tp4_valuta &","& tp5_valuta &", "& sostergp &", "_
                        &" "& mtsortorder &", "& mgruppe &","& afslutugekri &","& afslutugekri_proc &", "& noflex &", "_
                        &" "& kostpristarif_A &","& kostpristarif_B &","& kostpristarif_C &","& kostpristarif_D &", "& kp1_valuta &", "& mt_mobil_visstopur &", "& feriesats &", "& mtype_passiv &", "& maxflex &", "& maanedsnorm &", "& normtimer_rul &", "_
                        &" '"& normtimer_man_modetid &"','"& normtimer_tir_modetid &"','"& normtimer_ons_modetid &"','"& normtimer_tor_modetid &"','"& normtimer_fre_modetid &"','"& normtimer_lor_modetid &"','"& normtimer_son_modetid &"', "_
                        &" "& normtimer_man_lige &","& normtimer_tir_lige &","& normtimer_ons_lige &","& normtimer_tor_lige &","& normtimer_fre_lige &","& normtimer_lor_lige &","& normtimer_son_lige &", "_
                        &" '"& normtimer_man_modetid_lige &"','"& normtimer_tir_modetid_lige &"','"& normtimer_ons_modetid_lige &"','"& normtimer_tor_modetid_lige &"','"& normtimer_fre_modetid_lige &"','"& normtimer_lor_modetid_lige &"','"&normtimer_son_modetid_lige &"', "_
                        &" "& normtimer_man_ulige2 &","& normtimer_tir_ulige2 &","& normtimer_ons_ulige2 &","& normtimer_tor_ulige2 &","& normtimer_fre_ulige2 &","& normtimer_lor_ulige2 &","&normtimer_son_ulige2 &", "_
                        &" '"& normtimer_man_modetid_ulige2 &"','"& normtimer_tir_modetid_ulige2 &"','"& normtimer_ons_modetid_ulige2 &"','"& normtimer_tor_modetid_ulige2 &"','"& normtimer_fre_modetid_ulige2 &"','"&normtimer_lor_modetid_ulige2 &"','"&normtimer_son_modetid_ulige2 &"'"_
                        &" )")
				
                
                        strSQlast = "SELECT id FROM medarbejdertyper WHERE id <> 0 ORDER BY id DESC"
                        oRec6.open strSQlast, oConn, 3
                        if not oRec6.EOF then

                                lastId = oRec6("id")

                        end if
                        oRec6.close


                     '*** Tilf�jer sidst oprettet medarbejder hvis mtype 1:1
                     call medarbtypligmedarb_fn()

                     if cint(medarbtypligmedarb) = 1 then

                                if len(trim(request("mtypeIdforvlgt"))) <> 0 then '** Metype er oprettet p� vej fra medarbejeroprettelse. Sidste medarbejder flyttes over I denne type
                                mtypeIdforvlgt = request("mtypeIdforvlgt")
                                else
                                mtypeIdforvlgt = 0
                                end if

                                if mtypeIdforvlgt <> 0 then

                                lastMid = 0
                                strSQlmlast = "SELECT mid, ansatdato FROM medarbejdere WHERE mid <> 0 ORDER BY mid DESC LIMIT 1"
                                oRec6.open strSQlmlast, oConn, 3 
                                if not oRec6.EOF then

                                lastMid = oRec6("mid") 
                                ansatdato = oRec6("ansatdato")

                                end if
                                oRec6.close

                                strSQlmlastTilfoj = "UPDATE medarbejdere SET medarbejdertype = " & lastId & " WHERE mid = " & lastMid
                                oConn.execute(strSQlmlastTilfoj)


                                '*** Tilf�jer til medarbejdertype_historik. Skal g�res pga 1:1 og da der ellers ikke er norm i Forecast_kapacitet og Simulering
                                '*** 20180709
                                '*** Medarbejder  type historik ****
                                ansatdato = year(ansatdato) & "-" & month(ansatdato) & "-" & day(ansatdato)
						        strSQLmthupd = "INSERT INTO medarbejdertyper_historik (mid, mtype, mtypedato) "_
			                    &" VALUES ("& lastMid &", "& lastId &", '"& ansatdato &"') "
			                    oConn.execute(strSQLmthupd)


                                end if

                     end if
                                
                else

                strSQLupd = "UPDATE medarbejdertyper SET type ='"& strNavn &"', timepris = "& strTimepris &", editor = '" &strEditor &"', dato = '" & strDato &"', kostpris = "& dubKostpris &", "_
				&" normtimer_son = "& normtimer_son &", normtimer_man="& normtimer_man &", normtimer_tir="& normtimer_tir &", "_
				&" normtimer_ons="& normtimer_ons &", normtimer_tor="& normtimer_tor &", normtimer_fre="& normtimer_fre &", normtimer_lor="& normtimer_lor &", "_
				&" timepris_a1 = "& strTimepris1 &", timepris_a2 = "& strTimepris2 &", timepris_a3 = "& strTimepris3 &", "_
				&" timepris_a4 = "& strTimepris4 &", timepris_a5 = "& strTimepris5 &", "_
				&" tp0_valuta = "& tp0_valuta &", tp1_valuta = "& tp1_valuta &", "_
				&" tp2_valuta = "& tp2_valuta &", tp3_valuta = "& tp3_valuta &", "_
				&" tp4_valuta = "& tp4_valuta &", tp5_valuta = "& tp5_valuta &", sostergp = "& sostergp &", mtsortorder = "& mtsortorder &", "_
                &" mgruppe = "& mgruppe &", afslutugekri = "& afslutugekri &", afslutugekri_proc = "& afslutugekri_proc &", noflex = "& noflex &", "_
                &" kostpristarif_A = "& kostpristarif_A &", kostpristarif_B = "& kostpristarif_B &", kostpristarif_C = "& kostpristarif_C &", kostpristarif_D = "& kostpristarif_D &","_
                &" kp1_valuta = "& kp1_valuta &", mt_mobil_visstopur = "& mt_mobil_visstopur &", feriesats = "& feriesats &", mtype_passiv = "& mtype_passiv &", mtype_maxflex = "& maxflex &", maanedsnorm = "& maanedsnorm &", normtimer_rul = "& normtimer_rul &","_
                &" normtimer_man_modetid = '"& normtimer_man_modetid &"', normtimer_tir_modetid = '"& normtimer_tir_modetid&"', normtimer_ons_modetid = '"& normtimer_ons_modetid &"', normtimer_tor_modetid = '"& normtimer_tor_modetid &"', normtimer_fre_modetid = '"& normtimer_fre_modetid &"', normtimer_lor_modetid = '"& normtimer_lor_modetid &"', normtimer_son_modetid = '"& normtimer_son_modetid &"',"_
                &" normtimer_man_lige = "& normtimer_man_lige &", normtimer_tir_lige = "& normtimer_tir_lige &", normtimer_ons_lige = "& normtimer_ons_lige &", normtimer_tor_lige = "& normtimer_tor_lige &", normtimer_fre_lige = "& normtimer_fre_lige &", normtimer_lor_lige = "& normtimer_lor_lige &", normtimer_son_lige = "& normtimer_son_lige &","_
                &" normtimer_man_modetid_lige = '"& normtimer_man_modetid_lige &"', normtimer_tir_modetid_lige = '"& normtimer_tir_modetid_lige &"', normtimer_ons_modetid_lige = '"& normtimer_ons_modetid_lige &"', normtimer_tor_modetid_lige = '"& normtimer_tor_modetid_lige &"', normtimer_fre_modetid_lige = '"& normtimer_fre_modetid_lige &"', normtimer_lor_modetid_lige = '"& normtimer_lor_modetid_lige &"', normtimer_son_modetid_lige = '"& normtimer_son_modetid_lige &"',"_
                &" normtimer_man_ulige2 = "& normtimer_man_ulige2 &", normtimer_tir_ulige2 = "& normtimer_tir_ulige2 &", normtimer_ons_ulige2 = "& normtimer_ons_ulige2 &", normtimer_tor_ulige2 = "& normtimer_tor_ulige2 &", normtimer_fre_ulige2 = "& normtimer_fre_ulige2 &", normtimer_lor_ulige2 = "& normtimer_lor_ulige2 &", normtimer_son_ulige2 = "& normtimer_son_ulige2 &","_
                &" normtimer_man_modetid_ulige2 = '"& normtimer_man_modetid_ulige2 &"', normtimer_tir_modetid_ulige2 = '"& normtimer_tir_modetid_ulige2 &"', normtimer_ons_modetid_ulige2 = '"& normtimer_ons_modetid_ulige2 &"', normtimer_tor_modetid_ulige2 = '"& normtimer_tor_modetid_ulige2 &"', normtimer_fre_modetid_ulige2 = '"& normtimer_fre_modetid_ulige2 &"', normtimer_lor_modetid_ulige2 = '"& normtimer_lor_modetid_ulige2 &"', normtimer_son_modetid_ulige2 = '"& normtimer_son_modetid_ulige2 &"'"_
                &" WHERE id = "&id&""

                'response.write strSQLupd
                'response.Flush

				oConn.execute(strSQLupd)


				end if
                

                '** Fra dato ved opdater timepriser ***'
                fraDato = request("FM_opdatertpfra") 
                if isDate(fraDato) = false then
                fraDato = year(now) &"/"& month(now) & "/"& day(now)
                else
                fraDato = year(fraDato) &"/"& month(fraDato) & "/"& day(fraDato) 
                end if

                if func = "dbopr" then
                id = lastId
                end if


                call valutaKurs(kp1_valuta)
                kpvaluta_kurs = dblKurs



                if request("FM_opdater_timepriser") = "1" OR len(trim(request("FM_opd_intern"))) <> 0 AND mt = 0 then
                %>
                <div style="position:absolute; left:200px; top:0px; width:400px;">
                    <b>Opdaterer timepriser p� medarbejdere af denne type.</b><br />
                    Hver medarbejder kan tage op til 15 sekunder at opdatere. <br /><br />

                    <%if instr(lto, "epi") <> 0 then 'Tjekker ikke mtypehistorik %>
                    <b>Nb: Hvis der er mere end 50 medarbejdere af denne type, kontakt da TimeOut.</b><br /><br />
                    <%end if  %>

                <%
                end if

               
                       

                if instr(lto, "epi") <> 0 then 'Tjekker ikke mtypehistorik
                strSQLm = "SELECT m.mid, m.mnr, m.ansatdato AS mtypedato, init FROM medarbejdere AS m "_
                &" WHERE m.medarbejdertype = "& id & " AND (mansat = 1 OR mansat = 3) ORDER BY mid LIMIT 50" 'Aktive og passive
               
                else
                strSQLm = "SELECT m.mid, m.mnr, mth.mtypedato, mth.id, mth.mtype, init FROM medarbejdere AS m "_
                &" LEFT JOIN medarbejdertyper_historik AS mth ON (mth.mid = m.mid AND mth.mtype = m.medarbejdertype) WHERE m.medarbejdertype = "& id & " AND (mansat = 1 OR mansat = 3) LIMIT 200" 'Aktive og passive
                end if
                

                'Response.write strSQLm & "<br>"
                'Response.flush
                
                mt = 0
                oRec.open strSQLm, oConn, 3
                While Not oRec.EOF 
                        
                       
                                

                            
                                '*****************************************************************************************************************
                                '************** Opdater eksisterende timer og timepristabel p� �bne job + aktiviteter p� denne medarb.type 
                                '*****************************************************************************************************************

                                


                                            if len(trim(request("FM_opdater_timepriser_closedjobs"))) <> 0 then 'also closed jobs

                                                closedjobsSQL = "jobstatus = 1 OR jobstatus = 3 OR (jobstatus = 0 AND lukkedato >= '"& year(now) &"-01-01')"
                                            else
                                                closedjobsSQL = "jobstatus = 1 OR jobstatus = 3"

                                            end if


                                        if request("FM_opdater_timepriser") = "1" then

                                           
                                                'if session("mid") = 1 then
                        
                                                'Response.write "========================================================================================<br><br>"
                                                Response.write "Opdaterer timepriser for medarbejder: ["& oRec("init") &"] - "& mt & "<br>"
                                                response.flush
                                
                                                'end if



                                            '** Finder alle �bne job
                                            '** Finder tilh�rende Aktiviteter KUN DEM DER IKKE ER SAT TIL FAST TP
                                            strSQLjob = "SELECT j.id AS jobid, j.jobnr, a.id AS aktid, jobstatus FROM job AS j "_
                                            &" LEFT JOIN aktiviteter AS a ON (a.job = j.id AND (fakturerbar = 1 OR fakturerbar = 2 OR fakturerbar = 6) AND brug_fasttp = 0) "_
                                            &" WHERE ("& closedjobsSQL &") AND a.id IS NOT NULL"
                                            oRec5.open strSQLjob, oConn, 3
                                            while not oRec5.EOF  

                                     
                                            'if session("mid") = 1 then
                    
                                            'else

                                                '** Finder valgt_tp_alt fra timepriser
                                                call alttimepris(oRec5("aktid"), oRec5("jobid"), oRec("mid"), 1)

                                                if foundone = "n" then
                                                    call alttimepris(0, oRec5("jobid"), oRec("mid"), 1)
                                                end if

                                            'end if
                                           

                                            '** Valgte TP og kurs 
                                            'fasttp_val = tp0_valuta
                                            select case alttp_timeprisAlt
                                            case 1
                                            intTimepris = strTimepris1 
                                            case 2
                                            intTimepris = strTimepris2
                                            case 3
                                            intTimepris = strTimepris3
                                            case 4
                                            intTimepris = strTimepris4
                                            case 5
                                            intTimepris = strTimepris5
                                            case 6
                                            intTimepris = strTimepris
                                            case else
                                            intTimepris = strTimepris
                                            end select

                                        
                                            'intTimepris = strTimepris            

										    intTimepris = replace(intTimepris, ",", ".")
                                            alttp_valutaKurs = replace(alttp_valutaKurs, ",", ".")
                                            
                                            'if session("mid") = 1 then
                                            'intTimepris = 151
                                            'alttp_valutaKurs = 100
                                            'intValuta = 1
                                            'end if

                                                    '**** Opdaterer timerpriser p� timeregistreringer
			                                        strSQLtp = "UPDATE timer SET timepris = "& intTimepris &", valuta = "& intValuta &", kurs = "& alttp_valutaKurs &""_
			                                        &" WHERE tdato >= '"& fraDato &"' AND tmnr = "& oRec("mid") &" AND taktivitetid = "& oRec5("aktid") & " AND tjobnr = '"& oRec5("jobnr") &"'"
					                  
                                                   ' Response.flush
                                                    'Response.write strSQLtp & "<br>"
                                                    'Response.end
					                                oConn.execute(strSQLtp)
                                                    
                                                    'Response.end

                                                    if lto = "oko" then

                                                    strSQLrensTp = "DELETE FROM timepriser WHERE jobid = "& oRec5("jobid") &" AND aktid = "& oRec5("aktid") &" AND medarbid =  "& oRec("mid")
                                                    oConn.execute(strSQLrensTp)

                                                    '**** Opdaterer timerpriser i timepristabellen til fremtidige timeregistreringer (stamdata p� job)
                                                    strSQLinsTp = "INSERT INTO timepriser ( "_
							                        &" jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta) "_
							                        &" VALUES ("& oRec5("jobid") &", "& oRec5("aktid") &", "& oRec("mid") &", "& alttp_timeprisAlt &", "& intTimepris &", "& intValuta &")"

                                                    'Response.Write strSQLinsTp & "<br>"
                                                    'Response.flush

                                                    oConn.execute(strSQLinsTp)

                                                    end if

                                            oRec5.movenext
                                            wend
                                            oRec5.close


                                        end if





                        '*** SLUT OPDATER timepriser for denne medarbejdertype
                        '*****************************************************************************************************************




                    
 

                        '*****************************************************************************************************************
                         '*** Opdater �bne job med internkostpris **'
		                '*****************************************************************************************************************
                         if len(trim(request("FM_opd_intern"))) <> 0 then


                             Response.write "Opdaterer kostpriser for medarbejder: ["& oRec("init") &"] - "& mt & "<br>"
                             response.flush


                            if len(trim(request("FM_opd_intern_prdato"))) <> 0 then 'pr. dato ANd floow closed jobs

                                fromDt = fraDato 'fraDato salgspriser
                                closedjobsSQL = closedjobsSQL

                            else

                                if IsNull(oRec("mtypedato")) = true then
                                fromDt = "2000-1-1"
                                else
                                fromDt = year(oRec("mtypedato")) &"/"& month(oRec("mtypedato")) &"/"& day(oRec("mtypedato"))
                                end if

                                closedjobsSQL = "jobstatus = 1 OR jobstatus = 3"

                            end if

                            brugSimpelKostprisUpdate = 1
                            if cint(brugSimpelKostprisUpdate) = 1 then 'Tjekker ikke adgange via projektgrupper mtypehistorik men opdaterer alle kostpriser p� valgte type 

                                strSQLt = "UPDATE timer SET kostpris = "& dubKostpris &", kpvaluta = "& kp1_valuta &",  kpvaluta_kurs = "& kpvaluta_kurs &" WHERE tmnr = " & oRec("mid") & " AND tdato >= '"& fromDt &"'"
                                oConn.execute(strSQLt)

                            else

                            '*** Ignorer projektgruppe rel, da hvis medarbejderen har v�ret med p� jobbet, 
                            '****m� det v�re via sin projektgruppe
                            '*** Ellers bliver timer ikke opdateret ****
                            strSQLj = "SELECT jobnr, a.id AS aid, a.kostpristarif FROM job AS j "_
                            &"LEFT JOIN aktiviteter AS a ON (a.job = j.id) WHERE " & closedjobsSQL
                           
                            'response.write "strSQLj: " & strSQLj & "<br>"    
                                
                            oRec2.open strSQLj, oConn, 3
                            While Not oRec2.EOF 


                             '*** Overskrifer kostpris med kostpristarif fra medab.type
                             kostpristarif = oRec2("kostpristarif")
                             select case kostpristarif
                             case "0"
                             dubKostprisUse = dubKostpris
                             case "A"
                             dubKostprisUse = kostpristarif_A
                             case "B"
                             dubKostprisUse = kostpristarif_B
                             case "C"
                             dubKostprisUse = kostpristarif_C
                             case "D"
                             dubKostprisUse = kostpristarif_D
                             case else
                             dubKostprisUse = dubKostpris 
                             end select

                            if isNull(oRec2("aid")) <> true then 'AND dubKostpris <> ""
                        
                            strSQLt = "UPDATE timer SET kostpris = "& dubKostprisUse &", kpvaluta = "& kp1_valuta &",  kpvaluta_kurs = "& kpvaluta_kurs &" WHERE tjobnr = '" & oRec2("jobnr") & "' AND tmnr = " & oRec("mid") & " AND taktivitetid = "& oRec2("aid") &" AND tdato >= '"& fromDt &"'"
                            'response.write "kostprisTarif: " & oRec2("kostpristarif") & "<br>"
                            'Response.Write strSQLt & "<br>"
                            'Response.flush
                            oConn.execute(strSQLt)

                            end if

                            oRec2.movenext
                            wend
                            oRec2.close


                            end if


                        end if





                        '*****************************************************************************************************************
                            '*** Opdaterer stamaktiviteter med de nye priser ****'
                            '*** kun hvis det er tilvalgt p� den specifikekgruppe '****
                        '*****************************************************************************************************************

                            if len(trim(request("FM_opd_stamgrp"))) <> 0 then
                            opd_stamgrp = 1
                            else
                            opd_stamgrp = 0
                            end if


                            if cint(opd_stamgrp) = 1 then

                            
                            if len(trim(request("FM_opd_stamgrp_ids"))) <> 0 then
                            opd_stamgrp_ids = split(request("FM_opd_stamgrp_ids"), ", ") 

                            for a = 0 to UBOUND(opd_stamgrp_ids)

                            'Response.write "<hr>opd_stamgrp_ids(a): "& opd_stamgrp_ids(a) & "<br>"
                            'Response.flush

                            if len(trim(opd_stamgrp_ids(a))) <> 0 then

                            ct = 0
                            '*** henter alle aktiviteter i valgte stam-grupper ***'
						    strSQLtpe = "SELECT tp.id, tp.timeprisalt, tp.6timepris, a.aktfavorit, a.job FROM aktiviteter AS a "_
                            &" LEFT JOIN timepriser AS tp ON (tp.aktid = a.id AND tp.jobid = 0 AND tp.medarbid = "& oRec("mid") &") WHERE a.job = 0 AND aktfavorit = "& trim(opd_stamgrp_ids(a)) & " AND tp.aktid = a.id AND tp.jobid = 0 AND tp.medarbid = "& oRec("mid")
						    
                            'Response.write "strSQLtpe: "& strSQLtpe & "<br>"
    						
						    oRec5.open strSQLtpe, oConn, 3 
						    while not oRec5.EOF 
    						    

                                select case oRec5("timeprisalt")
                                case 1
                                tpThis = strTimepris1
                                valThis = tp1_valuta
                                case 2
                                tpThis = strTimepris2
                                valThis = tp2_valuta
                                case 3
                                tpThis = strTimepris3
                                valThis = tp3_valuta
                                case 4
                                tpThis = strTimepris4
                                valThis = tp4_valuta
                                case 5
                                tpThis = strTimepris5
                                valThis = tp5_valuta
                                case 6
                                tpThis = strTimepris
                                valThis = tp0_valuta
                                case else
                                tpThis = strTimepris
                                valThis = tp0_valuta
                                end select

                                if isNull(oRec5("id")) <> true then

							    strSQLinsTp = "UPDATE timepriser SET 6timepris = "& tpThis &", 6valuta = "& valThis &" WHERE id = " & oRec5("id")

                                'Response.Write strSQLinsTp  & "<br>" '& oRec5("timeprisalt")
                                'Response.flush

                                oConn.execute(strSQLinsTp)

                                end if
    						    
                                ct = ct + 1
						    oRec5.movenext
						    wend
						    oRec5.close 


                            end if 'len

                            next

                            end if

                            end if'opddater stamgrp

                               

                mt = mt + 1
                oRec.movenext
                wend
                oRec.close


            'if session("mid") = 1 then
            '            
            'Response.write "KKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKK<br><br>"
            'Response.write "KKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKK<br><br>"
            'response.flush
            'Response.end
            'end if


                '*****************************************************************************************************************
                '**** Opdater hovedgruppe p� medarbejderlinier
                '*****************************************************************************************************************

                if len(trim(request("FM_gruppe_opr"))) <> 0 then
                oprHovedgruppe = request("FM_gruppe_opr")
                else 
                oprHovedgruppe = 0
                end if
                


                nyHovedgruppe = mgruppe 'request("FM_gruppe")

                strSQlmed = "UPDATE medarbejdere SET medarbejdertype_grp = "& nyHovedgruppe &" WHERE medarbejdertype_grp = "& oprHovedgruppe
                
                'Response.write strSQlmed
                                
                oConn.execute(strSQlmed)




              
                '**********************************************************************'
                '*** Opdatrerer m�ls�tninger p� medarbejdertype **********************
                '**********************************************************************

                if func = "dbopr" then
                id = lastId
                end if

               

                fomr_mal_arr = split(request("FM_fomr_mal"), ", ##,")
                fomr_mal_arr_id = split(request("FM_fomr_mal_fomrid"), ", ")   
                
                for f = 0 TO UBOUND (fomr_mal_arr_id)

                    if f = 0 then
                    strSQLDeleteFomrMal = "DELETE FROM mtype_mal_fordel_fomr WHERE mmff_mtype = "& id
                    oConn.execute(strSQLDeleteFomrMal)
                    end if
                    
                    fomr_mal_arr(f) = replace(fomr_mal_arr(f), "#", "")
                    if len(trim(fomr_mal_arr(f))) <> 0 AND fomr_mal_arr(f) > 0 then
                    
                    fomr_mal_arr(f) = replace(fomr_mal_arr(f), ",", ".") 
                    strSQLInsertFomrMal = "INSERT INTO mtype_mal_fordel_fomr SET mmff_mal = "& fomr_mal_arr(f) & ", mmff_mtype = "& id &", mmff_fomr = "& fomr_mal_arr_id(f)
                    oConn.execute(strSQLInsertFomrMal)

                    end if

                next
                
                                
                '** Ved tjek SQL            
                'Response.end
				
                if request("FM_opdater_timepriser") = "1" OR len(trim(request("FM_opd_intern"))) <> 0 then
                Response.write "<br><br>Timepriser er opdateret. <a href=""medarbtyper.asp?menu=medarb""> Forts�t >> </a></div>"
                Response.end
                else
                Response.redirect "medarbtyper.asp?menu=medarb"
                end if

				end if 'validering
				end if 'validering
				end if 'validering
				end if 'validering
			end if 'validering
		end if 'validering




	
	case "opret", "red"
	'*** Her indl�ses form til rediger/oprettelse af ny type ***


     %><script src="js/medarbtyper.js" type="text/javascript"></script><%

    call GetSytemSetup()

    if useTrainerlog = 1 then
        %>
        <style>
            .removefromtrailerlog {
                display:none;
            }
        </style>
        
        <% 
    end if
	

    '***** Medarbejdertype 1:1
    if len(trim(request("mtypenavnforvlgt"))) <> 0 then 'mtype = 1:1 med medarbejder
    strNavnFrom = request("mtypenavnforvlgt")
    else
	strNavnFrom = ""
    end if

    if len(trim(request("mtypeIdforvlgt"))) <> 0 then 'mtype = 1:1 med medarbejder
    strFromId = request("mtypeIdforvlgt")
    else
	strFromId = 0
    end if


	if func = "opret" AND cint(strFromId) = 0 then


	strTimepris = ""
	varSubVal = medarbtyp_txt_110 
	varbroedkrumme = medarbtyp_txt_091
	dbfunc = "dbopr"
	
	strTimepris = 0
	dubKostpris = 0
	strTimepris1 = 0
	strTimepris2 = 0
	strTimepris3 = 0
	strTimepris4 = 0
	strTimepris5 = 0
	
	normtimer_son = 0
	normtimer_man = "7,5"
	normtimer_tir = "7,5"
	normtimer_ons = "7,5"
	normtimer_tor = "7,5"
	normtimer_fre = "7"
	normtimer_lor = 0
	
    sostergp = 0
    mtsortorder = 1000
    noflex = 0

                         
    
    feriesats = 0
    maxflex = 0


    kostpristarif_A = 0
    kostpristarif_B = 0
    kostpristarif_C = 0
    kostpristarif_D = 0

    maanedsnorm = 0



           normtimer_rul = 0

            normtimer_man_modetid = "00:00"
            normtimer_tir_modetid = "00:00"  
            normtimer_ons_modetid = "00:00"
            normtimer_tor_modetid = "00:00"
            normtimer_fre_modetid = "00:00" 
            normtimer_lor_modetid = "00:00"
            normtimer_son_modetid = "00:00"
                     
            normtimer_man_lige = 0
            normtimer_tir_lige = 0
            normtimer_ons_lige = 0
            normtimer_tor_lige = 0
            normtimer_fre_lige = 0
            normtimer_lor_lige = 0
            normtimer_son_lige = 0

            normtimer_man_modetid_lige = "00:00"
            normtimer_tir_modetid_lige = "00:00" 
            normtimer_ons_modetid_lige = "00:00"
            normtimer_tor_modetid_lige = "00:00" 
            normtimer_fre_modetid_lige = "00:00"  
            normtimer_lor_modetid_lige = "00:00" 
            normtimer_son_modetid_lige = "00:00"
            
            normtimer_man_ulige2 = 0
            normtimer_tir_ulige2 = 0
            normtimer_ons_ulige2 = 0
            normtimer_tor_ulige2 = 0
            normtimer_fre_ulige2 = 0
            normtimer_lor_ulige2 = 0
            normtimer_son_ulige2 = 0

            normtimer_man_modetid_ulige2 = "00:00"
            normtimer_tir_modetid_ulige2 = "00:00"
            normtimer_ons_modetid_ulige2 = "00:00"
            normtimer_tor_modetid_ulige2 = "00:00"
            normtimer_fre_modetid_ulige2 = "00:00" 
            normtimer_lor_modetid_ulige2 = "00:00"
            normtimer_son_modetid_ulige2 = "00:00"



	else

    '**** Medarbtype 1:1
    if cint(strFromId) <> 0 then
    id = strFromId
    end if

	strSQL = "SELECT type, editor, dato, timepris, kostpris, normtimer_son, normtimer_man, "_
	&" normtimer_tir, normtimer_ons, normtimer_tor, normtimer_fre, normtimer_lor, timepris_a1, "_
	&" timepris_a2, timepris_a3, timepris_a4, timepris_a5, "_
	&" tp0_valuta, tp1_valuta, tp2_valuta, tp3_valuta, tp4_valuta, tp5_valuta, sostergp, mtsortorder, mgruppe, afslutugekri, afslutugekri_proc, noflex, "_
    &" kostpristarif_A, kostpristarif_B, kostpristarif_C, kostpristarif_D, kp1_valuta, mt_mobil_visstopur, feriesats, mtype_passiv, mtype_maxflex AS maxflex, maanedsnorm, normtimer_rul,"_
    &" normtimer_man_modetid, normtimer_tir_modetid, normtimer_ons_modetid, normtimer_tor_modetid, normtimer_fre_modetid, normtimer_lor_modetid, normtimer_son_modetid,"_
    &" normtimer_man_lige, normtimer_tir_lige, normtimer_ons_lige, normtimer_tor_lige, normtimer_fre_lige, normtimer_lor_lige, normtimer_son_lige,"_
    &" normtimer_man_modetid_lige, normtimer_tir_modetid_lige, normtimer_ons_modetid_lige, normtimer_tor_modetid_lige, normtimer_fre_modetid_lige, normtimer_lor_modetid_lige, normtimer_son_modetid_lige,"_
    &" normtimer_man_ulige2, normtimer_tir_ulige2, normtimer_ons_ulige2, normtimer_tor_ulige2, normtimer_fre_ulige2, normtimer_lor_ulige2, normtimer_son_ulige2,"_
    &" normtimer_man_modetid_ulige2, normtimer_tir_modetid_ulige2, normtimer_ons_modetid_ulige2, normtimer_tor_modetid_ulige2, normtimer_fre_modetid_ulige2, normtimer_lor_modetid_ulige2, normtimer_son_modetid_ulige2"_
    &" FROM medarbejdertyper WHERE id=" & id


  

	oRec.open strSQL, oConn, 3
	
	if not oRec.EOF then
	
    if cint(strFromId) <> 0 then
    strNavn = strNavnFrom
    strNavnFolger = oRec("type")
    else
	strNavn = oRec("type")
    strNavnFolger = ""
    end if

	strTimepris = oRec("timepris")
	strDato = oRec("dato")
	strEditor = oRec("editor")
	dubKostpris = oRec("kostpris")
	normtimer_son = oRec("normtimer_son")
	normtimer_man = oRec("normtimer_man")
	normtimer_tir = oRec("normtimer_tir")
	normtimer_ons = oRec("normtimer_ons")
	normtimer_tor = oRec("normtimer_tor")
	normtimer_fre = oRec("normtimer_fre")
	normtimer_lor  = oRec("normtimer_lor")
	strTimepris1 = oRec("timepris_a1")
	strTimepris2 = oRec("timepris_a2")
	strTimepris3 = oRec("timepris_a3")
	strTimepris4 = oRec("timepris_a4")
	strTimepris5 = oRec("timepris_a5")
	
	tp0_valuta = oRec("tp0_valuta")
	tp1_valuta = oRec("tp1_valuta")
	tp2_valuta = oRec("tp2_valuta")
	tp3_valuta = oRec("tp3_valuta")
	tp4_valuta = oRec("tp4_valuta")
	tp5_valuta = oRec("tp5_valuta")

    sostergp = oRec("sostergp")
    mtsortorder = oRec("mtsortorder")

    mgruppe = oRec("mgruppe")

    afslutugekri_proc = oRec("afslutugekri_proc")
	afslutugekri = oRec("afslutugekri")

    noflex = oRec("noflex")

    kostpristarif_A = oRec("kostpristarif_A")
    kostpristarif_B = oRec("kostpristarif_B")
    kostpristarif_C = oRec("kostpristarif_C")
    kostpristarif_D = oRec("kostpristarif_D")

    kp1_valuta = oRec("kp1_valuta")
    mt_mobil_visstopur = oRec("mt_mobil_visstopur")

    feriesats = oRec("feriesats")
    maxflex = oRec("maxflex")
    mtype_passiv = oRec("mtype_passiv")

    maanedsnorm = oRec("maanedsnorm")




        normtimer_rul = oRec("normtimer_rul")

            normtimer_man_modetid = left(formatdatetime(oRec("normtimer_man_modetid"), 3), 5)
            normtimer_tir_modetid = left(formatdatetime(oRec("normtimer_tir_modetid"), 3), 5)  
            normtimer_ons_modetid = left(formatdatetime(oRec("normtimer_ons_modetid"), 3), 5)  
            normtimer_tor_modetid = left(formatdatetime(oRec("normtimer_tor_modetid"), 3), 5)  
            normtimer_fre_modetid = left(formatdatetime(oRec("normtimer_fre_modetid"), 3), 5) 
            normtimer_lor_modetid = left(formatdatetime(oRec("normtimer_lor_modetid"), 3), 5)  
            normtimer_son_modetid = left(formatdatetime(oRec("normtimer_son_modetid"), 3), 5) 
                     
            normtimer_man_lige = oRec("normtimer_man_lige") 
            normtimer_tir_lige = oRec("normtimer_tir_lige") 
            normtimer_ons_lige = oRec("normtimer_ons_lige") 
            normtimer_tor_lige = oRec("normtimer_tor_lige") 
            normtimer_fre_lige = oRec("normtimer_fre_lige") 
            normtimer_lor_lige = oRec("normtimer_lor_lige") 
            normtimer_son_lige = oRec("normtimer_son_lige")

            normtimer_man_modetid_lige = left(formatdatetime(oRec("normtimer_man_modetid_lige"), 3), 5) 
            normtimer_tir_modetid_lige = left(formatdatetime(oRec("normtimer_tir_modetid_lige"), 3), 5) 
            normtimer_ons_modetid_lige = left(formatdatetime(oRec("normtimer_ons_modetid_lige"), 3), 5)
            normtimer_tor_modetid_lige = left(formatdatetime(oRec("normtimer_tor_modetid_lige"), 3), 5) 
            normtimer_fre_modetid_lige = left(formatdatetime(oRec("normtimer_fre_modetid_lige"), 3), 5)  
            normtimer_lor_modetid_lige = left(formatdatetime(oRec("normtimer_lor_modetid_lige"), 3), 5) 
            normtimer_son_modetid_lige = left(formatdatetime(oRec("normtimer_son_modetid_lige"), 3), 5)
            
            normtimer_man_ulige2 = oRec("normtimer_man_ulige2")
            normtimer_tir_ulige2 = oRec("normtimer_tir_ulige2") 
            normtimer_ons_ulige2 = oRec("normtimer_ons_ulige2")
            normtimer_tor_ulige2 = oRec("normtimer_tor_ulige2") 
            normtimer_fre_ulige2 = oRec("normtimer_fre_ulige2") 
            normtimer_lor_ulige2 = oRec("normtimer_lor_ulige2") 
            normtimer_son_ulige2 = oRec("normtimer_son_ulige2")

            normtimer_man_modetid_ulige2 = left(formatdatetime(oRec("normtimer_man_modetid_ulige2"), 3), 5) 
            normtimer_tir_modetid_ulige2 = left(formatdatetime(oRec("normtimer_tir_modetid_ulige2"), 3), 5)
            normtimer_ons_modetid_ulige2 = left(formatdatetime(oRec("normtimer_ons_modetid_ulige2"), 3), 5)
            normtimer_tor_modetid_ulige2 = left(formatdatetime(oRec("normtimer_tor_modetid_ulige2"), 3), 5)
            normtimer_fre_modetid_ulige2 = left(formatdatetime(oRec("normtimer_fre_modetid_ulige2"), 3), 5) 
            normtimer_lor_modetid_ulige2 = left(formatdatetime(oRec("normtimer_lor_modetid_ulige2"), 3), 5)
            normtimer_son_modetid_ulige2 = left(formatdatetime(oRec("normtimer_son_modetid_ulige2"), 3), 5)





	end if
	oRec.close
	
	
    if cint(noflex) <> 0 then
    noflexCHK = "CHECKED"
    else
    noflexCHK = ""
    end if

    if cint(mtype_passiv) <> 0 then
    mtype_passivCHK = "CHECKED"
    else
    mtype_passivCHK = ""
    end if

                    

    if cint(mt_mobil_visstopur) <> 0 then
    mt_mobil_visstopurCHK = "CHECKED"
    else
    mt_mobil_visstopurCHK = ""
    end if


	if len(trim(dubKostpris)) <> 0 then
	dubKostpris = dubKostpris
	else
	dubKostpris = 0
	end if
	
	
	ugetotal = formatnumber(normtimer_son + normtimer_man + normtimer_tir + normtimer_ons + normtimer_tor + normtimer_fre + normtimer_lor, 2)

    if cint(strFromId) <> 0 then
    varSubVal = medarbtyp_txt_110
	varbroedkrumme = medarbtyp_txt_091
	dbfunc = "dbopr"
    else
	dbfunc = "dbred"
	varbroedkrumme = medarbtyp_txt_125
	varSubVal = medarbtyp_txt_126 
    end if
   

	end if 
	


 
	
    %>

    <!--Rediger sidensinhold-->
            <%if level = 1 then %>
    <div class="container">
    <div class="portlet">
        <h3 class="portlet-title"><u><%=medarbtyp_txt_001 & " " %> <%=varbroedkrumme %></u></h3>

        <form action="medarbtyper.asp?menu=medarb&func=<%=dbfunc%>" method="post">
	        <input type="hidden" name="id" value="<%=id%>">
            <input type="hidden" name="mtypeIdforvlgt" value="<%=strFromId%>">
    
                <!-- Denne fejler og er lang tid om load. Derfor fjernet. 20200120
                <div class="row">
                <div class="col-lg-10">&nbsp</div>
                <div class="col-lg-2 pad-b10">
                <button type="submit" class="btn btn-success btn-sm pull-right"><b><%=medarbtyp_txt_029 %></b></button>
                </div>
                </div>
                    -->

        <div class="portlet-body">
            <div class="well well-white">
            <div class="row">
                <div class="col-lg-1">&nbsp</div>
                <div class="col-lg-2"><%=medarbtyp_txt_008 %>:&nbsp<span style="color:red;">*</span></div>
                <div class="col-lg-5"><input class="form-control input-small" type="text" name="FM_navn" value="<%=strNavn%>">

                    <%if len(trim(strNavnFolger)) <> 0 then %>
                    <%=medarbtyp_txt_030 %><br />
                    <%=medarbtyp_txt_031 %>: <b><%=strNavnFolger %></b><br />
                    <%=medarbtyp_txt_032 %> 
                    <%end if %>
                </div>
                <div class="col-lg-4"><input type="checkbox" name="FM_mtype_passiv" value="<%=mtype_passiv%>" <%=mtype_passivChk %> /> Passiv (vises ikke p� lister)
            </div>

             <div class="row"><div class="col-lg-1 pad-b10">&nbsp;</div></div>

                <div class="row">
                    <div class="col-lg-1">&nbsp</div>
                    <div class="col-lg-2"><%=medarbtyp_txt_033 %>: &nbsp<span style="color:red;">*</span>

                        <%select case lto
                        case "kongeaa", "xintrant - local", "dencker"%>
                        <br />
                        <%select case normtimer_rul 
                            case 1
                            normtimer_rulSEL1 = "SELECTED"
                            normtimer_rulSEL0 = ""
                            case else
                            normtimer_rulSEL1 = ""
                            normtimer_rulSEL0 = "SELECTED"
                            end select
                            %>

                        <select name="FM_normtimer_rul" id="FM_normtimer_rul" class="form-control input-small">
                            <option value="0" <%=normtimer_rulSEL0 %>>Fast ugenorm</option>
                            <option value="1" <%=normtimer_rulSEL1 %>>Skiftende ugenorm</option>
                            
                        </select>
                        <%case else %>
                        <input type="hidden" name="FM_normtimer_rul" id="FM_normtimer_rul" value="0" />
                        <%end select %>
                    </div>
                    <div class="col-lg-1"><%=medarbtyp_txt_034 %>: <br />
                        <input type="text" class="form-control input-small" name="FM_norm_man" value="<%=normtimer_man%>">
                    </div>
                    <div class="col-lg-1"><%=medarbtyp_txt_035 %>: <br />
                        <input type="text" class="form-control input-small" name="FM_norm_tir" value="<%=normtimer_tir%>">
                    </div>
                    <div class="col-lg-1"><%=medarbtyp_txt_036 %>: <br />
                        <input type="text" class="form-control input-small" name="FM_norm_ons" value="<%=normtimer_ons%>">
                    </div>
                    <div class="col-lg-1"><%=medarbtyp_txt_037 %>: <br />
                        <input type="text" class="form-control input-small" name="FM_norm_tor" value="<%=normtimer_tor%>">
                    </div>
                    <div class="col-lg-1"><%=medarbtyp_txt_038 %>: <br />
                        <input type="text" class="form-control input-small" name="FM_norm_fre" value="<%=normtimer_fre%>">
                    </div>
                    <div class="col-lg-1"><%=medarbtyp_txt_039 %>: <br />
                        <input type="text" class="form-control input-small" name="FM_norm_lor" value="<%=normtimer_lor%>">
                    </div>
                    <div class="col-lg-1"><%=medarbtyp_txt_040 %>: <br />
                        <input type="text" class="form-control input-small" name="FM_norm_son" value="<%=normtimer_son%>">
                    </div>
                     <div class="col-lg-1"><br /><b><%=medarbtyp_txt_041 %>:</b>&nbsp;<u><%=ugetotal%></u></div>
                </div>

              
                 <div class="row">
                    <div class="col-lg-1">&nbsp</div>
                    <div class="col-lg-2">M�detider: (08:00)</div>

                    <div class="col-lg-1">
                        <input type="text" class="form-control input-small" name="FM_norm_man_modetid" value="<%=normtimer_man_modetid%>">
                    </div>
                    <div class="col-lg-1">
                        <input type="text" class="form-control input-small" name="FM_norm_tir_modetid" value="<%=normtimer_tir_modetid%>">
                    </div>
                    <div class="col-lg-1">
                        <input type="text" class="form-control input-small" name="FM_norm_ons_modetid" value="<%=normtimer_ons_modetid%>">
                    </div>
                    <div class="col-lg-1">
                        <input type="text" class="form-control input-small" name="FM_norm_tor_modetid" value="<%=normtimer_tor_modetid%>">
                    </div>
                    <div class="col-lg-1">
                        <input type="text" class="form-control input-small" name="FM_norm_fre_modetid" value="<%=normtimer_fre_modetid%>">
                    </div>
                    <div class="col-lg-1">
                        <input type="text" class="form-control input-small" name="FM_norm_lor_modetid" value="<%=normtimer_lor_modetid%>">
                    </div>
                    <div class="col-lg-1">
                        <input type="text" class="form-control input-small" name="FM_norm_son_modetid" value="<%=normtimer_son_modetid%>">
                    </div>
                     <div class="col-lg-1"></div>
                </div>

                 <div class="row div_normtimer_rul">
                    <div class="col-lg-1">&nbsp</div>
                     </div>


                 <div class="row div_normtimer_rul">
                    <div class="col-lg-1">&nbsp</div>
                    <div class="col-lg-2">(Lige uger)</div>
                    <div class="col-lg-1"><%=medarbtyp_txt_034 %>: <br />
                        <input type="text" class="form-control input-small" name="FM_norm_man_lige" value="<%=normtimer_man_lige%>">
                    </div>
                    <div class="col-lg-1"><%=medarbtyp_txt_035 %>: <br />
                        <input type="text" class="form-control input-small" name="FM_norm_tir_lige" value="<%=normtimer_tir_lige%>">
                    </div>
                    <div class="col-lg-1"><%=medarbtyp_txt_036 %>: <br />
                        <input type="text" class="form-control input-small" name="FM_norm_ons_lige" value="<%=normtimer_ons_lige%>">
                    </div>
                    <div class="col-lg-1"><%=medarbtyp_txt_037 %>: <br />
                        <input type="text" class="form-control input-small" name="FM_norm_tor_lige" value="<%=normtimer_tor_lige%>">
                    </div>
                    <div class="col-lg-1"><%=medarbtyp_txt_038 %>: <br />
                        <input type="text" class="form-control input-small" name="FM_norm_fre_lige" value="<%=normtimer_fre_lige%>">
                    </div>
                    <div class="col-lg-1"><%=medarbtyp_txt_039 %>: <br />
                        <input type="text" class="form-control input-small" name="FM_norm_lor_lige" value="<%=normtimer_lor_lige%>">
                    </div>
                    <div class="col-lg-1"><%=medarbtyp_txt_040 %>: <br />
                        <input type="text" class="form-control input-small" name="FM_norm_son_lige" value="<%=normtimer_son_lige%>">
                    </div>
                     <div class="col-lg-1"><br /><b><%=medarbtyp_txt_041 %>:</b>&nbsp;<u><%=ugetotal_lige%></u></div>
                </div>


                  <div class="row div_normtimer_rul">
                    <div class="col-lg-1">&nbsp</div>
                    <div class="col-lg-2">M�detider: (08:00)</div>

                    <div class="col-lg-1">
                        <input type="text" class="form-control input-small" name="FM_norm_man_modetid_lige" value="<%=normtimer_man_modetid_lige%>">
                    </div>
                    <div class="col-lg-1">
                        <input type="text" class="form-control input-small" name="FM_norm_tir_modetid_lige" value="<%=normtimer_tir_modetid_lige%>">
                    </div>
                    <div class="col-lg-1">
                        <input type="text" class="form-control input-small" name="FM_norm_ons_modetid_lige" value="<%=normtimer_ons_modetid_lige%>">
                    </div>
                    <div class="col-lg-1">
                        <input type="text" class="form-control input-small" name="FM_norm_tor_modetid_lige" value="<%=normtimer_tor_modetid_lige%>">
                    </div>
                    <div class="col-lg-1">
                        <input type="text" class="form-control input-small" name="FM_norm_fre_modetid_lige" value="<%=normtimer_fre_modetid_lige%>">
                    </div>
                    <div class="col-lg-1">
                        <input type="text" class="form-control input-small" name="FM_norm_lor_modetid_lige" value="<%=normtimer_lor_modetid_lige%>">
                    </div>
                    <div class="col-lg-1">
                        <input type="text" class="form-control input-small" name="FM_norm_son_modetid_lige" value="<%=normtimer_son_modetid_lige%>">
                    </div>
                     <div class="col-lg-1"></div>
                </div>

                 <div class="row div_normtimer_rul">
                    <div class="col-lg-1">&nbsp</div>
                     </div>

                 <div class="row div_normtimer_rul">
                    <div class="col-lg-1">&nbsp</div>
                    <div class="col-lg-2">(Uge 3 rul)</div>
                    <div class="col-lg-1"><%=medarbtyp_txt_034 %>: <br />
                        <input type="text" class="form-control input-small" name="FM_norm_man_ulige2" value="<%=normtimer_man_ulige2%>">
                    </div>
                    <div class="col-lg-1"><%=medarbtyp_txt_035 %>: <br />
                        <input type="text" class="form-control input-small" name="FM_norm_tir_ulige2" value="<%=normtimer_tir_ulige2%>">
                    </div>
                    <div class="col-lg-1"><%=medarbtyp_txt_036 %>: <br />
                        <input type="text" class="form-control input-small" name="FM_norm_ons_ulige2" value="<%=normtimer_ons_ulige2%>">
                    </div>
                    <div class="col-lg-1"><%=medarbtyp_txt_037 %>: <br />
                        <input type="text" class="form-control input-small" name="FM_norm_tor_ulige2" value="<%=normtimer_tor_ulige2%>">
                    </div>
                    <div class="col-lg-1"><%=medarbtyp_txt_038 %>: <br />
                        <input type="text" class="form-control input-small" name="FM_norm_fre_ulige2" value="<%=normtimer_fre_ulige2%>">
                    </div>
                    <div class="col-lg-1"><%=medarbtyp_txt_039 %>: <br />
                        <input type="text" class="form-control input-small" name="FM_norm_lor_ulige2" value="<%=normtimer_lor_ulige2%>">
                    </div>
                    <div class="col-lg-1"><%=medarbtyp_txt_040 %>: <br />
                        <input type="text" class="form-control input-small" name="FM_norm_son_ulige2" value="<%=normtimer_son_ulige2%>">
                    </div>
                     <div class="col-lg-1"><br /><b><%=medarbtyp_txt_041 %>:</b>&nbsp;<u><%=ugetotal_ulige2%></u></div>
                </div>


                 <div class="row div_normtimer_rul">
                    <div class="col-lg-1">&nbsp</div>
                    <div class="col-lg-2">M�detider: (08:00)</div>

                    <div class="col-lg-1">
                        <input type="text" class="form-control input-small" name="FM_norm_man_modetid_ulige2" value="<%=normtimer_man_modetid_ulige2%>">
                    </div>
                    <div class="col-lg-1">
                        <input type="text" class="form-control input-small" name="FM_norm_tir_modetid_ulige2" value="<%=normtimer_tir_modetid_ulige2%>">
                    </div>
                    <div class="col-lg-1">
                        <input type="text" class="form-control input-small" name="FM_norm_ons_modetid_ulige2" value="<%=normtimer_ons_modetid_ulige2%>">
                    </div>
                    <div class="col-lg-1">
                        <input type="text" class="form-control input-small" name="FM_norm_tor_modetid_ulige2" value="<%=normtimer_tor_modetid_ulige2%>">
                    </div>
                    <div class="col-lg-1">
                        <input type="text" class="form-control input-small" name="FM_norm_fre_modetid_ulige2" value="<%=normtimer_fre_modetid_ulige2%>">
                    </div>
                    <div class="col-lg-1">
                        <input type="text" class="form-control input-small" name="FM_norm_lor_modetid_ulige2" value="<%=normtimer_lor_modetid_ulige2%>">
                    </div>
                    <div class="col-lg-1">
                        <input type="text" class="form-control input-small" name="FM_norm_son_modetid_ulige2" value="<%=normtimer_son_modetid_ulige2%>">
                    </div>
                     <div class="col-lg-1"></div>
                </div>

                <div class="row">
                    <div class="col-lg-1">&nbsp</div>
                     </div>
                
               <div class="row">
                   <div class="col-lg-1">&nbsp</div>
                   <div class="col-lg-2">Gns. m�nedsnorm:</div>
                   <div class="col-lg-1 pad-b10"><input type="text" class="form-control input-small" name="FM_norm_month" value="<%=maanedsnorm%>"></div>
                    <div class="col-lg-6">
                        37 timer om ugen x 52 uger = 1.924 timer om �ret.<br />
                        1.924 timer / 12 m�neder = 160,33 om m�neden.</div>



               </div>

            </div>
            </div>

            <div class="row"><div class="pad-b10"></div></div>

            <section class="removefromtrailerlog">
                <!-- Accordion -->
                <div class="panel-group accordion-panel" id="accordion-paneled">
                    <!-- PersonData -->
                    <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseOne"><%=medarbtyp_txt_043 %></a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseOne" class="panel-collapse collapse">

                        <div class="panel-body">
                           
                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2"><b><%=medarbtyp_txt_044 %>:</b><br /> (<%=medarbtyp_txt_045 %> 1)</div>
                                <div class="col-lg-2"><input type="text" class="form-control input-small" name="FM_timepris" value="<%=strTimepris%>"></div>
                                <div class="col-lg-2"> <%call valutaKoder(0, tp0_valuta, 1) %></div>
                            </div>
                            <div class="row"><div class="pad-b10"></div></div>

                            <div class="row">
                                <div class="col-lg-1 pad-t10">&nbsp</div>
                                <div class="col-lg-2 pad-t10"><%=medarbtyp_txt_046 %></div>
                            </div>
                            <!--
                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2">Alt. Timepris 1:</div>
                                <div class="col-lg-2"><input type="text" class="form-control input-small" name="FM_timepris1" value="<%=strTimepris1%>" disabled></div>
		                        <div class="col-lg-2"><%call valutaKoder(1, tp1_valuta, 1) %></div>
                            </div>
                            -->
                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2"><%=medarbtyp_txt_045 %> 2:</div>
                                <div class="col-lg-2"><input type="text" class="form-control input-small" name="FM_timepris2" value="<%=strTimepris2%>"></div>
		                        <div class="col-lg-2"><%call valutaKoder(2, tp2_valuta, 1) %></div>
                            </div>
                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2"><%=medarbtyp_txt_045 %> 3:</div>
                                <div class="col-lg-2"><input type="text" class="form-control input-small" name="FM_timepris3" value="<%=strTimepris3%>"></div>
		                        <div class="col-lg-2"><%call valutaKoder(3, tp3_valuta, 1) %></div>
                            </div>
                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2"><%=medarbtyp_txt_045 %> 4:</div>
                                <div class="col-lg-2"><input type="text" class="form-control input-small" name="FM_timepris4" value="<%=strTimepris4%>"></div>
		                        <div class="col-lg-2"><%call valutaKoder(4, tp4_valuta, 1) %></div>
                            </div>
                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2"><%=medarbtyp_txt_045 %> 5:</div>
                                <div class="col-lg-2"><input type="text" class="form-control input-small" name="FM_timepris5" value="<%=strTimepris5%>"></div>
		                        <div class="col-lg-2"><%call valutaKoder(5, tp5_valuta, 1) %></div>
                            </div>

                            <div class="row pad-t20">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2"><%=medarbtyp_txt_047 %></div>
                                <div class="col-lg-4"><input id="Checkbox2" type="checkbox" name="FM_opd_stamgrp" value="1" />&nbsp <%=medarbtyp_txt_049 %>  <!--opdater timepriserne for alle medarbejdere af denne type p� alle aktiviteterne, i nedenst�ende stamaktivitets-grupper--><br />
                                    <select name="FM_opd_stamgrp_ids" multiple class="form-control input-small">
                                        <%
                                        strSQL = "SELECT navn, id FROM akt_gruppe WHERE id <> 0 AND skabelontype = 0 ORDER BY navn" 
                                        oRec3.open strSQL, oConn, 3
                                        while not oRec3.EOF 
                                        %>
                                        <option value="<%=oRec3("id") %>"><%=oRec3("navn") %></option>
                                        <%
                                        oRec3.movenext
                                        wend 
                                        oRec3.close
                                        %>
                                   </select>
                                </div>
                            </div>

                             <%if func = "red" OR len(trim(strNavnFolger)) <> 0 then 
                                 
                                 useDate = formatdatetime(now,2)%>
                                <div class="row">
                                    <div class="col-lg-1">&nbsp</div>
                                    <div class="col-lg-9"><br /><br />
                                        <%=medarbtyp_txt_048 %>:<br />
                                        <input id="Checkbox1" type="checkbox" name="FM_opdater_timepriser" value="1" />&nbsp <%=medarbtyp_txt_050 %>:<br />
                                        - <%=medarbtyp_txt_051 %>
                                        <br />
                                        - <%=medarbtyp_txt_052 %> <u><%=medarbtyp_txt_053 %></u> <%=medarbtyp_txt_054 %> <b><%=medarbtyp_txt_055 %> </b><input type="text" name="FM_opdatertpfra" value="<%=useDate %>" style="font-size:9px; width:60px;" /> <%=medarbtyp_txt_056 %>
                                        <br /><input id="Checkbox1" type="checkbox" name="FM_opdater_timepriser_closedjobs" value="1" /> Include closed jobs (after 1-1-<%=year(now) %>).
                                        <br />- <%=medarbtyp_txt_057 %>
                                        <br />
                                     (<%=medarbtyp_txt_058 %>)
                                    </div>
                                </div>
                           <%end if %>

                            <div class="row pad-t20">
                                <div class="col-lg-1"><br /><br />&nbsp</div>
                                </div>

                            <div class="row pad-t20">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2"><b><%=medarbtyp_txt_059 %>:</b></div>
                                <div class="col-lg-2"><input type="text" class="form-control input-small" name="FM_kostpris" value="<%=dubKostpris%>"></div>
                                <div class="col-lg-2"><%call valutaKoder(6, kp1_valuta, 1) %></div>
                            </div>

                            <div class="row pad-t20">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-4"><%=medarbtyp_txt_060 %><!--tariffer der kan bruges som till�g til kostpris p� specielle aktiviteter, aften og nat aktiviteter mm.---></div>
                            </div>
                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2"><%=medarbtyp_txt_061 %>:</div>
                                <div class="col-lg-2"><input type="text" class="form-control input-small" value="<%=kostpristarif_A %>" name="FM_kostpristarif_A" /></div>
                                <div class="col-lg-4"><%=basisValISO %></div>
                            </div>
                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2"><%=medarbtyp_txt_062 %>:</div>
                                <div class="col-lg-2"><input type="text" class="form-control input-small" value="<%=kostpristarif_B %>" name="FM_kostpristarif_B" /></div>
                                <div class="col-lg-4"><%=basisValISO %></div>
                            </div>
                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2"><%=medarbtyp_txt_063 %>:</div>
                                <div class="col-lg-2"><input type="text" class="form-control input-small" value="<%=kostpristarif_C %>" name="FM_kostpristarif_C" /></div>
                                <div class="col-lg-4"><%=basisValISO %></div>
                            </div>
                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2"><%=medarbtyp_txt_064 %>:</div>
                                <div class="col-lg-2"><input type="text" class="form-control input-small" value="<%=kostpristarif_D %>" name="FM_kostpristarif_D" /></div>
                                <div class="col-lg-4"><%=basisValISO %></div>
                            </div>
                            <%if func = "red" then %>
                                <div class="row">
                                    <div class="col-lg-3">&nbsp</div>
                                    <div class="col-lg-4"><input id="Checkbox1" type="checkbox" name="FM_opd_intern" value="1" />&nbsp <%=medarbtyp_txt_065 %> <br /><%=medarbtyp_txt_066 %> <u><%=medarbtyp_txt_053 %></u> <%=medarbtyp_txt_067 %>

                                        <br /><input id="Checkbox1" type="checkbox" name="FM_opd_intern_prdato" value="1" />&nbsp Follow "from date" and closed job settings from <%=medarbtyp_txt_048 %> above. (otherwise from the first day each employee was added to this type)
                                      

                                    </div>
                                </div>
                           <%end if %>


                    
                            </div>
                            </div>
                        </div>            
                    </div>
                </section>

               

            <section class="removefromtrailerlog">
                <!-- Accordion -->
                <div class="panel-group accordion-panel" id="accordion-paneled">
                    <!-- PersonData -->
                    <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseTwo"><%=medarbtyp_txt_068 %></a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseTwo" class="panel-collapse collapse">

                        <div class="panel-body">


                             <div class="row pad-t20">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2"><%=medarbtyp_txt_072 %>:</div>
                                <div class="col-lg-3">
                                    <select name="FM_gruppe" id="FM_gruppe" class="form-control input-small">         
                                         <option value="0">Ingen</option>     
                                       <%
                                        mttypSELid = 1
                                        strSQLmt = "SELECT id, navn FROM medarbtyper_grp WHERE id <> 0 ORDER BY navn"
                                        oRec3.open strSQLmt, oConn, 3
                                        while not oRec3.EOF

                                        if cint(mgruppe) = oRec3("id") then
                                            mttypSEL = "SELECTED"
                                            mttypSELid = oRec3("id")
                                        else
                                            mttypSEL = ""
                                        end if 
                                        %>
                                         <option value="<%=oRec3("id") %>" <%=mttypSEL %>><%=oRec3("navn") %></option>
                                        <%
                                        oRec3.movenext
                                        wend
                                        oRec3.close
            
            
            
                                        if cint(sostergp) > 0 then
                                         mttypSEL = "SELECTED"
                                        else
                                         mttypSEL = ""
                                        end if%>

                                        <option value="-1" <%=mttypSEL %> style="background-color:#CCCCCC;">// <%=medarbtyp_txt_073 %> //</option> 

    
                                      </select> 
                                </div>
                            </div>

                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2"><%=medarbtyp_txt_069 %>:</div>
                                <div class="col-lg-3">
                                    <select name="FM_soster" id="FM_soster" class="form-control input-small"> 
                                        <%  if cint(sostergp) = 0 then
                                        mttypSEL = "SELECTED"
                                        else
                                        mttypSEL = ""
                                        end if  %>

                                        <option value="0" <%=mttypSEL  %>><%=medarbtyp_txt_070 %></option>

                                        <%
        
                                        if len(id) <> 0 then
                                            tyid = id
                                        else
                                         tyid = 0
                                        end if
        
                                        strSQLmt = "SELECT id, type FROM medarbejdertyper WHERE id <>  "& tyid & " ORDER BY type"
                                        oRec3.open strSQLmt, oConn, 3
                                        while not oRec3.EOF

                                        if cint(sostergp) = oRec3("id") then
                                            mttypSEL = "SELECTED"
                                        else
                                            mttypSEL = ""
                                        end if 
                                        %>
                                         <option value="<%=oRec3("id") %>" <%=mttypSEL %>><%=oRec3("type") %></option>
                                        <%
                                        oRec3.movenext
                                        wend
                                        oRec3.close  %>

                                        <%  if cint(sostergp) = -1 then
                                            mttypSEL = "SELECTED"
                                        else
                                         mttypSEL = ""
                                        end if  %>

                                        <option value="-1" <%=mttypSEL %>><%=medarbtyp_txt_071 %></option> <!-- -->

                                        </select>
                                </div>
                                <div class="col-lg-3">In reports. If sister group is selected, this group will be selected to.</div>
                            </div>


                           

                             <div class="row"><div class="col-lg-1 pad-b10">&nbsp;</div></div>
                                <div class="row">
                   
                                    <div class="col-lg-1">&nbsp;</div>
                                    <div class="col-lg-2">Feriesats pr. md:</div>
                                    <div class="col-lg-1"><input type="text" name="feriesats" value="<%=feriesats %>" class="form-control input-small" /></div>
                                    <div class="col-lg-4">&nbsp;</div>
                    

                    
                    
                                </div>
                 
                                
               
                                <div class="row">
                                    <div class="col-lg-1">&nbsp;</div>
                                     <div class="col-lg-11"><br /><br /><h4>S�raftaler</h4></div>
                                </div>

                                 <div class="row">
                                     <div class="col-lg-1">&nbsp;</div>
                    
                                     <div class="col-lg-2">Max. Flex:</div>
                                     <div class="col-lg-1"><input type="text" name="FM_maxflex" value="<%=maxflex %>" class="form-control input-small"/> </div>
                     
                                 </div>

                                 <div class="row">
                                     <div class="col-lg-1">&nbsp;</div>
                                     <div class="col-lg-4"><input type="checkbox" name="FM_noflex" value="<%=noflex %>" <%=noflexChk %> /> <%=medarbtyp_txt_042 %></div>
                                     <div class="col-lg-1">&nbsp;</div>
                 
                                 </div>

                            <%select case lto
                            case "intrant - local", "foa"
                            %>
                             <div class="row">
                                     <div class="col-lg-1">&nbsp;</div>
                                     <div class="col-lg-4"><input type="checkbox" name="" value="" <%=noflexChk %> /> Senior Ordning</div>
                                     <div class="col-lg-1">&nbsp;</div>
                 
                                 </div>
                             <div class="row">
                                     <div class="col-lg-1">&nbsp;</div>
                                     <div class="col-lg-4"><input type="checkbox" name="" value="" <%=noflexChk %> /> Omsorgsdgae</div>
                                     <div class="col-lg-1">&nbsp;</div>
                 
                                 </div>
                            <%end select%>



                            <div class="row pad-t20">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2"><%=medarbtyp_txt_074 %>:</div>
                                <div class="col-lg-1"><input type="text" class="form-control input-small" value="<%=mtsortorder%>" name="FM_sortorder" id="FM_sortorder"></div>
                             
                            </div>
                            
                             <div class="row pad-t20">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2"><%=medarbtyp_txt_111 %>:<br />
                                   <%=medarbtyp_txt_112 %>
                                
                                </div>
                                <div class="col-lg-4"><input type="checkbox" value="1" name="FM_mt_mobil_visstopur" <%=mt_mobil_visstopurCHK %> /><br />
                                    <%=medarbtyp_txt_113 %><br />
                                    <%=medarbtyp_txt_114 %><br />
                                    <%=medarbtyp_txt_115 %>
                                </div>
                             
                            </div>


                           
                                  
                                 <%
                            call ersmileyaktiv() 
                            call erStempelurOn()   
            
                            if cint(smilaktiv) = 1 then 
          
                                afslutugekri0 = ""
                                afslutugekri1 = ""
                                afslutugekri2 = ""

                                select case afslutugekri
                                case 1
                                afslutugekri1 = "SELECTED"
                                case 2
                                afslutugekri2 = "SELECTED"
                                case else
                                afslutugekri0 = "SELECTED"
                                end select
            
                              %>

                              <div class="row pad-t20">
                                  <div class="col-lg-1">&nbsp</div>
                                 <div class="col-lg-2">
                           
                                       
                                       <%=medarbtyp_txt_075 %>
		                 
		                  
                                </div>
                                 <div class="col-lg-6"><%=medarbtyp_txt_076 %><br /><%=medarbtyp_txt_077 %>:<br />             
                                <select name="FM_afslutugekri">
                                        <option value="0" <%=afslutugekri0 %>><%=medarbtyp_txt_078 %></option>
                                        <option value="1" <%=afslutugekri1 %>><%=medarbtyp_txt_079 %></option>
                                        <option value="2" <%=afslutugekri2 %>><%=medarbtyp_txt_080 %></option>
                
		                            </select> <%=medarbtyp_txt_081 %> <input type="text" style="width:40px;" value="<%=afslutugekri_proc %>" name="FM_afslutugekri_proc"> % 
                                            <%select case lto
                                               case "dencker"
                                                %>
                                                <%=medarbtyp_txt_082 %>
                                                <%
                                               case else  %>
                                            <%=medarbtyp_txt_083 %> 
                                            <%end select %>
                                    </div>


                                 </div>
                                </div>

                            <%end if %>


                             <div class="row pad-t20"><!-- RoW -->
                                  <div class="col-lg-1">&nbsp</div>
                                 <div class="col-lg-6">
                           
                                       
                                       <h4><%=medarbtyp_txt_084 & " " %> <%=medarbtyp_txt_085 %></h4>
		                 
		                  
                                </div>
                            </div>

                           <div class="row pad-t20"><!-- RoW -->
                                 <div class="col-lg-2">&nbsp</div>
                                 <div class="col-lg-6">
                                     <table><tr>



                                     <% afn = 0
                                        strSQLsel_fomr = "SELECT navn, id FROM fomr WHERE id <> 0" 
                                        oRec3.open strSQLsel_fomr, oConn, 3
                                        while not oRec3.EOF

                                           select case right(afn,1)
                                           case 3,6,9
                                          %>
                                            </tr><tr>
                                          <%
                                            case else
                                         
                                          end select %>
                                          <td style="padding:2px; width:100px;">
                                            <input type="hidden" name="FM_fomr_mal_fomrid" value="<%=oRec3("id")%>" />
                                             <span style="white-space:nowrap;"><%=left(oRec3("navn"), 20) %>:
                                            <%
                                            fomrfundet = 0
                                            strSQLsel_mmff = "SELECT mmff_id, mmff_mal FROM mtype_mal_fordel_fomr WHERE mmff_fomr = "& oRec3("id") 
                                            oRec4.open strSQLsel_mmff, oConn, 3
                                            if not oRec4.EOF then
                                            %>
                                             <input type="text" value="<%=oRec4("mmff_mal") %>" name="FM_fomr_mal" style="width:60px" class="form-control input-small">
                                            <%
                                            ' else
                                           
                                            fomrfundet = 1
                                            end if
                                            oRec4.close


                                            if cint(fomrfundet) = 0 then%>
                                             <input type="text" name="FM_fomr_mal" value="0" style="width:60px" class="form-control input-small">% &nbsp;&nbsp;&nbsp;&nbsp;
                                            <%end if

                                            %>
                                            <input type="hidden" name="FM_fomr_mal" value="##" /></span>
                                            
                                              
                                            </td>
                                            <%

                                            afn = afn + 1
                                        oRec3.movenext
                                        wend
                                        oRec3.close

                                       
                                      
                                   
                                       
                                         
                                    %></tr></table>
                                     <br /><br />&nbsp;
                                     </div>

                                 </div><!-- RoW end -->


                                 </div>
                                </div>
                            </div>
                        </section>
                        </div>

                
             <br /><br /><br />
            <%if func = "red" then %>    
            <div style="font-weight: lighter;"><%=medarbtyp_txt_086 %> <b><%=strDato%></b> <%=medarbtyp_txt_087 %> <b><%=strEditor%></b></div>
            <%end if %>
        
        </div>
             <div style="margin-top:15px; margin-bottom:15px;">
            <button type="submit" class="btn btn-success btn-sm pull-right"><b><%=medarbtyp_txt_088 %></b></button>
             <div class="clearfix"></div>
        </div>

        </form>
      

        <%else%> 
    	<div><%=medarbtyp_txt_089 %></div>      
        <%end if %>

<%case else 
    
    
    
    if len(trim(request("formsubmitted"))) <> 0 then

            if len(trim(request("FM_vispassive"))) <> 0 then
            vispassive = 1
            vispassiveCHK = "CHECKED"
            else
            vispassive = 0
            vispassiveCHK = ""
            end if

            response.Cookies("tsa")("vispassive") = vispassive

     else

            if request.Cookies("tsa")("vispassive") <> "" AND request.Cookies("tsa")("vispassive") <> "0" then
            vispassive = 1
            vispassiveCHK = "CHECKED"
            else
            vispassive = 0
            vispassiveCHK = ""
            end if
        

     end if
     %>



     <%if level = 1 then %>
    <script src="js/medarbtyper_liste.js" type="text/javascript"></script>
    <input type="hidden" id="soogtekst" value="<%=medarb_txt_134 %>" />                            

    <div class="container">
    <div class="portlet">
        <h3 class="portlet-title"><u><%=medarbtyp_txt_090 %></u></h3>
        
       <!-- <form action="medarbtyper.asp?func=opret" method="post">-->
            <input type="hidden" name="lto" id="lto" value="<%=lto%>"><!-- Fejler p� PLAN -->
                <section>
                         <div class="row">
                             <div class="col-lg-10">&nbsp;</div>
                             <div class="col-lg-2">
                              <!--   <input type="submit" value="<%=medarbtyp_txt_091 %> +" class="btn btn-sm btn-success pull-right">-->
                                  <form action="medarbtyper.asp?func=opret" method="post">
                            <button type="submit" class="btn btn-sm btn-success pull-right"><b><%=medarbtyp_txt_091 %> +</b></button><br />&nbsp;
                                      </form>
                               
                                 <!--<a href="medarbtyper.asp?func=opret"><%=medarbtyp_txt_091 %> +</a>-->
                            </div>
                        </div>
                </section>
         <!--</form>-->


          <form action="medarbtyper.asp?formsubmitted=1" method="post">
           <br /><br />
                <section>
                         <div class="row">
                             <div class="col-lg-2"><input type="checkbox" value="1" name="FM_vispassive" <%=vispassiveCHK %> /> Show Passive</div>
                             <div class="col-lg-2">
                            <button type="submit" class="btn btn-sm btn-secondary"><b><%=godkend_txt_005 %> >></b></button>

                            </div>
                        </div>
                </section>
         </form>

    </div>

    <div class="portlet-body">
        <table id="example" class="table dataTable table-striped table-bordered table-hover ui-datatable">
            <thead>
                <tr>
                    <th style="width: 2%"><%=medarbtyp_txt_092 %></th>
                    <th style="width: 5%"><%=medarbtyp_txt_093 %></th>
                    <th ><%=medarbtyp_txt_094 %></th>

                    <%if useTrainerlog <> 1 then %>
                    <th style="width: 20%"><%=medarbtyp_txt_095 %></th>
                    
                    <th style="width: 5%"><%=medarbtyp_txt_096 %>?</th>
                    <%end if %>

                    <th style="width: 20%"><%=medarbtyp_txt_097 %> <br /><span style="font-size:10px;">(<%=medarbtyp_txt_098 %>)</span></th>

                    <%if useTrainerlog <> 1 then %>
                    <th style="width: 8%"><%=medarbtyp_txt_099 %></th>
                    <th style="width: 8%"><%=medarbtyp_txt_100 %></th>
                    <%end if %>

                    <th><%=medarbtyp_txt_101 %></th>
                    
		            <th style="max-width:5px;"><%=medarbtyp_txt_102 %></th>
		          
                </tr>
            </thead>
            <tbody>
                <%
	                sort = Request("sort")
	                if sort = "nr" then
	                odrBy = " ORDER BY id"
	                else
	                odrBy = " ORDER BY mgruppe, mtsortorder, type"
	                end if

                    if cint(vispassive) <> 0 then
                    strSQLpassiv = ""
                    else
                    strSQLpassiv = " AND mtype_passiv = 0"
                    end if

	                strSQL = "SELECT mt.id AS id, type, timepris, kostpris, normtimer_son, "_
	                &" normtimer_man, normtimer_tir, normtimer_ons, normtimer_tor, "_
	                &" normtimer_fre, normtimer_lor, tp0_valuta, v.valutakode, mtsortorder, sostergp, mgruppe, mtg.navn AS mtgnavn, mtype_passiv FROM medarbejdertyper mt "_
	                &" LEFT JOIN medarbtyper_grp AS mtg ON (mtg.id = mt.mgruppe)"_
                    &" LEFT JOIN valutaer v ON (v.id = tp0_valuta) "_
	                &" WHERE mt.id <> 0 "& strSQLpassiv &" "& odrBy
	
	
	                oRec.open strSQL, oConn, 3
	
	               
	                'c = 0
	                
	                while not oRec.EOF

                    'x = 0
                    antalx = 0
                    ugetotal = 0

	                ugetotal = formatnumber(oRec("normtimer_son") + oRec("normtimer_man") + oRec("normtimer_tir") + oRec("normtimer_ons") + oRec("normtimer_tor") + oRec("normtimer_fre") + oRec("normtimer_lor"), 2)
	
	                '** Antal medab i type Aktive ***'
	                strSQL2 = "SELECT count(Mid) AS mids FROM medarbejdere WHERE medarbejdertype = "& oRec("id") & " AND mansat <> 2 GROUP BY medarbejdertype"
	                
                   
                    oRec2.open strSQL2, oConn, 3
	                if not oRec2.EOF then

                    if IsNull(oRec2("mids")) <> true then
	                'x = oRec2("mids")
	                antalx = oRec2("mids")
                    else
                    'x = 0
	                antalx = 0
                    end if
                    'oRec2.movenext
	                
                    end if
	                oRec2.close



                    if cdbl(antalx) <> 0 then
	                Antal = antalx
                    else
                    antalx = 0
                    Antal = 0
                    end if


                     '** Antal medab i type Passive ***'
	                strSQL2 = "SELECT count(Mid) AS mids FROM medarbejdere WHERE medarbejdertype = "& oRec("id") & " AND mansat = 3 GROUP BY medarbejdertype"
	                
                    antalxPas = 0
                    oRec2.open strSQL2, oConn, 3
	                if not oRec2.EOF then

                    if IsNull(oRec2("mids")) <> true then
	                'x = oRec2("mids")
	                antalxPas = oRec2("mids")
                    else
                    'x = 0
	                antalxPas = 0
                    end if
                    'oRec2.movenext
	                
                    end if
	                oRec2.close
	
	
	                t = 0
	                '** Antal medarb i medarbtype historik ***'
	                strSQL2 = "SELECT COUNT(mtype) AS antalhistorik FROM medarbejdertyper_historik WHERE mtype = "& oRec("id") & " GROUP BY mtype, mid"
	
	                'Response.write strSQL2
	                'Response.flush
	                oRec2.open strSQL2, oConn, 3
	                if not oRec2.EOF then
	                t = cint(oRec2("antalhistorik"))
	                end if
	                oRec2.close
	
	                if t <> 0 then
	                t = t 
	                else
	                t = 0
	                end if
	
	            %>
                <tr>
                    <td><%=oRec("id") %></td>
                    <td><%=oRec("mtsortorder") %></td>
                    <td><a href="medarbtyper.asp?menu=medarb&func=red&id=<%=oRec("id")%>"><%=oRec("type")%></a>

                        <%if oRec("mtype_passiv") <> 0 then %>
                        (passiv)
                        <%end if %>
                    </td>

                    <%if useTrainerlog <> 1 then %>
                    <td><%=oRec("mtgnavn") %></td>
                   
                    <td><%if oRec("sostergp") <> 0 then %>
                    (id: <%=oRec("sostergp") %>)
                    <%end if %></td> 

                    <%end if %>

                    <td><a href="medarbtyper.asp?menu=medarb&func=med&id=<%=oRec("id")%>"><%=medarbtyp_txt_103 %></a> <b>(<%=antalx%>)</b><br />
                       <span style="font-size:10px"><%=medarbtyp_txt_104 %>: <%=antalxPas %>, <%=medarbtyp_txt_105 %>: <%=t %></span></td>

                    <%if useTrainerlog <> 1 then %>
                    <td><%=oRec("timepris") &" "& oRec("valutakode") %></td>
                    <td><%=oRec("kostpris") &" "& basisValISO%></td>
                    <%end if %>

                    <td><%=formatnumber(ugetotal)%> <%=medarbtyp_txt_116 %> (<%=formatnumber(ugetotal/5, 1)%>)</td>

                    <%if cint(antalx) = 0 AND cint(t) = 0 then%>
		                    <td style="text-align:center; vertical-align:middle;"><a href="medarbtyper.asp?menu=medarb&func=slet&id=<%=oRec("id")%>"><span style="color:darkred;" class="fa fa-times"></span></a></td>
		              
		                <%else%>
                        <td>&nbsp;</td>
		                     <%end if%>
                </tr>
               <%
                   oRec.movenext
	               wend
               %>
            </tbody>
            <tfoot>

                    <th><%=medarbtyp_txt_092 %></th>
                    <th><%=medarbtyp_txt_093 %></th>
                    <th><%=medarbtyp_txt_094 %></th>

                    <%if useTrainerlog <> 1 then %>
                    <th><%=medarbtyp_txt_095 %></th>
                    
                    <th><%=medarbtyp_txt_096 %>?</th>
                    <%end if %>

                    <th><%=medarbtyp_txt_097 %> <br /><span style="font-size:10px;">(<%=medarbtyp_txt_098 %>)</span></th>

                    <%if useTrainerlog <> 1 then %>
                    <th><%=medarbtyp_txt_099 %></th>
                    <th><%=medarbtyp_txt_0100 %></th>
                    <%end if %>

                    <th><%=medarbtyp_txt_101 %></th>
                    
		            <th><%=medarbtyp_txt_102 %></th>
                    

            </tfoot>
        </table>
         </div>

      


        <br /><br />

          <section>
                <div class="row">
                     <div class="col-lg-12">
                        <b><%=medarbtyp_txt_106 %></b>
                        </div>
                    </div>
                    <form action="medarbtyper.asp?media=eksport&func=med&id=0" method="Post" target="_blank">
                  
                    <div class="row">
                     <div class="col-lg-12 pad-r30">
                         
                    <input id="Submit5" type="submit" value="<%=medarbtyp_txt_107 %>" class="btn btn-sm" /><br />
                         
                         </div>


                </div>
                </form>
                
            </section>    
            
        
        <br /><br />


   
        <%else%> 
    	<div><%=medarbtyp_txt_108 %></div>      
        <%end if %>


<%end select %>

      </div>
    </div>
</div>


<!--#include file="../inc/regular/footer_inc.asp"-->