

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->


<%


if len(session("user")) = 0 then

	errortype = 5
	call showError(errortype)
	response.End
    end if
	

    call smileyAfslutSettings()

	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	if len(request("lastid")) <> 0 then
	lastID = request("lastid")
	else
	lastID = 0
	end if 
	
	level = session("rettigheder")
	
	'Response.Write "level" & level
	
	if len(request("medarbSel")) <> 0 then
	medarbSel = request("medarbSel")
	else
	medarbSel = session("mid")
	end if
	
	if level <> 1 then
	showalle = 0
	medarbSQLkri = " m.mid = " & medarbSel
	else
	showalle = 1
	medarbSQLkri = " m.mid <> 0 "
	end if
	
	'if len(request("hidemenu")) <> 0 then
	'hidemenu = request("hidemenu")
	'else
	'hidemenu = 0
	'end if
	
	rdir = request("rdir")
		
	thisfile = "stat_korsel"
	
	'*** Smiley ***''
    call ersmileyaktiv()
	
	media = request("media")
	'printdo = request("print")
	
	
	function ArTilDatoKm()
	%>
	  År -> Dato (<%=formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 2) %>) :
      
       
        
        <% 
        '** Sum år til dato ***'
        strSQLsum = "SELECT SUM(timer) AS kmsum, tfaktim FROM timer "_
        &" WHERE tmnr = "& ltmnr &" AND "_
        &" tdato BETWEEN '"& strAar_slut & "/1/1" &"' AND '"& sqlDatoSlut &"'"_
        &" AND tfaktim = 5 GROUP BY tmnr, tfaktim"
        kmAarDato = 0
        
        'Response.Write strSQLsum
        'response.flush
        
        oRec3.open strSQLsum, oConn, 3
        if not oRec3.EOF then
        
        kmAarDato = oRec3("kmsum")
        
        end if
        oRec3.close
	    
	    Response.Write "<b>"& formatnumber(kmAarDato, 2) & "</b>"
	
	end function 
    
    %>



<style type="text/css">

    .center 
    {
        text-align:center;
    }

</style>



<div class="wprapper">
    <div class="content">

        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u>Kørsel</u></h3>
                <div class="portlet-body">
                
                    <div class="well well-white">

                        <div class="row">
                            <div class="col-lg-1">Fra: <span style="color:red;">*</span></div>
                            <div class="col-lg-4"><input type="text" class="form-control input-small" value="Hovgårdsvej 51, 7100 Vejle" /></div>

                            <div class="col-lg-1">&nbsp</div>
                            <div class="col-lg-1">Dato: <span style="color:red;">*</span></div>
                            <div class="col-lg-4">
                                <div class='input-group date'>
                                      <input type="text" class="form-control input-small" name="FM_datoer" id="jq_dato" value="<%=day(now) &"/"& month(now) &"/"& year(now) %>" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                            </div>
                        </div>
                        
                        <div class="row">
                            <div class="col-lg-1">Via:</div>
                            <div class="col-lg-4"><input type="text" class="form-control input-small" value="Dragørvej 45, 819 Dragør"/></div>

                            <div class="col-lg-1">&nbsp</div>
                            <div class="col-lg-1">Kunde/Job:</div>
                            <div class="col-lg-4"><input type="text" class="form-control input-small" /></div>
                        </div>

                        <div class="row">
                            <div class="col-lg-1">Til: <span style="color:red;">*</span></div>
                            <div class="col-lg-4"><input type="text" class="form-control input-small" value="Køkebnahvnvej 14, 2100, Købehavnvej" /></div>
                            <div class="col-lg-1">&nbsp</div>
                            <div class="col-lg-1">Aktivitet:</div>
                            <div class="col-lg-4"><input type="text" class="form-control input-small" /></div>
                        </div>

                        <div class="row">
                            <div class="col-lg-6">&nbsp</div>
                            <div class="col-lg-1">Formål: <span style="color:red;">*</span></div>
                            <div class="col-lg-4"><input type="text" class="form-control input-small" value="" /></div>
                        </div>

                        <div class="row">
                            <div class="col-lg-2"><a>Tilføj</a></div>
                        </div>
                        
                        <div class="row">
                            <div class="col-lg-11"><button type="submit" class="btn btn-success btn-sm pull-right"><b>Indlæs >></b></button></div>
                        </div>
                    </div>


                    <br /><br /><br /><br /><br />


                    <div class="row">
                        <h5 class="col-lg-6">Kørsels statistik</h5>
                    </div>
                    <form action="stat_korsel.asp?menu=stat&FM_usedatokri=1&hidemenu=<%=hidemenu%>" method="post">
                    <div class="row">
                        <div class="col-lg-1">Medarbejder:</div>
                        <div class="col-lg-4">

                            <%
                            strSQL = "SELECT mnavn, mid, mnr FROM medarbejdere m WHERE m.mid <> 0 AND mansat <> 2 AND mansat <> 3 AND "& medarbSQLkri &" ORDER BY mnavn"
			                'Response.write strSQL
			                'Response.flush
			                oRec.open strSQL, oConn, 3 
                            
                            if media <> "print" then
			                %>
			                <select name="medarbSel" id="medarbSel" class="form-control input-small">
			
			                <%if showalle = 1 then%>
			                <option value="0">Alle</option>
			                <%end if
			
			                else
			                medarbPrVal = "Alle"
			                end if
			
			
			                while not oRec.EOF 
			
			                 if cint(medarbSel) = oRec("mid") then
			                 mSEL = "SELECTED"
			                 medarbPrVal = oRec("mnavn") &"&nbsp;("& oRec("mnr")&")"
			                 else
			                 mSEL = ""
			                 end if 
			 
			 
			                if media <> "print" then%>
			                <option value="<%=oRec("mid")%>" <%=mSEL%>><%=oRec("mnavn")%>&nbsp;(<%=oRec("mnr")%>)</option>
			                <%
			                end if
			
			                oRec.movenext
			                wend
			                oRec.close
			
			
			                if media <> "print" then
			                 %>
			                </select><br><br>
			                <%else %>
			                <%=medarbPrVal%>
			                <%end if %>
                             
                        </div>
                        
                        <div class="col-lg-1"></div>

                        <div class="col-lg-1">Periode:</div>
                        <div class="col-lg-2">
                            <div class='input-group date'>
                                      <input type="text" class="form-control input-small" name="FM_datoer" id="jq_dato" value="<%=day(now) &"/"& month(now) &"/"& year(now) %>" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>      
                        </div>
                        <div class="col-lg-2">
                            <div class='input-group date'>
                                      <input type="text" class="form-control input-small" name="FM_datoer" id="jq_dato" value="<%=day(now) &"/"& month(now) &"/"& year(now) %>" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>      
                        </div>   
                        <div class="col-lg-1"><button type="submit" class="btn btn-secondary btn-sm pull-right"><b>Vis >></b></button></div>                     
                    </div>
                    </form>

                    <br />
                   <!-- <div class="row">
                        <div class="col-lg-11"><button type="submit" class="btn btn-secondary btn-sm pull-right"><b>Vis >></b></button></div>
                    </div> -->

                    <br /><br />

                    <!--<b>Periode: <%=formatdatetime(datoStart, 1)%> til <%=formatdatetime(datoSlut, 1)%>  </b> -->
                   <!-- <table class="table">
                        <thead>
                            <tr>
                                <th style="width:10%">Dato</th>
                                <th style="width:20%">Fra</th>
                                <th style="width:20%">Til</th>
                                <th style="width:10%">Formål</th>
                                <th style="width:10%">Projekt</th>
                                <th style="width:5%">Retur</th>
                                <th style="width:5%">Kørte km</th>
                                <th style="width:5%">km i alt</th>
                                <th style="width:5%">Afregnet</th>
                            </tr>
                        </thead>

                        <tbody>
                            <tr>
                                <td>08-12-2016</td>
                                <td>Vejle</td>
                                <td>Kbh</td>
                                <td>Møde</td>
                                <td>TimeOut Opsætning & Support (527)</td>
                                <td class="center"><input type="checkbox" value="1" /></td>
                                <td><input type="text" class="form-control input-small" value="250" /></td>
                                <td><input type="text" class="form-control input-small" value="500" readonly /></td>
                                <th class="center"><input type="checkbox"/></th>
                            </tr>
                            <tr>
                                <td>12-12-2016</td>
                                <td>A. P. Møllers Allé 11, 2791 Dragø</td>
                                <td>Energivej 230, 5020 Odense C</td>
                                <td>Møde</td>
                                <td>TimeOut Opsætning & Support (527)</td>
                                <td class="center"><input type="checkbox" value="1" /></td>
                                <td><input type="text" class="form-control input-small" value="150"/></td>
                                <td></td>
                                <th></th>
                            </tr>
                            <tr>
                                <td></td>
                                <td>Energivej 230, 5020 Odense C</td>
                                <td>En gade vest på, 3333 ogenby</td>
                                <td>Møde</td>
                                <td>TimeOut Opsætning & Support (527)</td>
                                <td class="center"><input type="checkbox" value="1" /></td>
                                <td><input type="text" class="form-control input-small" value="100"/></td>
                                <td></td>
                                <th></th>
                            </tr>
                            <tr>
                                <td></td>
                                <td>Vestergade 2, 2871, KBH</td>
                                <td>Nørregade 6, 2871, KBH</td>
                                <td>Møde</td>
                                <td>TimeOut Opsætning & Support (527)</td>
                                <td class="center"><input type="checkbox" value="1" /></td>
                                <td><input type="text" class="form-control input-small" value="20" /></td>
                                <td><input type="text" class="form-control input-small" value="270" readonly/></td>
                                <th class="center"><input type="checkbox"/></th>
                            </tr>
                            
                        </tbody>
                    </table> -->

                    <table class="table" style="border:hidden;">

                        <thead>
                            <tr>
                                <th>&nbsp</th>
                                <th style="width:10%">Dato</th>
                                <th style="width:20%">Fra</th>
                                <th style="width:20%">Til</th>
                                <th style="width:10%">Formål</th>
                                <th style="width:10%">Projekt</th>
                                <th style="width:5%">Retur</th>
                                <th style="width:5%">Kørte km</th>
                                <th style="width:5%">km i alt</th>
                                <th style="width:5%">Afregnet</th>
                            </tr>
                        </thead>

                        <tbody>

                          <% 
                            strSQL = "SELECT tid, taktivitetnavn, timer, tfaktim, tjobnavn, tjobnr, tdato, "_
                            &" tknr, tknavn, tmnavn, tmnr, kkundenr, godkendtstatus, timerkom, bopal, "_
                            &" m.mnr, m.init, m.mid, destination FROM timer "_
                            &" LEFT JOIN kunder K on (kid = tknr) "_
                            &" LEFT JOIN medarbejdere m ON (mid = tmnr) WHERE "& medarbSQLkri &" AND "_
                            &" tdato BETWEEN '2016-01-01' AND '2016-12-31' AND tfaktim = 5 ORDER BY tmnr, tdato DESC "
   
                            'Response.Write strSQL
                            'Response.flush
   
                            oRec.open strSQL, oConn, 3
                            while not oRec.EOF 

                            if lastMid <> oRec("mid") then
                            call afsluger(oRec("mid"), sqlDatoStart, sqlDatoSlut)
                            end if
      
    
                          %>

                            <tr>
                                <td><%=oRec("tmnavn") %></td>
                                <td><%=left(weekdayname(weekday(oRec("Tdato"))), 3) %>. <%=day(oRec("Tdato")) &" "&left(monthname(month(oRec("Tdato"))), 3) &". "& right(year(oRec("Tdato")), 2)%></td>
                                <td></td>
                                <td></td>
                                <%
                                    'if len(trim(oRec("timerkom"))) <> 0 then
                                    'timerKom = replace(oRec("timerkom"), "Til:", "<br><b>Til:</b> ")
                                    'timerKom = replace(timerKom, "Fra:", "<b>Fra:</b> ")
                                    'timerKom = replace(timerKom, "Via:", "<br><b>Via:</b> ")
                                    'else
                                    timerKom = oRec("timerkom")
                                    'end if 
                                %>
                                <td><%=timerkom %></td>

                                

                                <td><%=oRec("tjobnavn") %> (<%=oRec("tjobnr") %>)</td>
                                <td></td>
                                <td>
                                <%
                                '** er periode godkendt ***'
		                                tjkDag = oRec("Tdato")
		                                erugeafsluttet = instr(afslUgerMedab(oRec("tmnr")), "#"&datepart("ww", tjkDag,2,2)&"_"& datepart("yyyy", tjkDag) &"#")
                

                                        strMrd_sm = datepart("m", tjkDag, 2, 2)
                                        strAar_sm = datepart("yyyy", tjkDag, 2, 2)
                                        strWeek = datepart("ww", tjkDag, 2, 2)
                                        strAar = datepart("yyyy", tjkDag, 2, 2)

                                        if cint(SmiWeekOrMonth) = 0 then
                                        usePeriod = strWeek
                                        useYear = strAar
                                        else
                                        usePeriod = strMrd_sm
                                        useYear = strAar_sm
                                        end if

                
                                        call erugeAfslutte(useYear, usePeriod, oRec("tmnr"), SmiWeekOrMonth, 0)
		        
		                                'Response.Write "smilaktiv: "& smilaktiv & "<br>"
		                                'Response.Write "SmiWeekOrMonth: "& SmiWeekOrMonth &" ugeNrAfsluttet: "& ugeNrAfsluttet & " tjkDag: "& tjkDag &"<br>"
		                                'Response.Write "autolukvdatodato: "& autolukvdatodato & "<br>"
		                                'Response.Write "tjkDag: "& tjkDag & "<br>"
		                                'Response.Write "autolukvdato: "& autolukvdato & "<br>"
		                                'Response.Write "erugeafsluttet:" & erugeafsluttet & "<br>"
		        
		                                 strSQL2 = "SELECT risiko FROM job WHERE jobnr = '"& oRec("Tjobnr") &"'"
						                        oRec2.open strSQL2, oConn, 3 
						                        if not oRec2.EOF then
						                        risiko = oRec2("risiko")
						                        end if
						                        oRec2.close

                                         call lonKorsel_lukketPer(tjkDag, risiko)
              
		         
                                        'if (cint(erugeafsluttet) <> 0 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                                         if ( ( datepart("ww", ugeNrAfsluttet, 2, 2) = usePeriod AND cint(SmiWeekOrMonth) = 0) OR (datepart("m", ugeNrAfsluttet, 2, 2) = usePeriod AND cint(SmiWeekOrMonth) = 1 ) AND cint(ugegodkendt) = 1 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                                         (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", tjkDag) = year(now) AND DatePart("m", tjkDag) < month(now)) OR _
                                        (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", tjkDag) < year(now) AND DatePart("m", tjkDag) = 12)) OR _
                                        (smilaktiv = 1 AND autolukvdato = 1 AND DatePart("yyyy", tjkDag) < year(now) AND DatePart("m", tjkDag) <> 12) OR _
                                        (smilaktiv = 1 AND autolukvdato = 1 AND (year(now) - DatePart("yyyy", tjkDag) > 1))) OR cint(lonKorsel_lukketIO) = 1 then
                
                                        ugeerAfsl_og_autogk_smil = 1
                                        else
                
                                        ugeerAfsl_og_autogk_smil = 0
                                        end if 
                
                                        if oRec("godkendtstatus") = 1 then
                                        ugeerAfsl_og_autogk_smil = 1
                                        else
                                        ugeerAfsl_og_autogk_smil = ugeerAfsl_og_autogk_smil
                                        end if
                
                
                                        if ugeerAfsl_og_autogk_smil = 0 OR (level = 10) then%>
                                        
                                        <a href="#" onclick="Javascript:window.open('rediger_tastede_dage_2006.asp?id=<%=oRec("tid") %>', '', 'width=650,height=500,resizable=yes,scrollbars=yes')" class=vmenu><%=formatnumber(oRec("timer"), 2) %></a>
	                                    
                                        <%else %>
                                        <%=formatnumber(oRec("timer"), 2) %>
                                        <%end if %>
                                        </td>
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
    </div>
</div>
    

<!--#include file="../inc/regular/footer_inc.asp"-->