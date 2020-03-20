


<%



level = session("rettigheder")



	
	%>
	
	
    
         <%sub hiddenforms%>
	   
         <input id="FM_lastactiveTD" name="FM_lastactiveTD" value="0" type="hidden" />
         <input id="id" name="id" type="hidden" value="0"/>
        <input id="oprFM_usedatokri" name="FM_usedatokri" type="hidden" value="1"/>
        <input id="oprFM_job" name="FM_job" type="hidden" value="<%=jobid %>"/>
        <input id="oprFM_aftale" name="FM_aftale" type="hidden" value="<%=aftid %>"/>
         <!--<input id="jobelaft" name="jobelaft" type="hidden" value="<%=jobelaft %>"/>-->
         <input id="oprFM_kunde" name="FM_kunde" type="hidden" value="<%=kid %>"/>
         
         <input id="oprFM_start_dag" name="FM_start_dag" type="hidden" value="<%=strDag%>"/>
         <input id="oprFM_start_mrd" name="FM_start_mrd" type="hidden" value="<%=strMrd%>"/>
         <input id="oprFM_start_aar" name="FM_start_aar" type="hidden" value="<%=strAar%>"/>
         
          <input id="oprFM_slut_dag" name="FM_slut_dag" type="hidden" value="<%=strDag_slut%>"/>
         <input id="oprFM_slut_mrd" name="FM_slut_mrd" type="hidden" value="<%=strMrd_slut%>"/>
         <input id="oprFM_slut_aar" name="FM_slut_aar" type="hidden" value="<%=strAar_slut%>"/>
         
         <%end sub %>
	
	
	<%
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	

	
	
	
	
    'func = request("func")
    'thisfile = "erp_opr_faktura"
    'print = request("print")
    err = 0
    

  
   
    
    if cdbl(jobid) = 0 AND cdbl(aftid) = 0 then
    err = 88
    end if
    
 
    
    if err <> 0 then
    %>

	<!--#include file="../inc/regular/header_lysblaa_nojava_inc.asp"-->
	<%
	useleftdiv = "h2"
	errortype = err
	call showError(errortype)
    
    Response.End 
    else
	%>
	
	
	
	
	
	

	
 <form action="erp_opr_faktura_fs.asp?formsubmitted=3&visfaktura=1&visjobogaftaler=1&visminihistorik=1" method="POST">


      
     





<!-- HISTORIK -->
     <br />
  <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseTwo">
                            <%=erp_txt_230 %>                  
                            </a>                           
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseTwo" class="panel-collapse collapse">
                          <div class="panel-body">

 <div class="row">
     <div class="col-lg-4">

	<!-- Fakturaer mini historik --->
      
 <a href="#" id="showfakind" style="color:#999999; font-size:9px; font-weight:lighter;">[ <%=erp_txt_498 %> ]</a>
             <br />
         <br /><br />
         <b><%=erp_txt_230 %></b> (last 20)<br />
         <table style="width:100%;" border="0">
     <!--
	<div id=fakturaer style="position:relative; height:260px; top:0px; background-color:#ffffff; visibility:visible; display:; width:275px; overflow:scroll; z-index:1;">
        -->
	
         
       <%
	 
        igDato = 1
	    'if cdbl(igDato) = 1 then
        SQLkriPer = ""
	    'else
	    'SQLkriPer = " AND ((f.fakdato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"') OR (f.labeldato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"' AND brugfakdatolabel = 1)) "
	    'end if
	    
	    
	    lastFak = "2001/01/01"
		fakfindes = 0
		lastFaknr = 0
		
		
		if jobid = 0 then
		jobaftSQL = "f.aftaleid = "& aftid &""
		len_jobtilkaftSQLkri = len(jobtilkaftSQLkri)
		left_jobtilkaftSQLkri = left(jobtilkaftSQLkri, len_jobtilkaftSQLkri - 3)
		jobtilkaftSQLkri = left_jobtilkaftSQLkri & ")"
		
		jobaftSQL = jobaftSQL & jobtilkaftSQLkri
		
		else
		jobaftSQL = "f.jobid = "& jobid &""
		end if


        select case lto
        case "bf"
        '** Medarb. må kun se nationale kontore
        call medariprogrpFn(session("mid"))
        call meStamdata(session("mid"))

                useasfakSQL = " AND afsender = -1"

                if meMed_lincensindehaver = 2 then 'Togo
                useasfakSQL = " AND afsender = 6"
                end if

                if meMed_lincensindehaver = 3 then 'Binin
                useasfakSQL = " AND afsender = 7"
                end if

                if meMed_lincensindehaver = 5 then 'Burkino Faso
                useasfakSQL = " AND afsender = 11"
                end if

                if meMed_lincensindehaver = 4 then 'Mali
                useasfakSQL = " AND afsender = 10"
                end if

                if meMed_lincensindehaver = 0 then 'CPH Head Office 
                useasfakSQL = " AND afsender = 1"
                end if

                if cint(level) = 1 then 'Administrator
                useasfakSQL = " AND afsender <> -1"
                end if

        case else
          useasfakSQL = " AND afsender <> -1"
        end select


		
		strSQLFak = "SELECT f.jobid, f.aftaleid, f.fid, f.fakdato, f.faknr, f.betalt, "_
		&" f.faktype, f.beloeb, f.shadowcopy, f.valuta, f.kurs, v.valutakode, brugfakdatolabel, labeldato, medregnikkeioms, fak_laast "_
		&" FROM fakturaer f LEFT JOIN valutaer v ON (v.id = f.valuta) "_
		&" WHERE ("& jobaftSQL &")"& SQLkriPer &" "& useasfakSQL &""_
		&" ORDER BY f.faknr DESC LIMIT 20"
	        
	        'Response.Write strSQLFak &"<br>" 
	        'Response.Flush
	        
		f = 1
		lastFakfundet = 0
		oRec3.open strSQLFak, oConn, 3
        while not oRec3.EOF 
            
            if f = 1 then 
            fakfindes = 1
            end if
            
            select case oRec3("faktype")
            case 0
                if cdbl(lastFakfundet) <> 1 AND cdbl(oRec3("medregnikkeioms")) <> 1 AND cdbl(oRec3("medregnikkeioms")) <> 2 then
                lastFakfundet = 1
                lastFak = oRec3("fakdato")
                lastFaknr = oRec3("faknr")
                end if
            tyThis = "Fak."
            belthis = oRec3("beloeb")
            case 1
            tyThis = "Kre."
            belthis = -(oRec3("beloeb"))
            'case 2
            'tyThis = "Ryk."
            'belthis = oRec3("beloeb")
            end select
            
            if cdbl(id) = oRec3("fid") then
            bgthis = "#ffff99"
            else
            bgthis = "#ffffff"
            end if
            
            if oRec3("betalt") = 1 then
            acls = "erp_green"
            else
            acls = "erp_silver"
            end if


           %><tr>
               <%

            
            if (oRec3("shadowcopy") <> 1 AND jobid <> 0) OR (aftid <> 0 AND oRec3("aftaleid") <> 0) then 
                  %>
            
            <!--id="td_1_<%=f%> --> 

            <td>
            
            <%if oRec3("brugfakdatolabel") = 1 then%>
            L: <%=replace(formatdatetime(oRec3("labeldato"),2),"-",".") %>
            <%else %>
            F: <%=replace(formatdatetime(oRec3("fakdato"),2),"-",".") %>
            <%end if %>
            &nbsp;<%=tyThis%>


            <%if oRec3("medregnikkeioms") = 1 then %>
            (i)
            <%end if %>

            <%if oRec3("medregnikkeioms") = 2 then %>
            (h)
            <%end if %>
            
            </td>
            <td>
            
            <!-- id="td_3_<%=f%> -->
             <a href="erp_opr_faktura_fs.asp?func=red&id=<%=oRec3("fid")%>&visminihistorik=1&visfaktura=2&visjobogaftaler=1&FM_job=<%=jobid%>&FM_aftale=<%=aftid%>&FM_kunde=<%=kid%>" class=<%=acls %> onclick="activetd(<%=f%>)"><%=oRec3("faknr")%></a><br />
            <!-- id="td_4_<%=f%> --> 
            </td>
            <td>
                   <%
            if oRec3("betalt") <> 1 AND (cdate(oRec3("fakdato")) > cdate("25/8/2007")) then%>
            <a href="erp_opr_faktura_fs.asp?func=red&id=<%=oRec3("fid")%>&visminihistorik=1&visfaktura=1&visjobogaftaler=1&FM_usedatokri=1&FM_job=<%=jobid%>&FM_aftale=<%=aftid%>&FM_kunde=<%=kid%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd %>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>"><img src="../ill/redigerfak.gif" border="0" alt="<%=erp_txt_510 %>" /></a><br />
             
            <!-- &rykkreopr=j -->
            </td>
            <td>
            <%
                if oRec3("fak_laast") <> 1 then ' har faktura været låst?? 
                %><a href="erp_opr_faktura_fs.asp?func=slet&id=<%=oRec3("fid")%>&visminihistorik=1&visfaktura=0&visjobogaftaler=1"><img src="../ill/sletfak.gif" border="0" alt="<%=erp_txt_508 %>" /></a><br />
                <%end if %>

            </td>

                <!-- &FM_usedatokri=1&FM_job=<%=jobid%>&FM_aftale=<%=aftid%>&FM_kunde=<%=kid%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd %>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%> -->

            <%else 
        
                %> <td>

                         <%
                        if level = 1 AND (cdate(oRec3("fakdato")) > cdate("25/8/2007")) then%>
               
                  
                        <!-- Skal IKKE være muligt pga huller i faknr reække følge ELLER HVAD:: ***' --> 
                        <a href="erp_opr_faktura_fs.asp?func=fortryd&id=<%=oRec3("fid")%>&visminihistorik=1&visfaktura=0&visjobogaftaler=1" class=vmenualt>
                            <img src="../ill/fortryd.gif" border="0" alt="Fortryd godkend" /></a>
                 
                
                        <%end if 
                %></td><%
               
                    
                   end if %>
            
               <td>
               <!--id="td_2_<%=f%> --> 
             <%=formatnumber(belthis, 0) &" "& oRec3("valutakode") %>
            </td>
            
            <%
            
            call beregnValuta(belthis,oRec3("kurs"),100)
            totbel = totbel + (valBelobBeregnet)
            
            else 'shadowcopy %>
                
                <!-- id="td_1_<%=f%>-->
                <td>
                 <%if oRec3("brugfakdatolabel") = 1 then%>
            L: <%=replace(formatdatetime(oRec3("labeldato"),2),"-",".") %>
            <%else %>
            F: <%=replace(formatdatetime(oRec3("fakdato"),2),"-",".") %>
            <%end if %>
                
              &nbsp;<%=tyThis%>
              
               <%if oRec3("medregnikkeioms") = 1 then %>
            (i)
            <%end if %>
            </td>
                <td>
                <!-- id="td_3_<%=f%> -->
                    &nbsp;<%=oRec3("faknr")%>
              <!-- id="td_4_<%=f%> -->
                </td>
                <td>
                <% 
                
                '*** Hvis man er inde på et job og skal de de
                '*** fakturaer der ligger på aftaler som dettet job er tilknyttet **
                if jobid <> 0 then
                
                aftnavn = ""
                aftnr = ""
                
                strSQLsc = "SELECT f.aftaleid, s.navn, s.aftalenr FROM fakturaer f "_
                &"LEFT JOIN serviceaft s ON (s.id = f.aftaleid) WHERE f.faknr = "& oRec3("faknr") &" AND f.shadowcopy = 0"
                
                oRec.open strSQLsc, oConn, 3
                if not oRec.EOF then
                
                aftnavn = oRec("navn")
                aftnr = oRec("aftalenr")
                
                end if
                oRec.close%>
                
                    <%if len(aftnavn) > 12 then %>
                    <%=left(aftnavn, 12) & ".. ("& aftnr &")" %>
                    <%else %>
                    <%=aftnavn & " ("& aftnr &")" %>
                    <%end if %>
                    </td>
                
                <%else 
                
                jobnavn = ""
                jobnr = ""
                
                strSQLsc = "SELECT f.jobid, j.jobnavn, j.jobnr FROM fakturaer f "_
                &"LEFT JOIN job j ON (j.id = f.jobid) WHERE f.faknr = "& oRec3("faknr") &" AND f.shadowcopy = 0"
                
                oRec.open strSQLsc, oConn, 3
                if not oRec.EOF then
                
                jobnavn = oRec("jobnavn")
                jobnr = oRec("jobnr")
                
                end if
                oRec.close
                %>
                
                    <%if len(jobnavn) > 12 then %>
                    <%=left(jobnavn, 12) & ".. ("& jobnr &")" %>
                    <%else %>
                    <%=jobnavn& " ("& jobnr &")" %>
                    <%end if %>
               </td>
                    
                
                
            <%end if %>
               
                
                
            <%
            
            totbel = totbel 'shadowcopy
            end if %>

            </tr>
            
            <% 
            
       
        f = f + 1
        oRec3.movenext
        wend 
        oRec3.close
                
        
       
       %><tr><td colspan="5">
            
            <%if f = 1 then %>
            (<%=erp_txt_446 %>)
            <%else %>
            <br /><%=erp_txt_447 %>: <b> <%=formatnumber(totbel, 2) &" "& basisValISO%></b> <br /><%=erp_txt_448 %></td></tr>
            <%end if %>
        
 
         </table>
          <%
                
              
      
            call hiddenforms %>
            
            </div>
        </div><!-- row -->

                   </div> <!-- /.panel-body -->
                        </div> <!-- /.panel-collapse -->
                      </div> <!-- /.panel -->




     	<div class="row">
        <div class="col-lg-2 pad-t20 pad-b20">
	
	
	
	<% 
	sqlDatostart = strAar&"/"&strMrd&"/"&strDag                'year(datointervalstart)&"/"& month(datointervalstart)&"/"&day(datointervalstart) 
	sqlDatoslut = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut  'year(datointervalslut)&"/"& month(datointervalslut) &"/"& day(datointervalslut)
    %>
	
	
	

	
	
        
             <input type="hidden" id="FM_job_tilknyttet_aftale" name="FM_job_tilknyttet_aftale" value=<%=jobtilkaft%> />
    
       
   
		<%
		'** Faktype ***
		select case intType
		case "0"
		selTypNUL = "SELECTED" 
		selTypET = ""
		selTypTO = ""
		strFaktypeNavn = "Faktura"
		case "1"
		selTypET = "SELECTED"
		selTypTO = ""
		selTypNUL = ""
		strFaktypeNavn = "Kreditnota"
		'case "2"
		'selTypTO = "SELECTED"
		'selTypET = ""
		'selTypNUL = ""
		'strFaktypeNavn = "Rykker"
		case else
		selTypNUL = "SELECTED"
		selTypET = ""
		selTypTO = ""
		strFaktypeNavn = "Faktura"
		end select
		%>
		
        <select id="FM_type" name="FM_type" class="form-control input-small">
                <%select case lto
                    case "intranet - local", "bf"
                    %>
                     <option value=0 <%=selTypNUL %>>Invoice</option>
                <option value=1 <%=selTypET %>>Creditnote</option>
                    <%
                        opretTxt = "Create new"


                case "epi2017"
                        
                        opretTxt = "Create new"
                        
                        %>
                    
                <option value=0 <%=selTypNUL %>><%=erp_txt_495 %></option>

                <%if level = 1 then %>
                <option value=1 <%=selTypET %>><%=erp_txt_496 %></option>
                <%end if %>

                <%case else %>    
                <option value=0 <%=selTypNUL %>><%=erp_txt_495 %></option>
                <option value=1 <%=selTypET %>><%=erp_txt_496 %></option>
                <%
                    opretTxt = erp_txt_497
                    
                    end select %>
             
            <!--<option value=2 <%=selTypTO%>>Rykker</option>-->
         </select> 
       
        </div>
        <div class="col-lg-1  pad-t20">
       
   
       
     
       <input id="Submit2" type="submit" value=" <%=opretTxt%> >> " class="btn btn-sm btn-success" />

        </div>
        </div> <!-- row -->




      </form>
  

         

               

              
              
    
	
	
       
        
        
        
	
	
	 <div id="fakind_d" style="position:absolute; width:275px; height:500px; left:220px; top:114px; border:10px #CCCCCC solid; visibility:hidden; display:none; z-index:10000; padding:10px 10px 10px 10px; background-color:#ffffff;">
        <table cellpadding=0 cellspacing=4 border=0 width=100%>
        <form id=prefak action="erp_opr_faktura_fs.asp?formsubmitted=3&visfaktura=1&visjobogaftaler=1&visminihistorik=1&func=opdprefak" method="post">
            <tr><td align="right" style="padding:10px 10px 10px 10px;"><span id="luk_fakpre" style="color:red;">[X]</span></td></tr>
        <tr><td style="padding:10px 10px 10px 10px;">
        <h4><%=erp_txt_449 %>:</h4>
        
        <b>A)</b> <%=erp_txt_450 %>
		<%=erp_txt_451 %>:<br />
		<%
		if request.cookies("erp")("tvmedarb") = "1" then
		chkA = "CHECKED"
		chkB = ""
		else
		chkB = "CHECKED"
		chkA = ""
		end if
		%>
		<input type="radio" name="FM_chkmed" value="1" <%=chkA%>> <%=erp_txt_453 %> 
		<input type="radio" name="FM_chkmed" value="0" <%=chkB%>> <%=erp_txt_454 %>
		
		<br /><br />
		<b>B)</b> <%=erp_txt_452 %> <br />
		<%
		if request.cookies("erp")("tvlogs") = "1" then
		chklogA = "CHECKED"
		chklogB = ""
		else
		chklogB = "CHECKED"
		chklogA = ""
		end if
		%>
		<input type="radio" name="FM_chklog" value="1" <%=chklogA%>> <%=erp_txt_453 %> 
		<input type="radio" name="FM_chklog" value="0" <%=chklogB%>> <%=erp_txt_454 %>
		
		
		<br /><br />
		<b>C)</b> <%=erp_txt_455 %> <br />
		<%
		if request.cookies("erp")("tvtlffax") = "1" then
		chktlffaxA = "CHECKED"
		chktlffaxB = ""
		else
		chktlffaxB = "CHECKED"
		chktlffaxA = ""
		end if
		%>
		<input type="radio" name="FM_chktlffax" value="1" <%=chktlffaxA%>> <%=erp_txt_453 %> 
		<input type="radio" name="FM_chktlffax" value="0" <%=chktlffaxB%>> <%=erp_txt_454 %>
		
		<br /><br />
		<b>D)</b> <%=erp_txt_456 %> <br />
		<%
		if request.cookies("erp")("tvemail") = "1" then
		chkemailA = "CHECKED"
		chkemailB = ""
		else
		chkemailB = "CHECKED"
		chkemailA = ""
		end if
		%>
		<input type="radio" name="FM_chkemail" value="1" <%=chkemailA%>> <%=erp_txt_453 %> 
		<input type="radio" name="FM_chkemail" value="0" <%=chkemailB%>> <%=erp_txt_454 %>
		
		<!--
		<br /><br />
		<b>E)</b> Vis alle medarbejdere uanset om de er de-aktiverede eller tilknyttet via projektgrupper. <br />
		Der kan forekomme lang load tid.<br />
		<%
		'if request.cookies("erp")("deak") = "1" then
		'chkdeakA = "CHECKED"
		'chkdeakB = ""
		'else
		'chkdeakB = "CHECKED"
		'chkdeakA = ""
		'end if
		%>
		<input type="radio" name="FM_chkdeak" value="1" <%=chkdeakA%>> Ja 
		<input type="radio" name="FM_chkdeak" value="0" <%=chkdeakB%>> Nej
		-->
		
		
		
		
		<br /><br />
		<b>E)</b> <%=erp_txt_457 %><br />
		<%
		if request.cookies("erp")("lukjob") = "1" then
		chklukjobA = "CHECKED"
		chklukjobB = ""
		else
		chklukjobB = "CHECKED"
		chklukjobA = ""
		end if
		%>
		<input type="radio" name="FM_chklukjob" value="1" <%=chklukjobA%>> <%=erp_txt_453 %> 
		<input type="radio" name="FM_chklukjob" value="0" <%=chklukjobB%>> <%=erp_txt_454 %>


        <br /><br />
        <%if cdbl(ign_akttype_inst) = 1 then
        ign_akttype_inst_CHK = "CHECKED"
        else
        ign_akttype_inst_CHK = ""
        end if %>
        <b>F)</b>  <input id="Checkbox1" name="ign_akttype_inst" value="1" type="checkbox" <%=ign_akttype_inst_CHK %> /> <%=erp_txt_458 %>
	  

		
		 <%call hiddenforms %>
	    
	    <br />
            &nbsp;
	    </td>
	    
	    </tr>
         

            <tr><td align=right style="padding:10px 10px 10px 10px;">
            <input id="Submit3" type="submit" value="<%=erp_txt_499 %> >>"  /> 
	    </td></tr>
		</form>
		</table>

     

		</div>
	
	
	    
	   
       
        
        
      
     




     
     <!--- Lommeregner --->
	<div id=lommeregner style="position:absolute; background-color:#ffffff; width:275px; height:230px; top:64px; border:0px #5582d2 solid; visibility:hidden; display:none; z-index:1;">
	<table cellspacing=0 cellpadding=0 border=0><form id=beregn name=beregn>
	<tr>
		<td style="padding:10px 2px 0px 10px;" valign=top><%=erp_txt_459 %>: <input type="text" name="beregn_belob" id="beregn_belob" value="0" size="4" style="font-size:9px;"> <b>/</b> </td>
	    <td style="padding:10px 2px 0px 2px;" valign=top><%=erp_txt_460 %>: <input type="text" name="beregn_timer" id="beregn_timer" value="0" size="4" style="font-size:9px;"></td>
	    <td style="padding:10px 2px 0px 2px;" valign=top><input type="button" name="beregn" id="beregn" value=" = " onClick="beregntimepris()" style="font-size:9px;">&nbsp; <input type="text" name="beregn_tp" id="beregntp" value="0" style="width:58px; font-size:9px;"></td>
	</tr></form></table>								
	</div>
	<!-- ASLUT lommeregner -->
	    
    
    
      <!--
     <br />
	 <a href="erp_fak.asp?FM_usedatokri=1&FM_job=<%=jobid%>&FM_aftale=<%=aftid%>&jobelaft=<%=jobelaft%>&FM_kunde=<%=kid%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd %>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>" class=vmenu target="_blank">Opret Faktura på valgte job/aftale og med den angivne fkatura dato.</a>
	 -->
 
   
	<%end if 'validering%>
	