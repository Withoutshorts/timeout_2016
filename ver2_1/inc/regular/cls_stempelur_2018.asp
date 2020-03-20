 
<%
function fLonTimerPer_html_2018(stDato, periode, visning, medid)
	
	call thisWeekNo53_fn(stDato)
    thisWeekNo53_stDato = thisWeekNo53

    call thisWeekNo53_fn(now)
    thisWeekNo53_now = thisWeekNo53


	select case visning 
	case 0, 20%>


	<%if visning <> 20 then %>
	

  

<style>

    /* The Modal (background) */
    .modal {
        display: none; /* Hidden by default */
        position: fixed; /* Stay in place */
        z-index: 1; /* Sit on top */
        padding-top: 100px; /* Location of the box */
        left: 0;
        top: 0;
        width: 100%; /* Full width */
        height: 100%; /* Full height */
        overflow: auto; /* Enable scroll if needed */
        background-color: rgb(0,0,0); /* Fallback color */
        background-color: rgba(0,0,0,0.4); /* Black w/ opacity */
    }

    /* Modal Content */
    .modal-content {
        background-color: #fefefe;
        margin: auto;
        padding: 20px;
        border: 1px solid #888;
        width: 1100px;
        height: 450px;
    }

    .picmodal:hover,
    .picmodal:focus {
    text-decoration: none;
    cursor: pointer;
}
   
</style>


    <br /><br /><br /><br />
    <div class="row">
        <div class="col-lg-4">

            <!-- Denne uge, nuværende login -->
             <%call erStempelurOn() 

              

            if cint(stempelur_hideloginOn) = 0 then 'skriv ikke login, men åben komme/gå tom

		    if cint(thisWeekNo53_stDato) = cint(thisWeekNo53_now) then%>
		    <%=tsa_txt_134 %>:
		    <%
		
		    sLoginTid = "00:00:00"
		
		    strSQL = "SELECT l.id AS lid, l.login "_
		    &" FROM login_historik l WHERE "_
		    &" l.mid = " & medid &" AND stempelurindstilling <> -1"_
		    &" ORDER BY l.id DESC" 
					
		    'Response.write strSQL
		    'Response.flush
		
		
		    oRec.open strSQL, oConn, 3 
		    if not oRec.EOF then 
        
            sLoginTid = oRec("login") 
        
            end if
            oRec.close
		
		    %>
		    <b><%=formatdatetime(sLoginTid, 3) %></b>
		
		    <% 
		    logindiffSidste = datediff("n", sLoginTid, now, 2, 2) 
		
            
            %>
		    <br /><%=tsa_txt_340 %>
		    <%call timerogminutberegning(logindiffSidste) %>
		    <b><%=thoursTot &":"& left(tminTot, 2) %>&nbsp;t.</b>
		
		
		    <%end if '*** Denne uge / nuværende login **'
              end if ' if cint(stempelur_hideloginOn) = 1 then 'skriv ikke login, men åben komme/gå tom %>

        </div>
    </div>
		
		
		
	<table class="table dataTable table-bordered table-hover ui-datatable">
    
    
    <thead>
	    <tr>
		    <td style="width:100px;">
                &nbsp;</td>
		
		    <th style="text-align:center"><%=tsa_txt_128 %> d. <%=formatdatetime(stDato, 2) %></th>
		    <th style="text-align:center"><%=tsa_txt_129 %> d. <%=formatdatetime(dateadd("d", 1, stDato), 2) %></th>
		    <th style="text-align:center"><%=tsa_txt_130 %> d. <%=formatdatetime(dateadd("d", 2, stDato), 2) %></th>
		    <th style="text-align:center"><%=tsa_txt_131 %> d. <%=formatdatetime(dateadd("d", 3, stDato), 2) %></th>
		    <th style="text-align:center"><%=tsa_txt_132 %> d. <%=formatdatetime(dateadd("d", 4, stDato), 2) %></th>
	        <th style="text-align:center"><%=tsa_txt_133 %> d. <%=formatdatetime(dateadd("d", 5, stDato), 2) %></th>
		    <th style="text-align:center"><%=tsa_txt_127 %> d. <%=formatdatetime(dateadd("d", 6, stDato), 2) %></th>
		    <th style="text-align:center"><%=global_txt_167 %></th>
	    </tr>

	    <%
        cspsStur = 1
        else 
        cspsStur = 2%>

        <tr><th colspan=10><br /><br /><b><%=tsa_txt_340 %></b> </th></tr>
    </thead>


    <%end if 'visning %>



    <tbody>






    <%if cint(showkgtim) = 1 then %>
    <tr>
		<td colspan="<%=cspsStur %>"><%=tsa_txt_137 %>:</td>
		<td style="text-align:right"><%call timerogminutberegning(manMin)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIalt = ugeIalt + manMin/1%>
		<td style="text-align:right"><%call timerogminutberegning(tirMin)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIalt = ugeIalt + tirMin/1%>
		<td style="text-align:right"><%call timerogminutberegning(onsMin)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIalt = ugeIalt + onsMin/1%>
		<td style="text-align:right"><%call timerogminutberegning(torMin)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIalt = ugeIalt +  torMin/1%>
		<td style="text-align:right"><%call timerogminutberegning(freMin)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIalt = ugeIalt +  freMin/1%>
		<td style="text-align:right"><%call timerogminutberegning(lorMin)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIalt = ugeIalt +  lorMin/1%>
		<td style="text-align:right"><%call timerogminutberegning(sonMin)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIalt = ugeIalt +  sonMin/1%>
		<td style="text-align:right">= 
		<%call timerogminutberegning(ugeIalt)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
	</tr>
    <%
     loginTimerTot = ugeIalt - (ugeIaltPause - ugeIaltFraTilTimer)     
    end if %>


    <%
       



    if cint(showkgpau) = 1 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="<%=cspsStur %>"><%=tsa_txt_138 %>:</td>
		<td style="text-align:right"><%call timerogminutberegning(-manMinPause)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIaltPause = ugeIaltPause + manMinPause%>
		<td style="text-align:right"><%call timerogminutberegning(-tirMinPause)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIaltPause = ugeIaltPause + tirMinPause%>
		<td style="text-align:right"><%call timerogminutberegning(-onsMinPause)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIaltPause = ugeIaltPause + onsMinPause%>
		<td style="text-align:right"><%call timerogminutberegning(-torMinPause)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIaltPause = ugeIaltPause + torMinPause%>
		<td style="text-align:right"><%call timerogminutberegning(-freMinPause)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIaltPause = ugeIaltPause + freMinPause%>
		<td style="text-align:right"><%call timerogminutberegning(-lorMinPause)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIaltPause = ugeIaltPause + lorMinPause%>
		<td style="text-align:right"><%call timerogminutberegning(-sonMinPause)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIaltPause = ugeIaltPause + sonMinPause%>
		<td style="text-align:right">= <%call timerogminutberegning(-ugeIaltPause)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
	</tr>
	<%end if %>

	<!-- Fradrag / Tillæg via Realtimer -->
	
    <%if cint(showkgtil) = 1 then %>
    <tr bgcolor="#ffffff">
		<td colspan="<%=cspsStur %>" style="white-space:nowrap;"><span id="modal_1" class="picmodal" ><u><%=global_txt_168 %>: +</u></span></td>
		<td style="text-align:right"><%call timerogminutberegning(manFraTimer)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIaltFraTilTimer = ugeIaltFraTilTimer + (manFraTimer)%>
		<td style="text-align:right"><%call timerogminutberegning(tirFraTimer)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIaltFraTilTimer = ugeIaltFraTilTimer + (tirFraTimer)%>
		<td style="text-align:right"><%call timerogminutberegning(onsFraTimer)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIaltFraTilTimer = ugeIaltFraTilTimer + (onsFraTimer)%>
		<td style="text-align:right"><%call timerogminutberegning(torFraTimer)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIaltFraTilTimer = ugeIaltFraTilTimer + (torFraTimer)%>
		<td style="text-align:right"><%call timerogminutberegning(freFraTimer)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIaltFraTilTimer = ugeIaltFraTilTimer + (freFraTimer)%>
		<td style="text-align:right"><%call timerogminutberegning(lorFraTimer)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIaltFraTilTimer = ugeIaltFraTilTimer + (lorFraTimer)%>
		<td style="text-align:right"><%call timerogminutberegning(sonFraTimer)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		<%ugeIaltFraTilTimer = ugeIaltFraTilTimer + (sonFraTimer)%>
		<td style="text-align:right">= <%call timerogminutberegning(ugeIaltFraTilTimer)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
	</tr>
	<%end if %>

 


    <%if cint(showkgtot) = 1 then %>
	<!-- total -->
	
	
	 <tr bgcolor="#cccccc">
		<td colspan="<%=cspsStur %>"><b><%=global_txt_167%>:</b></td>
		<td style="text-align:right"><b><%call timerogminutberegning(totMan)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </b></td>
		<td style="text-align:right"><b><%call timerogminutberegning(totTir)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </b></td>
		<td style="text-align:right"><b><%call timerogminutberegning(totOns)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </b></td>
		<td style="text-align:right"><b><%call timerogminutberegning(totTor)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </b></td>
	    <td style="text-align:right"><b><%call timerogminutberegning(totFre)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </b></td>
		<td style="text-align:right"><b><%call timerogminutberegning(totLor)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </b></td>
		<td style="text-align:right"><b><%call timerogminutberegning(totSon)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </b></td>
		<td style="text-align:right">= <b><%call timerogminutberegning(ugeIalt - (ugeIaltPause - (ugeIaltFraTilTimer)))
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </b></td>
		
	</tr>
	<%end if %>
 
	<!-- Normtimer -->
	
    <%if cint(showkgnor) = 1 then %>
    <tr bgcolor="#ffffff">
		<td colspan="<%=cspsStur %>"><%=tsa_txt_259 %>:</td>
		
		<%call normtimerper(medid, varTjDatoUS_man, 6, 0) %>
		
		<td style="text-align:right"><%call timerogminutberegning(ntimMan*60)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		
		<td style="text-align:right"><%call timerogminutberegning(ntimTir*60)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		
		<td style="text-align:right"><%call timerogminutberegning(ntimOns*60)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		
		<td style="text-align:right"><%call timerogminutberegning(ntimTor*60)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		
		<td style="text-align:right"><%call timerogminutberegning(ntimFre*60)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		
		<td style="text-align:right"><%call timerogminutberegning(ntimLor*60)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		
		
		<td style="text-align:right"><%call timerogminutberegning(ntimSon*60)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		
		<%
		NormTimerWeekTot = 0
		NormTimerWeekTot = (ntimMan + ntimTir + ntimOns + ntimTor + ntimFre + ntimLor + ntimSon) * 60 %>
		
		<td style="text-align:right">= <%call timerogminutberegning(NormTimerWeekTot)
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </td>
		</tr>

      <%end if %>
	    
	  <!-- Saldo -->
      
      <%if cint(showkgsal) = 1 then %>  
	  <tr bgcolor="#DCF5BD">
		<td colspan="<%=cspsStur %>" style="height:20px;"><b><%=global_txt_163 %>:</b></td>
		<td style="text-align:right"><b><%call timerogminutberegning(totMan - (ntimMan*60))
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </b></td>
		<td style="text-align:right"><b><%call timerogminutberegning(totTir - (ntimTir*60))
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </b></td>
		<td style="text-align:right"><b><%call timerogminutberegning(totOns - (ntimOns*60))
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </b></td>
		<td style="text-align:right"><b><%call timerogminutberegning(totTor - (ntimTor*60))
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </b></td>
	    <td style="text-align:right"><b><%call timerogminutberegning(totFre - (ntimFre*60))
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </b></td>
		<td style="text-align:right"><b><%call timerogminutberegning(totLor - (ntimLor*60))
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </b></td>
		<td style="text-align:right"><b><%call timerogminutberegning(totSon - (ntimSon*60))
		Response.write ""& thoursTot &":"& left(tminTot, 2)%> </b></td>
		<td style="text-align:right">=
        <%
        call timerogminutberegning((ugeIalt - (ugeIaltPause - (ugeIaltFraTilTimer))) - NormTimerWeekTot)
        %>

        <%
        stDatoUS = year(stDato) & "/" & month(stDato) &"/"& day(stDato)  
        slDatoUS = year(slutDato) & "/" & month(slutDato) &"/"& day(slutDato)
        %>
        <a href="afstem_tot.asp?usemrn=<%=usemrn%>&show=5&varTjDatoUS_man=<%=stDatoUS%>&varTjDatoUS_son=<%=slDatoUS%>" class=vmenu><%=thoursTot &":"& left(tminTot, 2)%></a></td>
		
	</tr>
    <%end if %>


    <%if visning <> 20 then '** Ikke fra timereg. siden da det er for tungt  %>
        <%if cint(showkgsaa) = 1 then '*** IKKE aktiv - se aftamte tot istedet for
        
        
        %>


        <%end if %>
    <%end if %>

   </tbody>


    <%if visning <> 20 then %>
	</table>
	
    <%if cint(stempelur_hideloginOn) = 0 then 'skriv ikke login, men åben komme/gå tom
	
    if cint(thisWeekNo53_stDato) = cint(thisWeekNo53_now) then%>
	<!-- Denne uge incl. nuværende  -->
	<%=tsa_txt_139 %>: <% 
	call timerogminutberegning(logindiffSidste+(loginTimerTot))
	%>
	<b><%=thoursTot &":"& left(tminTot, 2) %>&nbsp;t.</b>

 
	
    <%end if ' *** Denne uge / Nuværende login **'
    end if    
    %>

	

    
	<br /><br />
	<table><tr><td >
	
	</td></tr></table>

    <%
	if cint(showkguds) = 1 OR lto = "intranet - local" then

    ' Ligger i modal nu
	'Response.Write "*) <b> Tillægs typer:</b> " & akttypenavnTil
	'Response.Write "<br><b>Fradrags typer:</b> " & akttypenavnFra
	
	%>

    <%if media <> "print" then %>
    <div id="myModal_1" class="modal">

        <!-- Modal content -->
        <div class="modal-content"> 
            <div id="udspecdiv" style="position:relative; width:100%;">

                <div class="row">
                    <div class="col-lg-6">
                        <b> Tillægs typer:</b> <%=akttypenavnTil %>
                    </div>
                </div>
                <div class="row">
                    <div class="col-lg-6">
                        <b>Fradrags typer:</b> <%=akttypenavnFra %>
                    </div>
                </div>

                <br />

                <div class="row">
                    <div class="col-lg-12"><b>Udspecificering på fraværstyper</b>
                        <br /> Ikke medregnet i saldo, med mindre de er en del af <%=global_txt_168 %> typerne *
                    </div>
                </div>

	            <table class="table dataTable table-bordered table-hover ui-datatable">    

                    <thead>
	                    <tr>
		                    <td style="width:100px;">&nbsp;</td>	
		                    <td style="text-align:right"><b><%=tsa_txt_128 %> d. <%=formatdatetime(stDato, 2) %></b></td>
		                    <td style="text-align:right"><b><%=tsa_txt_129 %> d. <%=formatdatetime(dateadd("d", 1, stDato), 2) %></b></td>
		                    <td style="text-align:right"><b><%=tsa_txt_130 %> d. <%=formatdatetime(dateadd("d", 2, stDato), 2) %></b></td>
		                    <td style="text-align:right"><b><%=tsa_txt_131 %> d. <%=formatdatetime(dateadd("d", 3, stDato), 2) %></b></td>
		                    <td style="text-align:right"><b><%=tsa_txt_132 %> d. <%=formatdatetime(dateadd("d", 4, stDato), 2) %></b></td>
	                        <td style="text-align:right"><b><%=tsa_txt_133 %> d. <%=formatdatetime(dateadd("d", 5, stDato), 2) %></b></td>
		                    <td style="text-align:right"><b><%=tsa_txt_127 %> d. <%=formatdatetime(dateadd("d", 6, stDato), 2) %></b></td>
		                    <td style="text-align:right"><b><%=global_txt_167 %></b></td>
	                    </tr>
                    </thead>
	
                        
                    <tbody>
	    
	                    <%if cint(aty_fleks_on) = 1 then %>
	                    <!-- Fleks Realtimer -->
	                    <tr>
		                    <td style="text-align:right"><%=global_txt_147 &" "& aty_fleks_tf%></td>
		                    <td style="text-align:right"><%call timerogminutberegning(manFlekstimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltFlekstimer = ugeIaltFlekstimer + (manFlekstimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(tirFlekstimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltFlekstimer = ugeIaltFlekstimer + (tirFlekstimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(onsFlekstimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltFlekstimer = ugeIaltFlekstimer + (onsFlekstimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(torFlekstimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltFlekstimer = ugeIaltFlekstimer + (torFlekstimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(freFlekstimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltFlekstimer = ugeIaltFlekstimer + (freFlekstimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(lorFlekstimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltFlekstimer = ugeIaltFlekstimer + (lorFlekstimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(sonFlekstimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltFlekstimer = ugeIaltFlekstimer + (sonFlekstimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(ugeIaltFlekstimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	                    </tr>
	
	                    <%end if %>
	
	
	
	                    <%if cint(aty_Ferie_on) = 1 then %>
	                    <!-- Ferie Realtimer -->
	                    <tr>
		                    <td style="text-align:right"><%=global_txt_135 &" "& aty_Ferie_tf %></td>
		                    <td style="text-align:right"><%call timerogminutberegning(manFerietimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltFerietimer = ugeIaltFerietimer + (manFerietimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(tirFerietimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltFerietimer = ugeIaltFerietimer + (tirFerietimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(onsFerietimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltFerietimer = ugeIaltFerietimer + (onsFerietimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(torFerietimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltFerietimer = ugeIaltFerietimer + (torFerietimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(freFerietimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltFerietimer = ugeIaltFerietimer + (freFerietimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(lorFerietimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltFerietimer = ugeIaltFerietimer + (lorFerietimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(sonFerietimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltFerietimer = ugeIaltFerietimer + (sonFerietimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(ugeIaltFerietimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	                    </tr>
	
	                    <%end if %>
	
	                    <%if cint(aty_Syg_on) = 1 then %>
	                    <!-- Syg Realtimer -->
	                    <tr>
		                    <td style="text-align:right"><%=global_txt_138 &" "& aty_Syg_tf %></td>
		                    <td style="text-align:right"><%call timerogminutberegning(manSygtimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltSygtimer = ugeIaltSygtimer + (manSygtimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(tirSygtimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltSygtimer = ugeIaltSygtimer + (tirSygtimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(onsSygtimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltSygtimer = ugeIaltSygtimer + (onsSygtimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(torSygtimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltSygtimer = ugeIaltSygtimer + (torSygtimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(freSygtimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltSygtimer = ugeIaltSygtimer + (freSygtimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(lorSygtimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltSygtimer = ugeIaltSygtimer + (lorSygtimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(sonSygtimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltSygtimer = ugeIaltSygtimer + (sonSygtimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(ugeIaltSygtimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	                    </tr>
	
	                    <%end if %>
	
	
	                    <%if cint(aty_BarnSyg_on) = 1 then %>
	                    <!-- BarnSyg Realtimer -->
	                    <tr>
		                    <td style="text-align:right"><%=global_txt_139 &" "& aty_BarnSyg_tf %></td>
		                    <td style="text-align:right"><%call timerogminutberegning(manBarnSygtimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltBarnSygtimer = ugeIaltBarnSygtimer + (manBarnSygtimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(tirBarnSygtimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltBarnSygtimer = ugeIaltBarnSygtimer + (tirBarnSygtimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(onsBarnSygtimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltBarnSygtimer = ugeIaltBarnSygtimer + (onsBarnSygtimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(torBarnSygtimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltBarnSygtimer = ugeIaltBarnSygtimer + (torBarnSygtimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(freBarnSygtimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltBarnSygtimer = ugeIaltBarnSygtimer + (freBarnSygtimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(lorBarnSygtimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltBarnSygtimer = ugeIaltBarnSygtimer + (lorBarnSygtimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(sonBarnSygtimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltBarnSygtimer = ugeIaltBarnSygtimer + (sonBarnSygtimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(ugeIaltBarnSygtimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	                    </tr>
	
	                    <%end if %>
	
	
	                    <%if cint(aty_Lage_on) = 1 then %>
	                    <!-- Lage Realtimer -->
	                    <tr>
		                    <td style="text-align:right"><%=global_txt_160 &" "& aty_Lage_tf %></td>
		                    <td style="text-align:right"><%call timerogminutberegning(manLagetimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltLagetimer = ugeIaltLagetimer + (manLagetimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(tirLagetimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltLagetimer = ugeIaltLagetimer + (tirLagetimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(onsLagetimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltLagetimer = ugeIaltLagetimer + (onsLagetimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(torLagetimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltLagetimer = ugeIaltLagetimer + (torLagetimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(freLagetimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltLagetimer = ugeIaltLagetimer + (freLagetimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(lorLagetimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltLagetimer = ugeIaltLagetimer + (lorLagetimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(sonLagetimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltLagetimer = ugeIaltLagetimer + (sonLagetimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(ugeIaltLagetimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	                    </tr>
	
	                    <%end if %>
	
		                    <%if cint(aty_Sund_on) = 1 then %>
	                    <!-- Sund Realtimer -->
	                    <tr>
		                    <td style="text-align:right"><%=global_txt_148 &" "& aty_Sund_tf %></td>
		                    <td style="text-align:right"><%call timerogminutberegning(manSundtimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltSundtimer = ugeIaltSundtimer + (manSundtimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(tirSundtimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltSundtimer = ugeIaltSundtimer + (tirSundtimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(onsSundtimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltSundtimer = ugeIaltSundtimer + (onsSundtimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(torSundtimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltSundtimer = ugeIaltSundtimer + (torSundtimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(freSundtimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltSundtimer = ugeIaltSundtimer + (freSundtimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(lorSundtimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltSundtimer = ugeIaltSundtimer + (lorSundtimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(sonSundtimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltSundtimer = ugeIaltSundtimer + (sonSundtimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(ugeIaltSundtimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	                    </tr>
	
	                    <%end if %>    
	
	                    <%if cint(aty_Frokost_on) = 1 then %>
	                    <!-- Frokost Realtimer -->
	                    <tr>
		                    <td style="text-align:right"><%=global_txt_133 &" "& aty_Frokost_tf %></td>
		                    <td style="text-align:right"><%call timerogminutberegning(manFrokosttimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltFrokosttimer = ugeIaltFrokosttimer + (manFrokosttimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(tirFrokosttimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltFrokosttimer = ugeIaltFrokosttimer + (tirFrokosttimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(onsFrokosttimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltFrokosttimer = ugeIaltFrokosttimer + (onsFrokosttimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(torFrokosttimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltFrokosttimer = ugeIaltFrokosttimer + (torFrokosttimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(freFrokosttimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltFrokosttimer = ugeIaltFrokosttimer + (freFrokosttimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(lorFrokosttimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltFrokosttimer = ugeIaltFrokosttimer + (lorFrokosttimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(sonFrokosttimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		                    <%ugeIaltFrokosttimer = ugeIaltFrokosttimer + (sonFrokosttimer)%>
		                    <td style="text-align:right"><%call timerogminutberegning(ugeIaltFrokosttimer)
		                    Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	                    </tr>
	
	                    <%end if %>    
	    
	                </tbody>
	            </table>

   
	        </div>
            
        </div>

    </div>
    <%end if %>
	

    
    
  <% end if%>
    
    
    
	<%

    end if 'vis <> 3
	
	case 2
	
	totalTimerPer100 = 0
	totalTimerPer100 = (totaltimerPer-totalpausePer)
	
    call timerogminutberegning(totalTimerPer100) %>
	<%=thoursTot &":"& left(tminTot, 2) %>
	
	<%
	case 3
	
	totalTimerPer100 = 0
	totalTimerPer100 = (totaltimerPer-totalpausePer)
	
	case 4 ''** Avg. Week ***' 
	
	if weekDiff <> 0 then
	weekDiff = weekDiff
	else
	weekDiff = 1
	end if
	
	totalTimerPer100 = 0
	totalTimerPer100 = (totaltimerPer-totalpausePer)/weekDiff
	
	
	call timerogminutberegning(((totaltimerPer-totalpausePer)/weekDiff)) %>
	<!-- formatnumber(totalTimerPer100, 2) -->
	
	<%=thoursTot &":"& left(tminTot, 2) %>
	<%

    case 21 'smiley på timereg.
	

     totalTimerPer100 = 0
	totalTimerPer100 = (totaltimerPer-totalpausePer)
	
    %>
	
     
    <% 
    case else
	
	
	totalTimerPer100 = 0
	totalTimerPer100 = (totaltimerPer-totalpausePer)
	
	
	call timerogminutberegning(totaltimerPer-totalpausePer) %>
	<!-- formatnumber(totalTimerPer100, 2) -->
	
	<%=thoursTot &":"& left(tminTot, 2) %> 
	
	<%
	end select '** Visinng ***'%>
	


<%
end function





















'public lonTimerGT, stempelUrEkspTxtShowTot   
function tottimer_2018(lastMnavn, lastMnr, totalhours, totalmin, lastMid, sqlDatoStart, sqlDatoslut, vis, lastDato)
		

        'if session("mid") = 1 then
        'Response.write totalhours &"+"& totalmin & "<br>"
        'end if

        'vis 1: total pr. uge
        'vis 2: GT i bunden af hver medarbejder

		totalmin = (60 * totalhours) + totalmin
		lonTimer = totalmin
		call timerogminutberegning(totalmin)

        if vis = 2 OR (cint(showTot) = 1 AND vis = 1) then
        lonTimerGT = lonTimerGT + lonTimer
	    end if	

		if cint(showTot) = 1 then
            
                select case lto
                case "cst"
	            csp = 5
                case else
                csp = 6
                end select

		
		else
		csp = 9
		end if

        ugeNummer = day(lastDato) & "/" & month(lastDato) & "/" & year(lastDato)
        'ugeNummer = 

        if cint(vis) = 1 then 'total i bunden / else uge sum
         trBg = "#FFFFFF"
        else
        trBg = "lightpink"
        end if
        %>
      
    <thead>
        <%if cint(showTot) = 1 AND vis = 1 then %>

                <% if media <> "export" then %>
    
                <tr>
                <!--<td style="text-align:right">&nbsp;</td> -->
                    <td style="text-align:right"><%=lastMnavn & " ["& lastMinit &"]" %></td>
                    <td style="text-align:right"><%=formatdatetime(sqlDatoStart, 1) &" - "& formatdatetime(sqlDatoSlut, 1) %></td>
                    <td align=right><%=thoursTot%>:<%=left(tminTot, 2)%></td>

      

                <%
                else

                stempelUrEkspTxtShowTot = stempelUrEkspTxtShowTot & ""& meNavn &";"& meInit & ";" & sqlDatoStart & ";" & sqlDatoSlut & ";" 
                stempelUrEkspTxtShowTot = stempelUrEkspTxtShowTot & ""& thoursTot&"."& left(tminTot, 2) &";"

                end if
     

        else 
        
                call thisWeekNo53_fn(ugeNummer)

                if vis <> 1 then %>
                <tr>
		               
                        <th colspan=<%=csp-4%>>Løntimer (komme/gå) uge <%=thisWeekNo53 %>:</th>
                        <th style="text-align:right"><b><%=thoursTot%>:<%=left(tminTot, 2)%></b></th>
                        <th colspan="100">&nbsp;</th>
		             
	             </tr>
		        <%end if %>      
	    <%end if %>
    </thead>
    
        <%if vis = 1 then
        
        trBg = "#FFFFFF"

        if cint(showTot) <> 1 then
        %>
       <!-- <tr>
		    <td colspan="<%=csp+3%>" style="height:10px;">&nbsp;</td>
	    </tr> -->
        <%end if %>


      <%if cint(showTot) <> 1 then%>
	 <tr>
		<!--<td style="text-align:right">&nbsp;</td>-->
		<td colspan=<%=csp-4%>>Fradrag til løntimer:</td>
        <%end if 
            
            
         if media <> "export" then%>
        <td style="text-align:right;">
		<%
        end if

		call akttyper2009(2)
		
		tiltimer = 0
        
            stDtKriTL = year(sqlDatoStart)&"/"&month(sqlDatoStart)&"/"&day(sqlDatoStart)
            slDtKriTL = year(sqlDatoSlut)&"/"&month(sqlDatoSlut)&"/"&day(sqlDatoSlut)

		strSQL2 = "SELECT sum(t.timer) AS tiltimer FROM timer t WHERE "_
		&" (tmnr = "& lastMid &" AND tdato BETWEEN '"& stDtKriTL &"' AND '"& slDtKriTL &"' AND ("& aty_sql_tilwhours &")) GROUP BY tmnr " 
		
            'if session("mid") = 1 then
            'Response.Write strSQL2
		'Response.flush
		'end if

		oRec2.open strSQL2, oConn, 3 
		if not oRec2.EOF then
			tiltimer = oRec2("tiltimer")
		end if
		oRec2.close 
		
		if len(trim(tiltimer)) <> 0 then
		tiltimer = tiltimer
		else
		tiltimer = 0
		end if
		
		fradtimer = 0
		strSQL2 = "SELECT sum(timer) AS fratimer FROM timer t WHERE "_
		&" (tmnr = "& lastMid &" AND tdato BETWEEN '"& stDtKriTL &"' AND '"& slDtKriTL &"' AND ("& aty_sql_frawhours &")) GROUP BY tmnr "
		'Response.Write strSQL2
		'Response.flush
		
		oRec2.open strSQL2, oConn, 3 
		if not oRec2.EOF then
			fradtimer = oRec2("fratimer")
		end if
		oRec2.close 
		
		if len(trim(fradtimer)) <> 0 then
		fradtimer = fradtimer
		else
		fradtimer = 0
		end if
		
		'Response.Write "fradtimer" & fradtimer & "# tilt#"& tiltimer
		
		totalmin = (60 * (-(fradtimer) + (tiltimer)))
		fradragTil = totalmin
		call timerogminutberegning(totalmin)
		
        if media <> "export" then
		%>
		&nbsp;&nbsp;<%=thoursTot%>:<%=left(tminTot, 2)%>
	    </td>
         <%end if %>

		<%if showTot <> 1 then%>
		    <td colspan="100">&nbsp;</td>
		   <!-- <td style="text-align:right">&nbsp;</td>
		    <td style="text-align:right">&nbsp;</td>
            <td style="text-align:right">&nbsp;</td>
            <td style="text-align:right">&nbsp;10</td> -->
	     </tr>

        <%else
            
            if media = "export" then
            stempelUrEkspTxtShowTot = stempelUrEkspTxtShowTot & ""& thoursTot&"."& left(tminTot, 2) &";"
            end if
     
		end if %>
		
	 

     <%if showTot <> 1 then%>

	  <tr>
		<!--<td style="padding:2px; border-bottom:1px #999999 solid;">&nbsp;</td>-->
		<td colspan=<%=csp-4%>>Grandtotal Løntimer (komme/gå - fradrag):</td>
     <%end if %>

         <%if cint(showTot) <> 1 then%>
		<td style="text-align:right;">
        <%else 
            
            if media <> "export" then%>
        <td style="text-align:right;">
        <%end if
            end if %>

	    <%'*** Løn Timer minus fradarg *** %>
		<%lonTimerBeregnet = lonTimerGT + (fradragTil)
		call timerogminutberegning(lonTimerBeregnet) %>

        <%if media <> "export" then %>
		&nbsp;&nbsp;<b><%=thoursTot%>:<%=left(tminTot, 2)%></b></td>
        <%end if

             lonTimerBeregnet = 0
             lonTimerGT = 0 
             fradragTil = 0 
        %>
		

          <%if showTot <> 1 then%>
        	<td colspan="100">&nbsp;</td>
		<!--
		<td style="text-align:right">&nbsp;</td>

		<td style="text-align:right">&nbsp;</td>
		<td style="text-align:right">&nbsp;13</td>
        <td style="text-align:right">&nbsp;14</td>
       <td style="text-align:right">&nbsp;15</td> -->
         </tr>
        <%else 

            if media = "export" then
            stempelUrEkspTxtShowTot = stempelUrEkspTxtShowTot & ""& thoursTot&"."& left(tminTot, 2) &";"
            end if

          
		end if %>
	
	
	 
	 <%if lto <> "cst" AND lto <> "cflow" then %>
	 
      <%if showTot <> 1 then%>
	  <tr bgcolor="<%=trBg %>">
		<!--<td style="text-align:right">&nbsp;</td>-->
		<td colspan=<%=csp-4%>>Realiseret timer (projekt):</td>
        <%end if 
            
            
        if media <> "export" then %>
        <td style="text-align:right;">
		<%end if
		
		
		regtimer = 0
		strSQL2 = "SELECT sum(timer) AS sumtimer FROM timer WHERE tmnr = "& lastMid &" AND tdato BETWEEN '"& stDtKriTL &"' AND '"& slDtKriTL &"' AND ("& aty_sql_realhours &") "
		'Response.Write strSQL2
		'Response.flush
		
		oRec2.open strSQL2, oConn, 3 
		if not oRec2.EOF then
			regtimer = oRec2("sumtimer")
		end if
		oRec2.close 
		
		totalmin = (60 * regtimer)
		call timerogminutberegning(totalmin)
		
		if media <> "export" then %>
		&nbsp;&nbsp;<%=thoursTot%>:<%=left(tminTot, 2)%>
		</td>
		<%end if
            
        if showTot <> 1 then%>
		<td colspan="100">&nbsp;</td>

		<!-- <td style="text-align:right">&nbsp;</td>
		<td style="text-align:right">&nbsp;</td>
        <td style="text-align:right">&nbsp;19</td>
        <td style="text-align:right">&nbsp;20</td> -->
	    </tr>
		<%
        else
            
             if media = "export" then
            stempelUrEkspTxtShowTot = stempelUrEkspTxtShowTot & ""& thoursTot&"."& left(tminTot, 2) &";"
            end if
          

        end if %>
		
	 <%end if %>

     <%
            
        '**** Overtid / Overarbejde *************    
        select case lto 
        case "cflow", "intranet - local"
        %>

         <%if showTot <> 1 then%>
         <tr>
		<!--<td style="text-align:right">&nbsp;</td> -->
		<td colspan=<%=csp-4%>>Overtid:</td>
          <%
         end if
         
              
        if media <> "export" then %>
        <td style="text-align:right;">
        <%end if

        overtid = 0
		strSQL2 = "SELECT sum(timer) AS overtid FROM timer t WHERE "_
		&" (tmnr = "& lastMid &" AND tdato BETWEEN '"& stDtKriTL &"' AND '"& slDtKriTL &"' AND (tfaktim = 30 AND godkendtstatus = 1)) GROUP BY tmnr "
		
        'if session("mid") = 1 then     
        'Response.Write strSQL2
		'Response.flush
        'end if
		
		oRec2.open strSQL2, oConn, 3 
		if not oRec2.EOF then
			overtid = oRec2("overtid")
		end if
		oRec2.close 
		
		if len(trim(overtid)) <> 0 then
		overtid = overtid
		else
		overtid = 0
		end if
		
		'Response.Write "fradtimer" & fradtimer & "# tilt#"& tiltimer
		
		totalmin = (60 * overtid)
		'fradragTil = totalmin
		call timerogminutberegning(totalmin)
		
		if media <> "export" then %>
		&nbsp;&nbsp;<%=thoursTot%>:<%=left(tminTot, 2)%>
        </td>
        <%end if
            
            
        if showTot <> 1 then%>
		<td colspan="100">&nbsp;</td>
		<!--<td style="text-align:right">&nbsp;</td>
		<td style="text-align:right">&nbsp;</td>
        <td style="text-align:right">&nbsp;</td>
       <td style="text-align:right">&nbsp;25</td> -->
       </tr>

        <%else 
        
             if media = "export" then
            stempelUrEkspTxtShowTot = stempelUrEkspTxtShowTot & ""& thoursTot&"."& left(tminTot, 2) &";"
            end if
           
            
        end if %>
	
	 


        
            
        <%end select%>


       <%if showTot = 1 then
           
            if media = "export" then
           stempelUrEkspTxtShowTot = stempelUrEkspTxtShowTot & "xx99123sy#z"
           end if

        %>
        </tr>
        <%else %>
       <!-- <tr><td colspan="10">&nbsp;<br /><br /><br /><br />&nbsp;</td></tr> -->
       <%end if %>

        


	 
     <%end if 'VIS: uge 2 ell. tot 1 %>

	<%
    Response.flush
	end function







  
   
   %>