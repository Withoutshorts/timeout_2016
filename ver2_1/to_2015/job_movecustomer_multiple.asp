

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->


<%

    func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function


    call menu_2014()

    %>

       

    <div id="wrapper">
        <div class="content">    

    <%
    select case func


    case "flyt"
	'*** Her indsættes en ny type i db ****
		if len(request("FM_jids")) = 0 OR len(trim(request("FM_kunde"))) = 0 then
		
		errortype = 185
		call showError(errortype)
		
		else
		

                    jobids = split(request("FM_jids"), ", ")
                    kid = request("FM_kunde")

                    kkundenavn = "Customer Not found on update multible jobs"
                    kknr = 0
                    strSQLk = "SELECT kid, kkundenavn, kkundenr FROM kunder WHERE kid = "& kid
                    
                    'Response.write strSQLk
                    'Response.flush
        
                    oRec.open strSQLk, oConn, 3
                    if not oRec.EOF then

                    kkundenavn = replace(oRec("kkundenavn"), "'", "''")
                    kknr = oRec("kkundenr")
                    
                    end if
                    oRec.close


                    for j = 0 TO UBOUND(jobids)
          
                				jobnr = 0			
                                strSQLjobnrs = "SELECT jobnr FROM job WHERE id = "& jobids(j)
                                oRec.open strSQLjobnrs, oConn, 3
                                if not oRec.EOF then

                                jobnr = oRec("jobnr")
                    
                                end if
                                oRec.close


                    '*** Opdaterer job *****'
                   
					strSQLjob = "UPDATE job SET "_
					& " jobknr = "& kid &""_
				    & " WHERE id = "& jobids(j)
                								
				    'Response.write strSQLtimer
					'Response.flush
                								
					oConn.execute(strSQLjob)



					'*** Opdaterer timereg tabellen ****
					strSQLtimer = "UPDATE timer SET "_
					& " Tknavn = '"& kkundenavn &"', Tknr = "& kid &""_
				    & " WHERE tjobnr = '"& jobnr & "'"
                								
				    'Response.write strSQLtimer
					'Response.flush
                								
					oConn.execute(strSQLtimer)
								                
								                
					'*** Opdaterer faktura tabel så faktura kunde id passer hvis der er skiftet kunde  ved rediger job.
					'*** Adr. i adresse felt på faktura behodes til revisor spor. **'
					strSQLFakadr = "UPDATE fakturaer SET fakadr = "& kid &" WHERE jobid = " & jobids(j)
					oConn.execute(strSQLFakadr)
								
						
                    next
                
        		
					
		

		
		Response.redirect "job_movecustomer_multiple.asp"
		end if


    case else
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	




%>


<div class="container">
    <div class="portlet">
            <h3 class="portlet-title">
            <u>Flyt job til ny kunde</u>
            </h3>

        <form action="job_movecustomer_multiple.asp?func=flyt" method="post">
             <input type="hidden" name="id" value="<%=id%>">
        <div class="row">
            <div class="col-lg-10">&nbsp</div>
            <div class="col-lg-2 pad-b10">
            <button type="submit" class="btn btn-success btn-sm pull-right"><b>Opdatér</b></button>
            </div>
        </div>

        <div class="portlet-body">
            <section>
                <div class="well well-white">
           
           
          

      
             <div class="row">
                <div class="col-lg-1">&nbsp</div>
                <div class="col-lg-6">
                    Flytter job fra oprindelige kunde til den valgte nedenstående kunde.<br />
                    Gælder også aktiviteter og fakturaer og timeforbrug.
                    <br /><br /><b>Vælg job og tilbud:</b> (vælg gerne flere)<br />
                    <%

                   whSQL = "(jobstatus = 1 OR jobstatus = 3)"
                    strSQLjob = "SELECT jobnavn, jobnr, j.id, kid, kkundenavn, kkundenr, count(a.id) AS antalA FROM job j "_
		            &" LEFT JOIN kunder ON (kid = jobknr) "_
		            &" LEFT JOIN aktiviteter a ON (a.job = j.id) "_
                    &" WHERE "& whSQL &" GROUP BY j.id ORDER BY kkundenavn, jobnavn"
		
		            'Response.Write strSQLjob
		            'Response.flush
		
		            %>
		
            <select name="FM_jids" multiple size=10 style="width:700px;">
          
              
		        <%
		
		        oRec2.open strSQLjob, oConn, 3
		        lastkid = 0
                antaljob = 0
		        while not oRec2.EOF


                    if lastkid <> oRec2("kid") then

                        if antaljob <> 0 then
                        %>
                        <option DISABLED></option>
                        <%end if

                    %><option DISABLED><%=oRec2("kkundenavn") %> (<%=oRec2("kkundenr") %>)</option>
                    <%end if %>
                        
                                <%
                                 antalF = 0   
                                 strSQLantalFak = "SELECT count(fid) AS antalF FROM fakturaer WHERE jobid = "& oRec2("id") & " GROUP BY jobid"
                                 oRec6.open strSQLantalFak, oConn, 3
                                if not oRec6.EOF then

                                antalF = oRec6("antalF")
                    
                                end if
                                oRec6.close
                                %>


		          <option value="<%=oRec2("id") %>"><%=oRec2("jobnavn") & " ("& oRec2("jobnr") &")"%> --------------------------- Antal akt: <%=oRec2("antalA") & " fak.: "& antalF %> </option>
           
		        <%
                lastKid = oRec2("kid")    
                antaljob = antaljob + 1 
		        oRec2.movenext
		        Wend
		        oRec2.close
		
		        %>
		         </select>
		         <br />
                Antal aktive job og tilbud: <%=antaljob %>
                    
                </div>
            </div>
            <div class="row">
                <div class="col-lg-1">&nbsp</div>
                <div class="col-lg-6"><br /><br /><b>Flyt til kunde:</b> (vælg 1)<br />
                    <select name="FM_kunde" size=10 style="width:700px;">
                    <%strSQL = "SELECT Kkundenavn, Kkundenr, Kid, kundeans1, kundeans2 FROM kunder WHERE ketype <> 'e' AND (useasfak = 1 OR useasfak = 0 OR useasfak = 5) ORDER BY Kkundenavn"
			        oRec.open strSQL, oConn, 3
			        kans1 = ""
			        kans2 = ""
			        while not oRec.EOF
				
                     %>
			        <option value="<%=oRec("Kid")%>"><%=oRec("Kkundenavn")%>&nbsp;&nbsp;(<%=oRec("Kkundenr")%>)</option>
			        <%
				
			        oRec.movenext
			        wend
			        oRec.close %>
                    </select>
                </div>
            </div>


        
                    
        </div>


             

         </section>

            
        </div>
        </form>
    </div>
</div>



    


<%end select  %>

</div>
</div>

<!--#include file="../inc/regular/footer_inc.asp"-->