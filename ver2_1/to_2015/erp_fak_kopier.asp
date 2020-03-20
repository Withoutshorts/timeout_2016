<%response.Buffer = true %>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->



<%
    if len(session("user")) = 0 then

  
	errortype = 5
	call showError(errortype)
	
	else

    
    thisfile = "erp_fak_kopier.asp"
    
    func = request("func")

    if len(trim(request("usemrn"))) <> 0 then
    usemrn = request("usemrn")
    else
    usemrn = 0
    end if

    if len(trim(request("fakid"))) <> 0 then
    fakid = request("fakid")
    else
    fakid = 0
    end if

    id = fakid

    erfakbetalt = 0
    faknr = "ikke fundet"
    strSQL = "SELECT faknr, erfakbetalt FROM fakturaer WHERE fid = " & fakid
    'Response.write strSQL
    'Response.Flush
    oRec.open strSQL, oConn, 3
    if not oRec.EOF then

    faknr = oRec("faknr")
    erfakbetalt = oRec("erfakbetalt")

    end if
    oRec.close

    '**SKAL kun kunne kopierer faktura IKEK kredit
    '** HUsk Epi LTO at kopiere rigtie rækkefølge nr.

    nytfaknr = 0
    call findFaknr("dbopr")
    nytfaknr = intFaknum

    lastFakId = 0
    strSQL = "SELECT fid FROM fakturaer WHERE fid <> 0 ORDER BY fid DESC"
    'Response.write strSQL
    'Response.Flush
    oRec.open strSQL, oConn, 3
    if not oRec.EOF then

    lastFakId = oRec("fid")

    end if
    oRec.close
    lastFakId = lastFakId +1


    lastFak_det_Id = 0
    strSQL = "SELECT id FROM faktura_det WHERE id <> 0 ORDER BY id DESC"
    'Response.write strSQL
    'Response.Flush
    oRec.open strSQL, oConn, 3
    if not oRec.EOF then

    lastFak_det_Id = oRec("id")

    end if
    oRec.close
    lastFak_det_Id = lastFak_det_Id + 1


    lastFak_mat_spec_Id = 0
    strSQL = "SELECT id FROM fak_mat_spec WHERE id <> 0 ORDER BY id DESC"
    'Response.write strSQL
    'Response.Flush
    oRec.open strSQL, oConn, 3
    if not oRec.EOF then

    lastFak_mat_spec_Id = oRec("id")

    end if
    oRec.close
    lastFak_mat_spec_Id = lastFak_mat_spec_Id + 1



     lastFak_med_spec_Id = 0
    strSQL = "SELECT id FROM fak_mat_spec WHERE id <> 0 ORDER BY id DESC"
    'Response.write strSQL
    'Response.Flush
    oRec.open strSQL, oConn, 3
    if not oRec.EOF then

    lastFak_med_spec_Id = oRec("id")

    end if
    oRec.close
    lastFak_med_spec_Id = lastFak_med_spec_Id + 1
    

   

     varTjDatoUS_man = day(now)&"-"&month(now)&"-"&year(now)
     varTjDatoUS_SQL = year(now)&"-"&month(now)&"-"&day(now)
     

    %>
    <div class="wrapper"><br /><br />
<!--  
    <div class="content">
      -->

        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><h3>Copy / move Invoice - <%=faknr %></h3>
                <div class="portlet-body">

                     <div class="row">
                       <div class="col-lg-6">
                        
                           
                           <%
                            select case func 
                            case "flyt"
                            %>

                           Du kan flytte denne faktura kladde til et nyt job (og dermed ny kunde)<br />
                           Modtager ref, bankkonto mm. skal ændres manuelt efter flytning. 
                            
                           <br /><br />
                           <form action="erp_fak_kopier.asp?func=flytja" method="post">
                               <input type="hidden" name="fakid" value="<%=fakid %>" />
                           <select name="nytjobid" class="form-control input-small" size="10">
                               <option value="0">Vælg job..</option>
                           <%
                           lastKID = 0
                           k = 0
                            strSQLnytjob = "SELECT jobnavn, id, kkundenavn, kid FROM job "_
                            &" LEFT JOIN kunder ON (kid = jobknr) WHERE jobstatus = 1 ORDER BY kkundenavn, jobnavn"
                            oRec.open strSQLnytjob, oConn, 3
                            while not oRec.EOF 

                               if cdbl(lastKID) <> cdbl(oRec("kid")) then

                               if k <> 0 then
                               %>
                               <option value="0" disabled></option>
                               <%end if %>

                               <option value="0" disabled><%=oRec("kkundenavn") %></option><%
                               lastKID = oRec("kid")
                               end if

                               %><option value="<%=oRec("id") %>"><%=oRec("jobnavn") %></option><%


                            k = k + 1
                            oRec.movenext
                            wend
                            oRec.close
                            %>
                          


                           </select>
                               <br /><br />
                               <input type="submit" class="btn btn-success btn-sm pull-right" value="Opdater" />

                            </form>



                            <%
                            case "flytja"

                                if len(trim(request("nytjobid"))) <> 0 AND request("nytjobid") <> "0" then
                                nytjobid = request("nytjobid")

                                kid = 0
                                strSQLnytjob = "SELECT jobnavn, id, kkundenavn, kid, jobnr FROM job "_
                                &" LEFT JOIN kunder ON (kid = jobknr) WHERE id = "& nytjobid  &" ORDER BY kkundenavn, jobnavn"
                                oRec.open strSQLnytjob, oConn, 3
                                if not oRec.EOF then
                                kid = oRec("kid")
                                jobnavn = oRec("jobnavn")
                                jobnr = oRec("jobnr")
                                kkundenavn = oRec("kkundenavn")
                                end if
                                oRec.close

                                if kid <> 0 then

                                '*** Opdaterer faktura tabel så faktura kunde id passer hvis der er skiftet kunde  ved rediger job.
					            '*** Adr. i adresse felt på faktura behodes til revisor spor. **'
					            strSQLFakadr = "UPDATE fakturaer SET fakadr = "& kid &" WHERE fid = " & fakid
					            oConn.execute(strSQLFakadr)


                                %>
                                
                                    <span style="color:yellowgreen"><b>Faktura <%=faknr%> er flyttet!</b></span>
                                    <br /><br />
                                    Fakturaen er flyttet til:
                                    <b><%=kkundenavn & "<br>"& jobnavn &" ("& jobnr &")" %></b>
                                   
                                  
                             
                                <%
                           
                                else
                                %>

                                <div style="position:relative; top:100px; left:100px; padding:20px;">
                                    <b>Fejl!</b>
                                    <br /><br />
                                    Kunde / Job er ikke fudnet og fakturaen er ikke flyttet.
                                    <br /><br />
                                    <a href="Javascript:history.back()"> << Tilbage</a>  
                                </div>


                                <%
                                end if

                                
                                end if
                                
                                %>



                           <%

                            case "kopierja"
                            %>
                            Du har kopieret faktura <%=faknr %>.
                            <%

                           
                            '***** Henter og kopiere faktura ***'
                            strSQL1 = "create table fakturaer_temp like fakturaer"
                            oConn.execute(strSQL1)

                            Response.Write "<br><br>Henter faktura."
                            Response.Flush

                            strSQL2 = "insert into fakturaer_temp select * from fakturaer where fid = "& fakid
                            oConn.execute(strSQL2)
    
                            strSQL3 = "UPDATE fakturaer_temp SET betalt = 0, erfakbetalt = 0, fak_laast = 0, overfort_erp = 0, faknr = '"& nytfaknr &"', fid = "& lastFakId &", dato = '"& varTjDatoUS_SQL &"', editor = '"& session("user") &"' WHERE fid = "& fakid
                            oConn.execute(strSQL3)
    
                            strSQL4 = "insert into fakturaer select * from fakturaer_temp where fid = "& lastFakId
                            oConn.execute(strSQL4)
    
                            strSQL5 = "DROP TABLE IF EXISTS fakturaer_temp"
                            oConn.execute(strSQL5)


                           call opdater_fakturanr_rakkefgl(opdFaknrSerie, intFaknumFindes, sqlfld, intFaknum)


                            Response.Write "<br><br>Behandler data."
                            Response.Flush
                              
                            '*******************************************************'
                            '****** FAKTURA_DET *****'
                            '***** Henter og kopiere ***'
                            '*******************************************************'
                            strSQL1 = "create table faktura_det_temp like faktura_det"
                            oConn.execute(strSQL1)

                                     '** looper igennem alle ids
                                    strSQL = "SELECT id FROM faktura_det WHERE fakid = "& fakid &" ORDER BY id"
                                    'Response.write strSQL
                                    'Response.Flush
                                    oRec.open strSQL, oConn, 3
                                    while not oRec.EOF 

                                   
                                               strSQL2 = "insert into faktura_det_temp select * from faktura_det where id = "& oRec("id")
                                               oConn.execute(strSQL2)

                                        
                                    oRec.movenext
                                    wend
                                    oRec.close
                          
    
                       

                                    
                                    '** looper igennem alle ids
                                    strSQL = "SELECT id FROM faktura_det_temp WHERE id <> 0 ORDER BY id"
                                    'Response.write strSQL
                                    'Response.Flush
                                    oRec.open strSQL, oConn, 3
                                    while not oRec.EOF 

                                              strSQL3 = "UPDATE faktura_det_temp SET fakid = "& lastFakId &" WHERE id = "& oRec("id")
                                               oConn.execute(strSQL3)

                                              strSQL6 = "UPDATE faktura_det_temp SET id = "& lastFak_det_Id &" WHERE id = "& oRec("id")
                                              oConn.execute(strSQL6)

                                          

                                    lastFak_det_Id = lastFak_det_Id + 1
                                    oRec.movenext
                                    wend
                                    oRec.close

                            strSQL4 = "insert into faktura_det select * from faktura_det_temp where fakid = "& lastFakId '+1 
                            oConn.execute(strSQL4)
 
                            strSQL5 = "DROP TABLE IF EXISTS faktura_det_temp"
                            oConn.execute(strSQL5)

                           
                              '*******************************************************'
                            '****** fak_mat_spec *****'
                            '***** Henter og kopiere ***'
                            '*******************************************************'
                            strSQL1 = "create table fak_mat_spec_temp like fak_mat_spec"
                            oConn.execute(strSQL1)

                                     '** looper igennem alle ids
                                    strSQL = "SELECT id FROM fak_mat_spec WHERE matfakid = "& fakid &" ORDER BY id"
                                    'Response.write strSQL
                                    'Response.Flush
                                    oRec.open strSQL, oConn, 3
                                    while not oRec.EOF 

                                   
                                               strSQL2 = "insert into fak_mat_spec_temp select * from fak_mat_spec where id = "& oRec("id")
                                               oConn.execute(strSQL2)

                                    oRec.movenext
                                    wend
                                    oRec.close
                          
    
                           

                                    
                                    '** looper igennem alle ids
                                    strSQL = "SELECT id FROM fak_mat_spec_temp WHERE id <> 0 ORDER BY id"
                                    'Response.write strSQL
                                    'Response.Flush
                                    oRec.open strSQL, oConn, 3
                                    while not oRec.EOF 

                                   
                                              strSQL3 = "UPDATE fak_mat_spec_temp SET matfakid = "& lastFakId &" WHERE id = "& oRec("id")
                                                oConn.execute(strSQL3)

                                              strSQL6 = "UPDATE fak_mat_spec_temp SET id = "& lastFak_mat_spec_Id &" WHERE id = "& oRec("id")
                                              oConn.execute(strSQL6)

                                             strSQL4 = "insert into fak_mat_spec select * from fak_mat_spec_temp where id = "& oRec("id")
                                             oConn.execute(strSQL4)

                                    lastFak_mat_spec_Id = lastFak_mat_spec_Id + 1
                                    oRec.movenext
                                    wend
                                    oRec.close

                         
   
    
                            strSQL5 = "DROP TABLE IF EXISTS fak_mat_spec_temp"
                            oConn.execute(strSQL5)



                            '*******************************************************'
                            '****** fak_med_spec *****'
                            '***** Henter og kopiere ***'
                            '*******************************************************'
                            strSQL1 = "create table fak_med_spec_temp like fak_med_spec"
                            oConn.execute(strSQL1)

                                     '** looper igennem alle ids
                                    strSQL = "SELECT id FROM fak_med_spec WHERE fakid = "& fakid &" ORDER BY id"
                                    'Response.write strSQL
                                    'Response.Flush
                                    oRec.open strSQL, oConn, 3
                                    while not oRec.EOF 

                                   
                                               strSQL2 = "insert into fak_med_spec_temp select * from fak_med_spec where id = "& oRec("id")
                                               oConn.execute(strSQL2)

                                    oRec.movenext
                                    wend
                                    oRec.close
                          
    
                           

                                    
                                    '** looper igennem alle ids
                                    strSQL = "SELECT id FROM fak_med_spec_temp WHERE id <> 0 ORDER BY id"
                                    'Response.write strSQL
                                    'Response.Flush
                                    oRec.open strSQL, oConn, 3
                                    while not oRec.EOF 

                                               strSQL3 = "UPDATE fak_med_spec_temp SET fakid = "& lastFakId &" WHERE id = "& oRec("id")
                                               oConn.execute(strSQL3)

                                   
                                              strSQL6 = "UPDATE fak_med_spec_temp SET id = "& lastFak_med_spec_Id &" WHERE id = "& oRec("id")
                                              oConn.execute(strSQL6)

                                             strSQL4 = "insert into fak_med_spec select * from fak_med_spec_temp where id = "& oRec("id")
                                             oConn.execute(strSQL4)

                                    lastFak_med_spec_Id = lastFak_med_spec_Id + 1
                                    oRec.movenext
                                    wend
                                    oRec.close

                         
   
    
                            strSQL5 = "DROP TABLE IF EXISTS fak_med_spec_temp"
                            oConn.execute(strSQL5)



                               %>
                                <br /><br />Den nye faktura er klar og har fået faktura nr.: <b><%=nytfaknr %></b> (kladde)
                               <%response.flush

                               case else



                               if cint(erfakbetalt) = 0 AND cint(level) = 1 then
                                   %> <br /><br />
                                   <b>Ønsker du at flytte denne faktura?</b>
                                  <br /><a href="erp_fak_kopier.asp?func=flyt&fakid=<%=fakid %>">Ja flyt faktura til nyt job/kunde</a> 
                                    <br /><br />
                               <%end if %>


                                <br /><br />
                                <b>Ønsker du at kopiere faktura <%=faknr %>?</b>
                                <br /><a href="erp_fak_kopier.asp?func=kopierja&fakid=<%=fakid %>">Ja kopier faktura</a> 
                                <br /><br />
                               <%
                               end select 'kopierja
                               %>

                           <br /><br /><a href="Javascript:window.close()">[Luk vindue]</a>
                        
                       </div>
                   </div>

   

  
      </div><!-- body -->
                </div> <!-- portlet -->
            </div> <!-- container -->
        </div><!-- wrapper -->
	
	<%end if %>
<!--#include file="../inc/regular/footer_inc.asp"-->
