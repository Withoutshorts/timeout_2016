

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<div class="wrapper">
    <div class="content">


<%
    if len(session("user")) = 0 then

	errortype = 5
	call showError(errortype)
	
	response.End
    end if
    

    func = request("func")

    
    select case func


    case "abo"

      '**** Abonner ****'
            

            'strConnect_admin = "driver={MySQL ODBC 3.51 Driver};server=62.182.173.226;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_admin;"
            'strConnect_admin = "driver={MySQL ODBC 3.51 Driver};server=localhost;Port=3306; uid=root;pwd=;database=timeout_admin;"
            'strConnect_admin = "driver={MySQL ODBC 3.51 Driver};server=localhost;Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_admin;" 
            strConnect_admin = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154;Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_admin;" 
            Set oConn_admin = Server.CreateObject("ADODB.Connection")
            Set oRec_admin = Server.CreateObject ("ADODB.Recordset")
            oConn_admin.open strConnect_admin
            
            editor = session("user")

            'Response.write request("FM_abo_mid") & "<br>"

            aMedid = split(request("FM_abo_mid"), ", ") 'session("mid")
            
            aMedTyp = request("FM_abo_mtyp")
            
            if trim(aMedTyp) = "-1" then 'kan kun abonnere p� enten eller
            aProgrp = request("FM_abo_progrp")
            else
            aProgrp = "-1"
            end if 

            lto = lto
            
            
            
            
            if request("FM_abo") = "1" then 'aboner
            
           for m = 0 to UBOUND(aMedid)
            

            'Response.write aMedid(m)

            dtNowSQL = year(now) &"/"& month(now) &"/"& day(now)
          
            call meStamdata(aMedid(m))

         

            'if meBrugergruppe = 3 then 'Admin
            if len(trim(request("FM_rapporttype"))) <> 0 then
            rapporttype = request("FM_rapporttype")
            
                    if cint(rapporttype) = 2 then
                    'Jobansvarlig rapport - Viser for alle medarbejdere
                    aMedTyp = 0
                    aProgrp = 0
                    end if

            else
            rapporttype = 0
            end if 
            
            'else
            'rapporttype = 1 'Viser kun tal for sig sel
            'end if

                '** Findes den i forvejen ==> g�r intet ***'


                '*** 20170104 �NDRET: Man skal kunne abonnere p� flere typer af rapporter ***
                'strSQLAbo = "SELECT medid FROM rapport_abo WHERE medid = "& aMedid(m) & " AND lto = '"& lto & "'"
                'oRec_admin.open strSQLAbo, oConn_admin, 3
                'if not oRec_admin.EOF then

                'else '**Insert

                 strSQLAbo = "INSERT INTO rapport_abo (dato, editor, lto, navn, email, medid, rapporttype, akttyper, medarbejdertyper, projektgrupper) "_
                 &" VALUES ('"& dtNowSQL &"','"& editor &"', '"& lto &"', '"& meNavn &"', '"& meEmail &"', "& aMedid(m) &", "& rapporttype &", '#0#', '"& aMedTyp &"', '"& aProgrp &"' )"
                 oConn_admin.execute(strSQLAbo)

                'end if
                'oRec_admin.close
          
            
            next


            else '**slet
               

            

            end if
            
            
          
            oConn_admin.close

            'Response.end


            Response.redirect "abonner.asp"

        case "abo_slet"


            'strConnect_admin = "driver={MySQL ODBC 3.51 Driver};server=62.182.173.226;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_admin;"
            'strConnect_admin = "driver={MySQL ODBC 3.51 Driver};server=localhost;Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_admin;" 
            strConnect_admin = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154;Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_admin;"
            Set oConn_admin = Server.CreateObject("ADODB.Connection")
            'Set oRec_admin = Server.CreateObject ("ADODB.Recordset")

            oConn_admin.open strConnect_admin
            
            aMedid = request("FM_abo_mid") 'session("mid")
            lto = lto


                strSQLAbo = "DELETE FROM rapport_abo WHERE medid = "& aMedid & " AND lto = '"& lto & "'"
               'Response.write strSQLAbo
                oConn_admin.execute(strSQLAbo)


         Response.redirect "abonner.asp"


        case "abo_tidspunkt"

            raptype = request("FM_raptype")
            klokkeslet = request("FM_klokkeslet")            
            navn = request("FM_navn")
            sendtype = request("FM_sendtype")

            ugedagkraevet = 0

            if sendtype <> 0 then
                ugedagkraevet = 0
            else
                ugedagkraevet = 1
            end if

            if ugedagkraevet = 1 then

                ugedag = request("FM_ugedag")
                select case ugedag
                case 0
                    errortype = 194
			        call showError(errortype)
			        Response.end
                case else
                    ' G�r Ingenting
                end select

            else
            
                ugedag = 0

            end if

            if klokkeslet = "00:00" then
                klokkeslet = "00:01"
            else
                klokkeslet = klokkeslet
            end if


            'response.Write " raptype " & raptype & " KLokkeslet " & klokkeslet & " Ugedag " & ugedag & " Navn " & navn & " Send type " & sendtype & "<br>"
            strConnect_admin = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154;Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_admin;"
            Set oConn_admin = Server.CreateObject("ADODB.Connection")
            'Set oRec_admin = Server.CreateObject ("ADODB.Recordset")
            

            strSQL = "INSERT INTO abonner_raptype_tidspunkt (lto, rapporttype, klokkeslet, ugedag, navn, type_send) "_
            &" VALUES ('"& lto &"', "& raptype &", '"& klokkeslet &"', "& ugedag &", '"& navn &"', "& sendtype &" )"
        
            'response.Write strSQL
            oConn_admin.execute(strSQL)

            Response.redirect "abonner.asp"

        case "abo_tidspunkt_slet"

            abo_tidsID = request("tidsID")

            strSQL = "DELETE FROM abonner_raptype_tidspunkt WHERE id =" & abo_tidsID
            'response.Write strSQL
            oConn.execute(strSQL)

            Response.redirect "abonner.asp"

   
case else 

    call menu_2014 

%>



        



        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u>Abonner</u></h3>

                <div class="portlet-body">
                    

                    <div class="row">
                        <div class="col-lg-9">&nbsp</div>
                        <div class="col-lg-3 pad-b10"><a class="btn btn-default btn-sm pull-right" href="../timereg_net/abonner_manuel_2017.aspx?lto=<%=lto %>&user=<%=session("user") %>" target="_blank">Send reports manuel now >> </a></div>
                    </div>

                    <br />

                    <!-- hvorn�r skal hvilke rapporttyper sendes ud -->
                    <form action="abonner.asp?func=abo_tidspunkt" method="post">
                    <div class="row">
                        <div class="col-lg-2"><b>Navn</b></div>
                        <div class="col-lg-2"><b>Rapporttype</b></div>
                        <div class="col-lg-1"></div>
                        <div class="col-lg-2"><b>Skal sendes</b></div>
                        <div class="col-lg-1"></div>
                        <div class="col-lg-2"><b>Ugedag</b></div>
                        <div class="col-lg-2"><b>Klokken</b></div>                        
                    </div>

                    <div class="row">

                        <div class="col-lg-2"><input type="text" class="form-control input-small" name="FM_navn" value="" placeholder="valgfrit" /></div>

                        <div class="col-lg-3">
                            <select class="form-control input-small" name="FM_raptype">
                                <option value="1" SELECTED><%=abonner_txt_014 %></option>
                                <option value="2"><%=abonner_txt_015 %> </option>
                                <option value="3"><%=abonner_txt_016 %> </option>
                                <option value="4">Ressource Forecast </option>
                            </select>
                        </div>

                        <div class="col-lg-3">
                            <select class="form-control input-small send_type" id="FM_sendtype" name="FM_sendtype">
                                <option value="0">Ugentlig</option>
                                <option value="1">F�rst dag i m�neden</option>
                                <option value="2">F�rste <u>arbejds</u>dag i m�nedenen</option>
                                <option value="3">Sidste dag i m�neden</option>
                            </select>
                        </div>

                        <div class="col-lg-2" style="text-align:center">
                            <select class="form-control input-small" id="FM_ugedag" name="FM_ugedag">
                                <option value="0">Ugedag</option>
                                <option value="1">Mandag</option>
                                <option value="2">Tirdag</option>
                                <option value="3">Onsdag</option>
                                <option value="4">Torsdag</option>
                                <option value="5">Fredag</option>
                                <option value="6">L�rdag</option>
                                <option value="7">S�ndag</option>
                            </select>
                            <span style="visibility:hidden;" id="ugedag_disabled"> - </span>
                        </div>

                        <div class="col-lg-2"><input type="time" class="form-control input-small" name="FM_klokkeslet" value="00:00" /></div>                        

                    </div>

                    <div class="row">                        
                        <div class="col-lg-10">&nbsp</div>
                        <div class="col-lg-2 pad-b10"><button type="submit" class="btn btn-success btn-sm pull-right"><b>Opret</b></button></div>
                    </div>

                    </form>

                    <br /><br />

                    <table class="table">
                        <thead>
                            <tr>
                                <th>Navn</th>
                                <th>Rapporttype</th>
                                <th>Sendes</th>
                                <th>Ugedag</th>
                                <th>Klokken</th>                                
                                <th>&nbsp</th>
                            </tr>
                        </thead>

                        <tbody>
                        <%

                             'strConnect_admin = "driver={MySQL ODBC 3.51 Driver};server=localhost;Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_admin;"
                            strConnect_admin = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154;Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_admin;" 
                            Set oConn_admin = Server.CreateObject("ADODB.Connection")
                            'Set oRec = Server.CreateObject ("ADODB.Recordset")
                            oConn_admin.open strConnect_admin
                           
                            strSQL = "SELECT id, rapporttype, klokkeslet, ugedag, navn, type_send FROM abonner_raptype_tidspunkt"
                            'response.Write strSQL & " her" 
                            'oRec.open strSQL, oConn, 3
                            oRec.open strSQL, oConn_admin, 3
                            while not oRec.EOF
                        %>

                            <tr>

                                <td><%=oRec("navn") %></td>

                                <td>
                                    <%
                                    select case oRec("rapporttype")
                                    case 1
                                    raptypeTxt = abonner_txt_014
                                    case 2
                                    raptypeTxt = abonner_txt_015
                                    case 3
                                    raptypeTxt = abonner_txt_016
                                    case 4
                                    raptypeTxt = "Ressource Forecast"
                                    end select

                                    %>

                                    <%=raptypeTxt %>
                                </td>

                                <td>
                                    <%
                                        select case oRec("type_send")
                                        case 0
                                        sendetypeTxt = "Ugentlig"
                                        case 1
                                        sendetypeTxt = "F�rste dag i m�neden"
                                        case 2
                                        sendetypeTxt = "F�rste arbejdsdag i m�neden"
                                        case 3
                                        sendetypeTxt = "F�rste dag i m�neden"
                                        end select

                                        response.Write sendetypeTxt
                                    %>
                                </td>

                                <td>
                                    <%
                                        select case oRec("ugedag")
                                        case 1
                                        ugedagTxt = "Mandag"
                                        case 2
                                        ugedagTxt = "Tirsdag"
                                        case 3
                                        ugedagTxt = "Onsdag"
                                        case 4
                                        ugedagTxt = "Torsdag"
                                        case 5
                                        ugedagTxt = "Fredag"
                                        case 6
                                        ugedaTxt = "L�rdag"
                                        case 7
                                        ugedagTxt = "S�ndag"
                                        case else
                                        ugedagTxt = "-"
                                        end select
                                    %>

                                    <%=ugedagTxt %>
                                </td>

                                <td>
                                    <%
                                        klokkenTxt = right(oRec("klokkeslet"), 8)
                                        klokkenTxt = left(klokkenTxt, 5)

                                        if klokkenTxt = "00:01" then
                                        klokkenTxt = "00:00"
                                        else
                                        klokkenTxt = klokkenTxt
                                        end if
                                    %>

                                    <%=klokkenTxt %>
                                </td>                                

                                <td><a href="abonner.asp?func=abo_tidspunkt_slet&tidsID=<%=oRec("id") %>"><span style="color:darkred;" class="fa fa-times"></span></a></td>
                            </tr>


                        <%
                            oRec.movenext
                            wend
                            oRec.close

                        %>

                        </tbody>


                    </table>




                    <br /><br /><br /><br />
                    <br /><br /><br /><br />


                    <!-- Tilfoj til abonner -->
                    <form method="post" action="abonner.asp?func=abo">

                    <div class="row">
                        <div class="col-lg-2"><b>A) <%=abonner_txt_010 %> </b></div>
                        <div class="col-lg-1"></div>
                        <div class="col-lg-2"><b>B) <%=abonner_txt_012 %> </b></div>
                        <div class="col-lg-1"></div>
                        <div class="col-lg-2"><b>C)<%=" "&abonner_txt_019 %>:</b></div>
                        <div class="col-lg-1"></div>
                        <div class="col-lg-2"><b>D) <%=abonner_txt_024 %>:</b></div>
                    </div>

                    <div class="row">
                        <!-- Modtagere -->
                        <div class="col-lg-3">
                            <%=abonner_txt_011 %><br />
                            <%strSQLm = "SELECT mnavn, mid, email FROM medarbejdere WHERE mansat = 1 AND email <> '' ORDER BY mnavn"

                            'if session("mid") = 1 then
                            'Response.write strSQLm
                            'end if
        
                            %>
                            <select class="form-control input-small" size="10" multiple name="FM_abo_mid">
                            <%
                            oRec2.open strSQLm, oConn, 3
                            While Not oRec2.EOF 

                            %>
                            <option value="<%=oRec2("mid")%>"><%=oRec2("mnavn") %></option>
                            <%
                            oRec2.movenext
                            wend
                            oRec2.close
                             %>
        
        
                            </select>

                            <input type="hidden" name="FM_abo" value="1" /> 

                        </div>

                        <!-- Rapportype -->
                        <div class="col-lg-3">
                            <%=abonner_txt_013 %>:<br />
                            <select class="form-control input-small" size="10" name="FM_rapporttype" id="FM_rapporttype">
                                <option value="1" SELECTED><%=abonner_txt_014 %></option>
                                <option value="2"><%=abonner_txt_015 %> </option>
                                <option value="3"><%=abonner_txt_016 %> </option>
                                <option value="4">Ressource Forecast </option>
                            </select>
                        </div>

                        <!-- Medarbtype -->
                        <div class="col-lg-3">
                            Enten <br />
                            <%strSQLm = "SELECT id, type FROM medarbejdertyper WHERE id <> 0 ORDER BY type"

                            'if session("mid") = 1 then
                            'Response.write strSQLm
                            'end if 
                            %>
                            <select class="form-control input-small" size="10" multiple name="FM_abo_mtyp" id="FM_abo_mtyp">
                                <option value="-2" SELECTED><%=abonner_txt_020 %></option>
                                  <option value="-1"><%=abonner_txt_021 %></option>
                                <option value="0"><%=abonner_txt_022 %></option>
                            <%
                            oRec2.open strSQLm, oConn, 3
                            While Not oRec2.EOF 

                            %>
                            <option value="<%=oRec2("id")%>"><%=oRec2("type") %> (id:<%=oRec2("id") %>)</option>
                            <%
                            oRec2.movenext
                            wend
                            oRec2.close
                             %>
                            </select>
                        </div>

                        <!-- Projektgruppe -->
                        <div class="col-lg-3">
                            Eller <br />
                            <%strSQLm = "SELECT id, navn FROM projektgrupper WHERE id <> 0 ORDER BY navn"

                            'if session("mid") = 1 then
                            'Response.write strSQLm
                            'end if
        
                            %>
                            <select class="form-control input-small" size="10" multiple name="FM_abo_progrp" id="FM_abo_progrp">
                                <option value="-1" SELECTED><%=abonner_txt_021 %></option>
                                <option value="0"><%=abonner_txt_022 %></option>
                            <%
                            oRec2.open strSQLm, oConn, 3
                            While Not oRec2.EOF 

                            %>
                            <option value="<%=oRec2("id")%>"><%=oRec2("navn") %> (id:<%=oRec2("id") %>)</option>
                            <%
                            oRec2.movenext
                            wend
                            oRec2.close
                             %>
                            </select>
                        </div>                       
                    </div>

                    <div class="row">
                        <div class="col-lg-10">&nbsp</div>
                        <div class="col-lg-2 pad-b10"><button type="submit" class="btn btn-success btn-sm pull-right"><b><%=medarb_txt_020 %></b></button></div>
                    </div>
                    </form>


                    <%  

                    if request.servervariables("PATH_TRANSLATED") <> "c:\www\timeout_xp\wwwroot\ver2_1\timereg\abonner.asp" then

                        'strConnect_admin = "driver={MySQL ODBC 3.51 Driver};server=localhost;Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_admin;"
                        strConnect_admin = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154;Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_admin;" 
                        Set oConn_admin = Server.CreateObject("ADODB.Connection")
                        Set oRec_admin = Server.CreateObject ("ADODB.Recordset")
                        oConn_admin.open strConnect_admin

                    %>



                    <table class="table">
                        <thead>
                            <tr>
                                <th><%=abonner_txt_029 %></th>
                                <th><%=abonner_txt_030 %></th>
                                <th><%=abonner_txt_031 %></th>
                                <th><%=abonner_txt_032 %></th>
                                <th>&nbsp</th>
                            </tr>
                        </thead>

                        <tbody>

                            <% '** subscribers ***'
                            strSQLAbo = "SELECT navn, medid, rapporttype, medarbejdertyper, projektgrupper FROM rapport_abo WHERE lto = '"& lto & "' ORDER BY navn "
                            oRec_admin.open strSQLAbo, oConn_admin, 3
                            while not oRec_admin.EOF
            
                            %>
                            <tr>
                                <td><%=oRec_admin("navn")%></td><td><%=oRec_admin("rapporttype") %></td>
                                <td>
                
                                    <%if oRec_admin("medarbejdertyper") <> "-1" then %>
                                    <%=oRec_admin("medarbejdertyper") %>
                                    <%end if %>
                                </td>        
                                <td>
                                    <%if oRec_admin("medarbejdertyper") = "-1" then %>
                                    <%=oRec_admin("projektgrupper") %>
                                    <%end if %>
                                </td>
                                <td><a href="abonner.asp?func=abo_slet&FM_abo_mid=<%=oRec_admin("medid") %>" class="slet"><span style="color:darkred;" class="fa fa-times"></span></a></td>
                            </tr>
                            <%
                            oRec_admin.movenext
                            wend

                            oRec_admin.close
                            oConn_admin.close 
                            %>
                            
                        </tbody>

                    </table>




                    <%end if %>



                </div>
            </div>
        </div>


        <script type="text/javascript">

            $(".send_type").change(function () {

                if (document.getElementById('FM_sendtype').value !== "0" ) {

                    document.getElementById("FM_ugedag").style.display = "none";
                    document.getElementById("ugedag_disabled").style.visibility = "";                   
                } else
                {
                    document.getElementById("FM_ugedag").style.display = "";
                    document.getElementById("ugedag_disabled").style.visibility = "hidden";
                }
            });

        </script>


    <%end select %>

</div>
</div>


<!--#include file="../inc/regular/footer_inc.asp"-->
