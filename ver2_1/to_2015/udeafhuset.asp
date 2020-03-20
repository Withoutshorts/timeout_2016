

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

            call menu_2014 

            func = request("func")


            if len(trim(request("FM_datoFra"))) <> 0 then
                fradato = request("FM_datoFra")
            else
                fradato = day(now) &"-"& month(now) &"-"& year(now) 
            end if

            if len(trim(request("FM_datoTil"))) <> 0 then
                tildato = request("FM_datoTil") 
            else                
                tildato = day(now) &"-"& month(now) &"-"& year(now)
            end if


            select case func
            
            case "ud_opr"

         
            if  request("FM_medid") <> 0 then
                medid = request("FM_medid")
            else      
                errortype = 197
                call showError(errortype)
                response.End
            end if

            sqlFra = year(fradato) &"-"& month(fradato) &"-"& day(fradato)
            sqlTil = year(tildato) &"-"& month(tildato) &"-"& day(tildato)

            fraklokken = request("FM_fratidspunkt_hour") &":"& request("FM_fratidspunkt_minute") 
            tilklokken = request("FM_tiltidspunkt_hour") &":"& request("FM_tiltidspunkt_minute")

            dateTimeFra = cdate(sqlFra)
            dateTimeTil = cdate(sqlTil)

           if cdate(dateTimeFra) > cdate(dateTimeTil) then
                errortype = 193
                call showError(errortype)
                response.End
            end if

            'response.Write "dateTIme " & dateTimeFra & " TIl " & dateTimeTil & "<br>"

            if len(trim(request("FM_heledagen"))) <> 0 then
                heledagen = 1
            else
                heledagen = 0
            end if

            if heledagen <> 1 then
                if (cdate(dateTimeFra) = cdate(dateTimeTil)) AND (fraklokken > tilklokken) then
                    errortype = 198
                    call showError(errortype)
                    response.End
                end if

                if len(trim(request("FM_fratidspunkt_hour"))) = 0 OR len(trim(request("FM_fratidspunkt_minute"))) = 0 OR len(trim(request("FM_tiltidspunkt_hour"))) = 0 OR len(trim(request("FM_tiltidspunkt_minute"))) = 0 then
                    errortype = 199
                    call showError(errortype)
                    response.End
                end if
            end if

            if (isdate(fraklokken) <> true OR isdate(tilklokken) <> true) AND heledagen <> 1 then
                errortype = 200
                call showError(errortype)
                response.End
            end if

            if heledagen = 1 then
                fraklokken = "00:01"
                tilklokken = "23:59"
            end if

            fromDateTim = sqlFra &" "& fraklokken
            toDateTime = sqlTil &" "& tilklokken

            if len(trim(request("FM_udetype"))) <> 0 then
                udetype = request("FM_udetype")
            else
                udetype = 0
            end if

            'response.Write "HerHerHerHerHerHer Medarbejder " & medid & " fradato " & sqlFra & "tildato " & sqlTil & " fraklokken " & fraklokken & " til klokken " & tilklokken & " Heledagen " & heledagen

            strSQL = "INSERT INTO udeafhuset (medid, fradato, fratidspunkt, tildato, tiltidspunkt, heledagen, fra, til, udeafhusettype) VALUES ("& medid &", '"& sqlFra &"', '"& fraklokken &"', '"& sqlTil &"', '"& tilklokken &"', "& heledagen &", '"& fromDateTim &"', '"& toDateTime &"', "& udetype &")"
            oConn.execute(strSQL)


            'Tjekekr om der skal timer ind i dag
            sqlNow = year(now) &"-"& month(now) &"-"& day(now)
            if sqlFra = sqlNow AND udetype <> 0 then
                
                aktid = 0
                select case cint(udetype)
                    case 1 'Forretnignsrejse
                        aktid = 33            
                    case 2 'Ferie
                        aktid = 13     
                    case 3 'Arbejde hjemme
                        aktid = 32
                    case 4 'Syg
                        aktid = 30
                end select

                sqlNow = year(now) &"-"& month(now) &"-"& day(now) 
                call normtimerPer(medid, sqlNow, 0, 0)
                antal = ntimper

                lastUid = 0
                strSQL = "SELECT id FROM udeafhuset ORDER BY id DESC" 
                oRec3.open strSQL, oConn, 3
                if not oRec3.EOF then
                    lastUid = oRec3("id")
                end if
                oRec3.close

                timerkom = ""
                koregnr = ""
                destination = ""
                usebopal = 0

                call indlasTimerTfaktimAktid(lto, medid, antal, 0, aktid, 0, sqlNow, lastUid, timerkom, koregnr, destination, usebopal)
            end if




            response.Redirect "udeafhuset.asp"

            
            case "ud_slet"

                udID = request("FM_udID")
                udeafhusettype = request("udeafhusettype")
                medid = request("medid")

                if udeafhusettype = 0 then 
                    strSQL = "DELETE FROM udeafhuset WHERE id = "& udID 
                    oConn.execute(strSQL)

                    response.Redirect "udeafhuset.asp"
                else 'Her skal der også slettes timer derfor blvier der spurgt inden om man er sikker på man vil slette
                          
                    udeafhusettype = request("udeafhusettype")

                    %>
                    <div class="container">
                        <h3 style="text-align:center;"><%=udeafhuset_txt_001 %></h3>
                        <div class="portlet-body">
                            <%

                            select case udeafhusettype
                                case "1", cint(1)
                                    aktivitetsid = 33
                                case "2", cint(2)
                                    aktivitetsid = 13
                                case "3", cint(3)
                                    aktivitetsid = 32
                                case "4", cint(4)
                                    aktivitetsid = 30
                            end select

                            deleteTimerTxt = ""
                            strSQL = "SELECT timer, tmnavn, taktivitetnavn, tdato FROM timer WHERE tmnr = "& medid & " AND taktivitetid = "& aktivitetsid & " AND extsysid = "& udID
                            'response.Write strSQL
                            oRec.open strSQL, oConn, 3
                            while not oRec.EOF
                                deleteTimerTxt = deleteTimerTxt & "<tr><td>"&oRec("tmnavn")&"</td>"
                                deleteTimerTxt = deleteTimerTxt & "<td>"&oRec("taktivitetnavn")&"</td>"
                                deleteTimerTxt = deleteTimerTxt & "<td>"&oRec("timer")&"</td>"
                                deleteTimerTxt = deleteTimerTxt & "<td>"&oRec("tdato")&"</td>"
                                deleteTimerTxt = deleteTimerTxt & "<tr>" 
                            oRec.movenext
                            wend
                            oRec.close
                            %>

                            <%if deleteTimerTxt <> "" then %>
                            <h4 class="col-lg-12" style="text-align:center;"><%=udeafhuset_txt_002 %></h4>
                            <div class="row">        
                                <div class="col-lg-4"></div>
                                <div class="col-lg-4">
                                    <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                                        <tbody>
                                            <%=deleteTimerTxt %>
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                            <%end if %>
                            <br /><br />
                            <div class="row">
                                <div class="col-lg-12" style="text-align:center;">
                                    <a href="udeafhuset.asp?func=sletmedtimer&medid=<%=medid %>&FM_udID=<%=udID %>&udeafhusettype=<%=udeafhusettype %>" class="btn btn-danger btn-sm" style="width:50px;"><b><%=udeafhuset_txt_003 %></b></a>
                                    &nbsp
                                    <a href="udeafhuset.asp" class="btn btn-default btn-sm" style="width:50px;"><b><%=udeafhuset_txt_004 %></b></a>
                                </div>
                            </div>

                        </div>
                        <br /><br /><br /><br />
                    </div>
                    <%

                end if ' udeafhuesttype
            
            case "sletmedtimer"

                    medid = request("medid")
                    udID = request("FM_udID")

                    if udID <> 0 then
                        udeafhusettype = request("udeafhusettype")

                        select case udeafhusettype
                            case "1", cint(1)
                                aktivitetsid = 33
                            case "2", cint(2)
                                aktivitetsid = 13
                            case "3", cint(3)
                                aktivitetsid = 32
                            case "4", cint(4)
                                aktivitetsid = 30
                        end select

                        strSQL = "DELETE FROM timer WHERE tmnr = "& medid & " AND taktivitetid = "& aktivitetsid & " AND extsysid = "& udID 
                        oConn.execute(strSQL)

                        strSQL = "DELETE FROM udeafhuset WHERE id = "& udID 
                        oConn.execute(strSQL)
                    end if

                    response.Redirect "udeafhuset.asp"
        %>

        <%case else %>
        
        <script src="js/datepicker.js" type="text/javascript"></script>
        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u><%=udeafhuset_txt_005 %></u></h3>
                <div class="portlet-body">


                    <table class="table table-striped table-bordered">
                        <thead>
                            <tr>
                                <th><%=udeafhuset_txt_016 %></th>
                                <%if lto = "cool" OR lto = "hidalgo" OR lto = "kongeaa" then %>
                                <th><%=udeafhuset_txt_006 %></th>
                                <%end if %>
                                <th><%=udeafhuset_txt_007 %></th>                                
                                <th><%=udeafhuset_txt_008 %></th>
                                <th><%=udeafhuset_txt_009 %></th>
                                <th><%=udeafhuset_txt_010 %></th>
                                <th><%=udeafhuset_txt_011 %></th>
                            </tr>
                        </thead>

                        <tbody>

                            <form action="udeafhuset.asp?func=ud_opr" method="post">
                                <tr>
                                    <td>
                                        <select class="form-control input-small" name="FM_medid">
                                            <option value="0"><%=udeafhuset_txt_012 %></option>
                                        
                                            <%
                                                strSQL = "SELECT Mid, mnavn FROM medarbejdere WHERE mansat = 1 ORDER BY mnavn"
                                                oRec.open strSQL, oConn, 3 
                                                while not oRec.EOF 
                                            %>
                                                <option value="<%=oRec("Mid") %>"><%=oRec("mnavn") %></option>
                                            <%
                                                oRec.movenext
                                                wend
                                                oRec.close
                                            %>
                                        </select>
                                    </td>

                                    <%if lto = "cool" OR lto = "hidalgo" OR lto = "kongeaa" then %>
                                    <td>
                                        <select class="form-control input-small" name="FM_udetype">
                                            <option value="0"><%=udeafhuset_txt_021 %></option>
                                            <option value="1"><%=udeafhuset_txt_017 %></option>
                                            <option value="2"><%=udeafhuset_txt_018 %></option>
                                            <option value="3"><%=udeafhuset_txt_019 %></option>
                                            <option value="4"><%=udeafhuset_txt_020 %></option>
                                        </select>
                                    </td>
                                    <%end if %>

                                    <td>
                                        <div class='input-group date'>
                                            <input type="text" style="width:100%;" class="form-control input-small" name="FM_datoFra" value="<%=fradato %>" placeholder="dd-mm-yyyy" />
                                            <span class="input-group-addon input-small"><span class="fa fa-calendar"></span></span>
                                        </div>
                                    </td>

                                    <td style="width:100px; text-align:center;">
                                        <input type="text" class="form-control input-small" name="FM_fratidspunkt_hour" style="display:inline-block; width:40%"/>
                                        <span>:</span>
                                        <input type="text" class="form-control input-small" name="FM_fratidspunkt_minute" style="display:inline-block; width:40%" />
                                    </td>

                                    <td>
                                        <div class='input-group date'>
                                            <input type="text" style="width:100%;" class="form-control input-small" name="FM_datoTil" value="<%=tildato %>" placeholder="dd-mm-yyyy" />
                                            <span class="input-group-addon input-small"><span class="fa fa-calendar"></span></span>
                                        </div>
                                    </td>
                                    
                                    <td style="width:100px; text-align:center;">
                                        <input type="text" class="form-control input-small" name="FM_tiltidspunkt_hour" style="display:inline-block; width:40%" />
                                        <span>:</span>
                                        <input type="text" class="form-control input-small" name="FM_tiltidspunkt_minute" style="display:inline-block; width:40%" />
                                    </td>

                                    <td style="text-align:center"><input type="checkbox" name="FM_heledagen" value="1" /></td>

                                    <!-- Opret knap -->
                                    <td style="text-align:center"><button type="submit" class="btn btn-default btn-sm"><b><%=udeafhuset_txt_013 %></b></button></td>

                                </tr>
                            </form>

                            <%
                                x = 0
                                strSQL = "SELECT id, medid, m.mnavn as mnavn, fradato, fratidspunkt, tildato, tiltidspunkt, heledagen, udeafhusettype FROM udeafhuset as u LEFT JOIN medarbejdere as m ON m.mid = u.medid ORDER BY id DESC"
                                'response.Write strSQl
                                oRec.open strSQL, oConn, 3
                                while not oRec.EOF
                                x = x + 1
                            %>
                                <tr>
                                    <td><%=oRec("mnavn") %></td>

                                    <%if lto = "cool" OR lto = "hidalgo" OR lto = "kongeaa" then %>
                                        <%
                                        select case oRec("udeafhusettype") 
                                            case cint(1)
                                            typeTxt = udeafhuset_txt_017
                                            case cint(2)
                                            typeTxt = udeafhuset_txt_018
                                            case cint(3)
                                            typeTxt = udeafhuset_txt_019
                                            case cint(4)
                                            typeTxt = udeafhuset_txt_020
                                            case else
                                            typeTxt = udeafhuset_txt_021
                                        end select
                                        %>
                                        <td><%=typeTxt %></td>
                                    <%end if %>

                                    <td>                              
                                        <div class='input-group date'>
                                            <input type="text" style="width:100%;" class="form-control input-small" value="<%=day(oRec("fradato"))&"-"&month(oRec("fradato"))&"-"&year(oRec("fradato")) %>" readonly />
                                            <span class="input-group-addon input-small"><span class="fa fa-calendar"></span></span>
                                        </div>
                                    </td>

                                    <td style="text-align:center;">
                                        <%
                                            fraklokken = right(oRec("fratidspunkt"), 8)
                                            fraklokken = left(fraklokken, 5)
                                        %>
                                       <!-- <input type="text" readonly value="<%=fraklokken %>" class="form-control input-small" /> -->
                                        <%=fraklokken %>

                                    </td>

                                    <td>
                                        <div class='input-group date'>
                                            <input type="text" style="width:100%;" class="form-control input-small" value="<%=day(oRec("tildato"))&"-"&month(oRec("tildato"))&"-"&year(oRec("tildato")) %>" readonly />
                                            <span class="input-group-addon input-small"><span class="fa fa-calendar"></span></span>
                                        </div>
                                    </td>

                                    <td style="text-align:center">
                                        <%
                                            tilklokken = right(oRec("tiltidspunkt"), 8)
                                            tilklokken = left(tilklokken, 5)
                                        %>
                                       <!-- <input type="text" readonly value="<%=tilklokken %>" class="form-control input-small" /> -->
                                        <%=tilklokken %>
                                    </td>

                                    <td style="text-align:center;">
                                        <%
                                            if oRec("heledagen") <> 0 then
                                                response.Write  "<span class=""fa fa-check""></span>"
                                            else
                                                response.Write "-"
                                            end if
                                        %>
                                    </td>

                                    <td style="text-align:center"><a href="udeafhuset.asp?func=ud_slet&FM_udID=<%=oRec("id") %>&udeafhusettype=<%=oRec("udeafhusettype") %>&medid=<%=oRec("medid") %>"><span style="color:darkred;" class="fa fa-times"></span></a></td>

                                </tr>
                            <%
                                oRec.movenext
                                wend
                                oRec.close
                            %>

                            <%if x = 0 then %>

                                <tr>
                                    <td colspan="20" style="text-align:center;"><%=udeafhuset_txt_014 %></td>
                                </tr>

                            <%end if %>

                        </tbody>

                    </table>


                </div>
            </div>
        </div>


        <%end select 'func %>











</div></div>

<!--#include file="../inc/regular/footer_inc.asp"-->