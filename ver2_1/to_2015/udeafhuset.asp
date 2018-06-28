

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


            'response.Write "HerHerHerHerHerHer Medarbejder " & medid & " fradato " & sqlFra & "tildato " & sqlTil & " fraklokken " & fraklokken & " til klokken " & tilklokken & " Heledagen " & heledagen

            strSQL = "INSERT INTO udeafhuset (medid, fradato, fratidspunkt, tildato, tiltidspunkt, heledagen, fra, til) VALUES ("& medid &", '"& sqlFra &"', '"& fraklokken &"', '"& sqlTil &"', '"& tilklokken &"', "& heledagen &", '"& fromDateTim &"', '"& toDateTime &"')"
            oConn.execute(strSQL)

            response.Redirect "udeafhuset.asp"

            
            case "ud_slet"

                udID = request("FM_udID")

                strSQL = "DELETE FROM udeafhuset WHERE id = "& udID 
                oConn.execute(strSQL)

                response.Redirect "udeafhuset.asp"
            

        %>

        <%case else %>
        
        <script src="js/datepicker.js" type="text/javascript"></script>
        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u>Ude af huset</u></h3>
                <div class="portlet-body">


                    <table class="table table-striped table-bordered">
                        <thead>
                            <tr>
                                <th>Medarbejder</th>
                                <th>Er ude af huset fra d.</th>                                
                                <th>Fra klokken</th>
                                <th>Til d.</th>
                                <th>Til klokken</th>
                                <th>Heledagen</th>
                            </tr>
                        </thead>

                        <tbody>

                            <form action="udeafhuset.asp?func=ud_opr" method="post">
                                <tr>
                                    <td>
                                        <select class="form-control input-small" name="FM_medid">
                                            <option value="0">Ingen</option>
                                        
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
                                    <td style="text-align:center"><button type="submit" class="btn btn-default btn-sm"><b>Opret</b></button></td>

                                </tr>
                            </form>

                            <%
                                x = 0
                                strSQL = "SELECT id, medid, m.mnavn as mnavn, fradato, fratidspunkt, tildato, tiltidspunkt, heledagen FROM udeafhuset as u LEFT JOIN medarbejdere as m ON m.mid = u.medid"
                                'response.Write strSQl
                                oRec.open strSQL, oConn, 3
                                while not oRec.EOF
                                x = x + 1
                            %>
                                <tr>
                                    <td><%=oRec("mnavn") %></td>

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

                                    <td style="text-align:center"><a href="udeafhuset.asp?func=ud_slet&FM_udID=<%=oRec("id") %>"><span style="color:darkred;" class="fa fa-times"></span></a></td>

                                </tr>
                            <%
                                oRec.movenext
                                wend
                                oRec.close
                            %>

                            <%if x = 0 then %>

                                <tr>
                                    <td colspan="20" style="text-align:center;">Ingen ude af huset registreringer</td>
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