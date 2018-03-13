


<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->





       
        
        <%
           

           

            func = request("func")

          

         select case func

        case "db"
           
                    nh_editor_date = year(now) &"/"& month(now) &"/"& day(now)
                    

                    'response.write "IDs:"& request("FM_id") & "<br>"
                    'Response.write request("FM_open") & " CC:  "& request("FM_cc")
                    'Response.write "Dates: "& request("FM_date")
                    'response.flush
            
                    nh_id = split(request("FM_id"), ", ")
                    nh_id_row = split(request("FM_id_row"), ", ")
            
                    'nh_open = split(request("FM_open"), ", ")
                    'nh_country = split(request("FM_cc"), ", ")
                    nh_date = split(request("FM_date"), ", ")
                    nh_date_year = split(request("FM_date_year"), ", ")
                    'nh_name = split(request("FM_name"), ", ")


                    'for y = 0 TO UBOUND(nh_id)
            
            
                      

                            for i = 0 TO UBOUND(nh_id)
                    
                            sqlDate = nh_date(i) &"-"& nh_date_year(i) 
                            sqlDate = year(sqlDate) &"-"& month(sqlDate) &"-"& day(sqlDate)

                                if isDate(sqlDate) = true then

                                nh_name = request("FM_name_"& nh_id_row(i))
                                nh_name = replace(nh_name, ",", "")
                                nh_name = replace(nh_name, "'", "")
                                nh_name = trim(nh_name)
                                nh_open = request("FM_open_"& nh_id_row(i))
                                nh_open = replace(nh_open, ",", "")
                                nh_open = replace(nh_open, " ", "")
                                nh_country = request("FM_cc_" & nh_id_row(i))
                                nh_country = replace(nh_country, ",", "")
                                nh_country = replace(nh_country, " ", "")

                                nh_projgrp = request("FM_projgrp_"& nh_id_row(i))
                                nh_projgrp = trim(nh_projgrp)


                                sortorder = request("FM_sortorder_"& nh_id_row(i))

                                strSQL6 = "UPDATE national_holidays SET nh_country = '"& nh_country &"', nh_name = '"& nh_name &"', nh_duration = 1, nh_date = '"& sqlDate &"', "_
                                &"nh_editor_date = '"& nh_editor_date &"', nh_open = "& nh_open &", nh_sortorder = "& sortorder &", nh_projgrp = '"& nh_projgrp &"' WHERE nh_id = "& nh_id(i)
                                'response.write "i"& i &" SQL: "& strSQL6 & "<br>"
                                'response.flush
                                oConn.execute(strSQL6)
                   
                                end if

                            next

                        

                    'next


                    '*** Tilføjer NY

                    nh_id = request("FM_id_0")
                    nh_country = request("FM_cc_0")
                    nh_country = replace(nh_country, " ", "")

                    nh_date = split(request("FM_date_0"), ", ")        
                    nh_name = request("FM_name_0")
                    nh_name = replace(nh_name, ",", "")
                    nh_name = replace(nh_name, "'", "")
                    nh_name = trim(nh_name)
                   
                    nh_open = request("FM_open_0")
                    nh_open = replace(nh_open, ",", "")
                    nh_open = replace(nh_open, " ", "")

                    nh_date_year = split(request("FM_date_year_0"), ", ")

                    nh_projgrp = request("FM_projgrp_0")
                    nh_projgrp = trim(nh_projgrp)
                    

                    for y = 0 TO UBOUND(nh_date)
            
            
                        sqlDate = nh_date(y) &"-"& nh_date_year(y) 
                        sqlDate = year(sqlDate) &"-"& month(sqlDate) &"-"& day(sqlDate)

                        if isDate(sqlDate) = true AND len(trim(nh_name)) <> 0 then

                        strSQLins = "INSERT INTO national_holidays (nh_country, nh_name, nh_duration, nh_date, nh_editor_date, nh_open, nh_projgrp) "_
                        &" VALUES ('"& nh_country &"', '"& nh_name &"', 1, '"& sqlDate &"', '"& nh_editor_date &"', "& nh_open &", '"& nh_projgrp &"')"

                        'response.write strSQLins & "<br>"
                        'response.flush

                        oConn.execute(strSQLins)

                        end if


                    next


                    Response.redirect "national_holidays.asp"

        case "del"

                id = request("id")


                strSQLdel = "SELECT nh_name, nh_country FROM national_holidays WHERE nh_id = "& id
                oRec.open strSQLdel, oConn, 3
                while not oRec.EOF 

                    strSQLdel_serie = "DELETE FROM national_holidays WHERE nh_name = '"& oRec("nh_name") & "' AND nh_country = '"& oRec("nh_country") & "'"
                    oConn.execute(strSQLdel_serie)    


                oRec.movenext
                wend
                oRec.close


                 Response.redirect "national_holidays.asp"

        case else


       

        call licensStartDato()
        licensstdato = licensstdato
       
        %>

        <div class="container" style="width:100%; height:100%">
            <div class="portlet">
                   
                  <h3 class="portlet-title"><u>National holidays - Scheme</u> </h3>
                  

                <div class="portlet-body">
                   

                        <div class="row" style="padding-left:100px;">
                            
                            
                            <form action="national_holidays.asp?func=db" method="post">
                            <table style="width:90%" class="table table-striped">
                                <thead>
                                    <tr>
                                    <td>Country</td>
                                    <td>Order</td>
                                    <td>Name</td>
                                    <td>Projgrp <br />(1,14, ect. minus if exclude)</td>
                                    <td>Holiday</td>
                                    <%
                                        stAar = year(now)
                                        for y = -1 to 14
                                        thisYear = (stAar + y)
                                        %><td style="text-align:center;"><%=thisYear%></td><%
                                        
                                    next%>
                                    
                                    
                                    </tr>
                                </thead>
                                
                             <%

                            row = 1
                            x = 0
                            stAarSQLkri = dateAdd("yyyy", -5, now)
                            stAarSQLkriYear = year(stAarSQLkri)

                            lastname = ""
                            strSQL6 = "SELECT nh_id, nh_country, nh_name, nh_duration, nh_date, nh_editor_date, nh_open, nh_sortorder, nh_projgrp FROM national_holidays WHERE nh_id <> 0 AND YEAR(nh_date) >= "& stAarSQLkriYear &" ORDER BY nh_country, nh_sortorder, nh_name, nh_date"
                            oRec.open strSQL6, oConn, 3
                            while not oRec.EOF

                                 
                                 if lastname <> oRec("nh_country") &"_"& oRec("nh_name") then

                                    if lastname <> oRec("nh_country") &"_"& oRec("nh_name") AND x > 0 then 
                                    %>

                                    <td><a href="national_holidays.asp?func=del&id=<%=lastid %>"><span style="color:red;">X</span></a></td>
                                    </tr>
                                    <%
                                    row = row + 1  
                                    end if


                                 %>
                                 <tr>
                                     <!--<input type="hidden" name="FM_id" value="<%=oRec("nh_id")%>" />-->
                                    <td valign="top"><input type="text" name="FM_cc_<%=row%>" value="<%=oRec("nh_country")%>" class="form-control input-small" /></td>
                                    <td style="width:40px;"><input type="text" value="<%=oRec("nh_sortorder") %>" name="FM_sortorder_<%=row%>" class="form-control input-small" /></td>
                                    <td style="width:140px;"><input type="text" value="<%=oRec("nh_name") %>" name="FM_name_<%=row%>" style="width:140px;" class="form-control input-small" /></td>
                                    <td style="width:60px;"><input type="text" value="<%=oRec("nh_projgrp") %>" name="FM_projgrp_<%=row%>" class="form-control input-small" /></td>
                                    <td><select name="FM_open_<%=row%>" class="form-control input-small">
                                        
                                        <%if cint(oRec("nh_open")) = 1 then
                                            opSEL1 = "SELECTED"
                                            opSEL0 = ""
                                          else
                                            opSEL1 = ""
                                            opSEL0 = "SELECTED"
                                        end if
                                            
                                            %>
                                        
                                        <option value="1" <%=opSEL1 %>>Yes</option>
                                        <option value="0" <%=opSEL0 %>>No</option>
                                        </select></td>
                                    
                                    

                                     <%

                                     end if

                                        'stAar = year(now)
                                        'for y = -5 to 10
                                        thisYear = year(oRec("nh_date")) '(stAar + y)
                                        %><td><input type="text"  name="FM_date" value="<%=left(formatdatetime(oRec("nh_date"), 2), 5) %>" class="form-control input-small" style="width:50px;"  />
                                            <!--<br /><%=oRec("nh_date")%>--> 
                                            <input type="hidden" name="FM_date_year" value="<%=thisYear %>" />
                                            <input type="hidden" name="FM_id" value="<%=oRec("nh_id")%>" />
                                            <input type="hidden" name="FM_id_row" value="<%=row%>" />
                                          </td><%
                                        
                                       'next
                                              
                                   

                                lastname = oRec("nh_country") &"_"& oRec("nh_name")
                                lastid = oRec("nh_id")
                                
                            
                            x = x + 1
                            oRec.movenext
                            wend 
                            oRec.close


                                  
                                    %>

                                    <td><a href="national_holidays.asp?func=del&id=<%=lastid %>"><span style="color:red;">X</span></a></td>
                                    </tr>
                                    <%
                                   
                                        %>

                                    <tr>
                                     <input type="hidden" name="FM_id_0" value="" />
                                    <td valign="top"><input type="text" name="FM_cc_0" value="" class="form-control input-small" /></td>
                                    <td>&nbsp;</td>
                                    <td><input type="text" value="" name="FM_name_0" style="width:140px;" class="form-control input-small" /></td>
                                    <td><input type="text" value="" name="FM_projgrp_0" class="form-control input-small" /></td>
                                    <td><select name="FM_open_0" class="form-control input-small">
                                        
                                        
                                        
                                        <option value="1">Yes</option>
                                        <option value="0">No</option>
                                        </select></td>
                                    
                                    

                                     <%
                                        stAar = year(now)
                                        for y = -1 to 14
                                        thisYear = (stAar + y)
                                        %><td><input type="text" name="FM_date_0" value="" class="form-control input-small" style="width:50px;"  />
                                            <input type="hidden" name="FM_date_year_0" value="<%=thisYear %>" />
                                          </td><%
                                        
                                         next%>

                                    <td>&nbsp;</td>
                                    </tr>

                            </table><br />


                                

                          
                        </div>
                        <div class="row">
                            <div class="col-lg-10" style="text-align:right;">
                            <input type="submit" class="btn btn-small btn-success" value="Opdater >>">
                                </div>
                    </form>
                            </div>
                                        </div>
            </div>
        </div>
<%end select %>

<!--#include file="../inc/regular/footer_inc.asp"-->