	 
<%sub timeforbrug4md %>
<table cellpadding="0" cellspacing="0" border="0">
        <tr>
            <%
            'call akttyper2009(2)
            
            for dt = 0 TO 3
             if dt <> 0 then
             dtnowHigh = dateadd("m", 1, dtnowHigh)
             else
             dtnowHigh = dateadd("m", -3, slutDatovlgt)
             end if

                select case month(dtnowHigh)
             case 1,3,5,7,8,10,12
                eday = 31
                case 2
                    select case right(year(dtnowHigh), 2)
                    case "00", "04", "08", "12", "16", "20", "24", "28", "30", "34", "38", "42", "46"
                    eday = 29
                    case else
                    eday = 28
                    end select
             case else
                eday = 30
             end select

            

            %>
            <td align=center style="font-size:7px; line-height:9px;">1 - <%=eday&"<br>"&left(monthname(month(dtnowHigh)), 3) %></td>
            <%next %>
         </tr>

         <tr bgcolor="#FFFFFF">
            <%for dt = 0 TO 3
             if dt <> 0 then
             dtnowHigh = dateadd("m", 1, dtnowHigh)
             else
             dtnowHigh = dateadd("m", -3, slutDatovlgt)
             end if

             select case month(dtnowHigh)
             case 1,3,5,7,8,10,12
                eday = 31
                case 2
                    select case right(year(dtnowHigh), 2)
                    case "00", "04", "08", "12", "16", "20", "24", "28", "30", "34", "38", "42", "46"
                    eday = 29
                    case else
                    eday = 28
                    end select
             case else
                eday = 30
             end select

             dtnowHighSQL = year(dtnowHigh) &"/"& month(dtnowHigh) &"/"& eday
             dtnowLowSQL = year(dtnowHigh) &"/"& month(dtnowHigh) &"/1"

               hgt = 0
               timerThis = 0

             strSQLjl = "SELECT sum(timer) AS timer FROM timer WHERE tjobnr = '"& oRec("jobnr") &"' AND tdato BETWEEN '"& dtnowLowSQL &"' AND '"& dtnowHighSQL &"'"_
             &" AND ("& aty_sql_realhours &") GROUP BY tjobnr"
             

             'if session("mid") = 1 then
            ' Response.Write strSQLjl
             'Response.flush
             'end if

             oRec2.open strSQLjl, oConn, 3
             if not oRec2.EOF then

             hgt = formatnumber((oRec2("timer")/timerforbrugt)*100, 0)
             timerThis = oRec2("timer")

             end if
             oRec2.close

            if cint(visSimpel) = 1 then
              height100 = 20
                hgtDivider = 5
            else
                  height100 = 50
                hgtDivider = 2
            end if

             if hgt > 100 then
             hgt = height100
             else
             hgt = formatnumber(hgt/hgtDivider,0) 
             end if
            %>
            <td class=lille align=right height=<%=height100 %> valign=bottom style="border-right:1px #CCCCCC solid;">
            <%if hgt <> 0 then 
            bdThis = 1
            timerThis = formatnumber(timerThis,0)
            else
            timerThis = ""
            bdThis = 0
            
            end if 
            
            bgThis = "#D6dff5"

            if hgt = 0 then
            bgThis = "#ffffff"
            end if

           
          
            %>
            <div style="height:<%=hgt%>px; width:25px; border-top:<%=bdThis%>px #CCCCCC solid; font-size:8px; padding:0px 0px 0px 0px; background-color:<%=bgThis%>;"><%=timerThis%></div> 
            
            </td>
            <%next %>
         </tr>
	    
        </table>
<%end sub %>