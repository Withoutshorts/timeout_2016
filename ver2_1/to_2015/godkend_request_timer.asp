


<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->





       
        
        <%
           

           

            func = request("func")

            if len(trim(request("usemrn"))) <> 0 then
            usemrn = request("usemrn")
            else
            usemrn = 0
            end if
            
            call meStamdata(usemrn)


         select case func

        case "db"
           

                '******* Tjekker medarbejderen ********'
                

                    '******* Tjekker logindtid ********'
                    strSQL = "SELECT id, logud FROM login_historik WHERE mid = "& medid &" AND stempelurindstilling > 0 ORDER BY id DESC"
                    oRec.open strSQL, oConn, 3
                    if not oRec.EOF then
                    logud = oRec("logud")
                    id = oRec("id")
                    end if
                    oRec.close
                   

        case else


        call akttyper2009(2)
        call licensStartDato()
        licensstdato = licensstdato
        meAnsatDato = meAnsatDato

        if cDate(meAnsatDato) > cDate(licensstdato) then
        useStDatoKri = meAnsatDato 
        else
        useStDatoKri = licensstdato
        end if

        call lonKorsel_lukketPerPrDato(now)
        if lonKorsel_lukketPerDt > useStDatoKri then
        useStDatoKri = dateAdd("d", 1, day(lonKorsel_lukketPerDt) &"/"& month(lonKorsel_lukketPerDt) &"/"& year(lonKorsel_lukketPerDt))
        else
        useStDatoKri = useStDatoKri
        end if


        sqlDatoStart = year(useStDatoKri)&"/"&month(useStDatoKri)&"/"&day(useStDatoKri)
        sqlDatoStartATD =  year(now)&"/1/1"

     
        ddNowMinusOne = dateAdd("d", -1, now)

        sqlDatoEnd = year(ddNowMinusOne)&"/"&month(ddNowMinusOne)&"/"&day(ddNowMinusOne)

        daysAnsat = dateDiff("d", sqlDatoStart, ddNowMinusOne)
                    
        ntimPer = 0
        realtimerIalt = 0
        strSQL2 = "SELECT tmnavn, SUM(timer) AS realtimer FROM timer WHERE tmnr = "& usemrn &" AND ("& aty_sql_realhours &" OR tfaktim = 114) AND tdato BETWEEN '"& sqlDatoStart &"' AND '"& sqlDatoEnd &"' GROUP BY tmnr"

        'if session("mid") = 1 then
        'Response.write strSQL2
        'Response.flush
        'end if

        oRec2.open strSQL2, oConn, 3
        if not oRec2.EOF then
        '114 = Korrektion Ultimo
        realtimerIalt = oRec2("realtimer")
        end if

        oRec2.close
                              
        call normtimerPer(usemrn, useStDatoKri, daysansat, 0)
        ntimPerFlex = ntimPer

      
        'balRealNormTxt = formatnumber(( realTimer(x) + korrektionReal(x) ) - normTimer(x),2)

        flexsaldo = (realtimerIalt - ntimPerFlex)



        %>

        <div class="container" style="width:100%; height:100%">
            <div class="portlet">
                   
                  <h3 class="portlet-title"><u>Godkend ønsket Ferie, Tillæg og overarbejde</u> </h3>
                  

                <div class="portlet-body">
                     <div class="portlet-title">
                          <table style="width:100%;">
                              <tr>
                                  <td><h3 style="color:black"><%=medarb_navn %> <!--(<%=medid %>)--></h3></td>
                                  <td style="text-align:right"><img style="width:15%; margin-bottom:2px; margin-left:55px" src="img/cflow_logo.png" /></td>
                              </tr>
                          </table>
                        </div>

                        <br /><br /><br /><br />
                        <div class="row" style="padding-left:100px;">
                            
                            Flexsaldo: <b><%=flexsaldo %></b><br />
                            <form>
                            <table style="width:80%">
                                <thead>
                                    <tr>
                                    <td>Navn</td>
                                    <td>Dato</td>
                                    <td>Komme/gå tid</td>
                                    <td>Overtid ønsket</td>
                                    <td>Godkend antal</td>
                                    </tr>
                                </thead>
                                
                             <%
                            'overtidtot = 0
                            strSQL6 = "SELECT timer As overtidtimer_onsket, tdato FROM timer WHERE tmnr = "&usemrn&" AND tfaktim = 51 ORDER BY tdato" 'AND tdato BETWEEN '"& sqlDatoStartATD &"' AND '"& sqlDatoEnd &"'
                            oRec8.open strSQL6, oConn, 3
                            while not oRec8.EOF

                                 %>
                                 <tr>
                                    <td valign="top"><%=meNavn %></td>
                                    <td><%=formatdatetime(oRec8("tdato"), 2) %></td>
                                    <td><%
                                           '******* Tjekker logindtid ********'
                                            strSQL = "SELECT id, logud, login FROM login_historik WHERE mid = "& usemrn &" AND stempelurindstilling > 0 ORDER BY id DESC LIMIT 3"
                                            oRec.open strSQL, oConn, 3
                                            while not oRec.EOF 
                                            
                                            login = oRec("login")
                                            logud = oRec("logud")
                                            
                                            Response.write left(formatdatetime(login, 3), 5) & " - " & left(formatdatetime(logud, 3), 5) & "<br>" 

                                            oRec.movenext
                                            wend
                                            oRec.close
                                        
                                        %></td>
                                    <td valign="top"><%=oRec8("overtidtimer_onsket") %></td>
                                    <td valign="top"><input type="text"  name="timer_gk" value="<%=oRec8("overtidtimer_onsket") %>" class="form-control input-small" style="width:30px; text-align:right;"  /></td>
                                    </tr>
                                <%
                            
        
                            oRec8.movenext
                            wend 
                            oRec8.close %>

                            </table><br />

                                

                          
                        </div>
                        <div class="row" style="padding-left:100px;">
                            <input type="submit" class="btn btn-small btn-success" value="Godkend >>">
                    </form>
                            </div>
                                        </div>
            </div>
        </div>
<%end select %>

<!--#include file="../inc/regular/footer_inc.asp"-->