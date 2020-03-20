
<%
    thisfile = "timetag_kpi" 
%>

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->

<!--#include file="../inc/regular/top_menu_mobile.asp"-->

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/regular/treg_func.asp"-->

<%
    if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/errors/error_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	response.End
    end if


    ddDato_ugedag = day(now) &"/"& month(now) &"/"& year(now)
    ddDato_ugedag_w = datepart("w", ddDato_ugedag, 2, 2)
            
    varTjDatoUS_man_tt = dateAdd("d", -(ddDato_ugedag_w-1), ddDato_ugedag)

    
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">



<head>

<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>TimeOut mobile</title>

</head>
    



    <%call mobile_header %>


    <div class="container" style="height:100%">
        <div class="portlet">
            <div class="portlet-body">
                
                <!------ Flex saldo ----->
                <%
                    
                    if session("mid") = 1 AND lto = "hestia" then
                    usemrn = 68
                    else
                    usemrn = session("mid")
                    end if


                    call akttyper2009(2)
                    call meStamdata(usemrn)
                    call licensStartDato()
                    licensstdato = licensstdato
                    meAnsatDato = meAnsatDato

                    yNow = year(now)

                    if cDate(meAnsatDato) > cDate(licensstdato) then
                    useStDatoKri = meAnsatDato 
                    else
                    useStDatoKri = licensstdato
                    end if

                    call lonKorsel_lukketPerPrDato(now, usemrn)
                    if lonKorsel_lukketPerDt > useStDatoKri then
                    useStDatoKri = dateAdd("d", 1, day(lonKorsel_lukketPerDt) &"/"& month(lonKorsel_lukketPerDt) &"/"& year(lonKorsel_lukketPerDt))
                    else
                    useStDatoKri = useStDatoKri
                    end if


                    sqlDatoStart = year(useStDatoKri)&"/"&month(useStDatoKri)&"/"&day(useStDatoKri)
                    sqlDatoStartATD =  year(now)&"/1/1"

                    '** Er dd i år 
                    if month(now) <= 4 then 
                    sqlDatoStartFerie =  year(now)-1&"/5/1"
                    sqlDatoEndFerie =  year(now)&"/4/30"
                    else
                    sqlDatoStartFerie =  year(now)&"/5/1"
                    sqlDatoEndFerie =  year(now)+1&"/4/30"
                    end if

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

                    'if daysansat <> 0 then
                    'nTimerPerPrDay = formatnumber((ntimPer/daysansat), 2)
                    'else
                    'nTimerPerPrDay = 0
                    'end if
                    
                    'balRealNormTxt = formatnumber(( realTimer(x) + korrektionReal(x) ) - normTimer(x),2)

                    flexsaldo = (realtimerIalt - ntimPerFlex)


                    'call nortimerStandardDag(meType)
            
                %>

                <!------ Ferie saldo ----->
                <%
                    
                    'ferieafholdt = 0
                    ferieaarAfholdt = 0
                    ferieaarUdbetalt = 0
                    ferieaarAfholdtUdlon = 0
                    'call ferieAfholdtPer(sqlDatoStartFerie, sqlDatoEndFerie, usemrn, 0) 
                     strSQL = "SELECT timer, tdato, month(tdato) AS month, tfaktim, tastedato FROM timer WHERE tmnr = " & usemrn & ""_
	                        &" AND (tfaktim = 14 OR tfaktim = 19 OR tfaktim = 16) AND tdato BETWEEN '"& sqlDatoStartFerie &"' AND '"& sqlDatoEndFerie &"' ORDER BY tdato"
                            'response.Write "sql afholdt " & strSQL
                            oRec7.open strSQL, oConn, 3
                            
                            while not oRec7.EOF

                                call normtimerPer(usemrn , oRec7("tdato"), 0, 0)
	                            if ntimPer <> 0 then
                                ntimPerUse = ntimPer
                                else
                                ntimPerUse = 1
                                end if 

                                if oRec7("tfaktim") = 14 then
                                    ferieaarAfholdt = ferieaarAfholdt + (oRec7("timer") / ntimPerUse)
                                end if

                                if oRec7("tfaktim") = 16 then
                                    ferieaarUdbetalt = ferieaarUdbetalt + (oRec7("timer") / ntimPerUse)
                                end if

                                if oRec7("tfaktim") = 19 then
                                    ferieaarAfholdtUdlon = ferieaarAfholdtUdlon + (oRec7("timer") / ntimPerUse)
                                end if
                                
                        oRec7.movenext
                        wend
                        oRec7.close

                        

                    
                    'strSQL3 = "SELECT SUM(timer) AS ferieafholdt FROM timer WHERE tmnr = "& usemrn &" AND tdato BETWEEN '"& sqlDatoStartFerie &"' AND '"& sqlDatoEndFerie &"' AND tfaktim = 14 GROUP BY tmnr"
                    'oRec3.open strSQL3, oConn, 3
                    'if not oRec3.EOF then
                    'ferieafholdt = oRec3("ferieafholdt")
                    'end if
                    'oRec3.close
                   

                    if ferieAFPerTimer(0) <> 0 then
                        ferieafholdt = ferieAFPerTimer(0)
                    else
                        ferieafholdt = 0
                    end if

                    'response.Write "<br> af:" & ferieAFPerTimer(0)
                    
                    ferieoptjent = 0
                    strSQL4 = "SELECT sum(timer) AS timer, tdato, tfaktim FROM timer WHERE tmnr = "& usemrn &" AND tdato BETWEEN '"& sqlDatoStartFerie &"' AND '"& sqlDatoEndFerie &"' AND (tfaktim = 15 OR tfaktim = 111 OR tfaktim = 112) GROUP BY tfaktim ORDER BY tdato"
                    'response.Write strSQL4
                    'if session("mid") = 1 then
                    'Response.write strSQL4
                    'Response.flush
                    'end if
                    ferieaarOptjent = 0
                    ferieaarOverfort = 0
                    ferieaarOptUdenLon = 0

                    oRec7.open strSQL4, oConn, 3
                    if not oRec7.EOF then
                    
                         call normtimerPer(usemrn, oRec7("tdato"), 6, 0)
	                     if ntimPer <> 0 then
                         normTimerGns5 = (ntimManIgnHellig + ntimTirIgnHellig + ntimOnsIgnHellig + ntimTorIgnHellig + ntimFreIgnHellig + ntimLorIgnHellig + ntimSonIgnHellig)  / 5
                         else
                         normTimerGns5 = 1
                         end if 

                                      
                        if oRec7("tfaktim") = 15 then
                            ferieaarOptjent = oRec7("timer") / normTimerGns5
                        end if

                        if oRec7("tfaktim") = 111 then
                            ferieaarOverfort = oRec7("timer") / normTimerGns5
                        end if

                        if oRec7("tfaktim") = 112 then
                            ferieaarOptUdenLon = oRec7("timer") / normTimerGns5
                        end if

                    
                    end if
                    oRec7.close
                   
                    'response.Write "<br> Ferie opt " & ferieaarOptjent
                    'response.Write "<br> Ferie overfor " & ferieaarOverfort
                    'response.Write "<br> Ferie optudenlon " & ferieaarOverfort
                
                    'ferieSaldo = (ferieaarOptjent - ferieafholdt)
                    call ferieBal_fn(ferieaarOptjent, ferieaarOverfort, ferieaarOptUdenLon, ferieaarAfholdt, ferieaarAfholdtUdlon, ferieaarUdbetalt)
                    ferieSaldo = feriebal

                    'response.Write "<br> tal: " & ferieoptjent
                %>

                <!-- Feriefri -->
                <% 
                    ugp = now
                    strAar = year(ugp)

                    select case lto
                        case "akelius", "xintranet - local" 
                            ferieaarST = strAar &"/1/1"
	                        ferieaarSL = strAar &"/12/31"
                        case else
	                        if cdate(ugp) >= cdate("1-5-"& strAar) AND cdate(ugp) <= cdate("31-12-"& strAar) then
	                            'Response.Write "OK her"
	                            ferieaarST = strAar &"/5/1"
	                            ferieaarSL = strAar+1 &"/4/30"
	                        else
	                            ferieaarST = strAar-1 &"/5/1"
	                            ferieaarSL = strAar &"/4/30"
	                    end if
                    end select


                    select case lto 
                        case "mi", "intranet - local", "plan", "akelius", "outz", "xoko"
                            FefriStartdato = year(now) &"-1-1"
                            FefriSlutdato = year(now) & "-12-31"
                        case else
                            FefriStartdato = ferieaarST
                            FefriSlutdato = ferieaarSL
                    end select




                    feFri = 0
                    feFriBr = 0
                    feFriUd = 0
                    feFriPl = 0

                    strSQLFefri = "SELECT timer, tfaktim, tdato FROM timer WHERE tmnr = "& usemrn & " AND tdato BETWEEN '"& FefriStartdato &"' AND '"& FefriSlutdato &"' AND (tfaktim = 12 OR tfaktim = 13 OR tfaktim = 17 OR tfaktim = 18) ORDER BY tdato"
                    'response.Write strSQLFefri
                    oRec7.open strSQLFefri, oConn, 3
                    while not oRec7.EOF

                        call normtimerPer(usemrn , oRec7("tdato"), 0, 0)
	                    if ntimPer <> 0 then
                        ntimPerUse = ntimPer
                        else
                        ntimPerUse = 1
                        end if

                        select case cint(oRec7("tfaktim"))
                            case 12

                                call normtimerPer(usemrn , oRec7("tdato"), 6, 0)

                                if ntimPer <> 0 then
                                'ntimPerUse = ntimPer/antalDageMtimer
                                normTimerGns5 = (ntimManIgnHellig + ntimTirIgnHellig + ntimOnsIgnHellig + ntimTorIgnHellig + ntimFreIgnHellig + ntimLorIgnHellig + ntimSonIgnHellig)  / 5
                                else
                                normTimerGns5 = 1
                                end if 

                                feFri = feFri + (oRec7("timer") / normTimerGns5)  
                    
                            case 13
                                feFriBr = feFriBr + (oRec7("timer") / ntimPerUse)
                            case 17
                                feFriUd = feFriUd + (oRec7("timer") / ntimPerUse)
                            case 18
                                feFriPl = feFriPl +  (oRec7("timer") / ntimPerUse)

                        end select                        
                    oRec7.movenext
                    wend
                    oRec7.close

                    'sqlNow = year(now) &"-"& month(now) &"-"& day(now)
                    'call normtimerPer(usemrn, sqlNow, 6, 0)
	     
                    'if ntimPer <> 0 then
                    'normTimerGns5 = (ntimManIgnHellig + ntimTirIgnHellig + ntimOnsIgnHellig + ntimTorIgnHellig + ntimFreIgnHellig + ntimLorIgnHellig + ntimSonIgnHellig) / 5
                    'else
                    'normTimerGns5 = 1
                    'end if 

                    'response.Write "<br> Feriefri opt " & feFri
                    'response.Write "<br> Feriefri br " & feFriBr
                    'response.Write "<br> Feriefri ud " & feFriUd
                    'response.Write "<br> Feriefri pl " & feFriPl

                    'feFri = feFri / normTimerGns5
                    'feFriBr = feFriBr / normTimerGns5
                    'feFriUd = feFriUd / normTimerGns5
                    'feFriPl = feFriPl / normTimerGns5

                    feFrisaldo = (feFri - (feFriBr + feFriUd))
                    feFrisaldo = feFrisaldo - feFriPl


                    'response.Write "<br> norm " & normTimerGns5 & " Saldo " & feFrisaldo
                    'response.Write "<br> fefri " & feFri
                    'response.Write "<br> fefriBR " &  feFriBr
                    'response.Write "<br> feFriUd " &  feFriUd
                    'response.Write "<br> feFriPl " &  feFriPl

                %>

                <!------ sygetal ----->
                <%

                    sygtot = 0
                    strSQL5 = "SELECT timer AS sygtimer, tdato FROM timer WHERE tmnr = "& usemrn &" AND tfaktim = 20 AND tdato BETWEEN '"& sqlDatoStartATD &"' AND '"& sqlDatoEnd &"' ORDER BY tdato"
                    
                    oRec7.open strSQL5, oConn, 3
                    while not oRec7.EOF 

                         call normtimerPer(usemrn, oRec7("tdato"), 0, 0)
	                     if ntimPer <> 0 then
                         ntimPerUse = ntimPer
                         else
                         ntimPerUse = 1
                         end if 

                    sygtot = sygtot + (oRec7("sygtimer")*1 / ntimPerUse)
                    oRec7.movenext
                    wend 
                    oRec7.close
                    

                    barnsygtot = 0
                    strSQL6 = "SELECT timer As barnsygtimer, tdato FROM timer WHERE tmnr = "&usemrn&" AND tfaktim = 21 AND tdato BETWEEN '"& sqlDatoStartATD &"' AND '"& sqlDatoEnd &"' ORDER BY tdato"
                    oRec7.open strSQL6, oConn, 3
                    while not oRec7.EOF

                         call normtimerPer(usemrn, oRec7("tdato"), 0, 0)
	                     if ntimPer <> 0 then
                         ntimPerUse = ntimPer
                         else
                         ntimPerUse = 1
                         end if 

                    barnsygtot = barnsygtot + (oRec7("barnsygtimer")*1/ntimPerUse)
                    oRec7.movenext
                    wend 
                    oRec7.close
                   

                   

                    if ferieSaldo < 0 then
                    flexferieCol = "red"
                    else
                    flexferieCol = "greenyellow"
                    end if

                    if feFrisaldo < 0 then
                    flexferiefriCol = "red"
                    else
                    flexferiefriCol = "greenyellow"
                    end if

                    if flexsaldo < 0 then
                    flexsaldoCol = "red"
                    else
                    flexsaldoCol = "greenyellow"
                    end if

                    if sygtot < 0 then
                    sygtotCol = "red"
                    else
                    sygtotCol = "greenyellow"
                    end if

                    if barnsygtot < 0 then
                    barnsygtotCol = "red"
                    else
                    barnsygtotCol = "greenyellow"
                    end if

                %>

                <table class="table dataTable" style="font-size:125%">

                    <tr>
                        <td><%=ttw_txt_018 %></td>
                        <td style="text-align:right"><%=formatnumber(ferieSaldo,2) %> d.</td>
                        <td style="vertical-align:middle"><div style="height:10px; width:10px; background-color:<%=flexferieCol%>;"></div></td>
                    </tr>

                    <tr>
                        <td><%=ttw_txt_028 %></td>
                        <td style="text-align:right"><%=formatnumber(feFrisaldo,2) %> d.</td>
                        <td style="vertical-align:middle"><div style="height:10px; width:10px; background-color:<%=flexferiefriCol%>;"></div></td>
                    </tr>

                    <tr>
                        <td><%=ttw_txt_019 %></td>
                        <td style="text-align:right"><%=formatnumber(flexsaldo,2) %> t.</td>
                        <td style="vertical-align:middle"><div style="height:10px; width:10px; background-color:<%=flexsaldoCol%>;"></div></td>
                    </tr>

                    <tr>
                        <td><%=ttw_txt_020 %></td>
                        <td style="text-align:right"><%=formatnumber(sygtot,2) %> d.</td>
                        <td style="vertical-align:middle"><!--<div style="height:10px; width:10px; background-color:<%=sygtotCol%>;"></div>--></td>
                    </tr>

                    <tr>
                        <td><%=ttw_txt_021 %></td>
                        <td style="text-align:right"><%=formatnumber(barnsygtot,2) %> d.</td>
                        <td style="vertical-align:middle"><!--<div style="height:10px; width:10px; background-color:<%=barnsygtotCol%>;"></div>--></td>
                    </tr>

                   <!-- <tr>
                        <td>Ferie fridage</td>
                    </tr>
                    <tr>
                        <td>Felx saldo</td>
                    </tr>
                    <tr>
                        <td>Syg åtd</td>
                    </tr>
                    <tr>
                        <td>Barn syg åtd</td>
                    </tr> -->

                </table>

            </div>
        </div>
    </div>



</html>
<!--#include file="../inc/regular/footer_inc.asp"-->