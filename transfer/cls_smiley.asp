




<%

function godkenderTimeriUge(medid, varTjDatoUS_manSQL, varTjDatoUS_sonSQL, SmiWeekOrMonth)


        call akttyper2009(2)

        '*** NÅR LM Godkender HR hos TIA må fakturerbare ikke godkendes 
        '*** Hvis godkend projekttimer er slået til. -> lav i kontrolpanel
        SELECT CASE lto
        case "tia", "intranet - local"
        strSQLKriEks = " AND (" & aty_sql_admin & ")"
        strSQLKriEks = replace(strSQLKriEks, "fakturerbar", "tfaktim")
        case else
        strSQLKriEks = ""
        end select

        godkendtdato = Year(now) &"-"& Month(now) &"-"& Day(now)


       '** GODKENDER TIMERNE DER ER INDTASTET
	    strSQLup = "UPDATE timer SET godkendtstatus = 1, godkendtstatusaf = '"& session("user") &"', godkendtdato = '"& godkendtdato &"' WHERE tmnr = "& medid
	    if cint(SmiWeekOrMonth) = 0 then
        strSQLup = strSQLup & " AND tdato BETWEEN '"& varTjDatoUS_manSQL &"' AND '" & varTjDatoUS_sonSQL & "'" 
        else
        varTjDatoUS_man_mth = datepart("m", varTjDatoUS_man,2,2)
        strSQLup = strSQLup & " AND MONTH(tdato) = '"& varTjDatoUS_man_mth & "'" 
        end if

        strSQLup = strSQLup & " AND godkendtstatus <> 1 " & strSQLKriEks & " AND overfort = 0"


        'if session("mid") = 1 then
	    'Response.Write strSQLup & "<br>"
	    'Response.flush
        'end if

	    oConn.execute(strSQLup)

 end function



'** Finder dag for sidste rettidgie afslutning. Baseret på kontrolpanel indstilllinger i smileyAfslutSettings()
public cDateUge, cDateAfs, cDateUgeTilAfslutning, smileysttxt, smileyImg, weekafslTxt, kommegaa60
        
    
        function rettidigafsl(ugeVal, rettidigafsl_do)
            
               'Response.write "rettidigafsl_do: " & rettidigafsl_do & " ugeVal: "& ugeVal &"<br>"
               'Response.end

               call smileyAfslutSettings()


                
                select case cint(SmiWeekOrMonth) 
                case 0 '0: ugebasis, 1: Måned
                weekafslTxt = funk_txt_001
                case 1
                weekafslTxt = funk_txt_002
                case 2
                weekafslTxt = funk_txt_003
                end select


                cDateUgeTilAfslutning = year(ugeVal) &"/"& month(ugeVal) &"/"& day(ugeVal)

                select case cint(SmiWeekOrMonth)
                case 0,1


                        'Response.write "SK SmiWeekOrMonth: "& SmiWeekOrMonth & " SmiantaldageCount: "& SmiantaldageCount

                        if SmiantaldageCount < 8 then 'antal dage der skal lægges til (mandag: 1 søndag: 7)
                            cDateUge = dateadd("d", SmiantaldageCount, year(ugeVal) &"/"& month(ugeVal) &"/"& day(ugeVal) &" "& SmiantaldageCountClock)
		                    cDateAfs = year(now) &"/"& month(now) &"/"& day(now) &" "& time
                        else '1 hverdag
                    
                            '1 i nsæte måned
                            forsteDag = dateAdd("m", 1, ugeVal)
                            forsteDag = "1-"& month(forsteDag) & "-"& year(forsteDag)
                
                           forsteDagWeekDay = datePart("w", forsteDag, 2,2)
                           'response.write "forsteDag: "& forsteDag & " forsteDagWeekDay: "& forsteDagWeekDay &"<br>"     


                   
                           if forsteDagWeekDay = 6 then  
                           forsteDag = dateAdd("d", 2, forsteDag)
                           end if

                           if forsteDagWeekDay = 7 then  
                           forsteDag = dateAdd("d", 1, forsteDag)
                           end if
         

                           if SmiantaldageCount = 9 then '1 +7 hverdag
                           forsteDag = dateAdd("d", 7, forsteDag)
                           end if

                           
                           if SmiantaldageCount = 10 then '1 +3 hverdag
                           forsteDag = dateAdd("d", 3, forsteDag)
                           end if

                           cDateUge = forsteDag &" "& SmiantaldageCountClock
		                   cDateAfs = year(now) &"/"& month(now) &"/"& day(now) &" "& time           

                        end if


                 case else
            
                     forsteDag = dateAdd("d", 1, ugeVal)
                     if datePart("w", forsteDag, 2,2) = 6 then 'lørdag
                     forsteDag = dateAdd("d", 2, forsteDag)
                     end if             

                     if datePart("w", forsteDag, 2,2) = 7 then 'søndag
                     forsteDag = dateAdd("d", 1, forsteDag)
                     end if           

                     forsteDag = formatdatetime(forsteDag, 2)

                     cDateUge = forsteDag &" "& SmiantaldageCountClock
		             cDateAfs = year(now) &"/"& month(now) &"/"& day(now) &" "& time

                 end select

                 'cDateUge = forsteDag &" "& SmiantaldageCountClock
		         'cDateAfs = year(now) &"/"& month(now) &"/"& day(now) &" "& time

                 'if session("mid") then
                 'response.write "HER: " & cDateUge & " forsteDag : " & forsteDag  & " "& datePart("w", forsteDag, 2,2) & ""
                 'end if
             
		     


    end function


public afslutugeBasisKri, afslProc, afslProcKri, weekSelectedTjk 
function afslutkri(varTjDatoUS_son, tjkTimeriUgeDt, usemrn, lto, SmiWeekOrMonth)

        select case cint(SmiWeekOrMonth) 
        case 0 
        weekSelectedTjk = dateAdd("d", 7, varTjDatoUS_son)

         call normtimerPer(usemrn, varTjDatoUS_son, 6, 0) 
         '** gns timer pr. iforhold til samlet arbejdsuge, baseret på en 5 dages arbejdsuge. ***'
         '** Således at det er underordnet om man indtaster på en tirsdag (7,5) eller en fredag (7,0) '****'
         normTimerprUge = (ntimManIgnHellig + ntimTirIgnHellig + ntimOnsIgnHellig + ntimTorIgnHellig + ntimFreIgnHellig + ntimLorIgnHellig + ntimSonIgnHellig)
         afslutugeBasisKri = normTimerprUge

        case 1
        varTjDatoUS_tjk = varTjDatoUS_son 'dateAdd("d", -6, varTjDatoUS_son) 'finder mandag i den valgte uge. St. dato for ugne bestemmer hvilken måned man er ved at afslutte.
        weekSelectedTjk = "1-"& month(varTjDatoUS_tjk) &"-"& year(varTjDatoUS_tjk)
        
        '**bruges ikke ved måndes aflutning ?
        'call normtimerPer(usemrn, varTjDatoUS_son, 30, 0) 
         

        case 2
        varTjDatoUS_tjk = formatdatetime(tjkTimeriUgeDt, 2)
        weekSelectedTjk = formatdatetime(tjkTimeriUgeDt, 2) '"1-"& month(varTjDatoUS_tjk) &"-"& year(varTjDatoUS_tjk)
        
        varTjDatoUS_tjkSQL = year(varTjDatoUS_tjk) & "/" & month(varTjDatoUS_tjk) & "/" & day(varTjDatoUS_tjk)

            select case lto 
            case "dencker", "intranet - local" 'Løn timer - komme/gå

           

            call fLonTimerPer(varTjDatoUS_tjk, 0, 3, usemrn) 
            
            afslutugeBasisKri = formatnumber(totalTimerPer100/60, 2)

            call timerogminutberegning(totalTimerPer100) 
	        kommegaa60 = thoursTot &":"& left(tminTot, 2) 

           
            case else 'Norm timer

            call normtimerPer(usemrn, tjkTimeriUgeDt, 0, 0) 
            afslutugeBasisKri = ntimper

            normTimerprUge = (ntimManIgnHellig + ntimTirIgnHellig + ntimOnsIgnHellig + ntimTorIgnHellig + ntimFreIgnHellig + ntimLorIgnHellig + ntimSonIgnHellig)
            afslutugeBasisKri = normTimerprUge
    
            end select

        end select

            'select case lto 
            'case "dencker", "intranet - local"

            '    freDagiSidsteUge = tjkTimeriUgeDt 'tjekdag(1) 'dateAdd("d", -3, tjekdag(1))
            '    freDagiSidsteUge = year(freDagiSidsteUge) &"/"& month(freDagiSidsteUge) &"/"& day(freDagiSidsteUge) 
            '    call fLonTimerPer(freDagiSidsteUge, 6, 21, usemrn)

               'response.write "totaltimerPer: " & totaltimerPer
               'response.write "freDagiSidsteUge: "& freDagiSidsteUge & "<br>"
               'response.write "<br>totalTimerPer100: "& totalTimerPer100 & "("& totalTimerPer100/60& ")"

            'afslutugeBasisKri = formatnumber(totalTimerPer100/60, 0) 
            'case else
             'end select
            
          



            if totTimerWeek > 0 AND afslutugeBasisKri > 0 then
            afslProc = formatnumber(((totTimerWeek/1)/(afslutugeBasisKri/1)) * 100, 0)
            else
            afslProc = 0
            end if


            if cdbl(afslProc) >= cdbl(afslutugekri_proc) OR (afslutugeBasisKri = 0) then 'Er Reltimer/NormtidsKriterie opflydt Eller der er ikke nongen login tid eller normtid = 0.
            afslProcKri = 1
            else
            afslProcKri = 0
            end if



end function


public akttypeKrism, afslutugekri, afslutugekri_proc, tjkTimeriUgeDt, tjkTimeriUgeDtDay, tjkTimeriUgeSQL, tjkTimeriUgeDtTxt, afslugeDatoTimerudenMatch
function timeKriOpfyldt(lto, sidsteUgenrAfsl, meType, usemrn, SmiWeekOrMonth)

    'response.write "sidsteUgenrAfsl: " & sidsteUgenrAfsl & "#: "
    select case cint(SmiWeekOrMonth)
    case 0,1
    tjkTimeriUgeDt = dateAdd("d", 7, sidsteUgenrAfsl)
    case 2
    '** Hvis dd er større en sidste afsluttet vises næste dag til afslutning. (med Submit)
    '** Er den = med dd eller mindre vise "Periode er afsluttet.."
        
        'response.write "<br>A: sidsteUgenrAfsl: " & sidsteUgenrAfsl & "#: "

        if cDate(formatdatetime(sidsteUgenrAfsl, 2)) < cDate(formatdatetime(now, 2)) AND cint(sidsteUgenrAfslFundet) = 1 then
        tjkTimeriUgeDt = dateAdd("d", 1, sidsteUgenrAfsl) '** Dagen efter sidste afsluttede dag 'sidsteUgenrAfsl
        else
        tjkTimeriUgeDt = sidsteUgenrAfsl 
        end if

         'response.write "<br>B: tjkTimeriUgeDt: " & tjkTimeriUgeDt & "#: "
    end select

    tjkTimeriUgeDtDay = datepart("w", tjkTimeriUgeDt, 2,2)
    

    select case cint(SmiWeekOrMonth)
    case 0,1

        if tjkTimeriUgeDtDay <> 1 then 
        tjkTimeriUgeDt = dateAdd("d", -(tjkTimeriUgeDtDay-1), tjkTimeriUgeDt)
        end if

    case 2
    
        tjkTimeriUgeDt = tjkTimeriUgeDt

         
        '*** Hvis næste dag til afslutning er en lørdag eller søndag
        if datepart("w", tjkTimeriUgeDt, 2,2) = 6 then
        tjkTimeriUgeDt = dateAdd("d", 2, tjkTimeriUgeDt)
        end if

        if datepart("w", tjkTimeriUgeDt, 2,2) = 7 then
        tjkTimeriUgeDt = dateAdd("d", 1, tjkTimeriUgeDt)
        end if


    end select
     
    tjkTimeriUgeDtTxt = tjkTimeriUgeDt

    'select case lto
    'case "dencker" 'tjekker fra fredag i sidste uge
    '      tjkTimeriUgeDt = dateAdd("d", -3, tjkTimeriUgeDt)
    'case else 
    '      tjkTimeriUgeDt = tjkTimeriUgeDt
    'end select

    tjkTimeriUgeSQL = year(tjkTimeriUgeDt)&"/"&month(tjkTimeriUgeDt)&"/"&day(tjkTimeriUgeDt)
    'response.write "tjkTimeriUgeDt: "& tjkTimeriUgeDt & "sidsteUgenrAfsl: "& sidsteUgenrAfsl & "tjkTimeriUgeDtDay: "&tjkTimeriUgeDtDay & " varTjDatoUS_man: "& datepart("w", varTjDatoUS_man, 2,2) & "<br><br>"
        
              
              
            '** Er der kriterier for antal timer derskal være opnået i% før uge må afsluttes? *'
            afslutugekri = 0
            afslutugekri_proc = 0
            strSQlmt = "SELECT afslutugekri, afslutugekri_proc FROM medarbejdertyper WHERE id = "& meType
            oRec6.open strSQlmt, oConn, 3
            if not oRec6.EOF then
            
               afslutugekri  = oRec6("afslutugekri")
               afslutugekri_proc  = oRec6("afslutugekri_proc")
                  
            end if 
            oRec6.close
                    
        
         if cint(afslutugekri) = 2 then 'Kun Fakturerbare timer
         akttypeKrism = replace(aty_sql_fakbar, "fakturerbar","tfaktim")
         else
         akttypeKrism = aty_sql_realhours
         end if


         '** Findes der timereg. uden match UANSET DATO, må periode ikke afsluttes ***
         strSQLherrelose = "SELECT Sum(timer), tdato FROM timer_import_temp WHERE medarbejderid = "& usemrn &" AND overfort = 0 GROUP BY medarbejderid" 
         oRec6.open strSQLherrelose, oConn, 3
            if not oRec6.EOF then
            
               afslutugekri = 10
               afslugeDatoTimerudenMatch = oRec6("tdato")
               
                  
            end if 
            oRec6.close


end function


function ugeAfsluttetStatus(tjkDato, showAfsuge, ugegodkendt, ugegodkendtaf, SmiWeekOrMonth)

     'if session("mid") = 1 then
    'response.write "tjkDato:" & tjkDato
     'end if
     'response.write "SmiWeekOrMonth XXX: "& SmiWeekOrMonth & " varTjDatoUS_son: "& varTjDatoUS_son & " ugeNrStatus: "& ugeNrStatus & " showAfsuge: " & showAfsuge

              if cint(showAfsuge) = 0 then 
    
                        select case cint(SmiWeekOrMonth)
                        case 0 
                        periodeTxt = tsa_txt_005 & " "& datepart("ww", tjkDato, 2 ,2)
                        case 1
                        periodeTxt = monthname(datepart("m", tjkDato, 2 ,2))
                        case 2
                        periodeTxt = tjkDato
                        end select%>
              

                        
                      <span style="color:#999999; font-size:12px; background-color:#F7f7f7; padding:5px;"><span style="color:yellowgreen;"><i>V</i></span> - <%=periodeTxt &" "& funk_txt_071 %> <b><%=funk_txt_004 %></b> <%=funk_txt_005 %></span>
                                <%

                                select case lto
                                case "intranet - local", "tia"

                                'Afviste timer
                                afvisteTimer = 0
                                tjkDatoSt = dateAdd("d", -6, tjkDato)
                                strSQLtim = "SELECT ROUND(SUM(timer),2) AS sumtimer FROM timer WHERE ("& aty_sql_realhours &") AND tdato BETWEEN '"& year(tjkDatoSt) &"-"& month(tjkDatoSt) &"-"& day(tjkDatoSt) &"' AND '"& year(tjkDato) &"-"& month(tjkDato) &"-"& day(tjkDato) &"' AND tmnr = "& usemrn & " AND godkendtstatus = 2 GROUP BY tmnr"

                                'Response.write strSQLtim
                                'Response.flush      

                                oRec3.open strSQLtim, oConn, 3
                                if not oRec3.EOF then 

                                afvisteTimer = oRec3("sumtimer") 
               
                                end if
                                oRec3.close

                                end select

                                if cdbl(afvisteTimer) <> 0 then%>
                                <span style="color:#FFFFFF; font-size:12px; background-color:red; padding:2px;"><%=afvisteTimer%></span>
                                <%end if %>


            
            
                        <%
                        select case lto
                        case "tec", "esn"
                        lukTxt = funk_txt_006
                        case else
                        lukTxt = funk_txt_007
                        end select    
                            
                            
                            
                        select case ugegodkendt
                        case 1
                        call meStamdata(ugegodkendtaf)
                        ugegodkendtStatusTxt = ""& funk_txt_008 &" <b>"& lukTxt &"</b> af <a href='mailto:"&meEmail&"'>"& left(meNavn, 10) & " ["& meInit &"]</a>"
                        ugstCol = "yellowgreen" '"#DCF5BD"
                        ugstFtc = "green"
                        ugstBd = "#6CAE1C"
                        case 2
                        call meStamdata(ugegodkendtaf)
                        ugegodkendtStatusTxt = " "& funk_txt_008 &" <b>"& funk_txt_009 &"</b> af <a href='mailto:"&meEmail&"'>"& left(meNavn, 10) & " ["& meInit &"]</a>"
                       
                        ugstCol = "#FF6666"
                        ugstFtc = "#000000"
                        ugstBd = "#CCCCCC"
                        case else
                        ugegodkendtStatusTxt = " "& funk_txt_011 &" "& lukTxt &"/"& funk_txt_009 &""
                        ugstCol = ""
                        ugstFtc = "#999999"
                        ugstBd = "#cccccc"
                        end select %>


        
                         <span style="color:<%=ugstFtc%>; font-size:12px; background-color:<%=ugstCol%>; padding:5px;">
                         <%=ugegodkendtStatusTxt %>
                        </span>

                            <%if cint(ugegodkendt) = 2 then 'afvist 
                              response.write "<br><br><div style=""color:#999999; padding:0px; font-size:12px;"">"& funk_txt_012 &"<br>"& funk_txt_013 &"<br><br>"_
                              &"<b>"& funk_txt_014 &":</b><br>"& ugegodkendtTxt &"</div>"
                              end if %>



            <%end if 


end function



'** Godkend ugeseddel funktionen, med felter og submit funktion
function godkendugeseddel(fmlink, usemrn, varTjDatoUS_man, rdir) 

               
              varTjDatoUS_son = dateadd("d", 6, varTjDatoUS_man)
                

              'response.write varTjDatoUS_son & "<br>" & fmlink

              call smileyAfslutSettings()

              select case cint(SmiWeekOrMonth) 
              case 0 'uge
                periodeTxt = ""& funk_txt_015 &""
                periodeNavn = datepart("ww", varTjDatoUS_son, 2,2)
              case 1
                periodeTxt = ""& funk_txt_016 &""
                periodeNavn = left(monthname(datepart("m", varTjDatoUS_son, 2,2)), 3) & "."
             case 2
                 periodeNavn = weekdayname(weekday(varTjDatoUS_son, 2), 0,2) 
                 periodeTxt = varTjDatoUS_son
              end select

             select case lto
                case "tec", "esn"
                lukTxt = ""& funk_txt_017 &""
                lukTxt1 = ""& funk_txt_018 &""
                lukTxt2 = ""& funk_txt_019 &""
                teamlederTxt = ""& funk_txt_020 &""
                case else
                lukTxt = ""& funk_txt_021 &""
                lukTxt1 = ""& funk_txt_022 &""
                lukTxt2 = ""& funk_txt_023 &""
                teamlederTxt = ""& funk_txt_024 &""
                end select
              
               %>
                <hr />
                <h4><%=lukTxt2 &" "& lcase(periodeTxt) &" ("& periodeNavn %>)<br /><span style="font-size:11px; color:#999999; font-weight:lighter;"><i><%=teamlederTxt %></i></span></h4>
               <%


                  'if session("mid") = 1 then
                  'response.write "varTjDatoUS_man: "&varTjDatoUS_man&" showAfsuge: "& showAfsuge & " ugegodkendt: "& ugegodkendt
                  'end if

               if cint(showAfsuge) = 0 then
               
               
                if cint(ugegodkendt) <> 1 then 

                
                       if cint(ugegodkendt) = 2 then

                        call meStamdata(ugegodkendtaf)%>
                        <div style="background-color:#FF6666; padding:5px;"><b><%=periodeTxt &" "& funk_txt_025 %>!</b><br />
                        <span style="font-size:9px; line-height:12px; color:#ffffff;"><i><%=ugegodkendtdt &" "& funk_txt_026 &" "& meNavn %></i></span>
                        <%if len(trim(ugegodkendtTxt)) <> 0 then %>
                        <br />
                        <span style="font-size:9px; line-height:12px; color:#000000;"><i><%=left(ugegodkendtTxt, 200) %></i></span>
                        <%end if %>
                        </div>

                        <%end if %>



                           <form action="<%=fmLink%>&func=godkendugeseddel" method="post">
                               <table width=90% cellpadding=0 cellspacing=0 border=0>
                            <tr><td class=lille><br /> 
                                

                           <span style="font-size:11px;"><b><%=lukTxt2 &" "& periodeTxt %></b></span><br />
                           <%=funk_txt_027 &" "& periodeTxt &" "& lukTxt %>, <%=lukTxt &" "& funk_txt_028 %><br /><br />
                
                           <input id="Submit2" type="submit" value="<%=lukTxt2 &" "& periodeTxt %> >>" style="font-size:9px; width:120px;" />
             
                           </td></tr></table>
                           </form>

                <%else 
                        
                        call meStamdata(ugegodkendtaf)%>
                        <div style="color:green; font-size:11px; background-color:yellowgreen; padding:5px;"><b><%=funk_txt_029 &" "& periodeTxt &" "& funk_txt_071 &" "& lukTxt1 %>!</b><br />
                        <span style="font-size:9px; line-height:12px; color:#ffffff;"><i><%=ugegodkendtdt &" "& funk_txt_026 &" "& meNavn %></i></span></div>
                
                <%end if %>

               <%else %>
                
                 <form action="<%=fmLink%>&func=adviserugeafslutning" method="post">
                <table width=90% cellpadding=0 cellspacing=0 border=0>
                <tr><td class=lille>

                  <span style="font-size:11px;"><b><%= lukTxt2 &" "&periodeTxt %></b></span><br />
                 <%=funk_txt_029 &" "& periodeTxt &" "& funk_txt_030 &" "& lukTxt &" "& funk_txt_031 %><br />
                
            

                <%if len(trim(request("showadviseringmsg"))) <> 0 then %>
                <br />
                  <div style="color:#000000; font-size:11px; background-color:#cccccc; padding:20px;"><b><%=funk_txt_032 %></b></div>
                <%else %>
                  <br />
                  <b><%=funk_txt_033 %></b> <%=funk_txt_034 &" "& periodeTxt &" "& funk_txt_035 %><br /><br />
                <input id="Submit1" type="submit" value="<%=godkendweek_txt_109 %> >>" style="font-size:9px; width:120px;" />
                <%end if %>
           
             
               </td></tr></table>
                </form>
               <%end if 
               
               
               if cint(ugegodkendt) <> 2 AND cint(showAfsuge) = 0 then%>

               <form action="<%=fmLink%>&func=afvisugeseddel" method="post">
               <table width=90% cellpadding=0 cellspacing=0 border=0>
               <tr><td class=lille>
               <br />
               <span style="font-size:11px;"><b><%=funk_txt_036 &" "& periodeTxt %></b></span><br />
               <%=funk_txt_037 %>:<br />
               <textarea name="FM_afvis_grund" style="width:200px; height:40px;"></textarea><br /><br />
                <input id="Submit3" type="submit" value="<%=godkendweek_txt_110 &" " %> <%=periodeTxt %> >>" style="font-size:9px; width:120px;" /><br />
                 <span style="color:#999999;"><%=funk_txt_038 &" "& lukTxt1 &" "& funk_txt_039 %></span>
                  </td></tr></table>
                 
               </form>

               <%end if 

                   %><br />&nbsp;<%
       
    end function




    
function smileyAfslutBtn(SmiWeekOrMonth)

             select case cint(SmiWeekOrMonth) 
             case 0 
             afslutTxt = tsa_txt_345
             case 1
             afslutTxt = tsa_txt_426
             case 2
             afslutTxt =  tsa_txt_533 '"Afslut dag"
             end select %>
          
        	<span style="color:#000000; background-color:#DCF5BD; font-size:12px; padding:5px;"><a href="#" id="s1_k">[+] <%=afslutTxt %></a></span>&nbsp;



<%end function

    
sub sweekmsg
%>
 <br /><b><%=session("user") &"</b> "& lcase(tsa_txt_087) %>?</b><br /><%=tsa_txt_383 %> 

	           
<%
end sub




'************************************************
'** Afslut uge funktionen MAIN '*****************
'************************************************
public erDetteuge53, maaAfslutteUge
function afslutuge(weekSelected, visning, tjkDag7, rdir, SmiWeekOrMonth)
    
    'response.write "weekSelected: "& weekSelected
    'varTjDatoUS_man = dateAdd("d", -7, varTjDatoUS_son) 


    'Response.write "HER RDIR: "& rdir

    if thisfile = "ugeseddel_2011.asp" OR thisfile = "favorit.asp" then
    menu2015lnk = "../timereg/"
    else
    menu2015lnk = ""
    end if


    select case cint(SmiWeekOrMonth)
        case 0 'uge aflsutning

	        if visning = 0 then 'denne uge
            sidsteUgeTxt = funk_txt_040
            sidstedagisidsteuge = dateadd("d", -7, weekSelected)
            else
            sidsteUgeTxt = funk_txt_041
            sidstedagisidsteuge = dateadd("d", -7, weekSelected)
            end if

        case 1 'Måned
        sidsteUgeTxt = funk_txt_042
        sidstedagisidsteuge = weekSelected 'dateadd("d", -31, weekSelected)
        case 2 'Dag
        sidsteUgeTxt = funk_txt_043
        sidstedagisidsteuge = year(weekSelected) & "-" & month(weekSelected) & "-" & day(weekSelected) 'dateadd("d", -31, weekSelected)
        end select
	    
     
	 

        ugeNrAfsluttet = "1-1-2044"
		detteaar = datepart("yyyy", sidstedagisidsteuge, 2,2)

       
        erDetteuge53 = 0
        select case cint(SmiWeekOrMonth) 
        case 0 'uge aflsutning
		sidstedagiuge = datepart("ww", sidstedagisidsteuge, 2, 2)

            if datePart("ww", sidstedagisidsteuge, 2,2) = 53 then
            erDetteuge53 = 1 
            end if

        case 1
        sidstedagiuge = datepart("m", sidstedagisidsteuge, 2, 2)
        case 2
        sidstedagiuge = sidstedagisidsteuge
        end select

        'if session("mid") = 1 then
        'response.write "HER: sidstedagisidsteuge: "& sidstedagisidsteuge & " weekSelected: "& weekSelected &" detteaar "& detteaar & "sidstedagiuge: " & sidstedagiuge & " erDetteuge53: " & erDetteuge53
        'end if

        '** finder kriterie for rettidig afslutning
        call rettidigafsl(sidstedagisidsteuge, 1)

        'Response.write "sidstedagiuge: " & sidstedagiuge & "<br>"

        '** tjekker om uge er afsluttet og viser afsluttet eller form til afslutning
        '** ugeNrAfsluttet, showAfsuge, cdAfs, ugegodkendt, ugegodkendtaf, ugegodkendtTxt, ugegodkendtdt
		call erugeAfslutte(detteaar, sidstedagiuge, usemrn, SmiWeekOrMonth, 0)
        
        '** tjekker om valgte dag/uger/md er afsluttet
        '** Viser HÅRD nedlukning ved SLIP overskreddet 
        '** sidsteUgenrAfsl, sidsteUgenrAfslFundet, slip_smiley_agg_lukper
        call afsluttedeUger(detteaar, usemrn, 0)
		

        'response.write "sidsteUgenrAfslFundet: "& sidsteUgenrAfslFundet & " dag i uge FØR: "& datepart("w", sidsteUgenrAfsl, 2,2) &" dt:("& sidsteUgenrAfsl &")  sidsteUgenrAfsl: " & sidsteUgenrAfsl & " showAfsuge: " & showAfsuge
        
                    if cint(sidsteUgenrAfslFundet) = 1 then

                        select case cint(SmiWeekOrMonth)
                        case 0 'uge aflsutning
                        
                            'if datePart("ww", sidsteUgenrAfsl, 2,2) = 53 then
                            'sidstedagisidsteAfsluge = dateAdd("d", 0, sidsteUgenrAfsl)
                            'else
                            sidstedagisidsteAfsluge = dateAdd("d", 7, sidsteUgenrAfsl) 
                            'end if

                        case 1
                        sidstedagisidsteAfsluge = dateAdd("m", 1, sidsteUgenrAfsl)
                        sidstedagisidsteAfsluge = dateAdd("d", -1, sidstedagisidsteAfsluge) 
                        case 2 
                        sidstedagisidsteAfsluge = dateAdd("d", 1, sidsteUgenrAfsl)
            
                        '*** Hvis næste dag til afslutning er en lørdag eller søndag
                        if datepart("w", sidstedagisidsteAfsluge, 2,2) = 6 then
                        sidstedagisidsteAfsluge = dateAdd("d", 2, sidstedagisidsteAfsluge)
                        end if

                        if datepart("w", sidstedagisidsteAfsluge, 2,2) = 7 then
                        sidstedagisidsteAfsluge = dateAdd("d", 1, sidstedagisidsteAfsluge)
                        end if

                        end select
                    else
                    sidstedagisidsteAfsluge = sidsteUgenrAfsl 
                    end if
            

                    'response.write "sidstedagisidsteAfsluge:" & sidstedagisidsteAfsluge & " dag i uge: "& datepart("w", sidstedagisidsteAfsluge, 2,2) &" sidsteUgenrAfsl: " & sidsteUgenrAfsl
                    'response.write "UGE: "& datePart("ww", sidsteUgenrAfsl, 2,2)

                    select case cint(SmiWeekOrMonth) 
                    case 0
                            '*** Sikrer at sidste ugedag der er fundet altid er søndag. (sidstedag i ugen) *** 
                            '*** Afslut på månedsbasis skal altid være den sidste dag i md. og ikke nødvendigvis søndag
                            dayOfWeekThis = datePart("w",sidstedagisidsteAfsluge,2,2)
                            '1 = (dayOfWeek - d) 
                            dThis = (dayOfWeekThis - 1)
                            firstdayOfFirstWeekThis = dateAdd("d", -dThis, sidstedagisidsteAfsluge)
                            firstThursdayOfFirstWeekThis = dateAdd("d", +6, firstdayOfFirstWeekThis)
            
                            sidstedagisidsteAfsluge = firstThursdayOfFirstWeekThis
                            'sidstedagisidsteAfsluge = dateAdd("d", 7, sidstedagisidsteAfsluge)

                    case 1 'måneds afslutning ==> Altid sidste dag i måned
                
                        select case datepart("m", sidstedagisidsteAfsluge, 2,2)
                        case "1","3","5","7","8","10","12"
                        dayNumber = "31"
                        case "2"
                        
                             select case datepart("yyyy", sidstedagisidsteAfsluge, 2,2)
                             case "2000", "2004", "2008", "2012", "2016", "2020", "2024", "2028", "2032", "2036", "2040", "2044"
                             dayNumber = "29"
                             case else
                             dayNumber = "28"
                             end select
                            

                        case else
                        dayNumber = "30"
                        end select 
                        

                        sidstedagisidsteAfsluge = dayNumber &"-"& month(sidstedagisidsteAfsluge) &"-"& year(sidstedagisidsteAfsluge)

                    case 2 'Dag
                    end select



        '*** Medarbejder og overskift, status og smiley Ikon
        weekSelectedThis = dateAdd("d", 7, now) 
        call showsmiley(weekSelectedThis, 1, usemrn, SmiWeekOrMonth)


        call meStamdata(usemrn)

   
    
    %>



<form action="<%=menu2015lnk%>timereg_akt_2006.asp?func=opdatersmiley&fromsdsk=<%=fromsdsk%>&rdir=<%=rdir %>&varTjDatoUS_man=<%=varTjDatoUS_man %>&usemrn=<%=usemrn%>" method="POST">
<table border='0' cellspacing='0' cellpadding='0' width="100%">

<tr>
   
	<td valign=top bgcolor="#FFFFFF">

        
   
	<%
	
        
        
        sidstedagisidsteugeTjk = sidstedagisidsteuge
        if (cDate(meAnsatDato) > cDate(sidstedagisidsteugeTjk) OR cDate(meOpsagtDato) < cDate(sidstedagisidsteugeTjk)) then
        '*** Funktione endnu ikke aktiv / 

                    select case lto
                    case "xtec", "xintranet - local"

                      %>
                    <div style="background-color:#ffdfdf; padding:10px;">
                        <b><%=funk_txt_044 %></b><br />
                     <%=funk_txt_045 %><br /> 
                     <%=funk_txt_046 %>
             
          
                        </div>

                    <%

                    case else

                    %>
                    <div style="background-color:#ffdfdf; padding:10px;">
                    <b><%=tsa_txt_405%></b><br />
                     <%=tsa_txt_406%>: <%=meAnsatDato %> (<%=tsa_txt_005 %>: <%=datepart("ww", meAnsatDato, 2,2) %>) - <%=meOpsagtDato %> 
                        <br /><%=tsa_txt_407%>: <%=datepart("ww", sidstedagisidsteugeTjk, 2,2) %>
                        </div>

                    <%
                    end select


        else


		if cint(showAfsuge) = 1 then%>
		
			
		
				<%

                    
				'*** Man må gerne lukke uger frem i tiden 27012009 ***'
				'if (datepart("ww", tjekdag(7), 2, 2) <= datepart("ww", now, 2, 2) AND year(tjekdag(7)) <= year(now)) OR  year(tjekdag(7)) < year(now) then%>
				<input type="hidden" name="FM_mid" id="FM_mid" value="<%=usemrn%>">
				<input type="hidden" name="FM_afslutuge_sidstedag" id="FM_afslutuge_sidstedag" value="<%=sidstedagisidsteAfsluge%>"><!-- sidstedagisidsteAfsluge -->
     
                <%if datepart("ww", sidstedagisidsteuge, 2, 2) >= 1 then ' kan ikke afslutte uge 1? eller
                
                select case cint(SmiWeekOrMonth)
                case 2    
                %>
                <%end select %>
                
                <span style="font-size:14px;"><%=tsa_txt_409 %>:</span><br />
                <span style="font-size:14px; display:block; padding:2px; border:0px #999999 solid; width:400px; font-weight:bolder;">

                    <%


                       '*** MÅ TEAMLEDER/Bruger AFSLUTTE uge for medarb ***
                         maaAfslutteUge = 0
                         select case lto
                         case "tec", "esn"
                        
                                if session("mid") = usemrn OR level = 1 then
                                maaAfslutteUge = 1
                                else
                                maaAfslutteUge = 0
                                end if

                         case else

                            '** Er kriterie for antal timer opfyldt? 
                            if (cint(afslProcKri) = 1 AND afslutugekri <> 0) OR afslutugekri = 0 then
                            '**Udbygges evt. med TEAMLEDER
                            maaAfslutteUge = 1
                            else
                            maaAfslutteUge = 0
                            end if

                         end select


                        if cint(maaAfslutteUge) = 1 OR cint(level) = 1 then
                            FM_afslutugeDIS = ""
                        else
                            FM_afslutugeDIS = "DISABLED"
                        end if
                        '***************************************************


                        '********** Der findes herrelosetimer / timer uden match fra f.eks TT. 
                        '********** Admin må heller ikke afslutte
                        if (cint(afslutugekri) = 10) then 

                            FM_afslutugeDIS = "DISABLED"

                        end if
                    

                
                select case cint(SmiWeekOrMonth)
                case 0, 2

                    if ((cint(maaAfslutteUge) = 1 AND cint(afslutugekri) <> 10) OR cint(level) = 1) then
                    'AND lto <> "tec" AND lto <> "esn"  
                    'Må ikke kune afslutte alle uger når der er Kriterier opstillet på min. timer pr. uge
                    'Eller hvis man afslutter på månedsbasis (afslut alle måender til d.d funktion ikke klar)
                                       
                        if cint(level) = 1 OR SmiWeekOrMonth = 2 then%>
                        <input type="checkbox" name="FM_afslutuge" id="FM_afslutuge" value="1" onClick="showlukalleuger()" <%=FM_afslutugeDIS %>> 
                        <%else %>
                        <input type="checkbox" name="FM_afslutuge" id="FM_afslutuge" value="1" <%=FM_afslutugeDIS %>>
                        <%end if 

                    end if
                  
                case 1

                if ((cint(maaAfslutteUge) = 1 AND cint(afslutugekri) <> 10) OR cint(level) = 1) then '** Måned skal ikke kunne afslutte alle %>
                <input type="checkbox" name="FM_afslutuge" id="FM_afslutuge" value="1" <%=FM_afslutugeDIS %>>
                <%end if
                    
                end select %>

                

                   <!--
                  sidstedagisidsteAfsluge:  <%=sidstedagisidsteAfsluge %><br />
                   sidstedagisidsteuge: <%=sidstedagisidsteuge %><br />
                   ww: <%=datepart("ww", sidstedagisidsteuge, 2, 2) %>
                    -->
                  
               
               
                <%
               '**** Skriver periode / dag / uge der afsluttes     
                    
                select case cint(SmiWeekOrMonth) 
                case 0 'uge aflsutning %>
                <%=tsa_txt_005 %>: <%=datePart("ww", sidstedagisidsteAfsluge, 2,2) %>,
                   
                   <%if datePart("ww", sidstedagisidsteAfsluge, 2,2) = 53 then%> 
                   <%=datePart("yyyy", dateAdd("yyyy", -1, sidstedagisidsteAfsluge), 2,2) %> / <%=datePart("yyyy", sidstedagisidsteAfsluge, 2,2) %>  
                   <%else%>
                   <%=datePart("yyyy", sidstedagisidsteAfsluge, 2,2) %>
                   <%end if %>

                    
                <%if cint(SmiWeekOrMonth_HR) = 1 AND cint(SmiWeekOrMonth) = 0 then

                '*** Afslut MD nu! (ugeafslutning) funktion skal kun vises hvis det er den sidste uge i måneden og den strækker sig ind i næste måned.
                'sidstedagisidsteAfsluge = Sidstedag (Søndag) i den der skal afsluttes som den næste
                '*** Tjekker om måned allerede har været afsluttet. Vis da ikke knap igen.

                denneugeStarter = datePart("m", dateAdd("d", -7, sidstedagisidsteAfsluge), 2,2) 
                denneugeSlutter = datePart("m", sidstedagisidsteAfsluge, 2,2) 
                tjkHRsplitdenneUge = year(sidstedagisidsteAfsluge) & "-" & month(sidstedagisidsteAfsluge) & "-"& day(sidstedagisidsteAfsluge)

                splithrdenneUge = 0
                strSQLselHr = "SELECT uge, splithr FROM ugestatus WHERE WEEK(uge, 3) = WEEK('"& tjkHRsplitdenneUge &"', 3) AND mid = " & usemrn & " AND splithr = 1"
                    
                'Response.write strSQLselHr
                'Response.flush
                oRec6.open strSQLselHr, oConn, 3
                if not oRec6.EOF then

                    splithrdenneUge = oRec6("splithr")

                end if
                oRec6.close

                %>
                <%if cint(denneugeStarter) <> cint(denneugeSlutter) AND cint(splithrdenneUge) = 0 then %>
                <br /><span style="font-weight:normal;"><input type="checkbox" name="FM_afslutuge_hr" id="FM_afslutuge_hr" value="1"> 
                Complete month now <!--HR spiltweek -->
                <!-- Afslut indeværende uge nu! (pr. dd. ved månedsskift eller lønkørsel) -->
                </span><br />
                <%end if %>
                
                <%
                end if%>

                <%case 1 'MD %>
                <%=monthname(datePart("m", sidstedagisidsteAfsluge, 2,2)) %>, <%=datePart("yyyy", sidstedagisidsteAfsluge, 2,2) %>
                <%case 2 'Dag %>
                <%=weekdayname(weekday(sidstedagisidsteAfsluge)) & " d. " & sidstedagisidsteAfsluge %>
                <%end select %>

               
        
                    <%select case cint(SmiWeekOrMonth) 
                    case 0
                    %>

                    <%if datepart("ww", tjkDag7, 2 ,2) <> datepart("ww", sidstedagisidsteAfsluge, 2 ,2) AND thisfile <> "ugeseddel_2011.asp" then  %>
                    <!--(<a href="<%=menu2015lnk%>timereg_akt_2006.asp?showakt=1&strdag=<%=day(sidstedagisidsteAfsluge)%>&strmrd=<%=month(sidstedagisidsteAfsluge)%>&straar=<%=year(sidstedagisidsteAfsluge)%>" class="vmenu"><%=funk_txt_047 &" "& datePart("ww", sidstedagisidsteAfsluge, 2,2) %>..</a>)-->
                    <%end if %> 
                        
                    <%
                    case 2
                       %><%    
                    end select%>
                     </span>
        
        
                    <%
                  '***********************************************
                  '**** LUK ALLE UGER  
                  '*********************************************** 
                  %>
                <div id="lukalleuger" name="lukalleuger" style="position:relative; visibility:hidden; display:none; width:400px; padding-left:2px;">
                        
                <%
                select case cint(SmiWeekOrMonth) 
                case 0, 1    
                LastWeekSelected = dateAdd("d", -7, now)
                tomTxt = datePart("ww", LastWeekSelected, 2,2) &", "& datePart("yyyy", LastWeekSelected, 2,2)
                case 2
                LastWeekSelected = dateAdd("d", -1, now)
                tomTxt = formatdatetime(LastWeekSelected, 2) 
                end select %>
				<input type="checkbox" name="FM_alleuger" id="FM_alleuger" value="1">&nbsp;<span style="font-size:14px;"><%=tsa_txt_090 & ": "& tomTxt %> </span>
				</div>
        
        
                    
                    <%
                    '****************************************************
                    '************ SUBMIT ********************************
                    '****************************************************
                            
                                            
                    'if cint(SmiWeekOrMonth) = 0 then 'uge aflsutning %>
                    <%'=tsa_txt_089%>
                    <%'else %>
                    <%'=tsa_txt_431 &" "& monthname(datePart("m", sidstedagisidsteAfsluge, 2,2)) =datePart("yyyy", sidstedagisidsteAfsluge, 2,2) %>
                    <%'end if 
                

                    '***** Submit knap afslut  *****

                     select case cint(SmiWeekOrMonth) 
                     case 0
                     afslutTxtbtn = tsa_txt_431 '&" "& tsa_txt_345
                     case 1
                     afslutTxtbtn = tsa_txt_431 
                     case 2
                     afslutTxtbtn = tsa_txt_431 '&" "& left(tsa_txt_275, 3) 'Dag
                     end select

                 

                 if (cint(maaAfslutteUge) = 1 AND cint(afslutugekri) <> 10) OR cint(level) = 1 then%>
                   <br />
                 <input type="submit" class="btn btn-success btn-sm" value="<%=afslutTxtbtn &" "& tsa_txt_432 %> >>">
                <%else %>
                    <!--<span style="background-color:#CCCCCC; display:block; padding:10px; color:red; font-size:12px;"><%=tsa_txt_537 %></span>-->
                <%end if %>

               
                
         
                
                    <%end if %>
				
				
            
                   
               
				 
        
             <%
             '** Status på antal registrederede projekttimer i den pågældende uge / dag. Bruges ikke på md   
             'select case lto
             'case "tec", "esn"
             select case cint(SmiWeekOrMonth) 
             case 1 'Måned

             case else %>
               
               <%call smiley_uge_kriterie_opfyldt %>

                <!-- DER FINDES TIMER UDEN MATCH FRA F.eks TT -->
                <%if afslutugekri = 10 then %>
                <div style="color:#000000; background-color:#ffdfdf; padding:10px;">
                <b><%=funk_txt_048 %></b> <%=funk_txt_049 %><br /> <%=funk_txt_050 %><br /><br />
                <%=funk_txt_051 %> <b><%=afslugeDatoTimerudenMatch %></b>
                 </div>
                <%end if %>
            
            <%end select %>


                     
                     <span style="font-size:11px; color:#999999; display:block;">
                     <%
                           '******************************************************************************************
                           '*** Sidste rettidige afslutning   
                           '******************************************************************************************
                            select case cint(SmiWeekOrMonth) 
                            case 0
                              %>
                                <%=tsa_txt_410 %>: <b><%=left(weekdayname(weekday(cDateUge,1)), 3) &".  kl. "& left(formatdatetime(cDateUge, 3), 5) %> </b> <%=weekafslTxt %>
                             <%
                            case 1 'MD aflsutning %>
                        
                          
                        
                               <%if SmiantaldageCount = 8 then %>
                              <%=tsa_txt_410 &": "& tsa_txt_429 &" "& weekafslTxt %>
                              <%end if %>

                              <%if SmiantaldageCount = 9 then %>
                              <%=tsa_txt_410 &": "& tsa_txt_429 &" + 7 d. "& weekafslTxt %>
                              <%end if %>

                                <%if SmiantaldageCount = 10 then %>
                              <%=tsa_txt_410 &": "& tsa_txt_429 &" + 3 d. "& weekafslTxt %>
                              <%end if %>

                          <%case 2 %>
                            
                                <%=tsa_txt_410 %>: <%=funk_txt_052 %> <%=left(formatdatetime(cDateUge, 3), 5) %>
                          <%end select %>

                         </span>


		<%else 'uge/md er afsluttet
         
            
           

              
            select case cint(SmiWeekOrMonth) 
            case 0 'uge aflsutning 
            perTxt = tsa_txt_005 & " "& sidstedagiuge
            case 1
            perTxt = monthname(month(ugeNrAfsluttet)) & " er"
            case 2
            perTxt = weekdayname(weekday(ugeNrAfsluttet)) & ", d. "& ugeNrAfsluttet & " er"
            end select
            
            %>
        <br />
		<span style="font-size:14px; font-weight:bolder;"><%=perTxt%>&nbsp;<%=lcase(tsa_txt_093) %> <font color=green><i>V</i></font></span>
		<br><div style="font-size:11px; padding-top:5px;"><%=tsa_txt_093 %>&nbsp;<%=weekdayname(weekday(cdAfs))%>&nbsp;<%=tsa_txt_092 %>&nbsp;<%=formatdatetime(cdAfs, 2)%>&nbsp;<%=formatdatetime(cdAfs, 3)%> 
		(<%=tsa_txt_095 %>&nbsp;<%=datepart("ww", cdAfs, 2, 2)%>)</div>

        

		<%
            '*** hvis Kritere for ugen skal vises
            'smiley_uge_kriterie_opfyldt    
            
        end if%>


        <%end if %>
	
	
	</td>
	</tr>
</table> 
 </form>
<%end function%>




<%


public antalAfsDato, useDateSmileyTjkWeek, antalAfsl
function showsmiley(weekSelected, visning, usemrn, SmiWeekOrMonth)

 '***************************** Smiley ******************************
	%>
	<table border='0' cellspacing='0' cellpadding='0' width=100%>
	<tr bgcolor="#FFFFFF"><td style="padding:0px 0px 0px 0px;">
	<%
	
	'**** Sur Smiley overruler ******'
	antalAfsDato = 0

    'sætter useDateSmileyTjk = sidste uge
	if visning = 1 then 'sidste uge
    useDateSmileyTjk = dateadd("d", -7, weekSelected)
    useDateSmileyTjk2 = dateadd("d", -8, weekSelected) '14
	else
    useDateSmileyTjk = dateadd("d", -7, weekSelected)
    useDateSmileyTjk2 = dateadd("d", -8, weekSelected)
    end if


    call meStamdata(usemrn)
    call licensStartDato()

    if cDate(meAnsatDato) > cDate(licensstdato) then
    useListDato_meAnsatDato = meAnsatDato
    else
    useListDato_meAnsatDato = licensstdato
    end if

    
    if year(useListDato_meAnsatDato) >= year(useDateSmileyTjk) then
            
            'if year(meAnsatDato) >= year(now) then
            surDatoSQLSTART = year(useListDato_meAnsatDato)&"/"& month(useListDato_meAnsatDato)&"/"& day(useListDato_meAnsatDato)

            select case cint(SmiWeekOrMonth) 
            case 0 'then 'uge aflsutning
            useDateSmileyTjkWeek = dateDiff("ww", useListDato_meAnsatDato, now, 2, 2)
            case 1
            useDateSmileyTjkWeek = dateDiff("m", useListDato_meAnsatDato, now, 2, 2)
            case 2
            useDateSmileyTjkWeek = dateDiff("d", useListDato_meAnsatDato, now, 2, 2)
            end select
        
            'else
            'surDatoSQLSTART = year(now)&"/1/1"
            'useDateSmileyTjkWeek = dateDiff("ww", surDatoSQLSTART, now, 2, 2)
            'end if
    
     else
                    '** Første Torsdag så vi er sikker på at vi er inde i uge 1
                    surDatoSQLSTART = year(useDateSmileyTjk)&"/1/1"

                    select case cint(SmiWeekOrMonth) 
                    case 0 'then 'uge aflsutning
                    dayOfWeekThis = datePart("w",surDatoSQLSTART,2,2)

                    dThis = (dayOfWeekThis - 1)
                    firstdayOfFirstWeekThis = dateAdd("d", -dThis, surDatoSQLSTART)
                    surDatoSQLSTART = firstdayOfFirstWeekThis 'dateAdd("d", +3, firstdayOfFirstWeekThis)
                    surDatoSQLSTART = year(surDatoSQLSTART)&"/"& month(surDatoSQLSTART)&"/"& day(surDatoSQLSTART)

                    useDateSmileyTjkWeek = datepart("ww", useDateSmileyTjk, 2, 2)
                    
                    case 1
                    useDateSmileyTjkWeek = datepart("m", useDateSmileyTjk, 2, 2)

                    case 2
                    useDateSmileyTjkWeek = datepart("d", useDateSmileyTjk, 2, 2)
                    end select

    
    end if


    if cDate(meOpsagtDato) < cDate(useDateSmileyTjk2) then
       surDatoSQLEND = year(meOpsagtDato)&"/"& month(meOpsagtDato)&"/"& day(meOpsagtDato)
    else
       surDatoSQLEND = year(useDateSmileyTjk2)&"/"&month(useDateSmileyTjk2)&"/"&day(useDateSmileyTjk2)
    end if
	
	
    dim antalAfsDato, antalAfsl
	antalAfsDato = 0
    antalAfsl = 0

	strSQL3 = "SELECT COALESCE(count(*)) AS antalafsluttede FROM ugestatus "_
	&" WHERE mid = "& usemrn &" AND uge BETWEEN '"& surDatoSQLSTART &"' AND '"& surDatoSQLEND &"'"_
	&" GROUP BY mid ORDER BY afsluttet DESC"
    
    'if lto = "assurator" AND session("mid") = 14 then
	'Response.write strSQL3 & "<br><br>useDateSmileyTjkWeek: "& useDateSmileyTjkWeek & "<br>" & strSQL3 &"<br>"
	'Response.flush
    'end if
	
    'antalAfsDato = antalAfsDato * 1    
    oRec3.open strSQL3, oConn, 3 
	if not oRec3.EOF then

        if isNull(oRec3("antalafsluttede")) <> true then


        'antalAfsl = 0
        'if lto = "assurator" then
        'response.write "afl: #" & oRec3("antalafsluttede") & "#<br>"
        'antalAfsl = 0
        'antalAfsl = getLong(oRec3("antalafsluttede")) 
        'response.write "antalAfsl: #" & antalAfsl & "#<br>"
        'antalAfsl = antalAfsl/1
        'response.Flush

         antalAfsl = cdbl(oRec3("antalafsluttede")) 

        'if lto <> "assurator" OR (lto = "assurator" AND session("mid") = 14) then

        '    if (lto = "assurator" AND session("mid") = 14) then
        '    antalAfsl = cdbl(oRec3("antalafsluttede")) 
        '    antalAfsDato = antalAfsDato/1 + antalAfsl/1
        '    else
        '    antalAfsl = oRec3("antalafsluttede") 
        '    antalAfsDato = antalAfsDato/1 + antalAfsl/1
        '    end if
        
        'end if

         antalAfsDato = antalAfsDato/1 + antalAfsl/1

        'end if
        'antalAfsDato = 0

        'antalAfsl = oRec3("antalafs")
        'if antalAfsl <> 0 then
        'antalAfsDato = antalAfsDato + antalAfsl
        'end if

        'antalAfsDato = antalAfsDato

        'response.write antalAfsDato & "<br>" 
        'response.flush

        else
        antalAfsl = 0
        antalAfsDato = antalAfsDato
        end if

	'oRec3.movenext	
	end if
	oRec3.close  
	
	'if lto = "plan" AND session("mid") = 1 then
	'Response.write "lastAfsDatoWeek: " & antalAfsDato+1 & " >= "&  useDateSmileyTjkWeek
	'end if

	
	'** Glade ***'
	if (antalAfsDato+1) >= useDateSmileyTjkWeek then
	surSmil = 0
	else
	surSmil = 1
	end if
	
	smileyRnd = right(cint(Second(now)), 1)
	
	select case smileyRnd
	case 0, 1 
	smilVal = 1
	case 2, 3
	smilVal = 2
	case 4, 5
	smilVal = 3
	case 6, 7
	smilVal = 4
	case else
	smilVal = 5
	end select
		
		%>
       
		<h4 style="color:#000000;"><%=meNavn &" ["& meInit %>] <span style="color:#999999; font-size:11px;"> - <%=tsa_txt_404 %>: <%=formatdatetime(meAnsatDato, 2) %></span>
         <!--   
         <br />
		 <span style="color:#000000; font-size:11px;">
             <%select case cint(SmiWeekOrMonth)
             case 0 'uge aflsutning %>
             <%=tsa_txt_411 &" "& lcase(tsa_txt_005) &" "& datepart("ww", useDateSmileyTjk2, 2, 2) &" "& year(useDateSmileyTjk2)%>
             <%case 1 %>
             <%=tsa_txt_411 &" "& monthname(datepart("m", useDateSmileyTjk2, 2, 2)) &" "& year(useDateSmileyTjk2)%>
             <%case 2 %>
             <%=tsa_txt_411 &" "& useDateSmileyTjk2 %>
             <%end select %>
            

		 </span>
             -->

		</h4>
		<%	

        if cint(hidesmileyicon) = 0 then 'skjul Smiley Icons

		if cint(surSmil) = 1 then%>
				
				<b><%=tsa_txt_096 %></b><!--(du har afsluttet: <b><%=antalAfsDato %> / <%=useDateSmileyTjkWeek %></b> uger)--></td>
        
                <td align=right style="padding:2px 5px 0px 0px;">

                <%if cint(SmiWeekOrMonth) <> 2 then 'smiley_agg %>
                <a href="#" class="sA1_k" style="color:#999999;">X</a>
                <%end if %>
                </td></tr>
	            <tr><td colspan=2 style="padding-top:3px;">
                
				<img src="../ill/sur_<%=smilVal%>.gif" alt="" style="border:0px;"><br>
                
				
                    

                     <%select case cint(SmiWeekOrMonth) 
                     case 0 'uge aflsutning %>
                     <span style="color:red;"><%=tsa_txt_403 %>!
                    <%case 1 %>
			         <span style="color:red;"><%=tsa_txt_427 %>!
                    <%case 2 %>
                     <span style="color:red;"><%=tsa_txt_427 %>!
                    <%end select %>

                    <!--<br /><=tsa_txt_428 %>-->
                         </span>

				
				
				<%
		else
		
				
				%>
				<b><%=tsa_txt_098 %></b>
				</td><td align=right style="padding:2px 5px 0px 0px;">
                    <%if cint(SmiWeekOrMonth) <> 2 then 'smiley_agg %>
                    <a href="#" class="sA1_k" style="color:#999999;">X</a>
                    <%end if %>
				     </td></tr>
	            <tr><td colspan=2 style="padding-top:3px;">

                  

				<img src="../ill/gladsmil_<%=smilVal%>.gif" alt="" style="border:0px;"><br>
                   
				
				
				<%
		end if%>
		

                   
		<!--<font class=megetlillesort><br>Nb: Smileyordning kan slås til og fra i medarbejder-profilen.</font>-->
		
	
	 
	
	</td></tr>

        <%end if 'smiley icon %>

		</table>
<%end function%>



<%

'** Viser afsluttede uger ***

public expTxtsm, ugeMdNrTxtTopKri
function smileystatus(medarbid, visning, useYear)

    call smileyAfslutSettings()

    'response.Write "medarbid: " & medarbid
    'response.flush 

    if instr(medarbid, ",") <> 0 then
    medarbids = split(medarbid, ", ")
    medarbKri = " m.mid = 0 "

        for m = 0 TO UBOUND(medarbids)

           medarbKri = medarbKri & " OR m.mid = "& medarbids(m) 

        next

    else
        medarbKri = " m.mid = "& medarbid  '" mansat <> 2 AND mansat <> 3 " 
    end if

    'if cint(medarbid) <> 0 then
	'medarbKri = " m.mid = "& medarbid 
	'else
	
	'end if


  
	startdato = useYear&"/1/1"
	slutdato = useYear&"/12/31"

    startdatoBeregn = "1/1/"& useYear
	slutdatoBeregn = now

	
	
	dim glade, sure, mnavn, mnr, ugerafsluttet, sletafsl, mansatDato
	Redim glade(500), sure(500), mnavn(500), mnr(500), mansatDato(500)
	
	m = 100 '100
	v = 4200 '5200
	redim ugerafsluttet(m,v) 
	redim sletafsl(m,v) 
	redim yearaf(m,v)
	
	'Response.write medarbid
	
	
	
	strSQL = "SELECT m.mid AS mid, m.mnavn, m.mnr, m.ansatdato, m.init FROM medarbejdere m"_
	&" WHERE ("& medarbKri &") ORDER BY m.mid LIMIT 100"
	oRec.open strSQL, oConn, 3 
	
	'Response.write strSQL & "<hr>"
	ugeMdNrTxtLast = 0
	x = 0
	lastmid = 0
	while not oRec.EOF 
	
	'Response.write lastmid &"<>"& oRec("mid") &"<br>"
	
		x = x + 1
		lastmid = oRec("mid")
		'Redim preserve mnavn(x)
        'Redim preserve mansatDato(x)
		'Redim preserve mnr(x)
		'Redim preserve glade(x)
		'Redim preserve sure(x)
		
		
        mansatDato(x) = oRec("ansatdato")
		mnavn(x) = oRec("mnavn")
		mnr(x) = oRec("init") 'oRec("mnr")
		
        if media = "export" then
        expTxtsm = expTxtsm & mnavn(x) &";"& mnr(x) &";"& mansatDato(x) & ";"
		end if
		        
                    if datepart("yyyy", oRec("ansatdato"), 2,2) >= datepart("yyyy", startdato, 2,2) then
                    tjkAnsatDatoKri = "AND (uge >= '"& year(oRec("ansatdato")) &"/"& month(oRec("ansatdato"))&"/"& day(oRec("ansatdato")) &"')"
                    else
                    tjkAnsatDatoKri = ""
                    end if
        
					strSQL2 = "SELECT u.status, u.afsluttet, u.uge, u.id, u.ugegodkendt, u.splithr FROM ugestatus u WHERE u.mid = "& oRec("mid") &" AND (uge BETWEEN '"& startdato &"' AND '"& slutdato &"') "& tjkAnsatDatoKri &" ORDER BY uge" 
					'AND (WEEK(uge,1) BETWEEN 1 AND "& denneuge &" 
					
                    'if session("mid") = 1 then
					'Response.write strSQL2 &"<br>"
	                'end if				

					v = 0
                    lastWrtBlnk = 0
					oRec2.open strSQL2, oConn, 3 
					while not oRec2.EOF 
			
    
                    'if cint(SmiWeekOrMonth) = 0 then
                    'ugeMdNrTxt = datepart("ww", oRec2("uge"), 2, 2)
                    'ugeMdNrTxtTopKri = 52
                    'else
                    'ugeMdNrTxt = datepart("m", oRec2("uge"), 2, 2)
                    'ugeMdNrTxtTopKri = 12
                    'end if


                    select case cint(SmiWeekOrMonth)
                    case 0
                    ugeMdNrTxt = datepart("ww", oRec2("uge"), 2, 2)
                    ugeMdNrTxtTopKri = 52
                    case 1
                    ugeMdNrTxt = datepart("m", oRec2("uge"), 2, 2)
                    ugeMdNrTxtTopKri = 12
                    case 2
                    ugeMdNrTxt = formatdatetime(oRec2("uge"), 2)
                    ugeMdNrTxtTopKri = 365
                    end select
                    
   

                    if media <> "export" then
    
    		
					        if v = 0 AND ugeMdNrTxt = ugeMdNrTxtTopKri AND cint(SmiWeekOrMonth) = 0 then
					
					        else


                            '** Perioderode overskrift
                            if cint(SmiWeekOrMonth) = 2 then 'dag
                             
                             if ugeMdNrTxtLast <> datepart("m", oRec2("uge"), 2, 2) then
                                ugeMdNrTxt = "<br><u>"& monthname(month(oRec2("uge")), 0) & "</u><br>" & ugeMdNrTxt 
                             end if

                            end if

					
					                    if oRec2("status") = 2 then
					                    glade(x) = glade(x) + 1

                                        
                                                if print <> "j" then
                           
					        
					                                if level = 1 AND (thisfile <> "timereg_akt_2006" AND thisfile <> "logindhist_2011.asp" AND thisfile <> "ugeseddel_2011.asp") then
						                            ugerafsluttet(x,v) = "<a href='smileystatus.asp?func=slet&id="&oRec2("id")&"' class=vmenuglobal>"& ugeMdNrTxt &"</a>"
					                                else
					                                ugerafsluttet(x,v) = ugeMdNrTxt
					                                end if

                                                else

                                                    ugerafsluttet(x,v) = ugeMdNrTxt 

                                                end if
					        
					                    else
					                    sure(x) = sure(x) + 1

                                                 if print <> "j" then            

					                                if level = 1 AND (thisfile <> "timereg_akt_2006" AND thisfile <> "logindhist_2011.asp" AND thisfile <> "ugeseddel_2011.asp" AND thisfile <> "favorit.asp") then
						                            ugerafsluttet(x,v) = "<a href='smileystatus.asp?func=slet&id="&oRec2("id")&"' class=vmenualt>"& ugeMdNrTxt &"</a>" 
					                                else
					                                ugerafsluttet(x,v) = ugeMdNrTxt
					                                end if

                                                else

                                                    ugerafsluttet(x,v) = ugeMdNrTxt
                                        
                                                end if
					                    end if



                                        '** Er periode uge/md godkendt af leder 
                                        select case oRec2("ugegodkendt")
                                        case 0
                                        ugerafsluttet(x,v) = ugerafsluttet(x,v)
                                        case 1
                                        ugerafsluttet(x,v) = ugerafsluttet(x,v) & "&nbsp;<i style=""color:green;"">V</i>"
                                        case 2
                                        ugerafsluttet(x,v) = ugerafsluttet(x,v) & "&nbsp;<i style=""color:red;"">%</i>"
                                        end select


                                       '** HR SPlit uge afsluttet midt i uge ved månedsskift til at køre løn. TIA
                                       if cint(oRec2("splithr")) = 1 then
                                          ugerafsluttet(x,v) = ugerafsluttet(x,v) & "&nbsp;<i style=""color:#5582d2;"">(Month)</i>"
                                       end if

					      
					
					        end if

                    else
                                


                    select case cint(SmiWeekOrMonth)
                    case 0
                    ugeMdNrTxt = datepart("ww", oRec2("uge"), 2, 2)
                    ugeMdNrTxtTopKri = 52
                    case 1
                    ugeMdNrTxt = datepart("m", oRec2("uge"), 2, 2)
                    ugeMdNrTxtTopKri = 12
                    case 2
                    ugeMdNrTxt = formatdatetime(oRec2("uge"), 2)
                    ugeMdNrTxtTopKri = 365
                    end select
                
                    for p = 0 TO ugeMdNrTxtTopKri


	                if cint(p) = cint(ugeMdNrTxt) then 				
                    expTxtsm = expTxtsm & ugeMdNrTxt & ";"
                    lastWrtBlnk = ugeMdNrTxt
                    else
                    
                    if (cint(p) < cint(ugeMdNrTxt) AND v = 0) OR (cint(p) > cint(lastWrtBlnk) AND cint(p) < cint(ugeMdNrTxt)) then
                    expTxtsm = expTxtsm & ";"
                    lastWrtBlnk = p
                    end if
                    
                    end if

                    next

            
                    end if 'media


                    sletafsl(x,v) = oRec2("id")
					yearaf(x,v) = datepart("yyyy", oRec2("uge"), 2, 2)
					v = v + 1

          
                    select case cint(SmiWeekOrMonth)
                    case 0
                    ugeMdNrTxtLast = datepart("m", oRec2("uge"), 2, 2)
                    case 1
                    ugeMdNrTxtLast = datepart("m", oRec2("uge"), 2, 2)
                    case 2
                    ugeMdNrTxtLast = datepart("m", oRec2("uge"), 2, 2)
                    end select

                    response.flush
					
					oRec2.movenext
					wend
					oRec2.close 


     if media = "export" then			
     expTxtsm = expTxtsm & "xx99123sy#z"
     end if

	
	oRec.movenext
	wend
	oRec.close 
	
	antalx = x
	'Response.write "thisfile: "& thisfile
	
	%>
	

<%

if media <> "export" then
    
if thisfile <> "smileystatus.asp" then
twdt = "725"
else
twdt = "805"
end if

if thisfile <> "smileystatus.asp" then 'fra tiemreg. siden
dvdsp = "none"
dvwzb = "hidden"
else
dvdsp = ""
dvwzb = "visible"
end if
%>	

<div id="dv_ugeafslutninger" style="position:relative; display:<%=dvdsp%>; visibility:<%=dvwzb%>;">

<table border=0 cellspacing=0 cellpadding=0 width="95%">

<tr>
    <td>&nbsp;</td>
	<td width=400><!--<b><%=tsa_txt_101 %></b>-->&nbsp;</td>
	<td align=right style="font-size:11px;"><%=funk_txt_053 &"<br>"& funk_txt_054 %></td>
	<td align=right style="font-size:11px;"><%=funk_txt_053 &"<br>"& funk_txt_055 %></td>
	<td align=right style="font-size:11px;"><%=tsa_txt_104 %></td>
	<td valign=top style="padding:1px 5px 2px 10px;">
	<%if visning = 100 then %>
	<a href="#" id="sA2_k" class=red><b>[x]</b></a>
	<%else %>
	&nbsp;
	<%end if %>
	</td>
    <td>&nbsp;</td>
</tr>

	
	<%
	
	
	
	lastyear = 2000
	'Response.flush
	if x <> 0 then
			for x = 1 to antalx ' - 1
			
			select case right(x, 1)
	case 0,2,4,6,8
	trbg = "#FFFFFF"
	case else
	trbg = "#Eff3ff"
	end select


           if cint(SmiWeekOrMonth) = 0 then 'uger afslutninger
            diffIntervalSet = "ww"
            else
           diffIntervalSet = "m"
           end if

          if thisfile <> "smileystatus.asp" then
          
                if year(mansatDato(x)) = year(startdatoBeregn) then
                  antalUger = datediff(diffIntervalSet, mansatDato(x), slutdatoBeregn, 2,2)
                  else
                       if year(mansatDato(x)) < year(startdatoBeregn) then
                       antalUger = datediff(diffIntervalSet, startdatoBeregn, slutdatoBeregn, 2,2) 
                       else
                       antalUger = 0 'datediff("ww", startdatoBeregn, slutdatoBeregn, 2,2)
                       end if
                  end if
        
          else

                if year(mansatDato(x)) >= year(startdatoBeregn) then
                antalUger = datediff(diffIntervalSet, mansatDato(x), slutdatoBeregn, 2,2)
                else
                antalUger = datediff(diffIntervalSet, startdatoBeregn, slutdatoBeregn, 2,2) 
                end if


           end if


           if antalUger < 0 then
           antalUger = 0
           end if

			
			%>
			<tr bgcolor="<%=trbg %>">
				<td height=20>&nbsp;</td>
				<td><b><%=mnavn(x)%> [<%=mnr(x)%>]</b> <span style="color:#999999; font-size:9px;"> - <%=tsa_txt_404 %>: <%=formatdatetime(mansatDato(x), 2) %></span></td>
				<td align=right><b><%=glade(x)%></b></td>
				<td align=right><%=sure(x)%></td>
				<td align=right><%=glade(x)+sure(x)%> / <%=antalUger%></td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			</tr>
			<tr bgcolor="<%=trbg %>"><td colspan=7 style="padding:5px 5px 5px 5px;">
				
				<%
				for v = 0 to UBOUND(ugerafsluttet, 1)
				if len(ugerafsluttet(x,v)) <> 0 then%>
					
					<%if right(v, 1) = 0 AND v > 0 then%>
					<br>
					<%end if%>
					
					<%if lastyear <> yearaf(x,v) then %>
					<span style="color:#5582d2;"><b><%=yearaf(x,v)%></b></span>:
                    <%else %>
                    - 
					<%end if %>
					
				    <%=ugerafsluttet(x,v)%>
				 		
				<%	
				
				lastyear = yearaf(x,v)
				end if
				next
				%>
			</td></tr>
			
			<%next
	else%>
	<tr bgcolor="#ffffff">
		<td style="height:100px;">&nbsp;</td>
		<td colspan=5 style=padding-left:20px;><%=tsa_txt_106 %></td>
		<td>&nbsp;</td>
	</tr>
	<%end if%>
	
	<%if visning <> 1 then %>
	<tr bgcolor="#5C75AA">
		<td style="height:20px">&nbsp;</td>
		<td colspan=5>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	<%else %>
	<tr bgcolor="#FFFFFF">
		<td style="height:20px">&nbsp;</td>
		<td colspan=5>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	
	<%end if %>
	


    </table>
    </div>
    <%end if 'export %>


	<%end function%>
	
	
	
	
	<% 
    	


    '***** Viser PopUp med status på periode afslutning '*****'
	sub opdaterSmiley
	
	
	
    '*** Afslutter uge ****'
    if len(request("FM_afslutuge")) <> 0 OR request("FM_afslutuge_hr") then
		        
          
        'sidstedagisidsteuge

               

		        call rettidigafsl(request("FM_afslutuge_sidstedag"), 0) 'kontrolpanel settings, uge eller månedbasis, dag mv
		        
		     
		       
        		'if lto = "dencker" then
                'response.write "request(FM_afslutuge_sidstedag)" & request("FM_afslutuge_sidstedag")
                'Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;lto:" & lto & "<br>"& cDateAfs & "<br>"
        		'Response.flush
        		'Response.write "&nbsp;&nbsp;&nbsp;"& cdate(cDateUge) &" >= " & cdate(cDateAfs)
                'end if
        	
		       ' if cint(SmiWeekOrMonth) = 0 then 'uge aflsutning 
        		
                'response.write cdate(cDateUge) &" >="& cdate(cDateAfs) & "<br><br>"

		        if cdate(cDateUge) >= cdate(cDateAfs) then
		        intStatusAfs = 2 '** Afsluttet korrekt
		        smileysttxt = "<span style=""color:green;""> - "& funk_txt_056 &" <i>V</i> </span>"
                smileyImg = "gladsmil_2.gif"
		        else
		        intStatusAfs = 1 '** Afsluttet forsent
		        smileysttxt = "<span style=""color:red;""> - "& funk_txt_057 &"</span>"
                smileyImg = "sur_1.gif"
		        end if
		        
                '**Afslutter SKÆV uge ved månedsskift eller lønkørsel. TIA
                splithr = 0
                if len(request("FM_afslutuge_hr")) <> 0 then
                cDateUgeTilAfslutning = "1-"& month(cDateUgeTilAfslutning) &"-"& year(cDateUgeTilAfslutning)
                cDateUgeTilAfslutning = dateAdd("d", -1, cDateUgeTilAfslutning)
                'cDateUgeTilAfslutning = dateAdd("m", 1, cDateUgeTilAfslutning)
                cDateUgeTilAfslutning = year(cDateUgeTilAfslutning) & "-" & month(cDateUgeTilAfslutning) & "-" & day(cDateUgeTilAfslutning)
                splithr = 1
                end if
               
        	    'cDateUgeSQL = year(cDateUge) & "/" & month(cDateUge) &"/"& day(cDateUge)	
		        medarbejderid = request("FM_mid")
                call meStamdata(medarbejderid)


                '**** SLETTER (Opdaterer for revisionsspor) EVT HRsplit midlertidig godkendt WEEK ***
                '**** SPLIT HR bruger altid datoen fra den sidste uge i måneden. 
                '**** FULD uge godkendes igen fra WEEK approval
                strSQLafslutDelHRsplit = "UPDATE ugestatus SET uge = '2002-01-01' WHERE WEEK(uge, 3) = WEEK('"& cDateUgeTilAfslutning &"', 3) AND mid = "& medarbejderid &" AND splithr = 1" 
		        'Response.Write strSQLafslutDelHRsplit
		        'Response.flush
		        oConn.execute(strSQLafslutDelHRsplit)
                
        		
		        strSQLafslut = "INSERT INTO ugestatus (uge, afsluttet, mid, status, splithr) VALUES ('"& cDateUgeTilAfslutning &"', '"& cDateAfs &"', "& medarbejderid &", "& intStatusAfs &", "& splithr &")" 
		        'Response.Write strSQLafslut
		        'Response.flush
		        oConn.execute(strSQLafslut)
		       


                '**************************************************************************************
                '*** hvis afslutning på daglig basis er en fredag, afsluttes lørdag og søndag samtidigt 
                '**************************************************************************************
                if cint(SmiWeekOrMonth) = 2 then

                denneDag_tjkFredag = datepart("w", cDateUgeTilAfslutning, 2, 2)
                
                select case denneDag_tjkFredag
                case 5

                cDateUgeTilAfslutning_lordag = dateAdd("d", 1, cDateUgeTilAfslutning) 
                cDateUgeTilAfslutning_lordag = year(cDateUgeTilAfslutning_lordag) &"/"& month(cDateUgeTilAfslutning_lordag) &"/"& day(cDateUgeTilAfslutning_lordag) 

                strSQLafslut = "INSERT INTO ugestatus (uge, afsluttet, mid, status) VALUES ('"& cDateUgeTilAfslutning_lordag &"', '"& cDateAfs &"', "& medarbejderid &", "& intStatusAfs &")" 
		        oConn.execute(strSQLafslut)

                cDateUgeTilAfslutning_sondag = dateAdd("d", 2, cDateUgeTilAfslutning)
                cDateUgeTilAfslutning_sondag = year(cDateUgeTilAfslutning_sondag) &"/"& month(cDateUgeTilAfslutning_sondag) &"/"& day(cDateUgeTilAfslutning_sondag)

                strSQLafslut = "INSERT INTO ugestatus (uge, afsluttet, mid, status) VALUES ('"& cDateUgeTilAfslutning_sondag &"', '"& cDateAfs &"', "& medarbejderid &", "& intStatusAfs &")" 
		        oConn.execute(strSQLafslut)

                end select

                end if




                'response.end
		        call autogktimer_fn()

		        if cint(autogktimer) = 1 OR cint(autogktimer) = 2 then

                select case cint(autogktimer)
                case 1
                godkendtstatus = 1
                case 2
                godkendtstatus = 3
                end select
		        
		        stDatoGKtimer = cDateUgeTilAfslutning
		        slDatoGKtimer = dateadd("d", -6, cDateUgeTilAfslutning)
		        
		        slDatoGKtimer = year(slDatoGKtimer) &"/"& month(slDatoGKtimer) &"/"& day(slDatoGKtimer)
		        
		        
		        strSQLGKtimer = "UPDATE timer SET godkendtstatus = "& godkendtstatus &", godkendtstatusaf = '"& session("user") &"' WHERE"_
		        &" tmnr = "& medarbejderid &" AND tdato BETWEEN '"& slDatoGKtimer &"' AND '"& stDatoGKtimer &"' AND godkendtstatus <> 1 AND overfort = 0"
		        
		        oConn.execute(strSQLGKtimer)
		       
        		end if
        		
        		
			        '*** Afslut alle uger frem til dd. ****
			        if len(request("FM_alleuger")) <> 0 then
			        afslutalle = request("FM_alleuger")
			        else
			        afslutalle = 0
			        end if
        			         

                    'Response.Write "afslutalle: " & afslutalle
                    'Response.end

                    '***********************************************************
			        '** Kun afslut på ugebasis kan afslutte ALLE uger i året.
                    '***********************************************************
			        if afslutalle = "1" then
        					
                   

                           
					        detteAar = year(request("FM_afslutuge_sidstedag"))
                            
                            sidsteUgeAfsl = request("FM_afslutuge_sidstedag")
                
                            select case cint(SmiWeekOrMonth) 
                            case 0,1
                            nuMinusSyv = dateAdd("d", -7, now)
                            case 2
                            nuMinusSyv = dateAdd("d", 0, now)
                            end select                    

                           
                            select case cint(SmiWeekOrMonth) 
                            case 0,1
                            dayOfWeekThis = datePart("w",now,2,2)
        
                            dThis = (dayOfWeekThis - 1)
                            firstdayOfFirstWeekThis = dateAdd("d", -dThis, now)
                            firstThursdayOfFirstWeekThis = dateAdd("d", +6, firstdayOfFirstWeekThis)
            
                            sidstedagisidsteAfsluge = firstThursdayOfFirstWeekThis


                            denneugeLast = dateAdd("d", -7, sidstedagisidsteAfsluge)
					        denneuge = datepart("ww", denneugeLast, 2, 2) 'cDateUge
					        denneugeOpr = denneuge

                            case 2
                            denneugeLast = dateAdd("d", -1, nuMinusSyv)
                            end select
        					
                            'response.write "nuMinusSyv: "& nuMinusSyv & "<br>"
                            select case cint(SmiWeekOrMonth) 
                            case 0,1
                            PeriodeInterVal = dateDiff("ww", sidsteUgeAfsl, nuMinusSyv, 2,2)
                            case 2
                            PeriodeInterVal = dateDiff("d", sidsteUgeAfsl, nuMinusSyv, 2,2)
                            end select

                            'response.write "denneugeLast = "& denneugeLast &", PeriodeInterVal: "& PeriodeInterVal & "<br>"
                            
                               

					        for u = 0 to PeriodeInterVal 'denneuge 
        					
                            select case cint(SmiWeekOrMonth) 
                            case 0,1
					        nyDateUge = dateadd("d", (u*(-7)), denneugeLast)
                            case 2
                            nyDateUge = dateadd("d", (u*(-1)), denneugeLast)
                            end select

                            select case cint(SmiWeekOrMonth) 
                            case 0,1
					        denneuge = datepart("ww", nyDateUge, 2, 2)
					        detteaar = datepart("yyyy", nyDateUge, 2,3)

                            perSqlKri = "YEAR(u.uge) = '"& detteaar &"' AND  WEEK(u.uge, 3) = '"& denneuge &"'"
                            case 2
                        
                            denneDato = Year(nyDateUge) &"/"& month(nyDateUge) & "/" & day(nyDateUge)
                            perSqlKri = "u.uge = '"& denneDato &"'"
                            end select
        					
					        ugefindes = 0
					        strSQL2 = "SELECT u.id, u.status, u.afsluttet, u.uge FROM ugestatus u WHERE "& perSqlKri &" AND u.mid = "& medarbejderid
					        'Response.write strSQL2 & "<br>"
        					
					        oRec2.open strSQL2, oConn, 3 
					        if not oRec2.EOF then
					        ugefindes = 1 
					        'Response.write oRec2("id") &" :: "& oRec2("uge") & " - "& datepart("ww", oRec2("uge"), 2, 2) & " - ok!<br>"
				            end if
					        oRec2.close 
        					
					        'Response.Write "SmiWeekOrMonth: "& SmiWeekOrMonth &" "&  denneuge &" < "& denneugeOpr &" AND ugefindes: "& ugefindes &" nyDateUge:"&nyDateUge&" uge: "& datepart("ww", nyDateUge, 2,2) &" meAnsatDato: "& meAnsatDato &" ansatuge: "& datepart("ww", meAnsatDato, 2,2) &"<br>"
                            'response.flush 
                      
        					'cint(denneuge) <= cint(denneugeOpr) AND 

                            
                             'Response.Write "<br>strSQL2: "& strSQL2 &" <br>afslutalle: " & afslutalle & " PeriodeInterVal: " & PeriodeInterVal & "SmiWeekOrMonth: " & SmiWeekOrMonth & " ugefindes: " & ugefindes & " nyDateUge: "& nyDateUge &" >= "& meAnsatDato & " " & datepart("ww", nyDateUge, 2,2) &">="& datepart("ww", meAnsatDato, 2,2)
                             'Response.end

                            select case cint(SmiWeekOrMonth) 
                            case 0,1

						            if cint(ugefindes) = 0 AND ((datepart("ww", nyDateUge, 2,2) >= datepart("ww", meAnsatDato, 2,2) AND datepart("yyyy", nyDateUge, 2,2) = datepart("yyyy", meAnsatDato, 2,2)) OR datepart("yyyy", nyDateUge, 2,2) > datepart("yyyy", meAnsatDato, 2,2)) then
							        
                                    'Response.write "OK"

                                    sqlDateUge = year(nyDateUge)&"/"&month(nyDateUge)&"/"&day(nyDateUge)
        		
							        intStatusAfs = 1 '** Afsl. forsent
							        strSQLafslut = "INSERT INTO ugestatus (uge, afsluttet, mid, status) VALUES ('"& sqlDateUge &"', '"& cDateAfs &"', "& medarbejderid &", "& intStatusAfs &")" 
							        oConn.execute(strSQLafslut)
                                    
                                            if cint(autogktimer) = 1 then 'Teamleder godkender UGE. Kan kun sætte godkendtstatus = 1
		        
		                                    stDatoGKtimer = sqlDateUge
		                                    slDatoGKtimer = dateadd("d", -6, sqlDateUge)
                            		        
		                                    slDatoGKtimer = year(slDatoGKtimer) &"/"& month(slDatoGKtimer) &"/"& day(slDatoGKtimer)
                            		        
                            		        
		                                    strSQLGKtimer = "UPDATE timer SET godkendtstatus = "& godkendtstatus &", godkendtstatusaf = '"& session("user") &"' WHERE"_
		                                    &" tmnr = "& medarbejderid &" AND tdato BETWEEN '"& slDatoGKtimer &"' AND '"& stDatoGKtimer &"' AND godkendtstatus <> 1 AND overfort = 0"
                            		        
                                            'response.write strSQLGKtimer
                                            'response.flush                    		        

		                                    oConn.execute(strSQLGKtimer)
		                                 
        		                            end if
							                    
							        
							        end if
							        
							case 2

                                    if cint(ugefindes) = 0 AND cDate(nyDateUge) >= cDate(meAnsatDato) then
							        
                                    sqlDateUge = denneDato
        		
							        intStatusAfs = 1 '** Afsl. forsent
							        strSQLafslut = "INSERT INTO ugestatus (uge, afsluttet, mid, status) VALUES ('"& sqlDateUge &"', '"& cDateAfs &"', "& medarbejderid &", "& intStatusAfs &")" 

                                    'response.write strSQLafslut &"<br>"
							        oConn.execute(strSQLafslut)
                                    
                                            if cint(autogktimer) = 1 then 'Teamleder godkender UGE. Kan kun sætte godkendtstatus = 1
		        
		                                    stDatoGKtimer = sqlDateUge
		                                    
                            		        strSQLGKtimer = "UPDATE timer SET godkendtstatus = "& godkendtstatus &", godkendtstatusaf = '"& session("user") &"' WHERE"_
		                                    &" tmnr = "& medarbejderid &" AND tdato = '"& slDatoGKtimer &"' AND godkendtstatus <> 1 AND overfort = 0"
                            		        
                            		        
		                                    oConn.execute(strSQLGKtimer)

		                                 
        		                            end if
							                    
							        
							        end if

                                

                            end select
							        
        					
					        next
        					
        					     'Response.write "<br>SLUT"
        					     'Response.end
        					
			        end if
	
	    end if
	    
	end sub 
	
	
	
	
	sub afslutMsgTxt
        %>
          <label style="color:#999999;">
          <%=funk_txt_058 %><br /><br />
          <%=funk_txt_059 %></label>

        <%
    end sub
	
    



    '***************************************************************************
    '**** Er uge afsluttet??? ******'
    '**** Tjekker om den aktuelle uge man har klikket på er afsluttet / lukket
	'***************************************************************************

    public showAfsugeTxt
    public showAfsugeTxt_tot, ugegodkendtTxt_tot, btnstyle, ugeNrStatus

    public ugeNrAfsluttet, showAfsuge, cdAfs, ugegodkendt, ugegodkendtaf, ugegodkendtTxt, ugegodkendtdt, showAfsugeVisAfsluttetpaaGodkendUgesedler, splithr, ugeAfsluttetAfMedarb
    function erugeAfslutte(sm_aar, sm_sidsteugedag, sm_mid, SmiWeekOrMonth, erugeAfslutte_do)
            
          
             'call smileyAfslutSettings()
             splithr = 0
             ugegodkendt = 0
             ugeAfsluttet = 0
             ugeAfsluttetAfMedarb = 0
             showAfsuge = 1
             ugeNrAfsluttet = "1-1-2044"
             ugeNrStatus = 0


             'if session("mid") = 1 then
             'response.write "<br>Thisfile: "& thisfile  &" HER: sm_sidsteugedag: "& sm_sidsteugedag & " SmiWeekOrMonth: "& SmiWeekOrMonth & "sm_aar: "& sm_aar
             'end if    

             select case cint(SmiWeekOrMonth)
             case 0 
             periodeKri = "WEEK(uge, 3)"
             case 1
             periodeKri = "MONTH(uge)"
             case 2
             periodeKri = "uge"
             end select
             
            
            if cint(SmiWeekOrMonth) <> 2 then

                
                if cint(sm_sidsteugedag) = 53 then 'tjekker for gammel SQL server 3.23
                strSQLgl323 = " OR " & periodeKri &" = 0" 
                sqlDatoKri = "("& periodeKri &" = "& sm_sidsteugedag &" "& strSQLgl323 &")"
                else
                strSQLgl323 = ""
                sqlDatoKri = periodeKri &" = "& sm_sidsteugedag 
                end if

                if cint(sm_sidsteugedag) >= 52 then
                sqlYearKri = " AND (YEAR(uge) = "& sm_aar &" OR YEAR(uge) = "& sm_aar+1 &")"
                else
                sqlYearKri = " AND YEAR(uge) = "& sm_aar &""    
                end if

            else
            
                sqlDatoKri = periodeKri &" = '"& sm_sidsteugedag & "'" 
                sqlYearKri = " AND YEAR(uge) = "& sm_aar &""   

            end if


            'if thisfile <> "week_norm_2010.asp" then
            'splithrKri = "AND splithr <> 1"
            'else
            splithrKri = ""
            'end if

            strSQLafslut = "SELECT status, afsluttet, uge, ugegodkendt, ugegodkendtaf, ugegodkendtTxt, ugegodkendtdt, splithr FROM ugestatus WHERE "_
            &" "& sqlDatoKri &" "& sqlYearKri &" AND mid = "& sm_mid &" "& splithrKri  
		    
            'if session("mid") = 21 then
            'Response.write "HER:<br>"& strSQLafslut & "<br><br>//" & sm_sidsteugedag & "//"& sldatoSQL & thisfile
		    'Response.flush
            'end if
            
            oRec3.open strSQLafslut, oConn, 3 
		    if not oRec3.EOF then
    			
                splithr = oRec3("splithr")
			    
                showAfsuge = 0
                ugeAfsluttetAfMedarb = 1
                cdAfs = oRec3("afsluttet")
			    ugeNrAfsluttet = oRec3("uge")
                ugeNrStatus = oRec3("status")

                ugegodkendt = oRec3("ugegodkendt")
                ugegodkendtaf = oRec3("ugegodkendtaf")
                ugegodkendtTxt = oRec3("ugegodkendtTxt")
                ugegodkendtdt = oRec3("ugegodkendtdt")
                
                
    		
		    end if
		    oRec3.close 
            
            'if session("mid") = 1 then
            'Response.write "<br>ugeNrStatus: "& ugeNrStatus &" ugeNrAfsluttet : "& ugeNrAfsluttet  & " ugegodkendt: "& ugegodkendt & " splithr: "& splithr  '& "//" & sm_sidsteugedag & "//"& sldatoSQL
		    
           'Response.flush
            'end if

            showAfsugeVisAfsluttetpaaGodkendUgesedler = showAfsuge


            '** Altid åben på dagsafslutning
            if (level = 1 AND cint(SmiWeekOrMonth) = 2) OR cint(splithr) = 1 then
            showAfsuge = 1
            end if

          

    end function



    public sidsteUgenrAfsl, sidsteUgenrAfslFundet, slip_smiley_agg_lukper
     function afsluttedeUger(sm_aar, sm_mid, sm_dothis)

        
            call meStamdata(sm_mid)
            call licensStartDato()
        
            slip_smiley_agg_lukper = 0
            sidsteUgenrAfslFundet = 0

            if cDate(meAnsatDato) > cDate(licensstdato) then
            sidsteUgenrAfsl = meAnsatDato 
            else
            sidsteUgenrAfsl = licensstdato '"1-1-2002"
            end if        

            select case SmiWeekOrMonth
            case 2
            sidsteUgenrAfsl_w = datePart("w", sidsteUgenrAfsl, 2, 2)

            if cint(sidsteUgenrAfsl_w) = 6 then 'lørdag
            sidsteUgenrAfsl = dateAdd("d", 2, sidsteUgenrAfsl)
            end if

            if cint(sidsteUgenrAfsl_w) = 7 then 'søndag
            sidsteUgenrAfsl = dateAdd("d", 1, sidsteUgenrAfsl)
            end if

            end select

            'if cDate(sidsteUgenrAfsl) > cDate("1-1-"&year(now)) then
            'sidsteUgenrAfsl = sidsteUgenrAfsl
            'else
            'sidsteUgenrAfsl = "1-1-"&year(now)
            'end if


            ansatDatoSQL = year(meAnsatDato) &"/"& month(meAnsatDato) &"/"& day(meAnsatDato)
            licensstDatoSQL = year(licensstdato) &"/"& month(licensstdato) &"/"& day(licensstdato) 
            
            strSQLafslut = "SELECT status, afsluttet, uge, ugegodkendt, ugegodkendtaf, ugegodkendtTxt, ugegodkendtdt FROM ugestatus WHERE "_
            &" ((YEAR(uge) = "& sm_aar &" AND uge >= '"& ansatDatoSQL &"' AND uge >= '"& licensstDatoSQL &"') OR (YEAR(uge) < "& sm_aar &")) AND (mid = "& sm_mid & ") AND splithr = 0 ORDER BY uge DESC LIMIT 1"
		    
            
            'response.write "strSQLafslut: <br>" & strSQLafslut
            'response.Flush
         
            oRec3.open strSQLafslut, oConn, 3 
		    while not oRec3.EOF 
    			
            sidsteUgenrAfslFundet = 1
            sidsteUgenrAfsl = oRec3("uge")
			   
    		
            oRec3.movenext
		    wend
		    oRec3.close 
            

            'Response.Write "afsluttedeUger_fn: sidsteUgenrAfsl: " & sidsteUgenrAfsl

            '*** Skal der lukkes ned for timereg. cint(smiley_agg) = 1 
            if cint(sm_dothis) = 1 then

            if cint(smiley_agg_lukhard) = 1 then

            slip_dd = now
            

            select case SmiWeekOrMonth
            case 0
            slip_limit = 3 'ophøj til kontrolpanel
            slip_dd_diff = dateDiff("ww", sidsteUgenrAfsl, slip_dd, 2,2)
            case 1
            slip_limit = 6
            slip_dd_diff = dateDiff("m", sidsteUgenrAfsl, slip_dd, 2,2)
            case 2

            if datePart("w", slip_dd, 2, 2) = 1 then
            slip_limit = 3 'På en mandag må der gerne være 3 dage
            else
            slip_limit = 1
            end if

            '**tilføjer en dag hvis igår var en halligdag
            slip_dd_dayminusone = dateAdd("d", -1, slip_dd)
            call helligdage(slip_dd_dayminusone, 0, lto)
            if erhellig = 1 then
            slip_limit = slip_limit + 1
            end if

            '*** Efter Jul / Nytår, Påske
            slip_dd_newyear = "05-01-"& year(slip_dd)
            'Response.write "slip_dd: " & slip_dd & " slip_dd_newyear: " & slip_dd_newyear
            if cDate(slip_dd) >= cDate("1-1-"& year(slip_dd)) AND (cDate(slip_dd) <= cDate(slip_dd_newyear)) then
            slip_limit = 4
            end if

            if helligdagnavn = "2.Påskedag" then 'HVIS igår var 2 påskedag
            slip_limit = 6
            end if

            
            slip_dd_diff = dateDiff("d", sidsteUgenrAfsl, slip_dd, 2,2)
            end select


            select case SmiWeekOrMonth
            case 0,1
               
                if (cint(slip_dd_diff) > cint(slip_limit)) then
                slip_smiley_agg_lukper = 1
                end if

            case 2

                if (cint(slip_dd_diff) > cint(slip_limit)) then
                slip_smiley_agg_lukper = 1
                end if

                'slip_smiley_agg_lukper = 0
                'slip_dd_diff = 1
                'slip_limit = 1
                if (cint(slip_dd_diff) = cint(slip_limit)) then

                    if formatdatetime(SmiantaldageCountClock, 3) <= formatdatetime(now, 3) then '* Luk kun efter kl.
                    slip_smiley_agg_lukper = 1
                    end if

                end if

            end select
            

            if session("rettigheder") = 1 then
            slip_smiley_agg_lukper = 0 
            end if

            'Response.write "slip_dd_diff: " & slip_dd_diff & " SmiantaldageCount: "& SmiantaldageCount &" <br> SmiantaldageCountClock: " & SmiantaldageCountClock & " <= " & formatdatetime(now, 3)

            'SmiantaldageCount, 'SmiantaldageCountClock
            'select case slip_dd_diff


            '**** SLIP og skal periode lukkes for yderligere indtastning???
             
            
                'response.write "<br>smiley_agg: " & smiley_agg
		            '**** Tjekker hvornår periode skal være afsluttet
                    '**** Hvis afslut kriterie slået til
                    'slip_smiley_agg_lukper = 0 ' Dencker Testmode
                    if cint(slip_smiley_agg_lukper) = 1 then
                    %>        
                    <div style="position:relative; background-color:#cccccc; height:2000px; width:2000px; top:20px; left:-80px; padding:40px; z-index:20000;">
                    
                        <%=now %>
                        <h4><%=funk_txt_063 %></h4>
                    
                        <%=funk_txt_064 %> <%=session("user") %><br /><br />
                        <%=funk_txt_065 %>
                        <br /><br />
                        <%=funk_txt_066 %>

                        <br /><br />
                        <%=funk_txt_067 %> <a href="mailto:ad@dencker.net?subject=TimeOut bruger har ikke fået afsluttet sin dag til tiden." target="_blank"><%=funk_txt_068 %></a> <%=funk_txt_069 %>

                       
                        <!--  <h4>Du kan ikke længere registrere timer</h4> 
                            Det skyldes højst sandsynligt en af følgende årsager:<br /><br />
                            A) Du har for mange uafsluttede uger. <br /><br />
                            B) Din bruger er blevet de-aktiviteret. <br /><br />
                            C) Din fratrædelsesdato er overskreddet. <br /><br />&nbsp;

                            -->

                    </div>
                    <%
                    '*** LUKKER HÅRDT FOR INDTASTNING
                    'Response.end
                    end if
                end if 'cint(smiley_agg) = 1 
                end if 'sm_dothis

            
    end function


function godekendugeseddel(thisfile, godkenderID, medarbid, stDatoSQL)

        call smileyAfslutSettings()
       
        'gkweekfundet = 0
        ugegodkendtdtnow = year(now) & "/" & month(now) & "/" & day(now)
        denneuge = datepart("ww", stDatoSQL, 2,2)
	    detteaar = datepart("yyyy", stDatoSQL, 2,2)
        dettemd = datepart("m", stDatoSQL, 2,2)

        lastDayInWeek = dateAdd("d", 6, stDatoSQL)

        '** tjekker om de er uge 52/53 ell. 1 og om der skal korrigeres på år.
        if cint(denneuge) = 1 AND (datepart("yyyy", stDatoSQL, 2,2) <> datepart("yyyy", lastDayInWeek, 2,2)) OR (denneuge = 53 AND cint(SmiWeekOrMonth) = 0) then 'datepart("d", stDatoSQL, 2,2) > 29 then  
        '29 er den tidligst mulige dato uge 1 kan starte på, hvor det er akturelt at skifte år fordi søndagen der afslutter året ligger i det nye år.
        detteaarAdd = dateAdd("yyyy", 1, stDatoSQL)
        detteaar = datepart("yyyy", detteaarAdd, 2,2)

             if cint(SmiWeekOrMonth) = 1 then
             dettemdAdd = dateAdd("m", 1, stDatoSQL)
             dettemd = datepart("m", dettemdAdd, 2,2)
             end if

        end if


                
             if cint(SmiWeekOrMonth) <> 2 then
 
                if cint(denneuge) >= 52 then
                sqlYearKri = " AND (YEAR(uge) = "& detteaar &" OR YEAR(uge) = "& detteaar+1 &")"
                else
                sqlYearKri = " AND YEAR(uge) = "& detteaar &""    
                end if

             else
                   
                sqlYearKri = " AND YEAR(uge) = "& detteaar &""   

             end if

         ugstatusSQL = "SELECT id, mid FROM ugestatus AS u WHERE mid = "& medarbid & " " & sqlYearKri

        if cint(SmiWeekOrMonth) = 0 then
        

            if cint(denneuge) = 53 then
            ugstatusSQL = ugstatusSQL & " AND (WEEK(u.uge, 3) = '"& denneuge &"' OR WEEK(u.uge, 3) = '0') "
            else
            ugstatusSQL = ugstatusSQL & " AND WEEK(u.uge, 3) = '"& denneuge &"'"
            end if


        else
        ugstatusSQL = ugstatusSQL & " AND MONTH(u.uge) = '"& dettemd &"'"
        end if

        'if session("mid") = 1 then
        'Response.Write ugstatusSQL &"<br><br>mon:"& stDatoSQL &"<br>son: "& lastDayInWeek & " denneuge : "& denneuge 
        'Response.end
        'end if


        oRec6.open ugstatusSQL, oConn, 3
        if not oRec6.EOF then
        
            strSQLup = "UPDATE ugestatus SET ugegodkendt = 1, ugegodkendtaf = "& godkenderID &", ugegodkendtdt = '"& ugegodkendtdtnow &"' WHERE id = "& oRec6("id") 
	        'Response.Write strSQLup
            'Response.end
            oConn.execute(strSQLup)

        'gkweekfundet = 1
        end if
        oRec6.close

        
      

end function


function afviseugeseddel(thisfile, afsenderMid, modtagerMid, varTjDatoUS_man, varTjDatoUS_son, txt)

                            '*** Henter afsender **
				            strSQL = "SELECT mnavn, email FROM medarbejdere"_
				            &" WHERE mid = "& afsenderMid
				            oRec.open strSQL, oConn, 3
            				
				            if not oRec.EOF then
            				
				            afsNavn = oRec("mnavn")
				            afsEmail = oRec("email")
            				
				            end if
				            oRec.close

                             '*** Henter modtager **
				            strSQL = "SELECT mnavn, email FROM medarbejdere"_
				            &" WHERE mid = "& modtagerMid
				            oRec.open strSQL, oConn, 3
            				
				            if not oRec.EOF then
            				
				            modtNavn = oRec("mnavn")
				            modtEmail = oRec("email")
            				
				            end if
				            oRec.close


                            '*** Afvis ugeseddel ænder ikke statusd på timerne ***'
                            'strSQLup = "UPDATE timer SET godkendtstatus = 2, godkendtstatusaf = '"& afsNavn &"' WHERE tmnr = "& modtagerMid & " AND tdato BETWEEN '"& varTjDatoUS_man &"' AND '" & varTjDatoUS_son & "'" 
	                        'oConn.execute(strSQLup)

                            call smileyAfslutSettings()
        
                            ugegodkendtTxt = txt
                              
                              ugegodkendtdtnow = year(now) & "/" & month(now) & "/" & day(now)
                              denneuge = datepart("ww", varTjDatoUS_man, 2,2)
	                          detteaar = datepart("yyyy", varTjDatoUS_man, 2,2)

                              dettemd = datepart("m", varTjDatoUS_man, 2,2)


                              lastDayInWeek = dateAdd("d", 6, varTjDatoUS_man)

                                '** tjekker om de er uge 52/53 ell. 1 og om der skal korrigeres på år.
                                if cint(denneuge) = 1 AND (datepart("yyyy", varTjDatoUS_man, 2,2) <> datepart("yyyy", lastDayInWeek, 2,2)) OR (denneuge = 53 AND cint(SmiWeekOrMonth) = 0) then 'datepart("d", stDatoSQL, 2,2) > 29 then  
                                '29 er den tidligst mulige dato uge 1 kan starte på, hvor det er akturelt at skifte år fordi søndagen der afslutter året ligger i det nye år.
                                detteaarAdd = dateAdd("yyyy", 1, varTjDatoUS_man)
                                detteaar = datepart("yyyy", detteaarAdd, 2,2)

                                     if cint(SmiWeekOrMonth) = 1 then
                                     dettemdAdd = dateAdd("m", 1, varTjDatoUS_man)
                                     dettemd = datepart("m", dettemdAdd, 2,2)
                                     end if

                                end if
                            

                            '** TIA hvis autogk = 2 then SLET uge else opdater
                            strSQLupUgeseddel = "UPDATE ugestatus SET ugegodkendt = 2, ugegodkendtaf = "& afsenderMid &", ugegodkendtdt = '"& ugegodkendtdtnow &"', ugegodkendtTxt = '"& ugegodkendtTxt &"'"
                        
                        
                            strSQLupUgeseddelWhere = " WHERE mid = "& modtagerMid 
                            
                            
                                if cint(SmiWeekOrMonth) <> 2 then
 
                                    if cint(denneuge) >= 52 then
                                    sqlYearKri = " AND (YEAR(uge) = "& detteaar &" OR YEAR(uge) = "& detteaar+1 &")"
                                    else
                                    sqlYearKri = " AND YEAR(uge) = "& detteaar &""    
                                    end if

                                 else
                   
                                    sqlYearKri = " AND YEAR(uge) = "& detteaar &""   

                                 end if


                            if cint(SmiWeekOrMonth) = 0 then
                    
                                    if cint(denneuge) = 53 then
                                    strSQLupUgeseddelWhere = strSQLupUgeseddelWhere &" "& sqlYearKri &" AND (WEEK(uge, 3) = '"& denneuge &"' OR WEEK(uge, 3) = '0')"
                                    else
                                    strSQLupUgeseddelWhere = strSQLupUgeseddelWhere &" "& sqlYearKri &" AND WEEK(uge, 3) = '"& denneuge &"'"
                                    end if
                            else
                            strSQLupUgeseddelWhere = strSQLupUgeseddelWhere &" "& sqlYearKri &" AND  MONTH(uge) = '"& dettemd &"'"
                            end if

                            strSQLupUgeseddel = strSQLupUgeseddel & strSQLupUgeseddelWhere
	                        'Response.Write strSQLupUgeseddel
                            'Response.end
                            oConn.execute(strSQLupUgeseddel)



                            '*** Nulstiller tentative ****'
                        
                            call autogktimer_fn()
                            call smileyAfslutSettings()
                            call ersmileyaktiv()

                                    if (cint(autogktimer) = 1 OR cint(autogktimer) = 2) AND cint(smilaktiv) = 1 AND cint(autogk) = 2 then
                                    thisUeId = 0
                                    strSQLugest = "SELECT id FROM ugestatus "& strSQLupUgeseddelWhere
                                    oRec6.open strSQLugest, oConn, 3
                                    if not oRec6.EOF then
                                
                                    thisUeId = oRec6("id")

                                    end if
                                    oRec6.close


                                    strSQLupUgeseddelDel = "DELETE from ugestatus WHERe id = "& thisUeId
                                    oConn.execute(strSQLupUgeseddelDel)

                                    call nulstilTentative(autogktimer, thisUeId)

                
		                  end if

	    
	    'Response.Write strSQLup
	    'Response.end
	    

          if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\"&thisfile AND request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\to_2015\"&thisfile  then

                                    

                                     '************************** TEST ***********************************
                                    'Sender notifikations mail
		                            'Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		                            ' Sætter Charsettet til ISO-8859-1
		                            'Mailer.CharSet = 2
		                            'Mailer.FromName = "TimeOut performance service "
		                            'Mailer.FromAddress = "timeout@outzource.dk"
		                            'Mailer.RemoteHost = "webout.smtp.nu" '"webmail.abusiness.dk" '"pasmtp.tele.dk"
						            'Mailer.AddRecipient "TimeOut Support", "sk@outzource.dk"
		                           
            						
                             
            						    
						            'Mailer.Subject = "Loadtid alert "& lto &", "& thisfile
		                            'strBody = "Hej TimeOut Support" & vbCrLf & vbCrLf
                                   
            		
		                            'Mailer.BodyText = strBody
            		
		                            'Mailer.sendmail()
		                            'Set Mailer = Nothing

                                    '************************************************************************************

                                    
                           
                                    if cint(SmiWeekOrMonth) = 0 then
                                    thisWeek = datepart("ww", varTjDatoUS_man, 2, 2) 
                                    thisWeekTxt = "Din ugeseddel uge:"
                                    else
                                    thisWeek = monthname(datepart("m", varTjDatoUS_man, 2, 2)) 
                                    thisWeekTxt = ""
                                    end if   

					  	            ''Sender notifikations mail
		                            'Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		                            '' Sætter Charsettet til ISO-8859-1
		                            'Mailer.CharSet = 2
		                            'Mailer.FromName = "TimeOut | " & afsNavn 
		                            'Mailer.FromAddress = afsEmail
		                            'Mailer.RemoteHost = "webout.smtp.nu" '"webmail.abusiness.dk" '"pasmtp.tele.dk"
						            'Mailer.AddRecipient "" & afsNavn & "", "" & afsEmail & ""
		                            'Mailer.AddRecipient "" & modtNavn & "", "" & modtEmail & ""
            						
                                    'if cint(SmiWeekOrMonth) = 0 then
                                    'Mailer.Subject = "Din ugeseddel uge: "& thisWeek &" - er afvist"
                                    'else
                                    'Mailer.Subject = "Periode: "& thisWeek &" - er afvist"
                                    'end if

		                            'strBody = "Hej " & modtNavn &  vbCrLf & vbCrLf
                                    'strBody = strBody & thisWeekTxt &" "& thisWeek &" - er afvist" & vbCrLf & vbCrLf
                                    'strBody = strBody &"Begrundelse: " & vbCrLf 
						            'strBody = strBody & txt & vbCrLf & vbCrLf
		                            'strBody = strBody &"Med venlig hilsen" & vbCrLf
		                            'strBody = strBody & afsNavn & vbCrLf & vbCrLf
            		                
            		
            		
		                            'Mailer.BodyText = strBody
            		
		                            'Mailer.sendmail()
		                            'Set Mailer = Nothing


                        
                                    Set myMail=CreateObject("CDO.Message")
                                    if cint(SmiWeekOrMonth) = 0 then
                                    myMail.Subject = "Timeout - Din ugeseddel uge: "& thisWeek &" - er afvist"
                                    else
                                    myMail.Subject = "Timeout - Periode: "& thisWeek &" - er afvist"
                                    end if
                                    
                                    myMail.From = "timeout_no_reply@outzource.dk"
				                     

                                    'myMail.To=strEmail
                                    if len(trim(modtEmail)) <> 0 then
                                    myMail.To= ""& modtNavn &"<"& modtEmail &">"
                                    end if

                                    strBody = "Hej " & modtNavn &  vbCrLf & vbCrLf
                                    strBody = strBody & thisWeekTxt &" "& thisWeek &" - er afvist" & vbCrLf & vbCrLf
                                    strBody = strBody &"Begrundelse: " & vbCrLf 
						            strBody = strBody & txt & vbCrLf & vbCrLf
		                            strBody = strBody &"Med venlig hilsen" & vbCrLf
		                            strBody = strBody & afsNavn & ", "& afsEmail & vbCrLf & vbCrLf                       

                                    myMail.TextBody= strBody

                        
                                    myMail.Configuration.Fields.Item _
                                    ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
                                    'Name or IP of remote SMTP server
                                   
                                    if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
                                       smtpServer = "webout.smtp.nu"
                                    else
                                       smtpServer = "formrelay.rackhosting.com" 
                                    end if
                    
                                    myMail.Configuration.Fields.Item _
                                    ("http://schemas.microsoft.com/cdo/configuration/smtpserver")= smtpServer

                                    'Server port
                                    myMail.Configuration.Fields.Item _
                                    ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25
                                    myMail.Configuration.Fields.Update
                    
                                    if len(trim(modtEmail)) <> 0 then
                                    myMail.Send
                                    end if
                                    set myMail=nothing


				            
                            end if' x


end function


function afslutugereminder(thisfile, afsenderMid, modtagerMid, varTjDatoUS_man, varTjDatoUS_son, txt)

                            '*** Henter afsender **
				            strSQL = "SELECT mnavn, email FROM medarbejdere"_
				            &" WHERE mid = "& afsenderMid
				            oRec.open strSQL, oConn, 3
            				
				            if not oRec.EOF then
            				
				            afsNavn = oRec("mnavn")
				            afsEmail = oRec("email")
            				
				            end if
				            oRec.close

                             '*** Henter modtager **
				            strSQL = "SELECT mnavn, email FROM medarbejdere"_
				            &" WHERE mid = "& modtagerMid
				            oRec.open strSQL, oConn, 3
            				
				            if not oRec.EOF then
            				
				            modtNavn = oRec("mnavn")
				            modtEmail = oRec("email")
            				
				            end if
				            oRec.close


                          
	    

          if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\"&thisfile then

                                    call smileyAfslutSettings()
                           
                                    if cint(SmiWeekOrMonth) = 0 then
                                    thisWeek = datepart("ww", varTjDatoUS_man, 2, 2) 
                                    thisWeekTxt = ""& funk_txt_072 &":"
                                    thisWeekSbj = funk_txt_073
                                    else
                                    thisWeek = monthname(datepart("m", varTjDatoUS_man, 2, 2)) 
                                    thisWeekTxt = ""& funk_txt_074 &":"
                                    thisWeekSbj = funk_txt_075
                                    end if   

					  	            'Sender notifikations mail
		                            'Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		                            ' Sætter Charsettet til ISO-8859-1
		                            'Mailer.CharSet = 2
		                            'Mailer.FromName = "TimeOut | " & afsNavn 
		                            'Mailer.FromAddress = afsEmail
		                            'Mailer.RemoteHost = "webout.smtp.nu" '"webmail.abusiness.dk" '"pasmtp.tele.dk"
						            '''Mailer.AddRecipient "" & afsNavn & "", "" & afsEmail & ""
		                            'Mailer.AddRecipient "" & modtNavn & "", "" & modtEmail & ""
            						
                                    
                                    'Mailer.Subject = "Du har endnu ikke afsluttet " '& lcase(thisWeekTxt) &" "& thisWeek 
		                            'strBody = "Hej " & modtNavn &  vbCrLf & vbCrLf
                                    'strBody = strBody &" - " & thisWeekTxt &" "& thisWeek &" - er endnu ikke afsluttet, husk at få den afsluttet snarest." & vbCrLf & vbCrLf
                                   
		                            'strBody = strBody &"Med venlig hilsen" & vbCrLf
		                            'strBody = strBody & afsNavn & vbCrLf & vbCrLf
            		                
            		
            		
		                            'Mailer.BodyText = strBody
            		
		                            'Mailer.sendmail()
		                            'Set Mailer = Nothing


                    
                                    Set myMail=CreateObject("CDO.Message")
                                    myMail.Subject="TimeOut - "& funk_txt_076 &" "& thisWeekSbj & ": "& thisWeek
                                    myMail.From = "timeout_no_reply@outzource.dk"
				                     

                                    'myMail.To=strEmail
                                    if len(trim(modtEmail)) <> 0 then
                                    myMail.To= ""& modtNavn &"<"& modtEmail &">"
                                    end if

                                    strBody = "Hej " & modtNavn &  vbCrLf & vbCrLf
                                    strBody = strBody & thisWeekTxt &" "& thisWeek &" - "& funk_txt_078 &"" & vbCrLf & vbCrLf
                                   
		                            strBody = strBody &""& funk_txt_079 &"" & vbCrLf
		                            strBody = strBody & afsNavn &", "& afsEmail & vbCrLf & vbCrLf                       

                                    myMail.TextBody= strBody

                        
                                    myMail.Configuration.Fields.Item _
                                    ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
                                    'Name or IP of remote SMTP server
                                   
                                    if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
                                       smtpServer = "webout.smtp.nu"
                                    else
                                       smtpServer = "formrelay.rackhosting.com" 
                                    end if
                    
                                    myMail.Configuration.Fields.Item _
                                    ("http://schemas.microsoft.com/cdo/configuration/smtpserver")= smtpServer

                                    'Server port
                                    myMail.Configuration.Fields.Item _
                                    ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25
                                    myMail.Configuration.Fields.Update
                    
                                    if len(trim(modtEmail)) <> 0 then
                                    myMail.Send
                                    end if
                                    set myMail=nothing


                                   
				            
          end if' x


end function



sub smiley_uge_kriterie_opfyldt
%>


                        <br /><br />
                   <%if cint(afslProcKri) = 1 then
                   sm_bdcol = "#DCF5BD"
                   okSymbol = "<span style=""color:green; font-size:16px;""><i>V</i></span>"
                   else
                   sm_bdcol = "mistyrose"
                   okSymbol = "<span style=""color:red; font-size:16px;""><i>`/.</i></span>"
                   end if 
                       
                       
                 if thisfile = "logindhist_2011.asp" OR thisfile = "timereg_akt_2006" then
                  lpx = 0
                 else
                  lpx = 15
                 end if %>


                <div class="row" style="padding-left:<%=lpx%>px;">
                <div class="col-md-10" style="background-color:<%=sm_bdcol%>; padding:10px 10px 10px 10px; font-size:12px;">
                    
                    <table border="0" width="100%"><tr><td style="font-size:12px;">
                    
                    <br />

                    <%=tsa_txt_398 %>:<br /><b><%=totTimerWeek%></b>&nbsp; 
                    
                    

                    <%if afslutugekri = 2 then %>
                    <%=tsa_txt_399 %>
                    <%end if %>

                    <%select case cint(SmiWeekOrMonth)
                    case 0, 1%>
                    <%=" "&tsa_txt_140 %> / <%=afslutugeBasisKri %> = <b><%=afslProc %> %</b> <%=tsa_txt_095 %> <b><%=datePart("ww", tjkTimeriUgeDtTxt, 2,2) %></b>
                       
                    <%case 2 %>
                     <%=" "&tsa_txt_140 %> / <%=afslutugeBasisKri %>
                    
                    <%'select case lto
                     'case "dencker", "intranet - local"
                       Response.Write " ("& kommegaa60 & ") "
                     'case else
                        
                     'end select %>

                     = <b><%=afslProc %> %</b> <%=okSymbol %>
                    
                    <%end select %>
                    
                    <%if afslutugekri <> 0 then %>
                      <br /> <span style="color:#999999; font-size:11px;">
                        <%=funk_txt_070 %>: <%=afslutugekri_proc %> %  

                           <%select case cint(SmiWeekOrMonth)
                        case 2%> 
                        &nbsp;<%=weekdayname(weekday(tjkTimeriUgeDtTxt), 0, 1) %> d. <%=formatdatetime(tjkTimeriUgeDtTxt, 2) %> 
                        <%end select %>
                        
                    </span>
                    <%else %>
                         <%select case cint(SmiWeekOrMonth)
                        case 2%> <br /> <span style="color:#999999; font-size:11px;">
                        <%=weekdayname(weekday(tjkTimeriUgeDtTxt), 0, 1) %> d. <%=formatdatetime(tjkTimeriUgeDtTxt, 2) %> </span>
                        <%end select %>

                    <%end if %>

                        


                     <%select case cint(SmiWeekOrMonth)
                        case 0,1%>
                        <br /><span style="color:#000000; font-size:9px;">(<%=left(weekdayname(weekday(formatdatetime(tjkTimeriUgeDt, 2))), 3) &". "& formatdatetime(tjkTimeriUgeDt, 2)%> - <%= left(weekdayname(weekday(formatdatetime(dateAdd("d", 6, tjkTimeriUgeDt), 2))), 3) &". "&formatdatetime(dateAdd("d", 6, tjkTimeriUgeDt), 2) %>)</span>
                        <%end select %>
                     

                    <%select case cint(SmiWeekOrMonth)
                    case 0
                    pxFaktor = 1.5
                    case 1
                    pxFaktor = 0.25 'bruges ikke PÅ MD basis
                    case 2
                    pxFaktor = 10
                    end select      
                        
                        
                    pxVal = formatnumber(totTimerWeek*pxFaktor, 0)
                    pxVal2 = formatnumber(afslutugeBasisKri*pxFaktor, 0)

                    if pxVal > 100 then
                    pxVal = 100
                    end if
                        
                    if pxVal2 > 100 then
                    pxVal2 = 100
                    end if %>

                     </td><td align="right" valign="bottom" style="width:20px;">
                          <div style="height:<%=pxVal%>px; width:20px; background-color:green; padding:2px; vertical-align:bottom; font-size:9px; color:#ffffff;"><%=totTimerWeek %></div>
                    </td><td align="right" valign="bottom" style="width:20px;">
                    <div style="height:<%=pxVal2%>px; width:20px; background-color:#cccccc; padding:2px; vertical-align:bottom; font-size:9px;"><%=afslutugeBasisKri %></div>

                    

                        

                        </td></tr></table>
                   </div>
                    </div><!--</row>-->


<%
end sub








sub xsmiley_statusTxt


                         select case lto
                         case "tec", "esn"
                         case else

                               if cint(afslProcKri) = 1 then
                               sm_bdcol = "#DCF5BD"
                               else
                                sm_bdcol = "mistyrose"
                               end if %>



                            <div style="color:#000000; background-color:<%=sm_bdcol%>; padding:10px; border:0px <%=sm_bdcol%> solid;">

                                     <%select case cint(SmiWeekOrMonth)
                                     case 0, 1  %>


                                     <%=tsa_txt_398 & ":<br> <b>"& totTimerWeek %></b> 
                
                                        <%if afslutugekri = 2 then %>
                                        <%=" " & tsa_txt_399 %>
                                        <%end if %>
                
                                    <%=" "&tsa_txt_140 %> / <%=afslutugeBasisKri %> = <b><%=afslProc %> %</b> <%=tsa_txt_095 %> <b><%=datePart("ww", tjkTimeriUgeDtTxt, 2,2) %></b>
                                    <br />(<%=left(weekdayname(weekday(formatdatetime(tjkTimeriUgeDt, 2))), 3) &". "& formatdatetime(tjkTimeriUgeDt, 2)%> - <%= left(weekdayname(weekday(formatdatetime(dateAdd("d", 6, tjkTimeriUgeDt), 2))), 3) &". "&formatdatetime(dateAdd("d", 6, tjkTimeriUgeDt), 2) %>)
               
                                     <%=tsa_txt_400 %>: <b><%=afslutugekri_proc %> %</b>  <%=tsa_txt_401%>.
                          
                       
                                    <%case 2 %>
                                     <%=tsa_txt_398 & ":<br> <b>"& totTimerWeek %></b> 

                                    <%if afslutugekri = 2 then %>
                                        <%=" " & tsa_txt_399 %>
                                        <%end if %>

                                       <%=" "&tsa_txt_140 %> / <%=afslutugeBasisKri %> = <b><%=afslProc %> %</b>

                                    <%end select

                             %></div><%

                            %>

                
                            <%
                            end select


end sub



function nulstilTentative(autogktimer, id)


                select case cint(autogktimer)
                case 1
                godkendtstatus = 1
                case 2
                godkendtstatus = 3
                end select

                periodeSt = "2002-01-01"
                periodeMid = 0
                strSQLthisPeriode = "SELECT uge, mid FROM ugestatus WHERE id = "& id
                oRec6.open strSQLthisPeriode, oConn, 3
                if not oRec6.EOF then

                 periodeSt = oRec6("uge")
                 periodeMid = oRec6("mid")

                end if
                oRec6.close

		        

		        stDatoGKtimer = periodeSt
		        slDatoGKtimer = dateadd("d", -6, stDatoGKtimer)
		        
		        stDatoGKtimer = year(stDatoGKtimer) &"/"& month(stDatoGKtimer) &"/"& day(stDatoGKtimer)
		        slDatoGKtimer = year(slDatoGKtimer) &"/"& month(slDatoGKtimer) &"/"& day(slDatoGKtimer)
		        
		        strSQLGKtimer = "UPDATE timer SET godkendtstatus = 0, godkendtstatusaf = '' WHERE"_
		        &" tmnr = "& periodeMid &" AND tdato BETWEEN '"& slDatoGKtimer &"' AND '"& stDatoGKtimer &"' AND godkendtstatus = 3 AND overfort = 0" '** Kun tentative
		        
		        oConn.execute(strSQLGKtimer)

                'Response.write "autogktimer: "& autogktimer &" SmiWeekOrMonth: " & SmiWeekOrMonth & "<br>"& strSQLGKtimer
                'Response.end


end function
    
%>