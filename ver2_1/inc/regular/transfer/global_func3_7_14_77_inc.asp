<%
'call treg_medarb_afstem_saldo 'cls_afstem
	  

       '7 =  Aftem. M�ned afstem_tot
       '77 = Afstem dag/dag EXPAND 7
       '14 = Godkend m�ned/uge WEEEK_REAL_NORM

      level = session("rettigheder")
      call smileyAfslutSettings()

     if show = 77 then
      '**** finder loginds p� de valgta datoer ***'
      strlogin = ""
      strloginPauser = 0
      strSQLloginds = "SELECT login, logud, stempelurindstilling, minutter FROM login_historik WHERE dato = '"& startDato & "' AND mid = "& intMid 
      
    'Response.write strSQLloginds 
    'Response.flush
    
    
      oRec6.open strSQLloginds, oConn, 3
      while not oRec6.EOF

    
     
    if oRec6("stempelurindstilling") <> "-1" then

      if strlogin <> "" then
     strlogin = strlogin & "<br>"
     end if

            if isNull(oRec6("login")) <> true AND isNull(oRec6("logud")) <> true then
             strlogin = strlogin & left(formatdatetime(oRec6("login"), 3), 5) & " - "& left(formatdatetime(oRec6("logud"), 3), 5)
            end if 
     else
     strloginPauser = strloginPauser - oRec6("minutter")
     end if  


      oRec6.movenext 
      wend
      oRec6.close
        


     '**** Henter kommentarer p� dagen *******'
    kommentarTxt = ""
    strSQLkom = " SELECT timerkom FROM timer WHERE tmnr = "& intMid &" AND tdato = '"& startDato & "' AND tfaktim > 10 AND timerkom IS NOT NULL " 'kun frav�rstyper
    
    'Response.write strSQLkom
    'Response.flush
    
     oRec6.open strSQLkom, oConn, 3
      while not oRec6.EOF


    if (len(trim(oRec6("timerkom")))) > 2 AND ISNull(oRec6("timerkom")) <> true then
    call htmlparseCSV(oRec6("timerkom"))

    kommentarTxt = kommentarTxt & htmlparseCSVtxt & " - "

    end if

    oRec6.movenext 
      wend
      oRec6.close



     end if


      '**** Beregninger til min/maks kriterie *****
	   ltimerStDato = year(stDato) &"/"& month(stDato) &"/"& day(stDato)
	   call fLonTimerPer(ltimerStDato, dageDiff, 3, intMid)
	   
       select case datepart("w", ltimerStDato, 2, 2)
       case 1
       varTjDatoUS_man_use = dateAdd("d", 0, stDato)
       case 2
       varTjDatoUS_man_use = dateAdd("d", -1, stDato)
       case 3
       varTjDatoUS_man_use = dateAdd("d", -2, stDato)
       case 4
       varTjDatoUS_man_use = dateAdd("d", -3, stDato)
       case 5
       varTjDatoUS_man_use = dateAdd("d", -4, stDato)
       case 6
       varTjDatoUS_man_use = dateAdd("d", -5, stDato)
       case 7
       varTjDatoUS_man_use = dateAdd("d", -6, stDato)
       end select
       
	   
       if len(trim(fradragTimer(x))) <> 0 then
        fradragTimer(x) = fradragTimer(x)
        else 
        fradragTimer(x) = 0
        end if

     
       ltimerKorFrad = 0
       'if cint(showkgtil) = 1 then 'skal beregnes selvom kolionner ikke vises fysisk
	   ltimerKorFrad = ((totalTimerPer100/60) + (fradragTimer(x)))
       'else
       'ltimerKorFrad = (totalTimerPer100/60)
       'end if

	   normtime_lontime = -((normTimer(x) - (ltimerKorFrad)) * 60)
	   
	   if len(trim(normtime_lontime)) <> 0 then
	   normtime_lontime = normtime_lontime
	   else
	   normtime_lontime = 0
	   end if 
	   
	   balRealLontimer = realTimer(x) - (ltimerKorFrad)
	   if len(trim(balRealLontimer)) <> 0 then
	   balRealLontimer = balRealLontimer
	   else
	   balRealLontimer = 0
	   end if
	   
	   balRealNormtimer = ((realTimer(x)+(korrektionReal(x))) - normTimer(x))
	   if len(trim(balRealNormtimer)) <> 0 then
	   balRealNormtimer = balRealNormtimer
	   else
	   balRealNormtimer = 0
	   end if
	   
       select case lto
       case "wk", "mi", "wk_no", "akelius"
       balRealNormtimerAkk = balRealNormtimerAkk + (balRealNormtimer) - afspTimerTim(x)
       case else
       balRealNormtimerAkk = balRealNormtimerAkk + (balRealNormtimer) '+ korrektionReal(x)
       end select

       normtime_lontimeAkk = normtime_lontimeAkk + (normtime_lontime)
        
	   showthisMedarb = 1 
	   
	   if visning = 14 then 'week_real_norm
	   '** ekstra s�gekriterier ***'
            
                'Response.write "startDato: "& startDato &" intMid: "& intMid & "wth: "& wth
                if cint(SmiWeekOrMonth) = 0 then
                sidsteDag = year(dateadd("d", 6, startDato)) & "-"& month(dateadd("d", 6, startDato)) & "-"& day(dateadd("d", 6, startDato))
	            sidsteDagKri = datepart("ww", sidsteDag, 2, 2)
                else
                sidsteDag = startDato
	            sidsteDagKri = datepart("m", sidsteDag, 2, 2)
                end if        
        
        
                'AND (cint(lastMidtjk) <> cint(intMid) OR 
                if (cint(SmiWeekOrMonth) = 0) _
                OR (cint(SmiWeekOrMonth) = 1 AND datepart("m", slutDatoLastm_A, 2,2) <> lastMth) then 'cint(wth) = 0 OR 
               
                '**  OR ((cint(SmiWeekOrMonth) = 1 AND datepart("m", slutDatoLastm_A, 2,2) <> lastMth))
               
    
                if cint(SmiWeekOrMonth) = 0 then 'UgeAflsutning
                startDatoTor = dateAdd("d", 3, startDato) 'altid inde i uge, vistigt ved �rsskifte
        
                         if datepart("ww", startDatoTor,2,2) = 53 then
                            startDatoTor = dateAdd("d", 3, startDatoTor) 'ind i n�ste �r
                         end if

                else 'm�nedsafslutning
                startDatoTor = startDato 
                end if


                'response.write "startDatoTor: "& startDatoTor &" sidsteDagKri: "& sidsteDagKri &" slutDatoLastm_A " & startDato & "<br>lastMth: " & lastMth & "<br>lastMidtjk: "& lastMidtjk  'slutDatoLastm_A

              
                
                '*** Tjekker om uge er afsluttet af medarbejder. Viser CHK boks og Godkendtstatus
                call erugeAfslutte(datepart("yyyy", startDatoTor,2,2), sidsteDagKri, intMid, SmiWeekOrMonth, 0) 
                
                lastMidtjk = intMid
                end if

	            call ekstrasogKri(useSogKri, useSogKriAfs, useSogKriGk, moreorless,timeKri,saldoKri,visning)

                

	   end if
	   
	   
    
    
        if cint(showthisMedarb) = 1 then
	            

	           if len(trim(bgc)) <> 0 then
	           bgc = bgc
	           else
	           bgc = 0
	           end if
	   
	          
	   
	           bgc = bgc + 1

               
               if visning = 77 then
               bgthis = "#FFFFFF"
               
               if (datepart("w", startDato,1,2) = 1 OR datepart("w", startDato,1,2) = 7) then
               bgthis = "#EFf3ff"
               end if

               else

                select case right(bgc, 1)
	           case 0,2,4,6,8
	           bgthis = "#FFFFFF"
	           case else
	           bgthis = "#F7F7F7"
	           end select
               
               end if

              
	   
	   if media <> "export" then%>
	    <!-- st -->
        
	    <tr bgcolor="<%=bgthis %>">
	 <%end if
	 
	 strEksportTxt = strEksportTxt & "xx99123sy#z" & meNavn & ";" & meNr & ";"& meInit & ";"


        if visning = 14 then 'Uge / M�nedsafslutning
        '*** projektgrupper / afdelinger ****'
        strSQLprogrp = "SELECT ProjektgruppeId, MedarbejderId, p.navn AS pgrpnavn FROM progrupperelationer AS pr "_
        &" LEFT JOIN projektgrupper AS p ON (p.id = ProjektgruppeId) WHERE MedarbejderId = " & intMid & " AND ProjektgruppeId <> 10 ORDER BY p.navn"
	   
        progrpNavneTxt = ""
        p = 0
        oRec6.open strSQLprogrp, oConn, 3
        while not oRec6.EOF 
         
         if p = 0 then
         progrpNavneTxt = oRec6("pgrpnavn") 
         else
         progrpNavneTxt = progrpNavneTxt & ", "& oRec6("pgrpnavn")
         end if

        p = p + 1
        oRec6.movenext
        wend
        oRec6.close


         strEksportTxt = strEksportTxt & Chr(34) & progrpNavneTxt & Chr(34) &";"

        end if
	   
	    select case visning 
        case 7 
             
            if media <> "export" then%>
	        <td style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" class=lille>
            <%if media <> "print" then %>
            <a href="afstem_tot.asp?usemrn=<%=intMid %>&show=77&varTjDatoUS_man=<%=startDato %>&varTjDatoUS_son=<%=startDato%>&yuse=<%=yuse%>" class=rmenu><%=monthname(datepart("m", startDato,2,2)) &" "& year(startDato) %></a>
            <%else %>
            <b><%=monthname(datepart("m", startDato,2,2)) &" "& year(startDato) %></b>
            <%end if %>
            </td>
	        <%
            end if

	    strEksportTxt = strEksportTxt & monthname(datepart("m", startDato,2,2)) &" "& year(startDato) &";"
        case 77 
        
            if media <> "export" then%>
            
                <td style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap; padding-right:3px;" class=lille align=right>
            <%if lastWeek <> datePart("ww", startDato, 2,2) then  %>
                    <span style="color:#999999; font-size:8px;">uge: <%=datePart("ww", startDato, 2,2) %> </span>
                    <%
                        lastWeek = datePart("ww", startDato, 2,2) 
                        end if %>
                </td>


	        <td style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap; padding-left:3px; padding-right:3px;" class=lille align=right>
      
             <%if media <> "print" then %>
            <a href="../to_2015/ugeseddel_2011.asp?usemrn=<%=intMid %>&varTjDatoUS_man=<%=varTjDatoUS_man_use%>&nomenu=1" class=rmenu target="_blank"><%=left(weekdayname(datepart("w", startDato,1,2)), 2) %> d. <%=formatdatetime(startDato, 2) %></a>
                <!-- weekpage_2010.asp?medarbid=<%=intMid %>&st_dato=<%=day(varTjDatoUS_man_use)&"/"&month(varTjDatoUS_man_use)&"/"&year(varTjDatoUS_man_use) %>&func=us -->
            <%else %>
            <b><%=left(weekdayname(datepart("w", startDato,1,2)), 2) %> d. <%=formatdatetime(startDato, 2) %></b>
            <%end if %>
            </td>
	        <%
            end if

	        strEksportTxt = strEksportTxt & formatdatetime(startDato, 2) & ";"
	    
	    case else
        
        if media <> "export" then%>
	    <td style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" class=lille>
	         <%if media <> "print" then %>
	        <a href="../to_2015/ugeseddel_2011.asp?usemrn=<%=intMid %>&varTjDatoUS_man=<%=varTjDatoUS_man_use%>&nomenu=1" class="rmenu" target="_blank"><%=datepart("ww", startDato,2,2) %> </a>
            <!-- weekpage_2010.asp?medarbid=<%=intMid %>&st_dato=<%=day(varTjDatoUS_man_use)&"/"&month(varTjDatoUS_man_use)&"/"&year(varTjDatoUS_man_use) %> -->
	        <%else %>
	        <%=datepart("ww", startDato,2,2) %>
	        <%end if %>
	        
            <%
            select case month(startdato)
                case cint(1)
                strMonthName = left(godkendweek_txt_157, 3)
                case cint(2)
                strMonthName = left(godkendweek_txt_158, 3)
                case cint(3)
                strMonthName = left(godkendweek_txt_159, 3)
                case cint(4)
                strMonthName = left(godkendweek_txt_160, 3)
                case cint(5)
                strMonthName = left(godkendweek_txt_161, 3)
                case cint(6)
                strMonthName = left(godkendweek_txt_162, 3)
                case cint(7)
                strMonthName = left(godkendweek_txt_163, 3)
                case cint(8)
                strMonthName = left(godkendweek_txt_164, 3)
                case cint(9)
                strMonthName = left(godkendweek_txt_165, 3)
                case cint(10)
                strMonthName = left(godkendweek_txt_166, 3)
                case cint(11)
                strMonthName = left(godkendweek_txt_167, 3)
                case cint(12)
                strMonthName = left(godkendweek_txt_168, 3)
            end select
            %>

	        - <%=day(startDato) &". "& strMonthName %>
            
	        
	        </td>
	    
	    <%
        end if

	    strEksportTxt = strEksportTxt & datepart("yyyy", startDato,2,2) & ";"& datepart("m", startDato,2,2) & ";"& datepart("ww", startDato,2,2) & ";"
	    end select
	 
	 
         if normTimer(x) <> 0 then 
            
            if session("stempelur") <> 0 then 
                
                call timerogminutberegning(normTimer(x)*60)
		        
                normTal = ""& thoursTot &":"& left(tminTot, 2)
                normTalExp = ""& thoursTot &"."& left(tminTot, 2)

              else
                normTal = formatnumber(normTimer(x),2)
                normTalExp = formatnumber(normTimer(x),2)
                
            end if

        else
        normTal = ""
        normTalExp = 0
        end if 


	    if media <> "export" then %>

        <!-- Norm -->

        <td align=right style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" class=lille>
        <%=normTal %>
        </td>
    
    
    <%end if %>

    <%strEksportTxt = strEksportTxt & normTalExp & ";"
         if visning = 14 then
         strEksportTxt = strEksportTxt & formatnumber(normTimer(x), 2) & ";" 
         end if%>
	
            
    <%anormTimerTot = anormTimerTot + (normTimer(x)) %>
	
    
    
    <!--- Komme / G� tider --> 
	<%if session("stempelur") <> 0 then %> 
                 
                 <%call timerogminutberegning(totalTimerPer100) %>
                 
                         <%if media <> "export" then %>

                                     <%if show = 77 then %>
                                     <td align=right style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" class=lille><%=strlogin %></td>

                      
                         
                                                <%select case lto
                                                    case "kejd_pb", "xintranet - local"
                                                    case else %>
                                                    <td align=right style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" class=lille><%=strloginPauser %></td>

                    
                                                <%end select %>
                         
                                    <%end if %>

            
             
        
            

	                     <td align=right style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" class=lille>
                
	             
	             
                             <%
                 
                             if visning = 14 OR visning = 77 then %>
	 
	                                 <%if media = "print" then %>
	                                 <%=thoursTot &":"& left(tminTot, 2) %>
	                                 <%else 
                                        
                                        slDato = dateAdd("d", 7, varTjDatoUS_man_use)%>
                                        <a href="logindhist_2011.asp?usemrn=<%=intMid %>&varTjDatoUS_man=<%=varTjDatoUS_man_use %>&varTjDatoUS_son=<%=slDato %>&rdir=<%=rdir%>&nomenu=1" class="rmenu" target="_blank"><%=thoursTot &":"& left(tminTot, 2) %></a>
                                        <!--weekpage_2010.asp?medarbid=<%=intMid %>&st_dato=<%=stDato%>&func=lo&rdir=<%=rdir %>" class="rmenu" target="_blank">-->
        
	                                 <%end if %>
	                         <%else%>
	                         <%=thoursTot &":"& left(tminTot, 2) %>
	                         <%end if 
                             %>
	             
                         </td>
               <%end if 'export%>
                
                         


                        <%
                        if show = 77 then    
                        strEksportTxt = strEksportTxt & strlogin & ";" %>


                         <%select case lto
                         case "kejd_pb", "xintranet - local"
                         case else %>
                         <%strEksportTxt = strEksportTxt & strloginPauser & ";" %>
                    
                        <%end select 
                            
                            
                         end if%>
                        
                        <%
                        '** Komme / G� l�ntimer  **'
                        strEksportTxt = strEksportTxt & thoursTot &"."& left(tminTot, 2) & ";"
                             
                             if visning = 14 then
                             strEksportTxt = strEksportTxt & formatnumber(totalTimerPer100/60, 2) & ";"
                             end if %>
              

        <%if cint(mtypNoflex) <> 1 then 'noflex %> 

                <%if cint(showkgtil) = 1 then 
    
                         if media <> "export" then%>
                            <td align=right style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap; color:#999999; font-size:9px;">
                            <%if fradragTimer(x) <> 0 then %>
                            <%=formatnumber(fradragTimer(x), 2) %>
                            <%else %>
                            &nbsp;
                            <%end if %></td>
	             
                 
                         <td align=right style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" class=lille>
                 
                          <%if ltimerKorFrad <> 0 then %>
                            <%=formatnumber(ltimerKorFrad, 2) %>
                            <%else %>
                            &nbsp;
                            <%end if %>
                 
                 
                   </td>
	             
                 
                 
                         <% end if

           

	         strEksportTxt = strEksportTxt & formatnumber(fradragTimer(x), 2) & ";"& formatnumber(ltimerKorFrad, 2) & ";"
	         end if
	 
	 atotalTimerPer100 = atotalTimerPer100 + (totalTimerPer100)
	 afradragTimerTot = afradragTimerTot + (fradragTimer(x))
	 altimerKorFradTot = altimerKorFradTot + (ltimerKorFrad)
	 
	  'end if 'stempelur


       
      '**BAL L�ntimer (komme / G�) / Nomeret
      'if session("stempelur") <> 0 then 
            

      select case lto 
      case "kejd_pb"
            
       case else
            
            call timerogminutberegning(normtime_lontime) 
             
                             
             if media <> "export" then%> 
	         <td align=right style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" class=lille>
	         <b><%=thoursTot &":"& left(tminTot, 2) %></b>
	         </td>
             <%end if 'exp
     
                  strEksportTxt = strEksportTxt & thoursTot &"."& left(tminTot, 2) & ";"
                 
                  if visning = 14 then
                  strEksportTxt = strEksportTxt & formatnumber(normtime_lontime/60, 2) & ";"
                  end if 
      
	     
        
          
         lontimeFlexAkk = (normtime_lontimeAkk * 1 + akuPreNormLontBal * 60 * 1) '(akuPreNormLontBal * 60))
         call timerogminutberegning(lontimeFlexAkk)  
         
         if media <> "export" then%> 
          <!-- Akkumuleretet -->
          <%
     
       %>
          <td align=right bgcolor="#DCF5BD" style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" class=lille>

           

	        
             <%if (stfor = -1 AND visning = 7) OR (stfor = 0 AND visning = 77) then
            %>
            <span style="color:#999999; font-size:9px;"><%=akuPreNormLontBal60 %>  >> </span><br />
            <%
            end if 
            %>
            
                 
		    <b><%=thoursTot &":"& left(tminTot, 2) %></b> 
              
              <!--
              <%if session("mid") = 1 then %>
              (<%=lontimeFlexAkk %>)
               <%=akuPreNormLontBal%> // * 60  =
              <%=akuPreNormLontBal * 60%>
              <%end if %>
              -->

        
	     </td>

         <%
         end if'exp
         
         strEksportTxt = strEksportTxt & thoursTot &"."& left(tminTot, 2) &";"

             if visning = 14 then
             strEksportTxt = strEksportTxt & formatnumber(lontimeFlexAkk/60, 2) & ";"
             end if

            end select 'kejd_pb

       end if 'flex

       end if'stempelur






     if media <> "export" then
      
      %>


      <!-- Real timer --->

	    <td align=right style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" class=lille>
        <% if realTimer(x) <> 0 then %>
        <%=formatnumber(realTimer(x),2)%>
        <%else %>
        &nbsp;
        <%end if %></td>
        
        <%end if %>

    <%arealTimerTot = arealTimerTot + (realTimer(x)) %>
    <%strEksportTxt = strEksportTxt & formatnumber(realTimer(x),2) & ";" %>

         
         
         <%
          '** Hereby Invoiceable / Heraf fakturerbare  
          if lto <> "cst" AND lto <> "tec" AND lto <> "esn" then 
      
          if media <> "export" then%>


              <%select case lto
                case "kejd_pb", "wap"
                case else %>

	         <td align=right class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;">
              <%if realfTimer(x) <> 0 then %>
              <%=formatnumber(realfTimer(x),2)%>
             <%else %>
             &nbsp;
             <%end if %>
             </td>
             

             <%end select %>


            <%'** Hereby Non Invoiceable / Heraf Ikke fakturerbare
            if lto = "tia" OR lto = "intranet - local" then %>
              
                <td align=right class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;">
                <%if realIfTimer(x) <> 0 then %>
                <%=formatnumber(realIfTimer(x),2)%>
                <%else %>
                &nbsp;
                <%end if %>
             </td> 

            <% end if%>


        <%if cint(mtypNoflex) <> 1 then 'noflex %>

               <%if cint(showkgtil) = 1 AND (lto <> "cst" AND lto <> "tec" AND lto <> "esn" AND lto <> "wap" AND lto <> "lm") then %>
                <td align=right style="border-bottom:1px silver solid; border-right:1px silver solid; font-size:9px; color:#999999; white-space:nowrap;">
                      <%if korrektionReal(x) <> 0 then %>
                     <%=formatnumber(korrektionReal(x),2)%>
                     <%else %>
                     &nbsp;
                     <%end if %>

               <%end if %>


         <%end if 'noflex %>

         <%end if 'exp %>



	     
                  <%select case lto
                        case "kejd_pb", "wap"
                        case else %>

                 <%arealfTimerTot = arealfTimerTot + (realfTimer(x)) %>
                 <%strEksportTxt = strEksportTxt & formatnumber(realfTimer(x),2) & ";" %>

                <%end select %>


                <%if lto = "intranet - local" OR lto = "tia" then 
                    %>
                    
                     <%arealifTimerTot = arealifTimerTot + (realifTimer(x)) %>
                 <%strEksportTxt = strEksportTxt & formatnumber(realifTimer(x),2) & ";" %>
                    
                <%end if%>

                      <%if cint(mtypNoflex) <> 1 then 'noflex %>

                           <%if cint(showkgtil) = 1 then %>
                              <%korrektionRealTot = korrektionRealTot + (korrektionReal(x)) %>
                         <%strEksportTxt = strEksportTxt & formatnumber(korrektionReal(x),2) & ";" %>

	                      <%
                              end if 'noflex

                      end if %>
         
         <%end if 'lto 
             
             
             
        if cint(mtypNoflex) <> 1 then 'noflex %>
	 
             <%


             if lto <> "cst" AND lto <> "tec" AND lto <> "esn" AND lto <> "tia" AND lto <> "lm" then 

                          if media <> "export" then
                         %>
	                     <td align=right style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" class=lille ><b><%=formatnumber(balRealNormtimer,2)%></b></td>
	     
                         <td align=right style="border-bottom:1px silver solid; border-right:1px #FFC0CB solid; white-space:nowrap;" bgcolor="#FFC0CB" class=lille>
                             <%if (stfor = -1 AND visning = 7) OR (stfor = 0 AND visning = 77) then
                            %>
                            <span style="color:#999999; font-size:9px;"><%=formatnumber(akuPreRealNormBal, 2) %> >> </span><br /> <!-- +(korrektionReal(x)) -->
                            <%
                            end if 


                            select case lto
                            case "wk", "wk_no", "wk_dk", "mi", "akelius", "mi_no"
                                balRealNormtimerAkk_beregnet = formatnumber(balRealNormtimerAkk+(akuPreRealNormBal),2)
                              
                            case else
                                balRealNormtimerAkk_beregnet = formatnumber(balRealNormtimerAkk+(akuPreRealNormBal),2)
                               
                            end select
                            %>
                            <b><%=formatnumber(balRealNormtimerAkk_beregnet,2)%></b>
                      </td>
                <%end if %>
    
    
             <%strEksportTxt = strEksportTxt & formatnumber(balRealNormtimer, 2) & ";" & formatnumber(balRealNormtimerAkk_beregnet, 2) & ";" %>
	 
	          <%
	         end if 'lto
     
     
          if session("stempelur") <> 0 then 
     
        
                '**BAL Real / L�ntimer 
                if lto <> "cst" AND lto <> "kejd_pb" AND lto <> "tec" AND lto <> "esn" AND lto <> "lm" then 
                  if media <> "export" then%>
	           <td align=right style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" class=lille> <%=formatnumber(balRealLontimer,2)%></td>
	            <%
                end if
                strEksportTxt = strEksportTxt & formatnumber(balRealLontimer,2) & ";" %>
	 
	    
	            <%end if 'lto %>
	     <%
         end if 'stempelur
         
     end if 'noflex%>
	 
	 
	     <!-- Afspad / Overarb --->
	       <%if (instr(akttype_sel, "#30#") <> 0 OR instr(akttype_sel, "#31#") <> 0) AND lto <> "cflow"  then 
           
                     if media <> "export" then
                     
                         if lto <> "kejd_pb" AND lto <> "fk" AND lto <> "adra" AND lto <> "cisu" then %>
	           
	                     <td align=right style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" class=lille>
                          
                         <%if afspTimer(x) <> 0 then %>

                                <% 
                                '*** Overarbejde optjent i timer Wolters WK, Magnus Informatik, WK_DK, MI, WK_NO  
                                select case lto
                                case "wk", "wk_no", "wk_dk", "mi", "akelius", "mi_no"
                                 %>
                                <span style="font-size:8px; color:#999999;"><%=formatnumber(afspTimerTim(x),2) %></span><br />
                                <%
                                end select%>

                            
                         <%=formatnumber(afspTimer(x), 2)%>
                         <%else %>
                         &nbsp;
                         <%end if 
                     
             
                         %></td>
	                 
                         <%end if %>
             
                     <td align=right class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;">
                     <%if afspTimerBr(x) <> 0 then %>
                     <%=formatnumber(afspTimerBr(x), 2)%>
                     <%else %>
                     &nbsp;
                     <%end if
                     
                      %></td>

                          <%
                     end if 
                    
                    if lto <> "kejd_pb" AND lto <> "fk" AND lto <> "adra" AND lto <> "cisu" then 
                    strEksportTxt = strEksportTxt & formatnumber(afspTimer(x), 2) &";"
                    end if

                    strEksportTxt = strEksportTxt & formatnumber(afspTimerBr(x), 2) &";"
                    aafspTimerBrTot = aafspTimerBrTot + afspTimerBr(x)
             

                                
                  '***** Afspadsering udbetalt og Saldo **'
                  if lto <> "lw" AND lto <> "kejd_pb" AND lto <> "fk" AND lto <> "adra" AND lto <> "cisu" AND lto <> "wap" then 
                  
                            if lto <> "tec" AND lto <> "esn" then

                                 if media <> "export" then%>
	                             <td align=right class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;">
                                         <%if afspTimerUdb(x) <> 0 then %>
                                         <%=formatnumber(afspTimerUdb(x), 2)%>
                                         <%else %>
                                         &nbsp;
                                         <%end if %>

                         
                                 </td>
	     
	                              <% end if
                                  strEksportTxt = strEksportTxt & formatnumber(afspTimerUdb(x), 2) &";"
                                   aafspTimerUdbTot = aafspTimerUdbTot + afspTimerUdb(x)

                            end if


	                      aafspTimerTot = aafspTimerTot + afspTimer(x) 
	            
	          

                    '**** Afspadsering Udbetalt ***'   
                 if lto <> "tec" AND lto <> "esn" then

	                 afspadUdbBal = 0
	                 afspadUdbBal = (afspTimerOUdb(x) - afspTimerUdb(x)) 
	         
	                 aafspadUdbBalTot = aafspadUdbBalTot + (afspadUdbBal)
	             
                       if media <> "export" then%>
	                  <td align=right class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;">
                              <%if afspadUdbBal <> 0 then %>
                              <%=formatnumber(afspadUdbBal, 2)%>
                              <%else %>
                              &nbsp;
                              <%end if %>
                      </td>
	                  <%end if 
                      
                        strEksportTxt = strEksportTxt & formatnumber(afspadUdbBal, 2) &";"
                   end if
	             
               
                      
                 '**** Afspadsering Saldo ***'       
                 AfspadBal = 0 
	             AfspadBal = (afspTimer(x) - (afspTimerBr(x)+ afspTimerUdb(x)))
	             aAfspadBalTot = aAfspadBalTot + (AfspadBal)
	              
                           if media <> "export" then%>
	                      <td align=right class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;">
                          <%if AfspadBal <> 0 then %>
                          <%=formatnumber(AfspadBal, 2)%> 
                          <%else %>
                          &nbsp;
                          <%end if %></td>
        	             <%end if %>
	         
	             <% strEksportTxt = strEksportTxt & formatnumber(AfspadBal, 2) &";" %>
	              <%end if %>
	         
	         <%end if '30/31 %>


             <%

        


         
	       
                '****************************************************************************
                'Clfow Special 
                '90 / 91 - Overtid 50 / 100 fastl�nn Cflow
               '**************************************************************************** 
               if (instr(akttype_sel, "#91#") <> 0 OR instr(akttype_sel, "#92#") <> 0) AND lto = "cflow" then 

               e2TimerTot = e2TimerTot + e2Timer(x)
               e3TimerTot = e3TimerTot + e3Timer(x)

                  if media <> "export" then%>

	                      <td align=right class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;">
                          <%if e2Timer(x) <> 0 then %>
                          <%=formatnumber(e2Timer(x), 2)%> 
                          <%else %>
                          &nbsp;
                          <%end if %></td>

                            <td align=right class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;">
                          <%if e3Timer(x) <> 0 then %>
                          <%=formatnumber(e3Timer(x), 2)%> 
                          <%else %>
                          &nbsp;
                          <%end if %></td>

        	             <%end if %>
	         
	             <% strEksportTxt = strEksportTxt & formatnumber(e2Timer(x), 2) &";" & formatnumber(e3Timer(x), 2) &";" %>
           <%
            end if
               



                 
              

             




                '****************************************************************************
                
                'TEC / ESN special omsorgsdage beregning
                'ESN
                'Mertid                     = 7         (Fleks)
                'Ulempe 1706                = 6         (Salg | Newbizz)
                'Ulempe "W                  = 60        (Ad-hoc)

                'Omsorgsdage2 optjent       = 50        (Dag)
                'Omsorgsdage2 afholdt       = 23        (Omsorgsdage)
                'Omsorgsdage10 optjent      = 51        (Nat)
                'Omsorgsdage10 afholdt      = 8         (Sundhed)
                'OmsorgsdageK optjent       = 52        (Weekend)
                'OmsorgsdageK afholdt       = 115       (Tjenestefri)

               

                        select case lto
                        case "intranet - local", "esn"
                    
                        'MERTID slettet her
                        
                        if media <> "export" then %>

                      
                         <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=right>

                                     <%if sTimer(x) <> 0 then %>
                                     <%=formatnumber(sTimer(x),2) %>
                                     <%else %>
                                     &nbsp;
                                     <%end if 

                                    
                                         
                                        
                         %></td>    

                         
                         <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=right>

                                     <%if adhocTimer(x) <> 0 then %>
                                     <%=formatnumber(adhocTimer(x),2) %>
                                     <%else %>
                                     &nbsp;
                                     <%end if 

                                        
                                         
                                        
                         %></td>   

                        
                        <% if visning <> 14 then 'periode afslutning%>

                          <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=right>

                                     <%if ulempe1706udb(x) <> 0 then %>
                                     <%=formatnumber(ulempe1706udb(x),2) %>
                                     <%else %>
                                     &nbsp;
                                     <%end if 

                                   
                                         
                                        
                         %></td>    


                               <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=right>

                                     <%if ulempeWudb(x) <> 0 then %>
                                     <%=formatnumber(ulempeWudb(x),2) %>
                                     <%else %>
                                     &nbsp;
                                     <%end if 

                                    
                                         
                                        
                         %></td>    

            
                        <% end if 'visning <> 14



                                 sTimer_tot = sTimer_tot + sTimer(x)
                                 adhocTimer_tot = adhocTimer_tot + adhocTimer(x)

                               if visning <> 14 then
                               ulempe1706udb_tot = ulempe1706udb_tot + ulempe1706udb(x)
                               ulempeWudb_tot = ulempeWudb_tot + ulempeWudb(x)
                               end if

                                


                            end if 'media

                             
                            
                            
                             strEksportTxt = strEksportTxt & formatnumber(sTimer(x),2) &";"
                             strEksportTxt = strEksportTxt & formatnumber(adhocTimer(x),2) &";"
                             


                            if visning <> 14 then

                               strEksportTxt = strEksportTxt & formatnumber(ulempe1706udb(x),2) &";"
                               strEksportTxt = strEksportTxt & formatnumber(ulempeWudb(x),2) &";"
                              
                            end if


                            

                        end select




                         ' if visning <> 14 then '14: periode afslutning


                        select case lto
                        case "intranet - local", "tec", "esn"
                        
                         if media <> "export" AND visning = 77 then
                         'Omsorgsdage2 optjent%>
	                     <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=right>

                       

                                     <%if dagTimer(x) <> 0 then %>
                                     <%=formatnumber(dagTimer(x),2) %>
                                     <%else %>
                                     &nbsp;
                                     <%end if 
                                         
                                        
                         %></td>

                        <%end if 

                            

                            if visning = 77 then
                                if media = "export" then
                                strEksportTxt = strEksportTxt & formatnumber(dagTimer(x),2) &";"
                                end if
                            
                            dagTimer_tot = dagTimer_tot + dagTimer(x)
                            end if



                         if media <> "export" and (lto = "tec" OR (lto = "esn" AND visning = 14)) then
                        'Omsorgsdage2 afholdt %>
                           <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=right>

                        

                                     <%if omsorg(x) <> 0 then %>
                                     <%=formatnumber(omsorg(x),2) %> 
                                     <%else %>
                                     &nbsp;
                                     <%end if 
                                         
                                        
                         %></td>                        
                        <%end if
                            
                       

                            'if (visning = 7 OR visning = 77) then
                            

                            if media = "export" and (lto = "tec" OR (lto = "esn" AND visning = 14))  then
                            strEksportTxt = strEksportTxt & formatnumber(omsorg(x),2) &";" 
                            end if





                            omsorg2afh_tot = omsorg2afh_tot + omsorg(x)

                            if normTimerDag(x) <> 0 then
                            omsorg2Saldo = (dagTimer(x) - (omsorg(x)))
                            else
                            omsorg2Saldo = 0
                            end if
                            'end if



                          



                         if media <> "export" AND (visning = 77 or (visning = 7 AND lto = "esn")) then
                          'Omsorgsdage2 Saldo 
                            %>
	                     <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=right>

                       

                                     <%if omsorg2Saldo <> 0 then %>
                                     <%=formatnumber(omsorg2Saldo,2) %>
                                     <%else %>
                                     &nbsp;
                                     <%end if 
                                         
                                         
                         %></td>

                        <%end if 

                            

                            if visning = 77 or (visning = 7 and lto = "esn") then

                                if media = "export" then
                                strEksportTxt = strEksportTxt & formatnumber(omsorg2Saldo,2) &";"
                                end if

                            omsorg2Saldo_tot = omsorg2Saldo_tot + omsorg2Saldo
                            end if


                         if media <> "export" AND visning = 77 then
                         'Omsorgsdage10 optjent%>
	                     <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=right>

                       

                                     <%if natTimer(x) <> 0 then %>
                                     <%=formatnumber(natTimer(x),2) %>
                                     <%else %>
                                     &nbsp;
                                     <%end if 
                                         
                                         
                         %></td>

                        <%end if 

                            

                            if visning = 77 then
                            strEksportTxt = strEksportTxt & formatnumber(natTimer(x),2) &";"
                            nattimer_tot = nattimer_tot + natTimer(x)
                            end if



                         if media <> "export" and lto <> "esn" then 'AND (visning = 7 OR visning = 77)
                          'Omsorgsdage10 afholdt %>
	                     <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=right>

                       

                                     <%if sundTimer(x) <> 0 then %>
                                     <%=formatnumber(sundTimer(x),2) %>
                                     <%else %>
                                     &nbsp;
                                     <%end if 
                                         
                                        
                         %></td>

                        <%end if 


                        if lto = "esn" AND visning = 14 AND media <> "export" then 
                          'Omsorgsdage10 afholdt %>
	                     <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=right>

                       

                                     <%if sundTimer(x) <> 0 then %>
                                     <%=formatnumber(sundTimer(x),2) %>
                                     <%else %>
                                     &nbsp;
                                     <%end if 
                                         
                                        
                         %></td>

                        <%end if


                            'if (visning = 7 OR visning = 77) then
                                sundTimer_tot = sundTimer_tot + sundTimer(x)
                           if media = "export" AND lto = "esn" AND visning = 14 then 
                                strEksportTxt = strEksportTxt & formatnumber(sundTimer(x),2) &";" 
                            end if
                           
                            
                           omsorg10Saldo = (natTimer(x) - sundTimer(x))
                            'end if




                         if media <> "export" AND (visning = 77 or (visning = 7 AND lto = "esn")) then
                         
                         'Omsorgsdage10 Saldo%>
	                     <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=right>

                       

                                     <%if omsorg10Saldo <> 0 then %>
                                     <%=formatnumber(omsorg10Saldo,2) %>
                                     <%else %>
                                     &nbsp;
                                     <%end if 
                                         
                                         
                         %></td>

                        <%end if 



                            if visning = 77 or (visning = 7 AND lto = "esn") then

                                if media = "export" then
                                strEksportTxt = strEksportTxt & formatnumber(omsorg10Saldo,2) &";"
                                end if
                            
                            omsorg10Saldo_tot = omsorg10Saldo_tot + omsorg10Saldo
                            end if



                     

                         if media <> "export" AND (visning = 77 OR (lto = "esn" AND (visning = 14))) then
                         'OmsorgsdageK Optjent%>
	                     <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=right>

                       

                                     <%if weekendTimer(x) <> 0 then %>
                                     <%=formatnumber(weekendTimer(x),2) %>
                                     <%else %>
                                     &nbsp;
                                     <%end if 
                                         
                                         
                         %></td>

                        <%end if 

                     
                            
                            
                            if visning = 77  OR (lto = "esn" AND visning = 14) then

                                if media = "export" then
                                strEksportTxt = strEksportTxt & formatnumber(weekendTimer(x),2) &";"
                                end if
                            
                            weekendTimer_tot = weekendTimer_tot + weekendTimer(x)
                            end if




                         if media <> "export" and lto <> "esn" then 'AND (visning = 7 OR visning = 77)
                         'OmsorgsdageK afholdt%>
	                     <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=right>

                       

                                     <%if tjenestefri(x) <> 0 then %>
                                     <%=formatnumber(tjenestefri(x),2) %>
                                     <%else %>
                                     &nbsp;
                                     <%end if 
                                         
                                         
                         %></td>

                        <%end if 

                            omsorgKAfh_tot = omsorgKAfh_tot + tjenestefri(x)
                            
                            if lto <> "esn" then 
                                strEksportTxt = strEksportTxt & formatnumber(tjenestefri(x),2) &";" 
                            end if
                            
                            OmsorgKSaldo = (weekendTimer(x) - tjenestefri(x))
                           
                        
                          'OK HERTIL ESN


                          if media <> "export" AND visning = 77 or (visning = 7 AND lto = "esn")  then
                         'OmsorgsdageK Saldo
                          %>
	                     <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=right>

                       
                           
                                     <%if OmsorgKSaldo <> 0 then %>
                                     <%=formatnumber(OmsorgKSaldo,2) %>
                                     <%else %>
                                     &nbsp;
                                     <%end if 
                                         
                                         
                         %></td>

                        <%end if 

                            
                            
                            if visning = 77 or (visning = 7 AND lto = "esn") then
                            strEksportTxt = strEksportTxt & formatnumber(OmsorgKSaldo,2) &";"
                            OmsorgKSaldo_tot = OmsorgKSaldo_tot + OmsorgKSaldo
                            end if
                        %>

                        <%if visning = 7 and lto = "esn" then %>

                        <!--
                        <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=right>
                            l:
                            <%if omsorg2Saldo <> 0 then %>
                             <%=formatnumber(omsorg2Saldo,2) %>
                            <%end if %> 
                        </td>
                        -->

                        
                        <!--
                        <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=right>
                            <%if omsorg10Saldo <> 0 then%>
                            f:<%=formatnumber(omsorg10Saldo,2) %>
                            <%end if
                            'response.Write natTimer(x) & "h " & sundTimer(x) %> 
                        </td>
                        -->
                        
                        
                        <!--

                        <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=right>
                            <%if OmsorgKSaldo <> 0 then%>
                            <%=formatnumber(OmsorgKSaldo,2) %>
                            <%end if %>
                        </td>
                            -->
                        <%end if %>

                        <%
                        end select 
                            
                    'end if 'visning 14 godkend ugesedler%>
	         
	                    
	        
	        
	 <%'** Slut TEC / ESN



    '*** Aldersreduktion *****
         if instr(akttype_sel, "#27#") <> 0 OR instr(akttype_sel, "#28#") <> 0 OR instr(akttype_sel, "#29#") <> 0 then 


            if (visning = 7 OR visning = 77) then

                if visning = 77 then 'Kun p� afstembning udspec. dage

                        if aldersreduktionOpj(x) <> 0 then
                        aldersreduktionOpjTxt = formatnumber(aldersreduktionOpj(x),2) 
                        aldersreduktionOpjTxtExp = aldersreduktionOpjTxt
                        else
                        aldersreduktionOpjTxt = ""
                        aldersreduktionOpjTxtExp = 0
                        end if

                        aldersreduktionOpjTot = aldersreduktionOpjTot + aldersreduktionOpj(x)
                                         
                         if media <> "export" then %>

	                     <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=right><%=aldersreduktionOpjTxt %></td>


                    <% 
                        else
                         strEksportTxt = strEksportTxt & formatnumber(aldersreduktionOpjTxtExp,2) &";"

                        end if


                 end if '* 77





                    '** 7 og 77
                    if aldersreduktionBr(x) <> 0 then
                    aldersreduktionBrTxt = formatnumber(aldersreduktionBr(x),2) 
                    aldersreduktionBrTxtExp = aldersreduktionBrTxt
                    else
                    aldersreduktionBrTxt = ""
                    aldersreduktionBrTxtExp = 0
                    end if

                    aldersreduktionBrTot = aldersreduktionBrTot + aldersreduktionBr(x)


                    if media <> "export" then
                    %>
                        <%if lto <> "esn" then %>
                        <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=right><%=aldersreduktionBrTxt%></td>
                        <%end if %>
                    <%
                    else
                        if lto <> "esn" then
                        strEksportTxt = strEksportTxt & formatnumber(aldersreduktionBrTxtExp,2) &";"
                        end if
                    end if 


                
                    
                    
                           if visning = 77 or (visning = 7 and lto = "esn") then 'Kun p� afstembning udspec. dage   + ugeseddel
                    
                                 if aldersreduktionUdb(x) <> 0 then
                                 aldersreduktionUdbTxt = formatnumber(aldersreduktionUdb(x),2) 
                                 aldersreduktionUdbTxtExp = aldersreduktionUdbTxt
                                 else
                                 aldersreduktionUdbTxt = ""
                                 aldersreduktionUdbTxtExp = 0
                                 end if 
                    
                    
                                   aldersreduktionUdbTot = aldersreduktionUdbTot + aldersreduktionUdb(x)  
                                   if visning = 77 AND lto <> "esn"  then
                                             if media <> "export" then  %>
                                        
                                         <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=right><%=aldersreduktionUdbTxt %></td>

                                        <%else
                                                     strEksportTxt = strEksportTxt & formatnumber(aldersreduktionUdbTxtExp,2) &";"
                
                                          end if  

                                    end if

           
                
                                     aldersreduktionSaldo = 0
                                     aldersreduktionSaldo = aldersreduktionOpj(x) - (aldersreduktionBr(x) + aldersreduktionUdb(x))

           
                                     if aldersreduktionSaldo <> 0 then
                                     aldersreduktionSaldoTxt = formatnumber(aldersreduktionSaldo, 2)
                                     aldersreduktionSaldoTxtExp = aldersreduktionSaldoTxt
                                     else
                                     aldersreduktionSaldoTxt = ""
                                     aldersreduktionSaldoTxtExp = 0
                                     end if 
                
                                     aldersreduktionSaldoTot = aldersreduktionSaldoTot + aldersreduktionSaldo  
                

                                                         if media <> "export" then
                                                         %>

                                                         <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=right><%=aldersreduktionSaldoTxt %></td>

                                                <%
                                                    else
                                                     strEksportTxt = strEksportTxt & formatnumber(aldersreduktionSaldoTxtExp,2) &";"

                                                    end if 
                            
                            
                            
                               end if 'vis ikke p� ugeseddel 7 
                            

                end if '7 / 77
                                             
           end if '#27/#28/#29








     '***************** FERIE FRI **********************************************    
         
         
     if normTimerDag(x) <> 0 then
	 ferieAfVal_md = (ferieAFPerTimer(x)+ferieUdbPer(x))/normTimerDag(x)
     else
	 ferieAfVal_md = 0
     end if


     if normTimerDag(x) <> 0 then
	 ferieAfUlonVal_md = ferieAFulonPer(x)/normTimerDag(x)

         if ferieAfUlonVal_md  <> 0 then
         feAfhUlonExpTxt = ferieAfUlonVal_md
         else
         feAfhUlonExpTxt = 0
         end if

	 else
	 ferieAfUlonVal_md  = 0
     feAfhUlonExpTxt = 0
	 end if
	 
	 ferieAfVal_md_tot = ferieAfVal_md_tot + ferieAfVal_md
     ferieAfulonVal_md_tot = ferieAfulonVal_md_tot + feAfhUlonExpTxt
     
     if normTimerDag(x) <> 0 then
	 ferieFriAfVal_md = fefriTimerPerBr(x)+fefriTimerUdbPer(x)/normTimerDag(x)
	 else
	 ferieFriAfVal_md = 0
	 end if
	 
	 ferieFriAfVal_md_tot = ferieFriAfVal_md_tot + ferieFriAfVal_md


	  %>
	 
	                     <%'if visning = 7 OR visning = 77 then 
                          if media <> "export" then%>
	                     <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=right>

                       

                        <%if ferieAfVal_md <> 0 then %>
                        <%=formatnumber(ferieAfVal_md,2) %>               
                        <%else %>
                        &nbsp;
                        <%end if %></td>

                        <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=right>
                        <%if ferieAfUlonVal_md <> 0 then %>
                        <%=formatnumber(ferieAfUlonVal_md,2)%>               
                        <%else %>
                        &nbsp;
                        <%end if %></td>

	                    <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=right>
                        <%if ferieFriAfVal_md <> 0 then %>
                        <%=formatnumber(ferieFriAfVal_md,2) %>
                        <%else %>
                        &nbsp;
                        <%end if %></td> 
                                           
	                       <%end if
                                
                                strEksportTxt = strEksportTxt & formatnumber(ferieAfVal_md,2) &";"& formatnumber(feAfhUlonExpTxt,2) &";"& formatnumber(ferieFriAfVal_md,2) & ";"




                        '1 maj timer
                        select case lto
                        case "xintranet - local", "fk", "plan"
                        
                             if media <> "export" then%>
	                         <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=right>

                      
         

                                         <%if divfritimer(x) <> 0 then %>
                                         <%=formatnumber(divfritimer(x),2) %>
                                         <%else %>
                                         &nbsp;
                                         <%end if 
                                       
                                             divfritimer_tot = divfritimer_tot + divfritimer(x)
                                         
                                             %></td>

                            <%end if 

                            strEksportTxt = strEksportTxt & formatnumber(divfritimer(x),2) &";"

                        end select



                        '1 Rejsedage
                        select case lto
                        case "intranet - local", "adra", "plan"
                        
                             if media <> "export" then%>
	                         <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=right>

                      
         

                                         <%if rejseDage(x) <> 0 then %>
                                         <%=formatnumber(rejseDage(x),2) %>
                                         <%else %>
                                         &nbsp;
                                         <%end if 
                                       
                                             rejsedage_tot = rejsedage_tot + rejseDage(x)
                                         
                                             %></td>


                            <%
                            end if    
                        end select            




                     '***************************************************************
                     '** Omsorg optjent
                     '***************************************************************
                     select case lto
                     case "ddc", "intrant - local" 
                 
                             if visning = 7 OR visning = 77 then

                             omsorgOpt(x) = omsorgOpt(x)

                             if omsorgOpt(x) <> 0 then
                             omsorgOptTxt = formatnumber(omsorgOpt(x), 2) 
                             else
                             omsorgOptTxt = ""
                             end if

                                 if media <> "export" then
                                 %>
                                <td align=right class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;"><%=omsorgOptTxt%></td>
                                <%
                                
                                omsorgOptjTot = omsorgOptjTot + omsorgOpt(x)
                                
                                end if

                            
                            strEksportTxt = strEksportTxt & formatnumber(omsorgOpt(x), 2) &";"
                            

                            end if

                            

                     end select



                        'Omsorgsdage
                        select case lto
                        case "xintranet - local", "fk", "kejd_pb", "adra", "ddc", "wap"
                        
                         if media <> "export" then%>
	                     <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=right>

                       

                                     <%if omsorg(x) <> 0 then %>
                                     <%=formatnumber(omsorg(x),2) %>
                                     <%else %>
                                     &nbsp;
                                     <%end if 
                                         
                                         omsorg_tot = omsorg_tot + omsorg(x)
                                         %></td>

                        <%end if 

                            strEksportTxt = strEksportTxt & formatnumber(omsorg(x),2) &";"
                        end select


                             'Tjenestefri 
                        select case lto
                        case "xintranet - local", "fk", "kejd_pb"
                        
                         if media <> "export" then%>
	                     <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=right>

                       

                                     <%if tjenestefri(x) <> 0 then %>
                                     <%=formatnumber(tjenestefri(x),2) %>
                                     <%else %>
                                     &nbsp;
                                     <%end if 
                                         
                                         tjenestefri_tot = tjenestefri_tot + tjenestefri(x)
                                         %></td>
                                

                                  
                        <%end if 

                            strEksportTxt = strEksportTxt & formatnumber(tjenestefri(x),2) &";"
                        end select



                             'Barsel
                        select case lto
                        case "intranet - local", "fk", "kejd_pb", "tia"
                        
                         if media <> "export" then%>
	                     <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=right>

                       

                                     <%if barsel(x) <> 0 then %>
                                     <%=formatnumber(barsel(x),2) %>
                                     <%else %>
                                     &nbsp;
                                     <%end if 
                                         
                                         barsel_tot = barsel_tot + barsel(x)
                                         %></td>
                                

                                  
                        <%end if 

                            strEksportTxt = strEksportTxt & formatnumber(barsel(x),2) &";"
                        end select

                        
                       






                                    'L�ge
                        select case lto
                        case "intranet - local", "fk", "xkejd_pb"
                        
                         if media <> "export" then%>
	                  
                                

                                     <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=right>

                       

                                     <%if lageTimer(x) <> 0 then %>
                                     <%=formatnumber(lageTimer(x),2) %>
                                     <%else %>
                                     &nbsp;
                                     <%end if 
                                         
                                         lageTimer_tot = lageTimer_tot + lageTimer(x)
                                         %></td>

                        <%end if 

                            strEksportTxt = strEksportTxt & formatnumber(lageTimer(x),2) &";"
                        end select




                                  if (level = 1 OR (session("mid") = usemrn)) OR (cint(erTeamlederForVilkarligGruppe) = 1) then %>
	         
	                              <%
              
                                  if normTimerDag(x) <> 0 then 
	                              sygeDage = sygDage(x) / normTimerDag(x) 
	                              else 
	                              sygeDage = 0
	                              end if
	          
	                              sygeDage_tot = sygeDage_tot + sygeDage
	                              
                                       if media <> "export" then%> 
	           
	                                         <td align=right class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space: nowrap;">
                                             <%if sygeDage <> 0 then %>
                                             <%=formatnumber(sygeDage,2) %>
                                             <%else %>
                                             &nbsp;
                                             <%end if %></td>
	                             
                                     <%end if 
                                   

                                   strEksportTxt = strEksportTxt & formatnumber(sygeDage,2) &";"

                                  end if 'level 


                                  if (level = 1 OR (session("mid") = usemrn)) OR (cint(erTeamlederForVilkarligGruppe) = 1) then %>
	         
	                              <%
              
                                  if normTimerDag(x) <> 0 then 
	                              barnSyg = barnSygDage(x) / normTimerDag(x) 
	                              else 
	                              barnSyg = 0
	                              end if
	          
	                              barnSyg_tot = barnSyg_tot + barnSyg
	                              
                                       if media <> "export" then%> 
	           
	                                         <td align=right class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space: nowrap;">
                                             <%if barnSyg <> 0 then %>
                                             <%=formatnumber(barnSyg,2) %>
                                             <%else %>
                                             &nbsp;
                                             <%end if %></td>
	                             
                                     <%end if 
                                   

                                   strEksportTxt = strEksportTxt & formatnumber(barnSyg,2) &";"

                                  end if 'level %>
	         
	         
	       
	        
	                     <%'end if '*Visning


                       
                        if realIkkeGKTimer(x) <> 0 then    
                          realIkkeGKTimerTxt = formatnumber(realIkkeGKTimer(x), 2)
                          realIkkeGKTimerExpTxt = formatnumber(realIkkeGKTimer(x), 2)
                        else
                          realIkkeGKTimerTxt = ""
                          realIkkeGKTimerExpTxt = 0
                        end if

                        select case lto
                        case "esn", "tec"
                        case else

                        if media <> "export" AND visning = 14 then
                        %>
                        <td style="border-bottom:1px silver solid; border-right:1px silver solid; color:#999999;" class=lille align="right"><%=realIkkeGKTimerTxt%>
                            
                            <%if realAfvistTimer(x) <> 0 then
                            Response.write "<span style='color:red;'> ("& formatnumber(realAfvistTimer(x), 2) &")</span>" 
                            end if%> </td>

                        <%end if
                       
                             if (visning = 14) then    
                            strEksportTxt = strEksportTxt & realIkkeGKTimerExpTxt &";" & realAfvistTimer(x) & ";"
                            end if     
                            
                       end select 

                        '**** Kommentar p� afstemning

                        if lto = "esn" then

                            kommentarTxt = ""
                            strSQLKommentarUge = "SELECT kommentar, dato FROM login_historik WHERE mid = "& intMid &" AND dato = '"& startDato & "' AND kommentar <> ''"
                            oRec7.open strSQLKommentarUge, oConn, 3
                            while not oRec7.EOF 
                            kommentarTxt = kommentarTxt & oRec7("kommentar") &"</br>"
                            oRec7.movenext
                            wend
                            oRec7.close

                        end if

                         if media <> "export" AND visning = 77 then %>     
                         
                            <%if lto <> "esn" then %>                       
                            <td class=lille valign="top" style="border-bottom:1px silver solid; border-right:1px silver solid; width:200px; color:#999999; padding-left:3px;"><%=kommentarTxt %>&nbsp;</td>
                            <%else %>                            
                             <td class=lille valign="top" style="border-bottom:1px silver solid; border-right:1px silver solid; width:200px; color:#999999; padding-left:3px;"><%=kommentarTxt %></td>
                            <%end if %>

                        <%end if %>

                        <%
                        if (visning = 77 AND media = "export") then    
                        strEksportTxt = strEksportTxt & Chr(34) & replace(kommentarTxt,"</br>", "") & Chr(34) &";" 
                        end if

                        '**** Kommentar p� m�neds/uge godkend
                        if visning = 14 then

                            if lto = "esn" then
                       
                            weekEndDate = DateAdd("d",6,startDato)
                            weekEndDateSQL = year(weekEndDate) &"-"& month(weekEndDate) &"-"& day(weekEndDate)
                            strDatoSQL = year(startDato) &"-"& month(startDato) &"-"& day(startDato)
                            strKommentar = ""
                            antalKommentare = 0
                            strSQLKommentarUge = "SELECT kommentar, dato FROM login_historik WHERE mid = "& intMid &" AND dato BETWEEN '"&strDatoSQL&"' AND '"&weekEndDateSQL&"' AND kommentar <> ''"
                            oRec7.open strSQLKommentarUge, oConn, 3
                            while not oRec7.EOF
                            
                           
                            strKommentar = strKommentar &"<b>"& oRec7("dato") &"</b> - "& oRec7("kommentar") & "<br>" 
                            antalKommentare = antalKommentare + 1

                            oRec7.movenext
                            wend
                            oRec7.close

                                 if media <> "export" then%>
                                        <td style="border-bottom:1px silver solid; border-right:1px silver solid; width:200px; font-size:9px">
                                            <%if antalKommentare <> 0 then %>
                                       
                                                 <%=strKommentar %>
                                           
                                            <%end if %>
                                        </td> 
                                <% 
                                    else

                                        call htmlreplace(strKommentar)
                                        strKommentar = htmlparseTxt
                                        strEksportTxt = strEksportTxt & Chr(34) & strKommentar & Chr(34) &";" 
                                 

                                    end if

                            end if 'esn
                        end if 'visning 14


                        

                         if visning = 14 then '** Afslutning af periode %>
	                    <%
	                  
	        
	                        if cint(showAfsugeVisAfsluttetpaaGodkendUgesedler) = 0 then
	                        showAfsugeTxt = godkendweek_txt_112
	                        else
	                        showAfsugeTxt = ""
	                        end if
	                        
                        if media <> "export" then
                            
                        
                                if cint(SmiWeekOrMonth) = 0 OR (useSogKriAfs = 1 OR useSogKriGk = 1 OR useSogKri = 1) then%>
	                            <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;" align=center><%=showAfsugeTxt %>&nbsp;

                                    <%if cint(splithr) = 1 then
                                        %>
                                        (Month)
                                    <%
                                    end if %>

	                            </td>
                                <%else
                            
                                    'if cint(wth) = 0 then  
                                    'if (datepart("m", slutDatoLastm_B, 2,2) <> lastMth) then
                                    showAfsugeTxt_tot = showAfsugeTxt
                                    'end if%>
                      
                                <%'if lto <> "esn" then %>            
                                <td style="border-bottom:1px silver solid; border-right:1px silver solid;">&nbsp;</td>
                                <%'end if %>
                        
                        
                               <%end if %>

            
                        <%end if %>

                      
	


                           <%
                           
                        '**** KUN Uge eller m�nedsaflsutnigner kan godkendes. Dags niveau kan ikke godkendes af leder 
                        if level <= 2 OR level = 6 then

                           select case ugegodkendt
                           case 0
                                
                                if media <> "print" then

                                    
                                    if cint(showAfsugeVisAfsluttetpaaGodkendUgesedler) = 0 then

                                        select case lto
                                        case "tec", "esn"
                                        gkTxt = godkendweek_txt_101&"/"&godkendweek_txt_107
                                        case else
                                        gkTxt = godkendweek_txt_111
                                        end select

                                        if cint(SmiWeekOrMonth) = 0 then
                                        strCheckBoxGodkenduge = "<input type=""checkbox"" name=""FM_afslutuge_medid_uge"" value='"&intMid &"_"& varTjDatoUS_man_use&"' class='gkuge_"& intMid &"'>" 
                                        end if

                                    else

                                    gkTxt = godkendweek_txt_103

                                        strCheckBoxGodkenduge = ""
                                        strmonthgodkendCHB = ""

                                    end if

                                    btnstyle = 1

                                    
                                    ugegodkendtTxt = strCheckBoxGodkenduge &"&nbsp;<a href=""godkenduge.asp?usemrn="&intMid&"&varTjDatoUS_man="&varTjDatoUS_man_use&""" target=""_blank"" style=""font-size:9px; font-weight:lighter;"">"& gkTxt &" >> </a>"
                                    
                                    if cint(showAfsugeVisAfsluttetpaaGodkendUgesedler) = 0 then
                                        strmonthgodkendCHB = "<input type='checkbox' name='FM_afslutuge_medid_uge' value='"&intMid&"_"&varTjDatoUS_man_use&"' /><br>" 
                                    end if
                                else
                                ugegodkendtTxt = ""
                                strmonthgodkendCHB = ""
                                end if

                               

                           case 1
                           btnstyle = 0
                           call meStamdata(ugegodkendtaf)

                              
                                   call lonKorsel_lukketPer(varTjDatoUS_man_use, -2)
                                   if lonKorsel_lukketIO = 0 then
                                   ugegodkendtTxt = "<a href=""godkenduge.asp?usemrn="&intMid&"&varTjDatoUS_man="&varTjDatoUS_man_use&""" target=""_blank"" style=""font-size:9px; color:yellowgreen;""><i>V</i></a> - "& meInit 
                                   else
                                   ugegodkendtTxt = "<span style=""font-size:9px; color:yellowgreen;""><i>V</i></span> - "& meInit
                                   end if

                           case 2
                               btnstyle = 0
                               call meStamdata(ugegodkendtaf)

                                    call lonKorsel_lukketPer(varTjDatoUS_man_use, -2)
                                    if lonKorsel_lukketIO = 0 then
                                    ugegodkendtTxt = "<a href=""godkenduge.asp?usemrn="&intMid&"&varTjDatoUS_man="&varTjDatoUS_man_use&""" target=""_blank"" style=""font-size:9px; color:red;""><b>!</b></a> - "& meInit
                                    else
                                    ugegodkendtTxt = "<span style=""font-size:9px; color:red;""><b>!</b></span> - "& meInit
                                    end if
                           case else
                            
                                    ugegodkendtTxt = ""


                           end select 


                        else


                           select case ugegodkendt
                           case 0
                           ugegodkendtTxt = ""
                           

                           case 1
                           call meStamdata(ugegodkendtaf)
                           ugegodkendtTxt = "<label style=""color:yellowgreen;""><i>V</i></label> - "& meInit 
                          
                           
                           case 2
                           call meStamdata(ugegodkendtaf)
                           ugegodkendtTxt = "<label style=""color:red;""><b>!</b></label> - "& meInit 
                           
                           
                           case else
                           ugegodkendtTxt = ""
                           
                           end select 


                        end if

                        if media <> "export" then%>
                        <%
                       
                                 if cint(SmiWeekOrMonth) <> 1 OR (useSogKriAfs = 1 OR useSogKriGk = 1 OR useSogKri = 1) then%>
	                             <td class=lille style="border-bottom:1px silver solid; border-right:1px silver solid; white-space:nowrap;"><%=ugegodkendtTxt %>&nbsp;</td>
                                 <%else 
                             
                             
                                     'if cint(wth) = 0 then  
                                    'if (datepart("m", slutDatoLastm_B, 2,2) <> lastMth) then
                                    ugegodkendtTxt_tot = ugegodkendtTxt
                                    'end if
                             
                                     %>            
                                <td style="border-bottom:1px silver solid; border-right:1px silver solid;">&nbsp;</td>
                                 <%end if     
                            
                        end if 

                        
                             
                           'select case ugegodkendt
                           'case 1
                           '  gkTxtExp = "Godkendt"
                           ' case 2
                           '  gkTxtExp = "Afvist"
                           '  case else
                           '  gkTxtExp "-"
                           '  end select%>


	                    <%
                        if showAfsugeTxt = godkendweek_txt_112 then
                            showAfsugeExp = 1
                        else
                            showAfsugeExp = 0
                        end if   
                            
                        strEksportTxt = strEksportTxt & showAfsugeExp &";"& ugegodkendt & ";" 
                            
                            
                       
                        end if '14?
                          
                       %>
                        
                        
                      

	 
	                    <% 

                          
        if media <> "export" then%>
	  </tr>
	  <%
      end if
      end if 'showmedarb %>
	  <!-- slut -->