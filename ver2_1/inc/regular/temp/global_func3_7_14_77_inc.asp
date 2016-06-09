<%
'call treg_medarb_afstem_saldo 'cls_afstem
	  


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
       
	   
	   
       ltimerKorFrad = 0
       if cint(showkgtil) = 1 then
	   ltimerKorFrad = ((totalTimerPer100/60) + (fradragTimer(x)))
       else
       ltimerKorFrad = (totalTimerPer100/60)
       end if

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
	   
	   balRealNormtimer = (realTimer(x) - normTimer(x))
	   if len(trim(balRealNormtimer)) <> 0 then
	   balRealNormtimer = balRealNormtimer
	   else
	   balRealNormtimer = 0
	   end if
	    
	   showthisMedarb = 1 
	   
	   if visning = 14 then
	   '** ekstra søgekriterier ***'
	        call ekstrasogKri(useSogKri,moreorless,timeKri,saldoKri,visning)
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
	  
	   
	    select case visning 
        case 7 
            if media <> "export" then%>
	    <td style="border-bottom:1px silver solid; white-space:nowrap;" class=lille>
        <%if media <> "print" then %>
        <a href="afstem_tot.asp?usemrn=<%=intMid %>&show=77&varTjDatoUS_man=<%=startDato %>&varTjDatoUS_son=<%=startDato %>" class=rmenu><%=monthname(datepart("m", startDato,2,2)) &" "& year(startDato) %></a>
        <%else %>
        <b><%=monthname(datepart("m", startDato,2,2)) &" "& year(startDato) %></b>
        <%end if %>
        </td>
	    <%
        end if

	    strEksportTxt = strEksportTxt & monthname(datepart("m", startDato,2,2)) &" "& year(startDato) &";"
        case 77 
        
        if media <> "export" then%>
	    <td style="border-bottom:1px silver solid; white-space:nowrap;" class=lille align=right>
         <%if media <> "print" then %>
        <a href="ugeseddel_2011.asp?usemrn=<%=intMid %>&varTjDatoUS_man=<%=varTjDatoUS_man_use%>" class=rmenu target="_blank"><%=left(weekdayname(datepart("w", startDato,1,2)), 2) %> d. <%=formatdatetime(startDato, 2) %></a>
        <!-- weekpage_2010.asp?medarbid= &st_dato day(varTjDatoUS_man_use)&"/"&month(varTjDatoUS_man_use)&"/"&year(varTjDatoUS_man_use) &func=us-->
           
        <%else %>
        <b><%=left(weekdayname(datepart("w", startDato,1,2)), 2) %> d. <%=formatdatetime(startDato, 2) %></b>
        <%end if %>
        </td>
	    <%
        end if

	    strEksportTxt = strEksportTxt & formatdatetime(startDato, 2) & ";"
	    
	    case else
        
        if media <> "export" then%>
	    <td style="border-bottom:1px silver solid; white-space:nowrap;" class=lille>
	         <%if media <> "print" then %>
	        <a href="ugeseddel_2011.asp?usemrn=<%=intMid %>&varTjDatoUS_man=<%=varTjDatoUS_man_use%>" class="rmenu" target="_blank"><%=datepart("ww", startDato,2,2) %></a>
            <!-- weekpage_2010.asp?medarbid=<%=intMid %>&st_dato=<%=stDato%>&func=us -->
	        <%else %>
	        <%=datepart("ww", startDato,2,2) %>
	        <%end if %>
	        
	        - <%=datepart("d", startDato,2,2) & ". " & left(monthname(datepart("m", startDato,2,2)), 3) %>
	        
	        </td>
	    
	    <%
        end if

	    strEksportTxt = strEksportTxt & datepart("ww", startDato,2,2) & ";"
	    end select
	 
	 
	 
	 
	  if media <> "export" then
      
      %>
	    <td align=right style="border-bottom:1px silver solid; white-space:nowrap;" class=lille>
        <% if realTimer(x) <> 0 then %>
        <%=formatnumber(realTimer(x),2)%>
        <%else %>
        &nbsp;
        <%end if %></td>
        
        <%end if %>
    <%strEksportTxt = strEksportTxt & formatnumber(realTimer(x),2) & ";" %>

         <%if lto <> "cst" AND lto <> "kejd_pb" then 
      
          if media <> "export" then%>
	         <td align=right class=lille style="border-bottom:1px silver solid; white-space:nowrap;">
              <%if realfTimer(x) <> 0 then %>
             (<%=formatnumber(realfTimer(x),2)%>)
             <%else %>
             &nbsp;
             <%end if %>
             </td>
     


         <%end if %>
	     <%arealfTimerTot = arealfTimerTot + (realfTimer(x)) %>
	     <%strEksportTxt = strEksportTxt & formatnumber(realfTimer(x),2) & ";" %>
	     <%end if %> 
    

    <% if media <> "export" then %>
        <td align=right style="border-bottom:1px silver solid; white-space:nowrap;" class=lille>
        <%if normTimer(x) <> 0 then %>
        <%=formatnumber(normTimer(x),2)%>
        <%else %>
        &nbsp;
        <%end if %>
        </td>
    
    
    <%end if %>
    <%strEksportTxt = strEksportTxt & formatnumber(normTimer(x),2) & ";" %>
	
	<%arealTimerTot = arealTimerTot + (realTimer(x)) %>
	<%anormTimerTot = anormTimerTot + (normTimer(x)) %>
	 
	<%if session("stempelur") <> 0 then %> 
                 
                 <%call timerogminutberegning(totalTimerPer100) %>
                 
                         <%if media <> "export" then %>
	                     <td align=right style="border-bottom:1px silver solid; white-space:nowrap;" class=lille>
                
	             
	             
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
                         <%end if %>

                <%strEksportTxt = strEksportTxt & thoursTot &"."& left(tminTot, 2) & ";" %>
              

    <%if cint(showkgtil) = 1 then 
    
                 if media <> "export" then%>
                    <td align=right style="border-bottom:1px silver solid; white-space:nowrap;" class=lille>
                    <%if fradragTimer(x) <> 0 then %>
                    <%=formatnumber(fradragTimer(x), 2) %>
                    <%else %>
                    &nbsp;
                    <%end if %></td>
	             
                 
                 <td align=right style="border-bottom:1px silver solid; white-space:nowrap;" class=lille>
                 
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
	 
	 end if 'stempelur%>
	 
	 
	 <%if lto <> "cst" AND lto <> "kejd_pb" then 
          if media <> "export" then%>
	     <td align=right style="border-bottom:1px silver solid; white-space:nowrap;" class=lille bgcolor="#FFC0CB"><b><%=formatnumber(balRealNormtimer,2)%></b></td>
	     <%end if %>
     <%strEksportTxt = strEksportTxt & formatnumber(balRealNormtimer, 2) & ";" %>
	 
	  <%
	 end if %>
	  

	  
	  <%
      '**BAL Løntimer / Nomeret
      if session("stempelur") <> 0 then 
       call timerogminutberegning(normtime_lontime) 
        if media <> "export" then%> 
	 <td bgcolor="#DCF5BD" align=right style="border-bottom:1px silver solid; white-space:nowrap;" class=lille>
	
		<b><%=thoursTot &":"& left(tminTot, 2) %></b>
	 </td>
	 <%
     end if

     strEksportTxt = strEksportTxt & thoursTot &"."& left(tminTot, 2) & ";" %>
	 
	 
	    <%
        
        '**BAL Real / Løntimer 
        if lto <> "cst" AND lto <> "kejd_pb" then 
          if media <> "export" then%>
	   <td align=right style="border-bottom:1px silver solid; white-space:nowrap;" class=lille> <%=formatnumber(balRealLontimer,2)%></td>
	    <%
        end if
        strEksportTxt = strEksportTxt & formatnumber(balRealLontimer,2) & ";" %>
	 
	    
	    <%end if %>
	 <%end if %>
	 
	 
	     <!-- Afspad / Overarb --->
	       <%if instr(akttype_sel, "#30#") <> 0 OR instr(akttype_sel, "#31#") <> 0 then 
           
                     if media <> "export" then%>
	           
	                 <td align=right style="border-bottom:1px silver solid; white-space:nowrap;" class=lille>
                     <%if afspTimer(x) <> 0 then %>
                     <%=formatnumber(afspTimer(x), 2)%>
                     <%else %>
                     &nbsp;
                     <%end if 
                     
             
                     %></td>
	         
             
                     <td align=right class=lille style="border-bottom:1px silver solid; white-space:nowrap;">
                     <%if afspTimerBr(x) <> 0 then %>
                     <%=formatnumber(afspTimerBr(x), 2)%>
                     <%else %>
                     &nbsp;
                     <%end if
                     
                      %></td>

                          <%
                     end if 

                    strEksportTxt = strEksportTxt & formatnumber(afspTimer(x), 2) &";"
                    strEksportTxt = strEksportTxt & formatnumber(afspTimerBr(x), 2) &";"
             
                  
                  if lto <> "lw" AND lto <> "kejd_pb" then 
                  
                         if media <> "export" then%>
	                     <td align=right class=lille style="border-bottom:1px silver solid; white-space:nowrap;">
                                 <%if afspTimerUdb(x) <> 0 then %>
                                 <%=formatnumber(afspTimerUdb(x), 2)%>
                                 <%else %>
                                 &nbsp;
                                 <%end if %>

                         
                         </td>
	     
	                      <% end if
                          strEksportTxt = strEksportTxt & formatnumber(afspTimerUdb(x), 2) &";"



	              aafspTimerTot = aafspTimerTot + afspTimer(x) 
	              aafspTimerBrTot = aafspTimerBrTot + afspTimerBr(x)
	              aafspTimerUdbTot = aafspTimerUdbTot + afspTimerUdb(x)
	          
	             afspadUdbBal = 0
	             afspadUdbBal = (afspTimerOUdb(x) - afspTimerUdb(x)) 
	         
	             aafspadUdbBalTot = aafspadUdbBalTot + (afspadUdbBal)
	             
                   if media <> "export" then%>
	              <td align=right class=lille style="border-bottom:1px silver solid; white-space:nowrap;">
                          <%if afspadUdbBal <> 0 then %>
                          <%=formatnumber(afspadUdbBal, 2)%>
                          <%else %>
                          &nbsp;
                          <%end if %>
                  </td>
	              <%end if 
	             
                 AfspadBal = 0 
	             AfspadBal = (afspTimer(x) - (afspTimerBr(x)+ afspTimerUdb(x)))
	             aAfspadBalTot = aAfspadBalTot + (AfspadBal)
	              
                           if media <> "export" then%>
	                      <td align=right class=lille style="border-bottom:1px silver solid; white-space:nowrap;">
                          <%if AfspadBal <> 0 then %>
                          <%=formatnumber(AfspadBal, 2)%>
                          <%else %>
                          &nbsp;
                          <%end if %></td>
        	             <%end if %>
	         
	             <% strEksportTxt = strEksportTxt & formatnumber(afspadUdbBal, 2) &";" & formatnumber(AfspadBal, 2) &";" %>
	              <%end if %>
	         
	         <%end if %>


	         
	                    <%if visning = 14 then %>
	                    <%
	                    sidsteDag = year(dateadd("d", 6, startDato)) & "-"& month(dateadd("d", 6, startDato)) & "-"& day(dateadd("d", 6, startDato))
	                    call erugeAfslutte(datepart("yyyy", startDato,2,2), datepart("ww", sidsteDag, 2, 2), intMid) 
	        
	                        if cint(showAfsuge) = 0 then
	                        showAfsugeTxt = "Ja"
	                        else
	                        showAfsugeTxt = "Nej"
	                        end if
	                        
                             if media <> "export" then%>
	                    <td class=lille style="border-bottom:1px silver solid; white-space:nowrap;" align=center><%=showAfsugeTxt %></td>
                        <%end if %>

                           <%select case ugegodkendt
                           case 0
                           case 1
                           ugegodkendtTxt = "Godkendt"
                           case 2
                           ugegodkendtTxt = "Afvist" 
                           end select 
                           
                        if media <> "export" then%>
                         <td class=lille style="border-bottom:1px silver solid; white-space:nowrap;" align=center><%=ugegodkendtTxt %>&nbsp;</td>
                        <%end if %>


	                    <%strEksportTxt = strEksportTxt & showAfsugeTxt &";"& ugegodkendtTxt & ";" %>
	 
	                    <%end if %>
	        
	        
	 <%if normTimerDag(x) <> 0 then
	 ferieAfVal_md = (ferieAFPerTimer(x)+ferieAFulonPer(x)+ferieUdbPer(x))/normTimerDag(x)
	 else
	 ferieAfVal_md = 0
	 end if
	 
	 ferieAfVal_md_tot = ferieAfVal_md_tot + ferieAfVal_md

     
     if normTimerDag(x) <> 0 then
	 ferieFriAfVal_md = fefriTimerPerBr(x)+fefriTimerUdbPer(x)/normTimerDag(x)
	 else
	 ferieFriAfVal_md = 0
	 end if
	 
	 ferieFriAfVal_md_tot = ferieFriAfVal_md_tot + ferieFriAfVal_md

	  %>
	 
	                     <%if visning = 7 OR visning = 77 then 
                          if media <> "export" then%>
	                     <td class=lille style="border-bottom:1px silver solid; white-space:nowrap;" align=right>

                       

                         <%if ferieAfVal_md <> 0 then %>
                         <%=formatnumber(ferieAfVal_md,2) %>
                         <%else %>
                         &nbsp;
                         <%end if %></td>

	                     <td class=lille style="border-bottom:1px silver solid; white-space:nowrap;" align=right>
                         <%if ferieFriAfVal_md <> 0 then %>
                         <%=formatnumber(ferieFriAfVal_md,2) %>
                         <%else %>
                         &nbsp;
                         <%end if %></td>       
	                       <%end if
                                
                                strEksportTxt = strEksportTxt & formatnumber(ferieAfVal_md,2) &";"& formatnumber(ferieFriAfVal_md,2) & ";"
                                  
                                  if level = 1 OR (session("mid") = usemrn) then %>
	         
	                              <%
              
                                  if normTimerDag(x) <> 0 then 
	                              sygDage_barnSyg = (sygDage(x) + barnSygDage(x)) / normTimerDag(x) 
	                              else 
	                              sygDage_barnSyg = 0
	                              end if
	          
	                              sygDage_barnSyg_tot = sygDage_barnSyg_tot + sygDage_barnSyg
	                              
                                   if media <> "export" then%> 
	           
	                             <td align=right class=lille style="border-bottom:1px silver solid; white-space: nowrap;">
                                 <%if sygDage_barnSyg <> 0 then %>
                                 <%=formatnumber(sygDage_barnSyg,2) %>
                                 <%else %>
                                 &nbsp;
                                 <%end if %></td>
	                             <%end if 
                                   

                                   strEksportTxt = strEksportTxt & formatnumber(sygDage_barnSyg,2) &";"

                                  end if 'level %>
	         
	         
	       
	        
	                     <%end if '*Visning
                         
                          
        if media <> "export" then%>
	  </tr>
	  <%
      end if
      end if 'showmedarb %>
	  <!-- slut -->