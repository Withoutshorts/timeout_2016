
<%
    
  public ds, ugp
    sub ds_ugp

                call dageimd(tjkMth, request("yuse"))
                ds = mthDays

                'if mnow <= 12 then
                ugp = ""& ds &"-"& tjkMth &"-"&request("yuse") 'mnow
                'end if

    end sub

    public tjkMth, useMthATD, ddato
    sub mnowgt12

                select case mnow
                case 13
                tjkMth = datepart("m", "1-1-"&request("yuse"), 2,2)
                ddato = "1-1-"&request("yuse")
                'ugp = "31-3-"&request("yuse") 'mnow
                case 14
                tjkMth = datepart("m", "1-4-"&request("yuse"), 2,2)
                ddato = "1-4-"&request("yuse")
                'ugp = "30-6-"&request("yuse") 'mnow
                case 15
                tjkMth = datepart("m", "1-7-"&request("yuse"), 2,2)
                ddato = "1-7-"&request("yuse")
                'ugp = "31-10-"&request("yuse") 'mnow
                case 16
                tjkMth = datepart("m", "1-10-"&request("yuse"), 2,2)
                ddato = "1-10-"&request("yuse")
                'ugp = "31-12-"&request("yuse") 'mnow
                case 17
                tjkMth = datepart("m", "1-1-"&request("yuse"), 2,2) 'useMthATD 'datepart("m", now, 2,2)
                ddato = "1-1-"&request("yuse")
                'ugp = now
                end select        



                'ddato = "1-"& tjkMth & "-" & request("yuse")

    end sub   
    
    
    
    
    

 sub totallinje %>

<tr>

      

	<td style="border-bottom:1px silver dashed; padding-bottom:20px;" class=lille>
        
        <%if cint(SmiWeekOrMonth) = 1 AND mnow > 12 AND (useSogKriAfs = 0 AND useSogKriGk = 0 AND useSogKri = 0) then 'måned
            perThis = left(monthname(lastMth), 3) &" "& year(lastslutDatoLastm_A)
          else
            perThis = "Total"
          end if %>
        
        <b><%=perThis %></b></td>


    <%call timerogminutberegning(anormTimerTot*60) %> 

     <td align=right style="border-bottom:1px silver dashed; padding-bottom:20px; white-space:nowrap;" class=lille><b><%=thoursTot &":"& left(tminTot, 2) %></b></td>

	

	        <%if session("stempelur") <> 0 then %> 
	         <td align=right style="border-bottom:1px silver dashed; padding-bottom:20px; white-space:nowrap;" class=lille>
	         <%call timerogminutberegning(atotalTimerPer100) %>
            <b><%=thoursTot &":"& left(tminTot, 2) %></b>
	        </td>

                <%if showkgtil = 1 then %>
	         <td align=right style="border-bottom:1px silver dashed; padding-bottom:20px; white-space:nowrap;" class=lille><b><%=formatnumber(afradragTimerTot, 2) %></b></td>
	         <td align=right style="border-bottom:1px silver dashed; padding-bottom:20px; white-space:nowrap;" class=lille><b><%=formatnumber(altimerKorFradTot, 2) %></b></td>
             <%end if %>

            <%select case lto
                case "kejd_pb"
                
                case else %>

              <%anormtime_lontimeTot = -((anormTimerTot - (altimerKorFradTot)) * 60) %>
	         <%call timerogminutberegning(anormtime_lontimeTot) %>

             <td align=right style="border-bottom:1px silver dashed; padding-bottom:20px; white-space:nowrap;" class=lille>
	
		        <b><%=thoursTot &":"& left(tminTot, 2) %></b>
	         </td>

           


               <% 
                   
                   normtime_lontimeAkkGT = (normtime_lontimeAkk + (akuPreNormLontBal * 60)) 'normtime_lontimeAkkGT +
                   call timerogminutberegning(normtime_lontimeAkkGT)  %>
                          <td align=right style="white-space:nowrap; border-bottom:1px silver dashed; padding-bottom:20px;" class=lille>
	
		                    <b><%=thoursTot &":"& left(tminTot, 2) %></b>
	                     </td>

	         <%
            end select


	         end if'stempelur %>
	 
	
	 
	
	        
     <td align=right style="border-bottom:1px silver dashed; padding-bottom:20px; white-space:nowrap;" class=lille><b><%=formatnumber(arealTimerTot,2)%></b></td>
	         <%if lto <> "cst" AND lto <> "kejd_pb" AND lto <> "tec" AND lto <> "esn" then %>
	        <td align=right style="border-bottom:1px silver dashed; padding-bottom:20px; white-space:nowrap;" class=lille>(<%=formatnumber(arealfTimerTot,2)%>)</td>
              <%end if %>

                        <%if cint(showkgtil) = 1 AND lto <> "tec" AND lto <> "esn" then %>
                        <td align=right style="white-space:nowrap; border-bottom:1px silver dashed; padding-bottom:20px;" class=lille>(<%=formatnumber(korrektionRealTot,2)%>)</td>
                        <%end if %>
	        
            <%if lto <> "cst" AND lto <> "tec" AND lto <> "esn" then %>
	         <td align=right style="border-bottom:1px silver dashed; padding-bottom:20px; white-space:nowrap;" class=lille><b><%=formatnumber((arealTimerTot - anormTimerTot),2)%></b></td>
               <td align=right style="white-space:nowrap; border-bottom:1px silver dashed; padding-bottom:20px;" class=lille><b><%=formatnumber(balRealNormtimerAkk+(akuPreRealNormBal),2)%></b></td>
             <%end if %>
                

                 <%if session("stempelur") <> 0 AND (lto <> "kejd_pb" AND lto <> "cst" AND lto <> "tec" AND lto <> "esn") then %>
                <td align=right style="border-bottom:1px silver dashed; padding-bottom:20px; white-space:nowrap;" class=lille><b><%=formatnumber((arealTimerTot - (altimerKorFradTot)),2)%></b></td>
	             <%end if %>

          


	    
	      <!-- Afspad / Overarb --->
	       <%if instr(akttype_sel, "#30#") <> 0 OR instr(akttype_sel, "#31#") <> 0 then %>
	         
         <%if lto <> "fk" AND lto <> "kejd_pb" AND lto <> "adra" AND lto <> "cisu" then %>       
	         <td align=right style="border-bottom:1px silver dashed; padding-bottom:20px; white-space:nowrap;" class=lille><b><%=formatnumber(aafspTimerTot, 2)%></b></td>
        <%end if %>

	         <td align=right class=lille style="border-bottom:1px silver dashed; padding-bottom:20px; white-space:nowrap;"><b><%=formatnumber(aafspTimerBrTot, 2)%></b></td>

         <%if lto <> "fk" AND lto <> "kejd_pb" AND lto <> "adra" AND lto <> "cisu" then 
             
                         if lto <> "tec" AND lto <> "esn" then %>
	                     <td align=right class=lille style="border-bottom:1px silver dashed; padding-bottom:20px; white-space:nowrap;"><b><%=formatnumber(aafspTimerUdbTot, 2)%></b></td>
	                        <td align=right class=lille style="border-bottom:1px silver dashed; padding-bottom:20px; white-space:nowrap;"><b><%=formatnumber(aafspadUdbBalTot, 2)%></b></td>
	                        <%end if %>     
	                      <td align=right class=lille style="border-bottom:1px silver dashed; padding-bottom:20px; white-space:nowrap;"><b><%=formatnumber(aAfspadBalTot, 2)%></b></td>
                          <%end if %>
        	 
        
	         <%end if %>

       


           <%'TEC / ESN special  
                   
       select case lto
       case "tec", "esn", "intranet - local"
                     %>
                     <td align=right style="border-bottom:1px silver dashed; padding-bottom:20px;" class=lille><b><%=formatnumber(omsorg2afh_tot, 2) %></b></td>
                     <td align=right style="border-bottom:1px silver dashed; padding-bottom:20px;" class=lille><b><%=formatnumber(sundTimer_tot, 2) %></b></td>
                     <td align=right style="border-bottom:1px silver dashed; padding-bottom:20px;" class=lille><b><%=formatnumber(omsorgKAfh_tot, 2) %></b></td>
                    <%


       end select



      select case lto
       case "esn", "intranet - local"
        %>

           <!--<td align=right style="border-bottom:1px silver dashed; padding-bottom:20px;" class=lille><b><%=formatnumber(flexTimer_tot, 2) %></b></td>-->
           <td align=right style="border-bottom:1px silver dashed; padding-bottom:20px;" class=lille><b><%=formatnumber(sTimer_tot, 2) %></b></td>
           <td align=right style="border-bottom:1px silver dashed; padding-bottom:20px;" class=lille><b><%=formatnumber(adhocTimer_tot, 2) %></b></td>

        <%
    end select

            %>


                 <td align=right class=lille style="border-bottom:1px silver dashed; padding-bottom:20px;">
      
       <b><%=formatnumber(ferieAfVal_md_tot, 2) %></b>
    
       </td>

           <td align=right  class=lille style="border-bottom:1px silver dashed; padding-bottom:20px;">
      
       <b><%=formatnumber(ferieAfulonVal_md_tot, 2) %></b>
    
       </td>
	
	 <td align=right  class=lille style="border-bottom:1px silver dashed; padding-bottom:20px;"><b><%=formatnumber(ferieFriAfVal_md_tot, 2) %></b></td>


         <%   select case lto
                        case "xintranet - local", "fk"
            %>
        	 <td align=right  class=lille style="border-bottom:1px silver dashed; padding-bottom:20px;"><b><%=formatnumber(divfritimer_tot, 2) %></b></td>

        <%
            end select %>

        
        <%   select case lto
                        case "xintranet - local", "fk", "kejd_pb"
            %>
        	 <td align=right  class=lille style="border-bottom:1px silver dashed; padding-bottom:20px;"><b><%=formatnumber(omsorg_tot, 2) %></b></td>

        <%
            end select %>


          <%   select case lto
                        case "xintranet - local", "fk", "kejd_pb"
            %>
        	 <td align=right  class=lille style="border-bottom:1px silver dashed; padding-bottom:20px;"><b><%=formatnumber(tjenestefri_tot, 2) %></b></td>
      

        <%
            end select %>

	 
	   <%   select case lto
                        case "xintranet - local", "fk", "kejd_pb"
            %>
        	 <td align=right  class=lille style="border-bottom:1px silver dashed; padding-bottom:20px;"><b><%=formatnumber(barsel_tot, 2) %></b></td>
      

        <%
            end select %>

         


           <%select case lto
        case "xxintranet - local", "fk"%>
        	 
        <td align=right  class=lille style="border-bottom:1px silver dashed; padding-bottom:20px;"><b><%=formatnumber(lageTimer_tot, 2) %></b></td>

        <%end select

	   select case lto
        case "intranet - local", "adra"%>
        	 
        <td align=right  class=lille style="border-bottom:1px silver dashed; padding-bottom:20px;"><b><%=formatnumber(rejsedage_tot, 2) %></b></td>

        <%end select
    


	 if (level <= 2 OR level <= 6 OR (session("mid") = usemrn)) then%>
	 <td align=right  class=lille style="border-bottom:1px silver dashed; padding-bottom:20px;"><b><%=formatnumber(sygeDage_tot, 2) %></b></td>
	 <%end if %>

          <%if (level <= 2 OR level <= 6 OR (session("mid") = usemrn)) then%>
	 <td align=right  class=lille style="border-bottom:1px silver dashed; padding-bottom:20px;"><b><%=formatnumber(barnSyg_tot, 2) %></b></td>
	 <%end if %>

      
   
    <% if cint(SmiWeekOrMonth) = 0 OR (useSogKriAfs = 1 OR useSogKriGk = 1 OR useSogKri = 1) then %> 
    	<td style="border-bottom:1px silver dashed; padding-bottom:20px; white-space:nowrap;">&nbsp;</td>
	<td style="border-bottom:1px silver dashed; padding-bottom:20px; white-space:nowrap;">&nbsp;</td>
    <%else%>
	<td style="border-bottom:1px silver dashed; padding-bottom:20px; white-space:nowrap;" class="lille" align=center><%=showAfsugeTxt_tot%>&nbsp;</td>
	<td style="border-bottom:1px silver dashed; padding-bottom:20px; white-space:nowrap;" align=center>
        
        <%if len(trim(ugegodkendtTxt_tot)) <> 0 then %>

            <%if cint(btnstyle) = 1 then
            %><span style="border:1px #5582d2 solid; background-color:#d6dff5; font-size:10px; padding:1px;"><%=ugegodkendtTxt_tot %></span><%
            else%>
            <%=ugegodkendtTxt_tot %>
            <%end if %>

        <%end if %>
        
        
        
        &nbsp;</td>
    <%end if %>

	</tr>

   <!-- <tr><td colspan="100"><br />&nbsp;</td></tr> -->

<%

thoursTot = 0 
tminTot = 0

anormTimerTot = 0
atotalTimerPer100 = 0
afradragTimerTot = 0
altimerKorFradTot = 0
anormtime_lontimeTot = 0 
                
                 
normtime_lontimeAkk = 0 
akuPreNormLontBal = 0
                  
arealTimerTot = 0
arealfTimerTot = 0 
korrektionRealTot = 0
                            
anormTimerTot = 0
balRealNormtimerAkk = 0 
akuPreRealNormBal = 0
                   
                   
aafspTimerTot = 0

aafspTimerBrTot = 0 
aafspTimerUdbTot = 0
aafspadUdbBalTot = 0
aAfspadBalTot = 0
omsorg2afh_tot = 0
sundTimer_tot = 0
omsorgKAfh_tot = 0 
sTimer_tot = 0
adhocTimer_tot = 0
      
ferieAfVal_md_tot = 0 
ferieAfulonVal_md_tot = 0
	
ferieFriAfVal_md_tot = 0


divfritimer_tot = 0
omsorg_tot = 0
tjenestefri_tot = 0
barsel_tot = 0
lageTimer_tot = 0 
sygeDage_tot = 0
barnSyg_tot = 0
barnSyg_tot = 0

showAfsugeTxt_tot = ""
ugegodkendtTxt_tot = ""

lastMth = 0
lastMidtjk = 0

'if lastMid <> intMids(m) AND m > 0 then
'    normtime_lontimeAkkGT = 0
'end if


    
end sub %>