

    <%
       '***********************************************************************************************************************************************
       '*** GT i budnen på HR rapporten
       '***********************************************************************************************************************************************

        sub hr_list_grandtotal


       %>


         <!-- Under overskrrifter --> <!--- Akt type navn -->
	 
	 <tr bgcolor="#8CAAe6">
	 
	 <td style="width:200px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid;"><b>Grandtotal:</b></td>
    
         <%if lto = "esn" then %>
     <td style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;" class=alt>&nbsp;</td>
     <td style="width:200px; border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;" class=alt>&nbsp;</td>
	 <%end if %>
	 <%
        

         if instr(akttype_sel, "#-5#") <> 0 OR instr(akttype_sel, "#-1#") then%>
             <td class=lille  valign=bottom bgcolor="#DCF5BD" style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
	        <td class=lille  valign=bottom bgcolor="#DCF5BD" style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td> 
         <%
        end if 
         %>
	 


	  <%if (instr(akttype_sel, "#-5#") <> 0 OR instr(akttype_sel, "#-10#") <> 0) AND stempelurOn = 1 then %>

        <%if instr(akttype_sel, "#-10#") = 0 then %>

   
	 <td class=alt_lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>

     <%if cint(showkgtil) = 1 then %>
	 <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
	 <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
     <%end if %>

	 <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>

         <%end if %>


	 <td class=lille valign=bottom bgcolor="#DCF5BD" style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
    <%end if %>





	 
	 <%if (instr(akttype_sel, "#-1#") <> 0 OR instr(akttype_sel, "#-20#")) <> 0 then
         
         if instr(akttype_sel, "#-20#") = 0 then %>
	 <td class=lille  valign=bottom bgcolor="#cccccc" style="border-right:1px #ffffff solid;">&nbsp;</td>
	 <td class=lille  valign=bottom style="border-right: 1px #ffffff solid;" bgcolor="#cccccc">&nbsp;</td>

      <td class=lille  valign=bottom style="border-right: 1px #ffffff solid;" bgcolor="#999999">&nbsp;</td>
	 <td class=lille  valign=bottom style="border-right: 1px #ffffff solid;" bgcolor="#cccccc">&nbsp;</td>
	 <td class=lille  valign=bottom style="border-right: 1px #ffffff solid;" bgcolor="#cccccc">&nbsp;</td>
    <td class=lille  valign=bottom style="border-right: 1px #ffffff solid;" bgcolor="#999999">&nbsp;</td>

      <%end if %>

	 <td class=lille  valign=bottom bgcolor="pink" style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">
      
        &nbsp;
     </td>
	
	 
    <%end if %>
	 
	 <%if instr(akttype_sel, "#1#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">
	 &nbsp;
	 </td>
	 <%end if %>
	 
	 
	 <%if instr(akttype_sel, "#2#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">
	&nbsp;
	 </td>
	
	<%end if %>
	
	<%if instr(akttype_sel, "#50#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">
	 &nbsp;
	 </td>
	 <%end if %>


     <%if instr(akttype_sel, "#54#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">
	 &nbsp;
	 </td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#51#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">
	&nbsp;
	 </td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#52#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">
	&nbsp;
	 </td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#53#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">
	&nbsp;
	 </td>
	 <%end if %>
	 
	  
	 
	  <%if instr(akttype_sel, "#55#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">
	&nbsp;
	 </td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#60#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">
	    &nbsp;
	 </td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#90#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">
	&nbsp;
	 </td>
	 <%end if %>
	 
	   <%if instr(akttype_sel, "#91#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">
	 &nbsp;
	 </td>
	 <%end if %>

         
	   <%if instr(akttype_sel, "#92#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">
	 &nbsp;
	 </td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#6#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">
	&nbsp;
	 </td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#7#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">
	&nbsp;
	 </td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#5#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;" style="">
	 &nbsp;
	 </td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#10#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">
	 &nbsp;
	 </td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#9#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">
	 &nbsp;
	 </td>
	 <%end if %>
	 


	 
	 <!--Ferie Fridage -->
	 <%if instr(akttype_sel, "#12#") <> 0 then %>
	 <td bgcolor="#FFFF99" class=lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(fefriValTxtGT, 2) %></td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#18#") <> 0 then %>
	 <td bgcolor="#FFFF99" class=lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(fefriplValTxtGT, 2) %></td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#13#") <> 0 then %>
	 <td bgcolor="#FFFF99" class=lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(fefribrValTxtGT, 2) %></td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#17#") <> 0 then %>
	 <td bgcolor="#FFFF99" class=lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(fefriudbValTxtGT, 2) %></td>
	 <%end if %>
	 
	 
	 <%if instr(akttype_sel, "#12#") <> 0 OR instr(akttype_sel, "#18#") <> 0 OR instr(akttype_sel, "#13#") <> 0 OR instr(akttype_sel, "#17#") <> 0 then %>
	 <td bgcolor="#FFF000" class=lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(fefriBalValTxtGT, 2)%></td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#13#") <> 0 AND cint(visikkeFerieogSygiPer) <> 1 then %>
	 <td bgcolor="#EFf3FF" class=lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(fefriPerbrValTxtGT, 2) %></td>
     <td bgcolor="#EFf3ff" class=lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(fefriTimerPerBrTimerValTxtGT, 2) %></td>
     <td bgcolor="#EFf3ff" class=lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
	 <%end if %>


       <%if instr(akttype_sel, "#17#") <> 0 AND cint(visikkeFerieogSygiPer) <> 1 then %>
	 <td bgcolor="#EFf3FF" class=lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(fefriudbValPerTxtGT, 2) %></td>
     <td bgcolor="#EFf3ff" class=lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(fefriudbValPerTimerTxtGT, 2) %></td>
     <%end if %>
	 
	
	 
	 <!-- Ferie -->
	 <%if instr(akttype_sel, "#15#") <> 0 then %>
	 <td bgcolor="#6CAE1C" class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(ferieValTxtExpGT, 2) %></td>
	 <%end if %>

     <%if instr(akttype_sel, "#111#") <> 0 then %>
	 <td bgcolor="#6CAE1C" class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(ferieValoverfortTxtGT, 2) %></td>
	 <%end if %>

     
     <%if instr(akttype_sel, "#112#") <> 0 then %>
	 <td bgcolor="#6CAE1C" class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(ferieValUlonTxtGT, 2) %></td>
	 <%end if %>

	 
	 <%if instr(akttype_sel, "#11#") <> 0 then %>
	 <td bgcolor="#6CAE1C" class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(feriePlValTxtGT, 2) %></td>
	 
	 
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#14#") <> 0 then %>
	 <td bgcolor="#6CAE1C" class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(ferieAftValTxtGT, 2) %></td>
     <%end if %>
	 
	 <%if instr(akttype_sel, "#19#") <> 0 then %>
	 <td bgcolor="#6CAE1C" class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(ferieAftulonValTxtGT, 2) %></td>
	 
	 
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#16#") <> 0 then %>
	 <td bgcolor="#6CAE1C" class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(ferieUdbValTxtGT, 2) %></td>
	 
	 
	 <%end if %>
	 
	 <% if instr(akttype_sel, "#11#") <> 0 OR instr(akttype_sel, "#14#") <> 0 OR instr(akttype_sel, "#19#") <> 0 OR _
	 instr(akttype_sel, "#15#") <> 0 OR instr(akttype_sel, "#16#") <> 0 OR instr(akttype_sel, "#111#") <> 0 OR instr(akttype_sel, "#112#") <> 0 then %>
	  <td bgcolor="#6CA000" class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;" style=""><%=formatnumber(ferieBalValTxtGT, 2)%></td>
	 <%end if %>
	 

	 <!-- Ferie i periode -->

	 <%if instr(akttype_sel, "#14#") <> 0 AND cint(visikkeFerieogSygiPer) <> 1 then %>
	  <td bgcolor="#EFf3ff" class=lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(ferieAftPerValTxtGT, 2) %></td>
      <td bgcolor="#EFf3ff" class=lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(ferieAftPerValtimTxtGT, 2) %></td>
      <td bgcolor="#EFf3ff" class=lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
	 <%end if %>


     <%if instr(akttype_sel, "#19#") <> 0 AND cint(visikkeFerieogSygiPer) <> 1 then %>
	 <td bgcolor="#EFf3ff" class=lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(ferieAftulonPerValTxtGT, 2) %></td>
     <td bgcolor="#EFf3ff" class=lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(ferieAftulonPerValtimTxtGT, 2) %></td>
     <%end if %>


     <%if instr(akttype_sel, "#16#") <> 0 AND cint(visikkeFerieogSygiPer) <> 1 then %>
	 <td bgcolor="#EFf3ff" class=lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(ferieUdbPerValTxtGT, 2) %></td>
     <td bgcolor="#EFf3ff" class=lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(ferieUdbPerValtimTxtGT, 2) %></td>
 
	 <%end if %>

     

    

     <!-- Divfritimer 1 maj timer -->
	  <%if instr(akttype_sel, "#25#") <> 0 then %>
	     <td class=alt_lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(divfritimerTxtGT, 2) %></td>
	  <%end if %>

      <!-- FerieOptjAfh 1.5.2020 - 31.8.2020 -->
	  <%if instr(akttype_sel, "#128#") <> 0 then %>
	     <td class=alt_lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(ferieAfh152020timerTxtGT, 2) %></td>
	  <%end if %>

      <!-- FerieOptjIndefrosst -->
	  <%if instr(akttype_sel, "#129#") <> 0 then %>
	     <td class=alt_lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(ferieIndefrossettimerTxtGT, 2) %></td>
	  <%end if %>
	 

      <!-- Rejsedage -->
	  <%if instr(akttype_sel, "#125#") <> 0 then %>
	     <td class=alt_lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(rejseDageTxtGT, 2) %></td>
	  
      <%end if %>
	 
	 
	 <!-- Afspad --->
	 <%if instr(akttype_sel, "#30#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
	 
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#31#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
     
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#32#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
	 
	 <%end if %>
	 
	<%if instr(akttype_sel, "#33#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
	 
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#30#") <> 0 OR instr(akttype_sel, "#31#") <> 0 OR instr(akttype_sel, "#32#") <> 0 OR instr(akttype_sel, "#33#") <> 0 then %>
	 <td bgcolor="#FFC0CB" class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
	 
	 <%end if %>
	 
	  <%
        if cint(visikkeFerieogSygiPer) <> 1 then 
    
          
              if instr(akttype_sel, "#30#") <> 0 then %>
	         <td bgcolor="#EFf3FF" class=lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
	         <%end if %>
	 
	              <%if instr(akttype_sel, "#31#") <> 0 then %>
	             <td bgcolor="#EFf3FF" class=lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
                 <td bgcolor="#EFf3FF" class=lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
	 
	             <%
                 end if

         end if %>

      
	 
	 
	
	 <!-- Sygdom -->
	  <%if instr(akttype_sel, "#20#") <> 0 then %>
	     <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">
	     <%=formatnumber(sygTimerTxtGT, 2) %>
	     </td>

	      <td class=alt_lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(sygDageTxtGT, 2) %></td>

         <%if cint(visikkeFerieogSygiPer) <> 1 then %>
    	 <td bgcolor="#EFf3FF" class=lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(sygTimerPerTxtGT, 2) %></td>
	     <td bgcolor="#EFf3FF" class=lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(sygDagePerTxtGT, 2) %></td>
           <td bgcolor="#EFf3FF" class=lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
          <%end if 
         end if %>
	 
	  <%if instr(akttype_sel, "#21#") <> 0 then %>

	         <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(barnsygTimerTxtGT, 2) %></td>
	          <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(barnSygDageTxtGT, 2)%></td>

                 <%if cint(visikkeFerieogSygiPer) <> 1 then%>
    	         <td bgcolor="#EFf3FF" class=lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(barnSygTimerPerTxtGT, 2)%></td>
	              <td bgcolor="#EFf3FF" class=lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(barnsygDagePerTxtGT, 2)%></td>
                  <td bgcolor="#EFf3FF" class=lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
         
	 
	         <%end if
       end if %>

	 
	 <%if (instr(akttype_sel, "#20#") <> 0 OR instr(akttype_sel, "#21#") <> 0) AND cint(visikkeFerieogSygiPer) <> 1 then %>
	    <td bgcolor="lightpink" class=lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
     <%end if %>


     <!-- Barsel -->
	  <%if instr(akttype_sel, "#22#") <> 0 then %>
	     <td class=alt_lille valign="bottom" style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;"><%=formatnumber(barselTxtGT, 2) %></td>
         <td class=alt_lille valign="bottom" style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
	  
      <%end if %>

         <!-- Omsorg -->
	  <%if instr(akttype_sel, "#23#") <> 0 then %>
	     <td class=alt_lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">
	    <%=formatnumber(omsorgTxtGT, 2) %>
	     </td>
         <td class=alt_lille valign="bottom" style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
	  
      <%end if %>

         <!-- Senior -->
	  <%if instr(akttype_sel, "#24#") <> 0 then %>
	     <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">
	     <%=formatnumber(seniorTxtGT, 2) %></td>
	  
      <%end if %>

          <!-- Aldersreduktion -->
            <%if instr(akttype_sel, "#27#") <> 0 OR instr(akttype_sel, "#28#") <> 0 OR instr(akttype_sel, "#29#") <> 0 then %>
	        <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
           <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
           <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
           <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>    
           <%end if %>

     

          <%if instr(akttype_sel, "#26#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
	 <%end if %>

          <%if instr(akttype_sel, "#120#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
	 <%end if %>

          <%if instr(akttype_sel, "#121#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
	 <%end if %>

          <%if instr(akttype_sel, "#122#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
	 <%end if %>


          <%if instr(akttype_sel, "#123#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
	 <%end if %>


          <%if instr(akttype_sel, "#124#") <> 0 then %>
	 <td class=alt_lille  valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
	 <%end if %>


	 
	 <%if instr(akttype_sel, "#8#") <> 0 then %>
	 <td class=alt_lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#81#") <> 0 then %>
	 <td class=alt_lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
	 <%end if %>

         <%if instr(akttype_sel, "#115#") <> 0 then %>
	 <td class=alt_lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
	 <%end if %>
	 
	
	 <%if instr(akttype_sel, "#-2#") <> 0 then %>
	 <td class=alt_lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#-3#") <> 0 then %>
	 <td class=alt_lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#61#") <> 0 then %>
	 <td class=alt_lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
	 <%end if %>

       <%if instr(akttype_sel, "#113#") <> 0 then %>
	 <td class=alt_lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
	 <%end if %>
	 

       <%if instr(akttype_sel, "#114#") <> 0 then %>
	 <td class=alt_lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
	 <%end if %>

	 <%if instr(akttype_sel, "#-4#") <> 0 then %>
	  <td class=alt_lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
	  <%end if %>

          <%if instr(akttype_sel, "#-70#") <> 0 then %>
	  <td class=alt_lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
	  <%end if %>

     <%if instr(akttype_sel, "#-30#") <> 0 then %>
	  <td class=alt_lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
          <td class=alt_lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
	  <%end if %>

     <%if instr(akttype_sel, "#-40#") <> 0 then %>
	  <td class=alt_lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
    <td class=alt_lille valign=bottom style="border-right:1px #D6DfF5 solid; border-bottom:1px #D6DfF5 solid; text-align:right;">&nbsp;</td>
	  <%end if %>

	 </tr>


       <%end sub%>