 <%
'***********************************************************************************************
'************** Logind historik / Komme / gå timer                              ****************
'***********************************************************************************************
'public lastMnavn, lastMnr, totalhours, totalmin, totalhoursWeek, totalminWeek, stempelUrEkspTxt, lastMinit
function stempelurlist_2018(medarbSel, showtot, layout, sqlDatoStart, sqlDatoSlut, typ, d_end, lnk)


     'Response.write "browstype_client: " & browstype_client

if media <> "export" then
%>
        
<script language="javascript">


    $(document).ready(function () {




        $(".loginhh").keyup(function () {


            //alert(window.event.keyCode)

            if (window.event.keyCode != '9') {


                var thisid = this.id
                var idlngt = thisid.length
                var idtrim = thisid.slice(12, idlngt)

                eValhh = $("#FM_login_hh_" + idtrim).val()
                eValmm = $("#FM_login_mm_" + idtrim).val()

                $("#FM_kommentar_" + idtrim).attr("disabled", "");




                if (eValhh.length == 1) {

                    if (eValhh > 2) {
                        $("#FM_login_hh_" + idtrim).val('0' + eValhh)
                        eValhh = $("#FM_login_hh_" + idtrim).val()
                    }

                }


                if (eValhh.length == 2 && eValmm.length == 0) {
                    $("#FM_login_mm_" + idtrim).val('00')
                    $("#FM_login_mm_" + idtrim).focus();
                    $("#FM_login_mm_" + idtrim).select();

                }



            }




        });


        $(".logudhh").keyup(function () {



            if (window.event.keyCode != '9') {

                var thisid = this.id
                var idlngt = thisid.length
                var idtrim = thisid.slice(12, idlngt)

                eValhh = $("#FM_logud_hh_" + idtrim).val()
                eValmm = $("#FM_logud_mm_" + idtrim).val()

                $("#FM_kommentar_" + idtrim).attr("disabled", "");

                if (eValhh.length == 1) {

                    if (eValhh > 2) {
                        $("#FM_logud_hh_" + idtrim).val('0' + eValhh)
                        eValhh = $("#FM_logud_hh_" + idtrim).val()
                    }

                }


                if (eValhh.length == 2 && eValmm.length == 0) {
                    $("#FM_logud_mm_" + idtrim).val('00')
                    $("#FM_logud_mm_" + idtrim).focus();
                    $("#FM_logud_mm_" + idtrim).select();

                }

            }

        });


        $("#komme_gaa").submit(function () {

            $(".FM_kommentar").attr("disabled", "");

            //alert("her")
        });




    });

</script>
    <%
	end if 'media


    'sqlDatoStart = day(sqlDatoStart) &"-"& month(sqlDatoStart) &"-"& year(sqlDatoStart)

	'Response.Write "layout: " & layout & "<br>"
    'Response.write "medarbSel: "& medarbSel & "<br>"
	
    '** Fra stempelur historik / komme gå tider stat ***'
	if layout <> 1 then
	medarbSel = 0
	            
	'            for m = 0 to UBOUND(intMids)
	'		    
	'		     if m = 0 then
	'		     medSQLkri = " AND (l.mid = " & intMids(m)
	'		     else
	'		     medSQLkri = medSQLkri & " OR l.mid = " & intMids(m)
	'		     end if
	'		     
	'		    next
	'		    
	'		    medSQLkri = medSQLkri & ")"
			    
	else		    
	Redim intMids(0)

	if medarbSel <> 0 then
	medSQLkri = " AND l.mid = " & medarbSel
	else
	medSQLkri = " AND l.mid <> 0 "
	end if
	
	end if



    
	if media <> "export" then

    '** H&L timer

        select case lto
        case "intranet - local", "cflow"

          'if browstype_client <> "ip" then
          call cflow_hl_timer(medarbSel, sqlDatoStart, sqlDatoSlut)   
          'end if
            
            
       end select



    '*****


	'*** Finder afslutte uger på aktive medarbejdere ***'
	if cdbl(medarbSel) = 0 then
	strSQLmedarb = "SELECT mid FROM medarbejdere WHERE mansat <> '2' AND mansat <> '3' AND mansat <> '4'"
	oRec4.open strSQLmedarb, oConn, 3
	while not oRec4.EOF 
	    
	    call afsluger(oRec4("mid"), sqlDatoStart, sqlDatoSlut) 
	    
	oRec4.movenext
	wend
	oRec4.close
	
	else
	
	call afsluger(medarbSel, sqlDatoStart, sqlDatoSlut) 
	
	end if
	
	end if 'media



    call erStempelurOn() 
  

    if media <> "export" then


	
	if cint(showTot) = 1 then
          
        select case lto
        case "cst"
	    csp = 5
        case else
        csp = 6
        end select
	
    else
	csp = 10
	end if%>
	
	
    <%
    
   
    
    if layout = 1 then
    tblwdt = "100%"
    lnkTarget = "_blank"
   
    hidemenu = "1"
    
  
    else
    
    
    tblwdt = "100%"
     tTop = 10
    tLeft = 0
    tWdth = 1004
    lnkTarget = "_self"
    hidemenu = ""
    call tableDiv(tTop,tLeft,tWdth)
    end if
    
    
    url = "../timereg/stempelur.asp?menu=stat&func=oprloginhist&id=0&medarbSel="&medarbSel&"&showonlyone="&showonlyone&"&hidemenu="&hidemenu
    '&"&rdir="&rdir
    
    text = "Opret ny / pause "
    otoppx = 0
    oleftpx = 0
    owdtpx = 140
    java = "Javascript:window.open('"&url&"','','width=850,height=700,resizable=yes,scrollbars=yes')"
    urlhex = "#"
    
   

    %>

   
   
    <form method="post" id="komme_gaa" action="../timereg/stempelur.asp?func=dbloginhist<%=lnk%>">

        <!--
        <%if media <> "export" AND media <> "print" AND layout = 1 then %>
          <div class="row">
                <div class="col-lg-12">
                    <button type="submit" class="btn btn-success btn-sm pull-right"><b><%=medarb_txt_020 %></b></button> <!-- el. pause -,->
                </div>
            </div>
        <%end if %> -->

        <!-- Kommer l;ngere nede
         <%if media <> "print" AND layout = "1" then %>
            <div class="row">
                <div class="col-lg-12">
                    <a href="#" class="btn btn-success btn-sm pull-right" onclick="Javascript:window.open('<%=url%>','','width=750,height=850,resizable=yes,scrollbars=yes')" class="alt">Tilføj komme/gå </a>
                </div>
            </div>
        <%end if %> -->

  <table class="table dataTable table-bordered ui-datatable" style="border-top:none;">

      <%if media <> "print" AND layout = "1" then %>
      <thead>
          <tr>
              <%if showTot <> 1 then%>

                  <th>Dato</th>
                  <th>Logget ind</th>
                  <th>Logget ud</th>

                  <%select case lto %>
                      <%case "tec", "esn" %>                                         
                      <th>&nbsp</th>
                  <%case else %>
                      <th>IP</th>
                  <%end select %>

                  <th>Indstilling</th>
                  <th>Tid</th>

                <!-- Ligger i systembesked  <%if cint(stempelur_hideloginOn) = 0 then %>
		          <th>Manuelt opr?</th>
                  <%else %>
                  <th>&nbsp;</th>
                  <%end if %> -->

                  <th>System besked<!--<br />Logud ændret--></th>
		          <th>Kommentar</th>
                  <th>&nbsp</th>

              <%else %>
                        
                  <th>Medarbejder</th>
                  <th>Periode</th>
                  <th>Komme/gå (løntimer)</th>
                  <th>Fradrag</th>
                  <th>Ialt</th>

                  <%if lto = "cflow" then %>
                      <th>Overtid</th>
                  <%end if %>

                  <%if lto <> "cflow" AND lto <> "cst" then  %>
                      <th>Projekttid<br />Real. timer</th>
                  <%end if %>

              <%end if %>

          </tr>
      </thead>
      <%end if 'media %>

      <tbody>
      <%
      '****** Uge Loop ****'
      lastMid = 0
      g = 0
      medIDNavnWrt = "#0#"


      totalhoursWeek = 0 
      totalminWeek = 0

        totalhours = 0 
        totalmin = 0

        stempelUrEkspTxto = ""
        stempelUrEkspTxtShowTot = ""

        lastWeek = 0 '"01-01-2002" 'datepart("ww", sqlDatoStart, 2,2) 

        lastDato = sqlDatoStart

          '** valgte medarbejdere 
          for m = 0 to UBOUND(intMids)

       

              if media <> "export" then

                  if cint(layout) = 0 AND browstype_client <> "ip" then
                  '*** Total ***
	                  if lastMid <> intMids(m) AND m > 0 ANd totalhours > 0 then
                          'Response.write "lastMId:"& lastMid &" lmid:"& oRec("lmid")

                                if cint(showTot) <> 1 then             

                              call tottimer_2018(lastMnavn, lastMnr, totalhoursWeek, totalminWeek, lastMid, sqlDatoStart, sqlDatoSlut, 2, lastDato)
		                      totalhoursWeek = 0 
		                      totalminweek = 0
            
                              end if 'showTot
 
		                      call tottimer_2018(lastMnavn, lastMnr, totalhours, totalmin, lastMid, sqlDatoStart, sqlDatoSlut, 1, lastDato)
		                      totalhours = 0 
		                      totalmin = 0
                   
	                  end if
                  end if


              else

                        if cint(showTot) = 1 AND browstype_client <> "ip" then

                             
                              if lastMid <> intMids(m) AND m > 0 ANd totalhours > 0 then

                              call tottimer_2018(lastMnavn, lastMnr, totalhours, totalmin, lastMid, sqlDatoStart, sqlDatoSlut, 1, lastDato)
		                      totalhours = 0 
		                      totalmin = 0

                              end if

                        end if

              end if 'EXPORT

        
	
      if cint(layout) = 0 then
      medSQLkri = "AND l.mid = " & intMids(m)
      end if

    
    
      if cint(showTot) = 1 then
      d_end = 0
      end if
    

      lastDato = "01-01-2002"
    
      if media = "print" then
      td_height = ""
      else
      td_height = "height:30px;"
      end if
    
      for d = 0 to d_end '6

      dtUse = dateAdd("d", d, sqlDatoStart)
    
      'dtUseSQL = 
      'medIDNavnWrt = "#0#"

      if cint(showTot) = 1 then
      dtUseSQL = "l.dato BETWEEN '"& year(sqlDatoStart) &"/"& month(sqlDatoStart) &"/"& day(sqlDatoStart) &"' AND '"& year(sqlDatoSlut) &"/"& month(sqlDatoSlut) &"/"& day(sqlDatoSlut)&"'"
      else
      dtUseSQL = "l.dato = '"& year(dtUse) &"/"& month(dtUse) &"/"& day(dtUse) &"'"
      end if

      if cint(typ) <> 0 then
      strTypSQL = " AND l.stempelurindstilling = "& typ
      else
      strTypSQL = " AND l.stempelurindstilling <> 0 "
      end if
	
	  if cint(showTot) = 1 then
	
	  strSQL = "SELECT l.id AS lid, count(l.id) AS antal, l.mid AS lmid, l.login, l.logud, "_
	  &" s.navn AS stempelurnavn, s.faktor, s.minimum, l.stempelurindstilling, sum(l.minutter) AS minutter, l.dato, manuelt_afsluttet, login_first, logud_first FROM login_historik l"_
	  &" LEFT JOIN stempelur s ON (s.id = l.stempelurindstilling) WHERE "_
      &" "& dtUseSQL &" "& medSQLkri & strTypSQL &""_
	  &" GROUP BY l.mid, l.stempelurindstilling ORDER BY l.mid, l.login, l.stempelurindstilling DESC " 
    
	
	
	  else
	
	  strSQL = "SELECT l.id AS lid, l.mid AS lmid, l.login, l.logud, "_
	  &" s.navn AS stempelurnavn, s.faktor, s.minimum, "_
	  &" l.stempelurindstilling, l.minutter, manuelt_afsluttet, manuelt_oprettet, l.dato, "_
	  &" l.kommentar, logud_first, login_first, l.ipn FROM login_historik l"_
	  &" LEFT JOIN stempelur s ON (s.id = l.stempelurindstilling) WHERE "_
      &" "& dtUseSQL &" "& medSQLkri & strTypSQL &""_
	  &" ORDER BY l.mid, l.login, l.stempelurindstilling DESC" 
	
	  end if
	
	  'AND l.logud IS NOT NULL GROUP BY l.mid, l.dato, l.stempelurindstilling
	  'Response.write "<br><br>::"& medSQLkri &"<br><hr>:"& medarbSel
	  'Response.flush

      ' if lto = "tec" AND (session("mid") = 331 OR session("mid") = 1) then
      ' Response.write strSQL & "<br><br>"
      'end if
	

	  x = 0
  
	  oRec.open strSQL, oConn, 3 
	  while not oRec.EOF 
	
	  'afslSTDato = "2009-1-1"
	  'afslSLDato = "2009-12-31"
	  'call afsluger(oRec("lmid"), afslSTDato, afslSLDato)
	
       call thisWeekNo53_fn(oRec("dato"))

       if media <> "export" then       

            if cint(lastWeek) <> cint(thisWeekNo53) AND (cint(layout) = 0) then
    
                    if cint(layout) = 0 AND browstype_client <> "ip" then
                      '*** uge total ***
	                  if totalhoursWeek > 0 AND d > 0 AND lastMid = oRec("lmid") then
                          'Response.write "lastMId:"& lastMid &" lmid:"& oRec("lmid")

                          if cint(showTot) <> 1 then

		                  call tottimer_2018(lastMnavn, lastMnr, totalhoursWeek, totalminWeek, lastMid, sqlDatoStart, sqlDatoSlut, 2, lastDato)
		                  totalhoursWeek = 0 
		                  totalminweek = 0

                          end if
                      
                       
	                  end if
              end if

      end if

      end if

	  thours = 0
	  tmin = 0


      '****** Minutter / time beregning *****
          timerThis = 0
		  timerThisDIFF = 0
		
       
		  if len(oRec("login")) <> 0 AND len(oRec("logud")) <> 0 then
		    
		      if cint(oRec("stempelurindstilling")) = -1 then
		        
		          timerThisDIFF = oRec("minutter")
		          useFaktor = 1
		        
		      else
		    
		          timerThisDIFF = oRec("minutter")
    		
		          if timerThisDIFF < oRec("minimum") then
			          timerThisDIFF = oRec("minimum")
		          end if
    		    
		          if oRec("faktor") > 0 then
			      useFaktor = oRec("faktor")
			      else
			      useFaktor = 0
			      end if
			
			  end if
		
		
		  timerThis = (timerThisDIFF * useFaktor)
		  'totaltimer = totaltimer + timerThis
		  end if
		
		  'Response.write timerThis
		  'call timerogminutberegning(0,timerThis)
		
		  timTemp = formatnumber(timerThis/60, 3)
		  timTemp_komma = split(timTemp, ",")
		
		  for x = 0 to UBOUND(timTemp_komma)
			
			  if x = 0 then
			  thours = timTemp_komma(x)
			  end if
			
			  if x = 1 then
			  tmin = timerThis - (thours * 60)
			  end if
			
		  next
		
		  if cint(oRec("stempelurindstilling")) = -1 then
		  totalhours = totalhours - cint(thours)
		  totalmin = totalmin - tmin

          totalhoursWeek = totalhoursWeek - cint(thours)
		  totalminWeek = totalminWeek - tmin

		  else
		  totalhours = totalhours + cint(thours)
		  totalmin = totalmin + tmin
		
          totalhoursWeek = totalhoursWeek + cint(thours)
		  totalminWeek = totalminWeek + tmin
        
          end if


      
		
		  if len(tmin) <> 0 then
		
			  tmin = tmin
			
			  if tmin = 0 then
			  tmin = "00"
			  end if
			
			  if len(tmin) = 1 then
			  tmin = "0"&tmin
			  end if
			
		  else
		  tmin = "00"
		  end if

    
              'if session("mid") = 1 then
              'Response.write "thours: " &thours& "tmin: " & tmin & " tot: "& totalhours &"+"& totalmin & "<br>"
              'end if

       

      if cint(oRec("stempelurindstilling")) = -1 then 
		  timerMinExpTxt = "-"&thours&":"&left(tmin, 2)
          minutterExpTxt = "-"&oRec("minutter")
		  else
          timerMinExpTxt = thours&":"&left(tmin, 2)
          minutterExpTxt = oRec("minutter")
      end if
	
      '** Slut timeberegning
	  '******************************************


      '**** System besked ************************
      sysBeskedTxt = ""
      select case cint(oRec("manuelt_afsluttet"))
		  case 1 
        
		  '** tider manuelt ændret **'
		
		  sysBeskedTxt = "Ja:"
		
		      if len(trim(oRec("login_first"))) <> 0 then 
			  sysBeskedTxt = sysBeskedTxt & "("&left(formatdatetime(oRec("login_first"), 3), 5)
			  end if
			
			  if len(trim(oRec("logud_first"))) <> 0 then 
			  sysBeskedTxt = sysBeskedTxt & "-"&left(formatdatetime(oRec("logud_first"), 3), 5)&")"
			  end if 

           
			
		
		  case 2 '** Glemt at logge ud ***'
       
		  sysBeskedTxt = "Glemt at logge ud!"
		
		
		  case 3 'Tidertilret pga stempelur indstillinger og tider manuelt ændret **'
		  sysBeskedTxt = "Ja:"
		
		
		      if len(trim(oRec("login_first"))) <> 0 then 
			  sysBeskedTxt = sysBeskedTxt & "("&left(formatdatetime(oRec("login_first"), 3), 5)
			  end if 
			
			  if len(trim(oRec("logud_first"))) <> 0 then 
			  sysBeskedTxt = sysBeskedTxt & "-"&left(formatdatetime(oRec("logud_first"), 3), 5)&")"
			  end if 
			
			  sysBeskedTxt = sysBeskedTxt &"<br />Logind tilpasset pga.<br /> stempelur indstilllinger."
		
		
		
		  case else 
		
		  end select 
            
            
          '**** SLUT system besked ****





              if media <> "export" then    
                
                      if lastMid <> oRec("lmid") then
    
                              if instr(medIDNavnWrt, "#"&oRec("lmid")&"#") = 0 then
                    
                              call meStamdata(oRec("lmid"))

            
        
                                  dtUseTxt = dateadd("d", 3, dtUse) '** Sikere det er mid i uge, hvis ugen løber over årsskift
       
                                      if cint(showTot) = 1 then 'total
        
                                   
        
                                      else%>
	
                                      <tr style="display:none">
		                                  <!--<td>&nbsp;</td> -->
		                                  <td colspan="<%=csp%>" style="height:20px; padding:5px;"><h4><span style="font-size:11px; font-weight:lighter;">Komme / Gå tid </span><br />
            
                                              <%if len(trim(meInit)) <> 0 then %>
                                              <%=meNavn & "  ["& meInit &"]"%>
                                              <%else %>
                                              <%=meTxt%>
                                              <%end if %>
                                            </h4>
		                                  </td>
		                                  <!-- <td>&nbsp;</td> -->
	                                  </tr>


	                                  <%end if


   


                              medIDNavnWrt = medIDNavnWrt & ",#"& oRec("lmid") & "#" 
                              end if
    
                      end if

      else

                  if lastMid <> oRec("lmid") then
                  if instr(medIDNavnWrt, "#"&oRec("lmid")&"#") = 0 then


                  call meStamdata(oRec("lmid"))

                  end if

                  end if

                  if isNull(oRec("login")) <> true then
                  loginExpTxt = formatdatetime(oRec("login"), 3)
                  else
                  loginExpTxt = ""
                  end if

                  if isNULL(oRec("logud")) <> true then
                  logudExpTxt = formatdatetime(oRec("logud"), 3)
                  else
                  logudExpTxt = ""
                  end if


                  if cint(showTot) <> 1 then
                  stempelUrEkspTxt = stempelUrEkspTxt & ""& meNavn &";"& meInit & ";" & oRec("dato") & ";" & loginExpTxt & ";"& logudExpTxt & ";"& timerMinExpTxt &";"& minutterExpTxt &";" & useFaktor &";" & oRec("ipn") & ";" & oRec("stempelurindstilling") &";" & oRec("manuelt_oprettet") &";"& chr(34) & sysBeskedTxt & chr(34) &";"& chr(34) & oRec("kommentar") & chr(34) &";" & "xx99123sy#z"
                  end if

                  medIDNavnWrt = medIDNavnWrt & ",#"& oRec("lmid") & "#" 

      end if 'media
                
                
      if media <> "export" then            
      %>
	  <!-- <tr>
		  <td colspan="<%=csp+2%>" style="height:1px;"></td>
	  </tr> -->
	
      <%end if


      


      if media <> "export" AND cint(showTot) <> 1 then

    
      call thisWeekNo53_fn(oRec("dato"))
      if (cint(lastWeek) <> cint(thisWeekNo53)) AND (cint(layout) = 0) then

      call thisWeekNo53_fn(dtUse)
      %>
	
	  <tr>
		  <!--<td>&nbsp;</td> -->
		  <td colspan="<%=csp%>" style="<%=td_height%> padding:5px;"><b>Uge <%=thisWeekNo53 &" "& datepart("yyyy", dtUse, 2,2) %> </b></td>
		  <!-- <td>&nbsp;</td> -->
	  </tr>
	
	  <%
      end if



	
	  if cdbl(lastID) = oRec("lid") then
	  bgThis = "#ffff99"
	  else
	      'if cint(oRec("stempelurindstilling")) = -1 then
	      'bgThis = "#c4c4c4"
	      'else
          'select case right(g, 1)
          'case 0,2,4,6,8
	      bgThis = "#ffffff"
          'case else
          'bgThis = "#8CAAE6"
          'end select
	      'end if
	  end if
	
	
	  if len(trim(oRec("dato"))) <> 0 then
	  loginDTShow = left(weekdayname(weekday(oRec("dato"))), 3) & ". " & formatdatetime(oRec("dato"), 2)
	  else
	  loginDTShow = "--"
	  end if%>


          <% 

                    'tjekker helligdag
                    call helligdage(dtUse, 0, lto, usemrn)

                    if datepart("w", dtUse, 2,2) = 6 or datepart("w", dtUse, 2,2) = 7 or erHellig = 1 then
                    bgcolorFRI = "#e8e8e8"
                    else
                    bgcolorFRI = ""
                    end if
                    %>


                    <tr style="background:<%=bgcolorFRI%>">

           

              <td style="white-space:nowrap; text-align:center;">
                  <%if showTot = 1 then%>
			          <%=oRec("antal")%> stk.&nbsp;
		          <%else
                      '** er periode godkendt ***'
		              tjkDag = oRec("dato")
                      call thisWeekNo53_fn(tjkDag)
		              erugeafsluttet = instr(afslUgerMedab(oRec("lmid")), "#"&thisWeekNo53&"_"& datepart("yyyy", tjkDag) &"#")
                
                      strMrd_sm = datepart("m", tjkDag, 2, 2)
                      strAar_sm = datepart("yyyy", tjkDag, 2, 2)

                      call thisWeekNo53_fn(tjkDag)
                      strWeek = thisWeekNo53 'datepart("ww", tjkDag, 2, 2)
                      strAar = datepart("yyyy", tjkDag, 2, 2)

                      if cint(SmiWeekOrMonth) = 0 then
                      usePeriod = strWeek
                      useYear = strAar
                      else
                      usePeriod = strMrd_sm
                      useYear = strAar_sm
                      end if

                      'erugeAfslutte(
                      call erugeAfslutte(useYear, usePeriod, oRec("lmid"), SmiWeekOrMonth, 0, tjkDag)
		        
		              
		              call lonKorsel_lukketPer(tjkDag, -2, oRec("lmid"))
		         
                     

                        '*** tjekker om uge er afsluttet / lukket / lønkørsel
                      call tjkClosedPeriodCriteria(tjkDag, ugeNrAfsluttet, usePeriod, SmiWeekOrMonth, splithr, smilaktiv, autogk, autolukvdato, lonKorsel_lukketIO, ugegodkendt)
				
                
                      '* Admin får vist stipledede bokse
                      if cint(level) = 1 AND ugeerAfsl_og_autogk_smil = 1 then
                      boxStyleBorder = "border: 1px #999999 dashed;"
                      else
                      boxStyleBorder = ""
                      end if

                
                      if (ugeerAfsl_og_autogk_smil = 0 OR level = 1 OR erTeamlederForVilkarligGruppe = -100) AND media <> "print" then%>
			          <a href="#" onclick="Javascript:window.open('../timereg/stempelur.asp?id=<%=oRec("lid")%>&menu=stat&func=redloginhist&medarbSel=<%=medarbSel%>&showonlyone=<%=showonlyone%>&hidemenu=<%=hidemenu%>&rdir=popup','','width=850,height=700,resizable=yes,scrollbars=yes')" class="vmenu"><%=loginDTShow%></a>
			          <%else %>
		              <%=loginDTShow%>
		              <%end if %>
		
		          <%end if 'showtot%>             
              </td>

              <%if showTot <> 1 then%>
		    
		          <%if cint(oRec("stempelurindstilling")) <> -1 then %>
                      <td style="text-align:center; white-space:nowrap;">
                    
                              <%if len(trim(oRec("login"))) <> 0 then
                    
                              loginTidThisHH = left(formatdatetime(oRec("login"), 3), 2)
                              loginTidThisMM = mid(formatdatetime(oRec("login"), 3),4,2)

                                  if len(trim(loginTidThisHH)) <> 0 then
                                  loginTidThisHH_opr = loginTidThisHH
                                  loginTidThisMM_opr = loginTidThisMM
                                  else
                                  loginTidThisHH_opr = "00"
                                  loginTidThisMM_opr = "00"
                                  end if
                    
                              else

                              loginTidThisHH = ""
                              loginTidThisMM = ""
                                loginTidThisHH_opr = "00"
                              loginTidThisMM_opr = "00"

			                  end if  
                    
                      if (layout = 1) AND (ugeerAfsl_og_autogk_smil = 0 OR level = 1 OR erTeamlederForVilkarligGruppe = -100) AND media <> "print" then

                            %>
                          <input type="hidden" value="<%=oRec("lid") %>" name="id" />
                          <input type="hidden" value="<%=oRec("lmid") %>" name="mid" />
                          <input type="hidden" value="<%=formatdatetime(dtUse, 2) %>" name="logindato" />
                       

                          
                 
                          <input type="text" class="loginhh form-control input-small" name="FM_login_hh" id="FM_login_hh_<%=d %>" value="<%=loginTidThisHH%>" style="width:40px; display:inline-block;" /> :
                          <input type="text" class="form-control input-small" name="FM_login_mm" id="FM_login_mm_<%=d %>" value="<%=loginTidThisMM%>" style="width:40px; display:inline-block;" />
                          &nbsp;</td>  
			
                          <input type="hidden" name="oprFM_login_hh" id="Hidden5" value="<%=loginTidThisHH_opr%>">
		                  <input type="hidden" name="oprFM_login_mm" id="Hidden6"  value="<%=loginTidThisMM_opr%>">
	    
	                        <!-- bruges til arrray split -->
	                        <input type="hidden" name="FM_login_hh" id="Hidden7" value="#">
		                    <input type="hidden" name="FM_login_mm" id="Hidden8"  value="#"> 
                      <%else %>
                              <%=loginTidThisHH%>:<%=loginTidThisMM%>
                  
                      <%end if 'layout%> 

                      <td style="text-align:center; white-space:nowrap;">            
                      <%			
                          if oRec("logud") = oRec("login") then
                          fColsl = "#cccccc"%>
                          <%else 
                          fColsl = "#000000"%>
                          <%end if 
                
                
                          if len(oRec("logud")) <> 0 then 
                          logudTidThisHH = left(formatdatetime(oRec("logud"),3),2) 
                          logudTidThisMM = mid(formatdatetime(oRec("logud"),3),4,2) 

                          if len(logudTidThisHH) <> 0 then
                          logudTidThisHH_opr = logudTidThisHH
                          logudTidThisMM_opr = logudTidThisMM
                          else
                          logudTidThisHH_opr = "00"
                          logudTidThisMM_opr = "00"
                          end if

                          else
                          logudTidThisHH = ""
                          logudTidThisMM = ""

                          logudTidThisHH_opr = "00"
                          logudTidThisMM_opr = "00"

                          end if

                          if (layout = 1) AND (ugeerAfsl_og_autogk_smil = 0 OR level = 1 OR erTeamlederForVilkarligGruppe = -100) AND media <> "print"  then
                    
                          %>
                            <input type="text" class="logudhh form-control input-small" name="FM_logud_hh" id="FM_logud_hh_<%=d %>" value="<%=logudTidThisHH%>" style="width:40px; display:inline-block; color:<%=fColsl%>; <%=boxStyleBorder%>" /> :
                            <input type="text" class="form-control input-small" name="FM_logud_mm" id="FM_logud_mm_<%=d %>" value="<%=logudTidThisMM%>" style="width:40px; display:inline-block; color:<%=fColsl%>; <%=boxStyleBorder%>" />
                            <input type="hidden" name="oprFM_logud_hh" id="Hidden1" value="<%=logudTidThisHH_opr%>">
		                    <input type="hidden" name="oprFM_logud_mm" id="Hidden2"  value="<%=logudTidThisMM_opr%>">
	    
	                        <!-- bruges til arrray split -->
	                        <input type="hidden" name="FM_logud_hh" id="Hidden3" value="#">
		                    <input type="hidden" name="FM_logud_mm" id="Hidden4"  value="#">  

                          <%else %>
                          <%=logudTidThisHH%>:<%=logudTidThisMM%>
                          <%end if %>
                      </td> 

                      <%select case lto
                          case "tec", "esn"
                              %><td>&nbsp;</td><%
                          case else %>
		                      <td><%=oRec("ipn") %></td>
                      <%end select %>

                  <%else 'stempelurindstilling%>

                  <td>&nbsp;</td>
		          <td>&nbsp;</td>
		          <td>&nbsp;</td>

                  <%end if %>

              <%end if %>


              <td>
                  <%      
                  if cint(oRec("stempelurindstilling")) <> -1 then 

                          if (layout = 1) AND (ugeerAfsl_og_autogk_smil = 0 OR level = 1 OR erTeamlederForVilkarligGruppe = -100) AND media <> "print"  then
        
                        %>
                              <select name="FM_stur" class="form-control input-small" style="width:100px;">
		                      <%if lto <> "fk_bpm" AND lto <> "kejd_pb" AND lto <> "kejd_pb2" then %>
                              <!--<option value="0">Ingen (slet)</option>-->
                              <%end if %>
		                      <%
		                      strSQL5 = "SELECT id, navn, faktor, minimum FROM stempelur ORDER BY navn"
		                      oRec5.open strSQL5, oConn, 3 
		                      while not oRec5.EOF 
		
		                      if oRec("stempelurindstilling") = oRec5("id") then
		                      sel = "SELECTED"
		                      else
		                      sel = ""
		                      end if
		                      %>
		                      <option value="<%=oRec5("id")%>" <%=sel%>><%=oRec5("navn")%> (<%=oRec5("faktor")%> / <%=oRec5("minimum")%> min.)</option>
		
		                      <%
		                      oRec5.movenext
		                      wend
		                      oRec5.close 

        
		                      %>
                              </select>
               
            
               
                          <%else
                
            

            
		                      strSQL5 = "SELECT id, navn, faktor, minimum FROM stempelur WHERE id = "& oRec("stempelurindstilling") &" ORDER BY navn"
		                      oRec5.open strSQL5, oConn, 3 
		                      if not oRec5.EOF then
		
		
		                      %>
                              <input type="hidden" name="FM_stur" value="<%=oRec5("id") %>" /> 
		                    <%=oRec5("navn")%> (<%=oRec5("faktor")%> / <%=oRec5("minimum")%> min.) 
		        
		                      <%
		    
		                      end if
		                      oRec5.close 

                          end if

               

                  else
                  %>
                  Pause
           
                  <%
                  end if %>
		      </td>

              <td style="text-align:right">
		          <%if cint(oRec("stempelurindstilling")) = -1 then %>
                  <%=timerMinExpTxt%>
		          <!---<%=thours%>:<%=left(tmin, 2)%>-->
		          <%else %>
                  <%=timerMinExpTxt %>
		          <!--<%=thours%>:<%=left(tmin, 2)%>-->
                  <%end if %>
		      </td>

              <%if showTot <> 1 then %>

                  <!-- <td> Kommer frem i systembesked i stedet
		              <%if oRec("manuelt_oprettet") <> 0 AND cint(stempelur_hideloginOn) = 0 then%> 
		              <span style="color:#999999; font-size:9px;">Ja</span>
		              <%else %>
		              &nbsp;
		              <%end if %>
		          </td> -->

                  <td>
                      <span style="color:#999999; font-size:9px;"><%=sysBeskedTxt %></span>

                      <%if oRec("manuelt_oprettet") <> 0 AND cint(stempelur_hideloginOn) = 0 then %>
                          <span style="color:#999999; font-size:9px;">Manu. opr. af </span>
                      <%end if %>

                        <%
                          select case lto
                          case "cflow"
                          case "esn", "tec", "cst"
                              call findesDerTimer(1, oRec("dato"), medarbSel)
                          case else
                              call findesDerTimer(1, oRec("dato"), medarbSel)
                          end select
             
                          %>
		          </td>

              <%end if %>


              <%if showTot <> 1 then%>

                  <td style="width:300px;">
		
                      <%
                      if cint(oRec("stempelurindstilling")) <> -1 then 
            
                          if (layout = 1) AND (ugeerAfsl_og_autogk_smil = 0 OR level = 1 OR erTeamlederForVilkarligGruppe = -100) AND media <> "print"  then

                

                          %>
		                  <textarea id="FM_kommentar_<%=d %>" name="FM_kommentar" style="width:100%; resize:none; height:25px;" class="form-control input-small"><%=oRec("kommentar") %></textarea>
                          <input type="hidden" name="FM_kommentar" id="Hidden13"  value="#">  
                          <%else %>
                          <%=left(oRec("kommentar"), 100) %>
                          <%end if %>
                      <!-- &nbsp; -->
                      <%else %>
                      <!-- &nbsp; -->
                          <%if len(trim(oRec("kommentar"))) <> 0 then %>
                          <%=left(oRec("kommentar"), 100) %>
                          <%end if %>
		    
                      <%end if %>
		          </td>

                  <td>
		                <% if ((ugeerAfsl_og_autogk_smil = 0 AND showTot <> 1 AND len(oRec("logud")) <> 0) OR (level = 1 AND showTot <> 1)) AND media <> "print" then%>
		                <a href="#" onclick="Javascript:window.open('../timereg/stempelur.asp?menu=<%=menu%>&func=sletlogind&id=<%=oRec("lid")%>&medarbSel=<%=medarbSel%>&showonlyone=<%=showonlyone%>&hidemenu=<%=hidemenu%>&rdir=<%=rdir%>', '', 'width=400,height=200,resizable=no')" class="slet"><span style="color:darkred;" class="fa fa-times"></span></a>
		                <%else %>
		                &nbsp;
		                <%end if %>      
		          </td>

              <%end if %>

              <!-- <td>&nbsp</td> -->

          </tr>

          <%
              'Response.write "lastMid b:" & lastMid & "<br>"
              end if 'media + showtotal
              call thisWeekNo53_fn(oRec("dato"))
              lastWeek = thisWeekNo53 'datepart("ww", oRec("dato"), 2,2)
	          lastMnavn = meNavn
	          lastMnr = meNr
	          lastMid = oRec("lmid")
              'lastMenavnStempelur =  meNavn & "  ["& meInit &"]"
              lastMinit = meInit

              lastDato = oRec("dato")

	          x = x + 1
	          g = g + 1
	          Response.Flush
	          oRec.movenext
	          wend
	          oRec.close 

              if d > 1 then 'hvis der ikke er nogen komme/gå tider skal der alligevel tjekkes for overtid

                  if intMids(m) <> 0 then
                  thisIntMids = intMids(m)
                  else
                  thisIntMids = usemrn
                  end if

                  call fn_overtid(lto, dtUse, thisIntMids, 1)

              end if

          %>


          <%
              if media <> "export" then

              '*** Kun fra egen logind historik ***'
              if cint(layout) = 1 then

    
              '*** Forvalgt stempelurr ****'
              call fv_stempelur() 
          %>

              <%if x = 0 then %>

                  <%if instr(medIDNavnWrt, "#"&medarbsel&"#") = 0 then
    
    
                  call meStamData(medarbsel)
                   
                
                  dtUseTxt = dateadd("d", 3, dtUse) '** Sikere det er mid i uge, hvis ugen løber over årsskift
                 
    
                      medIDNavnWrt = medIDNavnWrt & ",#"& medarbsel & "#" 
                      end if
    
    
                      '** er periode godkendt ***'
                      'if session("mid") = 331 then
                      'tjkDag = dateadd("d", 3, dtUse) '** Periode altid torsdag, pga månedsskift / årsskift
                      'else
		              tjkDag = dtUse
                      'end if
                      call thisWeekNo53_fn(tjkDag)

                      if cint(SmiWeekOrMonth) = 0 then
		              erugeafsluttet = instr(afslUgerMedab(medarbsel), "#"&thisWeekNo53&"_"& datepart("yyyy", tjkDag) &"#")
                      else
                      erugeafsluttet = instr(afslUgerMedab(medarbsel), "#"&datepart("m", tjkDag,2,2)&"_"& datepart("yyyy", tjkDag) &"#")
                      end if
		        

                      strMrd_sm = datepart("m", tjkDag, 2, 2)
                      strAar_sm = datepart("yyyy", tjkDag, 2, 2)
                      strWeek = thisWeekNo53 'datepart("ww", tjkDag, 2, 2)
                      strAar = datepart("yyyy", tjkDag, 2, 2)

                      if cint(SmiWeekOrMonth) = 0 then
                      usePeriod = strWeek
                      useYear = strAar
                      else
                      usePeriod = strMrd_sm
                      useYear = strAar_sm
                      end if

                
                      call erugeAfslutte(useYear, usePeriod, medarbsel, SmiWeekOrMonth, 0, tjkDag)
		        
                      call lonKorsel_lukketPer(tjkDag, -2, medarbsel)

                    
                      '*** tjekker om uge er afsluttet / lukket / lønkørsel
                      call tjkClosedPeriodCriteria(tjkDag, ugeNrAfsluttet, usePeriod, SmiWeekOrMonth, splithr, smilaktiv, autogk, autolukvdato, lonKorsel_lukketIO, ugegodkendt)



                          'if session("mid") = 1 then

		                  'Response.Write "smilaktiv: "& smilaktiv & "<br>"
		                  'Response.Write "SmiWeekOrMonth: "& SmiWeekOrMonth &" ugeNrAfsluttet: "& ugeNrAfsluttet & " tjkDag: "& tjkDag &"<br>"
		                  'Response.Write "autolukvdatodato: "& autolukvdatodato & "<br>"
		                  'Response.Write "tjkDag: "& tjkDag & "<br>"
		                  'Response.Write "autolukvdato: "& autolukvdato & "<br>"
		                  'Response.Write "erugeafsluttet:" & erugeafsluttet & "<br>"
                          'Response.write "ugeerAfsl_og_autogk_smil: "& ugeerAfsl_og_autogk_smil & "<br>"
                          'Response.write "erTeamlederForVilkarligGruppe: " & erTeamlederForVilkarligGruppe & "<br>"
		        
                          'end if


                      '* Admin får vist stipledede bokse
                      if cint(level) = 1 AND ugeerAfsl_og_autogk_smil = 1 then
                      boxStyleBorder = "border: 1px #999999 dashed;"
                      else
                      boxStyleBorder = ""
                      end if
                
    
                      '** erTeamlederForVilkarligGruppe = 1 20171115, ellers var den -100. Teamleder må gerne se
                      '** Således at TEAMleder kan se ferie kommentarer i uger hvor der ikke er logget ind.

                      if (ugeerAfsl_og_autogk_smil = 0 OR level = 1 OR erTeamlederForVilkarligGruppe = 1) AND media <> "print" then
    
                      %>
  
                      <!-- <tr>
		                  <td colspan="<%=csp+2%>" style="height:1px;"></td>
	                  </tr> -->

                      <% 

                    'tjekker helligdag
                    call helligdage(dtUse, 0, lto, usemrn)

                    if datepart("w", dtUse, 2,2) = 6 or datepart("w", dtUse, 2,2) = 7 or erHellig = 1 then
                    bgcolorFRI = "#e8e8e8"
                    else
                    bgcolorFRI = ""
                    end if
                    %>


                    <tr style="background:<%=bgcolorFRI%>">
                          <input type="hidden" value="0" name="id" />
                          <input type="hidden" value="<%=usemrn%>" name="mid" />
                          <input type="hidden" value="<%=formatdatetime(dtUse, 2) %>" name="logindato" />
                          <!--<input type="hidden" value="<%=forvalgt_stempelur %>" name="FM_stur" />-->
                          <!--<td>&nbsp</td>-->
                          <td style="<%=td_height%> white-space:nowrap; text-align:center;"><%=left(weekdayname(weekday(dtUse)), 3) %>. <%=formatdatetime(dtUse, 2) %></td>
                                
                          <td style="white-space:nowrap; text-align:center;">
                              <%if (ugeerAfsl_og_autogk_smil = 0 OR level = 1) then %>
                              <input type="text" class="loginhh form-control input-small" name="FM_login_hh" id="FM_login_hh_<%=d %>" value="" style="width:40px; display:inline-block; <%=boxStyleBorder%>" /> :
                              <input type="text" class="form-control input-small" name="FM_login_mm" id="FM_login_mm_<%=d %>" value="" style="width:40px; display:inline-block; <%=boxStyleBorder%>" />
                              <%end if %>
                                &nbsp; 
                          </td>

                          <td style="white-space:nowrap; text-align:center;">
                                <%if (ugeerAfsl_og_autogk_smil = 0 OR level = 1) then %>
                                <input type="text" class="logudhh form-control input-small" name="FM_logud_hh" id="FM_logud_hh_<%=d %>" value="" style="width:40px; display:inline-block; <%=boxStyleBorder%>" /> :
                              <input type="text" class="form-control input-small" name="FM_logud_mm" id="FM_logud_mm_<%=d %>" value="" style="width:40px; display:inline-block; <%=boxStyleBorder%>" />
                                <%else %>
                                &nbsp;
                                <%end if %>   
                          </td>

                          <td style="width:20px;">&nbsp;</td>

                          <td>           
                              <%if (ugeerAfsl_og_autogk_smil = 0 OR level = 1) then %>    
                              <select name="FM_stur" class="form-control input-small" style="width:100px;">
		                      <%if lto <> "fk_bpm" AND lto <> "kejd_pb" AND lto <> "kejd_pb2"  then %>
                              <!--<option value="0">Ingen</option>-->
                              <%end if %>
		
                              <%
		                      strSQL5 = "SELECT id, navn, faktor, minimum, forvalgt FROM stempelur ORDER BY navn"
		                      oRec5.open strSQL5, oConn, 3 
		                      while not oRec5.EOF 
		
		                      if cint(oRec5("forvalgt")) = 1 then
		                      sel = "SELECTED"
		                      else
		                      sel = ""
		                      end if
		                      %>
		                      <option value="<%=oRec5("id")%>" <%=sel%>><%=oRec5("navn")%> (<%=oRec5("faktor")%> / <%=oRec5("minimum")%> min.)</option>
		
		                      <%
		                      oRec5.movenext
		                      wend
		                      oRec5.close 

                              %>
                              </select>
                              <%end if %>
                          </td>

                          <td>&nbsp;</td>
                          <!-- Manulet oprettet kollonnen er fjernet er ligger i systembesked i stedet <td>&nbsp;</td> -->

                          <td>
                              <%
                              select case lto
                              case "cflow"
                              case "esn", "tec", "cst"
                                  call findesDerTimer(0, dtUse, medarbSel)
                              case else
                              end select%>
                          </td>

                        <td style="width:300px;"><textarea id="FM_kommentar_<%=d %>" class="FM_kommentar form-control input-small" name="FM_kommentar" style="width:100%; resize:none; height:25px;"></textarea>
                        <input type="hidden" name="FM_kommentar" id="Hidden18"  value="#">  </td>
                        <td>&nbsp;</td>
		                <!-- <td>&nbsp;</td> -->

                    </tr>

                      <input type="hidden" name="oprFM_login_hh" id="Hidden16" value="00">
		              <input type="hidden" name="oprFM_login_mm" id="Hidden17"  value="00">
                
                      <input type="hidden" name="FM_login_hh" id="Hidden14" value="#">
		              <input type="hidden" name="FM_login_mm" id="Hidden15"  value="#">  


                    <input type="hidden" name="oprFM_logud_hh" id="Hidden9" value="00">
		            <input type="hidden" name="oprFM_logud_mm" id="Hidden10"  value="00">
	    
	                <!-- bruges til arrray split -->
	                <input type="hidden" name="FM_logud_hh" id="Hidden11" value="#">
		            <input type="hidden" name="FM_logud_mm" id="Hidden12"  value="#">  


                      <%
                          end if '** ugeafsluttet


                          end if '** x=0

                          end if ' layout

                          end if 'media


                          next 'dage/datoer

                          next 'intmmids 



                          '** Overtid ***
                          if cint(layout) = 0 AND d > 1 then

                              if lastMid <> 0 then
                              lastMid = 0
                              else
                              lastMid = usemrn
                              end if

                              call fn_overTid(lto, lastDato, lastMid, 1)
                          end if

                          if media = "export" AND showTot <> 1 then
                              stempelUrEkspTxt = stempelUrEkspTxt & "xx99123sy#z" & stempelUrEkspTxto
                          end if

                          if media <> "export" then %>
                                    

                          <!---- Oprettelse af ny komme g[ tid ---->
                          <%
                          url = "../timereg/stempelur.asp?menu=stat&func=oprloginhist&id=0&medarbSel="&medarbSel&"&showonlyone="&showonlyone&"&hidemenu="&hidemenu

                          if media <> "print" AND layout = "1" then
                          %>
                            <!--
                          <tr>
                              <td  colspan="20" style="text-align:right; border:0px;">

                                    <button type="submit" class="btn btn-success btn-sm pull-right"><b><%=medarb_txt_020 %></b></button> <br /><br />&nbsp;
                              </td>

                          </tr>-->
                          <tr>
                              <td colspan="20" style="text-align:center;">                                                                                 
                                  <a class="btn btn-default btn-sm" onclick="Javascript:window.open('<%=url%>','','width=750,height=850,resizable=yes,scrollbars=yes')" style="width:50%"><b>Tilføj komme/gå </b></a>
                              <br /><br />&nbsp;</td>
                          </tr>
                          <%end if %>




                          <% if g = 0 AND layout <> 1 then%>
	                          <!-- <tr>
		                          <tdcolspan="<%=csp+2%>" style="height:1px;"></td>
	                          </tr> -->
	                          <tr>
		                          <!--<td>&nbsp;</td>-->
		                          <td colspan="<%=csp%>"><br><br>Der findes <b>ikke</b> nogen komme/gå tider for de(n) valgte medarbejder(e) i den valgte periode.<br><br>&nbsp;</td>
		                          <!--<td style="border-right:1px #8caae6 solid;">&nbsp;</td> -->
	                          </tr>
	                          <%
	                      end if %>

                          </tbody>

                              <% 
                               if browstype_client <> "ip" then
                                  
                                      if cint(layout) <> 1 AND totalhours <> 0 OR totalmin <> 0 then 
		                              'Response.write "lastMnavn: " & lastMnavn & " lastMid:"& lastMid &" lastMnr. "&lastMnr &" totalhours "& totalhours  &" ¤"& totalmin &"datoer:"& sqlDatoStart &","& sqlDatoSlut
                                      'Response.end

                                      if cint(showTot) <> 1 then

                                          call tottimer_2018(lastMnavn, lastMnr, totalhoursWeek, totalminWeek, lastMid, sqlDatoStart, sqlDatoSlut, 2, lastDato)
                                          totalminWeek = 0
                                          totalhoursWeek = 0

                                      end if

		                              call tottimer_2018(lastMnavn, lastMnr, totalhours, totalmin, lastMid, sqlDatoStart, sqlDatoSlut, 1, lastDato)
		                              totalhours = 0 
		                              totalmin = 0

	                              end if
                                  
                              end if%>

                              <!-- <tr>
		                          <td bgcolor="#cccccc" colspan="<%=csp+2%>" style="height:1px;"></td>
	                          </tr> -->

                              <!-- <tr bgcolor="#5582d2">
		                          <td width="8" valign=top height=20 style="border-bottom:1px #8caae6 solid; border-left:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		                          <td colspan=<%=csp%> valign="top" style="border-bottom:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		                          <td align=right valign=top style="border-bottom:1px #8caae6 solid; border-right:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	                          </tr> -->

                              <%if media <> "print" AND layout = 1 then %> 
                              <!-- <tr>
		                          <td width="8"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		                          <td colspan=<%=csp%> align=right><br /><input type="submit" class="opdaterliste" value="Opdater liste >>"</td>
		                          <td><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	                          </tr> -->
                              <%end if %>


                          <%end if %>

            

  </table>
   
    
  <%if media <> "export" AND media <> "print" AND layout = 1 then %>

        
        <div class="row">
          <div class="col-lg-12">

              <%select case lto
              case "cflow", "intranet - local"

               'if session("mid") = 1 OR session("mid") = 26 OR session("mid") = 36 then
                call medariprogrpFn(session("mid"))
                if (instr(medariprogrpTxt, "#14#") <> 0 OR instr(medariprogrpTxt, "#16#") <> 0 OR instr(medariprogrpTxt, "#3#") <> 0 OR instr(medariprogrpTxt, "#19#") <> 1) OR level = 1 then

                %>
              <input type="checkbox" value="1" name="FM_opdater_hl" /> Tilpas H&L timer til ovenstående komme/gå tider
              <br />Sletter nuværende H&L timer denne uge og indlæser nye.
              <%end if

              end select%>

              <button type="submit" class="btn btn-success btn-sm pull-right"><b><%=medarb_txt_020 %></b></button>
          </div>
      </div>


  <%end if %>
         
    
</form>



<%
else
        
  if cint(showTot) = 1 AND browstype_client <> "ip" then
      call tottimer_2018(lastMnavn, lastMnr, totalhours, totalmin, lastMid, sqlDatoStart, sqlDatoSlut, 1, lastDato)
	totalhours = 0 
	totalmin = 0
  end if

end if 'media
        
end function
%>