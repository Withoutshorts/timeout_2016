
<%
   

sub showafslutuge_ugeseddel

          '** Variable der benyttes i denne SUB
          'usemrn
          'varTjDatoUS_man
          'smilaktiv
          'media
          'meType
          'afslutugekri
          'timerdenneuge_dothis       
    
  
         if cint(smilaktiv) <> 0 AND media <> "print" then

         varTjDatoUS_son = dateAdd("d", 6, varTjDatoUS_man) 'sldatoSQL 'dateAdd("d", 7, sldatoSQL) 'FEJL ved �rsskioft, detteaar bliver forkert

         'call medrabSmilord(usemrn) '** F�lger kontrolpanel. Kan fjernes. 20161004
         
         '** Skal smiley visess agrresiv og skal smiley ikon vises
         '** smiley_agg, hidesmileyicon
          call smiley_agg_fn()        
         
         '** Skal der afsluttes p� Dag:2/UGE:0/MD:1/, Hvilken dag skal der v�re afsluttet og godkendt
         '** SmiWeekOrMonth, SmiantaldageCount, SmiantaldageCountClock, SmiTeamlederCount 
         call smileyAfslutSettings()

         call meStamdata(usemrn)     '** ME stamdata

          '** Henter den sensete dag/uge/md der ER afsluttet - s� man ved hvilken periode der st�r for tur. ***'
          '** sidsteUgenrAfsl, sidsteUgenrAfslFundet
          call afsluttedeUger(year(now), usemrn, 0)

         '** Er kriterie for afslutuge m�dt? Ifht. medatype mindstimer og der m� ikke v�re herrel�se timer fra. f.eks TT
         '** akttypeKrism, afslutugekri, afslutugekri_proc, tjkTimeriUgeDt, tjkTimeriUgeDtDay, tjkTimeriUgeSQL, tjkTimeriUgeDtTxt, afslugeDatoTimerudenMatch
         call timeKriOpfyldt(lto, sidsteUgenrAfsl, meType, usemrn, SmiWeekOrMonth)

         'response.write "tjkTimeriUgeDt: " & tjkTimeriUgeDt & " // " & tjkTimeriUgeSQL


         '** timerdenneuge_dothis indstilling for timerDenneUge( 
         if len(trim(timerdenneuge_dothis)) <> 0 then
         timerdenneuge_dothis = timerdenneuge_dothis
         else
         timerdenneuge_dothis = 0
         end if

         select case SmiWeekOrMonth
         case 0,1
         timerdenneuge_dothis = 0 
         case 2
         timerdenneuge_dothis = 1
         end select
    
        '*** Hvor mange timer er der tastet i den peridoe der skal afsluttes
        '*** dag: DD ELLER DEN N�STE DAG DER SKAL AFSLUTTES
         call timerDenneUge(usemrn, lto, tjkTimeriUgeSQL, akttypeKrism, timerdenneuge_dothis, SmiWeekOrMonth)
     
        
         '** Periode aflsutning Dag / UGE / MD
         select case cint(SmiWeekOrMonth) 
         case 0
         weekMonthDate = datepart("ww", varTjDatoUS_son,2,2)
         case 1
         weekMonthDate = datepart("m", varTjDatoUS_son,2,2)
         case 2
         weekMonthDate = formatdatetime(now, 2) ' SKAL DET ALTID V�RE DD?
         weekMonthDate = year(weekMonthDate) & "-" & month(weekMonthDate) & "-" & day(weekMonthDate)
         end select

        '** Er periode afsluttet?
        '** ugeNrAfsluttet, showAfsuge, cdAfs, ugegodkendt, ugegodkendtaf, ugegodkendtTxt, ugegodkendtdt
        call erugeAfslutte(datepart("yyyy", varTjDatoUS_son,2,2), weekMonthDate, usemrn, SmiWeekOrMonth, 0) 

       
        
        '** Faneblade med afslutnings status
        select case cint(SmiWeekOrMonth) 
        case 0, 1

        '*** Viser faneblad til afslutning af periode (s1_k)
        call smileyAfslutBtn(SmiWeekOrMonth)

        '*** Viser besked/kvittering om at uge er afsluttet, hvis den er afsluttet.
        call ugeAfsluttetStatus(varTjDatoUS_son, showAfsuge, ugegodkendt, ugegodkendtaf, SmiWeekOrMonth) 

        end select



        '**** Hvis periode IKKE er afsluttet: ****'
        '**** Afslut uge ***'
	    '**** Smiley vises hvis sidste uge ikke er afsluttet, og dag for afslutninger ovwerskreddet ***' 12 
         '** Periode aflsutning Dag / UGE / MD
         select case cint(SmiWeekOrMonth) 
         case 0
         denneUgeDag = datePart("w", now, 2,2)
         s0Show_sidstedagsidsteuge = dateAdd("d", -denneUgeDag, now) 'now
         case 1
         s0Show_sidstedagsidsteuge = now
         case 2
         s0Show_sidstedagsidsteuge = now ' SKAL DET ALTID V�RE DD?
         s0Show_sidstedagsidsteuge = year(s0Show_sidstedagsidsteuge) & "-" & month(s0Show_sidstedagsidsteuge) & "-" & day(s0Show_sidstedagsidsteuge)
         end select
       
      
         '** finder kriterie for rettidig afslutning
         call rettidigafsl(s0Show_sidstedagsidsteuge, 0)

        
         '*** Finder dato p� den peridoeder skal afsluttes
         select case cint(SmiWeekOrMonth) 
         case 0
         s0Show_weekMd = datePart("ww", s0Show_sidstedagsidsteuge, 2,2)
         case 1
         s0Show_weekMd = datePart("m", s0Show_sidstedagsidsteuge, 2,2)
         case 2
         s0Show_weekMd = s0Show_sidstedagsidsteuge 'formatdatetime(s0Show_sidstedagsidsteuge,2)
         's0Show_weekMd = year(s0Show_weekMd) & "-" & month(s0Show_weekMd) & "-" & day(s0Show_weekMd)
         end select

        'response.write "s0Show_weekMd: " & s0Show_weekMd

        
        '** tjekker om uge er afsluttet og viser afsluttet eller form til afslutning
		call erugeAfslutte(year(s0Show_sidstedagsidsteuge), s0Show_weekMd, usemrn, SmiWeekOrMonth, 0)

       
      
     
            
	        '**** Afslut uge ***'
	        '**** Smiley vises om mandagen, f�rste gang me logger p�, hvis sidste uge ikke er afsluttet ***'
	        if (cDate(formatdatetime(now, 2)) >= cDate(formatdatetime(cDateUge, 2)) AND (datepart("w", now, 2,2) = 1 AND datepart("h", now, 2,2) < 12) AND cint(ugeNrStatus) = 0) OR cint(smiley_agg) = 1 then
            'if (datepart("w", now, 2,2) = 1 AND datepart("h", now, 2,2) <= 23 AND session("smvist") <> "j") OR cint(smiley_agg) = 1 then
    	    smVzb = "visible"
    	    smDsp = ""
    	    session("smvist") = "j"
    	    showweekmsg = "j"
    	    else
    	    smVzb = "hidden"
    	    smDsp = "none"
    	    showweekmsg = "n"
    	    end if
    	
    	

            select case thisfile
            case "stempelur"
            sDivWth = "100%"
            case else
            sDivWth = "725"
            end select      

              
            'smVzb = "visible"
            'smDsp = ""
            'Response.write "thisfile: " & thisfile
    	    
            '*** Auto popup ThhisWEEK now SMILEY
	        %>
	        <!--<div id="s0" class="well" style="position:relative; left:0px; top:26px; width:<%=sDivWth%>px; visibility:<%=smVzb%>; display:<%=smDsp%>; z-index:2; background-color:#ffffff; padding:20px;">-->
            <br />&nbsp;<div id="s0" class="well" style="width:800px;">
	        <%

            '*** Viser denne uge
            'weekSelectedThis = dateAdd("d", 7, now) 

            '*** Medarbejder og overskift, status og smiley Ikon
	        'call showsmiley(weekSelectedThis, 1, usemrn, SmiWeekOrMonth)

            '**** Kriterie
            '*** afslutugeBasisKri, afslProc, afslProcKri, weekSelectedTjk
            'Response.write "tjkTimeriUgeDt: " & tjkTimeriUgeDt

            call afslutkri(varTjDatoUS_son, tjkTimeriUgeDt, usemrn, lto, SmiWeekOrMonth)


            
               
                        'if cint(afslutugekri) = 0 OR ((cint(afslutugekri) = 1 OR cint(afslutugekri) = 2) AND cint(afslProcKri) = 1) OR cint(level) = 1 then 
            
                            select case thisfile
                            case "stempelur"
                            rdir = "stempelur"
                            case "favorit.asp"
                            rdir = "favorit"
                            case else
                            rdir = "ugeseddel"
                            end select
                     
                            'Response.write "XXX weekSelectedTjk: " & weekSelectedTjk & "<br>"

                            '*** Vis checkbox og submuit til at afslutte periode
                            call afslutuge(weekSelectedTjk, 1, varTjDatoUS_son, rdir, SmiWeekOrMonth)

                        'else
                       

                         '********************************************************************
                         '** Status p� antal registrerede projekttimer i den p�g�ldende uge    

                         '** B�r s�ttes i kontrolpanel
                         '********************************************************************
                         'call smiley_statusTxt


                        'end if
	        
	                   
                            
                      '********************************************************************     
                      '*** Viser allerede afsluttede peridoer 
                      '********************************************************************
                      if cint(SmiWeekOrMonth) <> 2 then 'Afslut p� dag. Giver ikke mening at vise. Uoverskueligt%>
               
	                  <br /><br />
	                    <span id="se_uegeafls_a" style="color:#5582d2;">[+] <%=tsa_txt_402 & " "& year(varTjDatoUS_son) %></span><br /><%

                        varTjDatoUS_ons = dateAdd("d", -3, varTjDatoUS_son)

	                    useYear = year(varTjDatoUS_ons)
                        '** Hvilke uger er afsluttet '***
	                    call smileystatus(usemrn, 1, useYear)
	        
                      end if
                
                
             %>
              
	        

               
	        </div>
        
        <br /><br />
                    
        <%                
      end if 'PRINT / SMILEY = 1 
        


end sub








'*********************************************************************************************************
'*********************************************************************************************************
'*** Ugeseddel 2011, dag for dag registreringer


'*********************************************************************************************************
'*********************************************************************************************************

public fmLink, timertot, ucount
function ugeseddel(medarbid, stdatoSQL, sldatoSQL, visning, showtotal, showheader)


timerThisDayLontimerTotTxt = 0

st_dato = day(stdatoSQL) &"-"& month(stdatoSQL) & "-"& year(stdatoSQL)

    stdatoSQL = year(stdatoSQL) &"-"& month(stdatoSQL) & "-" & day(stdatoSQL)
    sldatoSQL = year(sldatoSQL) &"-"& month(sldatoSQL) & "-" & day(sldatoSQL)




'** Altid i den rigtige uge uanset om det er slutdaot eller startato eller om ugen skifter �r s� den. f.eks starter i 2014 og slutter i 2015
varTjDatoUS_tor = dateAdd("d", 3, varTjDatoUS_man)

%>

 <!-- **** Ugeseddel *** -->

<%if showheader = 1 then %>
    <%
        call browsertype() 
        if browstype_client <> "ip" then
    %>
   <table cellpadding=0 cellspacing=0 border=0 width="100%">
   <tr>
	<td valign="top" style="width:90%;"><br />
	
	<%'if cint(SmiWeekOrMonth) <> 2 then 'Dag ==> Logud flow %>
    <%call showafslutuge_ugeseddel %>
    <%'end if %>
    <!-- afslut uge -->
    
   &nbsp;
	</td>

       <td valign="top" align="right"><br />
   
    <%
               'if thisfile = "ugeseddel_2011.asp" then
               fmLink = "ugeseddel_2011.asp?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&varTjDatoUS_son="&varTjDatoUS_son&"&nomenu="&nomenu 
               'else
               'fmLink = "weekpage_2010.asp?medarbid="&medarbid&"&st_dato="&st_dato
               'end if


    oimg = "blank.gif" '"ikon_stat_joblog_48.png"
	oleft = 0
	otop = 0
	owdt = 500
	oskrift = tsa_txt_338 &" "& datepart("ww", varTjDatoUS_tor, 2, 2) &" - "& datepart("yyyy", varTjDatoUS_tor, 2, 2) 
	
	'call sideoverskrift_2014(oleft, otop, owdt, oskrift)

    call meStamdata(usemrn)     
	
    %>
           <h4>
       <%
       thisWeek = datepart("ww", sldatoSQL, 2,2)
       prevWeek = datepart("ww", dateAdd("d", -7, sldatoSQL), 2,2) 
       nextWeek = datepart("ww", dateAdd("d", 7, sldatoSQL), 2,2) 



       if media <> "print" then       
       %>
       
       <a href="ugeseddel_2011.asp?usemrn=<%=usemrn%>&varTjDatoUS_man=<%=prev_varTjDatoUS_man %>&varTjDatoUS_son=<%=prev_varTjDatoUS_son %>&nomenu=<%=nomenu %>&media=<%=media%>"><</a> &nbsp <%=tsa_txt_005 &" "& thisWeek %> &nbsp <a href="ugeseddel_2011.asp?usemrn=<%=usemrn%>&varTjDatoUS_man=<%=next_varTjDatoUS_man %>&varTjDatoUS_son=<%=next_varTjDatoUS_son %>&nomenu=<%=nomenu %>">></a>
        <%else %>
           &nbsp <%=tsa_txt_005 &" "& thisWeek %> &nbsp
        <%end if %>
           </h4>
    </td>

     </tr>
    </table>
    <%end if %>


    <%
        call browsertype() 
        if browstype_client = "ip" then
    %>

    <%thisWeek = datepart("ww", sldatoSQL, 2,2) %>

    <br />
    <div class="row">
        <h4 class="col-lg-12"> <a href="ugeseddel_2011.asp?usemrn=<%=usemrn%>&varTjDatoUS_man=<%=prev_varTjDatoUS_man %>&varTjDatoUS_son=<%=prev_varTjDatoUS_son %>&nomenu=<%=nomenu %>&media=<%=media%>"><</a> &nbsp <%=week_txt_004 %> <%=thisWeek %> &nbsp <a href="ugeseddel_2011.asp?usemrn=<%=usemrn%>&varTjDatoUS_man=<%=next_varTjDatoUS_man %>&varTjDatoUS_son=<%=next_varTjDatoUS_son %>&nomenu=<%=nomenu %>">></a></h4>
    </div>

    <%end if %>




  <%end if 'showheader
      
   'end if   'visLukUge%>

  

<%if media <>  "print" then 
    
   
    %>

 
<%end if%>

 

  <table cellpadding=0 cellspacing=0 border=0 width=100%>
      <%
          nowdate = (DatePart("w", now,2,2))
          dateid = DatePart("w", st_dato,2,2)
      %>

   <tr><td valign=top colspan=2>

         

       <div class="panel-group accordion-panel" id="accordion-paneled">
            <div class="panel panel-default">

                <div class="panel-heading">
                    
                <%
                if media = "print" then
                   %> 
                    <h4 class="panel-title"><a class="accordion-toggle"><b><%=weekdayname(weekday(st_dato, 1)) &" "& formatdatetime(st_dato, 1)%></b> <span style="font-size:95%">- <%=meTxt %></span></a></h4>
                    </div> <!-- /.panel-heading -->
                   
                    <div> <!-- class="panel-collapse collapse" -->
                    <%
                else    

                        %>
                        <h4 class="panel-title"><a class="accordion-toggle" data-toggle="collapse" href="#<%=dateid %>"><b><%=weekdayname(weekday(st_dato, 1)) &" "& formatdatetime(st_dato, 1)%></b> 
                            <span style="font-size:85%; width:60%;">- <%=meTxt %></span>
                            
                             <span style="float:right;"> 

                            <% 
                            showKommega = 1 'sl� den til n�r den er klar til CFLOW    
                            if cint(stempelurOn) = 1 AND session("mid") = 1 AND cint(showKommega) = 1 then %>
                            <span id="sp_sumlontimer_dag_<%=datepart("w", st_dato, 2,2)%>">&nbsp;</span> vs.
                            <%end if %>
                            <span id="sp_sumtimer_dag_<%=datepart("w", st_dato, 2,2)%>">&nbsp;</span></a></h4>
                         
                            </span>

                        </div>
                <!-- /.panel-heading -->
                

                    <%
                    if nowdate = dateid AND (datePart("ww", now, 2,2) = datePart("ww", st_dato, 2,2)) then %>
                    <div id="<%=dateid %>" class="panel-collapse collapse in">
                    <%else %>
                    <div id="<%=dateid %>" class="panel-collapse collapse">
                    <%
                    end if   
                        
                end if %>

                    <div class="panel-body">


                          <%
                         'select case lto 
                         'case "cflow", "intranet - local"
                      
                        if cint(stempelurOn) = 1 AND lto <> "epi2017" AND showKommega = 1 then

                             'if session("mid") = 1 then%>
              
                                <div class="col-lg-5 well">
                                <b><a href="logindhist_2011.asp?usemrn=<%=usemrn%>&varTjDatoUS_man=<%=varTjDatoUS_man %>">Komme/G� tid:</a></b>
                                <%
                             
                                call logindhistorik_week_60_100(usemrn, 2, stdatoSQL, sldatoSQL) %>

                                </div>

                             
                        <% if nowdate = dateid AND (datePart("ww", now, 2,2) = datePart("ww", st_dato, 2,2)) then %>

                                <div class="col-lg-5 well-white">
                                <%


                                    if lto <> "tec" AND lto <> "esn" AND lto <> "cflow" AND lto <> "intranet - local" then %>
                                    <a class="btn btn-default" href="<%=toSubVerPath14 %>stempelur.asp?func=redloginhist&medarbSel=<%=session("mid")%>&showonlyone=1&hidemenu=1&id=0&rdir=sesaba" target="_top" style="color:#FFFFFF;"><%=tsa_txt_435 %></a>
                                    <%else 

                                       %>

                                         
                                        <!--
                                        'if cint(aabentLogudfindes) = 1 then%>

                                               
                                        
                                        <a class="btn btn-info" href="=toSubVerPath14 %>../sesaba.asp?fromweblogud=1&rdir=ugeseddel&varTjDatoUS_man==varTjDatoUS_man %>" target="_top">1) Frys aktuelle komme/g� tid</a><br />
                                                
                                        <if lto = "cflow" OR lto = "intranet - local" then  %>
                                                
                                            
                                                <span style="font-size:11px;">Stopper aktuelle komme/g� tid og beregner l�n-timer. Du kan herefter registrere projekttimer og till�g s� de passer til din komme/g� tid (100 tals).</span><br /><br />
                                                end if 

                                                <span style="background-color:#cccccc; padding:8px;">2) =tsa_txt_435 %></span>

                                        'else

                                             <span style="background-color:#cccccc; padding:8px;">1) Frys aktuelle komme/g� tid</span><br /><br />
                                    
                                                

                                               <a class="btn btn-danger" href="=toSubVerPath14 %>../sesaba.asp?fromweblogud=1&logudDone=1" target="_top" style="color:#FFFFFF;">2) =tsa_txt_435 %></a>

                                                
                                  
                                 
                                        'end if 
                                                    
                                    'end if 
                                     -->
                              

                                      <a class="btn btn-danger" href="<%=toSubVerPath14 %>../sesaba.asp?fromweblogud=1" target="_top" style="color:#FFFFFF;"><%=tsa_txt_435 %></a><%

                                end if%>

                                </div>
           
                            <%end if 'nowdate
                                
                                
                         'end if
                
                        end if   
                        'end select
                         %>



                        <table width="100%" class="table dataTable ui-datatable" style="font-size:90%">
                         

 
                            <thead>
                                <tr >
                                    <!--<th style="width:25%;"><b><%=tsa_txt_237 %>2</b></th> -->

                                    <%
                                        
                                    if browstype_client <> "ip" then
                                    %>
                                        <th style="width:25%;"><b><%=tsa_txt_022 %> (<%=tsa_txt_066 %>)</b></th>
                                    <%else if browstype_client = "ip" then %>
                                        <th style="width:25%;"><b><%=tsa_txt_022 %></b></th>
                                    <%end if
                                    end if %>

                                    <th style="width:30%;"><b><%=tsa_txt_244 %></b></th>

                                    <%
                                    
                                    if browstype_client <> "ip" then
                                    %>
                                        <th style="width:30%"><b><%=tsa_txt_068 %></b> <%=tsa_txt_329 %> & <%=tsa_txt_296 %></th>
                                    <%else if browstype_client = "ip" then  %>
                                        <th style="width:30%"><b><%=tsa_txt_068 %></b></th>
                                    <%end if
                                    end if %>

                                    <!--<th><b>Start - Stop</b></th>-->
                                    <th style="text-align:right"><b><%=tsa_txt_148 %></b></th>

                                   
                             </tr>
                        </thead>
    
                           <%'*** Ugeseddel ***' 



   
                           call afsluger(medarbid, stdatoSQL, sldatoSQL)

                            select case lto
                            case "cflow"
                            'strSQLTFaktimKri = " AND tfaktim = 1 "
                            strSQLTFaktimKri = " AND (tfaktim = 1 OR tfaktim = 2 OR tfaktim = 50 OR tfaktim = 52 OR tfaktim = 60 OR tfaktim = 61 OR tfaktim = 14 OR tfaktim = 22 OR tfaktim = 20 OR tfaktim = 21) "
                            'strSQLTFaktimKri = ""
                            strSQLodrBy = " tdato, sttid "
                            case "nonstop"
                            strSQLodrBy = " tdato, tfaktim, sttid "
                            strSQLTFaktimKri = ""
                            case else
                            strSQLodrBy = " tdato, sttid "
                            strSQLTFaktimKri = ""
                            end select


                           strSQL = "SELECT tid, taktivitetnavn, timer, tfaktim, tjobnavn, tjobnr, tdato, "_
                           &" tknr, tknavn, tmnavn, tmnr, kkundenr, godkendtstatus, godkendtstatusaf, m.mnr, timerkom, init, a.fase, a.easyreg, j.risiko, sttid, sltid, j.id as jobid, tidspunkt FROM timer "_
                           &" LEFT JOIN medarbejdere AS m ON (m.mid = tmnr)"_
                           &" LEFT JOIN aktiviteter AS a ON (a.id = taktivitetid)"_
                           &" LEFT JOIN job AS j ON (j.jobnr = tjobnr)"_
                           &" LEFT JOIN kunder K on (kid = tknr) WHERE tmnr = "& medarbid &" "& strSQLTFaktimKri &" AND "_
                           &" tdato BETWEEN '"& stdatoSQL &"' AND '"& sldatoSQL &"' ORDER BY "& strSQLodrBy &" DESC "
   

                        'if session("mid") = 1 then
                        '   Response.Write strSQL
                        '   Response.flush
                        'end if


                           at = 0
                           if showheader = 1 then
                           timertot = 0
                           end if
                           lwedaynm = ""
                           timerDag = 0
                           antalEasyReg = 0
                           eaDag = 0
                           atDag = 0
                           'ugegodkendtTxt = ""
   
                           oRec.open strSQL, oConn, 3
                           while not oRec.EOF 


                                        '** Er periode/timeregistrering godkendt ***'
		                                tjkDag = oRec("Tdato")
		                                erugeafsluttet = instr(afslUgerMedab(oRec("tmnr")), "#"&datepart("ww", tjkDag,2,2)&"_"& datepart("yyyy", tjkDag) &"#")

                                        strMrd_sm = datepart("m", tjkDag, 2, 2)
                                        strAar_sm = datepart("yyyy", tjkDag, 2, 2)
                                        strWeek = datepart("ww", tjkDag, 2, 2)
                                        strAar = datepart("yyyy", tjkDag, 2, 2)

                                        if cint(SmiWeekOrMonth) = 0 then
                                        usePeriod = strWeek
                                        useYear = strAar
                                        else
                                        usePeriod = strMrd_sm
                                        useYear = strAar_sm
                                        end if

                
                                        call erugeAfslutte(useYear, usePeriod, usemrn, SmiWeekOrMonth, 0)
                
                                        call lonKorsel_lukketPer(tjkDag, oRec("risiko"))
                
		                                'if ( (( datepart("ww", ugeNrAfsluttet, 2, 2) = usePeriod AND cint(SmiWeekOrMonth) = 0) OR (datepart("m", ugeNrAfsluttet, 2, 2) = usePeriod AND cint(SmiWeekOrMonth) = 1 )) AND cint(ugegodkendt) = 1 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                                        '(smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", tjkDag) = year(now) AND DatePart("m", tjkDag) < month(now)) OR _
                                        '(smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", tjkDag) < year(now) AND DatePart("m", tjkDag) = 12)) OR _
                                        '(smilaktiv = 1 AND autolukvdato = 1 AND DatePart("yyyy", tjkDag) < year(now) AND DatePart("m", tjkDag) <> 12) OR _
                                        '(smilaktiv = 1 AND autolukvdato = 1 AND (year(now) - DatePart("yyyy", tjkDag) > 1))) OR cint(lonKorsel_lukketIO) = 1 then
              
                                        'ugeerAfsl_og_autogk_smil = 1
                                        'else
                                        'ugeerAfsl_og_autogk_smil = 0
                                        'end if 

                                        '*** tjekker om uge er afsluttet / lukket / l�nk�rsel
                                        call tjkClosedPeriodCriteria(tjkDag, ugeNrAfsluttet, usePeriod, SmiWeekOrMonth, splithr, smilaktiv, autogk, autolukvdato, lonKorsel_lukketIO)

                
                                        if oRec("godkendtstatus") = 1 OR oRec("godkendtstatus") = 3 then 'Godkendt ell. Tentative
                                        ugeerAfsl_og_autogk_smil = 1
                                        else
                                        ugeerAfsl_og_autogk_smil = ugeerAfsl_og_autogk_smil
                                        end if







    
                            select case right(at, 1)
                            case 0,2,4,6,8
                            bgcol = "#ffffff"
                            case else
                            bgcol = "#EFF3ff"
                            end select
    
                            call akttyper2009prop(oRec("tfaktim"))
    
    
                            if lwedaynm <> weekdayname(weekday(oRec("Tdato"))) then
                             if at <> 0 then%>
                            <tr>
                                <td colspan=6 align=right><b><%=formatnumber(timerDag, 2) %></b>
                                    <!--<br />
                                <span style="color:#999999;">Fordelt p� <%=atDag %> registreringer, heraf <%=eaDag %> Easyreg.</span><br /><br />&nbsp;--></td>
                                <td>&nbsp;</td>

                            </tr>
                            <%
                            atDag = 0
                            eaDag = 0
                            timerDag = 0
    
                            end if 
                            
                             if visning <> 1 then%>
                            <tr bgcolor="#D6Dff5">
                                <td colspan=7 valign=bottom style="height:30px; padding:3px;"><b><%=weekdayname(weekday(oRec("Tdato"), 1)) %></b></td>
                            </tr>
                            <%end if


                            end if
    
                            %>
                           <tr bgcolor="<%=bgcol%>">
                          <!-- <td valign=top style="border-bottom:1px #cccccc solid; padding-right:20px; height:20px; white-space:nowrap;">
                               <%
                                tmp = 1
                               if tmp = 900 then
                               'if (((ugeerAfsl_og_autogk_smil = 0 AND cint(ugegodkendt) <> 1) OR (level = 1)) AND media <> "print") then %>
                               <input type="text" value="<%=oRec("tdato")%>" style="width:80px;" />
                                <%else %>
                               <%=oRec("tdato")%>
                               <%end if %>

                           </td> -->
                         <!--  <td valign=top style="border-bottom:1px #cccccc solid;"><%=left(oRec("tmnavn"), 10) %> 
                           <%if len(trim(oRec("init"))) <> 0 then %>
                           [<%=oRec("init") %>]
                           <%end if %></td> -->
                           
                           <% if browstype_client <> "ip" then %>
                               <td valign=top style="width:25%"><%=oRec("tknavn") %> (<%=oRec("kkundenr") %>)</td>

                               <%if lto = "mpt" OR (level = 1 OR level = 2 OR level = 6) then %>
                               <td valign=top style="width:30%"><a href="../timereg/jobs.asp?menu=job&func=red&id=<%=oRec("jobid")%>" target="_blank" ><%=oRec("tjobnavn") %> (<%=oRec("tjobnr") %>)</a></td>
                               <%else %>
                               <td valign=top style="width:30%"><%=oRec("tjobnavn") %> (<%=oRec("tjobnr") %>)</td>
                               <%end if %>

                               <td style="width:30%"><%=oRec("taktivitetnavn") %>
                                   <%if oRec("easyreg") <> 0 then %>
                                    <span style="color:#999999;">(E!)</span>
                                   <%antalEasyReg = antalEasyReg + 1  
                                   eaDag = eaDag + 1%>
                                   <%end if %>
   
    
                                    <%
	                                call akttyper(oRec("tfaktim"), 1)
	                                %>

                                    <span style="color:yellowgreen;">(<%=lcase(akttypenavn)%>)</span>
	
   
                                   <%if len(trim(oRec("fase"))) <> 0 then %>
                                   <br />
                                   <span style="color:#5582d2;"><%=week_txt_005 %>: <%=replace(oRec("fase"), "_", " ")%></span>
                                   <%end if %>

                                    <%if len(trim(oRec("timerkom"))) <> 0 then %>
                                    <span style="color:#999999;">
                                    <%if len(oRec("timerkom")) > 100 then %>
                                    <i><%=left(oRec("timerkom"), 100) %>...</i>
                                    <%else %>
                                    <i><%=oRec("timerkom")%></i>
                                    <%end if %>
                                    </span>
                                    <%end if %>
                                </td>

                           <%else if browstype_client = "ip" then%>
                               <td valign=top style="width:25%"><%=left(oRec("tknavn"), 9) %></td>
                               
                               <td valign=top style="width:30%"><%=left(oRec("tjobnavn"), 15) %>
                                <% if len(oRec("tjobnr")) > 4 then %>
                                (<%=left(oRec("tjobnr"), 1) %>...)
                                <% else  %>
                                (<%=left(oRec("tjobnr"), 4) %>)
                                <%end if %> 
                               </td>
                                                                   

                               <td style="width:30%"><%=left(oRec("taktivitetnavn"), 15) %>
                                   <%if oRec("easyreg") <> 0 then %>
                                    <span style="color:#999999;">(E!)</span>
                                   <%antalEasyReg = antalEasyReg + 1  
                                   eaDag = eaDag + 1%>
                                   <%end if %>

                                    <%if len(trim(oRec("timerkom"))) <> 0 then %>
                                    <span style="color:#999999;">
                                    <%if len(oRec("timerkom")) > 15 then %>
                                    <i><%=left(oRec("timerkom"), 12) %>...</i>
                                    <%else %>
                                    <i><%=left(oRec("timerkom"), 15)%></i>
                                    <%end if %>
                                    </span>
                                   <%end if %>

                               </td>
                            <%end if

                              end if
                                
                                
                                
                                '*** tilf�j start stop p� hver linje%>
                                <!--
                                <td>
                                if lto = "cflow" then
                                <span style="font-size:12px;">
                                =left(formatdatetime(oRec("tidspunkt"), 3), 5) %> </span> - <a class="btn btn-xs btn-danger" href="#">Stop</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                end if%>
                                </td>
                                end if %> 
                                -->
                               

                               <td style="text-align:right"><%
                
                
                                if (((ugeerAfsl_og_autogk_smil = 0 AND cint(ugegodkendt) <> 1) OR (level = 1)) AND media <> "print") then
                                      
                                      select case cint(oRec("godkendtstatus"))
                                         case 2
                                         gkbgcol = "red"
                                            
                                         case 1
                                         gkbgcol = "yellowgreen"
                                         
                                         case 3
                                       
                                         gkbgcol = "orange"
                                         

                                         case else
                                         
                                         gkbgcol = "#5582d2"
                                         end select
                                      
                                      %>

                                    <%if lto <> "tbg" OR (lto = "tbg" AND level = 1) then %>
                                        <a href="#" onclick="Javascript:window.open('rediger_tastede_dage_2006.asp?id=<%=oRec("tid") %>', '', 'width=450,height=675,resizable=yes,scrollbars=yes')" class=vmenu style="color:<%=gkbgcol%>;"><%=formatnumber(oRec("timer"), 2) %></a>
                                    <%else 
                                      
                                      if gkbgcol = "#5582d2" then
                                      gkbgcol = "#000000"
                                      end if
                                      %>
                                      <span style="color:<%=gkbgcol%>;"><%=formatnumber(oRec("timer"), 2) %></span>
                                    <%end if %>
	            
                                <%else 
                                  
                                   select case cint(oRec("godkendtstatus"))
                                         case 2
                                         gkbgcol = "red"
                                            
                                         case 1
                                         gkbgcol = "yellowgreen"
                                         
                                         case 3
                                       
                                         gkbgcol = "orange"
                                         

                                         case else
                                         
                                         gkbgcol = "#000000"
                                         end select
                                  
                                  %>
                                  <span style="color:<%=gkbgcol%>;"><%=formatnumber(oRec("timer"), 2) %></span>
                                  
                                <%end if %>

                
                                <%if formatdatetime(oRec("sttid"), 3) <> "00:00:00" then %>

                                <br /><span style="color:#999999; font-size:9px;"><%=left(formatdatetime(oRec("sttid"), 3), 5) %> - <%=left(formatdatetime(oRec("sltid"), 3), 5) %></span>

                                <%end if %>
            

                            </td>
                                        <%
                                            'Skjuler Status klik
                                            vis = 0

                                        if vis = 1 then
                                        %>
                                        <td align="right" valign="top" style="border-bottom:1px #cccccc solid;">
                                         
                                            <label>

                                               <%

                           
                    
                       
                                             '  else 
                       
                                                     'response.write "teamLederEditOK: "& teamLederEditOK & "ugeerAfsl_og_autogk_smil: " & ugeerAfsl_og_autogk_smil & " ugegodkendt: "& ugegodkendt &" level "& level 


                                        '***  Admin kan godkende for alle medarb. Teamleder kun for dem man er teamleder for. 
                                         '*** visAlleMedarb = 1 skal v�re sl�et til f�r man kan godkende andres (undt. admin), da det ikke giver mening at goende sine egne timer
                                         select case cint(oRec("godkendtstatus"))
                                         case 2
                                         gkTxt = "<b>!</b>"
                                         gkbgcol = "red"
                                             if ((ugeerAfsl_og_autogk_smil = 0 AND cint(ugegodkendt) <> 1) OR cint(level) = 1) AND media <> "print" then
                                             gkInputype = "CHECKBOX"
                                             gkInputName = "ids"
                                             else
                                             gkInputype = "hidden"
                                             gkInputName = "ids_hd"
                                             end if
                                         case 1
                                         gkTxt = "<b><i>V</i></b>"
                                         gkbgcol = "green"
                                          if ((ugeerAfsl_og_autogk_smil = 0 AND cint(ugegodkendt) <> 1) OR cint(level) = 1) AND media <> "print" then
                                             gkInputype = "CHECKBOX"
                                                   gkInputName = "ids"
                                             else
                                             gkInputype = "hidden"
                                                   gkInputName = "ids_hd"
                                             end if
                                        
                                         case 3
                                         gkTxt = "<b><i>(V)</i></b>"
                                         gkbgcol = "orange"
                                          if ((ugeerAfsl_og_autogk_smil = 0 AND cint(ugegodkendt) <> 1) OR cint(level) = 1) AND media <> "print" then
                                             gkInputype = "CHECKBOX"
                                                   gkInputName = "ids"
                                             else
                                             gkInputype = "hidden"
                                                   gkInputName = "ids_hd"
                                             end if

                                         case else
                                         gkTxt = ""
                                             if ((ugeerAfsl_og_autogk_smil = 0 AND cint(ugegodkendt) <> 1) OR cint(level) = 1) AND media <> "print" then
                                             gkInputype = "CHECKBOX"
                                                   gkInputName = "ids"
                                             else
                                             gkInputype = "hidden"
                                                   gkInputName = "ids_hd"
                                             end if
                                         gkbgcol = "#000000"
                                         end select

                
                 
                                         Response.write "<span style=""color:"& gkbgcol &";"">"& gkTxt &"</span>" 

                                         select case cint(oRec("godkendtstatus"))
                                         case 1, 2, 3
                                         %>
                                         <br><span style='font-size:9px; color:#999999;'><%=left(oRec("godkendtstatusaf"), 12)  %></span><br />
                                         <%
                                          end select

                       
                                               %>

                                            <input class="status_chk_<%=st_dato%>" name="<%=gkInputName %>" type="<%=gkInputype %>" value="<%=oRec("tid")%>" /> 
                                                </label>



                                        </td> <%end if %>
                           </tr>
   
                           <%
                           lwedaynm = weekdayname(weekday(oRec("Tdato")))
                           lwedate = oRec("tdato")

                           

                           if cint(aty_real) = 1 OR (lto = "tec") OR (lto = "esn") then
                           timertot = timertot + oRec("timer")
                           timerDag = timerDag + oRec("timer")
                           end if
                           at = at + 1

                           atDag = atDag + 1
                           oRec.movenext
                           wend
                           oRec.close


       


   

   
                           if at <> 0 then

                                   %>

                            <tr>
                                <td colspan="6" style="text-align:right;"><b><%=formatnumber(timerDag, 2) %></b></td>
                              
                              <!--  <td colspan=5>&nbsp;</td>
                                <td align=right><b><%=formatnumber(timerDag, 2) %> t.1</b>
                               <!-- <br />
                                <span style="color:#999999;">Fordelt p� <%=atDag %> registreringer, heraf <%=eaDag %> Easyreg.</span>-->
                            </tr>
                         
   
                          <% end if
           
  
                          


           
    



       if cint(visning) = 1 then


    '**** Timer uden match, fra import Eller timeTag **** sttid, sltid
     strSQL = "SELECT timer, jobid, tip.jobnavn, tip.aktnavn AS tipaktnavn, aktid, tdato, mnavn, mnr, init, timerkom, a.navn AS aktnavn, tip.id AS extsysId, errmsg "_
     &" FROM timer_import_temp AS tip "_
     &" LEFT JOIN medarbejdere AS m ON (m.mid = tip.medarbejderid)"_
     &" LEFT JOIN job AS j ON (j.id = tip.jobid)"_
     &" LEFT JOIN aktiviteter AS a ON (a.id = tip.aktid)"_
     &" WHERE tdato = '"& stdatoSQL & "' AND medarbejderid = "& medarbid & " AND overfort = 0"
     

    'if (lto = "assurator" AND session("mid") = 14) then
    'response.write " strSQL: " & strSQL
    ' response.flush
    'end if
           
    if len(trim(ucount)) <> 0 then
       u = ucount
     else
       u = 1
    end if

        
           
        oRec.open strSQL, oConn, 3
        while not oRec.EOF 

       'select case right(u, 1)
       'case 0,2,4,6,8
       'timerTempBgCol = "#FFE6E6" '"#FFDFFF"
       'case else
       'timerTempBgCol = "#FFFFFF"
       'end select


      
        
     

           if media <> "print" then%>
               <form action="../timereg/timereg_akt_2006.asp?func=db&rdir=ugeseddel_2011&usemrn=<%=usemrn%>&FM_medid=<%=usemrn%>&varTjDatoUS_man=<%=varTjDatoUS_man%>" method="post">
                  
               <input type="hidden" value="<%=positiv_aktivering_akt_val%>" name="FM_pa" />
               <input type="hidden" id="Hidden4" name="FM_dager" value=", xx"/>
               <input type="hidden" id="Hidden2" name="FM_feltnr" value="0"/>
               <input type="hidden" id="Hidden5" name="year" value="<%=year(oRec("tdato")) %>"/>
              
            <%end if%>

                <%if lastDato <> oRec("tdato") then %>
               <tr>

                   <td colspan="7"><span style="color:red;"><%=week_txt_006 %>:</span></td>

               </tr>
                    <%end if %>
      
       
     
       <tr class="danger">
           <td style="text-align:left;" valign="top">

              

                <input type="hidden" id="extsysId" name="extsysId" value="<%=oRec("extsysId") %>"/>
               <input type="text" name="FM_datoer" value="<%=oRec("tdato")%>" style="width:80px;" class="form-control input-small" />

               <span style="color:#999999; font-size:9px;"><b>Err. msg: </b><%=oRec("errmsg") %></span><br />

              <%if media <> "print" then %>  
             <a href="../to_2015/ugeseddel_2011.asp?func=slet_tip&usemrn=<%=usemrn %>&varTjDatoUS_man=<%=varTjDatoUS_man%>&varTjDatoUS_son=<%=varTjDatoUS_son %>&nomenu=<%=nomenu%>&id=<%=oRec("extsysId") %>"><span style="color:darkred;" class="fa fa-times"></span></a>
                
               <%end if %>

           </td>
          
           
            <!--<td valign="top" style="border-bottom:1px #cccccc solid;"><input type="hidden" value="<%=usemrn %>" name="FM_medid" id="FM_medid" /><%=left(oRec("mnavn"), 10) %>
                <%if len(trim(oRec("init"))) <> 0 then %>
   [<%=oRec("init") %>]
   <%end if %>

           </td>
           -->

           <%if len(trim(oRec("jobid"))) <> 0 AND oRec("jobid") <> 0 then
               ugsJobid = oRec("jobid") 
               ugsJobidFundet = 0

                strSQLjobf = "SELECT id FROM job WHERE id = "& ugsJobid
                oRec6.open strSQLjobf, oConn, 3
                if not oRec6.EOF then 

                   ugsJobidFundet = 1

                end if
                oRec6.close

             else
               ugsJobid = 0
               end if %>

           <td valign="top" style="white-space:nowrap;">
               <input type="hidden" name="FM_jobid" id="FM_jobid_<%=u %>" value="<%=ugsJobid %>"/>
               <input type="hidden" name="FM_job_opr" id="FM_job_opr_<%=u %>" value="<%=oRec("jobnavn") %>"/>

               <%if cdbl(ugsJobid) <> 0 AND cint(ugsJobidFundet) = 1 then
                strJobFundet = "-- OK"
                else
                strJobFundet = ""
                end if %>


               <input type="text" name="FM_jobnavn" id="FM_job_<%=u %>" value="<%=oRec("jobnavn") & " " & strJobFundet %>" class="FM_job form-control input-small" />
               
             
               <!--<div id="dv_job_<%=u %>" class="dv_job" style="position:absolute; background-color:#FFFFFF; visibility:hidden; display:none; width:250px; height:200px; overflow:auto; padding:10px; border:1px #999999 solid;"></div>-->
               
                <select id="dv_job_<%=u %>" class="form-control input-small chbox_job" size="10" style="visibility:hidden; display:none;">
                                 <option><%=week_txt_007 %>..</option>
                             </select>


           </td>
           <td valign="top" style="white-space:nowrap;">
           <!--<input type="hidden" name="FM_aktivitetid" id="FM_aktid_<%=u %>" value="0"/>-->
    
            <%if len(trim(oRec("aktnavn"))) <> 0 AND oRec("aktnavn") <> "" then
               ugeAktnavn = oRec("aktnavn") 
             else
               ugeAktnavn = oRec("tipaktnavn")
             end if %>


                <%if len(trim(oRec("aktid"))) <> 0 AND oRec("aktid") <> 0 then
               ugsAktid = oRec("aktid") 


                     ugsAktidFundet = 0

                strSQLaf = "SELECT id FROM aktiviteter WHERE id = "& ugsAktid
                oRec6.open strSQLaf, oConn, 3
                if not oRec6.EOF then 

                   ugsAktidFundet = 1

                end if
                oRec6.close

               else
               ugsAktid = 0
               end if %>
            
               
           <input type="hidden" name="FM_aktivitetid" id="FM_aktid_<%=u %>" value="<%=ugsAktid %>"/>

               <%if cdbl(ugsAktid) <> 0 AND cint(ugsAktidFundet) = 1 then
                strAktFundet = "-- OK"
                else
                strAktFundet = ""
                end if %>

           <input type="text" name="FM_aktnavn" value="<%=ugeAktnavn & " " & strAktFundet %>" class="FM_akt form-control input-small" id="FM_akt_<%=u %>" />

              
              
               <!--<div id="dv_akt_<%=u %>" class="dv_akt" style="position:absolute; background-color:#FFFFFF; visibility:hidden; display:none; width:200px; height:200px; overflow:auto; padding:10px; border:1px #999999 solid;"></div>-->
               
                 <select id="dv_akt_<%=u %>" class="form-control input-small chbox_akt" size="10" style="visibility:hidden; display:none;">
                                      <option>S�ger..</option>
                             </select>

               <textarea id="FM_kom_<%=u %>" name="FM_kom_0" style="height:40px;" placeholder="Kommentar" class="form-control input-small"><%=trim(oRec("timerkom")) %></textarea>
               

            
           </td>
           <td valign="top" align="right"><input type="text" name="FM_timer" value="<%=formatnumber(oRec("timer"), 2) %>" style="width:60px;" class="form-control input-small" />
               <input type="hidden" name="FM_timer" value="xx"/>
               <button class="btn btn-sm btn-success"><b><%=week_txt_008 %> >></b></button>
             
            
           </td>

       </tr>
           
       <%if media <> "print" then %>   
       
       </form>
   
       <% end if
        lastDato = oRec("tdato") 
        timerUge = timerUge + timerDag 
        u = u + 1
        ucount = u

   oRec.movenext
   wend
   oRec.close


   end if

    %>



            </table>
                    </div>
                </div>

            </div>
       </div>
            <input type="hidden" id="sumtimer_dag_<%=dateid%>" value="<%=formatnumber(timerDag, 2)%>" /> 

           <% if cint(stempelurOn) = 1 then %>
           <input type="hidden" id="sumlontimer_dag_<%=dateid%>" value="<%=timerThisDayLontimerTotTxt%>" /> 
           <%end if 
               
           timerThisDayLontimerTotTxt = 0%>
    
           <!--
   
   </tr>
    </table>
    
   -->
    

    <%end function

%>