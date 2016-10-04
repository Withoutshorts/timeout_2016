




<%
'** Finder dag for sidste rettidgie afslutning. Baseret p� kontrolpanel indstilllinger i smileyAfslutSettings()
public cDateUge, cDateAfs, cDateUgeTilAfslutning, smileysttxt, smileyImg, weekafslTxt
        function rettidigafsl(ugeVal)
            

               call smileyAfslutSettings()


                
                if cint(SmiWeekOrMonth) = 0 then '0: ugebasis, 1: M�ned
                weekafslTxt = "i den f�lgende uge"
                else
                weekafslTxt = "i den f�lgende m�ned"
                end if


                cDateUgeTilAfslutning = year(ugeVal) &"/"& month(ugeVal) &"/"& day(ugeVal)

                if SmiantaldageCount < 8 then 'antal dage der skal l�gges til (mandag: 1 s�ndag: 7)
                cDateUge = dateadd("d", SmiantaldageCount, year(ugeVal) &"/"& month(ugeVal) &"/"& day(ugeVal) &" "& SmiantaldageCountClock)
		        cDateAfs = year(now) &"/"& month(now) &"/"& day(now) &" "& time
                else '1 hverdag
                    
                    '1 i ns�te m�ned
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

                'response.write "HER: " & cDateUge
                
		        
		        'select case lcase(lto)
                'case "intranet - local"
                'cDateUge = dateadd("d", -2, year(ugeVal) &"/"& month(ugeVal) &"/"& day(ugeVal) &" 14:00:00")
		        'cDateAfs = year(now) &"/"& month(now) &"/"& day(now) &" "& time
                'weekafslTxt = "i samme uge"
                'case "epi", "epi_no", "epi_sta", "kejd_pb", "kejd_pb2", "epi_ab"
		        'cDateUge = dateadd("d", 1, year(ugeVal) &"/"& month(ugeVal) &"/"& day(ugeVal) &" 12:00:00")
		        'cDateAfs = year(now) &"/"& month(now) &"/"& day(now) &" "& time
                'weekafslTxt = "i den f�lgende uge"
                'case "mi"
		        'cDateUge = dateadd("d", 1, year(ugeVal) &"/"& month(ugeVal) &"/"& day(ugeVal) &" 16:30:00")
		        'cDateAfs = year(now) &"/"& month(now) &"/"& day(now) &" "& time
                'weekafslTxt = "i den f�lgende uge"
		        'case "lw", "synergi1"
		        'cDateUge = dateadd("d", 1, year(ugeVal) &"/"& month(ugeVal) &"/"& day(ugeVal) &" 12:00:00")
		        'cDateAfs = year(now) &"/"& month(now) &"/"& day(now) &" "& time
		        ' weekafslTxt = "i den f�lgende uge"
                'case "dencker" 
		        'cDateUge = dateadd("d", 1, year(ugeVal) &"/"& month(ugeVal) &"/"& day(ugeVal) &" 9:00:00")
		        'cDateAfs = year(now) &"/"& month(now) &"/"& day(now) &" "& time
                'weekafslTxt = "i den f�lgende uge"
                'case "ngf"
                'cDateUge = dateadd("d", 2, year(ugeVal) &"/"& month(ugeVal) &"/"& day(ugeVal) &" 12:00:00")
		        'cDateAfs = year(now) &"/"& month(now) &"/"& day(now) &" "& time
                'weekafslTxt = "i den f�lgende uge"
		        'case else
		        'cDateUge = dateadd("d", 1, year(ugeVal) &"/"& month(ugeVal) &"/"& day(ugeVal) &" 9:00:00")
		        'cDateAfs = year(now) &"/"& month(now) &"/"& day(now) &" "& time
		         'weekafslTxt = "i den f�lgende uge"
                'end select


    end function


public afslutugeBasisKri, afslProc, afslProcKri, weekSelectedTjk 
function afslutkri(varTjDatoUS_son, tjkTimeriUgeDt, usemrn, lto, SmiWeekOrMonth)

        select case cint(SmiWeekOrMonth) 
        case 0 
        weekSelectedTjk = dateAdd("d", 7, varTjDatoUS_son)
        case 1
        varTjDatoUS_tjk = varTjDatoUS_son 'dateAdd("d", -6, varTjDatoUS_son) 'finder mandag i den valgte uge. St. dato for ugne bestemmer hvilken m�ned man er ved at afslutte.
        weekSelectedTjk = "1-"& month(varTjDatoUS_tjk) &"-"& year(varTjDatoUS_tjk)
        case 2
        varTjDatoUS_tjk = formatdatetime(tjkTimeriUgeDt, 2)
        weekSelectedTjk = formatdatetime(tjkTimeriUgeDt, 2) '"1-"& month(varTjDatoUS_tjk) &"-"& year(varTjDatoUS_tjk)
        end select

            select case lto 
            case "dencker", "intranet - local"

                freDagiSidsteUge = tjkTimeriUgeDt 'tjekdag(1) 'dateAdd("d", -3, tjekdag(1))
                freDagiSidsteUge = year(freDagiSidsteUge) &"/"& month(freDagiSidsteUge) &"/"& day(freDagiSidsteUge) 
                call fLonTimerPer(freDagiSidsteUge, 6, 21, usemrn)

               'response.write "totaltimerPer: " & totaltimerPer
               'response.write "freDagiSidsteUge: "& freDagiSidsteUge & "<br>"
               'response.write "<br>totalTimerPer100: "& totalTimerPer100 & "("& totalTimerPer100/60& ")"

            afslutugeBasisKri = formatnumber(totalTimerPer100/60, 0) 
            case else


            call normtimerPer(usemrn, varTjDatoUS_son, 6, 0) 
            '** gns timer pr. iforhold til samlet arbejdsuge, baseret p� en 5 dages arbejdsuge. ***'
            '** S�ledes at det er underordnet om man indtaster p� en tirsdag (7,5) eller en fredag (7,0) '****'
            normTimerprUge = (ntimManIgnHellig + ntimTirIgnHellig + ntimOnsIgnHellig + ntimTorIgnHellig + ntimFreIgnHellig + ntimLorIgnHellig + ntimSonIgnHellig)            
            afslutugeBasisKri = normTimerprUge
            end select



            if totTimerWeek > 0 AND afslutugeBasisKri > 0 then
            afslProc = formatnumber(((totTimerWeek/1)/(afslutugeBasisKri/1)) * 100, 0)
            else
            afslProc = 0
            end if


            if cdbl(afslProc) >= cdbl(afslutugekri_proc) then 'Er Reltimer/NormtidsKriterie opflydt
            afslProcKri = 1
            else
            afslProcKri = 0
            end if



end function


public akttypeKrism, afslutugekri, afslutugekri_proc, tjkTimeriUgeDt, tjkTimeriUgeDtDay, tjkTimeriUgeSQL, tjkTimeriUgeDtTxt, afslugeDatoTimerudenMatch
function timeKriOpfyldt(lto, sidsteUgenrAfsl, meType, usemrn, SmiWeekOrMonth)

    'response.write "afslutugekri: " & afslutugekri & "#<br>"
    select case cint(SmiWeekOrMonth)
    case 0,1
    tjkTimeriUgeDt = dateAdd("d", 7, sidsteUgenrAfsl)
    case 2
    '** Hvis dd er st�rre en sidste afsluttet vises n�ste dag til afslutning. (med Submit)
    '** Er den = med dd eller mindre vise "Epriode er afsluttet.."
        if cDate(sidsteUgenrAfsl) < cDate(now) then
        tjkTimeriUgeDt = dateAdd("d", 1, sidsteUgenrAfsl)
        else
        tjkTimeriUgeDt = sidsteUgenrAfsl '** Dagen efter sidste afsluttede dag 'sidsteUgenrAfsl
        end if
    end select

    tjkTimeriUgeDtDay = datepart("w", tjkTimeriUgeDt, 2,2)
    

    select case cint(SmiWeekOrMonth)
    case 0,1

        if tjkTimeriUgeDtDay <> 1 then 
        tjkTimeriUgeDt = dateAdd("d", -(tjkTimeriUgeDtDay-1), tjkTimeriUgeDt)
        end if

    case 2
        tjkTimeriUgeDt = tjkTimeriUgeDt
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
        
              
              
            '** Er der kriterier for antal timer derskal v�re opn�et i% f�r uge m� afsluttes? *'
            afslutugekri = 0
            afslutugekri_proc = 0
            strSQlmt = "SELECT afslutugekri, afslutugekri_proc FROM medarbejdertyper WHERE id = "& meType
            oRec6.open strSQlmt, oConn, 3
            if not oRec6.EOF then
            
               afslutugekri  = oRec6("afslutugekri")
               afslutugekri_proc  = oRec6("afslutugekri_proc")
                  
            end if 
            oRec6.close
                    
        
         if cint(afslutugekri) = 2 then
         akttypeKrism = replace(aty_sql_fakbar, "fakturerbar","tfaktim")
         else
         akttypeKrism = aty_sql_realhours
         end if


         '** Findes der timereg. uden match UANSET DATO, m� periode ikke afsluttes ***
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
     'response.write "SmiWeekOrMonth XXX: "& SmiWeekOrMonth & " varTjDatoUS_son: "& varTjDatoUS_son & " ugeNrStatus: "& ugeNrStatus

              if cint(showAfsuge) = 0 then 
    
                        select case cint(SmiWeekOrMonth)
                        case 0 
                        periodeTxt = tsa_txt_005 & " "& datepart("ww", tjkDato, 2 ,2)
                        case 1
                        periodeTxt = monthname(datepart("m", tjkDato, 2 ,2))
                        case 2
                        periodeTxt = tjkDato
                        end select%>
              


                      <span style="color:#999999; background-color:#F7f7f7; padding:5px;"><span style="color:yellowgreen;"><i>V</i></span> - <%=periodeTxt%> er <b>afsluttet</b> af medarbejderen</span>&nbsp;
            
            
                        <%
                        select case lto
                        case "tec", "esn"
                        lukTxt = "lukket"
                        case else
                        lukTxt = "godkendt"
                        end select    
                            
                            
                            
                        select case ugegodkendt
                        case 1
                        call meStamdata(ugegodkendtaf)
                        ugegodkendtStatusTxt = " Periode er <b>"& lukTxt &"</b> af <a href='mailto:"&meEmail&"'>"& left(meNavn, 10) & " ["& meInit &"]</a>"
                        ugstCol = "yellowgreen" '"#DCF5BD"
                        ugstFtc = "green"
                        ugstBd = "#6CAE1C"
                        case 2
                        call meStamdata(ugegodkendtaf)
                        ugegodkendtStatusTxt = " Periode er <b>afvist</b> af <a href='mailto:"&meEmail&"'>"& left(meNavn, 10) & " ["& meInit &"]</a>"
                       
                        ugstCol = "#FF6666"
                        ugstFtc = "#000000"
                        ugstBd = "#CCCCCC"
                        case else
                        ugegodkendtStatusTxt = " Perioden er endnu ikke "& lukTxt &"/afvist"
                        ugstCol = ""
                        ugstFtc = "#999999"
                        ugstBd = "#cccccc"
                        end select %>


        
                         <span style="color:<%=ugstFtc%>; background-color:<%=ugstCol%>; padding:5px;">
                         <%=ugegodkendtStatusTxt %>
                        </span>

                            <%if cint(ugegodkendt) = 2 then 'afvist 
                              response.write "<br><br><div style=""color:#999999; padding:0px;"">N�r en periode er afvist, skal Du blot rette afviste registreringer, og sende mail (klik p� navnet ovenfor) til din godkender om at dine registreringer er rettet. <br>Du skal IKKE godkende perioden p� ny.<br><br>"_
                              &"<b>Kommentar fra godkender:</b><br>"& ugegodkendtTxt &"</div>"
                              end if %>



            <%end if 


end function



'** Godkend ugeseddel funktionen, med felter og submit funktion
function godkendugeseddel(fmlink, usemrn, varTjDatoUS_man, rdir) 

               
              varTjDatoUS_son = dateadd("d", 6, varTjDatoUS_man)
                

              'response.write varTjDatoUS_son & "<br>" & fmlink

              call smileyAfslutSettings()

              if cint(SmiWeekOrMonth) = 0 then 'uge
                periodeTxt = "Uge"
                periodeNavn = datepart("ww", varTjDatoUS_son, 2,2)
              else
                periodeTxt = "M�ned"
                periodeNavn = left(monthname(datepart("m", varTjDatoUS_son, 2,2)), 3) & "."
              end if

             select case lto
                case "tec", "esn"
                lukTxt = "lukkes"
                lukTxt1 = "lukket"
                lukTxt2 = "Luk"
                teamlederTxt = "For godkender (leder)"
                case else
                lukTxt = "godkendes"
                lukTxt1 = "godkendt"
                lukTxt2 = "Godkend"
                teamlederTxt = "For godkender (Teamleder)"
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
                        <div style="background-color:#FF6666; padding:5px;"><b><%=periodeTxt %> er afvist!</b><br />
                        <span style="font-size:9px; line-height:12px; color:#ffffff;"><i><%=ugegodkendtdt %> af <%=meNavn %></i></span>
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
                           N�r en <%=periodeTxt &" "& lukTxt %>, <%=lukTxt %> alle periodens registreringer automatisk.<br /><br />
                
                           <input id="Submit2" type="submit" value="<%=lukTxt2 &" "& periodeTxt %> >>" style="font-size:9px; width:120px;" />
             
                           </td></tr></table>
                           </form>

                <%else 
                        
                        call meStamdata(ugegodkendtaf)%>
                        <div style="color:green; font-size:11px; background-color:yellowgreen; padding:5px;"><b>Denne <%=periodeTxt %> er <%=lukTxt1 %>!</b><br />
                        <span style="font-size:9px; line-height:12px; color:#ffffff;"><i><%=ugegodkendtdt %> af <%=meNavn %></i></span></div>
                
                <%end if %>

               <%else %>
                
                 <form action="<%=fmLink%>&func=adviserugeafslutning" method="post">
                <table width=90% cellpadding=0 cellspacing=0 border=0>
                <tr><td class=lille>

                  <span style="font-size:11px;"><b><%= lukTxt2 &" "&periodeTxt %></b></span><br />
                 Denne <%=periodeTxt %> kan IKKE <%=lukTxt%> f�r den er afsluttet af medarbejder.<br />
                
            

                <%if len(trim(request("showadviseringmsg"))) <> 0 then %>
                <br />
                  <div style="color:#000000; font-size:11px; background-color:#cccccc; padding:20px;"><b>Besked afsendt!</b></div>
                <%else %>
                  <br />
                  <b>Send email</b> med besked om at <%=periodeTxt %> mangler af blive afsluttet<br /><br />
                <input id="Submit1" type="submit" value="Send besked >>" style="font-size:9px; width:120px;" />
                <%end if %>
           
             
               </td></tr></table>
                </form>
               <%end if 
               
               
               if cint(ugegodkendt) <> 2 AND cint(showAfsuge) = 0 then%>

               <form action="<%=fmLink%>&func=afvisugeseddel" method="post">
               <table width=90% cellpadding=0 cellspacing=0 border=0>
               <tr><td class=lille>
               <br />
               <span style="font-size:11px;"><b>Afvis denne <%=periodeTxt %></b></span><br />
               Begrundelse:<br />
               <textarea name="FM_afvis_grund" style="width:200px; height:40px;"></textarea><br /><br />
                <input id="Submit3" type="submit" value="Afvis <%=periodeTxt %> >>" style="font-size:9px; width:120px;" /><br />
                 <span style="color:#999999;">Afsender email til medarbejer om at ugeseddel er afvist, og �bner evt. allerede <%=lukTxt1 %> ugeseddel op igen.</span>
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
          
        	<span style="color:#000000; background-color:#DCF5BD; padding:5px;"><a href="#" id="s1_k">[+] <%=afslutTxt %></a></span>&nbsp;

<%end function

    
sub sweekmsg
%>
 <br /><b><%=session("user") &"</b> "& lcase(tsa_txt_087) %>?</b><br /><%=tsa_txt_383 %> 

	           
<%
end sub




'********************************************
'** Afslut uge funktionen '*****************
'*********************************************
public erDetteuge53
function afslutuge(weekSelected, visning, tjkDag7, rdir, SmiWeekOrMonth)
    
    'response.write "weekSelected: "& weekSelected
    'varTjDatoUS_man = dateAdd("d", -7, varTjDatoUS_son) 


    if thisfile = "ugeseddel_2011.asp" then
    menu2015lnk = "../timereg/"
    else
    menu2015lnk = ""
    end if

   
    
    %>


<%=menu2015lnk%>timereg_akt_2006.asp?func=opdatersmiley&fromsdsk=<%=fromsdsk%>&rdir=<%=rdir %>&varTjDatoUS_man=<%=varTjDatoUS_man %>&usemrn=<%=usemrn%>
<form action="<%=menu2015lnk%>timereg_akt_2006.asp?func=opdatersmiley&fromsdsk=<%=fromsdsk%>&rdir=<%=rdir %>&varTjDatoUS_man=<%=varTjDatoUS_man %>&usemrn=<%=usemrn%>" method="POST">
<table border='0' cellspacing='0' cellpadding='0' width="100%">

<tr>
   
	<td valign=top bgcolor="#FFFFFF" style="padding-top:0px;">

        
   
	<%
	
        select case cint(SmiWeekOrMonth)
        case 0 'uge aflsutning

	        if visning = 0 then 'denne uge
            sidsteUgeTxt = "valgte uge"
            sidstedagisidsteuge = dateadd("d", -7, weekSelected)
            else
            sidsteUgeTxt = "sidste uge"
            sidstedagisidsteuge = dateadd("d", -7, weekSelected)
            end if

        case 1
        sidsteUgeTxt = "sidste m�ned"
        sidstedagisidsteuge = weekSelected 'dateadd("d", -31, weekSelected)
        case 2
        sidsteUgeTxt = "sidste dag"
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
        call rettidigafsl(sidstedagisidsteuge)

        'Response.write "sidstedagiuge: " & sidstedagiuge & "<br>"

        '** tjekker om uge er afsluttet og viser afsluttet eller form til afslutning
		call erugeAfslutte(detteaar, sidstedagiuge, usemrn, SmiWeekOrMonth)
        
        '** Viser liste over afsluttede uger
        call afsluttedeUger(detteaar, usemrn)
		

        'response.write "sidsteUgenrAfslFundet: "& sidsteUgenrAfslFundet
        
                    if cint(sidsteUgenrAfslFundet) = 1 then
                        if cint(SmiWeekOrMonth) = 0 then 'uge aflsutning
                        
                            'if datePart("ww", sidsteUgenrAfsl, 2,2) = 53 then
                            'sidstedagisidsteAfsluge = dateAdd("d", 0, sidsteUgenrAfsl)
                            'else
                            sidstedagisidsteAfsluge = dateAdd("d", 7, sidsteUgenrAfsl) 
                            'end if

                        else
                        sidstedagisidsteAfsluge = dateAdd("m", 1, sidsteUgenrAfsl)
                        sidstedagisidsteAfsluge = dateAdd("d", -1, sidstedagisidsteAfsluge)  
                        end if
                    else
                    sidstedagisidsteAfsluge = sidsteUgenrAfsl 
                    end if
            

                    'response.write "sidstedagisidsteAfsluge:" & sidstedagisidsteAfsluge & " sidsteUgenrAfsl: " & sidsteUgenrAfsl
                    'response.write "UGE: "& datePart("ww", sidsteUgenrAfsl, 2,2)

                    if cint(SmiWeekOrMonth) = 1 then 'm�neds afslutning ==> Altid sidste dag i m�ned
                
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


                    end if


                    if cint(SmiWeekOrMonth) = 0 then 'uge aflsutning
                    '*** Sikrer at sidste ugedag der er fundet altid er s�ndag. (sidstedag i ugen) *** 
                    '*** Afslut p� m�nedsbasis skal altid v�re den sidste dag i md. og ikke n�dvendigvis s�ndag
                    dayOfWeekThis = datePart("w",sidstedagisidsteAfsluge,2,2)
                    '1 = (dayOfWeek - d) 
                    dThis = (dayOfWeekThis - 1)
                    firstdayOfFirstWeekThis = dateAdd("d", -dThis, sidstedagisidsteAfsluge)
                    firstThursdayOfFirstWeekThis = dateAdd("d", +6, firstdayOfFirstWeekThis)
            
                    sidstedagisidsteAfsluge = firstThursdayOfFirstWeekThis
                    'sidstedagisidsteAfsluge = dateAdd("d", 7, sidstedagisidsteAfsluge)
                    end if



        call meStamdata(usemrn)
        
        sidstedagisidsteugeTjk = sidstedagisidsteuge
        if (cDate(meAnsatDato) > cDate(sidstedagisidsteugeTjk) OR cDate(meOpsagtDato) < cDate(sidstedagisidsteugeTjk)) then
        '*** Funktione endnu ikke aktiv / 

                    select case lto
                    case "xtec", "xintranet - local"

                      %>
                    <div style="background-color:#ffdfdf; padding:10px;">
                        <b>funktionen endnu ikke aktiv hos jer.</b><br />
                     Vi arbejder p� at tilpasse funktionen, s� I kan afslutte jeres perioder p� m�nedsbasis.<br /> 
                     I vil f� n�rmere besked n�r I skal begynde p� m�neds-afslutning.
             
          
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

                    
				'*** Man m� gerne lukke uger frem i tiden 27012009 ***'
				'if (datepart("ww", tjekdag(7), 2, 2) <= datepart("ww", now, 2, 2) AND year(tjekdag(7)) <= year(now)) OR  year(tjekdag(7)) < year(now) then%>
				<input type="hidden" name="FM_mid" id="FM_mid" value="<%=usemrn%>">
				<input type="hidden" name="FM_afslutuge_sidstedag" id="FM_afslutuge_sidstedag" value="<%=sidstedagisidsteAfsluge%>"><!-- sidstedagisidsteAfsluge -->
     
                <%if datepart("ww", sidstedagisidsteuge, 2, 2) >= 1 then ' kan ikke afslutte uge 1? eller
                %>
                <span style="font-size:11px;"><%=tsa_txt_409 %>:</span><br />
                <span style="font-size:14px; font-weight:bolder;"> 

                   <!--
                  sidstedagisidsteAfsluge:  <%=sidstedagisidsteAfsluge %><br />
                   sidstedagisidsteuge: <%=sidstedagisidsteuge %><br />
                   ww: <%=datepart("ww", sidstedagisidsteuge, 2, 2) %>
                   -->
                  
               
               
                <%select case cint(SmiWeekOrMonth) 
                case 0 'uge aflsutning %>
                <%=tsa_txt_005 %>: <%=datePart("ww", sidstedagisidsteAfsluge, 2,2) %>, 
                   
                   <%if datePart("ww", sidstedagisidsteAfsluge, 2,2) = 53 then%> 
                   <%=datePart("yyyy", dateAdd("yyyy", -1, sidstedagisidsteAfsluge), 2,2) %>   / <%=datePart("yyyy", sidstedagisidsteAfsluge, 2,2) %>  
                   <%else%>
                   <%=datePart("yyyy", sidstedagisidsteAfsluge, 2,2) %>
                   <%end if %>

                <%case 1 'MD %>
                <%=monthname(datePart("m", sidstedagisidsteAfsluge, 2,2)) %>, <%=datePart("yyyy", sidstedagisidsteAfsluge, 2,2) %>
                <%case 2 'Dag %>
                <%=weekdayname(weekday(sidstedagisidsteAfsluge)) & " d. " & sidstedagisidsteAfsluge %>
                <%end select %>
                </span> 
        
                    <%select case cint(SmiWeekOrMonth) 
                    case 0, 1 
                    %>

                    <%if datepart("ww", tjkDag7, 2 ,2) <> datepart("ww", sidstedagisidsteAfsluge, 2 ,2) AND thisfile <> "ugeseddel_2011.asp" then  %>
                    (<a href="<%=menu2015lnk%>timereg_akt_2006.asp?showakt=1&strdag=<%=day(sidstedagisidsteAfsluge)%>&strmrd=<%=month(sidstedagisidsteAfsluge)%>&straar=<%=year(sidstedagisidsteAfsluge)%>" class="vmenu">se uge <%=datePart("ww", sidstedagisidsteAfsluge, 2,2) %>..</a>)</span> 
                    <%end if %> <br />
                        
                    <%
                    case 2
                       %><br /><%    
                    end select%>
                    
                    <%


                 '*** M� TEAMLEDER/Bruger AFSLUTTE uge for medarb ***
                 maaAfslutteUge = 0
                 select case lto
                 case "tec", "esn"
                        
                        if session("mid") = usemrn OR level = 1 then
                        maaAfslutteUge = 1
                        else
                        maaAfslutteUge = 0
                        end if

                 case else
                    maaAfslutteUge = 1
                 end select


                if cint(maaAfslutteUge) = 1 then
                    FM_afslutugeDIS = ""
                else
                    FM_afslutugeDIS = "DISABLED"
                end if
                '***************************************************


                '********** Der findes herrelosetimer / timer uden match fra f.eks TT. 
                '********** Admin m� heller ikke afslutte
                if (cint(afslutugekri) = 10) then 

                    FM_afslutugeDIS = "DISABLED"

                end if
                    

                   
                    
                %>
                <div style="padding:3px;">
                <%
                if (cint(afslutugekri) = 0 OR cint(level) = 1) AND cint(SmiWeekOrMonth) = 0 AND lto <> "tec" AND lto <> "esn" then 
                'M� ikke kune afslutte alle uger n�r der er Kriterier opstillet p� min. timer pr. uge
                'Eller hvis man afslutter p� m�nedsbasis (afslut alle m�ender til d.d funktion ikke klar)%>
                <input type="checkbox" name="FM_afslutuge" id="FM_afslutuge" value="1" onClick="showlukalleuger()" <%=FM_afslutugeDIS %>>
                <%else %>
                <input type="checkbox" name="FM_afslutuge" id="Checkbox1" value="1" <%=FM_afslutugeDIS %>>
                <%end if %>

                 
                   
                    
                    <%'if cint(SmiWeekOrMonth) = 0 then 'uge aflsutning %>
                    <%'=tsa_txt_089%>
                    <%'else %>
                    <%'=tsa_txt_431 &" "& monthname(datePart("m", sidstedagisidsteAfsluge, 2,2)) =datePart("yyyy", sidstedagisidsteAfsluge, 2,2) %>
                    <%'end if 
                
                     select case cint(SmiWeekOrMonth) 
                     case 0
                     afslutTxtbtn = tsa_txt_431 '&" "& tsa_txt_345
                     case 1
                     afslutTxtbtn = tsa_txt_431 
                     case 2
                     afslutTxtbtn = tsa_txt_431 '&" "& left(tsa_txt_275, 3) 'Dag
                     end select

                 

                 if cint(maaAfslutteUge) = 1 AND cint(afslutugekri) <> 10 then%>
                   
                 <input type="submit" value="<%=afslutTxtbtn &" "& tsa_txt_432 %> >>">
                <%end if %>

                </div>
                
         
                <br />
                          <%if cint(SmiWeekOrMonth) = 1 then 'MD aflsutning %>
                          
                        
                               <%if SmiantaldageCount = 8 then %>
                              <%=tsa_txt_410 &" "& tsa_txt_429 &" "& weekafslTxt %>
                              <%end if %>

                              <%if SmiantaldageCount = 9 then %>
                              <%=tsa_txt_410 &" "& tsa_txt_429 &" + 7 d. "& weekafslTxt %>
                              <%end if %>

                                <%if SmiantaldageCount = 10 then %>
                              <%=tsa_txt_410 &" "& tsa_txt_429 &" + 3 d. "& weekafslTxt %>
                              <%end if %>

                          <%else %>
                           <%=tsa_txt_410 %> <b><%=left(weekdayname(weekday(cDateUge,1)), 3) &".  kl. "& formatdatetime(cDateUge, 3) %> </b> <%=weekafslTxt %>
                          <%end if %>

                         
                      
                             

                <%end if %>
				
				
                   
                	<div id="lukalleuger" name="lukalleuger" style="position:relative; visibility:hidden; display:none; width:300px;">
                        
                <%LastWeekSelected = dateAdd("d", -7, now) %>
				<input type="checkbox" name="FM_alleuger" id="FM_alleuger" value="1">&nbsp;<span><%=tsa_txt_090 %>: <%=datePart("ww", LastWeekSelected, 2,2) %>, <%=datePart("yyyy", LastWeekSelected, 2,2) %></span>
				</div>
				 
        
             <%
             '** Status p� antal registrederede projekttimer i den p�g�ldende uge    
             select case lto
             case "tec", "esn"
             case else %>
            <br /><br />
                <%if cint(afslProcKri) = 1 then
                   sm_bdcol = "#DCF5BD"
                   else
                    sm_bdcol = "mistyrose"
                   end if %>



                <div style="color:#000000; background-color:<%=sm_bdcol%>; padding:10px; border:0px <%=sm_bdcol%> solid;"><%=tsa_txt_398 %>: <b><%=totTimerWeek & " "%></b> 
                    
                    <%if afslutugekri = 2 then %>
                    <%=tsa_txt_399 %>
                    <%end if %>

                    <%=" "&tsa_txt_140 %> / <%=afslutugeBasisKri %> = <b><%=afslProc %> %</b> <%=tsa_txt_095 %> <b><%=datePart("ww", tjkTimeriUgeDtTxt, 2,2) %></b>
                      <br />(<%=left(weekdayname(weekday(formatdatetime(tjkTimeriUgeDt, 2))), 3) &". "& formatdatetime(tjkTimeriUgeDt, 2)%> - <%= left(weekdayname(weekday(formatdatetime(dateAdd("d", 6, tjkTimeriUgeDt), 2))), 3) &". "&formatdatetime(dateAdd("d", 6, tjkTimeriUgeDt), 2) %>)
               


                   </div>

                <!-- DER FINDES TIMER UDEN MATCH FRA F.eks TT -->
                <%if afslutugekri = 10 then %>
                <div style="color:#000000; background-color:#ffdfdf; padding:10px;">
                <b>Periode kan ikke afsluttes</b> da der findes herrel�se timer uden match.<br /> Disse timer er enten uploadet via excel eller indtastet via f.eks TimeTag.<br /><br />
                Der er bl.a timer d. <b><%=afslugeDatoTimerudenMatch %></b>
                 </div>
                <%end if %>
            
            <%end select %>


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
        
		<span style="font-size:14px; font-weight:bolder;"><%=perTxt%>&nbsp;<%=lcase(tsa_txt_093) %> <font color=green><i>V</i></font></span>
		<br><div style="font-size:11px; padding-top:5px;"><%=tsa_txt_093 %>&nbsp;<%=weekdayname(weekday(cdAfs))%>&nbsp;<%=tsa_txt_092 %>&nbsp;<%=formatdatetime(cdAfs, 2)%>&nbsp;<%=formatdatetime(cdAfs, 3)%> 
		(<%=tsa_txt_095 %>&nbsp;<%=datepart("ww", cdAfs, 2, 2)%>)</div>
		<%end if%>


       

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

    's�tter useDateSmileyTjk = sidste uge
	if visning = 1 then 'sidste uge
    useDateSmileyTjk = dateadd("d", -7, weekSelected)
    useDateSmileyTjk2 = dateadd("d", -8, weekSelected) '14
	else
    useDateSmileyTjk = dateadd("d", -7, weekSelected)
    useDateSmileyTjk2 = dateadd("d", -8, weekSelected)
    end if


    call meStamdata(usemrn)


    
    if year(meAnsatDato) >= year(useDateSmileyTjk) then
            
            'if year(meAnsatDato) >= year(now) then
            surDatoSQLSTART = year(meAnsatDato)&"/"& month(meAnsatDato)&"/"& day(meAnsatDato)

            if cint(SmiWeekOrMonth) = 0 then 'uge aflsutning
            useDateSmileyTjkWeek = dateDiff("ww", meAnsatDato, now, 2, 2)
            else
            useDateSmileyTjkWeek = dateDiff("m", meAnsatDato, now, 2, 2)
            end if
        
            'else
            'surDatoSQLSTART = year(now)&"/1/1"
            'useDateSmileyTjkWeek = dateDiff("ww", surDatoSQLSTART, now, 2, 2)
            'end if
    
     else
                    '** F�rste Torsdag s� vi er sikker p� at vi er inde i uge 1
                    surDatoSQLSTART = year(useDateSmileyTjk)&"/1/1"

                    if cint(SmiWeekOrMonth) = 0 then 'uge aflsutning
                    dayOfWeekThis = datePart("w",surDatoSQLSTART,2,2)

                    dThis = (dayOfWeekThis - 1)
                    firstdayOfFirstWeekThis = dateAdd("d", -dThis, surDatoSQLSTART)
                    surDatoSQLSTART = firstdayOfFirstWeekThis 'dateAdd("d", +3, firstdayOfFirstWeekThis)
                    surDatoSQLSTART = year(surDatoSQLSTART)&"/"& month(surDatoSQLSTART)&"/"& day(surDatoSQLSTART)

                    useDateSmileyTjkWeek = datepart("ww", useDateSmileyTjk, 2, 2)
                    
                    else
                    useDateSmileyTjkWeek = datepart("m", useDateSmileyTjk, 2, 2)
                    end if
    
    end if


    if cDate(meOpsagtDato) < cDate(useDateSmileyTjk2) then
       surDatoSQLEND = year(meOpsagtDato)&"/"& month(meOpsagtDato)&"/"& day(meOpsagtDato)
    else
       surDatoSQLEND = year(useDateSmileyTjk2)&"/"&month(useDateSmileyTjk2)&"/"&day(useDateSmileyTjk2)
    end if
	
	
	antalAfsDato = 0
    antalAfsl = 0

	strSQL3 = "SELECT uge, count(*) AS antalafs FROM ugestatus "_
	&" WHERE mid = "& usemrn &" AND uge BETWEEN '"& surDatoSQLSTART &"' AND '"& surDatoSQLEND &"'"_
	&" GROUP BY mid ORDER BY afsluttet DESC"
    'if lto = "dencker" then
	'Response.write strSQL3 & "<br><br>useDateSmileyTjkWeek: "& useDateSmileyTjkWeek & "<br>" & strSQL3 &"<br>"
	'Response.flush
    'end if
	oRec3.open strSQL3, oConn, 3 
	if not oRec3.EOF then

        if isNull(oRec3("antalafs")) <> true then

       
        'if session("mid") = 1 then
        'response.write "afl: #" & oRec3("antalafs") & "#<br>"
        antalAfsl = oRec3("antalafs")
        'response.write "antalAfsl: #" & antalAfsl & "#<br>"
        'antalAfsl = antalAfsl/1
        'response.Flush
        'end if
        antalAfsDato = antalAfsDato + antalAfsl
        
        'response.write antalAfsDato & "<br>" 
        'response.flush

        else
         antalAfsl = 0
         antalAfsDato = antalAfsDato
        end if

		
	end if
	oRec3.close  
	
	'if lto = "dencker" then
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
       
		<h4><%=meNavn &" ["& meInit %>] <span style="color:#999999; font-size:9px; font-weight:lighter;"> - <%=tsa_txt_404 %>: <%=formatdatetime(meAnsatDato, 2) %></span><br />
		 <span style="color:#000000; font-size:11px;  font-weight:lighter;">
             <%select case cint(SmiWeekOrMonth)
             case 0 'uge aflsutning %>
             <%=tsa_txt_411 &" "& lcase(tsa_txt_005) &" "& datepart("ww", useDateSmileyTjk2, 2, 2) &" "& year(useDateSmileyTjk2)%>
             <%case 1 %>
             <%=tsa_txt_411 &" "& monthname(datepart("m", useDateSmileyTjk2, 2, 2)) &" "& year(useDateSmileyTjk2)%>
             <%case 2 %>
             <%=tsa_txt_411 &" "& useDateSmileyTjk2 %>
             <%end select %>
            

		 </span></h4>
		<%	

        if cint(hidesmileyicon) = 0 then 'skjul Smiley Icons

		if cint(surSmil) = 1 then%>
				
				<b><%=tsa_txt_096 %></b><!--(du har afsluttet: <b><%=antalAfsDato %> / <%=useDateSmileyTjkWeek %></b> uger)--></td><td align=right style="padding:2px 5px 0px 0px;">
                <a href="#" class="sA1_k" style="color:#999999;">X</a></td></tr>
	            <tr><td colspan=2 style="padding-top:3px;">
                
				<img src="../ill/sur_<%=smilVal%>.gif" alt="" style="border:0px;"><br>
                
				
                    

                     <%if cint(SmiWeekOrMonth) = 0 then 'uge aflsutning %>
                     <span style="color:red;"><%=tsa_txt_403 %>!
                    <%else %>
			         <span style="color:red;"><%=tsa_txt_427 %>!
                    <%end if %>

                    <br /><%=tsa_txt_428 %>
                         </span>

				
				
				<%
		else
		
				
				%>
				<b><%=tsa_txt_098 %></b>
				</td><td align=right style="padding:2px 5px 0px 0px;"><a href="#" class="sA1_k" style="color:#999999;">X</a></td></tr>
	            <tr><td colspan=2 style="padding-top:3px;">

                  

				<img src="../ill/gladsmil_<%=smilVal%>.gif" alt="" style="border:0px;"><br>
                   
				
				
				<%
		end if%>
		

                   
		<!--<font class=megetlillesort><br>Nb: Smileyordning kan sl�s til og fra i medarbejder-profilen.</font>-->
		
	
	
	
	</td></tr>

        <%end if 'smiley icon %>

		</table>
<%end function%>



<%



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
        
					strSQL2 = "SELECT u.status, u.afsluttet, u.uge, u.id, u.ugegodkendt FROM ugestatus u WHERE u.mid = "& oRec("mid") &" AND (uge BETWEEN '"& startdato &"' AND '"& slutdato &"') "& tjkAnsatDatoKri &" ORDER BY uge" 
					'AND (WEEK(uge,1) BETWEEN 1 AND "& denneuge &" 
					
                    'if session("mid") = 1 then
					'Response.write strSQL2 &"<br>"
	                'end if				

					v = 0
                    lastWrtBlnk = 0
					oRec2.open strSQL2, oConn, 3 
					while not oRec2.EOF 
			
    
                    if cint(SmiWeekOrMonth) = 0 then
                    ugeMdNrTxt = datepart("ww", oRec2("uge"), 2, 2)
                    ugeMdNrTxtTopKri = 52
                    else
                    ugeMdNrTxt = datepart("m", oRec2("uge"), 2, 2)
                    ugeMdNrTxtTopKri = 12
                    end if

                    
    

                    if media <> "export" then
    
    		
					        if v = 0 AND ugeMdNrTxt = ugeMdNrTxtTopKri AND cint(SmiWeekOrMonth) = 0 then
					
					        else

            
					
					                    if oRec2("status") = 2 then
					                    glade(x) = glade(x) + 1

                                        
                                                if print <> "j" then
                           
					        
					                            if level = 1 AND (thisfile <> "timereg_akt_2006" AND thisfile <> "logindhist_2011.asp" AND thisfile <> "ugeseddel_2011.asp") then
						                        ugerafsluttet(x,v) = "<a href='smileystatus.asp?func=slet&id="&oRec2("id")&"' class=vmenuglobal>"& ugeMdNrTxt &"</a>"
					                            else
					                            ugerafsluttet(x,v) = "<b>"& ugeMdNrTxt &"</b>"
					                            end if

                                                else

                                                      ugerafsluttet(x,v) = "<b>"& ugeMdNrTxt &"</b>"

                                                end if
					        
					                    else
					                    sure(x) = sure(x) + 1

                                                 if print <> "j" then            

					                            if level = 1 AND (thisfile <> "timereg_akt_2006" AND thisfile <> "logindhist_2011.asp" AND thisfile <> "ugeseddel_2011.asp") then
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
                                        ugerafsluttet(x,v) = ugerafsluttet(x,v) & "<i style=""color:green; font-weight:bolder;"">V</i>"
                                        case 2
                                        ugerafsluttet(x,v) = ugerafsluttet(x,v) & "<i style=""color:red; font-weight:bolder;"">%</i>"
                                        end select



                   

					      
					
					        end if

                    else
                                


                    if cint(SmiWeekOrMonth) = 0 then
                    ugeMdNrTxt = datepart("ww", oRec2("uge"), 2, 2)
                    ugeMdNrTxtTopKri = 52
                    else
                    ugeMdNrTxt = datepart("m", oRec2("uge"), 2, 2)
                    ugeMdNrTxtTopKri = 12
                    end if
            
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

<table border=0 cellspacing=0 cellpadding=0 width="<%=twdt%>">

<tr bgcolor="#5c75AA">
    <td>&nbsp;</td>
	<td width=400 class=alt><b><%=tsa_txt_101 %></b></td>
	<td class=alt align=right>Afsluttet<br /> til tiden</td>
	<td class=alt align=right>Afsluttet<br /> for sent</td>
	<td class=alt align=right><b><%=tsa_txt_104 %></b></td>
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
				<td><b><%=mnavn(x)%> [<%=mnr(x)%>]</b> <span style="color:#999999; font-size:9px;"> - <%=tsa_txt_404 %>: <%=formatdatetime(mansatDato(x), 2) %></span>
				</td>
				<td align=right><b><%=glade(x)%></b></td>
				<td align=right><%=sure(x)%></td>
				<td align=right><%=glade(x)+sure(x)%> / <%=antalUger%></td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			</tr>
			<tr bgcolor="<%=trbg %>"><td colspan=7 style="padding:5px;">
				
				<%
				for v = 0 to UBOUND(ugerafsluttet, 1)
				if len(ugerafsluttet(x,v)) <> 0 then%>
					
					<%if right(v, 1) = 0 AND v > 0 then%>
					<br>
					<%end if%>
					
					<%if lastyear <> yearaf(x,v) then %>
					<label style="color:#5582d2;"><b><%=yearaf(x,v)%></b></label>:
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
    	


    '***** Viser PopUp med status p� periode afslutning '*****'
	sub opdaterSmiley
	
	
	
    '*** Afslutter uge ****'
        		
    if len(request("FM_afslutuge")) <> 0 then
		        
          
        'sidstedagisidsteuge
		       call rettidigafsl(request("FM_afslutuge_sidstedag")) 'kontrolpanel settings, uge eller m�nedbasis, dag mv
		        
		     
		       
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
		        smileysttxt = "<span style=""color:green;""> er afsluttet korrekt&nbsp;&nbsp;<img src=""../ill/check2.png"" border=0> </span>"
                smileyImg = "gladsmil_2.gif"
		        else
		        intStatusAfs = 1 '** Afsluttet forsent
		        smileysttxt = "<span style=""color:red;""> er afsluttet for sent! </span>"
                smileyImg = "sur_1.gif"
		        end if
		        
		       
        	    'cDateUgeSQL = year(cDateUge) & "/" & month(cDateUge) &"/"& day(cDateUge)	
		        medarbejderid = request("FM_mid")
                call meStamdata(medarbejderid)
        		
		        strSQLafslut = "INSERT INTO ugestatus (uge, afsluttet, mid, status) VALUES ('"& cDateUgeTilAfslutning &"', '"& cDateAfs &"', "& medarbejderid &", "& intStatusAfs &")" 
		        
		        'Response.Write strSQLafslut
		        'Response.flush
		        
		        oConn.execute(strSQLafslut)
		       

                'response.end
		        
		        
		        
		        '*** Godkender automatisk timer i de afsluttede uger **'
		        
		        autogktimer = 0
		        strSQL = "SELECT autogktimer FROM licens WHERE id = 1"
		        oRec3.open strSQL, oConn, 3
		        if not oRec3.EOF then
		        
		        autogktimer = oRec3("autogktimer")
		        
		        end if
		        
		        if cint(autogktimer) = 1 then
		        
		        stDatoGKtimer = cDateUgeTilAfslutning
		        slDatoGKtimer = dateadd("d", -6, cDateUgeTilAfslutning)
		        
		        slDatoGKtimer = year(slDatoGKtimer) &"/"& month(slDatoGKtimer) &"/"& day(slDatoGKtimer)
		        
		        
		        strSQLGKtimer = "UPDATE timer SET godkendtstatus = 1, godkendtstatusaf = '"& session("user") &"' WHERE"_
		        &" tmnr = "& medarbejderid &" AND tdato BETWEEN '"& slDatoGKtimer &"' AND '"& stDatoGKtimer &"'"
		        
		        oConn.execute(strSQLGKtimer)
		       
        		end if
        		
        		
			        '*** Afslut alle uger frem til dd. ****
			        if len(request("FM_alleuger")) <> 0 then
			        afslutalle = request("FM_alleuger")
			        else
			        afslutalle = 0
			        end if
        			
			        
			        if afslutalle = 1 then
        					
                            

					        detteAar = year(request("FM_afslutuge_sidstedag"))
                            
                            sidsteUgeAfsl = request("FM_afslutuge_sidstedag")
                            nuMinusSyv = dateAdd("d", -7, now)
                            
                            'response.write "detteAar: "& detteAar & "<br>"    

                            dayOfWeekThis = datePart("w",now,2,2)
        
                            '1 = (dayOfWeek - d) 
                            dThis = (dayOfWeekThis - 1)
                            firstdayOfFirstWeekThis = dateAdd("d", -dThis, now)
                            firstThursdayOfFirstWeekThis = dateAdd("d", +6, firstdayOfFirstWeekThis)
            
                            sidstedagisidsteAfsluge = firstThursdayOfFirstWeekThis


                            denneugeLast = dateAdd("d", -7, sidstedagisidsteAfsluge)
					        denneuge = datepart("ww", denneugeLast, 2, 2) 'cDateUge
					        denneugeOpr = denneuge
        					
                            'response.write "nuMinusSyv: "& nuMinusSyv & "<br>"

                            PeriodeInterVal = dateDiff("ww", sidsteUgeAfsl, nuMinusSyv, 2,2)

                            'response.write "PeriodeInterVal: "& PeriodeInterVal & "<br>"

					        for u = 0 to PeriodeInterVal 'denneuge 
        					
					        nyDateUge = dateadd("d", (u*(-7)), denneugeLast) 'cDateUge
					        denneuge = datepart("ww", nyDateUge, 2, 2)
					        detteaar = datepart("yyyy", nyDateUge, 2,3)
        					
					        ugefindes = 0
					        strSQL2 = "SELECT u.id, u.status, u.afsluttet, u.uge FROM ugestatus u WHERE YEAR(u.uge) = '"& detteaar &"' AND  WEEK(u.uge, 1) = '"& denneuge &"' AND u.mid = "& medarbejderid
					        'Response.write strSQL2 & "<br>"
        					
					        oRec2.open strSQL2, oConn, 3 
					        if not oRec2.EOF then
					        ugefindes = 1 
					        'Response.write oRec2("id") &" :: "& oRec2("uge") & " - "& datepart("ww", oRec2("uge"), 2, 2) & " - ok!<br>"
				            end if
					        oRec2.close 
        					
					        'Response.Write  denneuge &"<"& denneugeOpr &" AND ugefindes: "& ugefindes &" nyDateUge:"&nyDateUge&" uge: "& datepart("ww", nyDateUge, 2,2) &" meAnsatDato: "& meAnsatDato &" ansatuge: "& datepart("ww", meAnsatDato, 2,2) &"<br>"
                            'response.flush 
                      
        					'cint(denneuge) <= cint(denneugeOpr) AND 
						            if cint(ugefindes) = 0 AND ((datepart("ww", nyDateUge, 2,2) >= datepart("ww", meAnsatDato, 2,2) AND datepart("yyyy", nyDateUge, 2,2) = datepart("yyyy", meAnsatDato, 2,2)) OR datepart("yyyy", nyDateUge, 2,2) > datepart("yyyy", meAnsatDato, 2,2)) then
							        
                                    sqlDateUge = year(nyDateUge)&"/"&month(nyDateUge)&"/"&day(nyDateUge)
        		
							        intStatusAfs = 1 '** Afsl. forsent
							        strSQLafslut = "INSERT INTO ugestatus (uge, afsluttet, mid, status) VALUES ('"& sqlDateUge &"', '"& cDateAfs &"', "& medarbejderid &", "& intStatusAfs &")" 
							        oConn.execute(strSQLafslut)


                                    'response.write "strSQLafslut:<br> "& strSQLafslut & "<br><hr>"

							        'Response.write sqlDateUge &"-"& datepart("ww", nyDateUge, 2, 2) &" Afslutte uge<br>"
							                
							                
							                if cint(autogktimer) = 1 then
		        
		                                    stDatoGKtimer = sqlDateUge
		                                    slDatoGKtimer = dateadd("d", -6, sqlDateUge)
                            		        
		                                    slDatoGKtimer = year(slDatoGKtimer) &"/"& month(slDatoGKtimer) &"/"& day(slDatoGKtimer)
                            		        
                            		        
		                                    strSQLGKtimer = "UPDATE timer SET godkendtstatus = 1, godkendtstatusaf = '"& session("user") &"' WHERE"_
		                                    &" tmnr = "& medarbejderid &" AND tdato BETWEEN '"& slDatoGKtimer &"' AND '"& stDatoGKtimer &"'"
                            		        
                            		        
		                                    oConn.execute(strSQLGKtimer)
		                                    
		                                    
                            		       
        		                            end if
							                    
							        
							        end if
							        
							        
							        
        					
					        next
        					
        					
        					      'Response.end
        					
			        end if
	
	    end if
	    
	end sub 
	
	
	
	
	sub afslutMsgTxt
        %>
          <label style="color:#999999;">
          Du kan, som administrator eller leder/teamleder, godkende eller afvise de enkelte timeregisteringer. N�r perioden bliver lukket/godkendt, kan kun administator �ndre.<br /><br />
          En uge kan f�rst godkendes eller afvises, n�r medarbejderen har afsluttet perioden. En medarbejder kan �ndre i sine indtastninger indtil ugen er lukket/godkendt af lederen.</label>

        <%
    end sub
	
    




    '********** Er uge afsluttet??? ******'
	
    public showAfsugeTxt
    public showAfsugeTxt_tot, ugegodkendtTxt_tot, btnstyle, ugeNrStatus

    public ugeNrAfsluttet, showAfsuge, cdAfs, ugegodkendt, ugegodkendtaf, ugegodkendtTxt, ugegodkendtdt
    function erugeAfslutte(sm_aar, sm_sidsteugedag, sm_mid, SmiWeekOrMonth)
            
          
              'call smileyAfslutSettings()
             ugegodkendt = 0
             showAfsuge = 1
             ugeNrAfsluttet = "1-1-2044"
             ugeNrStatus = 0


             'if session("mid") = 1 then
             response.write "<br>Thisfile: "& thisfile  &" HER: sm_sidsteugedag: "& sm_sidsteugedag & " SmiWeekOrMonth "& SmiWeekOrMonth & "sm_aar: "& sm_aar
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

                

            else
            
                sqlDatoKri = periodeKri &" = '"& sm_sidsteugedag & "'"    

            end if


            strSQLafslut = "SELECT status, afsluttet, uge, ugegodkendt, ugegodkendtaf, ugegodkendtTxt, ugegodkendtdt FROM ugestatus WHERE "_
            &" "& sqlDatoKri &" AND YEAR(uge) = "& sm_aar &" AND mid = "& sm_mid 
		    
            'if session("mid") = 1 then
            Response.write "<br>"& strSQLafslut & "<br><br>" '& "//" & sm_sidsteugedag & "//"& sldatoSQL
		    Response.flush
            'end if
            
            oRec3.open strSQLafslut, oConn, 3 
		    if not oRec3.EOF then
    			
			    showAfsuge = 0
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
            'Response.write "<br>ugeNrAfsluttet : "& ugeNrAfsluttet  & " ugegodkendt: "& ugegodkendt  '& "//" & sm_sidsteugedag & "//"& sldatoSQL
		    'Response.flush
            'end if

            'response.write "<br>#showAfsuge: "& showAfsuge


    end function



    public sidsteUgenrAfsl, sidsteUgenrAfslFundet
     function afsluttedeUger(sm_aar, sm_mid)

        
            call meStamdata(sm_mid)
            call licensStartDato()
        
            
            sidsteUgenrAfslFundet = 0

            if cDate(meAnsatDato) > cDate(licensstdato) then
            sidsteUgenrAfsl = meAnsatDato 
            else
            sidsteUgenrAfsl = licensstdato '"1-1-2002"
            end if        

            'if cDate(sidsteUgenrAfsl) > cDate("1-1-"&year(now)) then
            'sidsteUgenrAfsl = sidsteUgenrAfsl
            'else
            'sidsteUgenrAfsl = "1-1-"&year(now)
            'end if


            ansatDatoSQL = year(meAnsatDato) &"/"& month(meAnsatDato) &"/"& day(meAnsatDato)
            licensstDatoSQL = year(licensstdato) &"/"& month(licensstdato) &"/"& day(licensstdato) 
            
            strSQLafslut = "SELECT status, afsluttet, uge, ugegodkendt, ugegodkendtaf, ugegodkendtTxt, ugegodkendtdt FROM ugestatus WHERE "_
            &" ((YEAR(uge) = "& sm_aar &" AND uge >= '"& ansatDatoSQL &"' AND uge >= '"& licensstDatoSQL &"') OR (YEAR(uge) < "& sm_aar &")) AND (mid = "& sm_mid & ")  ORDER BY uge DESC LIMIT 1"
		    
         
            oRec3.open strSQLafslut, oConn, 3 
		    while not oRec3.EOF 
    			
            sidsteUgenrAfslFundet = 1
            sidsteUgenrAfsl = oRec3("uge")
			   
    		
            oRec3.movenext
		    wend
		    oRec3.close 
            
            
    end function


function godekendugeseddel(thisfile, godkenderID, medarbid, stDatoSQL)

        call smileyAfslutSettings()
       
        'gkweekfundet = 0
        ugegodkendtdtnow = year(now) & "/" & month(now) & "/" & day(now)
        denneuge = datepart("ww", stDatoSQL, 2,2)
	    detteaar = datepart("yyyy", stDatoSQL, 2,2)
        dettemd = datepart("m", stDatoSQL, 2,2)

        lastDayInWeek = dateAdd("d", 6, stDatoSQL)

        '** tjekker om de er uge 52/53 ell. 1 og om der skal korrigeres p� �r.
        if cint(denneuge) = 1 AND (datepart("yyyy", stDatoSQL, 2,2) <> datepart("yyyy", lastDayInWeek, 2,2)) OR (denneuge = 53 AND cint(SmiWeekOrMonth) = 0) then 'datepart("d", stDatoSQL, 2,2) > 29 then  
        '29 er den tidligst mulige dato uge 1 kan starte p�, hvor det er akturelt at skifte �r fordi s�ndagen der afslutter �ret ligger i det nye �r.
        detteaarAdd = dateAdd("yyyy", 1, stDatoSQL)
        detteaar = datepart("yyyy", detteaarAdd, 2,2)

             if cint(SmiWeekOrMonth) = 1 then
             dettemdAdd = dateAdd("m", 1, stDatoSQL)
             dettemd = datepart("m", dettemdAdd, 2,2)
             end if

        end if

         ugstatusSQL = "SELECT id, mid FROM ugestatus AS u WHERE mid = "& medarbid & " AND YEAR(u.uge) = '"& detteaar &"'"

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


                            '*** Afvis ugeseddel �nder ikke statusd p� timerne ***'
                            'strSQLup = "UPDATE timer SET godkendtstatus = 2, godkendtstatusaf = '"& afsNavn &"' WHERE tmnr = "& modtagerMid & " AND tdato BETWEEN '"& varTjDatoUS_man &"' AND '" & varTjDatoUS_son & "'" 
	                        'oConn.execute(strSQLup)

                            call smileyAfslutSettings()
        
                            ugegodkendtTxt = txt
                              
                              ugegodkendtdtnow = year(now) & "/" & month(now) & "/" & day(now)
                              denneuge = datepart("ww", varTjDatoUS_man, 2,2)
	                          detteaar = datepart("yyyy", varTjDatoUS_man, 2,2)

                              dettemd = datepart("m", varTjDatoUS_man, 2,2)


                              lastDayInWeek = dateAdd("d", 6, varTjDatoUS_man)

                                '** tjekker om de er uge 52/53 ell. 1 og om der skal korrigeres p� �r.
                                if cint(denneuge) = 1 AND (datepart("yyyy", varTjDatoUS_man, 2,2) <> datepart("yyyy", lastDayInWeek, 2,2)) OR (denneuge = 53 AND cint(SmiWeekOrMonth) = 0) then 'datepart("d", stDatoSQL, 2,2) > 29 then  
                                '29 er den tidligst mulige dato uge 1 kan starte p�, hvor det er akturelt at skifte �r fordi s�ndagen der afslutter �ret ligger i det nye �r.
                                detteaarAdd = dateAdd("yyyy", 1, varTjDatoUS_man)
                                detteaar = datepart("yyyy", detteaarAdd, 2,2)

                                     if cint(SmiWeekOrMonth) = 1 then
                                     dettemdAdd = dateAdd("m", 1, varTjDatoUS_man)
                                     dettemd = datepart("m", dettemdAdd, 2,2)
                                     end if

                                end if
                            

        
                            strSQLupUgeseddel = "UPDATE ugestatus SET ugegodkendt = 2, ugegodkendtaf = "& afsenderMid &", ugegodkendtdt = '"& ugegodkendtdtnow &"', ugegodkendtTxt = '"& ugegodkendtTxt &"' WHERE mid = "& modtagerMid 
                            
                            if cint(SmiWeekOrMonth) = 0 then
                    
                                    if cint(denneuge) = 53 then
                                    strSQLupUgeseddel = strSQLupUgeseddel &" AND YEAR(uge) = '"& detteaar &"' AND  (WEEK(uge, 3) = '"& denneuge &"' OR WEEK(uge, 3) = '0')"
                                    else
                                    strSQLupUgeseddel = strSQLupUgeseddel &" AND YEAR(uge) = '"& detteaar &"' AND  WEEK(uge, 3) = '"& denneuge &"'"
                                    end if
                            else
                            strSQLupUgeseddel = strSQLupUgeseddel &" AND YEAR(uge) = '"& detteaar &"' AND  MONTH(uge) = '"& dettemd &"'"
                            end if
	                        'Response.Write strSQLupUgeseddel
                            'Response.end
                            oConn.execute(strSQLupUgeseddel)
	    
	    'Response.Write strSQLup
	    'Response.end
	    

          if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\"&thisfile then

                                    

                                     '************************** TEST ***********************************
                                    'Sender notifikations mail
		                            'Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		                            ' S�tter Charsettet til ISO-8859-1
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
		                            '' S�tter Charsettet til ISO-8859-1
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
                                    thisWeekTxt = "Din ugeseddel uge:"
                                    thisWeekSbj = "uge"
                                    else
                                    thisWeek = monthname(datepart("m", varTjDatoUS_man, 2, 2)) 
                                    thisWeekTxt = "M�ned:"
                                    thisWeekSbj = "m�ned"
                                    end if   

					  	            'Sender notifikations mail
		                            'Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		                            ' S�tter Charsettet til ISO-8859-1
		                            'Mailer.CharSet = 2
		                            'Mailer.FromName = "TimeOut | " & afsNavn 
		                            'Mailer.FromAddress = afsEmail
		                            'Mailer.RemoteHost = "webout.smtp.nu" '"webmail.abusiness.dk" '"pasmtp.tele.dk"
						            '''Mailer.AddRecipient "" & afsNavn & "", "" & afsEmail & ""
		                            'Mailer.AddRecipient "" & modtNavn & "", "" & modtEmail & ""
            						
                                    
                                    'Mailer.Subject = "Du har endnu ikke afsluttet " '& lcase(thisWeekTxt) &" "& thisWeek 
		                            'strBody = "Hej " & modtNavn &  vbCrLf & vbCrLf
                                    'strBody = strBody &" - " & thisWeekTxt &" "& thisWeek &" - er endnu ikke afsluttet, husk at f� den afsluttet snarest." & vbCrLf & vbCrLf
                                   
		                            'strBody = strBody &"Med venlig hilsen" & vbCrLf
		                            'strBody = strBody & afsNavn & vbCrLf & vbCrLf
            		                
            		
            		
		                            'Mailer.BodyText = strBody
            		
		                            'Mailer.sendmail()
		                            'Set Mailer = Nothing


                    
                                    Set myMail=CreateObject("CDO.Message")
                                    myMail.Subject="TimeOut - Du har endnu ikke afsluttet "& thisWeekSbj & ": "& thisWeek
                                    myMail.From = "timeout_no_reply@outzource.dk"
				                     

                                    'myMail.To=strEmail
                                    if len(trim(modtEmail)) <> 0 then
                                    myMail.To= ""& modtNavn &"<"& modtEmail &">"
                                    end if

                                    strBody = "Hej " & modtNavn &  vbCrLf & vbCrLf
                                    strBody = strBody & thisWeekTxt &" "& thisWeek &" - er endnu ikke afsluttet, husk at f� den afsluttet snarest." & vbCrLf & vbCrLf
                                   
		                            strBody = strBody &"Med venlig hilsen" & vbCrLf
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



    
%>