<%response.Buffer = true 
tloadA = now
%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->

<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/convertDate.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/treg_func.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->




<%'GIT 20160811 - SK
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else
	
	function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
	end function
	
	func = request("func")
	id = request("id")
	thisfile = "logindhist_2011.asp"
    rdir = "lgnhist"
	media = request("media")

    'usemrn = request("usemrn")
    if len(trim(request("usemrn"))) <> 0 then
    usemrn = request("usemrn")
    else
    usemrn = session("mid")
    end if

    level = session("rettigheder")
 

    

    if len(trim(request("medarbsel_form"))) <> 0 then

     

    if len(trim(request("FM_visAlleMedarb"))) <> 0 then
    visAlleMedarbCHK = "CHECKED"
    visAlleMedarb = 1
    else
    visAlleMedarbCHK = ""
    visAlleMedarb = 0
    end if

    else


                    if request.cookies("tsa")("visAlleMedarb") = "1" then
                    visAlleMedarbCHK = "CHECKED"
                    visAlleMedarb = 1
                    else
                    visAlleMedarbCHK = ""
                    visAlleMedarb = 0
                    'usemrn = session("mid")
                    end if


    end if
    response.cookies("tsa")("visAlleMedarb") = visAlleMedarb


    call medarb_teamlederfor


    if len(trim(request("nomenu"))) <> 0 AND request("nomenu") <> "0" then
    nomenu = 1
    else
    nomenu = 0
    end if

   

    varTjDatoUS_man = request("varTjDatoUS_man")
    varTjDatoUS_son = dateAdd("d", 6, varTjDatoUS_man)
    'varTjDatoUS_son = request("varTjDatoUS_son")

   

    lnk = "&usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&varTjDatoUS_son="&varTjDatoUS_son&"&rdir="&rdir&"&nomenu="&nomenu


    'Response.Write varTjDatoUS_man
    'Response.end

    datoMan = day(varTjDatoUS_man) &"/"& month(varTjDatoUS_man) &"/"& year(varTjDatoUS_man)
    datoSon = day(varTjDatoUS_son) &"/"& month(varTjDatoUS_son) &"/"& year(varTjDatoUS_son)

    next_varTjDatoUS_man = dateadd("d", 7, varTjDatoUS_man)
	next_varTjDatoUS_son = dateadd("d", 7, varTjDatoUS_son)
    next_varTjDatoUS_man = year(next_varTjDatoUS_man) &"/"& month(next_varTjDatoUS_man) &"/"& day(next_varTjDatoUS_man)
	next_varTjDatoUS_son = year(next_varTjDatoUS_son) &"/"& month(next_varTjDatoUS_son) &"/"& day(next_varTjDatoUS_son)


    prev_varTjDatoUS_man = dateadd("d", -7, varTjDatoUS_man)
	prev_varTjDatoUS_son = dateadd("d", -7, varTjDatoUS_son)
    prev_varTjDatoUS_man = year(prev_varTjDatoUS_man) &"/"& month(prev_varTjDatoUS_man) &"/"& day(prev_varTjDatoUS_man)
	prev_varTjDatoUS_son = year(prev_varTjDatoUS_son) &"/"& month(prev_varTjDatoUS_son) &"/"& day(prev_varTjDatoUS_son)


	'Response.Write "media " & media
	
    d_end = 6
	
	'*** Sætter lokal dato/kr format. *****
	Session.LCID = 1030
	
	
    select case func 
	case "-"
	
	case "xxxopdaterlist"

    Response.write Request("logindhist_id") & "<br>"
    Response.write Request("logindhist_dato") & "<br>"
    Response.Write Request("logindhist_sttid")  & "<br>"
    Response.Write Request("logindhist_sltid")  & "<br>"

    ids = split(Request("logindhist_id"), ", ")
    datoers = split(Request("logindhist_dato"), ", ")
    stTiders = split(Request("logindhist_sttid"), ", ")
    slTiders = split(Request("logindhist_sltid"), ", ")

    for i = 0 TO UBOUND(ids)

    '** start tid **'
    dato_tidSt = datoers(i) &" "& stTiders(i)
    if len(trim(stTiders(i))) <> 0 then
        call logindStatus(usemrn, 1, 2, dato_tidSt)
    end if

    Response.Flush
    
     '** slut tid **'
    dato_tidSl = datoers(i) &" "& slTiders(i)
    if len(trim(slTiders(i))) <> 0 then
        call logindStatus(usemrn, 1, 2, dato_tidSl)
    end if
    
    next


    Response.end
	

    case else
       

    call akttyper2009(2)
	
	
	
	if media <> "export" then
	
	
	if media <> "print" then
	
    if nomenu <> 1 then

	leftPos = 90
	topPos = 102

    else

    leftPos = 20
	topPos = 48

    end if
	
	%>

	

     <%call browsertype()
    
     %>

	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
    <SCRIPT language=javascript src="inc/smiley_jav.js"></script>
    <SCRIPT language=javascript src="inc/logindhist_2011_jav.js"></script>
  

  <%
      if nomenu <> 1 then

    if media <> "print" then
    %>

     <div id="loadbar" style="position:absolute; display:; visibility:visible; top:260px; left:200px; width:300px; background-color:#ffffff; border:1px #cccccc solid; padding:2px; z-index:100000000;">

	<table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" />
	</td></tr></table>

	</div>

    <% end if
	

	
         call menu_2014() 


      end if

    
       
    else 
	
	leftPos = 0
	topPos = 0
	
	%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
    <SCRIPT language=javascript src="inc/logindhist_2011_jav.js"></script>
	<%end if%>
	
	<%end if%>
	

	
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:<%=leftPos%>px; top:<%=topPos%>px; visibility:visible;">

        
          <div style="position:absolute; background-color:#ffffff; border:0px #5582d2 solid; padding:4px; width:75px; top:-20px; left:730px; z-index:0;"><a href="<%=lnkTimeregside%>" class="vmenu"><%=left(tsa_txt_031, 7) %>.</a></div>
    
  
    <div style="position:absolute; background-color:#ffffff; border:0px #5582d2 solid; padding:4px; top:-20px; width:85px; left:820px; z-index:0;"><a href="../to_2015/<%=lnkUgeseddel%>" class="vmenu"><%=tsa_txt_337 %></a></div>
  


  <%
    
    call erTeamLederForvlgtMedarb(level, usemrn, session("mid"))
	
	call ersmileyaktiv()

    if media <> "print" then
    tTop = 0
	tLeft = 0
    else
    tTop = 0
	tLeft = 0
    end if
	
    tWdth = 900
	
	
	call tableDiv(tTop,tLeft,tWdth)

    
      

    if media <> "print" AND len(trim(strSQLmids)) > 0 then 'Hvis man er level 1 eller teamleder vil len(trim(strSQLmids)) ALTID VÆRE > 16
      %>
    
	<form id="filterkri" method="post" action="logindhist_2011.asp">
        <input type="hidden" name="FM_sesMid" id="FM_sesMid" value="<%=session("mid") %>">
        <input type="hidden" name="medarbsel_form" id="medarbsel_form" value="1">
        <input type="hidden" name="nomenu" id="nomenu" value="<%=nomenu %>">
        <input type="hidden" name="varTjDatoUS_man" id="varTjDatoUS_man" value="<%=varTjDatoUS_man %>">
        
        <table>
       <tr bgcolor="#ffffff">
	<td valign=top> <b><%=tsa_txt_077 %>:</b> <br />
    <input type="CHECKBOX" name="FM_visallemedarb" id="FM_visallemedarb" value="1" <%=visAlleMedarbCHK %> /> <%=tsa_txt_388 %> (<%=tsa_txt_357 %>)
   
	<br />
                <% 
				call medarb_vaelgandre
                %>
        </td>
           </tr></table>
        </form>
      
	
	<%end if 'media
        %>

   
  

    <table cellpadding=0 cellspacing=0 border=0 width=100%>
     
    <%if media <> "print" then %>

       <%
       prevWeek = datepart("ww", prev_varTjDatoUS_man, 2,2) 
       nextWeek = datepart("ww", next_varTjDatoUS_man, 2,2) 
       %>

         <tr>
          <td><h3><%=tsa_txt_340 %></h3></td>
          <td align=right valign=top style="padding:10px 10px 0px 0px;">
	    <table cellpadding=0 cellspacing=0 border=0 width=120>
	    <tr>
	    <td valign=top align=right><a href="logindhist_2011.asp?usemrn=<%=usemrn%>&varTjDatoUS_man=<%=prev_varTjDatoUS_man%>&varTjDatoUS_son=<%=prev_varTjDatoUS_son %>&rdir=<%=rdir %>&nomenu=<%=nomenu %>">< <%=tsa_txt_005 & " " %> <%=prevWeek %></a></td>
       <td style="padding-left:20px;" valign=top align=right><a href="logindhist_2011.asp?usemrn=<%=usemrn%>&varTjDatoUS_man=<%=next_varTjDatoUS_man %>&varTjDatoUS_son=<%=next_varTjDatoUS_son %>&rdir=<%=rdir %>&nomenu=<%=nomenu %>"> <%=tsa_txt_005 & " " %>  <%=nextWeek %> ></a></td>
	    </tr>
	    </table>

              </td>

      </tr>
        <%if datePart("d", now, 2,2) < 4 OR datePart("d", now, 2,2) > 29 then %>
        <tr>
            <td valign=top style="padding:10px 10px 0px 0px;">
                   
                     <div style="width:600px;"><%
                    
                    if cint(smilaktiv) <> 0 AND media <> "print" then
                    call afslutMsgTxt 
            end if %></div>
                <br />&nbsp;
            </td>

        </tr>
        <%end if %>

    <%end if %>
	
        
          
    <tr><td colspan="2" style="padding-top:10px;">
        <%
            call smileyAfslutSettings()
            'Skal det være mandag for måned og søndag for uge??
            select case cint(SmiWeekOrMonth) 
            case 0 
            periodeTxt = tsa_txt_005 & " "& datepart("ww", tjkDato, 2 ,2)
            weekMonthDate = datepart("ww", varTjDatoUS_son,2,2)
            case 1
            periodeTxt = monthname(datepart("m", tjkDato, 2 ,2))
            weekMonthDate = datepart("m", varTjDatoUS_son,2,2)
            case 2
            weekMonthDate = formatdatetime(now, 2) ' SKAL DET ALTID VÆRE DD?
            weekMonthDate = year(weekMonthDate) & "-" & month(weekMonthDate) & "-" & day(weekMonthDate)
            periodeTxt = weekdayname(weekday(weekMonthDate, 2), 0, 2)
            end select




        if cint(smilaktiv) <> 0 AND media <> "print" then

         'call medrabSmilord(usemrn) '** Sættes på virksomhedsniveau
         call smiley_agg_fn()
         call meStamdata(usemrn)

        call smileyAfslutSettings()

          '** Henter timer i den uge der skal afsluttes ***'
          call afsluttedeUger(year(now), usemrn, 0)

        '** Er kriterie for afslutuge mødt? Ifht. medatype mindstimer og der må ikke være herreløse timer fra. f.eks TT
         call timeKriOpfyldt(lto, sidsteUgenrAfsl, meType, usemrn, SmiWeekOrMonth)

         timerdenneuge_dothis = 1
         call timerDenneUge(usemrn, lto, tjkTimeriUgeSQL, akttypeKrism, timerdenneuge_dothis, SmiWeekOrMonth)
     
       
         '** Periode aflsutning Dag / UGE / MD
         select case cint(SmiWeekOrMonth) 
         case 0
         weekMonthDate = datepart("ww", varTjDatoUS_son,2,2)
         case 1
         weekMonthDate = datepart("m", varTjDatoUS_son,2,2)
         case 2
         weekMonthDate = formatdatetime(now, 2) ' SKAL DET ALTID VÆRE DD?
         weekMonthDate = year(weekMonthDate) & "-" & month(weekMonthDate) & "-" & day(weekMonthDate)
         end select

         'Response.Write "Hej der<br>"

         call erugeAfslutte(datepart("yyyy", varTjDatoUS_son,2,2), weekMonthDate, usemrn, SmiWeekOrMonth, 0) 

         
         '** Faneblade med afslutnings status
        select case cint(SmiWeekOrMonth) 
        case 0, 1
        call smileyAfslutBtn(SmiWeekOrMonth)

        call ugeAfsluttetStatus(varTjDatoUS_son, showAfsuge, ugegodkendt, ugegodkendtaf, SmiWeekOrMonth) 
        end select
         
        


            
	     '**** Afslut uge ***'
	    '**** Smiley vises hvis sidste uge ikke er afsluttet, og dag for afslutninger ovwerskreddet ***' 12 
        if cint(SmiWeekOrMonth) = 0 then
        denneUgeDag = datePart("w", now, 2,2)
        s0Show_sidstedagsidsteuge = dateAdd("d", -denneUgeDag, now) 'now
        else
        s0Show_sidstedagsidsteuge = now
        end if

        '** finder kriterie for rettidig afslutning
        call rettidigafsl(s0Show_sidstedagsidsteuge, 0)

        if cint(SmiWeekOrMonth) = 0 then
            s0Show_weekMd = datePart("ww", s0Show_sidstedagsidsteuge, 2,2)
        else
            s0Show_weekMd = datePart("m", s0Show_sidstedagsidsteuge, 2,2)
        end if

        
         '** HVORFOR DENNE IGEN?
        '** tjekker om uge er afsluttet og viser afsluttet eller form til afslutning
		call erugeAfslutte(year(s0Show_sidstedagsidsteuge), s0Show_weekMd, usemrn, SmiWeekOrMonth, 0)
      
      
      
            
             
        if (cDate(formatdatetime(now, 2)) >= cDate(formatdatetime(cDateUge, 2)) AND cint(ugeNrStatus) = 0) OR cint(smiley_agg) = 1 then

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
    	
    	

          
              

    	    
            '*** Auto popup ThhisWEEK now SMILEY
	      %>
	        <div id="s0" style="position:relative; left:0px; top:5px; width:725px; visibility:<%=smVzb%>; display:<%=smDsp%>; z-index:2; background-color:#FFFFFF; padding:20px; border:0px #CCCCCC solid;">
	       <%
           '*** Viser sidste uge
            'weekSelected = tjekdag(7)

            '*** Viser denne uge
            'weekSelectedThis = dateAdd("d", 7, now) 
            'call showsmiley(weekSelectedThis, 1, usemrn, SmiWeekOrMonth)

            call afslutkri(varTjDatoUS_son, tjkTimeriUgeDt, usemrn, lto, SmiWeekOrMonth)

            call afslutuge(weekSelectedTjk, 1, varTjDatoUS_son, "logindhist", SmiWeekOrMonth)

            '** Status på antal registrederede projekttimer i den pågældende uge    
            'call smiley_statusTxt

        
	          
            '********************************************************************     
            '*** Viser allerede afsluttede peridoer 
            '********************************************************************
            if cint(SmiWeekOrMonth) <> 2 then 'Afslut på dag. Giver ikke mening at vise. Uoverskueligt %>
            <br /><br />
            <span id="se_uegeafls_a" style="color:#5582d2;">[+] <%=tsa_txt_402 & " "& year(varTjDatoUS_son)%></span><br /><%

            varTjDatoUS_ons = dateAdd("d", -3, varTjDatoUS_son)

	        useYear = year(varTjDatoUS_ons)
            '** Hvilke uger er afsluttet '***
	        call smileystatus(usemrn, 1, useYear)
	        
                %> <br />&nbsp;<%

             end if
                
                
             %>
              
	       

               
	        </div>
        
        <br /><br />
        <%end if %>

        <%if cint(smilaktiv) <> 0 AND media = "print" then 

           
          call erugeAfslutte(datepart("yyyy", varTjDatoUS_son,2,2) , weekMonthDate, usemrn, SmiWeekOrMonth, 0) 

          call ugeAfsluttetStatus(varTjDatoUS_son, showAfsuge, ugegodkendt, ugegodkendtaf, SmiWeekOrMonth) 

           

        end if %>
         <br /><br />
        &nbsp;
        </td></tr>

    
	<tr>
	<td valign=top colspan=2 style="padding:10px;">
	
	
	<%call stempelurlist(useMrn, 0, 1, varTjDatoUS_man, varTjDatoUS_son, 0, d_end, lnk) %>
	
	</td></tr>

         <tr>
     <td valign=top colspan=2 style="padding:10px;">
     
	<%call fLonTimerPer(varTjDatoUS_man, 7, 0, useMrn) %>
	
	
	
	</td>
	</tr>
     

    </table>
    
 
 
    
  

   <%
	
	
	
	 if media <> "print" then

ptop = 0
pleft = 945
pwdt = 220

call eksportogprint(ptop,pleft, pwdt)
%>

<form action="logindhist_2011.asp?usemrn=<%=usemrn%>&varTjDatoUS_man=<%=varTjDatoUS_man %>&varTjDatoUS_son=<%=varTjDatoUS_son %>&media=print" method="post" target="_blank">
<tr> 
    <td valign=top align=center>
   <input type=image src="../ill/printer3.png" />
    </td>
    <td class=lille><input id="Submit5" type="submit" value=" Print venlig >> " style="font-size:9px; width:130px;" /></td>
</tr>
</form>

        <tr>
            <td colspan="2" style="padding-top:40px;">

                 <%
    
        if (level <=2 OR level = 6) AND cint(smilaktiv) <> 0 then 

        sm_aar = year(varTjDatoUS_son)
        sm_mid = usemrn

        'response.write "weekMonthDate: "& weekMonthDate

        call erugeAfslutte(sm_aar, weekMonthDate, sm_mid, SmiWeekOrMonth, 0)
                     
        fmlink = "ugeseddel_2011.asp?usemrn="& usemrn &"&varTjDatoUS_man="& varTjDatoUS_man &"&varTjDatoUS_son= "& varTjDatoUS_son &"&nomenu="& nomenu &"&rdir=logindhist"
        %>
           
        <%call godkendugeseddel(fmlink, usemrn, varTjDatoUS_man, rdir) %>
        
       <%end if 'level %>
            </td>

        </tr>
   
	
   </table>
</div>
<%else

Response.Write("<script language=""JavaScript"">window.print();</script>")

end if %>
  

    </div><!-- table div>-->
    <br /><br />&nbsp;
    </div><!-- side div>-->
    <br /><br />&nbsp;
    <%
	
	
	end select
	
	
	end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
