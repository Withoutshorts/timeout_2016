

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/treg_func.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->



<div class="wrapper">
    <div class="content">

        <%
            if len(session("user")) = 0 then

	        errortype = 5
	        call showError(errortype)
            response.End
            end if

            call menu_2014
        

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
	        
            <SCRIPT language=javascript src="inc/logindhist_2011_jav.js"></script>
	        <%end if%>
	
	        <%end if%>


        <!-------------------------------Sideindhold------------------------------------->
        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u>Komme / Gå tid</u></h3>
                <div class="portlet-body">

                    <%
                        call erTeamLederForvlgtMedarb(level, usemrn, session("mid"))
	
	                    call ersmileyaktiv()

                        if media <> "print" then
                        'tTop = 0
	                    'tLeft = 0
                        else
                        'tTop = 0
	                    'tLeft = 0
                        end if
	
                        'tWdth = 900
	
	
	                    'call tableDiv(tTop,tLeft,tWdth)

                    if media <> "print" AND len(trim(strSQLmids)) > 0 then 'Hvis man er level 1 eller teamleder vil len(trim(strSQLmids)) ALTID VÆRE > 16
                    %>
                    
                    <div class="well">
                        <form id="filterkri" method="post" action="logindhist_2011.asp">
                            <input type="hidden" name="FM_sesMid" id="FM_sesMid" value="<%=session("mid") %>">
                            <input type="hidden" name="medarbsel_form" id="medarbsel_form" value="1">
                            <input type="hidden" name="nomenu" id="nomenu" value="<%=nomenu %>">
                            <input type="hidden" name="varTjDatoUS_man" id="varTjDatoUS_man" value="<%=varTjDatoUS_man %>">

                            <div class="row">
                                <div class="col-lg-4"><%=tsa_txt_077 %> <input type="CHECKBOX" name="FM_visallemedarb" id="FM_visallemedarb" value="1" <%=visAlleMedarbCHK %> /> <%=tsa_txt_388 %> (<%=tsa_txt_357 %>)</div>
                            </div>
                            <div class="row">
                                <div class="col-lg-4">
                                    <% 
				                    call medarb_vaelgandre
                                    %>
                                </div>
                            </div>

                        </form>
                    </div>


                    <%end if 'media %>

                   <%

                   if media <> "print" then


                       prevWeek = datepart("ww", prev_varTjDatoUS_man, 2,2) 
                       nextWeek = datepart("ww", next_varTjDatoUS_man, 2,2) 
                       %>

                        <div class="row">
                            <h6 class="col-lg-1"><a href="logindhist_2011.asp?usemrn=<%=usemrn%>&varTjDatoUS_man=<%=prev_varTjDatoUS_man%>&varTjDatoUS_son=<%=prev_varTjDatoUS_son %>&rdir=<%=rdir %>&nomenu=<%=nomenu %>">< uge <%=prevWeek %></a></h6>
                            <h6 class="col-lg-1"><a href="logindhist_2011.asp?usemrn=<%=usemrn%>&varTjDatoUS_man=<%=next_varTjDatoUS_man %>&varTjDatoUS_son=<%=next_varTjDatoUS_son %>&rdir=<%=rdir %>&nomenu=<%=nomenu %>">uge <%=nextWeek %> ></a></h6>
                        </div>

                        <%if datePart("d", now, 2,2) < 4 OR datePart("d", now, 2,2) > 29 then %>
                        <div class="row">
                            <div class="col-lg-4">
                            
                                <div style="width:600px;"><%                   
                                if cint(smilaktiv) <> 0 AND media <> "print" then
                                call afslutMsgTxt 
                                end if %>
                                </div>
                            
                            </div>
                        </div>
                        <%end if %>

                    <%end if %>

                    <table class="table">
                        
                    </table>

                </div>
            </div>
        </div>




        <%end select %>

    </div>
</div>    

<!--#include file="../inc/regular/footer_inc.asp"-->