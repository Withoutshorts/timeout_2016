				<div id="kom_son_<%=iRowLoop%>" name="kom_son_<%=iRowLoop%>"  Style="position: absolute; display: none; left:140; top:-70; z-index:2994; width:500; height:140; background-color:#D6DFF5; padding-top:3; padding-left:35; !border: 1px; border-color: #000000; border-style: solid;">Kommentarer til <b><%=left(aTable1Values(6, iRowLoop), 16) &" d. "&convertDate(varTjDatoUS_son)%></b><img src="../ill/blank.gif" width="150" height="1" alt="" border="0">(maks. 255)&nbsp;<input type="text" name="antch_son_<%=iRowLoop%>" id="antch_son_<%=iRowLoop%>" size="3" maxlength="4"><br><textarea cols="50" rows="6" id="FM_kom_son_<%=iRowLoop%>" name="FM_kom_son_<%=iRowLoop%>" onKeyup="antalchar('son',<%=iRowLoop%>);"><%=sonKomm%></textarea>&nbsp;<a href="#" onclick="closekomm('son',<%=iRowLoop%>);">Luk&nbsp;</a></div>
				<div id="kom_man_<%=iRowLoop%>" name="kom_man_<%=iRowLoop%>"  Style="position: absolute; display: none; left:140; top:-70; z-index:2995; width:500; height:140; background-color:#D6DFF5; padding-top:3; padding-left:35; !border: 1px; border-color: #000000; border-style: solid;">Kommentarer til <b><%=left(aTable1Values(6, iRowLoop), 16) &" d. "&convertDate(varTjDatoUS_man)%></b><img src="../ill/blank.gif" width="150" height="1" alt="" border="0">(maks. 255)&nbsp;<input type="text" name="antch_man_<%=iRowLoop%>" id="antch_man_<%=iRowLoop%>" size="3" maxlength="4"><br><textarea cols="50" rows="6" id="FM_kom_man_<%=iRowLoop%>" name="FM_kom_man_<%=iRowLoop%>" onKeyup="antalchar('man',<%=iRowLoop%>);"><%=manKomm%></textarea>&nbsp;<a href="#" onclick="closekomm('man',<%=iRowLoop%>);">Luk&nbsp;</a></div>
				<div id="kom_tir_<%=iRowLoop%>" name="kom_tir_<%=iRowLoop%>"  Style="position: absolute; display: none; left:140; top:-70; z-index:2996; width:500; height:140; background-color:#D6DFF5; padding-top:3; padding-left:35; !border: 1px; border-color: #000000; border-style: solid;">Kommentarer til <b><%=left(aTable1Values(6, iRowLoop), 16) &" d. "&convertDate(varTjDatoUS_tir)%></b><img src="../ill/blank.gif" width="150" height="1" alt="" border="0">(maks. 255)&nbsp;<input type="text" name="antch_tir_<%=iRowLoop%>" id="antch_tir_<%=iRowLoop%>" size="3" maxlength="4"><br><textarea cols="50" rows="6" id="FM_kom_tir_<%=iRowLoop%>" name="FM_kom_tir_<%=iRowLoop%>" onKeyup="antalchar('tir',<%=iRowLoop%>);"><%=tirKomm%></textarea>&nbsp;<a href="#" onclick="closekomm('tir',<%=iRowLoop%>);">Luk&nbsp;</a></div>
				<div id="kom_ons_<%=iRowLoop%>" name="kom_ons_<%=iRowLoop%>"  Style="position: absolute; display: none; left:140; top:-70; z-index:2997; width:500; height:140; background-color:#D6DFF5; padding-top:3; padding-left:35; !border: 1px; border-color: #000000; border-style: solid;">Kommentarer til <b><%=left(aTable1Values(6, iRowLoop), 16) &" d. "&convertDate(varTjDatoUS_ons)%></b><img src="../ill/blank.gif" width="150" height="1" alt="" border="0">(maks. 255)&nbsp;<input type="text" name="antch_ons_<%=iRowLoop%>" id="antch_ons_<%=iRowLoop%>" size="3" maxlength="4"><br><textarea cols="50" rows="6" id="FM_kom_ons_<%=iRowLoop%>" name="FM_kom_ons_<%=iRowLoop%>" onKeyup="antalchar('ons',<%=iRowLoop%>);"><%=onsKomm%></textarea>&nbsp;<a href="#" onclick="closekomm('ons',<%=iRowLoop%>);">Luk&nbsp;</a></div>
				<div id="kom_tor_<%=iRowLoop%>" name="kom_tor_<%=iRowLoop%>"  Style="position: absolute; display: none; left:140; top:-70; z-index:2998; width:500; height:140; background-color:#D6DFF5; padding-top:3; padding-left:35; !border: 1px; border-color: #000000; border-style: solid;">Kommentarer til <b><%=left(aTable1Values(6, iRowLoop), 16) &" d. "&convertDate(varTjDatoUS_tor)%></b><img src="../ill/blank.gif" width="150" height="1" alt="" border="0">(maks. 255)&nbsp;<input type="text" name="antch_tor_<%=iRowLoop%>" id="antch_tor_<%=iRowLoop%>" size="3" maxlength="4"><br><textarea cols="50" rows="6" id="FM_kom_tor_<%=iRowLoop%>" name="FM_kom_tor_<%=iRowLoop%>" onKeyup="antalchar('tor',<%=iRowLoop%>);"><%=torKomm%></textarea>&nbsp;<a href="#" onclick="closekomm('tor',<%=iRowLoop%>);">Luk&nbsp;</a></div>
				<div id="kom_fre_<%=iRowLoop%>" name="kom_fre_<%=iRowLoop%>"  Style="position: absolute; display: none; left:140; top:-70; z-index:2999; width:500; height:140; background-color:#D6DFF5; padding-top:3; padding-left:35; !border: 1px; border-color: #000000; border-style: solid;">Kommentarer til <b><%=left(aTable1Values(6, iRowLoop), 16) &" d. "&convertDate(varTjDatoUS_fre)%></b><img src="../ill/blank.gif" width="150" height="1" alt="" border="0">(maks. 255)&nbsp;<input type="text" name="antch_fre_<%=iRowLoop%>" id="antch_fre_<%=iRowLoop%>" size="3" maxlength="4"><br><textarea cols="50" rows="6" id="FM_kom_fre_<%=iRowLoop%>" name="FM_kom_fre_<%=iRowLoop%>" onKeyup="antalchar('fre',<%=iRowLoop%>);"><%=freKomm%></textarea>&nbsp;<a href="#" onclick="closekomm('fre',<%=iRowLoop%>);">Luk&nbsp;</a></div>
				<div id="kom_lor_<%=iRowLoop%>" name="kom_lor_<%=iRowLoop%>"  Style="position: absolute; display: none; left:140; top:-70; z-index:3000; width:500; height:140; background-color:#D6DFF5; padding-top:3; padding-left:35; !border: 1px; border-color: #000000; border-style: solid;">Kommentarer til <b><%=left(aTable1Values(6, iRowLoop), 16) &" d. "&convertDate(varTjDatoUS_lor)%></b><img src="../ill/blank.gif" width="150" height="1" alt="" border="0">(maks. 255)&nbsp;<input type="text" name="antch_lor_<%=iRowLoop%>" id="antch_lor_<%=iRowLoop%>" size="3" maxlength="4"><br><textarea cols="50" rows="6" id="FM_kom_lor_<%=iRowLoop%>" name="FM_kom_lor_<%=iRowLoop%>" onKeyup="antalchar('lor',<%=iRowLoop%>);"><%=lorKomm%></textarea>&nbsp;<a href="#" onclick="closekomm('lor',<%=iRowLoop%>);">Luk&nbsp;</a></div>
				<%
					Response.write "<tr><td style='width:30px; background-color: #5582D2; padding-left : 4px; border-left: 1px solid #003399'><font class='lillehvid'><b>"_
					& aTable1Values(1,iRowLoop) &"</b></td><td height=20 style='background-color: #5582D2; padding-left : 4px;padding-right : 4px; width:380px;'>"
					Response.write "<font class='storhvid'>"
					'** Er der nogen aktiviteter **
					if len(aTable1Values(5,iRowLoop)) <> 0 then%>
							<a href="javascript:expand('<%=iRowLoop%>');"><img src="ill/plus.gif" width="9" height="9" border="0" name="Menub<%=iRowLoop%>"></a>&nbsp;
							<%if level <= 3 OR level = 6 then%>
							<a href="jobs.asp?menu=job&func=red&id=<%=aTable1Values(0,iRowLoop)%>&int=1&rdir=treg" class="alt">
							<%else%>
							<a href="javascript:expand('<%=iRowLoop%>');" class=alt>
							<%end if%>
							<%
							'* finder antal brugte timer på job ***
							strSQL2 = "SELECT sum(timer) AS sumtimer FROM timer WHERE tjobnr = '"& aTable1Values(1,iRowLoop) &"' AND tfaktim <> 5 ORDER BY timer"
							oRec2.open strSQL2, oConn, 3
							if not oRec2.EOF then 
							timerbrugtthis = oRec2("sumtimer")
							end if
							oRec2.close
							
							if len(timerbrugtthis ) > 0 then
							timerbrugtthis = timerbrugtthis 
							else
							timerbrugtthis = 0
							end if
							
							Response.write aTable1Values(2,iRowLoop) &"</a></b>&nbsp;&nbsp;&nbsp;<font class='lilleblaa'>("&aTable1Values(4,iRowLoop)&")</font></td><td align=right width=120 style='padding-right:3px; background-color: #5582D2; border-right: 1px solid #003399;'><font class='lillehvid'><b>"&formatnumber(aTable1Values(17,iRowLoop) + aTable1Values(18,iRowLoop), 2)&" / <font class='lilleblaa'>"& formatnumber(timerbrugtthis, 2) &"</b>&nbsp;&nbsp;&nbsp;<input type='checkbox' name='FM_flyttilguide' value='"& aTable1Values(0,iRowLoop) &"'></td></tr>"
							%>
							<tr style="border-left: 1px solid #003399; border-right: 1px solid #003399">
							<td colspan="3"><div ID="Menu<%=iRowLoop%>" Style="position: relative; display: none;">
							<table cellspacing="0" cellpadding="0" border="0" width=100% bgcolor="#ffffff">
							
					<%else '*** (ingen aktiviteter) ***
							%>
							<img src='ill/blank.gif' width='18' height='0' border='0' alt=''>&nbsp;<a href="jobs.asp?menu=job&func=red&id=<%=aTable1Values(0,iRowLoop)%>&int=1&rdir=treg" class="alt">
							<%
							Response.write "<b>"&aTable1Values(2,iRowLoop)&"</b></a>&nbsp;&nbsp;&nbsp;<font class='lilleblaa'>("&aTable1Values(4,iRowLoop)&")</font></td><td align=right width=120 style='padding-right:3px; background-color: #5582D2; border-right: 1px solid #003399;'>&nbsp;<input type='checkbox' name='FM_flyttilguide' value='"& aTable1Values(0,iRowLoop) &"'></td></tr>"
							%>
							<tr style=" border-left: 1px solid #003399; border-right: 1px solid #003399">
							<td colspan="3"><div ID="Menu_<%=aTable1Values(2,iRowLoop)%>" style="position: relative; display: none;">
							<table cellspacing="0" cellpadding="0" border="0" width="100%">
							<%
					end if