<%  
   Response.ContentType = "text/x-component"
   dim Culture
   Culture = Trim(Request.QueryString("UC"))  
   dim FilePath
   FilePath = Trim(Request.QueryString("F"))  
   Public Function GetString(instring)
	    dim t
    	
	    t = GetStringByCulture(instring,Culture)
    	
	    If t = ""  then
		    t= GetStringByCulture(instring,"_default")
	    End If
    	
	    If t = ""  then
		    t= "{"&instring&"}"	
	    End If
	    GetString= t	
    End Function
    
    
   FilePath = Left(request.ServerVariables("PATH_INFO"),InStr(request.ServerVariables("PATH_INFO"),FilePath)+len(FilePath)-1)

    Public Function GetStringByCulture(instring,input_culture)
	    dim scriptname,xmlfilename,doc,temp
	    dim node,selectednode,optionnodelist,errobj
	    dim selectednodes

	    xmlfilename= Server.MapPath(FilePath&"/languages/"&input_culture&".xml")

	    ' Create an object to hold the XML
	    set doc = server.CreateObject("Microsoft.XMLDOM")

	    ' For ASP, wait until the XML is all ready before continuing
	    doc.async = False

	    ' Load the XML file or return an error message and stop the script
	    if not Doc.Load(xmlfilename) then
		    Response.Write "Failed to load the language text from the XML file.<br>"
		    Response.End
	    end if

	    ' Make sure that the interpreter knows that we are using XPath as our selection language
	    doc.setProperty "SelectionLanguage", "XPath"

	    set selectednode= doc.selectSingleNode("/resources/resource[@name='"&instring&"']")
	    if IsObject(selectednode) and not selectednode is nothing  then
		    GetStringByCulture=CuteEditor_Decode(selectednode.text)
	    else
		    GetStringByCulture=""		
	    end if
    End Function    
    
     PUBLIC FUNCTION CuteEditor_Decode(input_str)        
	    input_str=Replace(input_str,"#1","<")
	    input_str=Replace(input_str,"#2",">")
	    input_str=Replace(input_str,"#3","&")
	    input_str=Replace(input_str,"#4","*")
	    input_str=Replace(input_str,"#5","o")
	    input_str=Replace(input_str,"#6","O")
	    input_str=Replace(input_str,"#7","s")
	    input_str=Replace(input_str,"#8","S")
	    input_str=Replace(input_str,"#9","e")
	    input_str=Replace(input_str,"#a","E")
	    input_str=Replace(input_str,"#0","#")
	    CuteEditor_Decode = input_str
     END FUNCTION
%>
<PUBLIC:COMPONENT>
	<PUBLIC:EVENT ID="event_oncolorchange" name="oncolorchange" />
	<PUBLIC:EVENT ID="event_oncolorpopup" name="oncolorpopup" />
	<PUBLIC:PROPERTY name="selectedColor" GET="_get_selectedColor" PUT="_set_selectedColor"/>
	<PUBLIC:METHOD name="popupColor" INTERNALNAME="_mtd_popupColor" />
	<PUBLIC:ATTACH EVENT="onclick" ONEVENT="_mtd_onclick()" />
</PUBLIC:COMPONENT>

<script type="text/javascript">
var OxO9b5d=["#000000","#993300","#333300","#003300","#003366","#000080","#333399","#333333","#800000","#FF6600","#808000","#008000","#008080","#0000FF","#666699","#808080","#FF0000","#FF9900","#99CC00","#339966","#33CCCC","#3366FF","#800080","#999999","#FF00FF","#FFCC00","#FFFF00","#00FF00","#00FFFF","#00CCFF","#993366","#C0C0C0","#FF99CC","#FFCC99","#FFFF99","#CCFFCC","#CCFFFF","#99CCFF","#CC99FF","#FFFFFF","dialogWidth:500px;dialogHeight:330px;help:0;status:0;resizable:1","disableVisual","","\x3CDIV style=\x27width=140;cursor:default;position:absolute;z-index:88888888;background-color:white;border:0px;overflow:visible;\x27\x3E","length","\x3Ctable cellpadding=0 cellspacing=5 style=\x27width:100%;font-family: Verdana; font-size: 6px; BORDER: #666666 1px solid;\x27 bgcolor=#f9f8f7\x3E\x3Ctr\x3E\x3Ctd\x3E","\x3Ctable cellpadding=0 cellspacing=2 style=\x27font-size: 3px;\x27\x3E","\x3Ctr\x3E","\x3Ctd colspan=10 align=center style=\x22padding:1px;border:solid 1px #f9f8f7;margin:1px\x22 onmouseup=\x22document.all.","uniqueID","._cphtc_sel(this.ColorValue)\x22  ColorValue=\x22\x22 onmouseover=\x22CuteEditor_ColorPicker_ButtonOver(this);\x22\x3E","\x3Ctable cellspacing=0 cellpadding=0 border=0 width=90% style=\x22font-size:3px\x22\x3E","\x3C/table\x3E","\x3C/td\x3E","\x3C/tr\x3E","\x3Ctr\x3E\x3Ctd\x3E\x26nbsp;\x3C/td\x3E\x3C/tr\x3E","\x3Ctd title="," align=center style=\x22padding:1px;border:solid 1px #f9f8f7;\x22 onmouseover=\x22CuteEditor_ColorPicker_ButtonOver(this);\x22 ColorValue=\x22","\x22 onmouseup=\x22document.all.","._cphtc_sel(this.ColorValue);\x22\x3E","\x3Cdiv style=\x22background-color:","; border:solid 1px #808080; width:12px; height:12px; font-size: 3px;\x22\x3E\x26nbsp;\x3C/div\x3E","\x3C/td\x3E\x3C/tr\x3E","\x3Ctd colspan=10 align=center style=\x22padding:1px;border:solid 1px #f9f8f7;\x22 onmouseover=\x22CuteEditor_ColorPicker_ButtonOver(this);\x22 onmouseup=\x22document.all.","._cphtc_dlg();\x22\x3E","innerHTML","body","document","onclick","SELECT","all","visibility","currentStyle","hidden","runtimeStyle","style","_visibility","top","left","display","block","offsetHeight","px","unselectable","on","none"];var colorsArray= new Array(OxO9b5d[0],OxO9b5d[1],OxO9b5d[2],OxO9b5d[3],OxO9b5d[4],OxO9b5d[5],OxO9b5d[6],OxO9b5d[7],OxO9b5d[8],OxO9b5d[9],OxO9b5d[10],OxO9b5d[11],OxO9b5d[12],OxO9b5d[13],OxO9b5d[14],OxO9b5d[15],OxO9b5d[16],OxO9b5d[17],OxO9b5d[18],OxO9b5d[19],OxO9b5d[20],OxO9b5d[21],OxO9b5d[22],OxO9b5d[23],OxO9b5d[24],OxO9b5d[25],OxO9b5d[26],OxO9b5d[27],OxO9b5d[28],OxO9b5d[29],OxO9b5d[30],OxO9b5d[31],OxO9b5d[32],OxO9b5d[33],OxO9b5d[34],OxO9b5d[35],OxO9b5d[36],OxO9b5d[37],OxO9b5d[38],OxO9b5d[39]);var ShowMoreColors=true;var dlgurl='<%=FilePath%>/Dialogs/ColorPicker.Asp?<%=Request.ServerVariables("QUERY_STRING") %>&setting=<%= Request.Cookies("CESecurity") %>';function element._cphtc_sel(Ox4){_color=Ox4;event_oncolorchange.fire();} ;function element._cphtc_dlg(){CloseDiv();event_oncolorpopup.fire();var Ox5=OxO9b5d[40];if(element[OxO9b5d[41]]){var Ox6=showModalDialog(dlgurl,{color:Ox7},Ox5);if(Ox6!=null&&Ox6!=false){_color=Ox6;event_oncolorchange.fire();} ;} else {var Ox7=_color;var Ox6=showModalDialog(dlgurl,{color:Ox7,onchange:Ox8},Ox5);if(Ox6!=null&&Ox6!=false){_color=Ox6;} else {_color=Ox7;} ;event_oncolorchange.fire();function Ox8(Ox9){_color=Ox9;event_oncolorchange.fire();} ;} ;} ;var _color=OxO9b5d[42];function _get_selectedColor(){return _color;} ;function _set_selectedColor(Ox9){_color=Ox9;} ;var div;var selects;var isopen=false;function _mtd_onclick(){_mtd_popupColor();} ;function _mtd_popupColor(){if(div==null){div=document.createElement(OxO9b5d[43]);var Ox12=OxO9b5d[42];var Ox13=colorsArray[OxO9b5d[44]];var Ox14=8;Ox12+=OxO9b5d[45];Ox12+=OxO9b5d[46];Ox12+=OxO9b5d[47];Ox12+=OxO9b5d[48]+element[OxO9b5d[49]]+OxO9b5d[50];Ox12+=OxO9b5d[51];Ox12+='<tr><td width=18><div style="background-color:#000000; border:solid 1px #808080; width:12px; height:12px; font-size: 3px;">&nbsp;</div></td><td align=center style="font:normal 11px verdana;">&nbsp;<%= GetString("Automatic") %></td></tr>';Ox12+=OxO9b5d[52];Ox12+=OxO9b5d[53];Ox12+=OxO9b5d[54];Ox12+=OxO9b5d[55];for(var i=0;i<Ox13;i++){if((i%Ox14)==0){Ox12+=OxO9b5d[47];} ;Ox12+=OxO9b5d[56]+colorsArray[i]+OxO9b5d[57]+colorsArray[i]+OxO9b5d[58]+element[OxO9b5d[49]]+OxO9b5d[59];Ox12+=OxO9b5d[60]+colorsArray[i]+OxO9b5d[61];Ox12+=OxO9b5d[53];if(((i+1)>=Ox13)||(((i+1)%Ox14)==0)){Ox12+=OxO9b5d[54];} ;} ;Ox12+=OxO9b5d[55];Ox12+=OxO9b5d[52];Ox12+=OxO9b5d[62];if(ShowMoreColors){Ox12+=OxO9b5d[47];Ox12+=OxO9b5d[63]+element[OxO9b5d[49]]+OxO9b5d[64];Ox12+=OxO9b5d[51];Ox12+='<tr><td width=18><div style="background-color:#000000; border:solid 1px #808080; width:12px; height:12px;font-size: 3px;"></div></td><td align=center style="font-size:11px"><%= GetString("MoreColors") %></td></tr>';Ox12+=OxO9b5d[52];Ox12+=OxO9b5d[53];Ox12+=OxO9b5d[54];} ;Ox12+=OxO9b5d[52];div[OxO9b5d[65]]=Ox12;element[OxO9b5d[67]][OxO9b5d[66]].appendChild(div);div[OxO9b5d[68]]=CloseDiv;} ;if(isopen){CloseDiv();} ;isopen=true;selects=[];var Ox16=element[OxO9b5d[67]][OxO9b5d[70]].tags(OxO9b5d[69]);for(var i=0;i<Ox16[OxO9b5d[44]];i++){var Ox17=Ox16[i];if(Ox17[OxO9b5d[72]][OxO9b5d[71]]!=OxO9b5d[73]){selects[selects[OxO9b5d[44]]]=Ox17;var Ox18=Ox17[OxO9b5d[74]]||Ox17[OxO9b5d[75]];Ox18[OxO9b5d[76]]=Ox18[OxO9b5d[71]];Ox18[OxO9b5d[71]]=OxO9b5d[73];} ;} ;div[OxO9b5d[75]][OxO9b5d[77]]=0;div[OxO9b5d[75]][OxO9b5d[78]]=0;div[OxO9b5d[75]][OxO9b5d[79]]=OxO9b5d[80];var Ox19=CalcPosition(div,element);Ox19[OxO9b5d[77]]+=element[OxO9b5d[81]];AdjustMirror(div,element,Ox19);div[OxO9b5d[75]][OxO9b5d[78]]=Ox19[OxO9b5d[78]]+OxO9b5d[82];div[OxO9b5d[75]][OxO9b5d[77]]=Ox19[OxO9b5d[77]]+OxO9b5d[82];var Ox16=div[OxO9b5d[70]];for(var i=0;i<Ox16[OxO9b5d[44]];i++){Ox16[i][OxO9b5d[83]]=OxO9b5d[84];} ;div.focus();var Ox1a= new CaptureManager(element,handlelosecapture);Ox1a.AddElement(div);} ;function handlelosecapture(){CloseDiv();} ;function CloseDiv(){CaptureManager.Unregister(element);isopen=false;div[OxO9b5d[75]][OxO9b5d[79]]=OxO9b5d[85];for(var i=0;i<selects[OxO9b5d[44]];i++){var Ox17=selects[i];Ox17[OxO9b5d[74]][OxO9b5d[71]]=Ox17[OxO9b5d[74]][OxO9b5d[76]];} ;} ;

</script>

<script type="text/javascript">


var OxO5d83=["body","document","compatMode","CSS1Compat","documentElement","scrollLeft","scrollTop","clientLeft","clientTop","parentElement","position","currentStyle","absolute","relative","left","top","clientWidth","clientHeight","offsetWidth","offsetHeight","element","capturemanager","\x3CDIV style=\x27width:0px;height:0px;left:0px;top:0px;position:absolute\x27\x3E","afterBegin","onlosecapture","onmousedown","onmousemove","onmouseover","onmouseout","length"];function GetScrollPosition(Ox1f){var Ox20=window[OxO5d83[1]][OxO5d83[0]];var p=Ox20;if(window[OxO5d83[1]][OxO5d83[2]]==OxO5d83[3]){p=window[OxO5d83[1]][OxO5d83[4]];} ;if(Ox1f==Ox20){return {left:0,top:0};} ;with(Ox1f.getBoundingClientRect()){return {left:p[OxO5d83[5]]+left,top:p[OxO5d83[6]]+top};} ;} ;function GetClientPosition(Ox1f){var Ox20=window[OxO5d83[1]][OxO5d83[0]];var p=Ox20;if(window[OxO5d83[1]][OxO5d83[2]]==OxO5d83[3]){p=window[OxO5d83[1]][OxO5d83[4]];} ;if(Ox1f==Ox20){return {left:-p[OxO5d83[5]],top:-p[OxO5d83[6]]};} ;with(Ox1f.getBoundingClientRect()){return {left:left-p[OxO5d83[7]],top:top-p[OxO5d83[8]]};} ;} ;function GetStandParent(Ox1f){for(var Ox24=Ox1f[OxO5d83[9]];Ox24!=null;Ox24=Ox24[OxO5d83[9]]){var Ox25=Ox24[OxO5d83[11]][OxO5d83[10]];if(Ox25==OxO5d83[12]||Ox25==OxO5d83[13]){return Ox24;} ;} ;return window[OxO5d83[1]][OxO5d83[0]];} ;function CalcPosition(Ox27,Ox1f){var Ox28=GetScrollPosition(Ox1f);var Ox29=GetScrollPosition(GetStandParent(Ox27));var Ox2a=GetStandParent(Ox27);return {left:Ox28[OxO5d83[14]]-Ox29[OxO5d83[14]]-Ox2a[OxO5d83[7]],top:Ox28[OxO5d83[15]]-Ox29[OxO5d83[15]]-Ox2a[OxO5d83[8]]};} ;function AdjustMirror(Ox27,Ox1f,Ox19){var Ox2c=window[OxO5d83[1]][OxO5d83[0]][OxO5d83[16]];var Ox2d=window[OxO5d83[1]][OxO5d83[0]][OxO5d83[17]];if(window[OxO5d83[1]][OxO5d83[2]]==OxO5d83[3]){Ox2c=window[OxO5d83[1]][OxO5d83[4]][OxO5d83[16]];Ox2d=window[OxO5d83[1]][OxO5d83[4]][OxO5d83[17]];} ;var Ox2e=Ox27[OxO5d83[18]];var Ox2f=Ox27[OxO5d83[19]];var Ox30=GetClientPosition(GetStandParent(Ox27));var Ox31={left:Ox30[OxO5d83[14]]+Ox19[OxO5d83[14]]+Ox2e/2,top:Ox30[OxO5d83[15]]+Ox19[OxO5d83[15]]+Ox2f/2};var Ox32={left:Ox30[OxO5d83[14]]+Ox19[OxO5d83[14]],top:Ox30[OxO5d83[15]]+Ox19[OxO5d83[15]]};if(Ox1f!=null){var Ox33=GetClientPosition(Ox1f);Ox32={left:Ox33[OxO5d83[14]]+Ox1f[OxO5d83[18]]/2,top:Ox33[OxO5d83[15]]+Ox1f[OxO5d83[19]]/2};} ;var Ox34=true;if(Ox31[OxO5d83[14]]-Ox2e/2<0){if((Ox32[OxO5d83[14]]*2-Ox31[OxO5d83[14]])+Ox2e/2<=Ox2c){Ox31[OxO5d83[14]]=Ox32[OxO5d83[14]]*2-Ox31[OxO5d83[14]];} else {if(Ox34){Ox31[OxO5d83[14]]=Ox2e/2+4;} ;} ;} else {if(Ox31[OxO5d83[14]]+Ox2e/2>Ox2c){if((Ox32[OxO5d83[14]]*2-Ox31[OxO5d83[14]])-Ox2e/2>=0){Ox31[OxO5d83[14]]=Ox32[OxO5d83[14]]*2-Ox31[OxO5d83[14]];} else {if(Ox34){Ox31[OxO5d83[14]]=Ox2c-Ox2e/2-4;} ;} ;} ;} ;if(Ox31[OxO5d83[15]]-Ox2f/2<0){if((Ox32[OxO5d83[15]]*2-Ox31[OxO5d83[15]])+Ox2f/2<=Ox2d){Ox31[OxO5d83[15]]=Ox32[OxO5d83[15]]*2-Ox31[OxO5d83[15]];} else {if(Ox34){Ox31[OxO5d83[15]]=Ox2f/2+4;} ;} ;} else {if(Ox31[OxO5d83[15]]+Ox2f/2>Ox2d){if((Ox32[OxO5d83[15]]*2-Ox31[OxO5d83[15]])-Ox2f/2>=0){Ox31[OxO5d83[15]]=Ox32[OxO5d83[15]]*2-Ox31[OxO5d83[15]];} else {if(Ox34){Ox31[OxO5d83[15]]=Ox2d-Ox2f/2-4;} ;} ;} ;} ;Ox19[OxO5d83[14]]=Ox31[OxO5d83[14]]-Ox30[OxO5d83[14]]-Ox2e/2;Ox19[OxO5d83[15]]=Ox31[OxO5d83[15]]-Ox30[OxO5d83[15]]-Ox2f/2;} ;function CaptureManager(element,handlelosecapture){if(CaptureManager[OxO5d83[20]]&&CaptureManager[OxO5d83[20]][OxO5d83[21]]){CaptureManager[OxO5d83[20]][OxO5d83[21]].Unregister();} ;var Ox37=true;var Ox38=[];var Ox39=true;var Ox3a=false;element[OxO5d83[21]]=Ox3f;CaptureManager[OxO5d83[20]]=element;Ox3f.AddElement(element);var Ox3b=element[OxO5d83[1]].createElement(OxO5d83[22]);element[OxO5d83[1]][OxO5d83[0]].insertAdjacentElement(OxO5d83[23],Ox3b);Ox3b.attachEvent(OxO5d83[24],Ox42);Ox3c(Ox3b);Ox3b.setCapture(true);Ox3a=true;return Ox3f;function Ox3c(Ox3d){Ox3d.attachEvent(OxO5d83[25],Ox43);Ox3d.attachEvent(OxO5d83[26],Ox45);Ox3d.attachEvent(OxO5d83[27],Ox47);Ox3d.attachEvent(OxO5d83[28],Ox48);} ;function Ox3e(Ox3d){Ox3d.detachEvent(OxO5d83[25],Ox43);Ox3d.detachEvent(OxO5d83[26],Ox45);Ox3d.detachEvent(OxO5d83[27],Ox47);Ox3d.detachEvent(OxO5d83[28],Ox48);} ;function Ox3f(){} ;function Ox3f.Unregister(){Ox3b.detachEvent(OxO5d83[24],Ox42);Ox3e(Ox3b);Ox3b.removeNode(true);for(var i=0;i<Ox38[OxO5d83[29]];i++){var Ox3d=Ox38[i];Ox3e(Ox3d);} ;Ox37=false;element[OxO5d83[21]]=null;CaptureManager[OxO5d83[20]]=null;if(Ox3a){Ox3a=false;Ox3b.releaseCapture();Ox3f.FireLoseCapture();} ;} ;function Ox3f.AddElement(Ox3d){Ox3c(Ox3d);Ox38[Ox38[OxO5d83[29]]]=Ox3d;} ;function Ox3f.DelElement(Ox3d){var len=Ox38[OxO5d83[29]];for(var i=0;i<len;i++){if(Ox38[i]==Ox3d){Ox3e(Ox3d);for(var Ox41=i;Ox41<len-1;Ox41++){Ox38[Ox41]=Ox38[Ox41+1];} ;Ox38[OxO5d83[29]]=Ox38[OxO5d83[29]]-1;return ;} ;} ;} ;function Ox3f.FireLoseCapture(){handlelosecapture();} ;function Ox42(){if(Ox3a){Ox3a=false;try{Ox3f.FireLoseCapture();} finally{Ox3f.Unregister();} ;} ;} ;function Ox43(){var Ox44=element[OxO5d83[1]].elementFromPoint(event.clientX,event.clientY);for(var i=0;i<Ox38[OxO5d83[29]];i++){var Ox3d=Ox38[i];if(Ox3d.contains(Ox44)){return ;} ;} ;Ox3f.Unregister();} ;function Ox45(){var Ox44=element[OxO5d83[1]].elementFromPoint(event.clientX,event.clientY);Ox46(Ox44);} ;function Ox46(Ox44){for(var i=0;i<Ox38[OxO5d83[29]];i++){var Ox3d=Ox38[i];if(Ox3d.contains(Ox44)){if(Ox3a){Ox3a=false;Ox3b.releaseCapture();} ;return ;} ;} ;if(!Ox3a){Ox3a=true;Ox3b.setCapture(true);} ;} ;function Ox47(){var Ox44=element[OxO5d83[1]].elementFromPoint(event.clientX,event.clientY);Ox39=false;for(var i=0;i<Ox38[OxO5d83[29]];i++){var Ox3d=Ox38[i];if(Ox3d.contains(event.fromElement)){return ;} ;if(Ox3d.contains(Ox44)){if(Ox3a){Ox3a=false;Ox3b.releaseCapture();} ;} ;} ;} ;function Ox48(){var Ox44=element[OxO5d83[1]].elementFromPoint(event.clientX,event.clientY);Ox39=false;for(var i=0;i<Ox38[OxO5d83[29]];i++){var Ox3d=Ox38[i];if(Ox3d.contains(event.toElement)){return ;} ;} ;if(!Ox3a){Ox3a=true;Ox3b.setCapture(true);} ;} ;} ;function CaptureManager.Register(element,handlelosecapture){return  new CaptureManager(element,handlelosecapture);} ;function CaptureManager.Unregister(element){if(element&&element[OxO5d83[21]]){element[OxO5d83[21]].Unregister();} ;} ;
</script>
