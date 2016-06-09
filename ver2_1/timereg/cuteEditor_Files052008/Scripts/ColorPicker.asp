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
var OxOb93c=["#000000","#993300","#333300","#003300","#003366","#000080","#333399","#333333","#800000","#FF6600","#808000","#008000","#008080","#0000FF","#666699","#808080","#FF0000","#FF9900","#99CC00","#339966","#33CCCC","#3366FF","#800080","#999999","#FF00FF","#FFCC00","#FFFF00","#00FF00","#00FFFF","#00CCFF","#993366","#C0C0C0","#FF99CC","#FFCC99","#FFFF99","#CCFFCC","#CCFFFF","#99CCFF","#CC99FF","#FFFFFF","dialogWidth:500px;dialogHeight:330px;help:0;status:0;resizable:1","disableVisual","","\x3CDIV style=\x27width=140;cursor:default;position:absolute;z-index:88888888;background-color:white;border:0px;overflow:visible;\x27\x3E","length","\x3Ctable cellpadding=0 cellspacing=5 style=\x27width:100%;font-family: Verdana; font-size: 6px; BORDER: #666666 1px solid;\x27 bgcolor=#f9f8f7\x3E\x3Ctr\x3E\x3Ctd\x3E","\x3Ctable cellpadding=0 cellspacing=2 style=\x27font-size: 3px;\x27\x3E","\x3Ctr\x3E","uniqueID","\x3Ctd colspan=10 align=center style=\x22padding:1px;border:solid 1px #f9f8f7;margin:1px\x22 onmouseup=\x22document.all.","._cphtc_sel(this.ColorValue)\x22  ColorValue=\x22\x22 onmouseover=\x22CuteEditor_ColorPicker_ButtonOver(this);\x22\x3E","\x3Ctable cellspacing=0 cellpadding=0 border=0 width=90% style=\x22font-size:3px\x22\x3E","\x3C/table\x3E","\x3C/td\x3E","\x3C/tr\x3E","\x3Ctr\x3E\x3Ctd\x3E\x26nbsp;\x3C/td\x3E\x3C/tr\x3E","\x3Ctd title="," align=center style=\x22padding:1px;border:solid 1px #f9f8f7;\x22 onmouseover=\x22CuteEditor_ColorPicker_ButtonOver(this);\x22 ColorValue=\x22","\x22 onmouseup=\x22document.all.","._cphtc_sel(this.ColorValue);\x22\x3E","\x3Cdiv style=\x22background-color:","; border:solid 1px #808080; width:12px; height:12px; font-size: 3px;\x22\x3E\x26nbsp;\x3C/div\x3E","\x3C/td\x3E\x3C/tr\x3E","\x3Ctd colspan=10 align=center style=\x22padding:1px;border:solid 1px #f9f8f7;\x22 onmouseover=\x22CuteEditor_ColorPicker_ButtonOver(this);\x22 onmouseup=\x22document.all.","._cphtc_dlg();\x22\x3E","innerHTML","body","document","onclick","SELECT","all","visibility","currentStyle","hidden","style","runtimeStyle","_visibility","top","left","display","block","offsetHeight","px","unselectable","on","none"];var colorsArray= new Array(OxOb93c[0x0],OxOb93c[0x1],OxOb93c[0x2],OxOb93c[0x3],OxOb93c[0x4],OxOb93c[0x5],OxOb93c[0x6],OxOb93c[0x7],OxOb93c[0x8],OxOb93c[0x9],OxOb93c[0xa],OxOb93c[0xb],OxOb93c[0xc],OxOb93c[0xd],OxOb93c[0xe],OxOb93c[0xf],OxOb93c[0x10],OxOb93c[0x11],OxOb93c[0x12],OxOb93c[0x13],OxOb93c[0x14],OxOb93c[0x15],OxOb93c[0x16],OxOb93c[0x17],OxOb93c[0x18],OxOb93c[0x19],OxOb93c[0x1a],OxOb93c[0x1b],OxOb93c[0x1c],OxOb93c[0x1d],OxOb93c[0x1e],OxOb93c[0x1f],OxOb93c[0x20],OxOb93c[0x21],OxOb93c[0x22],OxOb93c[0x23],OxOb93c[0x24],OxOb93c[0x25],OxOb93c[0x26],OxOb93c[0x27]);var ShowMoreColors=true;var dlgurl='<%=FilePath%>/Dialogs/ColorPicker.Asp?<%=Request.ServerVariables("QUERY_STRING") %>&setting=<%= Request.Cookies("CESecurity") %>'; function element._cphtc_sel(Oxe4){ _color=Oxe4 ; event_oncolorchange.fire() ;}  ; function element._cphtc_dlg(){ CloseDiv() ; event_oncolorpopup.fire() ;var Oxe5=OxOb93c[0x28];if(element[OxOb93c[0x29]]){var Oxe6=showModalDialog(dlgurl,{color:Oxe7},Oxe5);if(Oxe6!=null&&Oxe6!=false){ _color=Oxe6 ; event_oncolorchange.fire() ;} ;} else {var Oxe7=_color;var Oxe6=showModalDialog(dlgurl,{color:Oxe7,onchange:Oxe8},Oxe5);if(Oxe6!=null&&Oxe6!=false){ _color=Oxe6 ;} else { _color=Oxe7 ;} ; event_oncolorchange.fire() ; function Oxe8(Oxe9){ _color=Oxe9 ; event_oncolorchange.fire() ;}  ;} ;}  ;var _color=OxOb93c[0x2a]; function _get_selectedColor(){return _color;}  ; function _set_selectedColor(Oxe9){ _color=Oxe9 ;}  ;var div;var selects;var isopen=false; function _mtd_onclick(){ _mtd_popupColor() ;}  ; function _mtd_popupColor(){if(div==null){ div=document.createElement(OxOb93c[0x2b]) ;var Oxf2=OxOb93c[0x2a];var Oxf3=colorsArray[OxOb93c[0x2c]];var Oxf4=0x8; Oxf2+=OxOb93c[0x2d] ; Oxf2+=OxOb93c[0x2e] ; Oxf2+=OxOb93c[0x2f] ; Oxf2+=OxOb93c[0x31]+element[OxOb93c[0x30]]+OxOb93c[0x32] ; Oxf2+=OxOb93c[0x33] ; Oxf2+='<tr><td width=18><div style="background-color:#000000; border:solid 1px #808080; width:12px; height:12px; font-size: 3px;">&nbsp;</div></td><td align=center style="font:normal 11px verdana;">&nbsp;<%= GetString("Automatic") %></td></tr>' ; Oxf2+=OxOb93c[0x34] ; Oxf2+=OxOb93c[0x35] ; Oxf2+=OxOb93c[0x36] ; Oxf2+=OxOb93c[0x37] ;for(var i=0x0;i<Oxf3;i++){if((i%Oxf4)==0x0){ Oxf2+=OxOb93c[0x2f] ;} ; Oxf2+=OxOb93c[0x38]+colorsArray[i]+OxOb93c[0x39]+colorsArray[i]+OxOb93c[0x3a]+element[OxOb93c[0x30]]+OxOb93c[0x3b] ; Oxf2+=OxOb93c[0x3c]+colorsArray[i]+OxOb93c[0x3d] ; Oxf2+=OxOb93c[0x35] ;if(((i+0x1)>=Oxf3)||(((i+0x1)%Oxf4)==0x0)){ Oxf2+=OxOb93c[0x36] ;} ;} ; Oxf2+=OxOb93c[0x37] ; Oxf2+=OxOb93c[0x34] ; Oxf2+=OxOb93c[0x3e] ;if(ShowMoreColors){ Oxf2+=OxOb93c[0x2f] ; Oxf2+=OxOb93c[0x3f]+element[OxOb93c[0x30]]+OxOb93c[0x40] ; Oxf2+=OxOb93c[0x33] ; Oxf2+='<tr><td width=18><div style="background-color:#000000; border:solid 1px #808080; width:12px; height:12px;font-size: 3px;"></div></td><td align=center style="font-size:11px"><%= GetString("MoreColors") %></td></tr>' ; Oxf2+=OxOb93c[0x34] ; Oxf2+=OxOb93c[0x35] ; Oxf2+=OxOb93c[0x36] ;} ; Oxf2+=OxOb93c[0x34] ; div[OxOb93c[0x41]]=Oxf2 ; element[OxOb93c[0x43]][OxOb93c[0x42]].appendChild(div) ; div[OxOb93c[0x44]]=CloseDiv ;} ;if(isopen){ CloseDiv() ;} ; isopen=true ; selects=[] ;var Ox87=element[OxOb93c[0x43]][OxOb93c[0x46]].tags(OxOb93c[0x45]);for(var i=0x0;i<Ox87[OxOb93c[0x2c]];i++){var Oxf5=Ox87[i];if(Oxf5[OxOb93c[0x48]][OxOb93c[0x47]]!=OxOb93c[0x49]){ selects[selects[OxOb93c[0x2c]]]=Oxf5 ;var Oxf6=Oxf5[OxOb93c[0x4b]]||Oxf5[OxOb93c[0x4a]]; Oxf6[OxOb93c[0x4c]]=Oxf6[OxOb93c[0x47]] ; Oxf6[OxOb93c[0x47]]=OxOb93c[0x49] ;} ;} ; div[OxOb93c[0x4a]][OxOb93c[0x4d]]=0x0 ; div[OxOb93c[0x4a]][OxOb93c[0x4e]]=0x0 ; div[OxOb93c[0x4a]][OxOb93c[0x4f]]=OxOb93c[0x50] ;var Oxf7=CalcPosition(div,element); Oxf7[OxOb93c[0x4d]]+=element[OxOb93c[0x51]] ; AdjustMirror(div,element,Oxf7) ; div[OxOb93c[0x4a]][OxOb93c[0x4e]]=Oxf7[OxOb93c[0x4e]]+OxOb93c[0x52] ; div[OxOb93c[0x4a]][OxOb93c[0x4d]]=Oxf7[OxOb93c[0x4d]]+OxOb93c[0x52] ;var Ox87=div[OxOb93c[0x46]];for(var i=0x0;i<Ox87[OxOb93c[0x2c]];i++){ Ox87[i][OxOb93c[0x53]]=OxOb93c[0x54] ;} ; div.focus() ;var Oxf8= new CaptureManager(element,handlelosecapture); Oxf8.AddElement(div) ;}  ; function handlelosecapture(){ CloseDiv() ;}  ; function CloseDiv(){ CaptureManager.Unregister(element) ; isopen=false ; div[OxOb93c[0x4a]][OxOb93c[0x4f]]=OxOb93c[0x55] ;for(var i=0x0;i<selects[OxOb93c[0x2c]];i++){var Oxf5=selects[i]; Oxf5[OxOb93c[0x4b]][OxOb93c[0x47]]=Oxf5[OxOb93c[0x4b]][OxOb93c[0x4c]] ;} ;}  ;

</script>

<script type="text/javascript">


var OxOa313=["body","document","compatMode","CSS1Compat","documentElement","scrollLeft","scrollTop","clientLeft","clientTop","parentElement","position","currentStyle","absolute","relative","left","top","clientWidth","clientHeight","offsetWidth","offsetHeight","capturemanager","element","\x3CDIV style=\x27width:0px;height:0px;left:0px;top:0px;position:absolute\x27\x3E","afterBegin","onlosecapture","onmousedown","onmousemove","onmouseover","onmouseout","length","clientX"]; function GetScrollPostion(Ox41){var Ox60=window[OxOa313[0x1]][OxOa313[0x0]];var p=Ox60;if(window[OxOa313[0x1]][OxOa313[0x2]]==OxOa313[0x3]){ p=window[OxOa313[0x1]][OxOa313[0x4]] ;} ;if(Ox41==Ox60){return {left:0x0,top:0x0};} ;with(Ox41.getBoundingClientRect()){return {left:p[OxOa313[0x5]]+left,top:p[OxOa313[0x6]]+top};} ;}  ; function GetClientPosition(Ox41){var Ox60=window[OxOa313[0x1]][OxOa313[0x0]];var p=Ox60;if(window[OxOa313[0x1]][OxOa313[0x2]]==OxOa313[0x3]){ p=window[OxOa313[0x1]][OxOa313[0x4]] ;} ;if(Ox41==Ox60){return {left:-p[OxOa313[0x5]],top:-p[OxOa313[0x6]]};} ;with(Ox41.getBoundingClientRect()){return {left:left-p[OxOa313[0x7]],top:top-p[OxOa313[0x8]]};} ;}  ; function GetStandParent(Ox41){for(var Ox100=Ox41[OxOa313[0x9]];Ox100!=null;Ox100=Ox100[OxOa313[0x9]]){var Ox101=Ox100[OxOa313[0xb]][OxOa313[0xa]];if(Ox101==OxOa313[0xc]||Ox101==OxOa313[0xd]){return Ox100;} ;} ;return window[OxOa313[0x1]][OxOa313[0x0]];}  ; function CalcPosition(Ox103,Ox41){var Ox104=GetScrollPostion(Ox41);var Ox105=GetScrollPostion(GetStandParent(Ox103));var Ox106=GetStandParent(Ox103);return {left:Ox104[OxOa313[0xe]]-Ox105[OxOa313[0xe]]-Ox106[OxOa313[0x7]],top:Ox104[OxOa313[0xf]]-Ox105[OxOa313[0xf]]-Ox106[OxOa313[0x8]]};}  ; function AdjustMirror(Ox103,Ox41,Oxf7){var Ox108=window[OxOa313[0x1]][OxOa313[0x0]][OxOa313[0x10]];var Ox109=window[OxOa313[0x1]][OxOa313[0x0]][OxOa313[0x11]];if(window[OxOa313[0x1]][OxOa313[0x2]]==OxOa313[0x3]){ Ox108=window[OxOa313[0x1]][OxOa313[0x4]][OxOa313[0x10]] ; Ox109=window[OxOa313[0x1]][OxOa313[0x4]][OxOa313[0x11]] ;} ;var Ox10a=Ox103[OxOa313[0x12]];var Ox10b=Ox103[OxOa313[0x13]];var Ox10c=GetClientPosition(GetStandParent(Ox103));var Ox10d={left:Ox10c[OxOa313[0xe]]+Oxf7[OxOa313[0xe]]+Ox10a/0x2,top:Ox10c[OxOa313[0xf]]+Oxf7[OxOa313[0xf]]+Ox10b/0x2};var Ox10e={left:Ox10c[OxOa313[0xe]]+Oxf7[OxOa313[0xe]],top:Ox10c[OxOa313[0xf]]+Oxf7[OxOa313[0xf]]};if(Ox41!=null){var Ox10f=GetClientPosition(Ox41); Ox10e={left:Ox10f[OxOa313[0xe]]+Ox41[OxOa313[0x12]]/0x2,top:Ox10f[OxOa313[0xf]]+Ox41[OxOa313[0x13]]/0x2} ;} ;var Ox110=true;if(Ox10d[OxOa313[0xe]]-Ox10a/0x2<0x0){if((Ox10e[OxOa313[0xe]]*0x2-Ox10d[OxOa313[0xe]])+Ox10a/0x2<=Ox108){ Ox10d[OxOa313[0xe]]=Ox10e[OxOa313[0xe]]*0x2-Ox10d[OxOa313[0xe]] ;} else {if(Ox110){ Ox10d[OxOa313[0xe]]=Ox10a/0x2+0x4 ;} ;} ;} else {if(Ox10d[OxOa313[0xe]]+Ox10a/0x2>Ox108){if((Ox10e[OxOa313[0xe]]*0x2-Ox10d[OxOa313[0xe]])-Ox10a/0x2>=0x0){ Ox10d[OxOa313[0xe]]=Ox10e[OxOa313[0xe]]*0x2-Ox10d[OxOa313[0xe]] ;} else {if(Ox110){ Ox10d[OxOa313[0xe]]=Ox108-Ox10a/0x2-0x4 ;} ;} ;} ;} ;if(Ox10d[OxOa313[0xf]]-Ox10b/0x2<0x0){if((Ox10e[OxOa313[0xf]]*0x2-Ox10d[OxOa313[0xf]])+Ox10b/0x2<=Ox109){ Ox10d[OxOa313[0xf]]=Ox10e[OxOa313[0xf]]*0x2-Ox10d[OxOa313[0xf]] ;} else {if(Ox110){ Ox10d[OxOa313[0xf]]=Ox10b/0x2+0x4 ;} ;} ;} else {if(Ox10d[OxOa313[0xf]]+Ox10b/0x2>Ox109){if((Ox10e[OxOa313[0xf]]*0x2-Ox10d[OxOa313[0xf]])-Ox10b/0x2>=0x0){ Ox10d[OxOa313[0xf]]=Ox10e[OxOa313[0xf]]*0x2-Ox10d[OxOa313[0xf]] ;} else {if(Ox110){ Ox10d[OxOa313[0xf]]=Ox109-Ox10b/0x2-0x4 ;} ;} ;} ;} ; Oxf7[OxOa313[0xe]]=Ox10d[OxOa313[0xe]]-Ox10c[OxOa313[0xe]]-Ox10a/0x2 ; Oxf7[OxOa313[0xf]]=Ox10d[OxOa313[0xf]]-Ox10c[OxOa313[0xf]]-Ox10b/0x2 ;}  ; function CaptureManager(element,handlelosecapture){if(CaptureManager[OxOa313[0x15]]&&CaptureManager[OxOa313[0x15]][OxOa313[0x14]]){ CaptureManager[OxOa313[0x15]][OxOa313[0x14]].Unregister() ;} ;var Ox112=true;var Ox113=[];var Ox114=true;var Ox115=false; element[OxOa313[0x14]]=Ox11a ; CaptureManager[OxOa313[0x15]]=element ; Ox11a.AddElement(element) ;var Ox116=element[OxOa313[0x1]].createElement(OxOa313[0x16]); element[OxOa313[0x1]][OxOa313[0x0]].insertAdjacentElement(OxOa313[0x17],Ox116) ; Ox116.attachEvent(OxOa313[0x18],Ox11c) ; Ox117(Ox116) ; Ox116.setCapture(true) ; Ox115=true ;return Ox11a; function Ox117(Ox118){ Ox118.attachEvent(OxOa313[0x19],Ox11d) ; Ox118.attachEvent(OxOa313[0x1a],Ox11f) ; Ox118.attachEvent(OxOa313[0x1b],Ox121) ; Ox118.attachEvent(OxOa313[0x1c],Ox122) ;}  ; function Ox119(Ox118){ Ox118.detachEvent(OxOa313[0x19],Ox11d) ; Ox118.detachEvent(OxOa313[0x1a],Ox11f) ; Ox118.detachEvent(OxOa313[0x1b],Ox121) ; Ox118.detachEvent(OxOa313[0x1c],Ox122) ;}  ; function Ox11a(){}  ; function Ox11a.Unregister(){ Ox116.detachEvent(OxOa313[0x18],Ox11c) ; Ox119(Ox116) ; Ox116.removeNode(true) ;for(var i=0x0;i<Ox113[OxOa313[0x1d]];i++){var Ox118=Ox113[i]; Ox119(Ox118) ;} ; Ox112=false ; element[OxOa313[0x14]]=null ; CaptureManager[OxOa313[0x15]]=null ;if(Ox115){ Ox115=false ; Ox116.releaseCapture() ; Ox11a.FireLoseCapture() ;} ;}  ; function Ox11a.AddElement(Ox118){ Ox117(Ox118) ; Ox113[Ox113[OxOa313[0x1d]]]=Ox118 ;}  ; function Ox11a.DelElement(Ox118){var len=Ox113[OxOa313[0x1d]];for(var i=0x0;i<len;i++){if(Ox113[i]==Ox118){ Ox119(Ox118) ;for(var j=i;j<len-0x1;j++){ Ox113[j]=Ox113[j+0x1] ;} ; Ox113[OxOa313[0x1d]]=Ox113[OxOa313[0x1d]]-0x1 ;return ;} ;} ;}  ; function Ox11a.FireLoseCapture(){ handlelosecapture() ;}  ; function Ox11c(){if(Ox115){ Ox115=false ;try{ Ox11a.FireLoseCapture() ;} finally{ Ox11a.Unregister() ;} ;} ;}  ; function Ox11d(){var Ox11e=element[OxOa313[0x1]].elementFromPoint(event[OxOa313[0x1e]],event.clientY);for(var i=0x0;i<Ox113[OxOa313[0x1d]];i++){var Ox118=Ox113[i];if(Ox118.contains(Ox11e)){return ;} ;} ; Ox11a.Unregister() ;}  ; function Ox11f(){var Ox11e=element[OxOa313[0x1]].elementFromPoint(event[OxOa313[0x1e]],event.clientY); Ox120(Ox11e) ;}  ; function Ox120(Ox11e){for(var i=0x0;i<Ox113[OxOa313[0x1d]];i++){var Ox118=Ox113[i];if(Ox118.contains(Ox11e)){if(Ox115){ Ox115=false ; Ox116.releaseCapture() ;} ;return ;} ;} ;if(!Ox115){ Ox115=true ; Ox116.setCapture(true) ;} ;}  ; function Ox121(){var Ox11e=element[OxOa313[0x1]].elementFromPoint(event[OxOa313[0x1e]],event.clientY); Ox114=false ;for(var i=0x0;i<Ox113[OxOa313[0x1d]];i++){var Ox118=Ox113[i];if(Ox118.contains(event.fromElement)){return ;} ;if(Ox118.contains(Ox11e)){if(Ox115){ Ox115=false ; Ox116.releaseCapture() ;} ;} ;} ;}  ; function Ox122(){var Ox11e=element[OxOa313[0x1]].elementFromPoint(event[OxOa313[0x1e]],event.clientY); Ox114=false ;for(var i=0x0;i<Ox113[OxOa313[0x1d]];i++){var Ox118=Ox113[i];if(Ox118.contains(event.toElement)){return ;} ;} ;if(!Ox115){ Ox115=true ; Ox116.setCapture(true) ;} ;}  ;}  ; function CaptureManager.Register(element,handlelosecapture){return  new CaptureManager(element,handlelosecapture);}  ; function CaptureManager.Unregister(element){if(element&&element[OxOa313[0x14]]){ element[OxOa313[0x14]].Unregister() ;} ;}  ;
</script>
