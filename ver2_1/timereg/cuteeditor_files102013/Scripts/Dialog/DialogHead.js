var OxOe806=["ua","userAgent","isOpera","opera","isSafari","safari","isGecko","gecko","isWinIE","msie","isMac","Mac","compatMode","document","CSS1Compat","XMLHttpRequest","","caller","(","\x0A","attachEvent","on","\x0D\x0A","closeeditordialog","close","top","returnValue","_dialog_returnvalue","opener","_dialog_arguments","dialogArguments","length","element \x27","\x27 not found","all","childNodes","nodeType","UNSELECTABLE","tabIndex","-1","//TODO: event not found? throw error ?","preventDefault","event","arguments","parent","frames","srcElement","target","//TODO: srcElement not found? throw error ?","fromElement","relatedTarget","toElement","keyCode","which","clientX","clientY","offsetX","offsetY","button","ctrlKey","altKey","shiftKey","cancelBubble","stopPropagation","buttonInitialized","oncontextmenu","onmouseout","onmousedown","onmouseup","isover","className","CuteEditorButtonOver","CuteEditorButton","CuteEditorButtonDown","CuteEditorDown","border","style","solid 1px #0A246A","backgroundColor","#b6bdd2","padding","1px","solid 1px #f5f5f4","inset 1px","IsCommandDisabled","CuteEditorButtonDisabled","IsCommandActive","CuteEditorButtonActive","onerror","\x0D\x0A\x0D\x0A","href","location","view-source:","?\x26view-source=","_blank","ondblclick","onkeydown","//duplicated?\x0D\x0A","var ","=Window_GetElement(window,\x22","\x22,true);","Text","clipboardData","addEventListener","isdir","path","url","text","return this.getAttribute(\x27","\x27);","prototype","attributes","\x3C","tagName","specified"," ","name","=\x22","value","\x22","\x3E","innerHTML","\x3C/","AREA","BASE","BASEFONT","COL","FRAME","HR","IMG","BR","INPUT","ISINDEX","LINK","META","PARAM","00","0123456789ABCDEF","#","object","undefined","Microsoft.XMLHTTP","head","script","language","javascript","type","text/javascript","src","id","_cpinstalled","1","../Scripts/ColorPicker.js","CuteEditorColorPickerInitialize","GET","onreadystatechange","readyState","responseText"," \x0D\x0A window.CuteEditorColorPickerInitialize=CuteEditorColorPickerInitialize","colorScript","oncolorselect","FireUIChanged","onclick","popupColorPicker"];var _Browser_TypeInfo=null;function Browser__InitType(){if(_Browser_TypeInfo!=null){return ;} ;var Ox11d={};Ox11d[OxOe806[0]]=navigator[OxOe806[1]].toLowerCase();Ox11d[OxOe806[2]]=(Ox11d[OxOe806[0]].indexOf(OxOe806[3])>-1);Ox11d[OxOe806[4]]=(Ox11d[OxOe806[0]].indexOf(OxOe806[5])>-1);Ox11d[OxOe806[6]]=(!Ox11d[OxOe806[2]]&&Ox11d[OxOe806[0]].indexOf(OxOe806[7])>-1);Ox11d[OxOe806[8]]=(!Ox11d[OxOe806[2]]&&Ox11d[OxOe806[0]].indexOf(OxOe806[9])>-1);Ox11d[OxOe806[10]]=navigator[OxOe806[1]].indexOf(OxOe806[11])!=-1;_Browser_TypeInfo=Ox11d;} ;Browser__InitType();function Browser_IsWinIE(){return _Browser_TypeInfo[OxOe806[8]];} ;function Browser_IsGecko(){return _Browser_TypeInfo[OxOe806[6]];} ;function Browser_IsOpera(){return _Browser_TypeInfo[OxOe806[2]];} ;function Browser_IsSafari(){return _Browser_TypeInfo[OxOe806[4]];} ;function Browser_UseIESelection(){return _Browser_TypeInfo[OxOe806[8]];} ;function Browser_IsCSS1Compat(){return window[OxOe806[13]][OxOe806[12]]==OxOe806[14];} ;function Browser_IsIE7(){return _Browser_TypeInfo[OxOe806[8]]&&window[OxOe806[15]];} ;function GetStackTrace(){var Ox2a=OxOe806[16];for(var Ox126=GetStackTrace[OxOe806[17]];Ox126!=null;Ox126=Ox126[OxOe806[17]]){var Ox127=Ox126+OxOe806[16];Ox127=Ox127.substr(0,Ox127.indexOf(OxOe806[18]));Ox2a+=Ox127+OxOe806[19];} ;return Ox2a;} ;function Event_Attach(obj,name,Ox12a){if(obj[OxOe806[20]]){if(name.substr(0,2)!=OxOe806[21]){name=OxOe806[21]+name;} ;obj.attachEvent(name,Ox12a);} else {if(name.substr(0,2)==OxOe806[21]){name=name.substring(2);} ;obj.addEventListener(name,Ox12a,false);} ;} ;var __ISDEBUG=false;function Debug_Todo(msg){if(!__ISDEBUG){return ;} ;throw ( new Error(msg+OxOe806[22]+Debug_Todo[OxOe806[17]]));} ;function Window_CloseDialog(Ox12f){(top[OxOe806[23]]||top[OxOe806[24]])();} ;function Window_SetDialogReturnValue(Ox95,Ox4f){var Ox131=Ox95[OxOe806[25]];Ox131[OxOe806[26]]=Ox4f;Ox131[OxOe806[13]][OxOe806[27]]=Ox4f;var Ox132=Ox131[OxOe806[28]];if(Ox132==null){return ;} ;try{Ox132[OxOe806[13]][OxOe806[27]]=Ox4f;} catch(x){} ;} ;function Window_GetDialogArguments(Ox95){var Ox131=Ox95[OxOe806[25]];try{var Ox132=Ox131[OxOe806[28]];if(Ox132&&Ox132[OxOe806[13]][OxOe806[29]]){return Ox132[OxOe806[13]][OxOe806[29]];} ;} catch(x){} ;if(Ox131[OxOe806[13]][OxOe806[29]]){return Ox131[OxOe806[13]][OxOe806[29]];} ;if(Ox131[OxOe806[30]]){return Ox131[OxOe806[30]];} ;return Ox131[OxOe806[13]][OxOe806[29]];} ;function Window_GetElement(Ox95,Oxaf,Ox135){var Ox136=Ox95[OxOe806[13]].getElementById(Oxaf);if(Ox136){return Ox136;} ;var Ox16=Ox95[OxOe806[13]].getElementsByName(Oxaf);if(Ox16[OxOe806[31]]>0){return Ox16.item(0);} ;if(Ox135){if(__ISDEBUG){alert(OxOe806[32]+Oxaf+OxOe806[33]);} ;throw ( new Error(OxOe806[32]+Oxaf+OxOe806[33]));} ;return null;} ;function Element_GetAllElements(p){var arr=[];if(Browser_IsWinIE()){for(var i=0;i<p[OxOe806[34]][OxOe806[31]];i++){arr.push(p[OxOe806[34]].item(i));} ;return arr;} ;Ox138(p);function Ox138(Ox136){var Ox139=Ox136[OxOe806[35]];var Ox69=Ox139[OxOe806[31]];for(var i=0;i<Ox69;i++){var Ox93=Ox139.item(i);if(Ox93[OxOe806[36]]!=1){continue ;} ;arr.push(Ox93);Ox138(Ox93);} ;} ;return arr;} ;function Element_SetUnselectable(element){element.setAttribute(OxOe806[37],OxOe806[21]);element.setAttribute(OxOe806[38],OxOe806[39]);var arr=Element_GetAllElements(element);var len=arr[OxOe806[31]];if(!len){return ;} ;for(var i=0;i<len;i++){arr[i].setAttribute(OxOe806[37],OxOe806[21]);arr[i].setAttribute(OxOe806[38],OxOe806[39]);} ;} ;function Event_GetEvent(Ox13c){Ox13c=Event_FindEvent(Ox13c);if(Ox13c==null){Debug_Todo(OxOe806[40]);} ;return Ox13c;} ;function Array_IndexOf(arr,Ox13e){for(var i=0;i<arr[OxOe806[31]];i++){if(arr[i]==Ox13e){return i;} ;} ;return -1;} ;function Array_Contains(arr,Ox13e){return Array_IndexOf(arr,Ox13e)!=-1;} ;function clearArray(Ox141){for(var i=0;i<Ox141[OxOe806[31]];i++){Ox141[i]=null;} ;} ;function Event_FindEvent(Ox13c){if(Ox13c&&Ox13c[OxOe806[41]]){return Ox13c;} ;if(Browser_IsGecko()){return Event_FindEvent_FindEventFromCallers();} else {if(window[OxOe806[42]]){return window[OxOe806[42]];} ;return Event_FindEvent_FindEventFromWindows();} ;return null;} ;function Event_FindEvent_FindEventFromCallers(){var Ox7b=Event_GetEvent[OxOe806[17]];for(var i=0;i<100;i++){if(!Ox7b){break ;} ;var Ox13c=Ox7b[OxOe806[43]][0];if(Ox13c&&Ox13c[OxOe806[41]]){return Ox13c;} ;Ox7b=Ox7b[OxOe806[17]];} ;} ;function Event_FindEvent_FindEventFromWindows(){var arr=[];return Ox145(window);function Ox145(Ox95){if(Ox95==null){return null;} ;if(Ox95[OxOe806[42]]){return Ox95[OxOe806[42]];} ;if(Array_Contains(arr,Ox95)){return null;} ;arr.push(Ox95);var Ox146=[];if(Ox95[OxOe806[44]]!=Ox95){Ox146.push(Ox95.parent);} ;if(Ox95[OxOe806[25]]!=Ox95[OxOe806[44]]){Ox146.push(Ox95.top);} ;if(Ox95[OxOe806[28]]){Ox146.push(Ox95.opener);} ;for(var i=0;i<Ox95[OxOe806[45]][OxOe806[31]];i++){Ox146.push(Ox95[OxOe806[45]][i]);} ;for(var i=0;i<Ox146[OxOe806[31]];i++){try{var Ox13c=Ox145(Ox146[i]);if(Ox13c){return Ox13c;} ;} catch(x){} ;} ;return null;} ;} ;function Event_GetSrcElement(Ox13c){Ox13c=Event_GetEvent(Ox13c);if(Ox13c[OxOe806[46]]){return Ox13c[OxOe806[46]];} ;if(Ox13c[OxOe806[47]]){return Ox13c[OxOe806[47]];} ;Debug_Todo(OxOe806[48]);return null;} ;function Event_GetFromElement(Ox13c){Ox13c=Event_GetEvent(Ox13c);if(Ox13c[OxOe806[49]]){return Ox13c[OxOe806[49]];} ;if(Ox13c[OxOe806[50]]){return Ox13c[OxOe806[50]];} ;return null;} ;function Event_GetToElement(Ox13c){Ox13c=Event_GetEvent(Ox13c);if(Ox13c[OxOe806[51]]){return Ox13c[OxOe806[51]];} ;if(Ox13c[OxOe806[50]]){return Ox13c[OxOe806[50]];} ;return null;} ;function Event_GetKeyCode(Ox13c){Ox13c=Event_GetEvent(Ox13c);return Ox13c[OxOe806[52]]||Ox13c[OxOe806[53]];} ;function Event_GetClientX(Ox13c){Ox13c=Event_GetEvent(Ox13c);return Ox13c[OxOe806[54]];} ;function Event_GetClientY(Ox13c){Ox13c=Event_GetEvent(Ox13c);return Ox13c[OxOe806[55]];} ;function Event_GetOffsetX(Ox13c){Ox13c=Event_GetEvent(Ox13c);return Ox13c[OxOe806[56]];} ;function Event_GetOffsetY(Ox13c){Ox13c=Event_GetEvent(Ox13c);return Ox13c[OxOe806[57]];} ;function Event_IsLeftButton(Ox13c){Ox13c=Event_GetEvent(Ox13c);if(Browser_IsWinIE()){return Ox13c[OxOe806[58]]==1;} ;if(Browser_IsGecko()){return Ox13c[OxOe806[58]]==0;} ;return Ox13c[OxOe806[58]]==0;} ;function Event_IsCtrlKey(Ox13c){Ox13c=Event_GetEvent(Ox13c);return Ox13c[OxOe806[59]];} ;function Event_IsAltKey(Ox13c){Ox13c=Event_GetEvent(Ox13c);return Ox13c[OxOe806[60]];} ;function Event_IsShiftKey(Ox13c){Ox13c=Event_GetEvent(Ox13c);return Ox13c[OxOe806[61]];} ;function Event_PreventDefault(Ox13c){Ox13c=Event_GetEvent(Ox13c);Ox13c[OxOe806[26]]=false;if(Ox13c[OxOe806[41]]){Ox13c.preventDefault();} ;} ;function Event_CancelBubble(Ox13c){Ox13c=Event_GetEvent(Ox13c);Ox13c[OxOe806[62]]=true;if(Ox13c[OxOe806[63]]){Ox13c.stopPropagation();} ;return false;} ;function Event_CancelEvent(Ox13c){Ox13c=Event_GetEvent(Ox13c);Event_PreventDefault(Ox13c);return Event_CancelBubble(Ox13c);} ;function CuteEditor_ButtonOver(element){if(!element[OxOe806[64]]){element[OxOe806[65]]=Event_CancelEvent;element[OxOe806[66]]=CuteEditor_ButtonOut;element[OxOe806[67]]=CuteEditor_ButtonDown;element[OxOe806[68]]=CuteEditor_ButtonUp;Element_SetUnselectable(element);element[OxOe806[64]]=true;} ;element[OxOe806[69]]=true;element[OxOe806[70]]=OxOe806[71];} ;function CuteEditor_ButtonOut(){var element=this;element[OxOe806[70]]=OxOe806[72];element[OxOe806[69]]=false;} ;function CuteEditor_ButtonDown(){if(!Event_IsLeftButton()){return ;} ;var element=this;element[OxOe806[70]]=OxOe806[73];} ;function CuteEditor_ButtonUp(){if(!Event_IsLeftButton()){return ;} ;var element=this;if(element[OxOe806[69]]){element[OxOe806[70]]=OxOe806[71];} else {element[OxOe806[70]]=OxOe806[74];} ;} ;function CuteEditor_ColorPicker_ButtonOver(element){if(!element[OxOe806[64]]){element[OxOe806[65]]=Event_CancelEvent;element[OxOe806[66]]=CuteEditor_ColorPicker_ButtonOut;element[OxOe806[67]]=CuteEditor_ColorPicker_ButtonDown;Element_SetUnselectable(element);element[OxOe806[64]]=true;} ;element[OxOe806[69]]=true;element[OxOe806[76]][OxOe806[75]]=OxOe806[77];element[OxOe806[76]][OxOe806[78]]=OxOe806[79];element[OxOe806[76]][OxOe806[80]]=OxOe806[81];} ;function CuteEditor_ColorPicker_ButtonOut(){var element=this;element[OxOe806[69]]=false;element[OxOe806[76]][OxOe806[75]]=OxOe806[82];element[OxOe806[76]][OxOe806[78]]=OxOe806[16];element[OxOe806[76]][OxOe806[80]]=OxOe806[81];} ;function CuteEditor_ColorPicker_ButtonDown(){var element=this;element[OxOe806[76]][OxOe806[75]]=OxOe806[83];element[OxOe806[76]][OxOe806[78]]=OxOe806[16];element[OxOe806[76]][OxOe806[80]]=OxOe806[81];} ;function CuteEditor_ButtonCommandOver(element){element[OxOe806[69]]=true;if(element[OxOe806[84]]){element[OxOe806[70]]=OxOe806[85];} else {element[OxOe806[70]]=OxOe806[71];} ;} ;function CuteEditor_ButtonCommandOut(element){element[OxOe806[69]]=false;if(element[OxOe806[86]]){element[OxOe806[70]]=OxOe806[87];} else {if(element[OxOe806[84]]){element[OxOe806[70]]=OxOe806[85];} else {element[OxOe806[70]]=OxOe806[72];} ;} ;} ;function CuteEditor_ButtonCommandDown(element){if(!Event_IsLeftButton()){return ;} ;element[OxOe806[70]]=OxOe806[73];} ;function CuteEditor_ButtonCommandUp(element){if(!Event_IsLeftButton()){return ;} ;if(element[OxOe806[84]]){element[OxOe806[70]]=OxOe806[85];return ;} ;if(element[OxOe806[69]]){element[OxOe806[70]]=OxOe806[71];} else {if(element[OxOe806[86]]){element[OxOe806[70]]=OxOe806[87];} else {element[OxOe806[70]]=OxOe806[72];} ;} ;} [CuteEditor_ButtonOver,CuteEditor_ColorPicker_ButtonOver];[Window_GetDialogArguments,Window_SetDialogReturnValue,Window_CloseDialog,Window_GetElement];function CancelEventIfNotDigit(){var Ox162=Event_GetKeyCode();if(Browser_IsWinIE()){if((Ox162>=48)&&(Ox162<=57)){return true;} ;} else {if((Ox162<48||Ox162>57)&&Ox162!=8&&(Ox162<35||Ox162>37)&&Ox162!=39&&Ox162!=46){} else {return true;} ;} ;return Event_CancelEvent();} ;window[OxOe806[88]]=function window_onerror(Ox164,Ox20,Ox139){if(!__ISDEBUG){return false;} ;alert(Ox164+OxOe806[22]+Ox20+OxOe806[22]+Ox139+OxOe806[89]+GetStackTrace());return true;} ;function DialogHandleDblClick(){if(Event_IsCtrlKey()){window[OxOe806[91]][OxOe806[90]]=OxOe806[92]+window[OxOe806[91]][OxOe806[90]]+OxOe806[93]+ new Date().getTime();} ;if(Event_IsShiftKey()){window.open(window[OxOe806[91]].href,OxOe806[94]);} ;} ;function DialogHandleKeyDown(){var Ox167=Event_GetKeyCode();if(Ox167==116){window[OxOe806[91]].reload();} ;if(Ox167==27){Window_SetDialogReturnValue(window,false);Window_CloseDialog(window);} ;} ;Event_Attach(document,OxOe806[95],DialogHandleDblClick);Event_Attach(document,OxOe806[96],DialogHandleKeyDown);function Debug_ReportElements(Ox169){var Ox16a={};var Ox16b=[];function Ox16c(Oxaf){if(!Oxaf){return ;} ;if(Ox16a[Oxaf]){Ox16b.push(OxOe806[97]);} ;Ox16a[Oxaf]=true;Ox16b.push(OxOe806[98]);Ox16b.push(Oxaf);Ox16b.push(OxOe806[99]);Ox16b.push(Oxaf);Ox16b.push(OxOe806[100]);Ox16b.push(OxOe806[22]);} ;var arr=Element_GetAllElements(Ox169);for(var i=0;i<arr[OxOe806[31]];i++){var Ox1f=arr[i];Ox16c(Ox1f.id);} ;var Oxf2=Ox16b.join(OxOe806[16]);window[OxOe806[102]].setData(OxOe806[101],Oxf2);} ;if(document[OxOe806[103]]){var rowprops=[OxOe806[104],OxOe806[105],OxOe806[106],OxOe806[107]];for(var rowpropi=0;rowpropi<rowprops[OxOe806[31]];rowpropi++){try{HTMLElement[OxOe806[110]].__defineGetter__(rowprops[rowpropi], new Function(OxOe806[108]+rowprops[rowpropi]+OxOe806[109]));} catch(x){} ;} ;} ;function outerHTML(element){var attr;var Ox171=element[OxOe806[111]];var Ox61=OxOe806[112]+element[OxOe806[113]];for(var i=0;i<Ox171[OxOe806[31]];i++){attr=Ox171[i];if(attr[OxOe806[114]]){Ox61+=OxOe806[115]+attr[OxOe806[116]]+OxOe806[117]+attr[OxOe806[118]]+OxOe806[119];} ;} ;if(!canHaveChildren(element)){return Ox61+OxOe806[120];} ;return Ox61+OxOe806[120]+element[OxOe806[121]]+OxOe806[122]+element[OxOe806[113]]+OxOe806[120];} ;function canHaveChildren(element){switch(element[OxOe806[113]]){case OxOe806[123]:;case OxOe806[124]:;case OxOe806[125]:;case OxOe806[126]:;case OxOe806[127]:;case OxOe806[128]:;case OxOe806[129]:;case OxOe806[130]:;case OxOe806[131]:;case OxOe806[132]:;case OxOe806[133]:;case OxOe806[134]:;case OxOe806[135]:return false;;} ;return true;} ;function RGBtoHex(Ox174,Ox175,Ox176){return toHex(Ox174)+toHex(Ox175)+toHex(Ox176);} ;function toHex(Ox178){if(Ox178==null){return OxOe806[136];} ;Ox178=parseInt(Ox178);if(Ox178==0||isNaN(Ox178)){return OxOe806[136];} ;Ox178=Math.max(0,Ox178);Ox178=Math.min(Ox178,255);Ox178=Math.round(Ox178);return OxOe806[137].charAt((Ox178-Ox178%16)/16)+OxOe806[137].charAt(Ox178%16);} ;var dec_pattern=/rgb\((\d{1,3})[,]\s*(\d{1,3})[,]\s*(\d{1,3})\)/gi;function revertColor(Ox17b){if(Ox17b.match(dec_pattern)){var Ox17c=Ox17b.replace(dec_pattern,function (Ox61,Ox17d,Ox17e,Ox17f){return (OxOe806[138]+RGBtoHex(Ox17d,Ox17e,Ox17f)).toLowerCase();} );return Ox17c;} ;return Ox17b;} ;function isNull(Ox164){return  typeof Ox164==OxOe806[139]&&!Ox164;} ;function CreateXMLHttpRequest(){try{if( typeof (XMLHttpRequest)!=OxOe806[140]){return  new XMLHttpRequest();} ;if( typeof (ActiveXObject)!=OxOe806[140]){return  new ActiveXObject(OxOe806[141]);} ;} catch(x){return null;} ;} ;function include(Ox183,Ox184){var Ox185=document.getElementsByTagName(OxOe806[142]).item(0);var Ox186=document.getElementById(Ox183);if(Ox186){Ox185.removeChild(Ox186);} ;var Ox187=document.createElement(OxOe806[143]);Ox187.setAttribute(OxOe806[144],OxOe806[145]);Ox187.setAttribute(OxOe806[146],OxOe806[147]);Ox187.setAttribute(OxOe806[148],Ox184);Ox187.setAttribute(OxOe806[149],Ox183);Ox185.appendChild(Ox187);} ;function SelectColor(Ox189,Ox18a){if(Ox189.getAttribute(OxOe806[150])==OxOe806[151]){return ;} ;var Ox18b=OxOe806[152];if(!window[OxOe806[153]]){var Ox18c=CreateXMLHttpRequest();if(Ox18c){Ox18c.open(OxOe806[154],Ox18b,true,null,null);Ox18c[OxOe806[155]]=function (){if(Ox18c[OxOe806[156]]<4){return ;} ;var Ox162=Ox18c[OxOe806[157]];Ox18c=null;eval(Ox162+OxOe806[158]);Ox18d();} ;Ox18c.send(OxOe806[16]);} else {include(OxOe806[159],Ox18b);setTimeout(Ox18d,100);} ;} else {Ox18d();} ;function Ox18d(){Ox189.setAttribute(OxOe806[150],OxOe806[151]);Ox189[OxOe806[118]]=OxOe806[16];window.CuteEditorColorPickerInitialize(Ox189,window.editor);Ox189[OxOe806[160]]=function (Ox4){if(Ox4!=null&&Ox4!==false){Ox189[OxOe806[118]]=Ox4.toUpperCase();Ox189[OxOe806[76]][OxOe806[78]]=Ox4;if(Ox18a){Ox18a[OxOe806[76]][OxOe806[78]]=Ox4;} ;if(window[OxOe806[161]]){window.FireUIChanged();} ;} ;} ;Ox189[OxOe806[162]]=Ox189[OxOe806[163]];if(Ox18a){Ox18a[OxOe806[162]]=function (){setTimeout(Ox189.popupColorPicker,100);} ;} ;setTimeout(Ox189.popupColorPicker,100);} ;} ;function row_click(src){} ;function do_Close(){Window_SetDialogReturnValue(window,null);Window_CloseDialog(window);} ;