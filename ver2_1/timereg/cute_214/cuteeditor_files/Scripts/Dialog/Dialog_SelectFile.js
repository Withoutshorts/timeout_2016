var OxOa928=["top","dialogArguments","opener","_dialog_arguments","document","onload","value","","upload_image","browse_Frame","contentWindow","btn_CreateDir","btn_zoom_in","btn_zoom_out","btn_Actualsize","divpreview","TargetUrl","Button1","Button2","editor","window","\x3Cbr\x3E",".",".jpeg",".jpg",".gif",".png","\x3CIMG src=\x27","\x27 width=\x27150\x27\x3E",".bmp","\x26nbsp;\x3Cembed src=\x22","\x22 quality=\x22high\x22 width=\x22150\x22 height=\x22150\x22 type=\x22application/x-shockwave-flash\x22 pluginspage=\x22http://www.macromedia.com/go/getflashplayer\x22\x3E\x3C/embed\x3E\x0A","\x26nbsp;",".swf",".avi",".mpg",".mp3",".mpeg","\x26nbsp;\x3Cembed name=\x22MediaPlayer1\x22 src=\x22","\x22 autostart=-1 showcontrols=-1  type=\x22application/x-mplayer2\x22 width=\x22150\x22 height=\x22150\x22 pluginspage=\x22http://www.microsoft.com/Windows/MediaPlayer\x22 \x3E\x3C/embed\x3E\x0A",".wav","URL: ","innerHTML","inp","zoom","style","display","none","wrapupPrompt","iepromptfield","body","div","id","IEPromptBox","promptBlackout","border","1px solid #b0bec7","backgroundColor","#f0f0f0","position","absolute","width","330px","zIndex","100","\x3Cdiv style=\x22width: 100%; padding-top:3px;background-color: #DCE7EB; font-family: verdana; font-size: 10pt; font-weight: bold; height: 22px; text-align:center; background:url(../Images/formbg2.gif) repeat-x left top;\x22\x3E","\x3C/div\x3E","\x3Cdiv style=\x22padding: 10px\x22\x3E","\x3CBR\x3E\x3CBR\x3E","\x3Cform action=\x22\x22 onsubmit=\x22return wrapupPrompt()\x22\x3E","\x3Cinput id=\x22iepromptfield\x22 name=\x22iepromptdata\x22 type=text size=46 value=\x22","\x22\x3E","\x3Cbr\x3E\x3Cbr\x3E\x3Ccenter\x3E","\x3Cinput type=\x22submit\x22 value=\x22\x26nbsp;\x26nbsp;\x26nbsp;","\x26nbsp;\x26nbsp;\x26nbsp;\x22\x3E","\x26nbsp;\x26nbsp;\x26nbsp;\x26nbsp;\x26nbsp;\x26nbsp;","\x3Cinput type=\x22button\x22 onclick=\x22wrapupPrompt(true)\x22 value=\x22\x26nbsp;","\x26nbsp;\x22\x3E","\x3C/form\x3E\x3C/div\x3E","100px","left","offsetWidth","px","block","onmouseover","CuteEditor_ColorPicker_ButtonOver(this)"];function Window_FindDialogArguments(Ox95){var Ox132=Ox95[OxOa928[0]];if(Ox132[OxOa928[1]]){return Ox132[OxOa928[1]];} ;var Ox133=Ox132[OxOa928[2]];if(Ox133==null){return Ox132[OxOa928[4]][OxOa928[3]];} ;var Ox34f=Ox133[OxOa928[4]][OxOa928[3]];if(Ox34f==null){return Window_FindDialogArguments(Ox133);} ;return Ox34f;} ;function reset_hiddens(){} ;Event_Attach(window,OxOa928[5],reset_hiddens);function RequireFileBrowseScript(){} ;function reset_hiddens(){if(TargetUrl[OxOa928[6]]!=OxOa928[7]&&TargetUrl[OxOa928[6]]!=null){do_preview();} ;} ;RequireFileBrowseScript();var upload_image=Window_GetElement(window,OxOa928[8],true);var browse_Frame=Window_GetElement(window,OxOa928[9],true);browse_Frame=browse_Frame[OxOa928[10]];var btn_CreateDir=Window_GetElement(window,OxOa928[11],true);var btn_zoom_in=Window_GetElement(window,OxOa928[12],true);var btn_zoom_out=Window_GetElement(window,OxOa928[13],true);var btn_Actualsize=Window_GetElement(window,OxOa928[14],true);var divpreview=Window_GetElement(window,OxOa928[15],true);var TargetUrl=Window_GetElement(window,OxOa928[16],true);var Button1=Window_GetElement(window,OxOa928[17],true);var Button2=Window_GetElement(window,OxOa928[18],true);var arg=Window_FindDialogArguments(window);var editor=arg[OxOa928[19]];var editwin=arg[OxOa928[20]];var editdoc=arg[OxOa928[4]];do_preview();function do_preview(Ox17c){var Ox61;Ox61=OxOa928[7];if(Ox17c!=OxOa928[7]&&Ox17c!=null){Ox61=Ox17c;} ;Ox61=Ox61+OxOa928[21];var Ox185=TargetUrl[OxOa928[6]];if(Ox185==OxOa928[7]){return ;} ;var Ox2a3=Ox185.substring(Ox185.lastIndexOf(OxOa928[22])).toLowerCase();switch(Ox2a3){case OxOa928[23]:;case OxOa928[24]:;case OxOa928[25]:;case OxOa928[26]:;case OxOa928[29]:Ox61=Ox61+OxOa928[27]+Ox185+OxOa928[28];break ;;case OxOa928[33]:var Ox2a4=OxOa928[30]+Ox185+OxOa928[31];Ox61=Ox61+Ox2a4+OxOa928[32];break ;;case OxOa928[34]:;case OxOa928[35]:;case OxOa928[36]:;case OxOa928[37]:;case OxOa928[40]:var Ox2a5=OxOa928[38]+Ox185+OxOa928[39];Ox61=Ox61+Ox2a5+OxOa928[32];break ;;default:Ox61=Ox61+OxOa928[41]+TargetUrl[OxOa928[6]];break ;;} ;divpreview[OxOa928[42]]=Ox61;} ;function do_insert(){var Ox351=arg[OxOa928[43]];if(Ox351){try{Ox351[OxOa928[6]]=TargetUrl[OxOa928[6]];} catch(x){} ;} ;Window_SetDialogReturnValue(window,TargetUrl.value);Window_CloseDialog(window);} ;function do_Close(){Window_SetDialogReturnValue(window,null);Window_CloseDialog(window);} ;function Zoom_In(){if(divpreview[OxOa928[45]][OxOa928[44]]!=0){divpreview[OxOa928[45]][OxOa928[44]]*=1.2;} else {divpreview[OxOa928[45]][OxOa928[44]]=1.2;} ;} ;function Zoom_Out(){if(divpreview[OxOa928[45]][OxOa928[44]]!=0){divpreview[OxOa928[45]][OxOa928[44]]*=0.8;} else {divpreview[OxOa928[45]][OxOa928[44]]=0.8;} ;} ;function Actualsize(){divpreview[OxOa928[45]][OxOa928[44]]=1;do_preview();} ;function ResetFields(){TargetUrl[OxOa928[6]]=OxOa928[7];} ;if(!Browser_IsWinIE()){btn_zoom_in[OxOa928[45]][OxOa928[46]]=btn_zoom_out[OxOa928[45]][OxOa928[46]]=btn_Actualsize[OxOa928[45]][OxOa928[46]]=OxOa928[47];} ;if(!Browser_IsWinIE()){btn_zoom_in[OxOa928[45]][OxOa928[46]]=btn_zoom_out[OxOa928[45]][OxOa928[46]]=btn_Actualsize[OxOa928[45]][OxOa928[46]]=OxOa928[47];} else {} ;if(Browser_IsIE7()){var _dialogPromptID=null;function IEprompt(Ox112,Ox113,Ox114){that=this;this[OxOa928[48]]=function (Ox115){val=document.getElementById(OxOa928[49])[OxOa928[6]];_dialogPromptID[OxOa928[45]][OxOa928[46]]=OxOa928[47];document.getElementById(OxOa928[49])[OxOa928[6]]=OxOa928[7];if(Ox115){val=OxOa928[7];} ;Ox112(val);return false;} ;if(Ox114==undefined){Ox114=OxOa928[7];} ;if(_dialogPromptID==null){var Ox116=document.getElementsByTagName(OxOa928[50])[0];tnode=document.createElement(OxOa928[51]);tnode[OxOa928[52]]=OxOa928[53];Ox116.appendChild(tnode);_dialogPromptID=document.getElementById(OxOa928[53]);tnode=document.createElement(OxOa928[51]);tnode[OxOa928[52]]=OxOa928[54];Ox116.appendChild(tnode);_dialogPromptID[OxOa928[45]][OxOa928[55]]=OxOa928[56];_dialogPromptID[OxOa928[45]][OxOa928[57]]=OxOa928[58];_dialogPromptID[OxOa928[45]][OxOa928[59]]=OxOa928[60];_dialogPromptID[OxOa928[45]][OxOa928[61]]=OxOa928[62];_dialogPromptID[OxOa928[45]][OxOa928[63]]=OxOa928[64];} ;var Ox117=OxOa928[65]+InputRequired+OxOa928[66];Ox117+=OxOa928[67]+Ox113+OxOa928[68];Ox117+=OxOa928[69];Ox117+=OxOa928[70]+Ox114+OxOa928[71];Ox117+=OxOa928[72];Ox117+=OxOa928[73]+OK+OxOa928[74];Ox117+=OxOa928[75];Ox117+=OxOa928[76]+Cancel+OxOa928[77];Ox117+=OxOa928[78];_dialogPromptID[OxOa928[42]]=Ox117;_dialogPromptID[OxOa928[45]][OxOa928[0]]=OxOa928[79];_dialogPromptID[OxOa928[45]][OxOa928[80]]=parseInt((document[OxOa928[50]][OxOa928[81]]-315)/2)+OxOa928[82];_dialogPromptID[OxOa928[45]][OxOa928[46]]=OxOa928[83];var Ox118=document.getElementById(OxOa928[49]);try{var Ox119=Ox118.createTextRange();Ox119.collapse(false);Ox119.select();} catch(x){Ox118.focus();} ;} ;} ;if(btn_CreateDir){btn_CreateDir[OxOa928[84]]= new Function(OxOa928[85]);} ;if(btn_zoom_in){btn_zoom_in[OxOa928[84]]= new Function(OxOa928[85]);} ;if(btn_zoom_out){btn_zoom_out[OxOa928[84]]= new Function(OxOa928[85]);} ;if(btn_Actualsize){btn_Actualsize[OxOa928[84]]= new Function(OxOa928[85]);} ;