var OxO3645=["value","","nodeName","nodeType","divpreview","Width","Height","AutoStart","ShowControls","ShowStatusBar","WindowlessVideo","TargetUrl","Button1","Button2","btn_zoom_in","btn_zoom_out","btn_Actualsize","btn_CreateDir","upload_image","contentWindow","checked","\x3Cembed name=\x22MediaPlayer1\x22 src=\x22","\x22 autostart=\x22","\x22 showcontrols=\x22","\x22  windowlessvideo=\x22","\x22 showstatusbar=\x22","\x22 width=\x22","\x22 height=\x22","\x22 type=\x22application/x-mplayer2\x22 pluginspage=\x22http://www.microsoft.com/Windows/MediaPlayer\x22 \x3E\x3C/embed\x3E\x0A","\x3Cobject classid=\x22CLSID:22D6F312-B0F6-11D0-94AB-0080C74C7E95\x22 "," codebase=\x22http://activex.microsoft.com/activex/"," controls/mplayer/en/nsmp2inf.cab#Version=6,0,02,902\x22 "," standby=\x22Loading Microsoft Windows Media Player components...\x22 "," type=\x22application/x-oleobject\x22","  height=\x22","\x22 \x3E","\x3Cparam name=\x22FileName\x22 value=\x22","\x22/\x3E","\x3Cparam name=\x22autoStart\x22 value=\x22","\x3Cparam name=\x22showControls\x22 value=\x22","\x3Cparam name=\x22showstatusbar\x22 value=\x22","\x3Cparam name=\x22windowlessvideo\x22 value=\x22","\x3C/object\x3E","innerHTML","onunload","onbeforeunload","Please choose a Media movie to insert","\x22 windowlessvideo=\x22","zoom","style","wrapupPrompt","iepromptfield","display","none","body","div","id","IEPromptBox","promptBlackout","border","1px solid #b0bec7","backgroundColor","#f0f0f0","position","absolute","width","330px","zIndex","100","\x3Cdiv style=\x22width: 100%; padding-top:3px;background-color: #DCE7EB; font-family: verdana; font-size: 10pt; font-weight: bold; height: 22px; text-align:center; background:url(../Images/formbg2.gif) repeat-x left top;\x22\x3E","\x3C/div\x3E","\x3Cdiv style=\x22padding: 10px\x22\x3E","\x3CBR\x3E\x3CBR\x3E","\x3Cform action=\x22\x22 onsubmit=\x22return wrapupPrompt()\x22\x3E","\x3Cinput id=\x22iepromptfield\x22 name=\x22iepromptdata\x22 type=text size=46 value=\x22","\x22\x3E","\x3Cbr\x3E\x3Cbr\x3E\x3Ccenter\x3E","\x3Cinput type=\x22submit\x22 value=\x22\x26nbsp;\x26nbsp;\x26nbsp;","\x26nbsp;\x26nbsp;\x26nbsp;\x22\x3E","\x26nbsp;\x26nbsp;\x26nbsp;\x26nbsp;\x26nbsp;\x26nbsp;","\x3Cinput type=\x22button\x22 onclick=\x22wrapupPrompt(true)\x22 value=\x22\x26nbsp;","\x26nbsp;\x22\x3E","\x3C/form\x3E\x3C/div\x3E","top","100px","offsetWidth","left","px","block","CuteEditor_ColorPicker_ButtonOver(this)","onmouseover"]; setMouseOver() ; function setMouseOver(){}  ; function ResetFields(){ TargetUrl[OxO3645[0x0]]=OxO3645[0x1] ;}  ; function RequireFileBrowseScript(){}  ; function Actualsize(){}  ; function getParent(Ox80,Oxaf){if(Ox80==null){return null;} else {if(Ox80[OxO3645[0x3]]==0x1&&Ox80[OxO3645[0x2]].toLowerCase()==Oxaf.toLowerCase()){return Ox80;} else {return getParent(Ox80.parentNode,Oxaf);} ;} ;}  ; RequireFileBrowseScript() ;var divpreview=Window_GetElement(window,OxO3645[0x4],true);var Width=Window_GetElement(window,OxO3645[0x5],true);var Height=Window_GetElement(window,OxO3645[0x6],true);var AutoStart=Window_GetElement(window,OxO3645[0x7],true);var ShowControls=Window_GetElement(window,OxO3645[0x8],true);var ShowStatusBar=Window_GetElement(window,OxO3645[0x9],true);var WindowlessVideo=Window_GetElement(window,OxO3645[0xa],true);var TargetUrl=Window_GetElement(window,OxO3645[0xb],true);var Button1=Window_GetElement(window,OxO3645[0xc],true);var Button2=Window_GetElement(window,OxO3645[0xd],true);var btn_zoom_in=Window_GetElement(window,OxO3645[0xe],true);var btn_zoom_out=Window_GetElement(window,OxO3645[0xf],true);var btn_Actualsize=Window_GetElement(window,OxO3645[0x10],true);var CreateDir=Window_GetElement(window,OxO3645[0x11],true);var upload_image=Window_GetElement(window,OxO3645[0x12],true); upload_image=upload_image[OxO3645[0x13]] ;var editor=Window_GetDialogArguments(window); do_preview() ; function do_preview(){var Ox33e;var Oxf4;var Ox2b9;if(TargetUrl[OxO3645[0x0]]==OxO3645[0x1]){return ;} ;var Ox33f,Ox340,Ox341;if(AutoStart[OxO3645[0x14]]){ Ox33f=0x1 ;} else { Ox33f=0x0 ;} ;if(ShowStatusBar[OxO3645[0x14]]){ Ox340=0x1 ;} else { Ox340=0x0 ;} ;if(ShowControls[OxO3645[0x14]]){ Ox341=0x1 ;} else { Ox341=0x0 ;} ;if(WindowlessVideo[OxO3645[0x14]]){ windowlessvideo=true ;} else { windowlessvideo=false ;} ; Oxf4=Width[OxO3645[0x0]]+OxO3645[0x1] ; Ox2b9=Height[OxO3645[0x0]]+OxO3645[0x1] ; Oxf4=parseInt(Oxf4) ; Ox2b9=parseInt(Ox2b9) ;var Ox312=OxO3645[0x15]+TargetUrl[OxO3645[0x0]]+OxO3645[0x16]+Ox33f+OxO3645[0x17]+Ox341+OxO3645[0x18]+windowlessvideo+OxO3645[0x19]+Ox340+OxO3645[0x1a]+Oxf4+OxO3645[0x1b]+Ox2b9+OxO3645[0x1c];var Ox2f0=OxO3645[0x1d]+OxO3645[0x1e]+OxO3645[0x1f]+OxO3645[0x20]+OxO3645[0x21]+OxO3645[0x22]+Ox2b9+OxO3645[0x1a]+Oxf4+OxO3645[0x23]; Ox2f0=Ox2f0+OxO3645[0x24]+TargetUrl[OxO3645[0x0]]+OxO3645[0x25] ; Ox2f0=Ox2f0+OxO3645[0x26]+Ox33f+OxO3645[0x25] ; Ox2f0=Ox2f0+OxO3645[0x27]+Ox341+OxO3645[0x25] ; Ox2f0=Ox2f0+OxO3645[0x28]+Ox340+OxO3645[0x25] ; Ox2f0=Ox2f0+OxO3645[0x29]+windowlessvideo+OxO3645[0x25] ; Ox2f0=Ox2f0+Ox312+OxO3645[0x2a] ; Ox312=Ox2f0 ; divpreview[OxO3645[0x2b]]=Ox312 ;}  ; window[OxO3645[0x2d]]=window[OxO3645[0x2c]]=function (){ divpreview[OxO3645[0x2b]]=OxO3645[0x1] ;}  ;var parameters= new Array(); function do_insert(){ divpreview[OxO3645[0x2b]]=OxO3645[0x1] ;if(TargetUrl[OxO3645[0x0]]==OxO3645[0x1]){ alert(OxO3645[0x2e]) ;return false;} ;var Ox33f,Ox340,Ox341;if(AutoStart[OxO3645[0x14]]){ Ox33f=0x1 ;} else { Ox33f=0x0 ;} ;if(ShowStatusBar[OxO3645[0x14]]){ Ox340=0x1 ;} else { Ox340=0x0 ;} ;if(ShowControls[OxO3645[0x14]]){ Ox341=0x1 ;} else { Ox341=0x0 ;} ;if(WindowlessVideo[OxO3645[0x14]]){ windowlessvideo=true ;} else { windowlessvideo=false ;} ; width=Width[OxO3645[0x0]]+OxO3645[0x1] ; height=Height[OxO3645[0x0]]+OxO3645[0x1] ; width=parseInt(width) ; height=parseInt(height) ;var Ox312=OxO3645[0x15]+TargetUrl[OxO3645[0x0]]+OxO3645[0x16]+Ox33f+OxO3645[0x17]+Ox341+OxO3645[0x19]+Ox340+OxO3645[0x2f]+windowlessvideo+OxO3645[0x1a]+width+OxO3645[0x1b]+height+OxO3645[0x1c];var Ox2f0=OxO3645[0x1d]+OxO3645[0x1e]+OxO3645[0x1f]+OxO3645[0x20]+OxO3645[0x21]+OxO3645[0x22]+height+OxO3645[0x1a]+width+OxO3645[0x23]; Ox2f0=Ox2f0+OxO3645[0x24]+TargetUrl[OxO3645[0x0]]+OxO3645[0x25] ; Ox2f0=Ox2f0+OxO3645[0x26]+Ox33f+OxO3645[0x25] ; Ox2f0=Ox2f0+OxO3645[0x27]+Ox341+OxO3645[0x25] ; Ox2f0=Ox2f0+OxO3645[0x28]+Ox340+OxO3645[0x25] ; Ox2f0=Ox2f0+OxO3645[0x29]+windowlessvideo+OxO3645[0x25] ; Ox2f0=Ox2f0+Ox312+OxO3645[0x2a] ; Ox312=Ox2f0 ; editor.PasteHTML(Ox312) ; Window_CloseDialog(window) ;}  ; function do_Close(){ divpreview[OxO3645[0x2b]]=OxO3645[0x1] ; Window_CloseDialog(window) ;}  ; function Zoom_In(){if(divpreview[OxO3645[0x31]][OxO3645[0x30]]!=0x0){ divpreview[OxO3645[0x31]][OxO3645[0x30]]*=1.2 ;} else { divpreview[OxO3645[0x31]][OxO3645[0x30]]=1.2 ;} ;}  ; function Zoom_Out(){if(divpreview[OxO3645[0x31]][OxO3645[0x30]]!=0x0){ divpreview[OxO3645[0x31]][OxO3645[0x30]]*=0.8 ;} else { divpreview[OxO3645[0x31]][OxO3645[0x30]]=0.8 ;} ;}  ; function Actualsize(){ divpreview[OxO3645[0x31]][OxO3645[0x30]]=0x1 ; do_preview() ;}  ;if(Browser_IsIE7()){var _dialogPromptID=null; function IEprompt(Ox18,Ox174,Ox175){ that=this ; this[OxO3645[0x32]]=function (Ox176){ val=document.getElementById(OxO3645[0x33])[OxO3645[0x0]] ; _dialogPromptID[OxO3645[0x31]][OxO3645[0x34]]=OxO3645[0x35] ; document.getElementById(OxO3645[0x33])[OxO3645[0x0]]=OxO3645[0x1] ;if(Ox176){ val=OxO3645[0x1] ;} ; Ox18(val) ;return false;}  ;if(Ox175==undefined){ Ox175=OxO3645[0x1] ;} ;if(_dialogPromptID==null){var Ox177=document.getElementsByTagName(OxO3645[0x36])[0x0]; tnode=document.createElement(OxO3645[0x37]) ; tnode[OxO3645[0x38]]=OxO3645[0x39] ; Ox177.appendChild(tnode) ; _dialogPromptID=document.getElementById(OxO3645[0x39]) ; tnode=document.createElement(OxO3645[0x37]) ; tnode[OxO3645[0x38]]=OxO3645[0x3a] ; Ox177.appendChild(tnode) ; _dialogPromptID[OxO3645[0x31]][OxO3645[0x3b]]=OxO3645[0x3c] ; _dialogPromptID[OxO3645[0x31]][OxO3645[0x3d]]=OxO3645[0x3e] ; _dialogPromptID[OxO3645[0x31]][OxO3645[0x3f]]=OxO3645[0x40] ; _dialogPromptID[OxO3645[0x31]][OxO3645[0x41]]=OxO3645[0x42] ; _dialogPromptID[OxO3645[0x31]][OxO3645[0x43]]=OxO3645[0x44] ;} ;var Ox178=OxO3645[0x45]+InputRequired+OxO3645[0x46]; Ox178+=OxO3645[0x47]+Ox174+OxO3645[0x48] ; Ox178+=OxO3645[0x49] ; Ox178+=OxO3645[0x4a]+Ox175+OxO3645[0x4b] ; Ox178+=OxO3645[0x4c] ; Ox178+=OxO3645[0x4d]+OK+OxO3645[0x4e] ; Ox178+=OxO3645[0x4f] ; Ox178+=OxO3645[0x50]+Cancel+OxO3645[0x51] ; Ox178+=OxO3645[0x52] ; _dialogPromptID[OxO3645[0x2b]]=Ox178 ; _dialogPromptID[OxO3645[0x31]][OxO3645[0x53]]=OxO3645[0x54] ; _dialogPromptID[OxO3645[0x31]][OxO3645[0x56]]=parseInt((document[OxO3645[0x36]][OxO3645[0x55]]-0x13b)/0x2)+OxO3645[0x57] ; _dialogPromptID[OxO3645[0x31]][OxO3645[0x34]]=OxO3645[0x58] ;var Ox179=document.getElementById(OxO3645[0x33]);try{var Ox17a=Ox179.createTextRange(); Ox17a.collapse(false) ; Ox17a.select() ;} catch(x){ Ox179.focus() ;} ;}  ;} ;if(!Browser_IsWinIE()){ btn_zoom_in[OxO3645[0x31]][OxO3645[0x34]]=btn_zoom_out[OxO3645[0x31]][OxO3645[0x34]]=btn_Actualsize[OxO3645[0x31]][OxO3645[0x34]]=OxO3645[0x35] ;} else {} ;if(CreateDir){ CreateDir[OxO3645[0x5a]]= new Function(OxO3645[0x59]) ;} ;if(btn_zoom_in){ btn_zoom_in[OxO3645[0x5a]]= new Function(OxO3645[0x59]) ;} ;if(btn_zoom_out){ btn_zoom_out[OxO3645[0x5a]]= new Function(OxO3645[0x59]) ;} ;if(btn_Actualsize){ btn_Actualsize[OxO3645[0x5a]]= new Function(OxO3645[0x59]) ;} ;