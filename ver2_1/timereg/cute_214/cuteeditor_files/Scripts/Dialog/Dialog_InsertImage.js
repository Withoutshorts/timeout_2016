var OxO40ba=["zoomcount","wheelDelta","zoom","style","0%","top","value","","onload","upload_image","contentWindow","browse_Frame","height","250px","btn_CreateDir","btn_zoom_in","btn_zoom_out","btn_bestfit","btn_Actualsize","divpreview","img_demo","Align","Border","bordercolor","bordercolor_Preview","inp_width","imgLock","inp_height","constrain_prop","HSpace","VSpace","TargetUrl","AlternateText","inp_id","longDesc","fieldsetUpload","Button1","Button2","img","editor","src","src_cetemp","width","id","vspace","hspace","border","borderColor"," ","backgroundColor","align","alt","file","complete","../images/1x1.gif","?","\x26time=","?time=","0","onmousewheel","Edit","longdesc","Code","parentNode","=\x22","\x22","checked","../Images/locked.gif","../Images/1x1.gif","length","onclick","display","none","wrapupPrompt","iepromptfield","body","div","IEPromptBox","promptBlackout","1px solid #b0bec7","#f0f0f0","position","absolute","330px","zIndex","100","\x3Cdiv style=\x22width: 100%; padding-top:3px;background-color: #DCE7EB; font-family: verdana; font-size: 10pt; font-weight: bold; height: 22px; text-align:center; background:url(../Images/formbg2.gif) repeat-x left top;\x22\x3E","\x3C/div\x3E","\x3Cdiv style=\x22padding: 10px\x22\x3E","\x3CBR\x3E\x3CBR\x3E","\x3Cform action=\x22\x22 onsubmit=\x22return wrapupPrompt()\x22\x3E","\x3Cinput id=\x22iepromptfield\x22 name=\x22iepromptdata\x22 type=text size=46 value=\x22","\x22\x3E","\x3Cbr\x3E\x3Cbr\x3E\x3Ccenter\x3E","\x3Cinput type=\x22submit\x22 value=\x22\x26nbsp;\x26nbsp;\x26nbsp;","\x26nbsp;\x26nbsp;\x26nbsp;\x22\x3E","\x26nbsp;\x26nbsp;\x26nbsp;\x26nbsp;\x26nbsp;\x26nbsp;","\x3Cinput type=\x22button\x22 onclick=\x22wrapupPrompt(true)\x22 value=\x22\x26nbsp;","\x26nbsp;\x22\x3E","\x3C/form\x3E\x3C/div\x3E","innerHTML","100px","left","offsetWidth","px","block","onmouseover","CuteEditor_ColorPicker_ButtonOver(this)"];function OnImageMouseWheel(){var Ox234=Event_GetEvent();var img=Event_GetSrcElement(Ox234);var Ox2ec=img[OxO40ba[0]]||3;if(Ox234[OxO40ba[1]]>=106){Ox2ec++;} else {if(Ox234[OxO40ba[1]]<=-106){Ox2ec--;} ;} ;img[OxO40ba[0]]=Ox2ec;img[OxO40ba[3]][OxO40ba[2]]=Ox2ec+OxO40ba[4];return false;} ;function Window_GetDialogTop(Ox95){return Ox95[OxO40ba[5]];} ;function row_click(Ox2ac){TargetUrl[OxO40ba[6]]=Ox2ac;Actualsize();} ;function do_preview(){} ;function reset_hiddens(){if(TargetUrl[OxO40ba[6]]!=OxO40ba[7]&&TargetUrl[OxO40ba[6]]!=null){do_preview();} ;} ;Event_Attach(window,OxO40ba[8],reset_hiddens);function RequireFileBrowseScript(){} ;function Actualsize(){} ;RequireFileBrowseScript();var upload_image=Window_GetElement(window,OxO40ba[9],true);upload_image=upload_image[OxO40ba[10]];var browse_Frame=Window_GetElement(window,OxO40ba[11],true);if(!Browser_IsWinIE()){browse_Frame[OxO40ba[3]][OxO40ba[12]]=OxO40ba[13];} ;browse_Frame=browse_Frame[OxO40ba[10]];var btn_CreateDir=Window_GetElement(window,OxO40ba[14],true);var btn_zoom_in=Window_GetElement(window,OxO40ba[15],true);var btn_zoom_out=Window_GetElement(window,OxO40ba[16],true);var btn_bestfit=Window_GetElement(window,OxO40ba[17],true);var btn_Actualsize=Window_GetElement(window,OxO40ba[18],true);var divpreview=Window_GetElement(window,OxO40ba[19],true);var img_demo=Window_GetElement(window,OxO40ba[20],true);var Align=Window_GetElement(window,OxO40ba[21],true);var Border=Window_GetElement(window,OxO40ba[22],true);var bordercolor=Window_GetElement(window,OxO40ba[23],true);var bordercolor_Preview=Window_GetElement(window,OxO40ba[24],true);var inp_width=Window_GetElement(window,OxO40ba[25],true);var imgLock=Window_GetElement(window,OxO40ba[26],true);var inp_height=Window_GetElement(window,OxO40ba[27],true);var constrain_prop=Window_GetElement(window,OxO40ba[28],true);var HSpace=Window_GetElement(window,OxO40ba[29],true);var VSpace=Window_GetElement(window,OxO40ba[30],true);var TargetUrl=Window_GetElement(window,OxO40ba[31],true);var AlternateText=Window_GetElement(window,OxO40ba[32],true);var inp_id=Window_GetElement(window,OxO40ba[33],true);var longDesc=Window_GetElement(window,OxO40ba[34],true);var fieldsetUpload=Window_GetElement(window,OxO40ba[35],true);var Button1=Window_GetElement(window,OxO40ba[36],true);var Button2=Window_GetElement(window,OxO40ba[37],true);var btn_zoom_in=Window_GetElement(window,OxO40ba[15],true);var btn_zoom_out=Window_GetElement(window,OxO40ba[16],true);var btn_Actualsize=Window_GetElement(window,OxO40ba[18],true);var btn_bestfit=Window_GetElement(window,OxO40ba[17],true);var CreateDir=Window_GetElement(window,OxO40ba[14],true);var obj=Window_GetDialogArguments(window);var element=obj[OxO40ba[38]];var editor=obj[OxO40ba[39]];var src=OxO40ba[7];if(element.getAttribute(OxO40ba[40])){src=element.getAttribute(OxO40ba[40]);} ;if(element.getAttribute(OxO40ba[41])){src=element.getAttribute(OxO40ba[41]);} ;if(element&&src){TargetUrl[OxO40ba[6]]=src;} ;inp_width[OxO40ba[6]]=element[OxO40ba[42]]||OxO40ba[7];inp_height[OxO40ba[6]]=element[OxO40ba[12]]||OxO40ba[7];inp_id[OxO40ba[6]]=element[OxO40ba[43]]||OxO40ba[7];if(element[OxO40ba[44]]<=0){VSpace[OxO40ba[6]]=OxO40ba[7];} else {VSpace[OxO40ba[6]]=element[OxO40ba[44]];} ;if(element[OxO40ba[45]]<=0){HSpace[OxO40ba[6]]=OxO40ba[7];} else {HSpace[OxO40ba[6]]=element[OxO40ba[45]];} ;Border[OxO40ba[6]]=element[OxO40ba[46]]||OxO40ba[7];if(Browser_IsWinIE()){bordercolor[OxO40ba[6]]=element[OxO40ba[3]][OxO40ba[47]];} else {var arr=revertColor(element[OxO40ba[3]].borderColor).split(OxO40ba[48]);bordercolor[OxO40ba[6]]=arr[0];} ;bordercolor[OxO40ba[3]][OxO40ba[49]]=bordercolor[OxO40ba[6]]||OxO40ba[7];bordercolor_Preview[OxO40ba[3]][OxO40ba[49]]=bordercolor[OxO40ba[6]]||OxO40ba[7];Align[OxO40ba[6]]=element[OxO40ba[50]]||OxO40ba[7];AlternateText[OxO40ba[6]]=element[OxO40ba[51]]||OxO40ba[7];longDesc[OxO40ba[6]]=element[OxO40ba[34]]||OxO40ba[7];var sCheckFlag=OxO40ba[52];function ResetFields(){TargetUrl[OxO40ba[6]]=OxO40ba[7];inp_width[OxO40ba[6]]=OxO40ba[7];inp_height[OxO40ba[6]]=OxO40ba[7];inp_id[OxO40ba[6]]=OxO40ba[7];VSpace[OxO40ba[6]]=OxO40ba[7];HSpace[OxO40ba[6]]=OxO40ba[7];Border[OxO40ba[6]]=OxO40ba[7];bordercolor[OxO40ba[6]]=OxO40ba[7];bordercolor[OxO40ba[3]][OxO40ba[49]]=OxO40ba[7];Align[OxO40ba[6]]=OxO40ba[7];AlternateText[OxO40ba[6]]=OxO40ba[7];longDesc[OxO40ba[6]]=OxO40ba[7];} ;function do_preview(){var Ox127=TargetUrl[OxO40ba[6]];if(Ox127==null){TargetUrl[OxO40ba[6]]=OxO40ba[7];Ox127==OxO40ba[7];} ;if(Ox127!=null&&Ox127!=OxO40ba[7]){var Ox2fa;var Ox2fb;var Ox2fa= new Image();Ox2fa[OxO40ba[40]]=Ox127;function Ox2fc(){if(Ox2fa[OxO40ba[53]]){window.clearInterval(Ox2fb);var Ox2fd= new Date();var Ox2fe=Ox2fd.getTime();if(Ox127==OxO40ba[7]){Ox127=OxO40ba[54];} ;if(Ox127.indexOf(OxO40ba[55])!=-1){Ox127=Ox127+OxO40ba[56]+Ox2fe;} else {Ox127=Ox127+OxO40ba[57]+Ox2fe;} ;if(inp_width[OxO40ba[6]]==OxO40ba[58]||inp_width[OxO40ba[6]]==OxO40ba[7]){inp_width[OxO40ba[6]]=Ox2fa[OxO40ba[42]];inp_height[OxO40ba[6]]=Ox2fa[OxO40ba[12]];} ;img_demo[OxO40ba[40]]=Ox127;if(Browser_IsWinIE()){Event_Attach(img_demo,OxO40ba[59],OnImageMouseWheel);} ;img_demo[OxO40ba[51]]=AlternateText[OxO40ba[6]];img_demo[OxO40ba[50]]=Align[OxO40ba[6]];img_demo[OxO40ba[42]]=inp_width[OxO40ba[6]];img_demo[OxO40ba[12]]=inp_height[OxO40ba[6]];img_demo[OxO40ba[44]]=VSpace[OxO40ba[6]];img_demo[OxO40ba[45]]=HSpace[OxO40ba[6]];if(parseInt(Border.value)>0){img_demo[OxO40ba[46]]=Border[OxO40ba[6]];} ;if(bordercolor[OxO40ba[6]]!=OxO40ba[7]){img_demo[OxO40ba[3]][OxO40ba[47]]=bordercolor[OxO40ba[6]];} ;Ox127=Ox127.toLowerCase();} ;} ;Ox2fb=window.setInterval(Ox2fc,100);} ;} ;function do_insert(){var img=element;img[OxO40ba[40]]=TargetUrl[OxO40ba[6]];if(editor.GetActiveTab()==OxO40ba[60]){img.setAttribute(OxO40ba[41],TargetUrl.value);} ;img[OxO40ba[42]]=inp_width[OxO40ba[6]];img[OxO40ba[12]]=inp_height[OxO40ba[6]];if(img[OxO40ba[3]][OxO40ba[42]]||img[OxO40ba[3]][OxO40ba[12]]){img[OxO40ba[3]][OxO40ba[42]]=inp_width[OxO40ba[6]];img[OxO40ba[3]][OxO40ba[12]]=inp_height[OxO40ba[6]];} ;img[OxO40ba[44]]=VSpace[OxO40ba[6]];img[OxO40ba[45]]=HSpace[OxO40ba[6]];img[OxO40ba[46]]=Border[OxO40ba[6]];var Ox27b=/[^a-z\d]/i;if(Ox27b.test(inp_id.value)){alert(ValidID);return ;} ;img[OxO40ba[43]]=inp_id[OxO40ba[6]];try{img[OxO40ba[3]][OxO40ba[47]]=bordercolor[OxO40ba[6]];} catch(er){alert(ValidColor);return false;} ;img[OxO40ba[50]]=Align[OxO40ba[6]];img[OxO40ba[51]]=AlternateText[OxO40ba[6]]||OxO40ba[7];img[OxO40ba[34]]=longDesc[OxO40ba[6]];if(TargetUrl[OxO40ba[6]]==OxO40ba[7]){alert(SelectImagetoInsert);return false;} ;if(img[OxO40ba[42]]==0){img.removeAttribute(OxO40ba[42]);} ;if(img[OxO40ba[12]]==0){img.removeAttribute(OxO40ba[12]);} ;if(img[OxO40ba[61]]==OxO40ba[7]||img[OxO40ba[61]]==null){img.removeAttribute(OxO40ba[61]);} ;if(img[OxO40ba[34]]==OxO40ba[7]||img[OxO40ba[34]]==null){img.removeAttribute(OxO40ba[34]);} ;if(img[OxO40ba[45]]==OxO40ba[7]){img.removeAttribute(OxO40ba[45]);} ;if(img[OxO40ba[44]]==OxO40ba[7]){img.removeAttribute(OxO40ba[44]);} ;if(img[OxO40ba[43]]==OxO40ba[7]){img.removeAttribute(OxO40ba[43]);} ;if(img[OxO40ba[46]]==OxO40ba[7]){img.removeAttribute(OxO40ba[46]);} ;if(img[OxO40ba[50]]==OxO40ba[7]){img.removeAttribute(OxO40ba[50]);} ;if(editor.GetActiveTab()==OxO40ba[62]){if(Browser_IsWinIE()){editor.PasteHTML(img.outerHTML);} else {editor.PasteHTML(outerHTML(img));} ;} else {if(!img[OxO40ba[63]]){editor.InsertElement(img);} ;} ;Window_CloseDialog(window);} ;function attr(name,Ox4f){if(!Ox4f||Ox4f==OxO40ba[7]){return OxO40ba[7];} ;return OxO40ba[48]+name+OxO40ba[64]+Ox4f+OxO40ba[65];} ;function do_Close(){Window_CloseDialog(window);} ;function Zoom_In(){if(Browser_IsWinIE()){if(divpreview[OxO40ba[3]][OxO40ba[2]]!=0){divpreview[OxO40ba[3]][OxO40ba[2]]*=1.2;} else {divpreview[OxO40ba[3]][OxO40ba[2]]=1.2;} ;} ;} ;function Zoom_Out(){if(Browser_IsWinIE()){if(divpreview[OxO40ba[3]][OxO40ba[2]]!=0){divpreview[OxO40ba[3]][OxO40ba[2]]*=0.8;} else {divpreview[OxO40ba[3]][OxO40ba[2]]=0.8;} ;} ;} ;function BestFit(){var img=img_demo;if(!img){return ;} ;var Ox266=280;var Ox14=290;if(Browser_IsWinIE()){divpreview[OxO40ba[3]][OxO40ba[2]]=1/Math.max(img[OxO40ba[42]]/Ox14,img[OxO40ba[12]]/Ox266);} ;} ;function Actualsize(){inp_width[OxO40ba[6]]=OxO40ba[7];inp_height[OxO40ba[6]]=OxO40ba[7];do_preview();if(Browser_IsWinIE()){divpreview[OxO40ba[3]][OxO40ba[2]]=1;} ;} ;function toggleConstrains(){if(constrain_prop[OxO40ba[66]]){imgLock[OxO40ba[40]]=OxO40ba[67];checkConstrains(OxO40ba[42]);} else {imgLock[OxO40ba[40]]=OxO40ba[68];} ;} ;var checkingConstrains=false;function checkConstrains(Ox302){if(checkingConstrains){return ;} ;checkingConstrains=true;try{if(constrain_prop[OxO40ba[66]]){var Ox2e9= new Image();Ox2e9[OxO40ba[40]]=TargetUrl[OxO40ba[6]];var Ox303=Ox2e9[OxO40ba[42]];var Ox304=Ox2e9[OxO40ba[12]];if((Ox303>0)&&(Ox304>0)){var Ox14=inp_width[OxO40ba[6]];var Ox266=inp_height[OxO40ba[6]];if(Ox302==OxO40ba[42]){if(Ox14[OxO40ba[69]]==0||isNaN(Ox14)){inp_width[OxO40ba[6]]=OxO40ba[7];inp_height[OxO40ba[6]]=OxO40ba[7];} else {Ox266=parseInt(Ox14*Ox304/Ox303);inp_height[OxO40ba[6]]=Ox266;} ;} ;if(Ox302==OxO40ba[12]){if(Ox266[OxO40ba[69]]==0||isNaN(Ox266)){inp_width[OxO40ba[6]]=OxO40ba[7];inp_height[OxO40ba[6]]=OxO40ba[7];} else {Ox14=parseInt(Ox266*Ox303/Ox304);inp_width[OxO40ba[6]]=Ox14;} ;} ;} ;} ;do_preview();} finally{checkingConstrains=false;} ;} ;function ImageEditor(){} ;bordercolor[OxO40ba[70]]=bordercolor_Preview[OxO40ba[70]]=function bordercolor_onclick(){SelectColor(bordercolor,bordercolor_Preview);} ;if(!Browser_IsWinIE()){btn_zoom_in[OxO40ba[3]][OxO40ba[71]]=btn_zoom_out[OxO40ba[3]][OxO40ba[71]]=btn_bestfit[OxO40ba[3]][OxO40ba[71]]=btn_Actualsize[OxO40ba[3]][OxO40ba[71]]=OxO40ba[72];} ;if(Browser_IsIE7()){var _dialogPromptID=null;function IEprompt(Ox112,Ox113,Ox114){that=this;this[OxO40ba[73]]=function (Ox115){val=document.getElementById(OxO40ba[74])[OxO40ba[6]];_dialogPromptID[OxO40ba[3]][OxO40ba[71]]=OxO40ba[72];document.getElementById(OxO40ba[74])[OxO40ba[6]]=OxO40ba[7];if(Ox115){val=OxO40ba[7];} ;Ox112(val);return false;} ;if(Ox114==undefined){Ox114=OxO40ba[7];} ;if(_dialogPromptID==null){var Ox116=document.getElementsByTagName(OxO40ba[75])[0];tnode=document.createElement(OxO40ba[76]);tnode[OxO40ba[43]]=OxO40ba[77];Ox116.appendChild(tnode);_dialogPromptID=document.getElementById(OxO40ba[77]);tnode=document.createElement(OxO40ba[76]);tnode[OxO40ba[43]]=OxO40ba[78];Ox116.appendChild(tnode);_dialogPromptID[OxO40ba[3]][OxO40ba[46]]=OxO40ba[79];_dialogPromptID[OxO40ba[3]][OxO40ba[49]]=OxO40ba[80];_dialogPromptID[OxO40ba[3]][OxO40ba[81]]=OxO40ba[82];_dialogPromptID[OxO40ba[3]][OxO40ba[42]]=OxO40ba[83];_dialogPromptID[OxO40ba[3]][OxO40ba[84]]=OxO40ba[85];} ;var Ox117=OxO40ba[86]+InputRequired+OxO40ba[87];Ox117+=OxO40ba[88]+Ox113+OxO40ba[89];Ox117+=OxO40ba[90];Ox117+=OxO40ba[91]+Ox114+OxO40ba[92];Ox117+=OxO40ba[93];Ox117+=OxO40ba[94]+OK+OxO40ba[95];Ox117+=OxO40ba[96];Ox117+=OxO40ba[97]+Cancel+OxO40ba[98];Ox117+=OxO40ba[99];_dialogPromptID[OxO40ba[100]]=Ox117;_dialogPromptID[OxO40ba[3]][OxO40ba[5]]=OxO40ba[101];_dialogPromptID[OxO40ba[3]][OxO40ba[102]]=parseInt((document[OxO40ba[75]][OxO40ba[103]]-315)/2)+OxO40ba[104];_dialogPromptID[OxO40ba[3]][OxO40ba[71]]=OxO40ba[105];var Ox118=document.getElementById(OxO40ba[74]);try{var Ox119=Ox118.createTextRange();Ox119.collapse(false);Ox119.select();} catch(x){Ox118.focus();} ;} ;} ;if(CreateDir){CreateDir[OxO40ba[106]]= new Function(OxO40ba[107]);} ;if(btn_zoom_in){btn_zoom_in[OxO40ba[106]]= new Function(OxO40ba[107]);} ;if(btn_zoom_out){btn_zoom_out[OxO40ba[106]]= new Function(OxO40ba[107]);} ;if(btn_Actualsize){btn_Actualsize[OxO40ba[106]]= new Function(OxO40ba[107]);} ;if(btn_bestfit){btn_bestfit[OxO40ba[106]]= new Function(OxO40ba[107]);} ;