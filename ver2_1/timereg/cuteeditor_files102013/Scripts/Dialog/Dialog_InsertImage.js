var OxO3380=["zoomcount","wheelDelta","zoom","style","0%","top","value","","onload","upload_image","contentWindow","browse_Frame","height","250px","btn_CreateDir","btn_zoom_in","btn_zoom_out","btn_bestfit","btn_Actualsize","divpreview","img_demo","Align","Border","bordercolor","bordercolor_Preview","inp_width","imgLock","inp_height","constrain_prop","HSpace","VSpace","TargetUrl","AlternateText","inp_id","longDesc","fieldsetUpload","Button1","Button2","img","editor","src","src_cetemp","width","id","vspace","hspace","border","borderColor"," ","backgroundColor","align","alt","file","complete","../images/1x1.gif","?","\x26time=","?time=","0","onmousewheel","Edit","longdesc","Code","parentNode","=\x22","\x22","checked","../Images/locked.gif","../Images/1x1.gif","length","onclick","display","none","wrapupPrompt","iepromptfield","body","div","IEPromptBox","promptBlackout","1px solid #b0bec7","#f0f0f0","position","absolute","330px","zIndex","100","\x3Cdiv style=\x22width: 100%; padding-top:3px;background-color: #DCE7EB; font-family: verdana; font-size: 10pt; font-weight: bold; height: 22px; text-align:center; background:url(../Images/formbg2.gif) repeat-x left top;\x22\x3E","\x3C/div\x3E","\x3Cdiv style=\x22padding: 10px\x22\x3E","\x3CBR\x3E\x3CBR\x3E","\x3Cform action=\x22\x22 onsubmit=\x22return wrapupPrompt()\x22\x3E","\x3Cinput id=\x22iepromptfield\x22 name=\x22iepromptdata\x22 type=text size=46 value=\x22","\x22\x3E","\x3Cbr\x3E\x3Cbr\x3E\x3Ccenter\x3E","\x3Cinput type=\x22submit\x22 value=\x22\x26nbsp;\x26nbsp;\x26nbsp;","\x26nbsp;\x26nbsp;\x26nbsp;\x22\x3E","\x26nbsp;\x26nbsp;\x26nbsp;\x26nbsp;\x26nbsp;\x26nbsp;","\x3Cinput type=\x22button\x22 onclick=\x22wrapupPrompt(true)\x22 value=\x22\x26nbsp;","\x26nbsp;\x22\x3E","\x3C/form\x3E\x3C/div\x3E","innerHTML","100px","left","offsetWidth","px","block","onmouseover","CuteEditor_ColorPicker_ButtonOver(this)"];function OnImageMouseWheel(){var Ox233=Event_GetEvent();var img=Event_GetSrcElement(Ox233);var Ox2eb=img[OxO3380[0]]||3;if(Ox233[OxO3380[1]]>=106){Ox2eb++;} else {if(Ox233[OxO3380[1]]<=-106){Ox2eb--;} ;} ;img[OxO3380[0]]=Ox2eb;img[OxO3380[3]][OxO3380[2]]=Ox2eb+OxO3380[4];return false;} ;function Window_GetDialogTop(Ox95){return Ox95[OxO3380[5]];} ;function row_click(Ox2ab){TargetUrl[OxO3380[6]]=Ox2ab;Actualsize();} ;function do_preview(){} ;function reset_hiddens(){if(TargetUrl[OxO3380[6]]!=OxO3380[7]&&TargetUrl[OxO3380[6]]!=null){do_preview();} ;} ;Event_Attach(window,OxO3380[8],reset_hiddens);function RequireFileBrowseScript(){} ;function Actualsize(){} ;RequireFileBrowseScript();var upload_image=Window_GetElement(window,OxO3380[9],true);upload_image=upload_image[OxO3380[10]];var browse_Frame=Window_GetElement(window,OxO3380[11],true);if(!Browser_IsWinIE()){browse_Frame[OxO3380[3]][OxO3380[12]]=OxO3380[13];} ;browse_Frame=browse_Frame[OxO3380[10]];var btn_CreateDir=Window_GetElement(window,OxO3380[14],true);var btn_zoom_in=Window_GetElement(window,OxO3380[15],true);var btn_zoom_out=Window_GetElement(window,OxO3380[16],true);var btn_bestfit=Window_GetElement(window,OxO3380[17],true);var btn_Actualsize=Window_GetElement(window,OxO3380[18],true);var divpreview=Window_GetElement(window,OxO3380[19],true);var img_demo=Window_GetElement(window,OxO3380[20],true);var Align=Window_GetElement(window,OxO3380[21],true);var Border=Window_GetElement(window,OxO3380[22],true);var bordercolor=Window_GetElement(window,OxO3380[23],true);var bordercolor_Preview=Window_GetElement(window,OxO3380[24],true);var inp_width=Window_GetElement(window,OxO3380[25],true);var imgLock=Window_GetElement(window,OxO3380[26],true);var inp_height=Window_GetElement(window,OxO3380[27],true);var constrain_prop=Window_GetElement(window,OxO3380[28],true);var HSpace=Window_GetElement(window,OxO3380[29],true);var VSpace=Window_GetElement(window,OxO3380[30],true);var TargetUrl=Window_GetElement(window,OxO3380[31],true);var AlternateText=Window_GetElement(window,OxO3380[32],true);var inp_id=Window_GetElement(window,OxO3380[33],true);var longDesc=Window_GetElement(window,OxO3380[34],true);var fieldsetUpload=Window_GetElement(window,OxO3380[35],true);var Button1=Window_GetElement(window,OxO3380[36],true);var Button2=Window_GetElement(window,OxO3380[37],true);var btn_zoom_in=Window_GetElement(window,OxO3380[15],true);var btn_zoom_out=Window_GetElement(window,OxO3380[16],true);var btn_Actualsize=Window_GetElement(window,OxO3380[18],true);var btn_bestfit=Window_GetElement(window,OxO3380[17],true);var CreateDir=Window_GetElement(window,OxO3380[14],true);var obj=Window_GetDialogArguments(window);var element=obj[OxO3380[38]];var editor=obj[OxO3380[39]];var src=OxO3380[7];if(element.getAttribute(OxO3380[40])){src=element.getAttribute(OxO3380[40]);} ;if(element.getAttribute(OxO3380[41])){src=element.getAttribute(OxO3380[41]);} ;if(element&&src){TargetUrl[OxO3380[6]]=src;} ;inp_width[OxO3380[6]]=element[OxO3380[42]]||OxO3380[7];inp_height[OxO3380[6]]=element[OxO3380[12]]||OxO3380[7];inp_id[OxO3380[6]]=element[OxO3380[43]]||OxO3380[7];if(element[OxO3380[44]]<=0){VSpace[OxO3380[6]]=OxO3380[7];} else {VSpace[OxO3380[6]]=element[OxO3380[44]];} ;if(element[OxO3380[45]]<=0){HSpace[OxO3380[6]]=OxO3380[7];} else {HSpace[OxO3380[6]]=element[OxO3380[45]];} ;Border[OxO3380[6]]=element[OxO3380[46]]||OxO3380[7];if(Browser_IsWinIE()){bordercolor[OxO3380[6]]=element[OxO3380[3]][OxO3380[47]];} else {var arr=revertColor(element[OxO3380[3]].borderColor).split(OxO3380[48]);bordercolor[OxO3380[6]]=arr[0];} ;bordercolor[OxO3380[3]][OxO3380[49]]=bordercolor[OxO3380[6]]||OxO3380[7];bordercolor_Preview[OxO3380[3]][OxO3380[49]]=bordercolor[OxO3380[6]]||OxO3380[7];Align[OxO3380[6]]=element[OxO3380[50]]||OxO3380[7];AlternateText[OxO3380[6]]=element[OxO3380[51]]||OxO3380[7];longDesc[OxO3380[6]]=element[OxO3380[34]]||OxO3380[7];var sCheckFlag=OxO3380[52];function ResetFields(){TargetUrl[OxO3380[6]]=OxO3380[7];inp_width[OxO3380[6]]=OxO3380[7];inp_height[OxO3380[6]]=OxO3380[7];inp_id[OxO3380[6]]=OxO3380[7];VSpace[OxO3380[6]]=OxO3380[7];HSpace[OxO3380[6]]=OxO3380[7];Border[OxO3380[6]]=OxO3380[7];bordercolor[OxO3380[6]]=OxO3380[7];bordercolor[OxO3380[3]][OxO3380[49]]=OxO3380[7];Align[OxO3380[6]]=OxO3380[7];AlternateText[OxO3380[6]]=OxO3380[7];longDesc[OxO3380[6]]=OxO3380[7];} ;function do_preview(){var Ox126=TargetUrl[OxO3380[6]];if(Ox126==null){TargetUrl[OxO3380[6]]=OxO3380[7];Ox126==OxO3380[7];} ;if(Ox126!=null&&Ox126!=OxO3380[7]){var Ox2f9;var Ox2fa;var Ox2f9= new Image();Ox2f9[OxO3380[40]]=Ox126;function Ox2fb(){if(Ox2f9[OxO3380[53]]){window.clearInterval(Ox2fa);var Ox2fc= new Date();var Ox2fd=Ox2fc.getTime();if(Ox126==OxO3380[7]){Ox126=OxO3380[54];} ;if(Ox126.indexOf(OxO3380[55])!=-1){Ox126=Ox126+OxO3380[56]+Ox2fd;} else {Ox126=Ox126+OxO3380[57]+Ox2fd;} ;if(inp_width[OxO3380[6]]==OxO3380[58]||inp_width[OxO3380[6]]==OxO3380[7]){inp_width[OxO3380[6]]=Ox2f9[OxO3380[42]];inp_height[OxO3380[6]]=Ox2f9[OxO3380[12]];} ;img_demo[OxO3380[40]]=Ox126;if(Browser_IsWinIE()){Event_Attach(img_demo,OxO3380[59],OnImageMouseWheel);} ;img_demo[OxO3380[51]]=AlternateText[OxO3380[6]];img_demo[OxO3380[50]]=Align[OxO3380[6]];img_demo[OxO3380[42]]=inp_width[OxO3380[6]];img_demo[OxO3380[12]]=inp_height[OxO3380[6]];img_demo[OxO3380[44]]=VSpace[OxO3380[6]];img_demo[OxO3380[45]]=HSpace[OxO3380[6]];if(parseInt(Border.value)>0){img_demo[OxO3380[46]]=Border[OxO3380[6]];} ;if(bordercolor[OxO3380[6]]!=OxO3380[7]){img_demo[OxO3380[3]][OxO3380[47]]=bordercolor[OxO3380[6]];} ;Ox126=Ox126.toLowerCase();} ;} ;Ox2fa=window.setInterval(Ox2fb,100);} ;} ;function do_insert(){var img=element;img[OxO3380[40]]=TargetUrl[OxO3380[6]];if(editor.GetActiveTab()==OxO3380[60]){img.setAttribute(OxO3380[41],TargetUrl.value);} ;img[OxO3380[42]]=inp_width[OxO3380[6]];img[OxO3380[12]]=inp_height[OxO3380[6]];if(img[OxO3380[3]][OxO3380[42]]||img[OxO3380[3]][OxO3380[12]]){img[OxO3380[3]][OxO3380[42]]=inp_width[OxO3380[6]];img[OxO3380[3]][OxO3380[12]]=inp_height[OxO3380[6]];} ;img[OxO3380[44]]=VSpace[OxO3380[6]];img[OxO3380[45]]=HSpace[OxO3380[6]];img[OxO3380[46]]=Border[OxO3380[6]];var Ox27a=/[^a-z\d]/i;if(Ox27a.test(inp_id.value)){alert(ValidID);return ;} ;img[OxO3380[43]]=inp_id[OxO3380[6]];try{img[OxO3380[3]][OxO3380[47]]=bordercolor[OxO3380[6]];} catch(er){alert(ValidColor);return false;} ;img[OxO3380[50]]=Align[OxO3380[6]];img[OxO3380[51]]=AlternateText[OxO3380[6]]||OxO3380[7];img[OxO3380[34]]=longDesc[OxO3380[6]];if(TargetUrl[OxO3380[6]]==OxO3380[7]){alert(SelectImagetoInsert);return false;} ;if(img[OxO3380[42]]==0){img.removeAttribute(OxO3380[42]);} ;if(img[OxO3380[12]]==0){img.removeAttribute(OxO3380[12]);} ;if(img[OxO3380[61]]==OxO3380[7]||img[OxO3380[61]]==null){img.removeAttribute(OxO3380[61]);} ;if(img[OxO3380[34]]==OxO3380[7]||img[OxO3380[34]]==null){img.removeAttribute(OxO3380[34]);} ;if(img[OxO3380[45]]==OxO3380[7]){img.removeAttribute(OxO3380[45]);} ;if(img[OxO3380[44]]==OxO3380[7]){img.removeAttribute(OxO3380[44]);} ;if(img[OxO3380[43]]==OxO3380[7]){img.removeAttribute(OxO3380[43]);} ;if(img[OxO3380[46]]==OxO3380[7]){img.removeAttribute(OxO3380[46]);} ;if(img[OxO3380[50]]==OxO3380[7]){img.removeAttribute(OxO3380[50]);} ;if(editor.GetActiveTab()==OxO3380[62]){if(Browser_IsWinIE()){editor.PasteHTML(img.outerHTML);} else {editor.PasteHTML(outerHTML(img));} ;} else {if(!img[OxO3380[63]]){editor.InsertElement(img);} ;} ;Window_CloseDialog(window);} ;function attr(name,Ox4f){if(!Ox4f||Ox4f==OxO3380[7]){return OxO3380[7];} ;return OxO3380[48]+name+OxO3380[64]+Ox4f+OxO3380[65];} ;function do_Close(){Window_CloseDialog(window);} ;function Zoom_In(){if(Browser_IsWinIE()){if(divpreview[OxO3380[3]][OxO3380[2]]!=0){divpreview[OxO3380[3]][OxO3380[2]]*=1.2;} else {divpreview[OxO3380[3]][OxO3380[2]]=1.2;} ;} ;} ;function Zoom_Out(){if(Browser_IsWinIE()){if(divpreview[OxO3380[3]][OxO3380[2]]!=0){divpreview[OxO3380[3]][OxO3380[2]]*=0.8;} else {divpreview[OxO3380[3]][OxO3380[2]]=0.8;} ;} ;} ;function BestFit(){var img=img_demo;if(!img){return ;} ;var Ox265=280;var Ox14=290;if(Browser_IsWinIE()){divpreview[OxO3380[3]][OxO3380[2]]=1/Math.max(img[OxO3380[42]]/Ox14,img[OxO3380[12]]/Ox265);} ;} ;function Actualsize(){inp_width[OxO3380[6]]=OxO3380[7];inp_height[OxO3380[6]]=OxO3380[7];do_preview();if(Browser_IsWinIE()){divpreview[OxO3380[3]][OxO3380[2]]=1;} ;} ;function toggleConstrains(){if(constrain_prop[OxO3380[66]]){imgLock[OxO3380[40]]=OxO3380[67];checkConstrains(OxO3380[42]);} else {imgLock[OxO3380[40]]=OxO3380[68];} ;} ;var checkingConstrains=false;function checkConstrains(Ox301){if(checkingConstrains){return ;} ;checkingConstrains=true;try{if(constrain_prop[OxO3380[66]]){var Ox2e8= new Image();Ox2e8[OxO3380[40]]=TargetUrl[OxO3380[6]];var Ox302=Ox2e8[OxO3380[42]];var Ox303=Ox2e8[OxO3380[12]];if((Ox302>0)&&(Ox303>0)){var Ox14=inp_width[OxO3380[6]];var Ox265=inp_height[OxO3380[6]];if(Ox301==OxO3380[42]){if(Ox14[OxO3380[69]]==0||isNaN(Ox14)){inp_width[OxO3380[6]]=OxO3380[7];inp_height[OxO3380[6]]=OxO3380[7];} else {Ox265=parseInt(Ox14*Ox303/Ox302);inp_height[OxO3380[6]]=Ox265;} ;} ;if(Ox301==OxO3380[12]){if(Ox265[OxO3380[69]]==0||isNaN(Ox265)){inp_width[OxO3380[6]]=OxO3380[7];inp_height[OxO3380[6]]=OxO3380[7];} else {Ox14=parseInt(Ox265*Ox302/Ox303);inp_width[OxO3380[6]]=Ox14;} ;} ;} ;} ;do_preview();} finally{checkingConstrains=false;} ;} ;function ImageEditor(){} ;bordercolor[OxO3380[70]]=bordercolor_Preview[OxO3380[70]]=function bordercolor_onclick(){SelectColor(bordercolor,bordercolor_Preview);} ;if(!Browser_IsWinIE()){btn_zoom_in[OxO3380[3]][OxO3380[71]]=btn_zoom_out[OxO3380[3]][OxO3380[71]]=btn_bestfit[OxO3380[3]][OxO3380[71]]=btn_Actualsize[OxO3380[3]][OxO3380[71]]=OxO3380[72];} ;if(Browser_IsIE7()){var _dialogPromptID=null;function IEprompt(Ox112,Ox113,Ox114){that=this;this[OxO3380[73]]=function (Ox115){val=document.getElementById(OxO3380[74])[OxO3380[6]];_dialogPromptID[OxO3380[3]][OxO3380[71]]=OxO3380[72];document.getElementById(OxO3380[74])[OxO3380[6]]=OxO3380[7];if(Ox115){val=OxO3380[7];} ;Ox112(val);return false;} ;if(Ox114==undefined){Ox114=OxO3380[7];} ;if(_dialogPromptID==null){var Ox116=document.getElementsByTagName(OxO3380[75])[0];tnode=document.createElement(OxO3380[76]);tnode[OxO3380[43]]=OxO3380[77];Ox116.appendChild(tnode);_dialogPromptID=document.getElementById(OxO3380[77]);tnode=document.createElement(OxO3380[76]);tnode[OxO3380[43]]=OxO3380[78];Ox116.appendChild(tnode);_dialogPromptID[OxO3380[3]][OxO3380[46]]=OxO3380[79];_dialogPromptID[OxO3380[3]][OxO3380[49]]=OxO3380[80];_dialogPromptID[OxO3380[3]][OxO3380[81]]=OxO3380[82];_dialogPromptID[OxO3380[3]][OxO3380[42]]=OxO3380[83];_dialogPromptID[OxO3380[3]][OxO3380[84]]=OxO3380[85];} ;var Ox117=OxO3380[86]+InputRequired+OxO3380[87];Ox117+=OxO3380[88]+Ox113+OxO3380[89];Ox117+=OxO3380[90];Ox117+=OxO3380[91]+Ox114+OxO3380[92];Ox117+=OxO3380[93];Ox117+=OxO3380[94]+OK+OxO3380[95];Ox117+=OxO3380[96];Ox117+=OxO3380[97]+Cancel+OxO3380[98];Ox117+=OxO3380[99];_dialogPromptID[OxO3380[100]]=Ox117;_dialogPromptID[OxO3380[3]][OxO3380[5]]=OxO3380[101];_dialogPromptID[OxO3380[3]][OxO3380[102]]=parseInt((document[OxO3380[75]][OxO3380[103]]-315)/2)+OxO3380[104];_dialogPromptID[OxO3380[3]][OxO3380[71]]=OxO3380[105];var Ox118=document.getElementById(OxO3380[74]);try{var Ox119=Ox118.createTextRange();Ox119.collapse(false);Ox119.select();} catch(x){Ox118.focus();} ;} ;} ;if(CreateDir){CreateDir[OxO3380[106]]= new Function(OxO3380[107]);} ;if(btn_zoom_in){btn_zoom_in[OxO3380[106]]= new Function(OxO3380[107]);} ;if(btn_zoom_out){btn_zoom_out[OxO3380[106]]= new Function(OxO3380[107]);} ;if(btn_Actualsize){btn_Actualsize[OxO3380[106]]= new Function(OxO3380[107]);} ;if(btn_bestfit){btn_bestfit[OxO3380[106]]= new Function(OxO3380[107]);} ;