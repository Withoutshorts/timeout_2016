var OxO1dcf=["inp_src","btnbrowse","AlternateText","inp_id","longDesc","Align","optNotSet","optLeft","optRight","optTexttop","optAbsMiddle","optBaseline","optAbsBottom","optBottom","optMiddle","optTop","Border","bordercolor","bordercolor_Preview","inp_width","imgLock","inp_height","constrain_prop","HSpace","VSpace","outer","img_demo","onclick","src","width","height","value","cssText","style","","src_cetemp","id","vspace","hspace","border","borderColor"," ","backgroundColor","align","alt","ValidNumber","ValidID","longdesc","checked","../Images/locked.gif","../Images/1x1.gif","length"];var inp_src=Window_GetElement(window,OxO1dcf[0],true);var btnbrowse=Window_GetElement(window,OxO1dcf[1],true);var AlternateText=Window_GetElement(window,OxO1dcf[2],true);var inp_id=Window_GetElement(window,OxO1dcf[3],true);var longDesc=Window_GetElement(window,OxO1dcf[4],true);var Align=Window_GetElement(window,OxO1dcf[5],true);var optNotSet=Window_GetElement(window,OxO1dcf[6],true);var optLeft=Window_GetElement(window,OxO1dcf[7],true);var optRight=Window_GetElement(window,OxO1dcf[8],true);var optTexttop=Window_GetElement(window,OxO1dcf[9],true);var optAbsMiddle=Window_GetElement(window,OxO1dcf[10],true);var optBaseline=Window_GetElement(window,OxO1dcf[11],true);var optAbsBottom=Window_GetElement(window,OxO1dcf[12],true);var optBottom=Window_GetElement(window,OxO1dcf[13],true);var optMiddle=Window_GetElement(window,OxO1dcf[14],true);var optTop=Window_GetElement(window,OxO1dcf[15],true);var Border=Window_GetElement(window,OxO1dcf[16],true);var bordercolor=Window_GetElement(window,OxO1dcf[17],true);var bordercolor_Preview=Window_GetElement(window,OxO1dcf[18],true);var inp_width=Window_GetElement(window,OxO1dcf[19],true);var imgLock=Window_GetElement(window,OxO1dcf[20],true);var inp_height=Window_GetElement(window,OxO1dcf[21],true);var constrain_prop=Window_GetElement(window,OxO1dcf[22],true);var HSpace=Window_GetElement(window,OxO1dcf[23],true);var VSpace=Window_GetElement(window,OxO1dcf[24],true);var outer=Window_GetElement(window,OxO1dcf[25],true);var img_demo=Window_GetElement(window,OxO1dcf[26],true);btnbrowse[OxO1dcf[27]]=function btnbrowse_onclick(){function Ox260(Ox6){if(Ox6){function Actualsize(){var Ox2e9= new Image();Ox2e9[OxO1dcf[28]]=Ox6;if(Ox2e9[OxO1dcf[29]]>0&&Ox2e9[OxO1dcf[30]]>0){inp_width[OxO1dcf[31]]=Ox2e9[OxO1dcf[29]];inp_height[OxO1dcf[31]]=Ox2e9[OxO1dcf[30]];FireUIChanged();} else {setTimeout(Actualsize,400);} ;} ;inp_src[OxO1dcf[31]]=Ox6;setTimeout(Actualsize,400);} ;} ;editor.SetNextDialogWindow(window);if(Browser_IsSafari()){editor.ShowSelectImageDialog(Ox260,inp_src.value,inp_src);} else {editor.ShowSelectImageDialog(Ox260,inp_src.value);} ;} ;UpdateState=function UpdateState_Image(){img_demo[OxO1dcf[33]][OxO1dcf[32]]=element[OxO1dcf[33]][OxO1dcf[32]];if(Browser_IsWinIE()){img_demo.mergeAttributes(element);} ;if(element[OxO1dcf[28]]){img_demo[OxO1dcf[28]]=element[OxO1dcf[28]];} else {img_demo.removeAttribute(OxO1dcf[28]);} ;} ;SyncToView=function SyncToView_Image(){var src;src=element.getAttribute(OxO1dcf[28])+OxO1dcf[34];if(element.getAttribute(OxO1dcf[35])){src=element.getAttribute(OxO1dcf[35])+OxO1dcf[34];} ;inp_src[OxO1dcf[31]]=src;inp_width[OxO1dcf[31]]=element[OxO1dcf[29]];inp_height[OxO1dcf[31]]=element[OxO1dcf[30]];inp_id[OxO1dcf[31]]=element[OxO1dcf[36]];if(element[OxO1dcf[37]]<=0){VSpace[OxO1dcf[31]]=OxO1dcf[34];} else {VSpace[OxO1dcf[31]]=element[OxO1dcf[37]];} ;if(element[OxO1dcf[38]]<=0){HSpace[OxO1dcf[31]]=OxO1dcf[34];} else {HSpace[OxO1dcf[31]]=element[OxO1dcf[38]];} ;Border[OxO1dcf[31]]=element[OxO1dcf[39]];if(Browser_IsWinIE()){bordercolor[OxO1dcf[31]]=element[OxO1dcf[33]][OxO1dcf[40]];} else {var arr=revertColor(element[OxO1dcf[33]].borderColor).split(OxO1dcf[41]);bordercolor[OxO1dcf[31]]=arr[0];} ;bordercolor[OxO1dcf[33]][OxO1dcf[42]]=bordercolor[OxO1dcf[31]]||OxO1dcf[34];bordercolor[OxO1dcf[33]][OxO1dcf[42]]=bordercolor[OxO1dcf[31]];bordercolor_Preview[OxO1dcf[33]][OxO1dcf[42]]=bordercolor[OxO1dcf[31]];Align[OxO1dcf[31]]=element[OxO1dcf[43]];AlternateText[OxO1dcf[31]]=element[OxO1dcf[44]];longDesc[OxO1dcf[31]]=element[OxO1dcf[4]];} ;SyncTo=function SyncTo_Image(element){element[OxO1dcf[28]]=inp_src[OxO1dcf[31]];element.setAttribute(OxO1dcf[35],inp_src.value);element[OxO1dcf[39]]=Border[OxO1dcf[31]];element[OxO1dcf[38]]=HSpace[OxO1dcf[31]];element[OxO1dcf[37]]=VSpace[OxO1dcf[31]];try{element[OxO1dcf[29]]=inp_width[OxO1dcf[31]];element[OxO1dcf[30]]=inp_height[OxO1dcf[31]];} catch(er){alert(CE_GetStr(OxO1dcf[45]));return false;} ;if(element[OxO1dcf[33]][OxO1dcf[29]]||element[OxO1dcf[33]][OxO1dcf[30]]){try{element[OxO1dcf[33]][OxO1dcf[29]]=inp_width[OxO1dcf[31]];element[OxO1dcf[33]][OxO1dcf[30]]=inp_height[OxO1dcf[31]];} catch(er){alert(CE_GetStr(OxO1dcf[45]));return false;} ;} ;var Ox27b=/[^a-z\d]/i;if(Ox27b.test(inp_id.value)){alert(CE_GetStr(OxO1dcf[46]));return ;} ;element[OxO1dcf[36]]=inp_id[OxO1dcf[31]];element[OxO1dcf[43]]=Align[OxO1dcf[31]];element[OxO1dcf[44]]=AlternateText[OxO1dcf[31]];element[OxO1dcf[4]]=longDesc[OxO1dcf[31]];element[OxO1dcf[33]][OxO1dcf[40]]=bordercolor[OxO1dcf[31]];if(element[OxO1dcf[47]]==OxO1dcf[34]||element[OxO1dcf[47]]==null){element.removeAttribute(OxO1dcf[47]);} ;if(element[OxO1dcf[4]]==OxO1dcf[34]||element[OxO1dcf[4]]==null){element.removeAttribute(OxO1dcf[4]);} ;if(element[OxO1dcf[29]]==0){element.removeAttribute(OxO1dcf[29]);} ;if(element[OxO1dcf[30]]==0){element.removeAttribute(OxO1dcf[30]);} ;if(element[OxO1dcf[38]]==OxO1dcf[34]){element.removeAttribute(OxO1dcf[38]);} ;if(element[OxO1dcf[37]]==OxO1dcf[34]){element.removeAttribute(OxO1dcf[37]);} ;if(element[OxO1dcf[36]]==OxO1dcf[34]){element.removeAttribute(OxO1dcf[36]);} ;if(element[OxO1dcf[43]]==OxO1dcf[34]){element.removeAttribute(OxO1dcf[43]);} ;if(element[OxO1dcf[39]]==OxO1dcf[34]){element.removeAttribute(OxO1dcf[39]);} ;} ;function toggleConstrains(){if(constrain_prop[OxO1dcf[48]]){imgLock[OxO1dcf[28]]=OxO1dcf[49];checkConstrains(OxO1dcf[29]);} else {imgLock[OxO1dcf[28]]=OxO1dcf[50];} ;} ;var checkingConstrains=false;function checkConstrains(Ox302){if(checkingConstrains){return ;} ;checkingConstrains=true;try{var Ox8a,Ox23c;if(constrain_prop[OxO1dcf[48]]){var Ox2e9= new Image();Ox2e9[OxO1dcf[28]]=inp_src[OxO1dcf[31]];var Ox303=Ox2e9[OxO1dcf[29]];var Ox304=Ox2e9[OxO1dcf[30]];if((Ox303>0)&&(Ox304>0)){var Ox14=inp_width[OxO1dcf[31]];var Ox266=inp_height[OxO1dcf[31]];if(Ox302==OxO1dcf[29]){if(Ox14[OxO1dcf[51]]==0||isNaN(Ox14)){inp_width[OxO1dcf[31]]=OxO1dcf[34];inp_height[OxO1dcf[31]]=OxO1dcf[34];} else {Ox266=parseInt(Ox14*Ox304/Ox303);inp_height[OxO1dcf[31]]=Ox266;} ;} ;if(Ox302==OxO1dcf[30]){if(Ox266[OxO1dcf[51]]==0||isNaN(Ox266)){inp_width[OxO1dcf[31]]=OxO1dcf[34];inp_height[OxO1dcf[31]]=OxO1dcf[34];} else {Ox14=parseInt(Ox266*Ox303/Ox304);inp_width[OxO1dcf[31]]=Ox14;} ;} ;} ;} ;} finally{checkingConstrains=false;} ;} ;bordercolor[OxO1dcf[27]]=bordercolor_Preview[OxO1dcf[27]]=function bordercolor_onclick(){SelectColor(bordercolor,bordercolor_Preview);} ;