var OxO8d69=["inp_type","inp_name","inp_value","row_txt1","inp_Size","row_txt2","inp_MaxLength","row_img","inp_src","btnbrowse","row_img2","sel_Align","optNotSet","optLeft","optRight","optTexttop","optAbsMiddle","optBaseline","optAbsBottom","optBottom","optMiddle","optTop","inp_Border","row_img3","inp_width","inp_height","row_img4","inp_HSpace","inp_VSpace","row_img5","AlternateText","inp_id","row_txt3","inp_access","row_txt4","inp_index","row_chk","inp_checked","row_txt5","inp_Disabled","row_txt6","inp_Readonly","onclick","value","Name","name","id","src","type","checked","disabled","readOnly","tabIndex","","accessKey","size","maxLength","width","height","vspace","hspace","border","align","alt","text","display","style","none","password","hidden","radio","checkbox","submit","reset","button","image","className","class"];var inp_type=Window_GetElement(window,OxO8d69[0],true);var inp_name=Window_GetElement(window,OxO8d69[1],true);var inp_value=Window_GetElement(window,OxO8d69[2],true);var row_txt1=Window_GetElement(window,OxO8d69[3],true);var inp_Size=Window_GetElement(window,OxO8d69[4],true);var row_txt2=Window_GetElement(window,OxO8d69[5],true);var inp_MaxLength=Window_GetElement(window,OxO8d69[6],true);var row_img=Window_GetElement(window,OxO8d69[7],true);var inp_src=Window_GetElement(window,OxO8d69[8],true);var btnbrowse=Window_GetElement(window,OxO8d69[9],true);var row_img2=Window_GetElement(window,OxO8d69[10],true);var sel_Align=Window_GetElement(window,OxO8d69[11],true);var optNotSet=Window_GetElement(window,OxO8d69[12],true);var optLeft=Window_GetElement(window,OxO8d69[13],true);var optRight=Window_GetElement(window,OxO8d69[14],true);var optTexttop=Window_GetElement(window,OxO8d69[15],true);var optAbsMiddle=Window_GetElement(window,OxO8d69[16],true);var optBaseline=Window_GetElement(window,OxO8d69[17],true);var optAbsBottom=Window_GetElement(window,OxO8d69[18],true);var optBottom=Window_GetElement(window,OxO8d69[19],true);var optMiddle=Window_GetElement(window,OxO8d69[20],true);var optTop=Window_GetElement(window,OxO8d69[21],true);var inp_Border=Window_GetElement(window,OxO8d69[22],true);var row_img3=Window_GetElement(window,OxO8d69[23],true);var inp_width=Window_GetElement(window,OxO8d69[24],true);var inp_height=Window_GetElement(window,OxO8d69[25],true);var row_img4=Window_GetElement(window,OxO8d69[26],true);var inp_HSpace=Window_GetElement(window,OxO8d69[27],true);var inp_VSpace=Window_GetElement(window,OxO8d69[28],true);var row_img5=Window_GetElement(window,OxO8d69[29],true);var AlternateText=Window_GetElement(window,OxO8d69[30],true);var inp_id=Window_GetElement(window,OxO8d69[31],true);var row_txt3=Window_GetElement(window,OxO8d69[32],true);var inp_access=Window_GetElement(window,OxO8d69[33],true);var row_txt4=Window_GetElement(window,OxO8d69[34],true);var inp_index=Window_GetElement(window,OxO8d69[35],true);var row_chk=Window_GetElement(window,OxO8d69[36],true);var inp_checked=Window_GetElement(window,OxO8d69[37],true);var row_txt5=Window_GetElement(window,OxO8d69[38],true);var inp_Disabled=Window_GetElement(window,OxO8d69[39],true);var row_txt6=Window_GetElement(window,OxO8d69[40],true);var inp_Readonly=Window_GetElement(window,OxO8d69[41],true);btnbrowse[OxO8d69[42]]=function btnbrowse_onclick(){function Ox25f(Ox6){if(Ox6){inp_src[OxO8d69[43]]=Ox6;SyncTo(element);} ;} ;editor.SetNextDialogWindow(window);if(Browser_IsSafari()){editor.ShowSelectImageDialog(Ox25f,inp_src.value,inp_src);} else {editor.ShowSelectImageDialog(Ox25f,inp_src.value);} ;} ;UpdateState=function UpdateState_Input(){} ;SyncToView=function SyncToView_Input(){if(element[OxO8d69[44]]){inp_name[OxO8d69[43]]=element[OxO8d69[44]];} ;if(element[OxO8d69[45]]){inp_name[OxO8d69[43]]=element[OxO8d69[45]];} ;inp_id[OxO8d69[43]]=element[OxO8d69[46]];inp_value[OxO8d69[43]]=(element[OxO8d69[43]]).trim();inp_src[OxO8d69[43]]=element[OxO8d69[47]];inp_type[OxO8d69[43]]=element[OxO8d69[48]];inp_checked[OxO8d69[49]]=element[OxO8d69[49]];inp_Disabled[OxO8d69[49]]=element[OxO8d69[50]];inp_Readonly[OxO8d69[49]]=element[OxO8d69[51]];if(element[OxO8d69[52]]==0){inp_index[OxO8d69[43]]=OxO8d69[53];} else {inp_index[OxO8d69[43]]=element[OxO8d69[52]];} ;if(element[OxO8d69[54]]){inp_access[OxO8d69[43]]=element[OxO8d69[54]];} ;if(element[OxO8d69[55]]){if(element[OxO8d69[55]]==20){inp_Size[OxO8d69[43]]=OxO8d69[53];} else {inp_Size[OxO8d69[43]]=element[OxO8d69[55]];} ;} ;if(element[OxO8d69[56]]){if(element[OxO8d69[56]]==2147483647||element[OxO8d69[56]]<=0){inp_MaxLength[OxO8d69[43]]=OxO8d69[53];} else {inp_MaxLength[OxO8d69[43]]=element[OxO8d69[56]];} ;} ;if(element[OxO8d69[57]]){inp_width[OxO8d69[43]]=element[OxO8d69[57]];} ;if(element[OxO8d69[58]]){inp_height[OxO8d69[43]]=element[OxO8d69[58]];} ;if(element[OxO8d69[59]]){inp_HSpace[OxO8d69[43]]=element[OxO8d69[59]];} ;if(element[OxO8d69[60]]){inp_VSpace[OxO8d69[43]]=element[OxO8d69[60]];} ;if(element[OxO8d69[61]]){inp_Border[OxO8d69[43]]=element[OxO8d69[61]];} ;if(element[OxO8d69[62]]){sel_Align[OxO8d69[43]]=element[OxO8d69[62]];} ;if(element[OxO8d69[63]]){alt[OxO8d69[43]]=element[OxO8d69[63]];} ;switch((element[OxO8d69[48]]).toLowerCase()){case OxO8d69[64]:;case OxO8d69[68]:row_img[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_img2[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_img3[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_img4[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_img5[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_chk[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];break ;;case OxO8d69[69]:row_img[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_img2[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_img3[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_img4[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_img5[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_chk[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_txt1[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_txt2[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_txt3[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_txt4[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_txt5[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_txt6[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];break ;;case OxO8d69[70]:;case OxO8d69[71]:row_img[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_img2[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_img3[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_img4[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_img5[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_txt1[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_txt2[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_txt6[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];break ;;case OxO8d69[72]:;case OxO8d69[73]:;case OxO8d69[74]:row_chk[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_img[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_img2[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_img3[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_img4[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_img5[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_txt1[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_txt2[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_txt6[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];break ;;case OxO8d69[75]:row_chk[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_txt1[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_txt2[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];row_txt6[OxO8d69[66]][OxO8d69[65]]=OxO8d69[67];break ;;} ;} ;SyncTo=function SyncTo_Input(element){element[OxO8d69[45]]=inp_name[OxO8d69[43]];if(element[OxO8d69[44]]){element[OxO8d69[44]]=inp_name[OxO8d69[43]];} else {if(element[OxO8d69[45]]){element.removeAttribute(OxO8d69[45],0);element[OxO8d69[44]]=inp_name[OxO8d69[43]];} else {element[OxO8d69[44]]=inp_name[OxO8d69[43]];} ;} ;element[OxO8d69[46]]=inp_id[OxO8d69[43]];if(inp_src[OxO8d69[43]]){element[OxO8d69[47]]=inp_src[OxO8d69[43]];} ;element[OxO8d69[49]]=inp_checked[OxO8d69[49]];element[OxO8d69[43]]=inp_value[OxO8d69[43]];element[OxO8d69[50]]=inp_Disabled[OxO8d69[49]];element[OxO8d69[51]]=inp_Readonly[OxO8d69[49]];element[OxO8d69[54]]=inp_access[OxO8d69[43]];element[OxO8d69[52]]=inp_index[OxO8d69[43]];element[OxO8d69[56]]=inp_MaxLength[OxO8d69[43]];element[OxO8d69[57]]=inp_width[OxO8d69[43]];element[OxO8d69[58]]=inp_height[OxO8d69[43]];element[OxO8d69[59]]=inp_HSpace[OxO8d69[43]];element[OxO8d69[60]]=inp_VSpace[OxO8d69[43]];element[OxO8d69[61]]=inp_Border[OxO8d69[43]];element[OxO8d69[62]]=sel_Align[OxO8d69[43]];element[OxO8d69[63]]=AlternateText[OxO8d69[43]];try{element[OxO8d69[55]]=inp_Size[OxO8d69[43]];} catch(e){element[OxO8d69[55]]=20;} ;if(element[OxO8d69[52]]==OxO8d69[53]){element.removeAttribute(OxO8d69[52]);} ;if(element[OxO8d69[54]]==OxO8d69[53]){element.removeAttribute(OxO8d69[54]);} ;if(element[OxO8d69[56]]==OxO8d69[53]){element.removeAttribute(OxO8d69[56]);} ;if(element[OxO8d69[55]]==0){element.removeAttribute(OxO8d69[55]);} ;if(element[OxO8d69[57]]==0){element.removeAttribute(OxO8d69[57]);} ;if(element[OxO8d69[58]]==0){element.removeAttribute(OxO8d69[58]);} ;if(element[OxO8d69[60]]==OxO8d69[53]){element.removeAttribute(OxO8d69[60]);} ;if(element[OxO8d69[59]]==OxO8d69[53]){element.removeAttribute(OxO8d69[59]);} ;if(element[OxO8d69[46]]==OxO8d69[53]){element.removeAttribute(OxO8d69[46]);} ;if(element[OxO8d69[44]]==OxO8d69[53]){element.removeAttribute(OxO8d69[44]);} ;if(element[OxO8d69[63]]==OxO8d69[53]){element.removeAttribute(OxO8d69[63]);} ;if(element[OxO8d69[62]]==OxO8d69[53]){element.removeAttribute(OxO8d69[62]);} ;if(element[OxO8d69[76]]==OxO8d69[53]){element.removeAttribute(OxO8d69[77]);} ;if(element[OxO8d69[76]]==OxO8d69[53]){element.removeAttribute(OxO8d69[76]);} ;switch((element[OxO8d69[48]]).toLowerCase()){case OxO8d69[64]:;case OxO8d69[68]:;case OxO8d69[69]:;case OxO8d69[70]:;case OxO8d69[71]:;case OxO8d69[72]:;case OxO8d69[73]:;case OxO8d69[74]:element.removeAttribute(OxO8d69[58]);element.removeAttribute(OxO8d69[61]);element.removeAttribute(OxO8d69[47]);break ;;case OxO8d69[75]:break ;;} ;} ;