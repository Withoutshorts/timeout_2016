var OxO668e=["rowSpan","removeNode","parentNode","firstChild","colSpan","nodeName","TABLE","Map","rowIndex","rows","length","cells","cellIndex","innerHTML","cell","\x26nbsp;","Can\x27t Get The Position ?","RowCount","ColCount","Unknown Error , pos ",":"," already have cell","Unknown Error , Unable to find bestpos","tbody","uniqueID","richDropDown","list_Templates","subcolumns","addcolumns","subcolspan","addcolspan","btn_row_dialog","btn_cell_dialog","inp_cell_width","inp_cell_height","btn_cell_editcell","tabledesign","subrows","addrows","subrowspan","addrowspan","display","style","none","disabled","value","width","height","ValidNumber","","\x3Ctable\x3E\x3Ctr\x3E\x3Ctd height=\x2224\x22\x3E\x3C/td\x3E\x3C/tr\x3E\x3C/table\x3E","\x3Ctable\x3E\x3Ctr\x3E\x3Ctd\x3E\x3C/td\x3E\x3Ctd\x3E\x3C/td\x3E\x3C/tr\x3E\x3C/table\x3E","\x3Ctable\x3E\x3Ctr\x3E\x3Ctd\x3E\x3C/td\x3E\x3Ctd\x3E\x3C/td\x3E\x3Ctd\x3E\x3C/td\x3E\x3C/tr\x3E\x3C/table\x3E","\x3Ctable border=\x220\x22 cellpadding=\x220\x22 cellspacing=\x220\x22\x3E\x3Ctr\x3E\x3Ctd valign=\x22top\x22 colspan=\x222\x22\x3E\x3C/td\x3E\x3C/tr\x3E","\x3Ctr\x3E\x3Ctd valign=\x22top\x22 rowspan=\x222\x22\x3E\x3C/td\x3E\x3Ctd valign=\x22top\x22\x3E\x3C/td\x3E\x3C/tr\x3E","\x3Ctr\x3E\x3Ctd valign=\x22top\x22\x3E\x3C/td\x3E\x3C/tr\x3E\x3C/table\x3E","\x3Ctr\x3E\x3Ctd valign=\x22top\x22\x3E\x3C/td\x3E\x3Ctd valign=\x22top\x22 rowspan=\x222\x22\x3E\x3C/td\x3E\x3C/tr\x3E\x3Ctr\x3E\x3Ctd valign=\x22top\x22\x3E\x3C/td\x3E\x3C/tr\x3E\x3C/table\x3E","\x3Ctable border=\x220\x22 cellpadding=\x220\x22 cellspacing=\x220\x22\x3E\x3Ctr\x3E\x3Ctd valign=\x22top\x22 colspan=\x223\x22\x3E\x3C/td\x3E\x3C/tr\x3E","\x3Ctr\x3E\x3Ctd valign=\x22top\x22\x3E\x3C/td\x3E\x3Ctd valign=\x22top\x22\x3E\x3C/td\x3E\x3Ctd valign=\x22top\x22\x3E\x3C/td\x3E\x3C/tr\x3E","\x3Ctr\x3E\x3Ctd valign=\x22top\x22 colspan=\x223\x22\x3E\x3C/td\x3E\x3C/tr\x3E\x3C/table\x3E","DIV","onclick","bgColor","#d4d0c8","onmouseover","onmouseout","ondblclick","contains","ToggleBorder","backgroundColor","highlight","cssText","runtimeStyle","background-color:gray","onload","body","document","csstext","font:normal 11px Tahoma;background-color:windowtext;","isOpen"];function Table_GetCellFromMap(Ox3ec,Ox3ed,Ox3ee){var Ox3ef=Ox3ec[Ox3ed];if(Ox3ef){return Ox3ef[Ox3ee];} ;} ;function Table_CanSubRowSpan(Ox3f1){return Ox3f1[OxO668e[0]]>1;} ;function Element_RemoveNode(element,Ox3f3){if(element[OxO668e[1]]){element.removeNode(Ox3f3);return ;} ;var p=element[OxO668e[2]];if(!p){return ;} ;if(Ox3f3){p.removeChild(element);return ;} ;while(true){var Ox13a=element[OxO668e[3]];if(!Ox13a){break ;} ;p.insertBefore(Ox13a,element);} ;p.removeChild(element);} ;function Table_CanSubColSpan(Ox3f1){return Ox3f1[OxO668e[4]]>1;} ;function Table_GetTable(Ox1f){for(;Ox1f!=null;Ox1f=Ox1f[OxO668e[2]]){if(Ox1f[OxO668e[5]]==OxO668e[6]){return Ox1f;} ;} ;return null;} ;function Table_SubRowSpan(Ox3f1){var Ox3f7=Table_GetTable(Ox3f1);var Ox11d=Table_CalculateTableInfo(Ox3f7);var Ox3ec=Ox11d[OxO668e[7]];var Ox19=Table_GetCellPositionFromMap(Ox3ec,Ox3f1);var Ox3f8=Ox3f7[OxO668e[9]].item(Ox3f1[OxO668e[2]][OxO668e[8]]+Ox3f1[OxO668e[0]]-1);var Ox3f9=0;for(var i=0;i<Ox3f8[OxO668e[11]][OxO668e[10]];i++){var Ox3fa=Ox3f8[OxO668e[11]].item(i);var Ox3fb=Table_GetCellPositionFromMap(Ox3ec,Ox3fa);if(Ox3fb[OxO668e[12]]<Ox19[OxO668e[12]]){Ox3f9=i+1;} ;} ;for(var i=0;i<Ox3f1[OxO668e[4]];i++){var Ox13a=Ox3f8.insertCell(Ox3f9);if(Browser_IsOpera()){Ox13a[OxO668e[13]]=OxO668e[14];} else {if(Browser_IsGecko()||Browser_IsSafari()){Ox13a[OxO668e[13]]=OxO668e[15];} ;} ;} ;Ox3f1[OxO668e[0]]--;} ;function Table_GetCellPositionFromMap(Ox3ec,Ox3f1){for(var Oxc4=0;Oxc4<Ox3ec[OxO668e[10]];Oxc4++){var Ox3ef=Ox3ec[Oxc4];for(var Oxf6=0;Oxf6<Ox3ef[OxO668e[10]];Oxf6++){if(Ox3ef[Oxf6]==Ox3f1){return {rowIndex:Oxc4,cellIndex:Oxf6};} ;} ;} ;throw ( new Error(-1,OxO668e[16]));} ;function Table_GetCellMap(Ox3f7){return Table_CalculateTableInfo(Ox3f7)[OxO668e[7]];} ;function Table_GetRowCount(Ox1f){return Table_CalculateTableInfo(Ox1f)[OxO668e[17]];} ;function Table_GetColCount(Ox1f){return Table_CalculateTableInfo(Ox1f)[OxO668e[18]];} ;function Table_CalculateTableInfo(Ox1f){var Ox3f7=Table_GetTable(Ox1f);var Ox401=Ox3f7[OxO668e[9]];for(var Oxb3=Ox401[OxO668e[10]]-1;Oxb3>=0;Oxb3--){var Ox402=Ox401.item(Oxb3);if(Ox402[OxO668e[11]][OxO668e[10]]==0){Element_RemoveNode(Ox402,true);continue ;} ;} ;var Ox403=Ox401[OxO668e[10]];var Ox404=0;var Ox405= new Array(Ox401.length);for(var Ox406=0;Ox406<Ox403;Ox406++){Ox405[Ox406]=[];} ;function Ox407(Oxb3,Ox13a,Ox3f1){while(Oxb3>=Ox403){Ox405[Ox403]=[];Ox403++;} ;var Ox408=Ox405[Oxb3];if(Ox13a>=Ox404){Ox404=Ox13a+1;} ;if(Ox408[Ox13a]!=null){throw ( new Error(-1,OxO668e[19]+Oxb3+OxO668e[20]+Ox13a+OxO668e[21]));} ;Ox408[Ox13a]=Ox3f1;} ;function Ox409(Oxb3,Ox13a){var Ox408=Ox405[Oxb3];if(Ox408){return Ox408[Ox13a];} ;} ;for(var Ox406=0;Ox406<Ox401[OxO668e[10]];Ox406++){var Ox402=Ox401.item(Ox406);var Ox40a=Ox402[OxO668e[11]];for(var Ox40b=0;Ox40b<Ox40a[OxO668e[10]];Ox40b++){var Ox3f1=Ox40a.item(Ox40b);var Ox40c=Ox3f1[OxO668e[0]];var Ox40d=Ox3f1[OxO668e[4]];var Ox408=Ox405[Ox406];var Ox40e=-1;for(var Ox40f=0;Ox40f<Ox404+Ox40d+1;Ox40f++){if(Ox40c==1&&Ox40d==1){if(Ox408[Ox40f]==null){Ox40e=Ox40f;break ;} ;} else {var Ox410=true;for(var Ox411=0;Ox411<Ox40c;Ox411++){for(var Ox412=0;Ox412<Ox40d;Ox412++){if(Ox409(Ox406+Ox411,Ox40f+Ox412)!=null){Ox410=false;break ;} ;} ;} ;if(Ox410){Ox40e=Ox40f;break ;} ;} ;} ;if(Ox40e==-1){throw ( new Error(-1,OxO668e[22]));} ;if(Ox40c==1&&Ox40d==1){Ox407(Ox406,Ox40e,Ox3f1);} else {for(var Ox411=0;Ox411<Ox40c;Ox411++){for(var Ox412=0;Ox412<Ox40d;Ox412++){Ox407(Ox406+Ox411,Ox40e+Ox412,Ox3f1);} ;} ;} ;} ;} ;var Ox6={};Ox6[OxO668e[17]]=Ox403;Ox6[OxO668e[18]]=Ox404;Ox6[OxO668e[7]]=Ox405;for(var Oxb3=0;Oxb3<Ox403;Oxb3++){var Ox408=Ox405[Oxb3];for(var Ox13a=0;Ox13a<Ox404;Ox13a++){} ;} ;return Ox6;} ;function Table_SubColSpan(Ox3f1){var Ox3f7=Table_GetTable(Ox3f1);Ox3f1[OxO668e[2]].insertCell(Ox3f1[OxO668e[12]]+1)[OxO668e[0]]=Ox3f1[OxO668e[0]];Ox3f1[OxO668e[4]]--;} ;function Table_CanAddRowSpan(Ox3f1){var Ox3f7=Table_GetTable(Ox3f1);var Ox11d=Table_CalculateTableInfo(Ox3f7);var Ox3ec=Ox11d[OxO668e[7]];var Ox19=Table_GetCellPositionFromMap(Ox3ec,Ox3f1);var Ox415=0;for(var Ox13a=0;Ox13a<Ox3f1[OxO668e[4]];Ox13a++){var Ox416=Table_GetCellFromMap(Ox3ec,Ox19[OxO668e[8]]+Ox3f1[OxO668e[0]],Ox19[OxO668e[12]]+Ox13a);if(Ox416==null){return false;} ;if(Ox415!=0&&Ox415!=Ox416[OxO668e[0]]){return false;} ;Ox415=Ox416[OxO668e[0]];var Ox417=Table_GetCellPositionFromMap(Ox3ec,Ox416);if(Ox417[OxO668e[12]]<Ox19[OxO668e[12]]){return false;} ;if(Ox417[OxO668e[12]]+Ox416[OxO668e[4]]>Ox19[OxO668e[12]]+Ox3f1[OxO668e[4]]){return false;} ;} ;return true;} ;function Table_AddRow(Ox3f1){var Ox3f7=Table_GetTable(Ox3f1);var Ox11d=Table_CalculateTableInfo(Ox3f7);var Ox3ec=Ox11d[OxO668e[7]];var Ox404=Ox11d[OxO668e[18]];var Ox403=Ox11d[OxO668e[17]];var Ox402;if(!Browser_IsSafari()){Ox402=Ox3f7.insertRow(-1);} else {var Ox116=document.createElement(OxO668e[23]);Ox3f7.appendChild(Ox116);Ox402=Ox116.insertRow(-1);} ;for(var i=0;i<Ox404;i++){var Ox13a=Ox402.insertCell(-1);if(Browser_IsOpera()){Ox13a[OxO668e[13]]=OxO668e[14];} else {if(Browser_IsGecko()||Browser_IsSafari()){Ox13a[OxO668e[13]]=OxO668e[15];} ;} ;} ;} ;function Table_AddCol(Ox3f1){var Ox3f7=Table_GetTable(Ox3f1);for(var Oxb3=0;Oxb3<Ox3f7[OxO668e[9]][OxO668e[10]];Oxb3++){var Ox402=Ox3f7[OxO668e[9]].item(Oxb3);var Ox13a=Ox402.insertCell(-1);if(Browser_IsOpera()){Ox13a[OxO668e[13]]=OxO668e[14];} else {if(Browser_IsGecko()||Browser_IsSafari()){Ox13a[OxO668e[13]]=OxO668e[15];} ;} ;} ;} ;function Table_CleanUpTableInfo(Ox3f7,Ox11d){var Ox401=Ox3f7[OxO668e[9]];for(var Oxb3=Ox401[OxO668e[10]]-1;Oxb3>=0;Oxb3--){var Ox402=Ox401.item(Oxb3);if(Ox402[OxO668e[11]][OxO668e[10]]==0){Element_RemoveNode(Ox402,true);continue ;} ;} ;var Ox3ec=Ox11d[OxO668e[7]];var Ox404=Ox11d[OxO668e[18]];var Ox403=Ox11d[OxO668e[17]];for(var Ox406=1;Ox406<Ox403;Ox406++){var Ox41b=true;for(var Ox40b=0;Ox40b<Ox404;Ox40b++){if(Table_GetCellFromMap(Ox3ec,Ox406-1,Ox40b)!=Table_GetCellFromMap(Ox3ec,Ox406,Ox40b)){Ox41b=false;break ;} ;} ;if(Ox41b){for(var Ox40b=0;Ox40b<Ox404;Ox40b++){var Ox3f1=Table_GetCellFromMap(Ox3ec,Ox406,Ox40b);if(Ox3f1){if(Ox3f1[OxO668e[0]]>1){Ox3f1[OxO668e[0]]=Ox3f1[OxO668e[0]]-1;} ;Ox40b+=Ox3f1[OxO668e[4]]-1;} ;} ;} ;} ;for(var Ox40b=1;Ox40b<Ox404;Ox40b++){var Ox41b=true;for(var Ox406=0;Ox406<Ox403;Ox406++){if(Table_GetCellFromMap(Ox3ec,Ox406,Ox40b-1)!=Table_GetCellFromMap(Ox3ec,Ox406,Ox40b)){Ox41b=false;break ;} ;} ;if(Ox41b){for(var Ox406=0;Ox406<Ox403;Ox406++){var Ox3f1=Table_GetCellFromMap(Ox3ec,Ox406,Ox40b);if(Ox3f1){if(Ox3f1[OxO668e[4]]>1){Ox3f1[OxO668e[4]]=Ox3f1[OxO668e[4]]-1;} ;Ox406+=Ox3f1[OxO668e[0]]-1;} ;} ;} ;} ;} ;function Table_SubRow(Ox3f1){var Ox3f7=Table_GetTable(Ox3f1);var Ox11d=Table_CalculateTableInfo(Ox3f7);var Ox3ec=Ox11d[OxO668e[7]];var Ox404=Ox11d[OxO668e[18]];var Ox403=Ox11d[OxO668e[17]];if(Ox403==0){return ;} ;var Ox41d={};var Ox406=Ox403-1;for(var Ox40b=0;Ox40b<Ox404;Ox40b++){var Ox3f1=Table_GetCellFromMap(Ox3ec,Ox406,Ox40b);if(Ox3f1){if(Ox41d[Ox3f1[OxO668e[24]]]){continue ;} ;Ox41d[Ox3f1[OxO668e[24]]]=true;if(Ox3f1[OxO668e[0]]==1){Element_RemoveNode(Ox3f1,true);} else {Ox3f1[OxO668e[0]]=Ox3f1[OxO668e[0]]-1;} ;} ;} ;Ox11d[OxO668e[17]]--;Table_CleanUpTableInfo(Ox3f7,Ox11d);} ;function Table_CanAddColSpan(Ox3f1){var Ox3f7=Table_GetTable(Ox3f1);var Ox11d=Table_CalculateTableInfo(Ox3f7);var Ox3ec=Ox11d[OxO668e[7]];var Ox19=Table_GetCellPositionFromMap(Ox3ec,Ox3f1);var Ox41f=0;for(var Oxb3=0;Oxb3<Ox3f1[OxO668e[0]];Oxb3++){var Ox416=Table_GetCellFromMap(Ox3ec,Ox19[OxO668e[8]]+Oxb3,Ox19[OxO668e[12]]+Ox3f1[OxO668e[4]]);if(Ox416==null){return false;} ;if(Ox41f!=0&&Ox41f!=Ox416[OxO668e[4]]){return false;} ;Ox41f=Ox416[OxO668e[4]];var Ox417=Table_GetCellPositionFromMap(Ox3ec,Ox416);if(Ox417[OxO668e[8]]<Ox19[OxO668e[8]]){return false;} ;if(Ox417[OxO668e[8]]+Ox416[OxO668e[0]]>Ox19[OxO668e[8]]+Ox3f1[OxO668e[0]]){return false;} ;} ;return true;} ;function Table_AddRowSpan(Ox3f1){var Ox3f7=Table_GetTable(Ox3f1);var Ox11d=Table_CalculateTableInfo(Ox3f7);var Ox3ec=Ox11d[OxO668e[7]];var Ox19=Table_GetCellPositionFromMap(Ox3ec,Ox3f1);var Ox415=0;for(var Ox13a=0;Ox13a<Ox3f1[OxO668e[4]];Ox13a++){var Ox416=Table_GetCellFromMap(Ox3ec,Ox19[OxO668e[8]]+Ox3f1[OxO668e[0]],Ox19[OxO668e[12]]+Ox13a);if(Ox415==0){Ox415=Ox416[OxO668e[0]];} ;Element_RemoveNode(Ox416,true);} ;Ox3f1[OxO668e[0]]=Ox3f1[OxO668e[0]]+Ox415;for(var Ox406=0;Ox406<Ox3f1[OxO668e[0]];Ox406++){for(var Ox40b=0;Ox40b<Ox3f1[OxO668e[4]];Ox40b++){Ox11d[OxO668e[7]][Ox19[OxO668e[8]]+Ox406][Ox19[OxO668e[12]]+Ox40b]=Ox3f1;} ;} ;Table_CleanUpTableInfo(Ox3f7,Ox11d);} ;function Table_AddColSpan(Ox3f1){var Ox3f7=Table_GetTable(Ox3f1);var Ox11d=Table_CalculateTableInfo(Ox3f7);var Ox3ec=Ox11d[OxO668e[7]];var Ox19=Table_GetCellPositionFromMap(Ox3ec,Ox3f1);var Ox41f=0;for(var Oxb3=0;Oxb3<Ox3f1[OxO668e[0]];Oxb3++){var Ox416=Table_GetCellFromMap(Ox3ec,Ox19[OxO668e[8]]+Oxb3,Ox19[OxO668e[12]]+Ox3f1[OxO668e[4]]);if(Ox41f==0){Ox41f=Ox416[OxO668e[4]];} ;Element_RemoveNode(Ox416,true);} ;Ox3f1[OxO668e[4]]+=Ox41f;for(var Ox406=0;Ox406<Ox3f1[OxO668e[0]];Ox406++){for(var Ox40b=0;Ox40b<Ox3f1[OxO668e[4]];Ox40b++){Ox11d[OxO668e[7]][Ox19[OxO668e[8]]+Ox406][Ox19[OxO668e[12]]+Ox40b]=Ox3f1;} ;} ;Table_CleanUpTableInfo(Ox3f7,Ox11d);} ;function Table_SubCol(Ox3f1){var Ox3f7=Table_GetTable(Ox3f1);var Ox11d=Table_CalculateTableInfo(Ox3f7);var Ox3ec=Ox11d[OxO668e[7]];var Ox404=Ox11d[OxO668e[18]];var Ox403=Ox11d[OxO668e[17]];if(Ox403==0){return ;} ;var Ox41d={};var Ox40b=Ox404-1;for(var Ox406=0;Ox406<Ox403;Ox406++){var Ox3f1=Table_GetCellFromMap(Ox3ec,Ox406,Ox40b);if(Ox41d[Ox3f1[OxO668e[24]]]){continue ;} ;Ox41d[Ox3f1[OxO668e[24]]]=true;if(Ox3f1[OxO668e[4]]==1){Element_RemoveNode(Ox3f1,true);} else {Ox3f1[OxO668e[4]]=Ox3f1[OxO668e[4]]-1;} ;} ;Ox11d[OxO668e[18]]--;Table_CleanUpTableInfo(Ox3f7,Ox11d);} ;var richDropDown=Window_GetElement(window,OxO668e[25],true);var list_Templates=Window_GetElement(window,OxO668e[26],true);var subcolumns=Window_GetElement(window,OxO668e[27],true);var addcolumns=Window_GetElement(window,OxO668e[28],true);var subcolspan=Window_GetElement(window,OxO668e[29],true);var addcolspan=Window_GetElement(window,OxO668e[30],true);var btn_row_dialog=Window_GetElement(window,OxO668e[31],true);var btn_cell_dialog=Window_GetElement(window,OxO668e[32],true);var inp_cell_width=Window_GetElement(window,OxO668e[33],true);var inp_cell_height=Window_GetElement(window,OxO668e[34],true);var btn_cell_editcell=Window_GetElement(window,OxO668e[35],true);var tabledesign=Window_GetElement(window,OxO668e[36],true);var subrows=Window_GetElement(window,OxO668e[37],true);var addrows=Window_GetElement(window,OxO668e[38],true);var subrowspan=Window_GetElement(window,OxO668e[39],true);var addrowspan=Window_GetElement(window,OxO668e[40],true);btn_cell_editcell[OxO668e[42]][OxO668e[41]]=OxO668e[43];UpdateState=function UpdateState_InsertTable(){btn_cell_editcell[OxO668e[44]]=btn_row_dialog[OxO668e[44]]=btn_cell_dialog[OxO668e[44]]=GetElementCell()==null;} ;SyncToView=function SyncToView_InsertTable(){var Ox435=GetElementCell();if(Ox435){inp_cell_width[OxO668e[45]]=Ox435[OxO668e[46]];inp_cell_height[OxO668e[45]]=Ox435[OxO668e[47]];} ;} ;SyncTo=function SyncTo_InsertTable(element){var Ox435=GetElementCell();if(Ox435){try{Ox435[OxO668e[46]]=inp_cell_width[OxO668e[45]];Ox435[OxO668e[47]]=inp_cell_height[OxO668e[45]];} catch(er){alert(CE_GetStr(OxO668e[48]));} ;} ;} ;function selectTemplate(Oxaf){var Ox438=OxO668e[49];switch(Oxaf){case 1:Ox438=OxO668e[50];break ;;case 2:Ox438=OxO668e[51];break ;;case 3:Ox438=OxO668e[52];break ;;case 4:Ox438=OxO668e[53];Ox438=Ox438+OxO668e[54];Ox438=Ox438+OxO668e[55];break ;;case 5:Ox438=OxO668e[53];Ox438=Ox438+OxO668e[56];break ;;case 6:Ox438=OxO668e[57];Ox438=Ox438+OxO668e[58];Ox438=Ox438+OxO668e[59];break ;;default:break ;;} ;try{var div=document.createElement(OxO668e[60]);div[OxO668e[13]]=Ox438;var Ox3f7=div.children(0);ApplyTable(Ox3f7,element);ApplyTable(Ox3f7,tabledesign);HandleTableChanged();hidePopup();} catch(x){} ;} ;subcolumns[OxO668e[61]]=function subcolumns_onclick(){Table_SubCol(GetElementCell());Table_SubCol(currentcell);HandleTableChanged();} ;addcolumns[OxO668e[61]]=function addcolumns_onclick(){Table_AddCol(GetElementCell());Table_AddCol(currentcell);HandleTableChanged();} ;subrows[OxO668e[61]]=function subrows_onclick(){Table_SubRow(GetElementCell());Table_SubRow(currentcell);HandleTableChanged();} ;addrows[OxO668e[61]]=function addrows_onclick(){Table_AddRow(currentcell);Table_AddRow(GetElementCell());HandleTableChanged();} ;subcolspan[OxO668e[61]]=function subcolspan_onclick(){Table_SubColSpan(GetElementCell());Table_SubColSpan(currentcell);HandleTableChanged();} ;addcolspan[OxO668e[61]]=function addcolspan_onclick(){Table_AddColSpan(GetElementCell());Table_AddColSpan(currentcell);HandleTableChanged();} ;subrowspan[OxO668e[61]]=function subrowspan_onclick(){Table_SubRowSpan(GetElementCell());Table_SubRowSpan(currentcell);HandleTableChanged();} ;addrowspan[OxO668e[61]]=function addrowspan_onclick(){Table_AddRowSpan(GetElementCell());Table_AddRowSpan(currentcell);HandleTableChanged();} ;function InitDesignTableCells(){for(var Oxb3=0;Oxb3<tabledesign[OxO668e[9]][OxO668e[10]];Oxb3++){var Ox402=tabledesign[OxO668e[9]][Oxb3];for(var Ox13a=0;Ox13a<Ox402[OxO668e[11]][OxO668e[10]];Ox13a++){var Ox3f1=Ox402[OxO668e[11]][Ox13a];Ox3f1.removeAttribute(OxO668e[46]);Ox3f1.removeAttribute(OxO668e[47]);Ox3f1[OxO668e[46]]=OxO668e[49];Ox3f1[OxO668e[47]]=OxO668e[49];Ox3f1[OxO668e[62]]=OxO668e[63];Ox3f1[OxO668e[64]]=cell_mouseover;Ox3f1[OxO668e[65]]=cell_mouseout;Ox3f1[OxO668e[61]]=cell_click;Ox3f1[OxO668e[66]]=cell_dblclick;} ;} ;} ;function Element_Contains(element,Ox6e){if(!Browser_IsOpera()){if(element[OxO668e[67]]){return element.contains(Ox6e);} ;} ;for(;Ox6e!=null;Ox6e=Ox6e[OxO668e[2]]){if(element==Ox6e){return true;} ;} ;return false;} ;function HandleTableChanged(){if(!Element_Contains(tabledesign,currentcell)){SetCurrentCell(tabledesign.rows(0).cells(0));} ;InitDesignTableCells();UpdateButtonStates();editor.ExecCommand(OxO668e[68]);editor.ExecCommand(OxO668e[68]);} ;function cell_mouseover(){var Ox3f1=this;Ox3f1[OxO668e[42]][OxO668e[69]]=OxO668e[70];} ;function cell_mouseout(){var Ox3f1=this;Ox3f1[OxO668e[42]][OxO668e[69]]=OxO668e[63];} ;function cell_click(){var Ox3f1=this;SetCurrentCell(Ox3f1);} ;function cell_dblclick(){var Ox3f1=this;SetCurrentCell(Ox3f1);} ;btn_cell_editcell[OxO668e[61]]=function btn_cell_editcell_onclick(){var Ox3f1=GetElementCell();var Ox17c=editor.EditInWindow(Ox3f1.innerHTML,window);if(Ox17c!=null&&Ox17c!==false){Ox3f1[OxO668e[13]]=Ox17c;} ;} ;btn_cell_dialog[OxO668e[61]]=function btn_cell_dialog_onclick(){editor.SetNextDialogWindow(window);editor.ShowTagDialogWithNoCancellable(GetElementCell());} ;btn_row_dialog[OxO668e[61]]=function btn_row_dialog_onclick(){editor.SetNextDialogWindow(window);editor.ShowTagDialogWithNoCancellable(GetElementCell().parentNode);} ;btn_row_dialog[OxO668e[42]][OxO668e[41]]=btn_cell_dialog[OxO668e[42]][OxO668e[41]]=OxO668e[43];var currentcell=null;function GetElementCell(){if(currentcell==null){return null;} ;return element[OxO668e[9]][currentcell[OxO668e[2]][OxO668e[8]]][OxO668e[11]][currentcell[OxO668e[12]]];} ;function SetCurrentCell(Ox3f1){if(currentcell!=null){if(Browser_IsWinIE()){currentcell[OxO668e[72]][OxO668e[71]]=OxO668e[49];} else {currentcell[OxO668e[42]][OxO668e[71]]=OxO668e[49];} ;} ;currentcell=Ox3f1;if(Browser_IsWinIE()){currentcell[OxO668e[72]][OxO668e[71]]=OxO668e[73];} else {currentcell[OxO668e[42]][OxO668e[71]]=OxO668e[73];} ;UpdateButtonStates();SyncToViewInternal();} ;function UpdateButtonStates(){subcolspan[OxO668e[44]]=!Table_CanSubColSpan(currentcell);addcolspan[OxO668e[44]]=!Table_CanAddColSpan(currentcell);subrowspan[OxO668e[44]]=!Table_CanSubRowSpan(currentcell);addrowspan[OxO668e[44]]=!Table_CanAddRowSpan(currentcell);subrows[OxO668e[44]]=Table_GetRowCount(currentcell)<2;subcolumns[OxO668e[44]]=Table_GetColCount(currentcell)<2;} ;function ApplyTable(src,Ox44f){if(Browser_IsSafari()){Ox44f[OxO668e[42]][OxO668e[47]]=OxO668e[49];} ;for(var Oxb3=Ox44f[OxO668e[9]][OxO668e[10]]-1;Oxb3>=0;Oxb3--){Ox44f[OxO668e[9]][Oxb3].removeNode(true);} ;for(var Oxb3=0;Oxb3<src[OxO668e[9]][OxO668e[10]];Oxb3++){var Ox450=src[OxO668e[9]][Oxb3];var Ox451;if(!Browser_IsSafari()){Ox451=Ox44f.insertRow(-1);} else {var Ox116=document.createElement(OxO668e[23]);Ox44f.appendChild(Ox116);Ox451=Ox116.insertRow(-1);} ;Ox451[OxO668e[42]][OxO668e[71]]=Ox450[OxO668e[42]][OxO668e[71]];for(var Ox13a=0;Ox13a<Ox450[OxO668e[11]][OxO668e[10]];Ox13a++){var Ox452=Ox450[OxO668e[11]][Ox13a];var Ox453=Ox451.insertCell(-1);Ox453[OxO668e[0]]=Ox452[OxO668e[0]];Ox453[OxO668e[4]]=Ox452[OxO668e[4]];Ox453[OxO668e[42]][OxO668e[71]]=Ox452[OxO668e[42]][OxO668e[71]];if(Browser_IsOpera()){Ox453[OxO668e[13]]=OxO668e[14];} else {if(Browser_IsGecko()||Browser_IsSafari()){Ox453[OxO668e[13]]=OxO668e[15];} ;} ;} ;} ;} ;window[OxO668e[74]]=function window_onload(){ApplyTable(element,tabledesign);InitDesignTableCells();SetCurrentCell(tabledesign[OxO668e[9]][0][OxO668e[11]][0]);} ;function updateList(){} ;var oPopup;if(Browser_IsWinIE()){oPopup=window.createPopup();} else {richDropDown[OxO668e[42]][OxO668e[41]]=OxO668e[43];} ;function selectTemplates(){if(Browser_IsWinIE()){oPopup[OxO668e[76]][OxO668e[75]][OxO668e[13]]=list_Templates[OxO668e[13]];oPopup[OxO668e[76]][OxO668e[75]][OxO668e[42]][OxO668e[77]]=OxO668e[78];oPopup.show(0,32,380,220,richDropDown);} ;} ;function hidePopup(){if(Browser_IsWinIE()){if(oPopup){if(oPopup[OxO668e[79]]){oPopup.hide();} ;} ;} ;} ;