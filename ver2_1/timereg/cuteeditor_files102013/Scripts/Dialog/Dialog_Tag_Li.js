var OxO11bf=["inp_src","box1","box2","box3","box4","box5","box6","box7","box8","box9","inp_start","CustomBullet","nodeName","LI","parentNode","none","decimal","upper-roman","upper-alpha","lower-alpha","lower-roman","disc","circle","square","listStyleType","style","border","solid 2px #708090","listStyleImage","","value","visibility","hidden","length","start","url(\x27","\x27)","visible","UL","OL","document","firstChild","element","solid 2px #ffffff","solid 2px #ECECF6","onclick"];var inp_src=Window_GetElement(window,OxO11bf[0],true);var box1=Window_GetElement(window,OxO11bf[1],true);var box2=Window_GetElement(window,OxO11bf[2],true);var box3=Window_GetElement(window,OxO11bf[3],true);var box4=Window_GetElement(window,OxO11bf[4],true);var box5=Window_GetElement(window,OxO11bf[5],true);var box6=Window_GetElement(window,OxO11bf[6],true);var box7=Window_GetElement(window,OxO11bf[7],true);var box8=Window_GetElement(window,OxO11bf[8],true);var box9=Window_GetElement(window,OxO11bf[9],true);var inp_start=Window_GetElement(window,OxO11bf[10],true);var CustomBullet=Window_GetElement(window,OxO11bf[11],true);OriginalnodeName=element[OxO11bf[12]];if(element[OxO11bf[12]]&&element[OxO11bf[12]]==OxO11bf[13]){OriginalnodeName=(element[OxO11bf[14]])[OxO11bf[12]];} ;var OriginalnodeName,CurrentnodeName,selectedObject;SyncToView=function SyncToView_LI(){if(element[OxO11bf[12]]==OxO11bf[13]){element=element[OxO11bf[14]];} ;switch((element[OxO11bf[25]][OxO11bf[24]]).toLowerCase()){case OxO11bf[15]:selectedObject=box1;break ;;case OxO11bf[16]:selectedObject=box2;break ;;case OxO11bf[17]:selectedObject=box3;break ;;case OxO11bf[18]:selectedObject=box4;break ;;case OxO11bf[19]:selectedObject=box5;break ;;case OxO11bf[20]:selectedObject=box6;break ;;case OxO11bf[21]:selectedObject=box7;break ;;case OxO11bf[22]:selectedObject=box8;break ;;case OxO11bf[23]:selectedObject=box9;break ;;default:selectedObject=box1;break ;;} ;selectedObject[OxO11bf[25]][OxO11bf[26]]=OxO11bf[27];if(element[OxO11bf[25]][OxO11bf[28]]==OxO11bf[29]){inp_src[OxO11bf[30]]=OxO11bf[29];CustomBullet[OxO11bf[25]][OxO11bf[31]]=OxO11bf[32];} else {var Ox2a;Ox2a=element[OxO11bf[25]][OxO11bf[28]];Ox2a=Ox2a.substring(4,Ox2a[OxO11bf[33]]-1);inp_src[OxO11bf[30]]=Ox2a;} ;} ;SyncTo=function SyncTo_LI(element){switch(selectedObject){case box1:;case box2:;case box3:;case box4:;case box5:;case box6:try{element.setAttribute(OxO11bf[34],inp_start.value);} catch(er){} ;break ;;case box7:;case box8:;case box9:break ;;} ;if(inp_src[OxO11bf[30]]){element[OxO11bf[25]][OxO11bf[28]]=OxO11bf[35]+inp_src[OxO11bf[30]]+OxO11bf[36];} ;} ;function ToggleCustomBullet(){if(CustomBullet[OxO11bf[25]][OxO11bf[31]]==OxO11bf[32]){CustomBullet[OxO11bf[25]][OxO11bf[31]]=OxO11bf[37];} else {CustomBullet[OxO11bf[25]][OxO11bf[31]]=OxO11bf[32];} ;} ;function doClick1(Ox273){if(element[OxO11bf[12]]==OxO11bf[13]){element=element[OxO11bf[14]];} ;selectedObject=Ox273;switch(selectedObject){case box1:element[OxO11bf[25]][OxO11bf[24]]=OxO11bf[15];break ;;case box2:element[OxO11bf[25]][OxO11bf[24]]=OxO11bf[16];break ;;case box3:element[OxO11bf[25]][OxO11bf[24]]=OxO11bf[17];break ;;case box4:element[OxO11bf[25]][OxO11bf[24]]=OxO11bf[18];break ;;case box5:element[OxO11bf[25]][OxO11bf[24]]=OxO11bf[19];break ;;case box6:element[OxO11bf[25]][OxO11bf[24]]=OxO11bf[20];break ;;case box7:element[OxO11bf[25]][OxO11bf[24]]=OxO11bf[21];break ;;case box8:element[OxO11bf[25]][OxO11bf[24]]=OxO11bf[22];break ;;case box9:element[OxO11bf[25]][OxO11bf[24]]=OxO11bf[23];break ;;} ;var Ox301=false;switch(selectedObject){case box1:;case box2:;case box3:;case box4:;case box5:;case box6:if(OriginalnodeName==OxO11bf[38]){OriginalnodeName=OxO11bf[39];Ox301=true;} ;break ;;case box7:;case box8:;case box9:if(OriginalnodeName==OxO11bf[39]){OriginalnodeName=OxO11bf[38];Ox301=true;} ;break ;;} ;if(Ox301){var Ox469=editwin[OxO11bf[40]].createElement(OriginalnodeName);while(element[OxO11bf[41]]){Ox469.appendChild(element.firstChild);} ;element[OxO11bf[14]].insertBefore(Ox469,element);element[OxO11bf[14]].removeChild(element);var arg=Window_FindDialogArguments(window);arg[OxO11bf[42]]=element=Ox469;} ;box1[OxO11bf[25]][OxO11bf[26]]=OxO11bf[43];box2[OxO11bf[25]][OxO11bf[26]]=OxO11bf[43];box3[OxO11bf[25]][OxO11bf[26]]=OxO11bf[43];box4[OxO11bf[25]][OxO11bf[26]]=OxO11bf[43];box5[OxO11bf[25]][OxO11bf[26]]=OxO11bf[43];box6[OxO11bf[25]][OxO11bf[26]]=OxO11bf[43];box7[OxO11bf[25]][OxO11bf[26]]=OxO11bf[43];box8[OxO11bf[25]][OxO11bf[26]]=OxO11bf[43];box9[OxO11bf[25]][OxO11bf[26]]=OxO11bf[43];selectedObject[OxO11bf[25]][OxO11bf[26]]=OxO11bf[27];inp_src[OxO11bf[30]]=OxO11bf[29];SyncTo();} ;function doMouseOut(Ox273){if(Ox273==selectedObject){Ox273[OxO11bf[25]][OxO11bf[26]]=OxO11bf[27];} else {Ox273[OxO11bf[25]][OxO11bf[26]]=OxO11bf[43];} ;} ;function doMouseOver(Ox273){Ox273[OxO11bf[25]][OxO11bf[26]]=OxO11bf[44];} ;btnbrowse[OxO11bf[45]]=function btnbrowse_onclick(){function Ox25f(Ox6){if(Ox6){inp_src[OxO11bf[30]]=Ox6;SyncTo(element);} ;} ;editor.SetNextDialogWindow(window);if(Browser_IsSafari()){editor.ShowSelectImageDialog(Ox25f,inp_src.value,inp_src);} else {editor.ShowSelectImageDialog(Ox25f,inp_src.value);} ;} ;