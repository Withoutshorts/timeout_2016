var OxOec46=["outer","btnbrowse","inp_src","onclick","value","cssText","style","src","FileName"];var outer=Window_GetElement(window,OxOec46[0x0],true);var btnbrowse=Window_GetElement(window,OxOec46[0x1],true);var inp_src=Window_GetElement(window,OxOec46[0x2],true); btnbrowse[OxOec46[0x3]]=function btnbrowse_onclick(){ function Ox2b3(Oxe6){if(Oxe6){ inp_src[OxOec46[0x4]]=Oxe6 ;} ;}  ; editor.SetNextDialogWindow(window) ; editor.ShowSelectFileDialog(Ox2b3,inp_src.value) ;}  ; UpdateState=function UpdateState_Media(){ outer[OxOec46[0x6]][OxOec46[0x5]]=element[OxOec46[0x6]][OxOec46[0x5]] ; outer.mergeAttributes(element) ;if(element[OxOec46[0x7]]){ outer[OxOec46[0x8]]=element[OxOec46[0x8]] ;} else { outer.removeAttribute(OxOec46[0x8]) ;} ;}  ; SyncToView=function SyncToView_Media(){ inp_src[OxOec46[0x4]]=element[OxOec46[0x8]] ;}  ; SyncTo=function SyncTo_Media(element){ element[OxOec46[0x8]]=inp_src[OxOec46[0x4]] ;}  ;