var OxO4e7b=["stringSearch","stringReplace","MatchWholeWord","MatchCase","document","checked","length","value","Nothing to search.","selection","body","type","Control","rangeCount","userAgent","innerText","text","Finished Searching the document. Would you like to start again from the top?","","textedit"," : ","Please use replace function."];var editwin=Window_GetDialogArguments(window);var stringSearch=Window_GetElement(window,OxO4e7b[0],true);var stringReplace=Window_GetElement(window,OxO4e7b[1],true);var MatchWholeWord=Window_GetElement(window,OxO4e7b[2],true);var MatchCase=Window_GetElement(window,OxO4e7b[3],true);var editdoc=editwin[OxO4e7b[4]];function get_ie_matchtype(){var Ox213=0;var Ox214=0;var Ox215=0;if(MatchCase[OxO4e7b[5]]){Ox214=4;} ;if(MatchWholeWord[OxO4e7b[5]]){Ox215=2;} ;Ox213=Ox214+Ox215;return (Ox213);} ;function checkInputString(){if(stringSearch[OxO4e7b[7]][OxO4e7b[6]]<1){alert(OxO4e7b[8]);return false;} else {return true;} ;} ;function IsMatchSearchValue(Ox61){if(!Ox61){return false;} ;if(stringSearch[OxO4e7b[7]]==Ox61){return true;} ;if(MatchCase[OxO4e7b[5]]){return false;} ;return stringSearch[OxO4e7b[7]].toLowerCase()==Ox61.toLowerCase();} ;var _ie_range=null;function IE_Restore(){editwin.focus();if(_ie_range!=null){_ie_range.select();} ;} ;function IE_Save(){editwin.focus();_ie_range=editdoc[OxO4e7b[9]].createRange();} ;function MoveToBodyStart(){if(Browser_UseIESelection()){range=document[OxO4e7b[10]].createTextRange();range.collapse(true);range.select();IE_Save();} else {editwin.getSelection().collapse(editdoc.body,0);} ;} ;function DoFind(){if(Browser_UseIESelection()){IE_Restore();var Ox17=editdoc[OxO4e7b[9]];if(Ox17[OxO4e7b[11]]==OxO4e7b[12]){MoveToBodyStart();} ;var Ox119=Ox17.createRange();Ox119.collapse(false);if(Ox119.findText(stringSearch.value,1000000000,get_ie_matchtype())){Ox119.select();IE_Save();return true;} ;} else {var Ox119;var Ox17=editwin.getSelection();if(Ox17[OxO4e7b[13]]>0){Ox119=editwin.getSelection().getRangeAt(0);} ;var Ox11e=!!navigator[OxO4e7b[14]].match(/Trident\/7\./);if(Ox11e){editdoc[OxO4e7b[10]][OxO4e7b[15]].indexOf(stringSearch.value)>-1;} else {if(editwin.find(stringSearch.value,MatchCase.checked,false,false,MatchWholeWord.checked,false,false)){return true;} ;} ;} ;} ;function DoReplace(){if(Browser_UseIESelection()){IE_Restore();var Ox17=editdoc[OxO4e7b[9]];if(Ox17[OxO4e7b[11]]!=OxO4e7b[12]){var Ox119=Ox17.createRange();if(IsMatchSearchValue(Ox119.text)){Ox119[OxO4e7b[16]]=stringReplace[OxO4e7b[7]];Ox119.collapse(false);IE_Save();return true;} ;} ;} else {var Ox17=editwin.getSelection();if(IsMatchSearchValue(Ox17.toString())){Ox17.deleteFromDocument();Ox17.getRangeAt(0).insertNode(editdoc.createTextNode(stringReplace.value));Ox17.getRangeAt(0).collapse(false);return true;} ;} ;return false;} ;function FindTxt(){if(!checkInputString()){return false;} ;while(true){if(DoFind()){return ;} ;if(!confirm(OxO4e7b[17])){return ;} ;MoveToBodyStart();} ;} ;function ReplaceTxt(){if(!checkInputString()){return ;} ;DoReplace();FindTxt();} ;function ReplaceAllTxt(){if(!checkInputString()){return ;} ;var Ox221=0;var msg=OxO4e7b[18];MoveToBodyStart();if(Browser_UseIESelection()){var Ox17=editdoc[OxO4e7b[9]];if(Ox17[OxO4e7b[11]]==OxO4e7b[12]){MoveToBodyStart();} ;var Ox222=Ox17.createRange();var Ox221=0;var msg=OxO4e7b[18];Ox222.expand(OxO4e7b[19]);Ox222.collapse();Ox222.select();while(Ox222.findText(stringSearch.value,1000000000,get_ie_matchtype())){Ox222.select();Ox222[OxO4e7b[16]]=stringReplace[OxO4e7b[7]];Ox221++;} ;if(Ox221==0){msg=WordNotFound;} else {msg=WordReplaced+OxO4e7b[20]+Ox221;} ;alert(msg);} else {if((stringReplace[OxO4e7b[7]]).indexOf(stringSearch.value)==-1){DoFind();while(DoReplace()){Ox221++;DoFind();FindTxt();} ;if(Ox221==0){msg=WordNotFound;} else {msg=WordReplaced+OxO4e7b[20]+Ox221;} ;alert(msg);} else {FindTxt();alert(OxO4e7b[21]);} ;} ;} ;