var OxOe815=["onload","contentWindow","idSource","innerHTML","body","document","","designMode","on","contentEditable","fontFamily","style","Tahoma","fontSize","11px","color","black","background","white","length","\x22","\x3C$1$3"," ","\x26nbsp;","$1","\x3Ch$1\x3E","\x3C$1\x3E$2\x3C/$1\x3E","\x27"];var editor=Window_GetDialogArguments(window);function cancel(){Window_CloseDialog(window);} ;window[OxOe815[0]]=function (){var iframe=document.getElementById(OxOe815[2])[OxOe815[1]];iframe[OxOe815[5]][OxOe815[4]][OxOe815[3]]=OxOe815[6];iframe[OxOe815[5]][OxOe815[7]]=OxOe815[8];iframe[OxOe815[5]][OxOe815[4]][OxOe815[9]]=true;iframe[OxOe815[5]][OxOe815[4]][OxOe815[11]][OxOe815[10]]=OxOe815[12];iframe[OxOe815[5]][OxOe815[4]][OxOe815[11]][OxOe815[13]]=OxOe815[14];iframe[OxOe815[5]][OxOe815[4]][OxOe815[11]][OxOe815[15]]=OxOe815[16];iframe[OxOe815[5]][OxOe815[4]][OxOe815[11]][OxOe815[17]]=OxOe815[18];iframe.focus();} ;function insertContent(){var iframe=document.getElementById(OxOe815[2])[OxOe815[1]];var Ox195=iframe[OxOe815[5]][OxOe815[4]][OxOe815[3]];if(Ox195&&Ox195[OxOe815[19]]>0){editor.PasteHTML(_RemoveWord(Ox195));Window_CloseDialog(window);} ;} ;function _RemoveWord(Ox23c){Ox23c=Ox23c.replace(/<[\/]?(base|meta|link|style|font|st1|shape|path|lock|imagedata|stroke|formulas|xml|del|ins|[ovwxp]:\w+)[^>]*?>/gi,OxOe815[6]);Ox23c=Ox23c.replace(/\s*mso-[^:]+:[^;"]+;?/gi,OxOe815[6]);Ox23c=Ox23c.replace(/<!--[\s\S]*?-->/gi,OxOe815[6]);Ox23c=Ox23c.replace(/\s*MARGIN: 0cm 0cm 0pt\s*;/gi,OxOe815[6]);Ox23c=Ox23c.replace(/\s*MARGIN: 0cm 0cm 0pt\s*"/gi,OxOe815[20]);Ox23c=Ox23c.replace(/\s*TEXT-INDENT: 0cm\s*;/gi,OxOe815[6]);Ox23c=Ox23c.replace(/\s*TEXT-INDENT: 0cm\s*"/gi,OxOe815[20]);Ox23c=Ox23c.replace(/\s*TEXT-ALIGN: [^\s;]+;?"/gi,OxOe815[20]);Ox23c=Ox23c.replace(/\s*PAGE-BREAK-BEFORE: [^\s;]+;?"/gi,OxOe815[20]);Ox23c=Ox23c.replace(/\s*FONT-VARIANT: [^\s;]+;?"/gi,OxOe815[20]);Ox23c=Ox23c.replace(/\s*tab-stops:[^;"]*;?/gi,OxOe815[6]);Ox23c=Ox23c.replace(/\s*tab-stops:[^"]*/gi,OxOe815[6]);Ox23c=Ox23c.replace(/<(\w[^>]*) class=([^ |>]*)([^>]*)/gi,OxOe815[21]);Ox23c=Ox23c.replace(/\s*style="\s*"/gi,OxOe815[6]);Ox23c=Ox23c.replace(/<SPAN\s*[^>]*>\s* \s*<\/SPAN>/gi,OxOe815[22]);Ox23c=Ox23c.replace(/<(\w+)[^>]*\sstyle="[^"]*DISPLAY\s?:\s?none(.*?)<\/\1>/ig,OxOe815[6]);Ox23c=Ox23c.replace(/<span\s*[^>]*>\s*&nbsp;\s*<\/span>/gi,OxOe815[23]);Ox23c=Ox23c.replace(/<SPAN\s*[^>]*><\/SPAN>/gi,OxOe815[6]);Ox23c=Ox23c.replace(/<(\w[^>]*) lang=([^ |>]*)([^>]*)/gi,OxOe815[21]);Ox23c=Ox23c.replace(/<SPAN\s*>(.*?)<\/SPAN>/gi,OxOe815[24]);Ox23c=Ox23c.replace(/<\/?\w+:[^>]*>/gi,OxOe815[6]);Ox23c=Ox23c.replace(/<\!--.*?-->/g,OxOe815[6]);Ox23c=Ox23c.replace(/<H\d>\s*<\/H\d>/gi,OxOe815[6]);Ox23c=Ox23c.replace(/<(\w[^>]*) language=([^ |>]*)([^>]*)/gi,OxOe815[21]);Ox23c=Ox23c.replace(/<(\w[^>]*) onmouseover="([^\"]*)"([^>]*)/gi,OxOe815[21]);Ox23c=Ox23c.replace(/<(\w[^>]*) onmouseout="([^\"]*)"([^>]*)/gi,OxOe815[21]);Ox23c=Ox23c.replace(/<H(\d)([^>]*)>/gi,OxOe815[25]);Ox23c=Ox23c.replace(/<(H\d)><FONT[^>]*>(.*?)<\/FONT><\/\1>/gi,OxOe815[26]);Ox23c=Ox23c.replace(/<(H\d)><EM>(.*?)<\/EM><\/\1>/gi,OxOe815[26]);Ox23c=Ox23c.replace(/<a name="?OLE_LINK\d+"?>((.|[\r\n])*?)<\/a>/gi,OxOe815[24]);Ox23c=Ox23c.replace(/<a name="?_Hlt\d+"?>((.|[\r\n])*?)<\/a>/gi,OxOe815[24]);Ox23c=Ox23c.replace(/<a name="?_Toc\d+"?>((.|[\r\n])*?)<\/a>/gi,OxOe815[24]);Ox23c=Ox23c.replace(/[\\]/gi,OxOe815[20]);Ox23c=Ox23c.replace(/[\\]/gi,OxOe815[27]);return Ox23c;} ;