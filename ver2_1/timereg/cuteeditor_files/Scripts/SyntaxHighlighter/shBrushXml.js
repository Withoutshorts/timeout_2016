var OxOf66b=["Xml","Brushes","sh","CssClass","dp-xml","Style",".dp-xml .cdata { color: #ff1493; }",".dp-xml .tag, .dp-xml .tag-name { color: #069; font-weight: bold; }",".dp-xml .attribute { color: red; }",".dp-xml .attribute-value { color: blue; }","prototype","Aliases","xml","xhtml","xslt","html","ProcessRegexList","length","(\x26lt;|\x3C)\x5C!\x5C[[\x5Cw\x5Cs]*?\x5C[(.|\x5Cs)*?\x5C]\x5C](\x26gt;|\x3E)","gm","cdata","(\x26lt;|\x3C)!--\x5Cs*.*?\x5Cs*--(\x26gt;|\x3E)","comments","([:\x5Cw-.]+)\x5Cs*=\x5Cs*(\x22.*?\x22|\x27.*?\x27|\x5Cw+)*|(\x5Cw+)","attribute","index","attribute-value","(\x26lt;|\x3C)/*\x5C?*(?!\x5C!)|/*\x5C?*(\x26gt;|\x3E)","tag","(?:\x26lt;|\x3C)/*\x5C?*\x5Cs*([:\x5Cw-.]+)","tag-name"];dp[OxOf66b[2]][OxOf66b[1]][OxOf66b[0]]=function (){this[OxOf66b[3]]=OxOf66b[4];this[OxOf66b[5]]=OxOf66b[6]+OxOf66b[7]+OxOf66b[8]+OxOf66b[9];} ;dp[OxOf66b[2]][OxOf66b[1]][OxOf66b[0]][OxOf66b[10]]= new dp[OxOf66b[2]].Highlighter();dp[OxOf66b[2]][OxOf66b[1]][OxOf66b[0]][OxOf66b[11]]=[OxOf66b[12],OxOf66b[13],OxOf66b[14],OxOf66b[15],OxOf66b[13]];dp[OxOf66b[2]][OxOf66b[1]][OxOf66b[0]][OxOf66b[10]][OxOf66b[16]]=function (){function Oxacb(Oxacc,Ox4f){Oxacc[Oxacc[OxOf66b[17]]]=Ox4f;} ;var Oxed=0;var Ox836=null;var Oxacd=null;this.GetMatches( new RegExp(OxOf66b[18],OxOf66b[19]),OxOf66b[20]);this.GetMatches( new RegExp(OxOf66b[21],OxOf66b[19]),OxOf66b[22]);Oxacd= new RegExp(OxOf66b[23],OxOf66b[19]);while((Ox836=Oxacd.exec(this.code))!=null){if(Ox836[1]==null){continue ;} ;Oxacb(this.matches, new dp[OxOf66b[2]].Match(Ox836[1],Ox836.index,OxOf66b[24]));if(Ox836[2]!=undefined){Oxacb(this.matches, new dp[OxOf66b[2]].Match(Ox836[2],Ox836[OxOf66b[25]]+Ox836[0].indexOf(Ox836[2]),OxOf66b[26]));} ;} ;this.GetMatches( new RegExp(OxOf66b[27],OxOf66b[19]),OxOf66b[28]);Oxacd= new RegExp(OxOf66b[29],OxOf66b[19]);while((Ox836=Oxacd.exec(this.code))!=null){Oxacb(this.matches, new dp[OxOf66b[2]].Match(Ox836[1],Ox836[OxOf66b[25]]+Ox836[0].indexOf(Ox836[1]),OxOf66b[30]));} ;} ;