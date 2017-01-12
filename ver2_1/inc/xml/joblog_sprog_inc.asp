
<% 
Dim objXMLHTTP_joblog, objXMLDOM_joblog, i_joblog, strHTML_joblog

Set objXMLDom_joblog = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_joblog = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_joblog.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/joblog_sprog.xml", False
'objXmlHttp_joblog.open "GET", "http://localhost/inc/xml/grandtotal_sprog.xml", False
'objXmlHttp_joblog.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/joblog_sprog.xml", False
objXmlHttp_joblog.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/joblog_sprog.xml", False
'objXmlHttp_joblog.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/joblog_sprog.xml", False
'objXmlHttp_joblog.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/joblog_sprog.xml", False

objXmlHttp_joblog.send


Set objXmlDom_joblog = objXmlHttp_joblog.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_joblog = Nothing



Dim Address_joblog, Latitude_joblog, Longitude_joblog
Dim oNode_joblog, oNodes_joblog
Dim sXPathQuery_joblog

sprog = 1 'DK
if len(trim(session("mid"))) <> 0 then
strSQL = "SELECT sprog FROM medarbejdere WHERE mid = " & session("mid")

oRec.open strSQL, oConn, 3
if not oRec.EOF then
sprog = oRec("sprog")
end if
oRec.close
end if

select case sprog
case 1
sXPathQuery_joblog = "//sprog/dk"
Session.LCID = 1030
case 2
sXPathQuery_joblog = "//sprog/uk"
Session.LCID = 2057
case 3
sXPathQuery_joblog = "//sprog/se"
Session.LCID = 1053
case 4
sXPathQuery_joblog = "//sprog/no"
Session.LCID = 2068
case 5
sXPathQuery_joblog = "//sprog/es"
Session.LCID = 1034
case 6
sXPathQuery_joblog = "//sprog/de"
Session.LCID = 1031
case 7
sXPathQuery_joblog = "//sprog/fr"
Session.LCID = 1036
case else
sXPathQuery_joblog = "//sprog/dk"
Session.LCID = 1030
end select





Set oNode_joblog = objXmlDom_joblog.documentElement.selectSingleNode(sXPathQuery_joblog)
Address_joblog = oNode_joblog.Text

Set oNodes_joblog = objXmlDom_joblog.documentElement.selectNodes(sXPathQuery_joblog)

    For Each oNode_joblog in oNodes_joblog

    joblog_txt_001 = oNode_joblog.selectSingleNode("txt_1").Text
    joblog_txt_002 = oNode_joblog.selectSingleNode("txt_2").Text
    joblog_txt_003 = oNode_joblog.selectSingleNode("txt_3").Text
    joblog_txt_003 = oNode_joblog.selectSingleNode("txt_3").Text
    joblog_txt_004 = oNode_joblog.selectSingleNode("txt_4").Text
    joblog_txt_005 = oNode_joblog.selectSingleNode("txt_5").Text

   joblog_txt_006 = oNode_joblog.selectSingleNode("txt_6").Text
   joblog_txt_007 = oNode_joblog.selectSingleNode("txt_7").Text
   joblog_txt_008 = oNode_joblog.selectSingleNode("txt_8").Text
   joblog_txt_009 = oNode_joblog.selectSingleNode("txt_9").Text
   joblog_txt_010 = oNode_joblog.selectSingleNode("txt_10").Text
   joblog_txt_011 = oNode_joblog.selectSingleNode("txt_11").Text
    
   joblog_txt_012 = oNode_joblog.selectSingleNode("txt_12").Text
   joblog_txt_013 = oNode_joblog.selectSingleNode("txt_13").Text
   joblog_txt_014 = oNode_joblog.selectSingleNode("txt_14").Text
   joblog_txt_015 = oNode_joblog.selectSingleNode("txt_15").Text
   joblog_txt_016 = oNode_joblog.selectSingleNode("txt_16").Text
   joblog_txt_017 = oNode_joblog.selectSingleNode("txt_17").Text
    
   joblog_txt_018 = oNode_joblog.selectSingleNode("txt_18").Text
   joblog_txt_019 = oNode_joblog.selectSingleNode("txt_19").Text
   joblog_txt_020 = oNode_joblog.selectSingleNode("txt_20").Text
   joblog_txt_021 = oNode_joblog.selectSingleNode("txt_21").Text
   joblog_txt_022 = oNode_joblog.selectSingleNode("txt_22").Text
   joblog_txt_023 = oNode_joblog.selectSingleNode("txt_23").Text
   joblog_txt_024 = oNode_joblog.selectSingleNode("txt_24").Text
    
   joblog_txt_025 = oNode_joblog.selectSingleNode("txt_25").Text
   joblog_txt_026 = oNode_joblog.selectSingleNode("txt_26").Text
   joblog_txt_027 = oNode_joblog.selectSingleNode("txt_27").Text
   joblog_txt_028 = oNode_joblog.selectSingleNode("txt_28").Text
   joblog_txt_029 = oNode_joblog.selectSingleNode("txt_29").Text
   joblog_txt_030 = oNode_joblog.selectSingleNode("txt_30").Text
   joblog_txt_031 = oNode_joblog.selectSingleNode("txt_31").Text
   joblog_txt_032 = oNode_joblog.selectSingleNode("txt_32").Text
   joblog_txt_033 = oNode_joblog.selectSingleNode("txt_33").Text
   joblog_txt_034 = oNode_joblog.selectSingleNode("txt_34").Text
    
    joblog_txt_035 = oNode_joblog.selectSingleNode("txt_35").Text
     joblog_txt_036 = oNode_joblog.selectSingleNode("txt_36").Text
      joblog_txt_037 = oNode_joblog.selectSingleNode("txt_37").Text
       
       joblog_txt_038 = oNode_joblog.selectSingleNode("txt_38").Text
        joblog_txt_039 = oNode_joblog.selectSingleNode("txt_39").Text
        joblog_txt_040 = oNode_joblog.selectSingleNode("txt_40").Text
        joblog_txt_041 = oNode_joblog.selectSingleNode("txt_41").Text
        joblog_txt_042 = oNode_joblog.selectSingleNode("txt_42").Text
         joblog_txt_043 = oNode_joblog.selectSingleNode("txt_43").Text
         joblog_txt_044 = oNode_joblog.selectSingleNode("txt_44").Text
          
         joblog_txt_045 = oNode_joblog.selectSingleNode("txt_45").Text
         joblog_txt_046 = oNode_joblog.selectSingleNode("txt_46").Text
         joblog_txt_047 = oNode_joblog.selectSingleNode("txt_47").Text
         joblog_txt_048 = oNode_joblog.selectSingleNode("txt_48").Text
         joblog_txt_049 = oNode_joblog.selectSingleNode("txt_49").Text
         joblog_txt_050 = oNode_joblog.selectSingleNode("txt_50").Text
         joblog_txt_051 = oNode_joblog.selectSingleNode("txt_51").Text
         joblog_txt_052 = oNode_joblog.selectSingleNode("txt_52").Text
         joblog_txt_053 = oNode_joblog.selectSingleNode("txt_53").Text
         joblog_txt_054 = oNode_joblog.selectSingleNode("txt_54").Text
         joblog_txt_055 = oNode_joblog.selectSingleNode("txt_55").Text
         joblog_txt_056 = oNode_joblog.selectSingleNode("txt_56").Text
         joblog_txt_057 = oNode_joblog.selectSingleNode("txt_57").Text
        
        joblog_txt_058 = oNode_joblog.selectSingleNode("txt_58").Text
        joblog_txt_059 = oNode_joblog.selectSingleNode("txt_59").Text
    

   joblog_txt_060 = oNode_joblog.selectSingleNode("txt_60").Text
   joblog_txt_061 = oNode_joblog.selectSingleNode("txt_61").Text
   joblog_txt_062 = oNode_joblog.selectSingleNode("txt_62").Text
   joblog_txt_063 = oNode_joblog.selectSingleNode("txt_63").Text
   joblog_txt_064 = oNode_joblog.selectSingleNode("txt_64").Text
   joblog_txt_065 = oNode_joblog.selectSingleNode("txt_65").Text
   joblog_txt_066 = oNode_joblog.selectSingleNode("txt_66").Text
   joblog_txt_067 = oNode_joblog.selectSingleNode("txt_67").Text
   joblog_txt_068 = oNode_joblog.selectSingleNode("txt_68").Text
   joblog_txt_069 = oNode_joblog.selectSingleNode("txt_69").Text

   joblog_txt_070 = oNode_joblog.selectSingleNode("txt_70").Text
   joblog_txt_071 = oNode_joblog.selectSingleNode("txt_71").Text
   joblog_txt_072 = oNode_joblog.selectSingleNode("txt_72").Text
   joblog_txt_070 = oNode_joblog.selectSingleNode("txt_73").Text
   joblog_txt_074 = oNode_joblog.selectSingleNode("txt_74").Text
   joblog_txt_075 = oNode_joblog.selectSingleNode("txt_75").Text
   joblog_txt_076 = oNode_joblog.selectSingleNode("txt_76").Text
   joblog_txt_077 = oNode_joblog.selectSingleNode("txt_77").Text
   joblog_txt_078 = oNode_joblog.selectSingleNode("txt_78").Text
   joblog_txt_079 = oNode_joblog.selectSingleNode("txt_79").Text
    
   joblog_txt_080 = oNode_joblog.selectSingleNode("txt_80").Text
   joblog_txt_081 = oNode_joblog.selectSingleNode("txt_81").Text
   joblog_txt_082 = oNode_joblog.selectSingleNode("txt_82").Text
   joblog_txt_083 = oNode_joblog.selectSingleNode("txt_83").Text
   joblog_txt_084 = oNode_joblog.selectSingleNode("txt_84").Text
   joblog_txt_085 = oNode_joblog.selectSingleNode("txt_85").Text
   joblog_txt_086 = oNode_joblog.selectSingleNode("txt_86").Text
   joblog_txt_087 = oNode_joblog.selectSingleNode("txt_87").Text

    joblog_txt_088 = oNode_joblog.selectSingleNode("txt_88").Text
    joblog_txt_089 = oNode_joblog.selectSingleNode("txt_89").Text
    joblog_txt_090 = oNode_joblog.selectSingleNode("txt_90").Text
          
    joblog_txt_091 = oNode_joblog.selectSingleNode("txt_91").Text
    joblog_txt_092 = oNode_joblog.selectSingleNode("txt_92").Text
    joblog_txt_093 = oNode_joblog.selectSingleNode("txt_93").Text
    joblog_txt_094 = oNode_joblog.selectSingleNode("txt_94").Text
    joblog_txt_095 = oNode_joblog.selectSingleNode("txt_95").Text
    joblog_txt_096 = oNode_joblog.selectSingleNode("txt_96").Text
    joblog_txt_097 = oNode_joblog.selectSingleNode("txt_97").Text
    joblog_txt_098 = oNode_joblog.selectSingleNode("txt_98").Text
    joblog_txt_099 = oNode_joblog.selectSingleNode("txt_99").Text
          
    joblog_txt_100 = oNode_joblog.selectSingleNode("txt_100").Text
    joblog_txt_101 = oNode_joblog.selectSingleNode("txt_101").Text
    joblog_txt_102 = oNode_joblog.selectSingleNode("txt_102").Text
    joblog_txt_103 = oNode_joblog.selectSingleNode("txt_103").Text
          
    joblog_txt_104 = oNode_joblog.selectSingleNode("txt_104").Text
    joblog_txt_105 = oNode_joblog.selectSingleNode("txt_105").Text
    joblog_txt_106 = oNode_joblog.selectSingleNode("txt_106").Text
         
          
    joblog_txt_116 = oNode_joblog.selectSingleNode("txt_116").Text
    joblog_txt_117 = oNode_joblog.selectSingleNode("txt_117").Text
    joblog_txt_118 = oNode_joblog.selectSingleNode("txt_118").Text
    joblog_txt_119 = oNode_joblog.selectSingleNode("txt_119").Text
         
    joblog_txt_120 = oNode_joblog.selectSingleNode("txt_120").Text
    joblog_txt_121 = oNode_joblog.selectSingleNode("txt_121").Text
    joblog_txt_122 = oNode_joblog.selectSingleNode("txt_122").Text
    joblog_txt_123 = oNode_joblog.selectSingleNode("txt_123").Text
          
    joblog_txt_124 = oNode_joblog.selectSingleNode("txt_124").Text
    joblog_txt_125 = oNode_joblog.selectSingleNode("txt_125").Text
    joblog_txt_126 = oNode_joblog.selectSingleNode("txt_126").Text
    joblog_txt_127 = oNode_joblog.selectSingleNode("txt_127").Text
    joblog_txt_128 = oNode_joblog.selectSingleNode("txt_128").Text
            
    joblog_txt_129 = oNode_joblog.selectSingleNode("txt_129").Text
    joblog_txt_130 = oNode_joblog.selectSingleNode("txt_130").Text
    joblog_txt_131 = oNode_joblog.selectSingleNode("txt_131").Text
    joblog_txt_132 = oNode_joblog.selectSingleNode("txt_132").Text
    joblog_txt_133 = oNode_joblog.selectSingleNode("txt_133").Text
            
    joblog_txt_134 = oNode_joblog.selectSingleNode("txt_134").Text
    joblog_txt_135 = oNode_joblog.selectSingleNode("txt_135").Text

    joblog_txt_136 = oNode_joblog.selectSingleNode("txt_136").Text
    joblog_txt_137 = oNode_joblog.selectSingleNode("txt_137").Text
    joblog_txt_138 = oNode_joblog.selectSingleNode("txt_138").Text
    joblog_txt_139 = oNode_joblog.selectSingleNode("txt_139").Text
            
    joblog_txt_140 = oNode_joblog.selectSingleNode("txt_140").Text
    joblog_txt_141 = oNode_joblog.selectSingleNode("txt_141").Text
            
    joblog_txt_142 = oNode_joblog.selectSingleNode("txt_142").Text
    joblog_txt_143 = oNode_joblog.selectSingleNode("txt_143").Text
    joblog_txt_144 = oNode_joblog.selectSingleNode("txt_144").Text
    joblog_txt_145 = oNode_joblog.selectSingleNode("txt_145").Text
    joblog_txt_146 = oNode_joblog.selectSingleNode("txt_146").Text
            
    joblog_txt_147 = oNode_joblog.selectSingleNode("txt_147").Text
    joblog_txt_148 = oNode_joblog.selectSingleNode("txt_148").Text
    joblog_txt_149 = oNode_joblog.selectSingleNode("txt_149").Text
            
    joblog_txt_150 = oNode_joblog.selectSingleNode("txt_150").Text
    joblog_txt_151 = oNode_joblog.selectSingleNode("txt_151").Text
    joblog_txt_152 = oNode_joblog.selectSingleNode("txt_152").Text
    joblog_txt_153 = oNode_joblog.selectSingleNode("txt_153").Text
    joblog_txt_154 = oNode_joblog.selectSingleNode("txt_154").Text
  
          
    next




'Response.Write "joblog_txt_001: " & joblog_txt_001 & "<br>"
'Response.Write "joblog_txt_002: " & joblog_txt_002 & "<br>"


%>