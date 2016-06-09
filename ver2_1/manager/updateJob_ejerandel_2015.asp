


<%
strConnect = "driver={MySQL ODBC 3.51 Driver};server=195.189.130.210; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_epi;"

	
Response.write strConnect & "<br><br>"

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject ("ADODB.Recordset")
Set oRec2 = Server.CreateObject ("ADODB.Recordset")
Set oRec3 = Server.CreateObject ("ADODB.Recordset")
Set oCmd = Server.CreateObject ("ADODB.Command")

oConn.open strConnect

%>
<!--#include file="../inc/regular/global_func.asp"-->
<%

strSQL = "SELECT jobnr, jobans1txt, jobans2txt, jobans3txt, jobans4txt, jobans5txt, jobans1_proc, jobans2_proc, jobans3_proc, jobans4_proc, jobans5_proc, epi_andel, "_
&" salgsans1txt, salgsans2txt, salgsans3txt, salgsans4txt, salgsans5txt, salgsans1_proc, salgsans2_proc, salgsans3_proc, salgsans4_proc, salgsans5_proc "_
&" FROM jobans_salg_proc_imp_2015 AS j WHERE id <> 0 AND jobnr > 0 ORDER BY jobnr" ' LIMIT 50" 

    'AND jobnr > 1000 AND id > 0
Response.Write strSQL & "<br><br>"
Response.flush

'Response.end
oRec.open strSQL, oConn, 3
while not oRec.EOF 














                    '***************** JOBANSVARLIGE ***********************************************************


                    
                    jobans1 = 0
					if len(trim(oRec("jobans1txt"))) <> 0 then		
                     strSQLmeid = "SELECT mid FROM medarbejdere WHERE init = '" & oRec("jobans1txt") & "'"       
                     oRec2.open strSQLmeid, oConn, 3
                     if not oRec2.EOF then      	
					 
                     jobans1 = oRec2("mid")
                     
                     end if
                     oRec2.close	
                     end if
                     

                     jobans2 = 0
                     if len(trim(oRec("jobans2txt"))) <> 0 then
                      strSQLmeid = "SELECT mid FROM medarbejdere WHERE init = '" & oRec("jobans2txt") & "'"      
                     oRec2.open strSQLmeid, oConn, 3
                     if not oRec2.EOF then      	
					 
                     jobans2 = oRec2("mid")
                     
                     end if
                     oRec2.close	
                     end if
                     
                     
                     jobans3 = 0
                     if len(trim(oRec("jobans3txt"))) <> 0 then
                      strSQLmeid = "SELECT mid FROM medarbejdere WHERE init = '" & oRec("jobans3txt")   & "'"    
                     oRec2.open strSQLmeid, oConn, 3
                     if not oRec2.EOF    then   	
					 
                     jobans3 = oRec2("mid")
                     
                     end if
                     oRec2.close	
                     end if
                     


                     jobans4 = 0

                     if len(trim(oRec("jobans4txt"))) <> 0 then
                      strSQLmeid = "SELECT mid FROM medarbejdere WHERE init = '" & oRec("jobans4txt")  & "'"     
                     oRec2.open strSQLmeid, oConn, 3
                     if not oRec2.EOF then      	
					 
                     jobans4 = oRec2("mid")
                     
                     end if
                     oRec2.close	

                     end if
                     

                      jobans5 = 0

                     if len(trim(oRec("jobans5txt"))) <> 0 then
                      strSQLmeid = "SELECT mid FROM medarbejdere WHERE init = '" & oRec("jobans5txt")  & "'"     
                     oRec2.open strSQLmeid, oConn, 3
                     if not oRec2.EOF then      	
					 
                     jobans5 = oRec2("mid")
                     
                     end if
                     oRec2.close	

                     end if
                     		

                    'Response.write  oRec("jobnr") &";"& oRec("jobans1") &" ("& oRec("jobans1andel") &");"& oRec("jobans2") &" ("& oRec("jobans2andel") &");"& oRec("jobans3") &" ("& oRec("jobans3andel") &");"& oRec("jobans4") &"  ("& oRec("jobans4andel") &");"& oRec("jobans5") & " ("& oRec("jobans_proc_5") &") <br>"
                    
                    if isNULL(oRec("jobans1_proc")) <> true then
                    jansprocent_1 = replace(oRec("jobans1_proc"), ",", ".") 
                    else
                    jansprocent_1 = 0
                    end if

                    if isNULL(oRec("jobans2_proc")) <> true then
                    jansprocent_2 = replace(oRec("jobans2_proc"), ",", ".") 
                    else
                    jansprocent_2 = 0
                    end if

                    if isNULL(oRec("jobans3_proc")) <> true then
                    jansprocent_3 = replace(oRec("jobans3_proc"), ",", ".") 
                    else
                    jansprocent_3 = 0
                    end if

                    if isNULL(oRec("jobans4_proc")) <> true then
                    jansprocent_4 = replace(oRec("jobans4_proc"), ",", ".") 
                    else
                    jansprocent_4 = 0
                    end if

                    if isNULL(oRec("jobans5_proc")) <> true then
                    jansprocent_5 = replace(oRec("jobans5_proc"), ",", ".") 
                    else
                    jansprocent_5 = 0
                    end if
                    
                    if isNULL(oRec("epi_andel")) <> true then
                    virkandel = replace(oRec("epi_andel"), ",", ".") 
                    else
                    virkandel = 0
                    end if



                    '**************** SALG ***********************************************




                 salgsans1 = 0
					if len(trim(oRec("salgsans1txt"))) <> 0 then		
                     strSQLmeid = "SELECT mid FROM medarbejdere WHERE init = '" & oRec("salgsans1txt") & "'"       
                     oRec2.open strSQLmeid, oConn, 3
                     if not oRec2.EOF then      	
					 
                     salgsans1 = oRec2("mid")
                     
                     end if
                     oRec2.close	
                     end if
                     

                     salgsans2 = 0
                     if len(trim(oRec("salgsans2txt"))) <> 0 then
                      strSQLmeid = "SELECT mid FROM medarbejdere WHERE init = '" & oRec("salgsans2txt") & "'"      
                     oRec2.open strSQLmeid, oConn, 3
                     if not oRec2.EOF then      	
					 
                     salgsans2 = oRec2("mid")
                     
                     end if
                     oRec2.close	
                     end if
                     
                     
                     salgsans3 = 0
                     if len(trim(oRec("salgsans3txt"))) <> 0 then
                      strSQLmeid = "SELECT mid FROM medarbejdere WHERE init = '" & oRec("salgsans3txt")   & "'"    
                     oRec2.open strSQLmeid, oConn, 3
                     if not oRec2.EOF    then   	
					 
                     salgsans3 = oRec2("mid")
                     
                     end if
                     oRec2.close	
                     end if
                     


                     salgsans4 = 0

                     if len(trim(oRec("salgsans4txt"))) <> 0 then
                      strSQLmeid = "SELECT mid FROM medarbejdere WHERE init = '" & oRec("salgsans4txt")  & "'"     
                     oRec2.open strSQLmeid, oConn, 3
                     if not oRec2.EOF then      	
					 
                     salgsans4 = oRec2("mid")
                     
                     end if
                     oRec2.close	

                     end if
                     

                      salgsans5 = 0

                     if len(trim(oRec("salgsans5txt"))) <> 0 then
                      strSQLmeid = "SELECT mid FROM medarbejdere WHERE init = '" & oRec("salgsans5txt")  & "'"     
                     oRec2.open strSQLmeid, oConn, 3
                     if not oRec2.EOF then      	
					 
                     salgsans5 = oRec2("mid")
                     
                     end if
                     oRec2.close	

                     end if
                     		

                    'Response.write  oRec("salgsnr") &";"& oRec("salgsans1") &" ("& oRec("salgsans1andel") &");"& oRec("salgsans2") &" ("& oRec("salgsans2andel") &");"& oRec("salgsans3") &" ("& oRec("salgsans3andel") &");"& oRec("salgsans4") &"  ("& oRec("salgsans4andel") &");"& oRec("salgsans5") & " ("& oRec("salgsans_proc_5") &") <br>"
                    
                    if isNULL(oRec("salgsans1_proc")) <> true then
                    salgsprocent_1 = replace(oRec("salgsans1_proc"), ",", ".") 
                    else
                    salgsprocent_1 = 0
                    end if

                    if isNULL(oRec("salgsans2_proc")) <> true then
                    salgsprocent_2 = replace(oRec("salgsans2_proc"), ",", ".") 
                    else
                    salgsprocent_2 = 0
                    end if

                    if isNULL(oRec("salgsans3_proc")) <> true then
                    salgsprocent_3 = replace(oRec("salgsans3_proc"), ",", ".") 
                    else
                    salgsprocent_3 = 0
                    end if

                    if isNULL(oRec("salgsans4_proc")) <> true then
                    salgsprocent_4 = replace(oRec("salgsans4_proc"), ",", ".") 
                    else
                    salgsprocent_4 = 0
                    end if

                    if isNULL(oRec("salgsans5_proc")) <> true then
                    salgsprocent_5 = replace(oRec("salgsans5_proc"), ",", ".") 
                    else
                    salgsprocent_5 = 0
                    end if



                















                    '************************************* OPDATETERER ******************************************
                   'strSQLjobupdatdt = "SELECT dato FROM job WHERE jobnr = '" & oRec("jobnr") &"'"
                   'oRec3.open strSQLjobupdatdt, oConn, 3
                   'if not oRec3.EOF then
                                

                                '**** Opdater alle jobansvarlige *************
                    '            if cdate(oRec3("dato")) < "1-10-2013" then

								strSQLupd = "UPDATE job SET jobans_proc_1 = "& jansprocent_1 &", jobans_proc_2 = "& jansprocent_2 &", "_
                                &" jobans_proc_3 = "& jansprocent_3 &","_
                                &" jobans_proc_4 = "& jansprocent_4 &", jobans_proc_5 = "& jansprocent_5 &", virksomheds_proc = "& virkandel &", " _
                                &" jobans1 = "& jobans1 &", "_
                                &" jobans2 = "& jobans2 &", "_
                                &" jobans3 = "& jobans3 &", "_
                                &" jobans4 = "& jobans4 &", "_
                                &" jobans5 = "& jobans5 &", "_
                                &" salgsans1_proc = "& salgsprocent_1 &", salgsans2_proc = "& salgsprocent_2 &", "_
                                &" salgsans3_proc = "& salgsprocent_3 &","_
                                &" salgsans4_proc = "& salgsprocent_4 &", salgsans5_proc = "& salgsprocent_5 &","_
                                &" salgsans1 = "& salgsans1 &", "_
                                &" salgsans2 = "& salgsans2 &", "_
                                &" salgsans3 = "& salgsans3 &", "_
                                &" salgsans4 = "& salgsans4 &", "_
                                &" salgsans5 = "& salgsans5 &" "_
                                &" WHERE jobnr = '" & oRec("jobnr") &"'"
                                
                                Response.Write strSQLupd & "<br><br>"
                                Response.flush
                                'oConn.execute(strSQLupd)

            
                                '****** Opdater kun EPI andele
                                'strSQLupd = "UPDATE job SET virksomheds_proc = "& virkandel &"" _
                                '&" WHERE jobnr = '" & oRec("jobnr") &"'"
                                
                                'Response.Write strSQLupd & "<br><br>"
                                'Response.flush
                                'oConn.execute(strSQLupd)

                     '           end if
                     
                     'end if
                     'oRec3.close           'end if

oRec.movenext
wend
oRec.close



%>
<br><br><br>
Opdatering gennemført!
