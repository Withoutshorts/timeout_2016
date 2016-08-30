
<!--#include file="../inc/connection/aktivedb_inc.asp"-->



<%



Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject("ADODB.Recordset")

useSQL = 9
Response.Write "her:"

					x = 1
					numberoflicens = 124
					For x = 1 To numberoflicens 
						
						call aktivedb(x)
						
						
						if strConnect_aktiveDB <> "nogo" then
						Response.write strConnect_aktiveDB &"<br>"
						
								
								oConn.open strConnect_aktiveDB
								
								
								select case useSQL
								case 1
								
								strSQL = "SELECT COUNT(fid) AS antal FROM fakturaer WHERE jobfaktype = 1"
								oRec.Open strSQL, oConn, 3 
								
								if not oRec.EOF then
								
								
								Response.Write "Antal:" & oRec("antal") & "<hr>"
								
								end if
								oRec.close
								
								Response.flush
								
								case 2
								
								strSQL = "SELECT fid, faknr, fakadr, kid, kkundenavn, adresse, postnr, city, land FROM fakturaer "_
								&" LEFT JOIN kunder ON (kid = fakadr) WHERE fid <> 0 ORDER BY fid"
								
								oRec.Open strSQL, oConn, 3 
								while not oRec.EOF
								    
								   if len(trim(oRec("adresse"))) <> 0 then
								   strAdr = replace(oRec("adresse"), "'", "")
								   else
								   strAdr = ""
								   end if
								   
								   if len(trim(oRec("kkundenavn"))) <> 0 then
								   knavn = replace(oRec("kkundenavn"), "'", "")
								   else
								   knavn = ""
								   end if
								   
								  strSQLupd = "UPDATE fakturaer SET modtageradr = '"& knavn & "<br>"_
								  & strAdr & "<br>"_
								  & oRec("postnr") &" "& oRec("city") &"<br>"& oRec("land") &"' WHERE fid = " & oRec("fid")
								  
								  
								  Response.Write "faknr:"& oRec("faknr") & ", "& oRec("fakadr") &"," &  oRec("kid") & "<br>"
								  'Response.Write strSQLupd & "<br><br><br>"
								  Response.flush
								  'oConn.Execute(strSQlupd)
								
								
								oRec.MoveNext
								wend
								oRec.close
								
								case 3
								
								strSQL = "SELECT COUNT(*) AS antal, jobnavn, jobnr, jobstatus FROM job_ulev_ju "_
								&" LEFT JOIN job AS j ON (j.id = ju_jobid) WHERE ju_jobid <> 0 GROUP BY ju_jobid"
								oRec.Open strSQL, oConn, 3 
								
								if not oRec.EOF then
								
								
								Response.Write "Job: "& oRec("jobnavn") &" ("& oRec("jobnr") &"), status: "& oRec("jobstatus")  &" , Antal ulev:" & oRec("antal") & "<br>"
								
								end if
								oRec.close

                                case 4

                                strSQL = "SELECT licens, timeout_version FROM licens WHERE id = 1"
                                oRec.Open strSQL, oConn, 3 
								
								if not oRec.EOF then
								
								if oRec("timeout_version") = "ver2_10" then
								Response.Write "<br>licens: "& oRec("licens") &" TO version: "& oRec("timeout_version") & "<hr>"
								end if

								end if
								oRec.close


                                '***

                                case 5

                                 '*** Forretningsområder ****'
                                lastjob = 0
                                strSQLa = "SELECT id, job, fomr FROM aktiviteter WHERE fomr <> 0 ORDER BY job"
                                oRec.open strSQLa, oConn, 3
                                while not oRec.EOF 

                                    if lastJob <> oRec("job") then

                                    strSQLfomr = "INSERT INTO fomr_rel "_
                                    &" (for_fomr, for_jobid, for_aktid, for_faktor) "_
                                    &" VALUES ("& oRec("fomr") &", "& oRec("job") &", 0, 100)"

                                    Response.Write strSQLfomr & "<br>"
                                    'Response.flush
                                    'oConn.execute(strSQLfomr)

                                    end if
                                    

                                    strSQLfomrai = "INSERT INTO fomr_rel "_
                                    &" (for_fomr, for_jobid, for_aktid, for_faktor) "_
                                    &" VALUES ("& oRec("fomr") &", "& oRec("job") &","& oRec("id") &", 100)"

                                    Response.Write strSQLfomrai & "<br>"
                                    'Response.flush
                                    'oConn.execute(strSQLfomrai)

                                lastJob = oRec("job")
                                oRec.movenext
                                wend
                                oRec.close

								
								
								
								
                                case 6

                                '****** Materialer på aftaler ****
                                lastjob = 0
                                strSQLa = "SELECT id, serviceaft FROM job WHERE serviceaft <> 0 ORDER BY id"
                                oRec.open strSQLa, oConn, 3
                                while not oRec.EOF 

                                   
                                    strSQLfomrai = "UPDATE materiale_forbrug SET "_
                                    &" serviceaft = "& oRec("serviceaft")  &" "_
                                    &" WHERE jobid = "& oRec("id") 

                                    Response.Write strSQLfomrai & "<br>"
                                    'Response.flush
                                    'oConn.execute(strSQLfomrai)

                                lastJob = oRec("id")
                                oRec.movenext
                                wend
                                oRec.close

                                case 7

                                strSQLa = "SELECT COUNT(tid) AS antal FROM timer WHERE godkendtstatus = 1 AND tdato > '2011-06-01' GROUP BY godkendtstatus"
                                
                                Response.write strSQLa
                                Response.flush
                                oRec.open strSQLa, oConn, 3
                                if not oRec.EOF then

                                Response.Write oRec("antal") & "<br>"

                                end if
                                oRec.close


                                case 8

                                strSQLa = "SELECT jobnavn, jobnr FROM job WHERE jobstatus = 3 AND jobstartdato > '2011-06-01'"
                                
                                Response.write strSQLa
                                Response.flush
                                oRec.open strSQLa, oConn, 3
                                if not oRec.EOF then

                                Response.Write oRec("jobnavn") & " ("& oRec("jobnr") &")" & "<br>"

                                end if
                                oRec.close


                                case 9

                                strSQLa = "SELECT kid FROM kunder WHERE useasfak = 1"
                                
                                Response.write strSQLa
                                Response.flush
                                oRec.open strSQLa, oConn, 3
                                if not oRec.EOF then

                                strSQlopdFak = "UPDATE fakturaer SET afsender = "& oRec("kid")
                                'oConn.execute(strSQlopdFak)

                                end if
                                oRec.close

								end select

								
						
						oConn.close
						end if
							
                    next

            
            
            Response.Write "<br><br>Alles klar<br><br>&nbsp;"


%>