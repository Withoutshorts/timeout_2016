<tr>
							<td>
								<div id="<%=oRec2("id")%>" Style="position: relative; display: none;">
								<table cellspacing="0" cellpadding="0" border="0">
								<%
								'********************************************
								'* Henter medarbejdere pr. aktivitet
								'********************************************
								%>
								<%m = 0
								For m = 0 to Ubound(intMedArbValThis)%>
								<tr>
									<td><%=intMedArbValThis(m)%></td>
								</tr>
								<tr>
									<td valign="top">
									<%
									call periode(periodeinterval, periodeintervalWeeks, StrTdato, StrUdato)
									call findtimer("medprakt", intMedArbValThis(m), oRec2("id"))
									%>
									</table>
									</td>
								</tr>
								<%next%>
								</table>
								</div>
							</td>
						 </tr>
