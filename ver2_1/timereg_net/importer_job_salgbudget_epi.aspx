<%@ Page Language="C#" AutoEventWireup="true" CodeFile="importer_job_salgbudget_epi.cs" Inherits="importer_job" Debug="true" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">


<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>TimeOut import Job-Salgsbudget Epinion</title>
</head>
<body style="font-family:Arial" leftmargin="20" topmargin="20">

    


  <table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	
	</td></tr></table>

    <form id="form1" runat="server">

          <label ID="lbl_importtype" runat="server">..</label>
          <asp:HiddenField id="hdn_importtype" runat="server" Value="" />
             <!--<asp:TextBox id="txt_importtype" runat="server" Value="" />-->
        

    <div>

       

    <h4>Add/Import salesbudget to exsisting job in TimeOut </h4>
  
        <asp:Label ID="lblUploadStatus" runat="server" Text=""></asp:Label><br />
        
        <div>
         Browse for salesbudget: (.CSV)
         <asp:FileUpload ID="fu" runat="server" />
          <asp:Button ID="btnUpload" runat="server" Text="Upload" 
                onclick="btnUpload_Click" />
        </div><br />

        <div>

        <table>

        <tr><td><h4>Basedata</h4></td></tr>
        <tr>
        <td>
        Jobnr.: (must exsist in TimeOut)
        </td>
        <td>
            <asp:DropDownList ID="ddlJobId" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlJobId_SelectedIndexChanged" 
                ondatabound="ddlJobId_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlJobId" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
         <td>
            <asp:Label ID="lblJobId" runat="server" Text=""></asp:Label>
        </td>
        </tr>
       
        <tr>
        <td>
        Projecttype:
        </td>
        <td>
            <asp:DropDownList ID="ddlKnavn" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlKnavn_SelectedIndexChanged" 
                ondatabound="ddlKnavn_DataBound" AppendDataBoundItems="True">
                <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator8" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlKnavn" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>      
        <td>
            <asp:Label ID="lblKnavn" runat="server" Text=""></asp:Label>
        </td>
        </tr>

          <tr>
        <td>
        Option:
        </td>
        <td>
            <asp:DropDownList runat ="server">
                <asp:ListItem Text="Budget" Value=""></asp:ListItem>
                  <asp:ListItem Text="Option 1" Value=""></asp:ListItem>
                  <asp:ListItem Text="Option 2" Value=""></asp:ListItem>
                  <asp:ListItem Text="Option 3" Value=""></asp:ListItem>
                 
            </asp:DropDownList>
           
        </td>      
        <td>
      
        </td>
        </tr>

        </table>



             <%
            //**********************************************************************************************************************
            // Project - Budget - Medarb. typer
            //********************************************************************************************************************** 
            %>

            <br /><br />

            <h4>Project - Budget</h4>
            <table>
            <tr>


                <%

                    for (int i = 0; i < 1; i++)
                    {
                        %>
                      
                <td>

                        <table>
                         <tr>
                        <td>
                        <b>Man.Director & Director:</b>
                        </td>
                        <td>
                            <asp:DropDownList ID="ddlMtype" runat="server" AutoPostBack="True" 
                                onselectedindexchanged="ddlMtype_SelectedIndexChanged" 
                                ondatabound="ddlMtype_DataBound" AppendDataBoundItems="True">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                            <asp:RequiredFieldValidator ID="RequiredFieldValidator_Mtype" runat="server" 
                                ErrorMessage="*" ControlToValidate="ddlMtype" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
                        </td>      
                        <td>
                            <asp:Label ID="lblMtype" runat="server" Text=""></asp:Label>
                        </td>
                        </tr>


                        <tr>
                        <td>
                        Total number of hours 

                        </td>
                        <td>
                            <asp:DropDownList ID="ddlMtypeHours" runat="server" AutoPostBack="True" 
                                onselectedindexchanged="ddlMtypeHours_SelectedIndexChanged" 
                                ondatabound="ddlMtypeHours_DataBound" AppendDataBoundItems="True">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                            <asp:RequiredFieldValidator ID="RequiredFieldValidator_MtypeHours" runat="server" 
                                ErrorMessage="*" ControlToValidate="ddlMtypeHours" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
                        </td>      
                        <td>
                            <asp:Label ID="lblMtypeHours" runat="server" Text=""></asp:Label>
                        </td>
                        </tr>


                        <tr>
                        <td>
                         Sales Price per hour 

                        </td>
                        <td>
                            <asp:DropDownList ID="ddlMtypeSalesPrice" runat="server" AutoPostBack="True" 
                                onselectedindexchanged="ddlMtypeSalesPrice_SelectedIndexChanged" 
                                ondatabound="ddlMtypeSalesPrice_DataBound" AppendDataBoundItems="True">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                            <asp:RequiredFieldValidator ID="RequiredFieldValidator9" runat="server" 
                                ErrorMessage="*" ControlToValidate="ddlMtypeSalesPrice" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
                        </td>      
                        <td>
                            <asp:Label ID="lblMtypeSalesPrice" runat="server" Text=""></asp:Label>
                        </td>
                        </tr>

                        <tr>
                        <td>
                          Revenue </td>
                        <td>
                            <asp:DropDownList ID="ddlMtypeRevenue" runat="server" AutoPostBack="True" 
                                onselectedindexchanged="ddlMtypeRevenue_SelectedIndexChanged" 
                                ondatabound="ddlMtypeRevenue_DataBound" AppendDataBoundItems="True">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                            <asp:RequiredFieldValidator ID="RequiredFieldValidator11" runat="server" 
                                ErrorMessage="*" ControlToValidate="ddlMtypeRevenue" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
                        </td>      
                        <td>
                            <asp:Label ID="lblMtypeRevenue" runat="server" Text=""></asp:Label>
                        </td>
                        </tr>

                         <tr>
                        <td>
                          Cost</td>
                        <td>
                            <asp:DropDownList ID="ddlMtypeCost" runat="server" AutoPostBack="True" 
                                onselectedindexchanged="ddlMtypeCost_SelectedIndexChanged" 
                                ondatabound="ddlMtypeCost_DataBound" AppendDataBoundItems="True">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                            <asp:RequiredFieldValidator ID="RequiredFieldValidator14" runat="server" 
                                ErrorMessage="*" ControlToValidate="ddlMtypeCost" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
                        </td>      
                        <td>
                            <asp:Label ID="lblMtypeCost" runat="server" Text=""></asp:Label>
                        </td>
                        </tr>
                    </table>



        </td>


                  <%

                      };

                      
                      for (int y = 0; y < 6; y++)
                      {
                          string mtypeLabeltxt = "";
                          
                         if (y == 0) { 
                           mtypeLabeltxt = "Sen. Manager & Manager";
                          };

                           if (y == 1) { 
                          mtypeLabeltxt = "Senior Consultant";
                          };

                           if (y == 2) { 
                          mtypeLabeltxt = "Consultant";
                          };

                           if (y == 3) { 
                          mtypeLabeltxt = "Junior Consultant";
                          };

                           if (y == 4) { 
                          mtypeLabeltxt = "Professional services";
                          };

                           if (y == 5) { 
                          mtypeLabeltxt = "Trainee";
                          };
                        
                        %>
                      
                <td>

                        <table>
                         <tr>
                        <td>
                        <b><%=mtypeLabeltxt%>:</b>
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        <td>
                           
                        </td>
                        </tr>


                        <tr>
                        <td>
                        Total number of hours 

                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                            
                        </td>      
                        <td>
                       
                        </td>
                        </tr>


                        <tr>
                        <td>
                         Sales Price per hour 

                        </td>
                      <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                            
                        </td>      
                        <td>
                         
                        </td>
                        </tr>

                        <tr>
                        <td>
                          Revenue </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                          
                        </td>      
                        <td>
                           
                        </td>
                        </tr>

                         <tr>
                        <td>
                          Cost</td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                            
                        </td>      
                        <td>
                          
                        </td>
                        </tr>
                    </table>



        </td>
                <%}; %>
             



    </tr>
            </table>

            
            <%
            //**********************************************************************************************************************
            // Epinion Data Collection
            //********************************************************************************************************************** 
            %>
            <br /><br /><br />

            <h4>Epinion Data Collection</h4>

            <table>
                <tr>
                    
                 <td>  CATI

                        <table>
                         <tr>
                        <td>
                        Sales price pr. Unit:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        <td>
                           
                        </td>
                        </tr>
                   <tr>
                        <td>
                        Revenue:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        <td>
                           
                        </td>
                        </tr>
                   <tr>
                        <td>
                        Cost:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        <td>
                           
                        </td>
                        </tr>
               </table>



                </td>


                 <td> CAPI Trans

                        <table>
                         <tr>
                        <td>
                        Sales price pr. Unit:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        <td>
                           
                        </td>
                        </tr>
                   <tr>
                        <td>
                        Revenue:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        <td>
                           
                        </td>
                        </tr>
                   <tr>
                        <td>
                        Cost:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        <td>
                           
                        </td>
                        </tr>
               </table>



                </td>


                 <td>   CAPI Aviation <br />


                        <table>
                         <tr>
                        <td>
                        Sales price pr. Unit:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        <td>
                           
                        </td>
                        </tr>
                   <tr>
                        <td>
                        Revenue:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        <td>
                           
                        </td>
                        </tr>
                   <tr>
                        <td>
                        Cost:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        <td>
                           
                        </td>
                        </tr>
               </table>



                </td>




                </tr>
            </table>

            <%//********************************************************************************************************************** %>




            <%
            //**********************************************************************************************************************
            //  CAWI Data collection 
            //********************************************************************************************************************** 
            %>





             <br /><br /><br />

            <h4>CAWI Data collection </h4>

            <table>
                <tr>
                    
                 <td>   CAWI - Completed 


                        <table>
                         <tr>
                        <td>
                        Sales price pr. Unit:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        <td>
                           
                        </td>
                        </tr>
                   <tr>
                        <td>
                        Revenue:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        <td>
                           
                        </td>
                        </tr>
                   <tr>
                        <td>
                        Cost:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        <td>
                           
                        </td>
                        </tr>
               </table>



                </td>


                 <td>  CAWI - Screened 


                        <table>
                         <tr>
                        <td>
                        Sales price pr. Unit:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        <td>
                           
                        </td>
                        </tr>
                   <tr>
                        <td>
                        Revenue:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        <td>
                           
                        </td>
                        </tr>
                   <tr>
                        <td>
                        Cost:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        <td>
                           
                        </td>
                        </tr>
               </table>



                </td>

                </tr>
            </table>





             <%
            //**********************************************************************************************************************
            //  CAWI Data collection 
            //********************************************************************************************************************** 
            %>





             <br /><br /><br />

            <h4>Focusgroup (FG) </h4>

            <table>
                <tr>
                    
                 <td>    Number of FG 



                        <table>
                         <tr>
                        <td>
                        Sales price pr. Unit:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        <td>
                           
                        </td>
                        </tr>
                   <tr>
                        <td>
                        Revenue:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        <td>
                           
                        </td>
                        </tr>
                   <tr>
                        <td>
                        Cost:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        <td>
                           
                        </td>
                        </tr>
               </table>



                </td>


                 <td>   Recruiting 

                        <table>
                         <tr>
                        <td>
                        Sales price pr. Unit:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        <td>
                           
                        </td>
                        </tr>
                   <tr>
                        <td>
                        Revenue:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        <td>
                           
                        </td>
                        </tr>
                   <tr>
                        <td>
                        Cost:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        <td>
                           
                        </td>
                        </tr>
               </table>



                </td>

                </tr>
            </table>






                 <%
            //**********************************************************************************************************************
            //   Other Purchases 

            //********************************************************************************************************************** 
            %>





             <br /><br /><br />

            <h4> Other Purchases </h4>

           

                        <table>
                         <tr>
                            <td> F53: <asp:DropDownList  runat="server">
                                <asp:ListItem Text="Gift Card Denmark" Value=""></asp:ListItem>
                            </asp:DropDownList></td>
                        <td>
                        Amount
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                     
                        <td>
                        Cost pr. item:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                      
                        <td>
                        Sales price pr. item:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        
                        </tr>

                          <tr>
                             <td> F54: <asp:DropDownList  runat="server">
                                <asp:ListItem Text="Sub-vendor xx" Value=""></asp:ListItem>
                            </asp:DropDownList></td>
                        <td>
                        Amount
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                     
                        <td>
                        Cost pr. item:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                      
                        <td>
                        Sales price pr. item:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        
                        </tr>

                          <tr>
                              <td> F55: <asp:DropDownList  runat="server">
                                <asp:ListItem Text="INCITE license per year" Value=""></asp:ListItem>
                            </asp:DropDownList></td>
                        <td>
                        Amount
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                     
                        <td>
                        Cost pr. item:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                      
                        <td>
                        Sales price pr. item:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        
                        </tr>


                          <tr>
                            <td> F56: <asp:DropDownList  runat="server">
                                <asp:ListItem Text="IT needed?" Value=""></asp:ListItem>
                            </asp:DropDownList></td>
                        <td>
                        Amount
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                     
                        <td>
                        Cost pr. item:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                      
                        <td>
                        Sales price pr. item:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        
                        </tr>

                          <tr>
                             <td> F57: <asp:DropDownList  runat="server">
                                <asp:ListItem Text="Postal" Value=""></asp:ListItem>
                            </asp:DropDownList></td>
                        <td>
                        Amount
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                     
                        <td>
                        Cost pr. item:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                      
                        <td>
                        Sales price pr. item:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        
                        </tr>


                          <tr>
                              <td> F58: <asp:DropDownList  runat="server">
                                <asp:ListItem Text="Other" Value=""></asp:ListItem>
                            </asp:DropDownList></td>
                        <td>
                        Amount
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                     
                        <td>
                        Cost pr. item:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                      
                        <td>
                        Sales price pr. item:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        
                        </tr>

                       <tr>
                             <td> F59: <asp:DropDownList  runat="server">
                                <asp:ListItem Text="Other" Value=""></asp:ListItem>
                            </asp:DropDownList></td>
                        <td>
                        Amount
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                     
                        <td>
                        Cost pr. item:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                      
                        <td>
                        Sales price pr. item:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        
                        </tr>

                          <tr>
                             <td> F60: <asp:DropDownList  runat="server">
                                <asp:ListItem Text="Other" Value=""></asp:ListItem>
                            </asp:DropDownList></td>
                        <td>
                        Amount
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                     
                        <td>
                        Cost pr. item:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                      
                        <td>
                        Sales price pr. item:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        
                        </tr>

                          <tr>
                             <td> F61: <asp:DropDownList  runat="server">
                                <asp:ListItem Text="Other" Value=""></asp:ListItem>
                            </asp:DropDownList></td>
                        <td>
                        Amount
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                     
                        <td>
                        Cost pr. item:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                      
                        <td>
                        Sales price pr. item:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        
                        </tr>

                          <tr>
                             <td> F62: <asp:DropDownList  runat="server">
                                <asp:ListItem Text="Other" Value=""></asp:ListItem>
                            </asp:DropDownList></td>
                        <td>
                        Amount
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                     
                        <td>
                        Cost pr. item:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                      
                        <td>
                        Sales price pr. item:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        
                        </tr>

                          <tr>
                             <td> F63: <asp:DropDownList  runat="server">
                                <asp:ListItem Text="Other" Value=""></asp:ListItem>
                            </asp:DropDownList></td>
                        <td>
                        Amount
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                     
                        <td>
                        Cost pr. item:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                      
                        <td>
                        Sales price pr. item:
                        </td>
                        <td>
                            <asp:DropDownList  runat="server">
                                <asp:ListItem Text="" Value=""></asp:ListItem>
                            </asp:DropDownList>
                           
                        </td>      
                        
                        </tr>

               </table>



              






















            <%//**** Standard felter **** %>
            <br /><br /><br />
            
            <!--

      <table>

            <tr>
        <td>
        Kundenr:
        </td>
        <td>
            <asp:DropDownList ID="ddlKnr" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlKnr_SelectedIndexChanged" 
                ondatabound="ddlKnr_DataBound" AppendDataBoundItems="True">
                <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator15" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlKnr" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>      
        <td>
            <asp:Label ID="lblKnr" runat="server" Text=""></asp:Label>
        </td>
        </tr>

        <tr>
        <td>
        Jobnavn:
        </td>
        <td>
            <asp:DropDownList ID="ddlJobnavn" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlJobnavn_SelectedIndexChanged" 
                ondatabound="ddlJobnavn_DataBound" AppendDataBoundItems="True">
                <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlJobnavn" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>      
        <td>
            <asp:Label ID="lblJobnavn" runat="server" Text=""></asp:Label>
        </td>
        </tr>

       




        <tr>
        <td>
        Jobansvarlig:
        </td>
        <td>
            <asp:DropDownList ID="ddlJobans" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlJobans_SelectedIndexChanged" 
                ondatabound="ddlJobans_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlJobans" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
         <td>
            <asp:Label ID="lblJobans" runat="server" Text=""></asp:Label>
        </td>
        </tr>
        
        <tr>
        <td>
        Startdato:
        </td>
        <td>
            <asp:DropDownList ID="ddlstDato" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlstDato_SelectedIndexChanged" 
                ondatabound="ddlstDato_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator4" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlstDato" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td> 
         <td>
            <asp:Label ID="lblstDato" runat="server" Text=""></asp:Label>
        </td>
        </tr>

         <tr>
        <td>
        Slutdato:
        </td>
        <td>
            <asp:DropDownList ID="ddlslDato" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlslDato_SelectedIndexChanged" 
                ondatabound="ddlslDato_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator5" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlslDato" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td> 
         <td>
            <asp:Label ID="lblslDato" runat="server" Text=""></asp:Label>
        </td>
        </tr>

         
       
        <tr>
        <td>
        <asp:label runat="server" ID="feltnr5navn">Faktureringsnavn:</asp:label>
        </td>
        <td>
        <asp:DropDownList ID="ddlTimerKom" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlTimerKom_SelectedIndexChanged" 
                ondatabound="ddlTimerKom_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator6" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlTimerKom" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
        <td>
        <asp:Label ID="lblTimerKom" runat="server" Text=""></asp:Label>
        </td>
        </tr>
            
     
      
             <tr>
        <td>
        <asp:label runat="server" ID="feltnr6navn">Costcenter: (NOT IN USE)</asp:label>
        </td>
        <td>
        <asp:DropDownList ID="ddlProjgrp" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlProjgrp_SelectedIndexChanged" 
                ondatabound="ddlProjgrp_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator10" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlProjgrp" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
        <td>
        <asp:Label ID="lblProjgrp" runat="server" Text=""></asp:Label>
        </td>
        </tr>
          
      
      

            
             <tr>
        <td>
        <asp:label runat="server" ID="feltnr7navn">Kunde kontaktperson</asp:label>
        </td>
        <td>
        <asp:DropDownList ID="ddlKpers" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlKpers_SelectedIndexChanged" 
                ondatabound="ddlKpers_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator12" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlKpers" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
        <td>
        <asp:Label ID="lblKpers" runat="server" Text=""></asp:Label>
        </td>
        </tr>

            
             <tr>
        <td>
        <asp:label runat="server" ID="feltnr8navn">Rek. nr</asp:label>
        </td>
        <td>
        <asp:DropDownList ID="ddlRekvnr" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlRekvnr_SelectedIndexChanged" 
                ondatabound="ddlRekvnr_DataBound" AppendDataBoundItems="True">
                 <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator13" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlRekvnr" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>
        <td>
        <asp:Label ID="lblRekvnr" runat="server" Text=""></asp:Label>
        </td>
        </tr>


          

        </table>
        </div>

        -->

        <br />
        Filnavn
            <asp:TextBox ID="txtFileName" runat="server"></asp:TextBox>
        <asp:RequiredFieldValidator ID="RequiredFieldValidator7" runat="server" ErrorMessage="*" ForeColor="Red" ControlToValidate="txtFileName" ValidationGroup="Send"></asp:RequiredFieldValidator>
            <asp:Button ID="btnAdd" runat="server" Text="Send" onclick="btnAdd_Click" ValidationGroup="Send" /> til TimeOut    
            <br /><br />
        <asp:Label ID="lblStatus" runat="server" Text="..."></asp:Label>         
    </div>
    </form>
</body>
</html>
