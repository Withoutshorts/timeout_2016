<%@ Page Language="C#" AutoEventWireup="true" CodeFile="importer_job.cs" Inherits="importer_job" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">


<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>TimeOut import job</title>
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
    <h4>Importer job til TimeOut - maks 8000 linjer <span style="font-size:small; font-weight:normal;"><a href="../inc/xls/indlaesjobTemplate.csv">Download excel template her...</a></span></h4>
    *Bemærk at der skal være en overskriftslinje(header) med kolonne navne i excel filen.<br />
        <asp:Label ID="lblUploadStatus" runat="server" Text=""></asp:Label><br />
        
        <div>
         <asp:FileUpload ID="fu" runat="server" />
          <asp:Button ID="btnUpload" runat="server" Text="Upload" 
                onclick="btnUpload_Click" />
        </div><br />

        <div>

        <table>

        <%if (importtype == "d1" || importtype == "t1")
            {  %>
        <tr>
        <td>
        Kundenavn:
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
        Kundenr:
        </td>
        <td>
            <asp:DropDownList ID="ddlKnr" runat="server" AutoPostBack="True" 
                onselectedindexchanged="ddlKnr_SelectedIndexChanged" 
                ondatabound="ddlKnr_DataBound" AppendDataBoundItems="True">
                <asp:ListItem Text="" Value=""></asp:ListItem>
            </asp:DropDownList>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator9" runat="server" 
                ErrorMessage="*" ControlToValidate="ddlKnr" ValidationGroup="Send" ForeColor="Red"></asp:RequiredFieldValidator>
        </td>      
        <td>
            <asp:Label ID="lblKnr" runat="server" Text=""></asp:Label>
        </td>
        </tr>

        <%}; %>           

      
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
        Jobnr.:
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
            
          <%if (importtype == "t1")
              {  %>
      
             <tr>
        <td>
        <asp:label runat="server" ID="Label1">Costcenter:</asp:label>
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
            <%}; %>

        </table>
        </div>
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
