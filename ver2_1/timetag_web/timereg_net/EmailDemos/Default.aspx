<%@ Page Language="VB" MasterPageFile="MasterPage.master" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="_Default" title="Untitled Page" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <h4>Simple Feedback Form</h4>
    
    <asp:Panel runat="server" ID="EmailForm">
    <p>The following example illustrates a simple feedback form. For completeness, take a moment to add RegularExpressionValidators
    to ensure that the email values are valid and RequiredFieldValidators to ensure that the required inputs are provided...
    <strong>Also be sure to update the <code>To</code> property in the code-behind class!!</strong></p>
    <table border="0">
        <tr>
            <td><b>Your Email:</b></td>
            <td><asp:TextBox runat="server" ID="UsersEmail" Columns="30"></asp:TextBox></td>
        </tr>
        <tr>
            <td><b>Subject:</b></td>
            <td><asp:TextBox runat="server" ID="Subject" Columns="30"></asp:TextBox></td>
        </tr>
        <tr>
            <td colspan="2">
                <b>Body:</b><br />
                <asp:TextBox runat="server" ID="Body" TextMode="MultiLine" Columns="55" Rows="10"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td colspan="2" align="center">
                <asp:Button runat="server" ID="SendEmail" Text="Send Feedback" />
            </td>
        </tr>
    </table>
    </asp:Panel>
    
    <asp:Panel runat="server" ID="EmailSentForm">
        Your feedback has been sent... thank you!
    </asp:Panel>
</asp:Content>

