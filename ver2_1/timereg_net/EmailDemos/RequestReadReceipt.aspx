<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="RequestReadReceipt.aspx.vb" Inherits="RequestReadReceipt" title="Untitled Page" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <h2>
        Request Read Receipt</h2>
    <p>
        This demo illustrates how to add the Disposition-Notification-To SMTP header to an outgoing
        email message. This SMTP header is responsible for requesting a read receipt. Keep in mind
        that not all email servers and clients support read receipt requests; moreover, even if
        such support is present, a system administrator or individual user may disable the functionality.
        In short, you are <i>requesting</i> a read receipt. You are <b>not</b> guaranteed to receive one!
    </p>

    <asp:Panel runat="server" ID="EmailForm">
    <strong>Be sure to update the <code>To</code> property in the code-behind class!!</strong>
        <br /><br />
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
        Your email has been sent with a request for a read receipt... thank you!
    </asp:Panel>    
</asp:Content>

