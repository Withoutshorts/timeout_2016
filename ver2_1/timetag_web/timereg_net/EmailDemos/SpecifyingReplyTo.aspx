<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="SpecifyingReplyTo.aspx.vb" Inherits="SpecifyingReplyTo" title="Untitled Page" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <h2>
        Specifying a Reply-To Address</h2>
    <p>
        This demo illustrates how to send an email where the From
        address is from one email account, but replying to the email
        sends it to another account. It also looks at
        how to indicate a priority.</p>

    <asp:Panel runat="server" ID="EmailForm">
    <strong>Be sure to update the <code>To</code> property in the code-behind class!!</strong>
        <br /><br />
    <table border="0">
        <tr>
            <td><b>The From Email Address:</b></td>
            <td><asp:TextBox runat="server" ID="UsersEmail" Columns="30"></asp:TextBox></td>
        </tr>
        <tr>
            <td><b>The Reply-To Email Address:</b></td>
            <td><asp:TextBox runat="server" ID="ReplyToEmail" Columns="30"></asp:TextBox></td>
        </tr>
        <tr>
            <td><b>Priority:</b></td>
            <td>
                <asp:DropDownList ID="EmailPriority" runat="server">
                    <asp:ListItem>Low</asp:ListItem>
                    <asp:ListItem Selected="True">Normal</asp:ListItem>
                    <asp:ListItem>High</asp:ListItem>
                </asp:DropDownList></td>
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
        Your email has been sent... thank you!
    </asp:Panel>    
</asp:Content>

