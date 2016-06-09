<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="SendAttachment.aspx.vb" Inherits="SendAttachment" title="Untitled Page" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <h4>Upload and Email an Attachment</h4>

    <asp:Panel runat="server" ID="EmailForm">  
        <p>
            This demo shows how to use the FileUpload control to allow a visitor to upload a file from their computer to the
            web server, which then emails that file as an attachment in an email.
        </p>

        <table border="0">
            <tr>
                <td><b>Your Email:</b></td>
                <td><asp:TextBox runat="server" ID="UsersEmail" Columns="30"></asp:TextBox></td>
            </tr>
            <tr>
                <td><b>File to Send:</b></td>
                <td>
                    <asp:FileUpload ID="AttachmentFile" runat="server" />
                </td>
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
        Your attachment has been emailed... thanks!
    </asp:Panel>
</asp:Content>

