<%@ Page Language="vb" Debug="true" %>

<%@ Import Namespace="System.Web.Services.Description" %>

<%@  Import Namespace="Microsoft.VisualBasic" %>
<%@  Import Namespace="System.Data"%>
<%@  Import Namespace="System.Data.Odbc" %>
<%@  Import Namespace="System.Web.Services" %> 
<%@  Import Namespace="System.Diagnostics" %>
<%@  Import Namespace="System.Web.UI" %> 
<%@  Import Namespace="System.Web.UI.Webcontrols" %>
<!--@ Page Language="C#" AutoEventWireup="true" CodeFile="importer_job_wilke.cs" Inherits="importer_job_wilke" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">


<script runat="server" language="vbscript">
    
    
    Sub Page_load()

  
   
      
        meMTxt.Text = "HENT DATA"

     
  
     

    End Sub
    
    Sub hentData()
     
        
        
        meMTxt.Text = "HEJ SØREN"
        
        Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_intranet;User=root;Password=;"
       
        Dim objConn As OdbcConnection
        objConn = New OdbcConnection(strConn)
        objConn.Open()
        
        Dim objCmd As OdbcCommand
        'Dim objDataSet As New DataSet
        Dim objDR As OdbcDataReader
        Dim objDR2 As OdbcDataReader
        Dim objDR3 As OdbcDataReader
        Dim dt As DataTable
        Dim dr As DataRow

        Dim tid As Integer
        Dim tmnavn As String
        Dim tmnr As String
        Dim timer As String
        Dim jobnr As String
        Dim jobnavn As String
        Dim tdato As String
        Dim taktivitetnavn, extsysid As String
        Dim taktivitetnavnLst As String = ""
      
        Dim jobstartdato, jobslutdato As String
        Dim jobans, kundenavn, lto As String
        Dim kategori As Integer
        
            
        
           
        'Dim strConn As String
     
        
        Dim Table1 As DataTable
        Table1 = New DataTable("tb_to_var")

        Dim column1 As DataColumn = New DataColumn("jobnr")
        column1.DataType = System.Type.GetType("System.Int32")

        Dim column2 As DataColumn = New DataColumn("jobnavn")
        column2.DataType = System.Type.GetType("System.String")
        
        Dim column3 As DataColumn = New DataColumn("jobansinit")
        column3.DataType = System.Type.GetType("System.String")
        
        Dim column4 As DataColumn = New DataColumn("jobstartdato")
        column4.DataType = System.Type.GetType("System.String")
        
        Dim column5 As DataColumn = New DataColumn("jobslutdato")
        column5.DataType = System.Type.GetType("System.String")
        
        Dim column6 As DataColumn = New DataColumn("kundenavn")
        column6.DataType = System.Type.GetType("System.String")
        
        Dim column7 As DataColumn = New DataColumn("kategori")
        column7.DataType = System.Type.GetType("System.Int32")
        
        Dim column8 As DataColumn = New DataColumn("lto")
        column8.DataType = System.Type.GetType("System.String")
        
        'Dim column3 As DataColumn = New DataColumn("Column2")
        'column3.DataType = System.Type.GetType("System.Int32")

        Table1.Columns.Add(column1)
        Table1.Columns.Add(column2)
        Table1.Columns.Add(column3)
        Table1.Columns.Add(column4)
        Table1.Columns.Add(column5)
        Table1.Columns.Add(column6)
        Table1.Columns.Add(column7)
        Table1.Columns.Add(column8)
        
        Dim row As DataRow
        Dim t As integer = 1
        'AND jobnr BETWEEN 7000 AND 72000
        Dim strSQLext As String = "SELECT jobnr, jobnavn, jobstartdato, jobslutdato, jobans, kategori, kundenavn, lto FROM job_import_temp WHERE id <> 0 AND jobnavn IS NOT NULL ORDER BY jobnr LIMIT 1000" 'AND origin = " & importFrom
        objCmd = New OdbcCommand(strSQLext, objConn)
        objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
        While objDR.Read() = True

            
            
            'tid = objDR("tid")
            'tmnavn = objDR("tmnavn")
            'tmnr = objDR("tmnr")
            'timer = objDR("timer")
            jobnr = objDR("jobnr")
            jobnavn = Replace(objDR("jobnavn"), "'", "")
            'tdato = objDR("tdato")
            'taktivitetnavn = objDR("taktivitetnavn")

            taktivitetnavnLst = taktivitetnavnLst & "<br>"& t & " " & jobnavn &" "& jobnr
            
            jobstartdato = "01-01-2016"  'objDR("jobstartdato")
            jobslutdato = "31-12-2016" 'objDR("jobslutdato")
            If IsDBNull(objDR("jobans")) <> True Then
                jobans = objDR("jobans")
            Else
                jobans = ""
            End If
            
                kategori = objDR("kategori")
                kundenavn = objDR("kundenavn")
                lto = objDR("lto")
            
            
           
                row = Table1.NewRow()
                row("jobnr") = jobnr
                row("jobnavn") = jobnavn
                row("jobstartdato") = jobstartdato
                row("jobslutdato") = jobslutdato
                row("jobansinit") = jobans
                row("kategori") = kategori
                row("kundenavn") = kundenavn
                row("lto") = lto
                Table1.Rows.Add(row)

           

                t = t + 1
            

                    
        End While
        
            
        objDR.Close()
        objConn.Close()
   
        
        datasrc.Text = taktivitetnavnLst
        
        'Dim ws As cati_timeout.CATI = New cati_timeout.CATI
        'ws.upDateCatiTimer(objDataSet)

        Dim ds As DataSet = New DataSet("ds")
        ds.Tables.Add(Table1)
      

     
        Dim CallWebService As New importjob2_ds.oz_importjob2_ds()
        CallWebService.createjob2_ds(ds)

        
        
        
        

        
    End Sub
    
    
    


</script>
   


<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <asp:TextBox runat="server" ID="meid"></asp:TextBox>
    <asp:TextBox runat="server" ID="meMTxt">SØREN</asp:TextBox>
    </div>
    <asp:Button runat="server" Text="Hent data" ID="bt" OnClick="hentData"  />

    <h4>Reading Data from the connection
    <asp:Label ID="datasrc" runat="server"></asp:Label> to the DataGrid</h4>

    <asp:DataGrid ID="dgNameList" runat="server" /><br />
    </form>
</body>
</html>






        




