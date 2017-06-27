<%@ Page Language="vb" Debug="true" %>

<%@ Import Namespace="System.Web.Services.Description" %>

<%@  Import Namespace="Microsoft.VisualBasic" %>
<%@  Import Namespace="System.Data"%>
<%@  Import Namespace="System.Data.Odbc" %>
<%@  Import Namespace="System.Web.Services" %> 
<%@  Import Namespace="System.Diagnostics" %>
<%@  Import Namespace="System.Web.UI" %> 
<%@  Import Namespace="System.Web.UI.Webcontrols" %>
<%@  Import Namespace="System.Collections.Generic" %>
<!--@ Page Language="C#" AutoEventWireup="true" CodeFile="importer_job_wilke.cs" Inherits="importer_job_wilke" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">


<script runat="server" language="vbscript">


    Sub Page_load()




        'meMTxt.Text = "HENT DATA"





    End Sub

    Sub hentData()








        'Dim tmnavn As String
        'Dim tmnr As String
        'Dim timer As String
        'Dim jobnr As String
        'Dim jobnavn As String
        'Dim tdato As String
        'Dim taktivitetnavn, extsysid As String
        Dim taktivitetnavnLst As String = ""

        'Dim jobstartdato, jobslutdato As String
        Dim jobans, kundenavn, lto As String
        'Dim kategori As Integer


        lto = "tia"
        'meMTxt.Text = "HEJ SØREN"

        'Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_intranet;User=root;Password=;"
        Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_" & lto & ";User=to_outzource2;Password=SKba200473;"

        Dim objConn As OdbcConnection
        objConn = New OdbcConnection(strConn)
        objConn.Open()

        Dim objCmd As OdbcCommand
        'Dim objDataSet As New DataSet
        Dim objDR As OdbcDataReader
        'Dim objDR2 As OdbcDataReader
        'Dim objDR3 As OdbcDataReader
        'Dim dt As DataTable
        'Dim dr As DataRow



        Dim row As DataRow
        Dim t As Integer = 1


        'Dim strConn As String


        Dim Table1 As DataTable
        Table1 = New DataTable("tb_to_var")

        Dim column1 As DataColumn = New DataColumn("lto")
        column1.DataType = System.Type.GetType("System.String")

        Dim column2 As DataColumn = New DataColumn("importtype")
        column2.DataType = System.Type.GetType("System.String")

        Dim column3 As DataColumn = New DataColumn("init")
        column3.DataType = System.Type.GetType("System.String")

        Dim column4 As DataColumn = New DataColumn("mnavn")
        column4.DataType = System.Type.GetType("System.String")

        Dim column5 As DataColumn = New DataColumn("email")
        column5.DataType = System.Type.GetType("System.String")

        Dim column6 As DataColumn = New DataColumn("normtid")
        column6.DataType = System.Type.GetType("System.String")

        Dim column7 As DataColumn = New DataColumn("ansatdato")
        column7.DataType = System.Type.GetType("System.String")

        Dim column8 As DataColumn = New DataColumn("opsagtdato")
        column8.DataType = System.Type.GetType("System.String")

        Dim column9 As DataColumn = New DataColumn("mansat")
        column9.DataType = System.Type.GetType("System.String")

        Dim column10 As DataColumn = New DataColumn("expvendorno")
        column10.DataType = System.Type.GetType("System.String")

        Dim column11 As DataColumn = New DataColumn("costcenter")
        column11.DataType = System.Type.GetType("System.String")

        Dim column12 As DataColumn = New DataColumn("linemanager")
        column12.DataType = System.Type.GetType("System.String")

        Dim column13 As DataColumn = New DataColumn("countrycode")
        column13.DataType = System.Type.GetType("System.String")


        Dim column14 As DataColumn = New DataColumn("weblang")
        column14.DataType = System.Type.GetType("System.String")





        Table1.Columns.Add(column1)
        Table1.Columns.Add(column2)
        Table1.Columns.Add(column3)
        Table1.Columns.Add(column4)
        Table1.Columns.Add(column5)
        Table1.Columns.Add(column6)
        Table1.Columns.Add(column7)
        Table1.Columns.Add(column8)

        Table1.Columns.Add(column9)
        Table1.Columns.Add(column10)
        Table1.Columns.Add(column11)
        Table1.Columns.Add(column12)
        Table1.Columns.Add(column13)
        Table1.Columns.Add(column14)




        '***************************************************************************************
        '**** VIA recordset
        '***************************************************************************************

        ''AND jobnr BETWEEN 7000 AND 72000

        Dim strSQLmedins As String = "SELECT id, origin, Init, "
        strSQLmedins += "mnavn, "
        strSQLmedins += "email, "
        strSQLmedins += "normtid, "
        strSQLmedins += "ansatdato, "
        strSQLmedins += "opsagtdato, "
        strSQLmedins += "mansat, "
        strSQLmedins += "expvendorno, "
        strSQLmedins += "costcenter, "
        strSQLmedins += "linemanager, "
        strSQLmedins += "countrycode, "
        strSQLmedins += "weblang, lto, editor, overfort FROM med_import_temp_ds WHERE id > 0 AND overfort = 0 AND errid = 0 ORDER BY id"

        objCmd = New OdbcCommand(strSQLmedins, objConn)
        objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

        While objDR.Read() = True




            row = Table1.NewRow()
            row("lto") = lto
            row("importtype") = "1"
            row("init") = objDR("init")
            row("mnavn") = objDR("mnavn")
            row("email") = objDR("email")
            row("normtid") = objDR("normtid")
            row("ansatdato") = objDR("ansatdato")
            row("opsagtdato") = objDR("opsagtdato")
            row("mansat") = objDR("mansat")

            row("expvendorno") = objDR("expvendorno")
            row("costcenter") = objDR("costcenter")
            row("linemanager") = objDR("linemanager")
            row("countrycode") = objDR("countrycode")
            row("weblang") = objDR("weblang")

            taktivitetnavnLst = taktivitetnavnLst & objDR("init") & "<br>"

            'Add row
            Table1.Rows.Add(row)



            t = t + 1



        End While


        objDR.Close()
        objConn.Close()

        If t > 0 Then

            'datasrc.Text = taktivitetnavnLst


            Dim ds As DataSet = New DataSet("ds")
            ds.Tables.Add(Table1)


            'Dim CallWebService As New dk_rack.cloud_timeout_importmed_ds.oz_importmed_ds()
            'CallWebService.addmed(ds)

        Else

            datasrc.Text = "Der blev ikke fundet medarbejder data at indlæse i denne kørsel"

        End If








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
    <asp:TextBox runat="server" ID="meMTxt" Style="width:600px; height:400px; vertical-align:top;">NAV Data:</asp:TextBox>
    </div>
    <asp:Button runat="server" Text="Hent data" ID="bt" OnClick="hentData"  />

    <h4>Reading Data from the connection
    <asp:Label ID="datasrc" runat="server"></asp:Label> to the DataGrid</h4>

    <asp:DataGrid ID="dgNameList" runat="server" /><br />
    </form>
</body>
</html>






        




