<%@  Import Namespace="Microsoft.VisualBasic" %>
<%@  Import Namespace="System.Data"%>
<%@  Import Namespace="System.Data.Odbc" %>
<%@  Import Namespace="System.Web.Services" %> 
<%@  Import Namespace="System.Diagnostics" %>
<%@  Import Namespace="System.Web.UI" %> 
<%@  Import Namespace="System.Web.UI.Webcontrols" %>

<script runat="server" language="vbscript">
    

 


    Sub Page_load()

  
   
        
     End Sub




     Public function CallWebservice()
    

       Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_intranet;User=outzource;Password=SKba200473;"
       
           Dim objConn As OdbcConnection
        Dim objCmd As OdbcCommand
        'Dim objDataSet As New DataSet
        Dim objDR As OdbcDataReader
        Dim objDR2 As OdbcDataReader
        Dim objDR3 As OdbcDataReader
        Dim dt As DataTable
        Dim dr As DataRow

        Dim tid AS Integer
        Dim tmnavn AS String
        Dim tmnr AS String
        Dim timer AS String
        Dim tjobnr AS String
        Dim tdato AS String
        Dim taktivitetnavn, extsysid AS String    

        'Dim strConn As String
     
      Dim strSQLext As String = "SELECT tid, tmnavn, tmnr, timer, tjobnr, tdato, taktivitetnavn, extsysid FROM Timer WHERE tdato > '2014-05-01'" 'AND origin = " & importFrom
                        objCmd = New OdbcCommand(strSQLext, objConn)
                        objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)
            
                        If objDR.Read() = True Then

                            tid = objDR("tid") 
                            tmnavn = objDR("tmnavn") 
                            tmnr = objDR("tmnr")
                            timer = objDR("timer")
                            tjobnr = objDR("tjobnr")
                            tdato = objDR("tdato")
                            taktivitetnavn = objDR("taktivitetnavn")

                            extsysid = objDR("extsysid")
                    
                        End If
            
                        objDR.Close()
   
      'Dim ws As cati_timeout.CATI = New cati_timeout.CATI
      'ws.upDateCatiTimer(objDataSet)


      'string strRet = string.Empty;


        '//// Create two DataTable instances.
        DataTable table1 = new DataTable("TB_to_var");
        table1.Columns.Add("lto");
        table1.Rows.Add(""+ ltoIn +"");

        '//// Create a DataSet and put both tables in it.
        DataSet dsData = new DataSet("DS_to_var");
        dsData.Tables.Add(table1);

        '//DataSet dsData = new DataSet();
        '//dsData = ReadDS(ltoIn);

        '//dk.outzource.to_import service = new dk.outzource.to_import();
        '//dk.outzource_importjob.oz_importjob service = new dk.outzource_importjob.oz_importjob();
        dk_rack.outzource_timeout2.oz_importjob2 service = new dk_rack.outzource_timeout2.oz_importjob2();
        strRet = service.createjob2(dsData);

        'return strRet;
   
     

    end function

     

   
        
   
    
</script>


<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <asp:TextBox runat="server" ID="meid"></asp:TextBox>
    <asp:TextBox runat="server" ID="meMTxt"></asp:TextBox>
    </div>
    <asp:Button runat="server" Text="Clikc mig" ID="bt"  />

    <h4>Reading Data from the connection
    <asp:Label ID="datasrc" runat="server"></asp:Label> to the DataGrid</h4>

    <asp:DataGrid ID="dgNameList" runat="server" /><br />
    </form>
</body>
</html>






        




