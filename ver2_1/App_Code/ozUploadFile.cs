using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Data.Odbc;
using System.Web.UI;
using System.Web.UI.WebControls;

using System.Globalization;
using System.Threading;

/// <summary>
/// Summary description for ozUploadFile
/// </summary>
public class ozUploadFile
{
    public string id = string.Empty;
    public string dato = string.Empty;
    public string editor = string.Empty;
    public string origin = string.Empty;
    public string medarbejderid = string.Empty;
    public string jobid = string.Empty;
    public string aktnavn = string.Empty;
    public string timer = string.Empty;
    public string tdato = string.Empty;
    public string lto = string.Empty;
    public string timerkom = string.Empty;

    const int ORIGIN = 10;

	public ozUploadFile()
	{
        // Sets the CurrentCulture property to Danish Denmark
        Thread.CurrentThread.CurrentCulture = new CultureInfo("da-DK");
	}

    public string Upload(FileUpload fu, string path)
    {
        string strRet = string.Empty;

        Boolean fileOK = false;
        string fileName = fu.FileName;

        if (fileName == string.Empty)
            strRet = "Filen blev ikke uploadet.";

        if (fu.HasFile)
        {
            String fileExtension =
                System.IO.Path.GetExtension(fu.FileName).ToLower();
            String[] allowedExtensions = { ".xlsx", ".xls", ".csv" };
            for (int i = 0; i < allowedExtensions.Length; i++)
            {
                if (fileExtension == allowedExtensions[i])
                {
                    fileOK = true;
                }
            }
        }

        if (fileOK)
        {
            try
            {
                fu.PostedFile.SaveAs(path + fu.FileName);
                strRet = "Fil uploadet.";
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        else
        {
            strRet = "*Fil format er ikke gyldigt. Brug kun excel filer (.xls, .xlsx, .csv).";
        }

        return strRet;
    }

    public string GetInit(string midIn, string ltoIn)
    {
        string strRet = string.Empty;

        if (ltoIn == "outz" || ltoIn == "intranet - local")
            ltoIn = "intranet";

        //changed back to this query instead of the one below on purpose 01-17-2014 by Lei
        string strSelect = "SELECT init FROM medarbejdere where mid = "+midIn; 

        //string strSelect = "SELECT init, brugergruppe FROM medarbejdere where mid = " + midIn + " AND brugergruppe <> 3";

        try
        {
            using (OdbcConnection connection = new OdbcConnection("driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=outzource; pwd=SKba200473; database=timeout_" + ltoIn + ";"))
            {
                OdbcCommand command = new OdbcCommand(strSelect, connection);

                connection.Open();

                // Execute the DataReader and access the data.
                OdbcDataReader myReader = command.ExecuteReader();
                while (myReader.Read())
                {
                    strRet = myReader.GetValue(0).ToString();
                }

                // Call Close when done reading.
                myReader.Close();
                connection.Close();
            }
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }

        return strRet;
    }

    public string Submit(string pathIn, string filenameIn, string[] headers, int[] intHeader, string folderIn, string editorIn, ref int countIgnore, ref int countInserted, string initIn, ref string errorLine, string midIn)
    {      
        string strRet = string.Empty;

        try
        {
            List<ozUploadFile> lstRow = Read(pathIn, filenameIn, headers, intHeader, ref countIgnore, ref countInserted, initIn, ref errorLine, midIn, folderIn);
            int intRow = Insert(lstRow, folderIn, editorIn);
            if (intRow > 0)
            {
                strRet = CallWebservice(folderIn);
            }

            if (strRet == null)
                strRet = "Null værdier retuneret fra webservice. Der er tomme felter i de importerede data.";

            else if (strRet.ToLower().Contains("succes"))
            {
                UpdateOverfort(folderIn, editorIn);
                MoveFile(pathIn, filenameIn, folderIn);
            }
        }
        catch (Exception ex)
        {
            throw new Exception("Submit error: "+ex.Message);
        }

        return strRet;
    }

    public List<string> ReadHeader(string pathIn, string filenameIn, string folder, string editorIn)
    {
        List<string> headerData = new List<string>();

        try
        {
            if (filenameIn.ToLower().Contains(".xlsx"))
                headerData = ReadXlsx2007Header(pathIn, filenameIn);
            else if (filenameIn.ToLower().Contains(".xls"))
                headerData = ReadXls2003Header(pathIn, filenameIn);
            else if (filenameIn.ToLower().Contains(".csv"))
                headerData = ReadCsvHeader(pathIn, filenameIn);
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }

        return headerData;
    }

    public DataSet ReadDS(string ltoIn)
    {
        DataSet dsRet = new DataSet();

        //intTempImpId    = 0
        //tempM          = 4
        //intJobNr       = 5
        //aktnavnUse      = 6
        //dldTimer       = 7
        //Tdato           = 8
        //LTO             = 9
        //timerkom        = 10

        //, medarbejderid, jobid, aktnavn, timer, tdato, lto, timerkom'
        string strSelect = "SELECT id, dato, editor, origin, medarbejderid, jobid, aktnavn, timer, tdato, lto, timerkom, aktid FROM timer_import_temp WHERE overfort = 0 ORDER BY id limit 3000";
        //'dato, editor, origin,'
        try
        {

            using (OdbcConnection connection = new OdbcConnection("driver={MySQL ODBC 3.51 Driver}; Server=194.150.108.154; Port=3306; Uid=outzource; Pwd=SKba200473; Database=timeout_" + ltoIn + ";"))
            {
                OdbcCommand command = new OdbcCommand(strSelect, connection);

                connection.Open();

                // Execute the DataReader and access the data.
                OdbcDataReader myReader = command.ExecuteReader();

                dsRet = ConvertDataReaderToDataSet(myReader);

                // Call Close when done reading.
                myReader.Close();
                connection.Close();
            }
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }

        return dsRet;
    }

    public DataSet ConvertDataReaderToDataSet(System.Data.Odbc.OdbcDataReader reader)
    {
        DataSet dataSet = new DataSet();

        do
        {
            // Create data table in runtime
            DataTable schemaTable = reader.GetSchemaTable();
            DataTable dataTable = new DataTable();

            if (schemaTable != null)
            {
                DataColumn column1 = new DataColumn("id", typeof(int));
                dataTable.Columns.Add(column1);
                DataColumn column2 = new DataColumn("dato", typeof(DateTime));
                dataTable.Columns.Add(column2);
                DataColumn column3 = new DataColumn("editor", typeof(string));
                dataTable.Columns.Add(column3);
                DataColumn column4 = new DataColumn("origin", typeof(int));
                dataTable.Columns.Add(column4);
                DataColumn column5 = new DataColumn("medarbejderid", typeof(string));
                dataTable.Columns.Add(column5);
                DataColumn column6 = new DataColumn("jobid", typeof(int));
                dataTable.Columns.Add(column6);
                DataColumn column7 = new DataColumn("aktnavn", typeof(string));
                dataTable.Columns.Add(column7);
                DataColumn column8 = new DataColumn("timer", typeof(double));
                dataTable.Columns.Add(column8);
                DataColumn column9 = new DataColumn("tdato", typeof(DateTime));
                dataTable.Columns.Add(column9);
                DataColumn column10 = new DataColumn("lto", typeof(string));
                dataTable.Columns.Add(column10);
                DataColumn column11 = new DataColumn("timerkom", typeof(string));
                dataTable.Columns.Add(column11);

                dataSet.Tables.Add(dataTable);

                // Fill the data table from reader data
                while (reader.Read())
                {
                    DataRow dataRow = dataTable.NewRow();

                    dataRow["id"] = reader.GetValue(0);
                    dataRow["dato"] = reader.GetValue(1);
                    dataRow["editor"] = reader.GetValue(2);
                    dataRow["origin"] = reader.GetValue(3);
                    dataRow["medarbejderid"] = reader.GetValue(4);
                    dataRow["jobid"] = reader.GetValue(5);
                    dataRow["aktnavn"] = reader.GetValue(6);
                    dataRow["timer"] = reader.GetValue(7);
                    dataRow["tdato"] = reader.GetValue(8);
                    dataRow["lto"] = reader.GetValue(9);    //10 Overfort omited here
                    dataRow["timerkom"] = reader.GetValue(11); 

                    dataTable.Rows.Add(dataRow);
                }
            }
            else
            {
                // No records were returned
                DataColumn column = new DataColumn("RowsAffected");
                dataTable.Columns.Add(column);
                dataSet.Tables.Add(dataTable);
                DataRow dataRow = dataTable.NewRow();
                dataRow[0] = reader.RecordsAffected;
                dataTable.Rows.Add(dataRow);
            }
        }
        while (reader.NextResult());

        return dataSet;
    }

    #region Private methods
    /// <summary>
    /// Check if user level is less than 2 or equals to 6, if so return true and has the right to upload for other users, otherwise return false
    /// </summary>
    /// <param name="folder"></param>
    /// <param name="midIn"></param>
    /// <returns>number</returns>
    private bool CheckUserLevel(string folder, string midIn)
    {
        bool isRet = false;

        if (folder == "outz" || folder == "intranet - local")
            folder = "intranet";

        string conn = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=outzource; pwd=SKba200473; database=timeout_" + folder + ";";

        using (OdbcConnection connection = new OdbcConnection(conn))
        {
            string strSelect = "SELECT b.rettigheder FROM medarbejdere AS m LEFT JOIN brugergrupper AS b ON (b.id = m.brugergruppe) WHERE mid = "+midIn;

            OdbcCommand command = new OdbcCommand(strSelect, connection);

            connection.Open();

            // Execute the DataReader and access the data.
            OdbcDataReader myReader = command.ExecuteReader();
            while (myReader.Read())
            {
                string level = myReader.GetValue(0).ToString();
                if (level != string.Empty)
                {
                    int intLevel = int.Parse(level);
                    if (intLevel <= 2 || intLevel == 6)
                    {
                        isRet = true;
                        break;
                    }
                }
            }
            myReader.Close();
            connection.Close();
        }

        return isRet;
    }

    private int Insert(List<ozUploadFile> lstData, string folder, string editorIn)
    {
        int intRow = 0;

        if (folder == "outz" || folder == "intranet - local")
            folder = "intranet";

        string conn = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=outzource; pwd=SKba200473; database=timeout_" + folder + ";";

        using (OdbcConnection connection = new OdbcConnection(conn))
        {
            foreach (ozUploadFile data in lstData)
            {
                
              

                data.tdato = ConvertDate(data.tdato);
                string strInsert = "INSERT INTO timer_import_temp (dato, origin, medarbejderid, jobid, aktnavn, timer, tdato, timerkom, lto, editor, overfort, jobnavn, aktid, errid, errmsg)";
                strInsert += " VALUES('" + DateTime.Now.ToString("yyyy-MM-dd") + "',+"+ORIGIN+",'" + data.medarbejderid + "'," + data.jobid + ",'" + data.aktnavn + "'," + data.timer.Replace(',', '.') + ",'" + data.tdato + "','" + data.timerkom + "','" + folder + "','" + editorIn + "',0,'',0,0,0)";

                OdbcCommand command = new OdbcCommand(strInsert, connection);

                connection.Open();

                // Execute the DataReader and access the data.
                intRow = command.ExecuteNonQuery();

                connection.Close();

               
            }
        }

        return intRow;
    }

    private void MoveFile(string pathIn, string fileIn, string folderIn)
    {
        if (folderIn == "outz" || folderIn == "intranet - local")
            folderIn = "intranet";

        //Move file
        string path = pathIn + fileIn;
        string path2 = @"D:\webserver\wwwroot\timeout_xp\wwwroot\ver2_14\inc\upload\" + folderIn + "\\" + fileIn;
        Move(path, path2);
    }

    private int UpdateOverfort(string folderIn, string editorIn)
    {
        int intRow = 0;

        if (folderIn == "outz" || folderIn == "intranet - local")
            folderIn = "intranet";

        string sqlIn = "UPDATE timer_import_temp SET Overfort=1 WHERE lto='" + folderIn + "' AND editor='" + editorIn + "'";
        using (OdbcConnection connection = new OdbcConnection("driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=outzource; pwd=SKba200473; database=timeout_" + folderIn + ";"))
        {
            OdbcCommand command = new OdbcCommand(sqlIn, connection);

            connection.Open();
            intRow = command.ExecuteNonQuery();
            connection.Close();
        }

        return intRow;
    }

    private string CallWebservice(string ltoIn)
    {
        if (ltoIn == "outz" || ltoIn == "intranet - local")
            ltoIn = "intranet";

        string strRet = string.Empty;

        DataSet dsData = new DataSet();
        dsData = ReadDS(ltoIn);

        //dk.outzource_importhours.to_import service = new dk.outzource_importhours.to_import();
        //dk_rack.outzource_importhours.to_import_hours service = new dk_rack.outzource_importhours.to_import_hours();
        dk_rack.outzource_timeout2_importhours.to_import_hours service = new dk_rack.outzource_timeout2_importhours.to_import_hours();
        service.Timeout = 1920000; // 30 min.

        strRet = service.timeout_importTimer_rack(dsData);

        return strRet;
    }

    private List<string> GetAllFolders(string pathIn)
    {
        List<string> lstRet = new List<string>();

        DirectoryInfo dirInfo = new DirectoryInfo(pathIn);
        DirectoryInfo[] dirInfos = dirInfo.GetDirectories("*.*");

        for (int i = 0; i < dirInfos.Length; i++)
        {
            lstRet.Add(dirInfos[i].Name);
        }

        return lstRet;
    }

    private List<string> GetAllFileNames(string folderIn, string pathIn)
    {
        List<string> lstRet = new List<string>();

        DirectoryInfo dirInfo = new DirectoryInfo(pathIn + folderIn + "\\");
        FileInfo[] files = dirInfo.GetFiles();

        for (int i = 0; i < files.Length; i++)
        {
            lstRet.Add(files[i].Name);
        }

        return lstRet;
    }

    private List<ozUploadFile> Read(string pathIn, string filenameIn, string[] headers, int[] intHeader, ref int countIgnore, ref int countInserted, string initIn, ref string errorLines, string midIn, string folderIn)
    {
        List<ozUploadFile> lstData = new List<ozUploadFile>();

        try
        {           
            if (filenameIn.ToLower().Contains(".xlsx"))
                lstData = ReadXlsx2007(pathIn, filenameIn, headers, ref countInserted, ref countIgnore, initIn, ref errorLines, midIn, folderIn);
            else if (filenameIn.ToLower().Contains(".xls"))
                lstData = ReadXls2003(pathIn, filenameIn, headers, ref countInserted, ref countIgnore, initIn, ref errorLines, midIn, folderIn);
            else if (filenameIn.ToLower().Contains(".csv"))
                lstData = ReadCsv(pathIn, filenameIn, intHeader, ref countInserted, ref countIgnore, initIn, ref errorLines, midIn, folderIn);          
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }

        return lstData;
    }

    private string ConvertDate(string dateIn)
    {
        string strRet = string.Empty;

        try
        {
            strRet = DateTime.Parse(dateIn).ToString("yyyy-MM-dd");
        }
        catch
        {
        }

        return strRet;
    }
   
    #region Read file
    private List<ozUploadFile> ReadXls2003(string pathIn, string filenameIn, string[] headers, ref int countInserted, ref int countIgnore, string initIn, ref string errorLines, string midIn, string folderIn)
    {
        List<ozUploadFile> lstRet = new List<ozUploadFile>();

        try
        {
            OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathIn + filenameIn + ";Extended Properties=Excel 8.0");
            con.Open();
            OleDbDataAdapter da = new OleDbDataAdapter("select * from [Sheet1$]", con);
            DataTable dt = new DataTable();
            da.Fill(dt);

            countIgnore = 0;
            countInserted = 0;
            int countAll = 0;
            foreach (DataRow row in dt.Rows)
            {
                countAll++;

                ozUploadFile fileRet = new ozUploadFile();
                fileRet.medarbejderid = row[headers[0]].ToString().Trim();
                fileRet.jobid = row[headers[1]].ToString().Trim();
                fileRet.timer = row[headers[2]].ToString().Trim();
                fileRet.tdato = row[headers[3]].ToString().Trim();
                fileRet.aktnavn = row[headers[4]].ToString().Trim();
                fileRet.timerkom = row[headers[5]].ToString().Trim();

                if (fileRet.medarbejderid == string.Empty && fileRet.jobid == string.Empty && fileRet.timer == string.Empty && fileRet.tdato == string.Empty && fileRet.aktnavn == string.Empty && fileRet.timerkom == string.Empty)
                {
                    countIgnore++;
                    errorLines += "Linje " + countAll + " - fejlkode: Tom linje.<br>";
                    continue;
                }
                else if (fileRet.jobid == string.Empty)
                    fileRet.jobid = "0";
                else if (fileRet.jobid == string.Empty)
                    fileRet.jobid = "0";
                else if (fileRet.jobid == string.Empty)
                    fileRet.jobid = "0";
                else if (fileRet.medarbejderid == string.Empty)
                    fileRet.medarbejderid = "0";
                else if (fileRet.timer == string.Empty)
                    fileRet.timer = "0";
                else if (fileRet.tdato == string.Empty)
                    fileRet.tdato = "01-01-2001";

                 bool isLevelEnough = CheckUserLevel(folderIn, midIn);

                 if (isLevelEnough)
                 {
                     lstRet.Add(fileRet);
                 }
                 else
                 {
                     if (fileRet.medarbejderid == initIn)// check level <=2 || =6
                         lstRet.Add(fileRet);
                     else
                     {
                         errorLines += "Linje "+countAll + " - fejlkode: Du har ikke rettigheder til at uploade timer for andre brugere end dig selv.<br>";
                         countIgnore++;
                     }
                 }
            }

            countInserted = lstRet.Count;

            con.Close();
        }
        catch(Exception ex)
        {
            throw new Exception("Der opstod en fejl i læsning af .xls filen: " + ex.Message);
        }        

        return lstRet;
    }

    private List<string> ReadXls2003Header(string pathIn, string filenameIn)
    {
        List<string> lstRet = new List<string>();

        OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathIn + filenameIn + ";Extended Properties='Excel 8.0;HDR=Yes'");
        con.Open();
        OleDbDataAdapter da = new OleDbDataAdapter("select * from [Sheet1$]", con);
        DataTable dt = new DataTable();
        da.Fill(dt);

        foreach (DataColumn column in dt.Columns)
        {
            lstRet.Add(column.ColumnName);
        }

        con.Close();

        return lstRet;
    }

    private List<ozUploadFile> ReadXlsx2007(string pathIn, string filenameIn, string[] headers, ref int countInserted, ref int countIgnore, string initIn, ref string errorLines, string midIn, string folderIn)
    {
        List<ozUploadFile> lstRet = new List<ozUploadFile>();

        try
        {
            OleDbConnection con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathIn + filenameIn + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=1\";");
            con.Open();
            OleDbDataAdapter da = new OleDbDataAdapter("select * from [Sheet1$]", con);
            DataTable dt = new DataTable();
            da.Fill(dt);

            countIgnore = 0;
            countInserted = 0;
            int countAll = 0;
            foreach (DataRow row in dt.Rows)
            {
                countAll++;

                ozUploadFile fileRet = new ozUploadFile();
                fileRet.medarbejderid = row[headers[0]].ToString().Trim();
                fileRet.jobid = row[headers[1]].ToString().Trim();
                fileRet.timer = row[headers[2]].ToString().Trim();
                fileRet.tdato = row[headers[3]].ToString().Trim();
                fileRet.aktnavn = row[headers[4]].ToString().Trim();
                fileRet.timerkom = row[headers[5]].ToString().Trim();

                if (fileRet.medarbejderid == string.Empty && fileRet.jobid == string.Empty && fileRet.timer == string.Empty && fileRet.tdato == string.Empty && fileRet.aktnavn == string.Empty && fileRet.timerkom == string.Empty)
                {
                    countIgnore++;
                    errorLines += "Linje " + countAll + " - fejlkode: Tom linje.<br>";
                    continue;
                }
                else if (fileRet.jobid == string.Empty)
                    fileRet.jobid = "0";
                else if (fileRet.jobid == string.Empty)
                    fileRet.jobid = "0";
                else if (fileRet.medarbejderid == string.Empty)
                    fileRet.medarbejderid = "0";
                else if (fileRet.timer == string.Empty)
                    fileRet.timer = "0";
                else if (fileRet.tdato == string.Empty)
                    fileRet.tdato = "01-01-2001";

                 bool isLevelEnough = CheckUserLevel(folderIn, midIn);

                 if (isLevelEnough)
                 {
                     lstRet.Add(fileRet);
                 }
                 else
                 {
                     if (fileRet.medarbejderid == initIn)
                         lstRet.Add(fileRet);
                     else
                     {
                         errorLines += "Linje " + countAll + " - fejlkode: Du har ikke rettigheder til at uploade timer for andre brugere end dig selv.<br>";
                         countIgnore++;
                     }
                 }
            }

            countInserted = lstRet.Count;


            con.Close();
        }
        catch (Exception ex)
        {
            throw new Exception("Der opstod en fejl i læsning af .xlsx filen: " + ex.Message);
        }        

        return lstRet;
    }

    private List<string> ReadXlsx2007Header(string pathIn, string filenameIn)
    {
        List<string> lstRet = new List<string>();
        try
        {          
            OleDbConnection con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathIn + filenameIn + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=1\";");
            con.Open();
            OleDbDataAdapter da = new OleDbDataAdapter("select * from [Sheet1$]", con);
            DataTable dt = new DataTable();
            da.Fill(dt);

            foreach (DataColumn column in dt.Columns)
            {
                lstRet.Add(column.ColumnName);
            }

            con.Close();
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }

        return lstRet;
    }

    private List<ozUploadFile> ReadCsv(string pathIn, string filenameIn, int[] headers, ref int countInserted, ref int countIgnore, string initIn, ref string errorLines, string midIn, string folderIn)
    {
        List<ozUploadFile> lstRet = new List<ozUploadFile>();

        try
        {
            string[] allLines = File.ReadAllLines(pathIn + filenameIn);

            countIgnore = 0;
            countInserted = 0;
            int countAll = 0;
            for (int i = 1; i < allLines.Count(); i++)
            {
                countAll++;
                if (allLines[i].Replace(";", "").Trim() == string.Empty)
                {
                    countIgnore++;
                    errorLines += "Linje " + countAll + " - fejlkode: Tom linje.<br>";
                    continue;
                }

                ozUploadFile fileRet = new ozUploadFile();
                string[] datas = allLines[i].Split(';');
                fileRet.medarbejderid = datas[headers[0]-1];
                fileRet.jobid = datas[headers[1]-1];
                fileRet.timer = datas[headers[2]-1];
                fileRet.tdato = datas[headers[3]-1];
                fileRet.aktnavn = datas[headers[4]-1];
                fileRet.timerkom = datas[headers[5]-1];

                if (fileRet.jobid == string.Empty)
                    fileRet.jobid = "0";
                if (fileRet.medarbejderid == string.Empty)
                    fileRet.medarbejderid = "0";
                if (fileRet.timer == string.Empty)
                    fileRet.timer = "0";
                if (fileRet.tdato == string.Empty)
                    fileRet.tdato = "01-01-2001";

                bool isLevelEnough = CheckUserLevel(folderIn, midIn);

                if (isLevelEnough)
                {
                    lstRet.Add(fileRet);
                }
                else
                {
                    if (fileRet.medarbejderid == initIn)
                        lstRet.Add(fileRet);
                    else
                    {
                        errorLines += "Linje " + countAll + " - fejlkode: Du har ikke rettigheder til at uploade timer for andre brugere end dig selv.<br>";
                        countIgnore++;
                    }
                }
            }

            countInserted = lstRet.Count;
        }
         catch (Exception ex)
         {
             throw new Exception("Der opstod en fejl i læsning af .csv filen: " + ex.Message);
         }     

        return lstRet;
    }

    private List<string> ReadCsvHeader(string pathIn, string filenameIn)
    {
        List<string> lstRet = new List<string>();

        string[] allLines = File.ReadAllLines(pathIn + filenameIn);

        string[] datas = allLines[0].Split(';');

        for (int i = 0; i < datas.Length; i++)
        {
            lstRet.Add(datas[i].ToString());
        }

        return lstRet;
    }
    #endregion

    private int Update(string sqlIn)
    {
        int intRow = 0;

        using (OdbcConnection connection = new OdbcConnection("driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=outzource; pwd=SKba200473; database=timeout_demo;"))
        {
            OdbcCommand command = new OdbcCommand(sqlIn, connection);

            connection.Open();
            intRow = command.ExecuteNonQuery();
            connection.Close();
        }

        return intRow;
    }

    private string Move(string pathSource, string pathTarget)
    {
        string strRet = string.Empty;

        try
        {
            if (!File.Exists(pathSource))
            {
                // This statement ensures that the file is created,
                // but the handle is not kept.
                using (FileStream fs = File.Create(pathSource)) { }
            }

            // Ensure that the target does not exist.
            if (File.Exists(pathTarget))
                File.Delete(pathTarget);

            // Move the file.
            File.Move(pathSource, pathTarget);
            strRet = pathSource + " var fjernet til " + pathTarget + ".";

            // See if the original exists now.
            if (File.Exists(pathSource))
            {
                strRet = "En fil med samme navn findes allerede.";
            }
            else
            {
                strRet = "Filen findes ikke længere.";
            }

        }
        catch (Exception e)
        {
            throw new Exception(e.Message);
        }

        return strRet;
    }

    #endregion

    
}