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
public class ozUploadFileMed
{
    public string id = string.Empty;
    public string dato = string.Empty;
    public string editor = string.Empty;
    public string origin = string.Empty;


    /// <summary>
    /// Bør slettes bruges IKKE
    /// </summary>
    public string medarbejderid = string.Empty;
    public string jobid = string.Empty;
    public string jobnavn = string.Empty;
    public string aktnavn = string.Empty;
    public string aktnr = string.Empty;
    public string akttimer = string.Empty;
    public string akttpris = string.Empty;
    public string aktsum = string.Empty;
    public string timer = string.Empty;
    public string tdato = string.Empty;
   
    public string timerkom = string.Empty;
    public string linjetype = string.Empty;
    public string konto = string.Empty;

    // END 

    public string lto = string.Empty;
    public string minit = string.Empty;
    public string navn = string.Empty;
    public string email = string.Empty;
    public string norm = string.Empty;
    public string mansat = string.Empty;
    public string ansatdato = string.Empty;
    public string opsagtdato = string.Empty;
    public string evn = string.Empty;
    public string costcenter = string.Empty;
    public string linemanager = string.Empty;
    public string countrycode = string.Empty;
    public string weblang = string.Empty;

    public string medstdato = string.Empty;
    public string medsldato = string.Empty;

    public string normthis = string.Empty;

    const int ORIGIN = 10;

	public ozUploadFileMed()
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
            }
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }


        //strRet = "SELECT init FROM medarbejdere where mid = " + midIn; 

        
        return strRet;
    }



    public string Submit(string pathIn, string filenameIn, string[] headers, int[] intHeader, string folderIn, string editorIn, ref int countIgnore, ref int countInserted, string initIn, ref string errorLine, string midIn, string importtype)
    {      
        string strRet = string.Empty;

        try
        {
            List<ozUploadFileMed> lstRow = Read(pathIn, filenameIn, headers, intHeader, ref countIgnore, ref countInserted, initIn, ref errorLine, midIn, folderIn, importtype);
            int intRow = Insert(lstRow, folderIn, editorIn, importtype);
            if (intRow > 0)
            {
                strRet = CallWebservice(folderIn, importtype);
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



    public List<string> ReadHeader(string pathIn, string filenameIn, string folder, string editorIn, string importtype)
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

            connection.Close();
        }

        return isRet;
    }

    private int Insert(List<ozUploadFileMed> lstData, string folder, string editorIn, string importtype)
    {
        int intRow = 0;

        if (folder == "outz" || folder == "intranet - local")
            folder = "intranet";

        string conn = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=outzource; pwd=SKba200473; database=timeout_" + folder + ";";

        using (OdbcConnection connection = new OdbcConnection(conn))
        {

            connection.Open();
            foreach (ozUploadFileMed data in lstData)
            {
              


                if (folder == "tia")
                {

                    if (String.IsNullOrEmpty(data.norm) != true && data.norm != "EXT" && data.norm != "80%")
                    {
                        normthis = data.norm.ToString().Replace(",", ".");
                         } else
                    {
                        normthis = "0";
                    };


                    if (String.IsNullOrEmpty(data.ansatdato) != true) { 
                    medstdato = ConvertDate(data.ansatdato);
                    } else
                    {
                        medstdato = "2002-01-01";
                    };

                    if (String.IsNullOrEmpty(data.opsagtdato) != true)
                    {
                        medsldato = ConvertDate(data.opsagtdato);
                    }
                    else
                    {
                        medsldato = "2002-01-01";
                    };


                    string mnavn = data.navn.ToString().Replace("'", "");

                    string strSQLmedinsTemp = "INSERT INTO med_import_temp_ds (dato, origin, Init, ";
                    strSQLmedinsTemp += "mnavn, ";
                    strSQLmedinsTemp += "email, ";
                    strSQLmedinsTemp += "normtid, ";
                    strSQLmedinsTemp += "ansatdato, ";
                    strSQLmedinsTemp += "opsagtdato, ";
                    strSQLmedinsTemp += "mansat, ";
                    strSQLmedinsTemp += "expvendorno, ";
                    strSQLmedinsTemp += "costcenter, ";
                    strSQLmedinsTemp += "linemanager, ";
                    strSQLmedinsTemp += "countrycode, ";
                    strSQLmedinsTemp += "weblang, lto, editor, overfort) ";
                    strSQLmedinsTemp += " VALUES ('" + DateTime.Now.ToString("yyyy-MM-dd") + "', '910','" + data.minit + "',";
                    strSQLmedinsTemp += "'" + mnavn + "','" + data.email + "',";
                    strSQLmedinsTemp += "'" + normthis + "',";
                    strSQLmedinsTemp += "'" + medstdato + "',";
                    strSQLmedinsTemp += "'" + medsldato +  "','" + data.mansat + "', '" + data.evn + "',";
                    strSQLmedinsTemp += "'" + data.costcenter + "', '" + data.linemanager + "','" + data.countrycode + "', '" + data.weblang + "',";
                    strSQLmedinsTemp += "'tia','Timeout - ImportMedService',0)";

                    OdbcCommand command = new OdbcCommand(strSQLmedinsTemp, connection);

                    // Execute the DataReader and access the data.
                    intRow = command.ExecuteNonQuery();
                    //intRow = "";
                    //command.ExecuteNonQuery();

                }
                else {

                    /// SKAL TILRETTES 20170629
                    string mnavn = data.navn.ToString().Replace("'", "");

                    string strInsert = "INSERT INTO med_import_temp (dato, origin, jobnr, aktnavn, aktnr, beskrivelse, lto, editor, overfort) ";
                    strInsert += " VALUES ('" + DateTime.Now.ToString("yyyy-MM-dd") + "',910,'" + data.jobid + "','" + mnavn + "','" + data.aktnr + "',";
                    strInsert += "'','" + folder + "','Timeout - ImportMedService ',0)";
                    OdbcCommand command = new OdbcCommand(strInsert, connection);

                    // Execute the DataReader and access the data.
                    intRow = command.ExecuteNonQuery();
                    //intRow = "";
                    //command.ExecuteNonQuery();

                }









            }
            connection.Close();
        }

        
        
        return intRow;
    }

    private void MoveFile(string pathIn, string fileIn, string folderIn)
    {
        if (folderIn == "outz" || folderIn == "intranet - local")
            folderIn = "intranet";

        //Move file
        string path = pathIn + fileIn;
        //string path2 = @"C:\www\timeout_xp\wwwroot\ver2_1\inc\upload\" + folderIn + "\\" + fileIn;
        //string path2 = @"D:\webserver\wwwroot\timeout_xp\wwwroot\ver2_10\inc\upload\" + folderIn + "\\" + fileIn;
        string path2 = @"D:\webserver\wwwroot\timeout_xp\wwwroot\ver2_14\inc\upload\" + folderIn + "\\" + fileIn;
        Move(path, path2);
    }

    private int UpdateOverfort(string folderIn, string editorIn)
    {
        int intRow = 0;

        if (folderIn == "outz" || folderIn == "intranet - local")
            folderIn = "intranet";

        string sqlIn = "UPDATE timer_import_temp SET Overfort=1 WHERE lto='" + folderIn + "' AND editor='" + editorIn + "'";
        using (OdbcConnection connection = new OdbcConnection("driver={MySQL ODBC 3.51 Driver};server=195.189.130.210; Port=3306; uid=outzource; pwd=SKba200473; database=timeout_" + folderIn + ";"))
        {
            OdbcCommand command = new OdbcCommand(sqlIn, connection);

            connection.Open();
            intRow = command.ExecuteNonQuery();
            connection.Close();
        }

        return intRow;
    }

    private string CallWebservice(string ltoIn, string importtype)
    {
        if (ltoIn == "outz" || ltoIn == "intranet - local")
            ltoIn = "intranet";

        string strRet = string.Empty;

        //// Create two DataTable instances.
        DataTable table1 = new DataTable("TB_to_var");
        table1.Columns.Add("lto");
        table1.Columns.Add("importtype");
        table1.Rows.Add("" + ltoIn, importtype + "");

        //// Create a DataSet and put both tables in it.
        DataSet dsData = new DataSet("DS_to_var");
        dsData.Tables.Add(table1);

        //DataSet dsData = new DataSet();
      
        //dsData = ReadDS(ltoIn);

        //dk.outzource.to_import service = new dk.outzource.to_import();
        //dk.outzource_addakt.oz_importakt service = new dk.outzource_addakt.oz_importakt();
        //dk_rack.cloud_timeout_importmed.oz_importmed service = new dk_rack.cloud_timeout_importmed.oz_importmed();
        dk_rack.outzource_timeout2_importmed_nav.oz_importmed_na service = new dk_rack.outzource_timeout2_importmed_nav.oz_importmed_na();

        strRet = service.addmed(dsData);

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

    private List<ozUploadFileMed> Read(string pathIn, string filenameIn, string[] headers, int[] intHeader, ref int countIgnore, ref int countInserted, string initIn, ref string errorLines, string midIn, string folderIn, string importtype)
    {
        List<ozUploadFileMed> lstData = new List<ozUploadFileMed>();

        try
        {           
            if (filenameIn.ToLower().Contains(".xlsx"))
                lstData = ReadXlsx2007(pathIn, filenameIn, headers, ref countInserted, ref countIgnore, initIn, ref errorLines, midIn, folderIn, importtype);
            else if (filenameIn.ToLower().Contains(".xls"))
                lstData = ReadXls2003(pathIn, filenameIn, headers, ref countInserted, ref countIgnore, initIn, ref errorLines, midIn, folderIn, importtype);
            else if (filenameIn.ToLower().Contains(".csv"))
                lstData = ReadCsv(pathIn, filenameIn, intHeader, ref countInserted, ref countIgnore, initIn, ref errorLines, midIn, folderIn, importtype);          
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
    private List<ozUploadFileMed> ReadXls2003(string pathIn, string filenameIn, string[] headers, ref int countInserted, ref int countIgnore, string initIn, ref string errorLines, string midIn, string folderIn, string importtype)
    {
        List<ozUploadFileMed> lstRet = new List<ozUploadFileMed>();

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

                ozUploadFileMed fileRet = new ozUploadFileMed();
                fileRet.medarbejderid = row[headers[0]].ToString().Trim();
                fileRet.jobid = row[headers[1]].ToString().Trim();
                fileRet.jobnavn = row[headers[2]].ToString().Trim();
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
                else if (fileRet.jobnavn == string.Empty)
                    fileRet.jobnavn = "0";
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

    private List<ozUploadFileMed> ReadXlsx2007(string pathIn, string filenameIn, string[] headers, ref int countInserted, ref int countIgnore, string initIn, ref string errorLines, string midIn, string folderIn, string importtype)
    {
        List<ozUploadFileMed> lstRet = new List<ozUploadFileMed>();

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

                ozUploadFileMed fileRet = new ozUploadFileMed();
                fileRet.medarbejderid = row[headers[0]].ToString().Trim();
                fileRet.jobid = row[headers[1]].ToString().Trim();
                fileRet.jobnavn = row[headers[2]].ToString().Trim();
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

    private List<ozUploadFileMed> ReadCsv(string pathIn, string filenameIn, int[] headers, ref int countInserted, ref int countIgnore, string initIn, ref string errorLines, string midIn, string folderIn, string importtype)
    {
        List<ozUploadFileMed> lstRet = new List<ozUploadFileMed>();

        try
        {
            string[] allLines = File.ReadAllLines(pathIn + filenameIn, System.Text.Encoding.UTF7);

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

                ozUploadFileMed fileRet = new ozUploadFileMed();
                string[] datas = allLines[i].Split(';');
                fileRet.minit = datas[headers[0]-1];
                fileRet.navn = datas[headers[1]-1];
                fileRet.email = datas[headers[2]-1];
                fileRet.norm = datas[headers[3]-1];
              
                fileRet.ansatdato = datas[headers[4]-1];
                fileRet.opsagtdato = datas[headers[5]-1];
                fileRet.mansat = datas[headers[6]-1];

                fileRet.evn = datas[headers[7]-1];
                fileRet.costcenter = datas[headers[8]-1];
                fileRet.linemanager = datas[headers[9] - 1];
                fileRet.countrycode = datas[headers[10] - 1];
                fileRet.weblang = datas[headers[11] - 1];


                //if (fileRet.jobid == string.Empty)
                //    fileRet.jobid = "0";
                //if (fileRet.medarbejderid == string.Empty)
                //    fileRet.medarbejderid = "0";
                //if (fileRet.timer == string.Empty)
                //    fileRet.timer = "0";
                //if (fileRet.tdato == string.Empty)
                //    fileRet.tdato = "01-01-2001";

                bool isLevelEnough = CheckUserLevel(folderIn, midIn);

                if (isLevelEnough)
                {
                    lstRet.Add(fileRet);
                }
                else
                {
                    if (fileRet.minit == initIn)
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

        using (OdbcConnection connection = new OdbcConnection("driver={MySQL ODBC 3.51 Driver};server=195.189.130.210; Port=3306; uid=outzource; pwd=SKba200473; database=timeout_demo;"))
        {
            OdbcCommand command = new OdbcCommand(sqlIn, connection);

            connection.Open();
            intRow = command.ExecuteNonQuery();
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