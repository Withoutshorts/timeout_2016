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
public class ozUploadFileAkt
{
    public string id = string.Empty;
    public string dato = string.Empty;
    public string editor = string.Empty;
    public string origin = string.Empty;
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
    public string lto = string.Empty;
    public string timerkom = string.Empty;
    public string linjetype = string.Empty;
    public string aktstdato = string.Empty;
    public string konto = string.Empty;

    const int ORIGIN = 10;

	public ozUploadFileAkt()
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
            List<ozUploadFileAkt> lstRow = Read(pathIn, filenameIn, headers, intHeader, ref countIgnore, ref countInserted, initIn, ref errorLine, midIn, folderIn, importtype);
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

    private int Insert(List<ozUploadFileAkt> lstData, string folder, string editorIn, string importtype)
    {
        int intRow = 0;

        if (folder == "outz" || folder == "intranet - local")
            folder = "intranet";

        string conn = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=outzource; pwd=SKba200473; database=timeout_" + folder + ";";

        using (OdbcConnection connection = new OdbcConnection(conn))
        {

            connection.Open();
            foreach (ozUploadFileAkt data in lstData)
            {

                if (importtype == "t2" || importtype == "cflow_prima")
                {

                    if (importtype == "t2")
                    {

                        data.aktnavn = data.aktnavn.Replace("'", "");


                        string strInsert = "INSERT INTO akt_import_temp (dato, origin, jobnr, aktnavn, aktnr, beskrivelse, lto, editor, overfort, akttype, aktstatus) ";
                        strInsert += " VALUES ('" + DateTime.Now.ToString("yyyy-MM-dd") + "',919,'" + data.jobid + "','" + data.aktnavn + "','" + data.jobid + "" + data.aktnr + "',";
                        strInsert += "'','tia','Timeout - ImportAktService ',0,'Posting',1)";
                        OdbcCommand command = new OdbcCommand(strInsert, connection);


                        // Execute the DataReader and access the data.
                        intRow = command.ExecuteNonQuery();
                        //intRow = "";
                        //command.ExecuteNonQuery();

                    } else //Cflow Prima
                    {


                        data.aktnavn = data.aktnavn.Replace("'", "");
                        data.aktstdato = DateTime.Now.ToString("yyyy-MM-dd");

                        string strInsert = "INSERT INTO akt_import_temp (dato, origin, jobnr, aktnavn, aktnr, beskrivelse, lto, editor, overfort, akttype, aktstatus, aktstdato) ";
                        strInsert += " VALUES ('" + DateTime.Now.ToString("yyyy-MM-dd") + "',921,'" + data.jobid + "','" + data.aktnavn + "','" + data.jobid + "" + data.aktnr + "',";
                        strInsert += "'','cflow','Timeout - ImportAktService ',0,'Posting',1,'" + data.aktstdato + "')";
                        OdbcCommand command = new OdbcCommand(strInsert, connection);


                        // Execute the DataReader and access the data.
                        intRow = command.ExecuteNonQuery();
                        //intRow = "";
                        //command.ExecuteNonQuery();
                    }



                }
                else {


                    if  (data.linjetype == "Kontrakt") {
                        data.aktstdato = ConvertDate(data.aktstdato);  //.ToString("yyyy-MM-dd");  //ConvertDate(data.aktstdato);
                      } else {
                         data.aktstdato = DateTime.Now.ToString("yyyy-MM-dd");
                    };

                data.akttimer = data.akttimer.Replace(".", "");
                //data.akttimer = data.akttimer.Replace(".", "");
                data.akttimer = data.akttimer.Replace(",", ".");

                data.akttpris = data.akttpris.Replace(".", "");
                //data.akttpris = data.akttpris.Replace(".", "");
                data.akttpris = data.akttpris.Replace(",", ".");

                data.aktsum = data.aktsum.Replace(".", "");
                //data.aktsum = data.aktsum.Replace(".", "");
                data.aktsum = data.aktsum.Replace(",", ".");

                data.aktnavn = data.aktnavn.Replace("'", "");

               
                //string strInsert = "INSERT INTO timer_import_temp (dato, origin, medarbejderid, jobid, aktnavn, timer, tdato, timerkom, lto, editor,overfort)VALUES('" + DateTime.Now.ToString("yyyy-MM-dd") + "',+"+ORIGIN+",'" + data.medarbejderid + "'," + data.jobid + ",'" + data.aktnavn + "'," + data.timer.Replace(',', '.') + ",'" + data.tdato + "','" + data.timerkom + "','" + folder + "','" + editorIn + "',0)";
                string strInsert = "INSERT INTO akt_import_temp (dato, origin, jobnr, aktnavn, aktnr, akttimer, akttpris, aktsum, beskrivelse, lto, editor, overfort, aktkonto, akttype, aktstdato) ";
                strInsert += " VALUES('" + DateTime.Now.ToString("yyyy-MM-dd") + "',910,'" + data.jobid + "','" + data.aktnavn + "','" + data.aktnr + "'";
                strInsert += "," + data.akttimer + "," + data.akttpris + "," + data.aktsum + ",'','oko','Timeout - ImportAktService ',0, '" + data.konto + "', '" + data.linjetype + "', '" + data.aktstdato + "')";
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
        dk_rack.outzource_timeout2_importakt.oz_importakt service = new dk_rack.outzource_timeout2_importakt.oz_importakt();

     
            strRet = service.addakt(dsData);
       
        

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

    private List<ozUploadFileAkt> Read(string pathIn, string filenameIn, string[] headers, int[] intHeader, ref int countIgnore, ref int countInserted, string initIn, ref string errorLines, string midIn, string folderIn, string importtype)
    {
        List<ozUploadFileAkt> lstData = new List<ozUploadFileAkt>();

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
    private List<ozUploadFileAkt> ReadXls2003(string pathIn, string filenameIn, string[] headers, ref int countInserted, ref int countIgnore, string initIn, ref string errorLines, string midIn, string folderIn, string importtype)
    {
        List<ozUploadFileAkt> lstRet = new List<ozUploadFileAkt>();

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

                ozUploadFileAkt fileRet = new ozUploadFileAkt();
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

    private List<ozUploadFileAkt> ReadXlsx2007(string pathIn, string filenameIn, string[] headers, ref int countInserted, ref int countIgnore, string initIn, ref string errorLines, string midIn, string folderIn, string importtype)
    {
        List<ozUploadFileAkt> lstRet = new List<ozUploadFileAkt>();

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

                ozUploadFileAkt fileRet = new ozUploadFileAkt();
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

    private List<ozUploadFileAkt> ReadCsv(string pathIn, string filenameIn, int[] headers, ref int countInserted, ref int countIgnore, string initIn, ref string errorLines, string midIn, string folderIn, string importtype)
    {
        List<ozUploadFileAkt> lstRet = new List<ozUploadFileAkt>();

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

                ozUploadFileAkt fileRet = new ozUploadFileAkt();
                string[] datas = allLines[i].Split(';');

                fileRet.jobid = datas[headers[0]-1];

                if (importtype != "cflow_prima")
                {
                    fileRet.aktnavn = datas[headers[1] - 1];
                    fileRet.aktnr = datas[headers[2] - 1];
                }

                if (importtype != "t2" && importtype != "cflow_prima") { 
                fileRet.akttimer = datas[headers[3]-1];
                fileRet.akttpris = datas[headers[4]-1];
                fileRet.aktsum = datas[headers[5]-1];
                //fileRet.timerkom = datas[headers[6]-1];
                fileRet.konto = datas[headers[6]-1];
                fileRet.linjetype = datas[headers[7]-1];
                fileRet.aktstdato = datas[headers[8]-1];
                }

                if (importtype == "cflow_prima")
                {
                    fileRet.aktnr = datas[headers[1] - 1];
                    fileRet.aktnavn = datas[headers[2] - 1];
                    fileRet.aktstdato = datas[headers[8] - 1];
                }

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