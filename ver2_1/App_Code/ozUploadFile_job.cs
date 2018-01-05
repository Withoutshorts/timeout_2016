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
public class ozUploadFileJob
{

    // Standard job opret 
    // Wilke, Øko
    public string id = string.Empty;
    public string dato = string.Empty;
    public string editor = string.Empty;
    public string origin = string.Empty;
    public string medarbejderid = string.Empty;
    public string jobid = string.Empty;
    public string jobnavn = string.Empty;
    public string jobans = string.Empty;
    public string aktnavn = string.Empty;
    public string timer = string.Empty;
    public string stdato = string.Empty;
    public string sldato = string.Empty;
    public string lto = string.Empty;
    public string timerkom = string.Empty;
    public string jq_str = string.Empty;

    //D1
    public string kundenr = string.Empty;
    public string kpers = string.Empty;
    public string rekvnr = string.Empty;

    //T1 TIA
    public string projgrp = string.Empty;

    //Monitor
    public string kundenavn = string.Empty;
    public string stkantal = string.Empty;
    public string aktstdato = string.Empty;
    public string aktsldato = string.Empty;
    public string sort = string.Empty;
    public string fomr = string.Empty;
    public string aktvarenr = string.Empty;

    const int ORIGIN = 10;
    //const int WIN_1252_CP = 1252; // Windows ANSI codepage 1252

	public ozUploadFileJob()
	{
        // Sets the CurrentCulture property to Danish Denmark
        Thread.CurrentThread.CurrentCulture = new CultureInfo("da-DK");
        
    }

    public string Upload(FileUpload fu, string path)
    {
        string strRet = string.Empty;

        Boolean fileOK = false;
        string fileName = fu.FileName;
        //string fileencoding = fu.GetEncoding();

        if (fileName == string.Empty)
            strRet = "Filen blev ikke uploadet.";

        if (fu.HasFile)
        {
            String fileExtension =
                System.IO.Path.GetExtension(fu.FileName).ToLower();
                //System.IO.StreamReader(fu.FileName);

              
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

        //if (ltoIn == "dencker")
        //    ltoIn = "dencker_test";

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
            List<ozUploadFileJob> lstRow = Read(pathIn, filenameIn, headers, intHeader, ref countIgnore, ref countInserted, initIn, ref errorLine, midIn, folderIn, importtype);
            int intRow = Insert(lstRow, folderIn, editorIn, importtype);
            if (intRow > 0)
            {
                //importtype = "D6";
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

        //if (folder == "dencker")
        //    folder = "dencker_test";

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



    private int Insert(List<ozUploadFileJob> lstData, string folder, string editorIn, string importtype)
    {


       

        int intRow = 0;

        if (folder == "outz" || folder == "intranet - local")
            folder = "intranet";

        // if (folder == "dencker")
        //     folder = "dencker_test";

        string conn = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=to_outzource2; pwd=SKba200473; database=timeout_" + folder + ";";

        using (OdbcConnection connection = new OdbcConnection(conn))
        {
            foreach (ozUploadFileJob data in lstData)
            {

                stdato = ConvertDate(data.stdato);
                sldato = ConvertDate(data.sldato);


                //string importtype = data.importtype;
                //Dencker Monitor
                if (folder == "dencker_test" || folder == "dencker" || folder == "tia")
                {

                    aktstdato = ConvertDate(data.aktstdato);
                    aktsldato = ConvertDate(data.aktsldato);

                    if ((importtype == "d1") || importtype == "t1") {

                        if (importtype == "d1") {

                            //jobnavn = data.jobnavn.Replace("'", "");
                            //jobnavn = jobnavn.Replace("'\'", "");

                            string strInsert = "INSERT INTO job_import_temp (dato, origin, jobnr, jobnavn, jobans, jobstartdato, jobslutdato, beskrivelse, lto, editor, overfort, kundenavn, kundenr, kpers, rekvnr)"+
                            " VALUES ('" + DateTime.Now.ToString("yyyy-MM-dd") + "',602,'" + data.jobid + "','" + data.jobnavn.Replace("'", "") + "','" + data.jobans + "','" + stdato + "','" + sldato + "',"+
                            "'" + data.timerkom + "','" + folder + "','Timeout - ImportJobService',0, '" + data.kundenavn + "', '" + data.kundenr + "', '"+ data.kpers.Replace("'", "") + "', '"+ data.rekvnr + "')";
                            OdbcCommand command = new OdbcCommand(strInsert, connection);
                            connection.Open();

                        // Execute the DataReader and access the data.
                        intRow = command.ExecuteNonQuery();

                        connection.Close();

                        }


                        if (importtype == "t1") {

                            //data.timerkom = "";
                            data.projgrp = "10";

                            string strInsert = "INSERT INTO job_import_temp (dato, origin, jobnr, jobnavn, jobans, jobstartdato, jobslutdato, beskrivelse, lto, editor, overfort, kundenavn, kundenr, projgrp) " +
                            " VALUES ('" + DateTime.Now.ToString("yyyy-MM-dd") + "',603,'" + data.jobid + "','" + data.jobnavn.Replace("'", "") + "','" + data.jobans + "'," +
                            "'" + stdato + "','" + sldato + "','" + data.timerkom + "','" + folder + "','Timeout - ImportJobService',0, '" + data.kundenavn + "', '" + data.kundenr + "', '" + data.projgrp + "')";
                            OdbcCommand command = new OdbcCommand(strInsert, connection);
                            connection.Open();

                            // Execute the DataReader and access the data.
                            intRow = command.ExecuteNonQuery();

                            connection.Close();

                        }


                    }
                    else { 

                    string aktnavn = data.aktnavn.ToString();

                    //System.Text.Encoding iso = System.Text.Encoding.GetEncoding("ISO-8859-1");
                    //System.Text.Encoding utf8 = System.Text.Encoding.UTF8;
                    //byte[] utfBytes = utf8.GetBytes(aktnavn);
                    //byte[] isoBytes = System.Text.Encoding.Convert(utf8, iso, utfBytes);
                    //aktnavn = iso.GetString(isoBytes);
                    
                    aktnavn = aktnavn.Replace("'", "");
                    //aktstdato = DateTime.Now.ToString("yyyy-MM-dd");

                    //string strInsert = "INSERT INTO job_import_temp (dato, origin, jobnr, jobnavn, jobans, jobstartdato, jobslutdato, aktnavn, lto, editor, overfort, " +
                    //" kundenavn, stkantal, aktstdato, aktsldato, sort, fomr, aktvarenr)" +
                    //" VALUES('" + DateTime.Now.ToString("yyyy-MM-dd") + "',600,'" + data.jobid + "','" + data.jobnavn.Replace("'", "") + "','" + data.jobans + "'," +
                    //"'" + stdato + "','" + sldato + "','" + data.aktnavn + "','" + folder + "','Timeout - ImportJobService',0 " +
                    //"'" + data.kundenavn.Replace("'", "") + "', " + data.stkantal + ",'"+ data.aktstdato + "', '" + data.aktsldato + "'," + data.sort + ",'"+ data.fomr +"','"+ data.aktvarenr +"'" +
                    //")";

                        string strInsert = "INSERT INTO job_import_temp (dato, origin, jobnr, jobnavn, jobans, jobstartdato, jobslutdato, beskrivelse, lto, editor, overfort, kundenavn, aktstdato, aktsldato, stkantal, aktnavn, sort, fomr, aktvarenr) " +
                        " VALUES ('" + DateTime.Now.ToString("yyyy-MM-dd") + "',600,'" + data.jobid + "','" + data.jobnavn.Replace("'", "") + "','" + data.jobans + "','" + stdato + "',"+
                        "'" + sldato + "','" + data.timerkom.Replace("'", "") + "','" + folder + "','Timeout - ImportJobService',0,'" + data.kundenavn + "','" + aktstdato + "','" + aktsldato + "', " + data.stkantal + ", '" + aktnavn + "'," + data.sort + ",'" + data.fomr + "','" + data.aktvarenr + "')";
                        OdbcCommand command = new OdbcCommand(strInsert, connection);
                        connection.Open();

                        // Execute the DataReader and access the data.
                        intRow = command.ExecuteNonQuery();

                        connection.Close();
                    }

                    

                   

                }
                //Standard Wilke, ØKO
                else {

                    //string strInsert = "INSERT INTO timer_import_temp (dato, origin, medarbejderid, jobid, aktnavn, timer, tdato, timerkom, lto, editor,overfort)VALUES('" + DateTime.Now.ToString("yyyy-MM-dd") + "',+"+ORIGIN+",'" + data.medarbejderid + "'," + data.jobid + ",'" + data.aktnavn + "'," + data.timer.Replace(',', '.') + ",'" + data.stdato + "','" + data.timerkom + "','" + folder + "','" + editorIn + "',0)";
                    string strInsert = "INSERT INTO job_import_temp (dato, origin, jobnr, jobnavn, jobans, jobstartdato, jobslutdato, beskrivelse, lto, editor, overfort) "+
                    " VALUES ('" + DateTime.Now.ToString("yyyy-MM-dd") + "',600,'" + data.jobid + "','" + data.jobnavn.Replace("'", "") + "','" + data.jobans + "',"+
                    "'" + stdato + "','" + sldato + "','" + data.timerkom.Replace("'", "") + "','" + folder + "','Timeout - ImportJobService',0)";
                    OdbcCommand command = new OdbcCommand(strInsert, connection);

                    connection.Open();

                    // Execute the DataReader and access the data.
                    intRow = command.ExecuteNonQuery();

                    connection.Close();
                }

                

               
            }
        }

        return intRow;
    }

    private void MoveFile(string pathIn, string fileIn, string folderIn)
    {
        if (folderIn == "outz" || folderIn == "intranet - local")
            folderIn = "intranet";

        //if (folderIn == "dencker")
        //    folderIn = "dencker_test";

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

        //if (folderIn == "dencker")
        //folderIn = "dencker_test";

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

    private string CallWebservice(string ltoIn, string importtype)
    {
        if (ltoIn == "outz" || ltoIn == "intranet - local")
            ltoIn = "intranet";

        //if (ltoIn == "dencker")
        //    ltoIn = "dencker_test";

        string strRet = string.Empty;
        //importtype = "D9";

        //// Create two DataTable instances.
        DataTable table1 = new DataTable("TB_to_var");
        table1.Columns.Add("lto");
        table1.Columns.Add("importtype");
        table1.Rows.Add(""+ ltoIn, importtype + "");

        //// Create a DataSet and put both tables in it.
        DataSet dsData = new DataSet("DS_to_var");
        dsData.Tables.Add(table1);

        //DataSet dsData = new DataSet();
        //dsData = ReadDS(ltoIn);

        //dk.outzource.to_import service = new dk.outzource.to_import();
        //dk.outzource_importjob.oz_importjob service = new dk.outzource_importjob.oz_importjob();
        dk_rack.outzource_timeout2.oz_importjob2 service = new dk_rack.outzource_timeout2.oz_importjob2();
        strRet = service.createjob2(dsData);

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

    private List<ozUploadFileJob> Read(string pathIn, string filenameIn, string[] headers, int[] intHeader, ref int countIgnore, ref int countInserted, string initIn, ref string errorLines, string midIn, string folderIn, string importtype)
    {
        List<ozUploadFileJob> lstData = new List<ozUploadFileJob>();

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
    private List<ozUploadFileJob> ReadXls2003(string pathIn, string filenameIn, string[] headers, ref int countInserted, ref int countIgnore, string initIn, ref string errorLines, string midIn, string folderIn, string importtype)
    {
        List<ozUploadFileJob> lstRet = new List<ozUploadFileJob>();

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

                ozUploadFileJob fileRet = new ozUploadFileJob();
                fileRet.medarbejderid = row[headers[0]].ToString().Trim();
                fileRet.jobid = row[headers[1]].ToString().Trim();
                fileRet.jobnavn = row[headers[2]].ToString().Trim();
                fileRet.stdato = row[headers[3]].ToString().Trim();
                fileRet.aktnavn = row[headers[4]].ToString().Trim();
                fileRet.timerkom = row[headers[5]].ToString().Trim();

                if (fileRet.medarbejderid == string.Empty && fileRet.jobid == string.Empty && fileRet.timer == string.Empty && fileRet.stdato == string.Empty && fileRet.aktnavn == string.Empty && fileRet.timerkom == string.Empty)
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
                else if (fileRet.stdato == string.Empty)
                    fileRet.stdato = "01-01-2001";

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

    private List<ozUploadFileJob> ReadXlsx2007(string pathIn, string filenameIn, string[] headers, ref int countInserted, ref int countIgnore, string initIn, ref string errorLines, string midIn, string folderIn, string importtype)
    {
        List<ozUploadFileJob> lstRet = new List<ozUploadFileJob>();

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

                ozUploadFileJob fileRet = new ozUploadFileJob();

               

                fileRet.jobnavn = row[headers[0]].ToString().Trim();
                fileRet.jobid = row[headers[1]].ToString().Trim();

                fileRet.stdato = row[headers[3]].ToString().Trim();
                fileRet.sldato = row[headers[4]].ToString().Trim();

                if (folderIn != "dencker" && folderIn != "dencker_test")
                {
                    fileRet.timerkom = row[headers[5]].ToString().Trim(); //Faktureringstype
                    fileRet.jobans = row[headers[2]].ToString().Trim(); //Jobans

                    if (fileRet.timerkom == string.Empty)
                        fileRet.timerkom = "";
                }

                if (fileRet.jobid == string.Empty)
                    fileRet.jobid = "0";
                if (fileRet.timer == string.Empty)
                    fileRet.timer = "0";
                if (fileRet.stdato == string.Empty)
                    fileRet.stdato = "01-01-2001";
                if (fileRet.sldato == string.Empty)
                    fileRet.sldato = "01-01-2001";




                if (folderIn == "dencker" || folderIn == "dencker_test")
                {
                    //Monitor Dencker

                    fileRet.stkantal = row[headers[5]].ToString().Trim();
                    fileRet.aktstdato = row[headers[8]].ToString().Trim();
                    fileRet.aktsldato = row[headers[9]].ToString().Trim();
                    fileRet.sort = row[headers[6]].ToString().Trim();
                    fileRet.fomr = row[headers[7]].ToString().Trim();
                    fileRet.aktnavn = row[headers[10]].ToString().Trim();


                    fileRet.aktvarenr = row[headers[11]].ToString().Trim();
                    fileRet.kundenavn = row[headers[2]].ToString().Trim(); ;


                    if (fileRet.stkantal == string.Empty)
                        fileRet.stkantal = "0";
                    if (fileRet.aktstdato == string.Empty)
                        fileRet.aktstdato = "01-01-2001";
                    if (fileRet.aktsldato == string.Empty)
                        fileRet.aktsldato = "01-01-2001";

                    if (fileRet.sort == string.Empty)
                        fileRet.sort = "0";
                    if (fileRet.fomr == string.Empty)
                        fileRet.fomr = "0";
                    if (fileRet.aktnavn == string.Empty)
                        fileRet.aktnavn = "-";
                    if (fileRet.aktvarenr == string.Empty)
                        fileRet.aktvarenr = "0";
                    if (fileRet.kundenavn == string.Empty)
                        fileRet.kundenavn = "-";

                }

                //

          

              
                //


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

    private List<ozUploadFileJob> ReadCsv(string pathIn, string filenameIn, int[] headers, ref int countInserted, ref int countIgnore, string initIn, ref string errorLines, string midIn, string folderIn, string importtype)
    {
        List<ozUploadFileJob> lstRet = new List<ozUploadFileJob>();

        try
        {



            string[] allLines = File.ReadAllLines(pathIn + filenameIn, System.Text.Encoding.Default);

            //GetEncoding(1250)
            //System.Text.Encoding.GetEncoding(1252) 
            //System.Text.Encoding.UTF7
            //iso-8859-1

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

                ozUploadFileJob fileRet = new ozUploadFileJob();
                string[] datas = allLines[i].Split(';');



                fileRet.jobnavn = datas[headers[0] - 1];
                fileRet.jobid = datas[headers[1]-1];
               
                fileRet.stdato = datas[headers[3]-1];
                fileRet.sldato = datas[headers[4]-1];

                if (folderIn == "oko")
                {
                    fileRet.timerkom = datas[headers[5] - 1]; //Faktureringstype
                    fileRet.jobans = datas[headers[2] - 1]; //Jobans

                    if (fileRet.timerkom == string.Empty)
                        fileRet.timerkom = "";
                } else { 


                if ((folderIn != "dencker" && folderIn != "dencker_test") || (importtype == "d1") || (importtype == "t1"))
                {
                    fileRet.timerkom = datas[headers[5] - 1]; //Faktureringstype
                    fileRet.jobans = datas[headers[2] - 1]; //Jobans

                    fileRet.kundenavn = datas[headers[6] - 1];
                    fileRet.kundenr = datas[headers[7] - 1];

                        if (fileRet.timerkom == string.Empty)
                        fileRet.timerkom = "";
                }


                    if (importtype == "d1")
                    {
                        fileRet.kpers = datas[headers[9] - 1];
                    fileRet.rekvnr = datas[headers[10] - 1];
                    }


                    if (importtype == "t1")
                {
                    fileRet.projgrp = datas[headers[8] - 1];
                }

                if (fileRet.jobid == string.Empty)
                    fileRet.jobid = "0";
                if (fileRet.timer == string.Empty)
                    fileRet.timer = "0";
                if (fileRet.stdato == string.Empty)
                    fileRet.stdato = "01-01-2001";
                if (fileRet.sldato == string.Empty)
                    fileRet.sldato = "01-01-2001";
             

              

                if ((folderIn == "dencker" || folderIn == "dencker_test") && (importtype != "d1")) { 
                //Monitor Dencker
                
                fileRet.stkantal = datas[headers[5] - 1];
                fileRet.aktstdato = datas[headers[8] - 1];
                fileRet.aktsldato = datas[headers[9] - 1];
                fileRet.sort = datas[headers[6] - 1];
                fileRet.fomr = datas[headers[7] - 1];
                fileRet.aktnavn = datas[headers[10] - 1];

                    
                    fileRet.aktvarenr = datas[headers[11] - 1];
                    fileRet.kundenavn = datas[headers[2] - 1];

                   
                    if (fileRet.stkantal == string.Empty)
                        fileRet.stkantal = "0";
                    if (fileRet.aktstdato == string.Empty)
                        fileRet.aktstdato = "01-01-2001";
                    if (fileRet.aktsldato == string.Empty)
                        fileRet.aktsldato = "01-01-2001";

                    if (fileRet.sort == string.Empty)
                        fileRet.sort = "0";
                    if (fileRet.fomr == string.Empty)
                        fileRet.fomr = "0";
                    if (fileRet.aktnavn == string.Empty)
                        fileRet.aktnavn = "-";
                    if (fileRet.aktvarenr == string.Empty)
                        fileRet.aktvarenr = "0";
                    if (fileRet.kundenavn == string.Empty)
                        fileRet.kundenavn = "-";

                }


                } //OKO





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

        string[] allLines = File.ReadAllLines(pathIn + filenameIn, System.Text.Encoding.Default);
        //ASCII
        //iso-8859-1
        //GetEncoding(1250)

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
                //fs = new FileStream(fileName, FileMode.CreateNew, System.Text.Encoding.UTF8);
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