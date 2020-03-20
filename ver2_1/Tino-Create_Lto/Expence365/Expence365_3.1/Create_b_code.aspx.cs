using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.Odbc;


public partial class Create_b_code : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        


    }

    // Denne metode er knappen som bliver kladt i Create_b_code.aspx under submit knappen.
    protected void BtnUpload_Click(object sender, EventArgs e)
    {

       // Opret_firma_I_egen_database();

       Opret_ny_log();

       Opret_ny_upload_folder();

       // Opret_ny_default_ASP();    Her skal pathen laves om!

       Opret_ny_database_for_bruger();

    }

    // Denne metode opretter et nyt table i databasen timeout_admin, og indsætter tablet med vadierne givet fra den nye kunde.
    protected void Opret_firma_I_egen_database()
    {

        string f_navn = Request.Form["CompanyName"];
        string f_initialer = Request.Form["CompanyInitials"];
        string f_Adresse = Request.Form["Address"];
        string f_by = Request.Form["City"];
        string f_postnr = Request.Form["ZipCode"];
        string f_telefonnummer = Request.Form["CompanyPhone"];
        string f_EAN_nummer = Request.Form["EANNumber"];
        string f_CVR_nummer = Request.Form["CVRNumber"];
        string Kontakt_person = Request.Form["ContactPerson"];
        string kp_Email = Request.Form["Email"];
        string kp_telefonnummer = Request.Form["ContactPersonPhone"];


        // connection to database
        var builder_for_admin_creation = new MySqlConnectionStringBuilder
        {
            Server = "194.150.108.154",
            Database = "",
            UserID = "to_tino",
            Password = "2HotPotatoes",
        };

        using (var connToDatabase = new MySqlConnection(builder_for_admin_creation.ConnectionString))
        {
          

            // create new table in timeout_admin
            using (var command = connToDatabase.CreateCommand())
            {
                command.Connection.Open();

                command.CommandText = "CREATE TABLE "+ f_initialer + "(id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                         "f_navn varchar(255) DEFAULT'" + f_navn + "'," +
                         "f_initialer varchar(255) DEFAULT'" + f_initialer + "'," +
                         "f_Adresse varchar(255) DEFAULT'" + f_Adresse + "'," +
                         "f_by varchar(255) DEFAULT'" + f_by + "'," +
                         "f_postnr varchar(255) DEFAULT'" + f_postnr + "'," +
                         "f_telefonnummer varchar(255) DEFAULT'" + f_telefonnummer + "'," +
                         "f_EAN_nummer varchar(255) DEFAULT'" + f_EAN_nummer + "'," +
                         "f_CVR_nummer varchar(255) DEFAULT'" + f_CVR_nummer + "'," +
                         "Kontakt_person varchar(255) DEFAULT'" + Kontakt_person + "'," +
                         "kp_Email varchar(255) DEFAULT'" + kp_Email + "'," +
                         "kp_telefonnummer varchar(255) DEFAULT'" + kp_telefonnummer + "');";
                command.ExecuteNonQueryAsync();

                command.Connection.Close();
            }
        }

    }

    // Denne motde sender dig tilbage til loginsiden.
    protected void Redirect_til_home(object sender, EventArgs e)
    {

        Response.Redirect("https://timeout.cloud/timeout_xp/wwwroot/ver2_14/login.asp");

    }

    // Denne metode opretter en ny log ud fra den nye LTO
    protected void Opret_ny_log()
    {
        string LTOname = Request.Form["CompanyInitials"];



        string path = @"D:\webserver\wwwroot\timeout_xp\wwwroot\ver4_22\inc\log\logfile_timeout_" + LTOname + ".txt";

        if (!File.Exists(path))
        {
            // Create a file to write to.
            using (StreamWriter sw = File.CreateText(path))
            {
                sw.WriteLine("   ");

            }

        }

    }

    // Denne metode opretter en ny folder ud fra den nye LTO
    protected void Opret_ny_upload_folder()
    {

        string LTOname = Request.Form["CompanyInitials"];

        string path = @"D:\webserver\wwwroot\timeout_xp\wwwroot\ver4_22\inc\upload\" + LTOname + "";
        Directory.CreateDirectory(path);

    }


    // Denne metode opretter en ny .asp fil ud fra den nye LTO
    protected void Opret_ny_default_ASP()
    {
        string LTOname = Request.Form["CompanyInitials"];

        string path = @"C:\Users\tinom\source\repos\Expence365_3.1\Expence365_3.1\" + LTOname + ".asp";

        if (!File.Exists(path))
        {
            // Create a file to write to.
            using (StreamWriter sw = File.CreateText(path))
            {
                sw.WriteLine("    ");
                sw.WriteLine("<%");
                sw.WriteLine("'** Lokal **'");
                sw.WriteLine("'** Brug kun i specel tilfæle eller ved indstallation på egen server **'");
                sw.WriteLine("'response-redirect ''../timeout_xp/wwwroot/ver2_1/login.aps?key=2.2009-1201-TO100&Lto=cst'' ");
                sw.WriteLine(" ");
                sw.WriteLine("'** Brug denne på produktions server **'");
                sw.WriteLine("response.redirect \"https://timeout.cloud/timeout_xp/wwwroot/ver2_14/login.asp?key=9K2016-0410-TO171&Lto=" + LTOname + "\"");
                sw.WriteLine(" ");
                sw.WriteLine("'** Version lukket ned ***'");
                sw.WriteLine("'response.redirect \"https://outxource.dk/timeout_xp/wwwroot/ver2_1/tak_login.ask" + "\"");
                sw.WriteLine("%>");
            }

        }

    }


    // Denne metode opretter en hel ny database som skal bruges af den nye kunde. Alle tables som databasen skal indeholde bliver også oprettet med DEFAULT VALUE.
    protected void Opret_ny_database_for_bruger()
    {

        string databaseName = Request.Form["CompanyInitials"];

        // Connection to the server
        string connStrServer = "server=194.150.108.154;user=to_tino;port=3306;password=2HotPotatoes;";
        MySqlConnection connToServer = new MySqlConnection(connStrServer);
        MySqlCommand cmd;
        string databaseCreation;

        // Creation of the new database for buyer.
        connToServer.Open();
        databaseCreation = "CREATE DATABASE IF NOT EXISTS `timeout_" + databaseName + "`;";
        MySqlCommand mySqlCommand = new MySqlCommand(databaseCreation, connToServer);
        cmd = mySqlCommand;
        cmd.ExecuteNonQuery();
        connToServer.Close();


        // Connection to the Database.
        var builder_for_new_database = new MySqlConnectionStringBuilder
        {
            Server = "194.150.108.154",
            Database = "timeout_"+databaseName,
            UserID = "to_tino",
            Password = "2HotPotatoes",

        };

        using (var connToDatabase = new MySqlConnection(builder_for_new_database.ConnectionString))
        {

            // Creation of Tables.
            using (var command = connToDatabase.CreateCommand())
            {
                command.Connection.Open();

                command.CommandText = "CREATE TABLE abonner_file_email (afe_id INT(11) NOT NULL AUTO_INCREMENT PRIMARY KEY, " +
                    "afe_file varchar(250) DEFAULT NULL, " +
                    "afe_email varchar(250) DEFAULT NULL, " +
                    "afe_sent int(11) NOT NULL DEFAULT'0'," +
                    "afe_date date NOT NULL DEFAULT '2010-01-01'," +
                    "afe_progrp int(11) NOT NULL DEFAULT'0'," +
                    "afe_init varchar(50) NOT NULL DEFAULT'');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE aft_enh_fordeling (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY, " +
                    "aft_id int(11) NOT NULL DEFAULT'0'," +
                    "fordeldato date NOT NULL DEFAULT '0000-00-00'," +
                    "aar int(11) NOT NULL DEFAULT'0'," +
                    "maned int(11) NOT NULL DEFAULT'0'," +
                    "dag int(11) NOT NULL DEFAULT'0'," +
                    "enheder double(12,2) NOT NULL DEFAULT '0.00'," +
                    "editor varchar(255) NOT NULL DEFAULT''," +
                    "dato date NOT NULL DEFAULT '0000-00-00')";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE akt_bookings (ab_id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY, " +
                    "ab_name int(11) NOT NULL DEFAULT '0'," +
                    "ab_date date NOT NULL DEFAULT '2010-01-01'," +
                    "ab_startdate datetime DEFAULT NULL," +
                    "ab_enddate datetime DEFAULT NULL," +
                    "ab_medid int(11) NOT NULL DEFAULT '0'," +
                    "ab_aktid int(11) NOT NULL DEFAULT '0'," +
                    "ab_jobid int(11) NOT NULL DEFAULT '0'," +
                    "ab_serie int(11) NOT NULL DEFAULT '0'," +
                    "ab_editor varchar(50) DEFAULT NULL," +
                    "ab_editor_date date NOT NULL DEFAULT '2010-01-01'," +
                    "ab_important int(11) NOT NULL DEFAULT '0'," +
                    "ab_end_after int(11) NOT NULL DEFAULT '0'," +
                    "ab_heading int(11) DEFAULT '0'," +
                    "ab_color int(11) DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE akt_bookings_rel (abl_id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "abl_bookid int(11) NOT NULL DEFAULT '0'," +
                    "abl_medid int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE akt_gruppe (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "navn varchar(50) DEFAULT NULL," +
                    "editor varchar(50) DEFAULT NULL," +
                    "dato varchar(50) DEFAULT NULL," +
                    "forvalgt int(11) NOT NULL DEFAULT '0'," +
                    "skabelontype int(11) NOT NULL DEFAULT '0'," +
                    "aktgrp_account int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE akt_typer (aty_id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "aty_label varchar(100) DEFAULT NULL," +
                    "aty_desc varchar(100) DEFAULT NULL," +
                    "aty_on int(11) NOT NULL DEFAULT '0'," +
                    "aty_on_realhours int(11) NOT NULL DEFAULT '0'," +
                    "aty_on_invoiceble int(11) NOT NULL DEFAULT '0'," +
                    "aty_on_invoice int(11) NOT NULL DEFAULT '0'," +
                    "aty_on_invoice_chk int(11) NOT NULL DEFAULT '0'," +
                    "aty_on_workhours int(11) NOT NULL DEFAULT '0'," +
                    "aty_pre double(12,2) NOT NULL DEFAULT '0.00'," +
                    "aty_sort double(4,2) NOT NULL DEFAULT '0.00'," +
                    "aty_on_recon int(11) NOT NULL DEFAULT '0'," +
                    "aty_enh int(11) NOT NULL DEFAULT '0'," +
                    "aty_on_adhoc int(11) NOT NULL DEFAULT '0'," +
                    "aty_hide_on_treg int(11) NOT NULL DEFAULT '0'," +
                    "aty_pre_dg varchar(50) NOT NULL DEFAULT ''," +
                    "aty_pre_prg varchar(50) NOT NULL DEFAULT '10'," +
                    "aty_on_calender int(11) NOT NULL DEFAULT '1'," +
                    "aty_mobile_btn int(11) DEFAULT '0'," +
                    "aty_mobile_btn_order int(11) DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE aktionsrelationer (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "aktionsid int(11) DEFAULT NULL," +
                    "medarbid int(11) DEFAULT NULL," +
                    "KEY akt_mad_id_inx (aktionsid,medarbid));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE aktiviteter (id int(11) NOT NULL AUTO_INCREMENT," +
                    "navn varchar(100) DEFAULT NULL," +
                    "beskrivelse longtext," +
                    "dato varchar(50) DEFAULT NULL," +
                    "editor varchar(50) DEFAULT NULL," +
                    "job int(11) DEFAULT '0'," +
                    "fakturerbar int(11) DEFAULT '0'," +
                    "projektgruppe1 int(11) DEFAULT '0'," +
                    "projektgruppe2 int(11) DEFAULT '0'," +
                    "projektgruppe3 int(11) DEFAULT '0'," +
                    "projektgruppe4 int(11) DEFAULT '0'," +
                    "projektgruppe5 int(11) DEFAULT '0'," +
                    "aktstartdato date DEFAULT NULL," +
                    "aktslutdato date DEFAULT NULL," +
                    "budgettimer double(11, 2) DEFAULT '0.00'," +
                    "aktnr int(11) DEFAULT NULL," +
                    "aktpris double DEFAULT '0'," +
                    "aktfavorit int(11) DEFAULT '0'," +
                    "aktstatus int(11) DEFAULT '1'," +
                    "projektgruppe6 int(10) DEFAULT '1'," +
                    "projektgruppe7 int(10) DEFAULT '1'," +
                    "projektgruppe8 int(10) DEFAULT '1'," +
                    "projektgruppe9 int(10) DEFAULT '1'," +
                    "projektgruppe10 int(10) DEFAULT '1'," +
                    "fomr int(11) NOT NULL DEFAULT '0'," +
                    "faktor double(10, 2) NOT NULL DEFAULT '0.00'," +
                    "aktbudget double(10, 2) NOT NULL DEFAULT '0.00'," +
                    "tidslaas int(11) NOT NULL DEFAULT '0'," +
                    "tidslaas_st time DEFAULT NULL," +
                    "tidslaas_sl time DEFAULT NULL," +
                    "incidentid int(11) NOT NULL DEFAULT '0'," +
                    "antalstk double NOT NULL DEFAULT '0'," +
                    "tidslaas_man int(11) NOT NULL DEFAULT '0'," +
                    "tidslaas_tir int(11) NOT NULL DEFAULT '0'," +
                    "tidslaas_ons int(11) NOT NULL DEFAULT '0'," +
                    "tidslaas_tor int(11) NOT NULL DEFAULT '0'," +
                    "tidslaas_fre int(11) NOT NULL DEFAULT '0'," +
                    "tidslaas_lor int(11) NOT NULL DEFAULT '0'," +
                    "tidslaas_son int(11) NOT NULL DEFAULT '0'," +
                    "fase varchar(150) DEFAULT NULL," +
                    "sortorder int(11) NOT NULL DEFAULT '1000'," +
                    "bgr int(11) NOT NULL DEFAULT '0'," +
                    "aktbudgetsum double(10, 2) NOT NULL DEFAULT '0.00'," +
                    "easyreg int(11) NOT NULL DEFAULT '0'," +
                    "fravalgt int(11) NOT NULL DEFAULT '0'," +
                    "brug_fasttp int(11) NOT NULL DEFAULT '0'," +
                    "brug_fastkp int(11) NOT NULL DEFAULT '0'," +
                    "fasttp double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "fasttp_val int(11) NOT NULL DEFAULT '0'," +
                    "fastkp double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "fastkp_val int(11) NOT NULL DEFAULT '0'," +
                    "avarenr varchar(250) DEFAULT NULL," +
                    "kostpristarif varchar(50) NOT NULL DEFAULT '0'," +
                    "aktkonto int(11) NOT NULL DEFAULT '0'," +
                    "extsysid double(12, 0) NOT NULL DEFAULT '0'," +
                    "aktsttid time NOT NULL DEFAULT '00:00:00'," +
                    "aktsltid time NOT NULL DEFAULT '00:00:00'," +
                    "pl_important_act int(11) NOT NULL DEFAULT '0'," +
                    "pl_header int(11) NOT NULL DEFAULT '0'," +
                    "pl_reccour int(11) NOT NULL DEFAULT '0'," +
                    "pl_employee int(11) NOT NULL DEFAULT '0'," +
                    "easyreg_max double(12, 2) DEFAULT '0.00'," +
                    "easyreg_timer_proc int(11) DEFAULT '0'," +
                    "UNIQUE KEY akt_id_inx (id)," +
                    "KEY aktiviteterJ_inx (job)," +
                    "KEY inx_aktgrp (aktfavorit)," +
                    "KEY easyreg (easyreg)," +
                    "KEY inx_a_fakturerbar (fakturerbar));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE brugergrupper (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "rettigheder int(11) DEFAULT '0'," +
                    "navn varchar(50) DEFAULT NULL," +
                    "dato date DEFAULT NULL," +
                    "editor varchar(50) DEFAULT NULL);";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE budget_bruttonetto_dt (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "navn varchar(255) NOT NULL DEFAULT ''," +
                    "budget_aar int(11) NOT NULL DEFAULT '0'," +
                    "dato timestamp(6) NULL DEFAULT NULL," +
                    "editor varchar(255) NOT NULL DEFAULT ''," +
                    "bd_1 int(11) NOT NULL DEFAULT '0'," +
                    "bd_2 int(11) NOT NULL DEFAULT '0'," +
                    "bd_3 int(11) NOT NULL DEFAULT '0'," +
                    "bd_4 int(11) NOT NULL DEFAULT '0'," +
                    "bd_5 int(11) NOT NULL DEFAULT '0'," +
                    "bd_6 int(11) NOT NULL DEFAULT '0'," +
                    "bd_7 int(11) NOT NULL DEFAULT '0'," +
                    "bd_8 int(11) NOT NULL DEFAULT '0'," +
                    "bd_9 int(11) NOT NULL DEFAULT '0'," +
                    "bd_10 int(11) NOT NULL DEFAULT '0'," +
                    "bd_11 int(11) NOT NULL DEFAULT '0'," +
                    "bd_12 int(11) NOT NULL DEFAULT '0'," +
                    "frd_ferie_1 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_ferie_2 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_ferie_3 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_ferie_4 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_ferie_5 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_ferie_6 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_ferie_7 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_ferie_8 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_ferie_9 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_ferie_10 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_ferie_11 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_ferie_12 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_syg_1 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_syg_2 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_syg_3 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_syg_4 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_syg_5 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_syg_6 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_syg_7 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_syg_8 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_syg_9 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_syg_10 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_syg_11 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_syg_12 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_andet_1 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_andet_2 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_andet_3 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_andet_4 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_andet_5 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_andet_6 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_andet_7 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_andet_8 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_andet_9 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_andet_10 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_andet_11 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "frd_andet_12 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "brutto_1 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "brutto_2 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "brutto_3 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "brutto_4 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "brutto_5 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "brutto_6 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "brutto_7 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "brutto_8 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "brutto_9 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "brutto_10 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "brutto_11 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "brutto_12 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "netto_1 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "netto_2 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "netto_3 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "netto_4 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "netto_5 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "netto_6 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "netto_7 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "netto_8 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "netto_9 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "netto_10 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "netto_11 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "netto_12 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "timerprdag double(12, 2) NOT NULL DEFAULT '0.00');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE budget_job (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "dato date DEFAULT NULL," +
                    "editor varchar(100) DEFAULT NULL," +
                    "budgetnavn varchar(100) DEFAULT NULL," +
                    "jobid int(11) NOT NULL DEFAULT '0'," +
                    "refno varchar(100) DEFAULT NULL," +
                    "prodno varchar(100) DEFAULT NULL," +
                    "budget_tot double(12, 2) DEFAULT NULL," +
                    "budget_extra_fo double(12, 2) DEFAULT NULL," +
                    "valuta int(11) DEFAULT NULL," +
                    "kurs double(12, 2) DEFAULT NULL," +
                    "budget_view int(11) NOT NULL DEFAULT '0'," +
                    "rapport_view int(11) NOT NULL DEFAULT '0'," +
                    "stdato date DEFAULT NULL," +
                    "sldato date DEFAULT NULL," +
                    "budgetstatus int(11) DEFAULT NULL," +
                    "b_valuta int(11) NOT NULL DEFAULT '0'," +
                    "b_valuta_kurs double(12, 2) NOT NULL DEFAULT '0.00');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE budget_job_exp (id int(11) NOT NULL AUTO_INCREMENT  PRIMARY KEY," +
                    "budgetid int(11) DEFAULT NULL," +
                    "periodeid int(11) DEFAULT NULL," +
                    "fase varchar(250) DEFAULT NULL," +
                    "aktid int(11) DEFAULT NULL," +
                    "aktnavn varchar(250) DEFAULT NULL," +
                    "konto varchar(250) DEFAULT NULL," +
                    "budget_app double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "valuta int(11) DEFAULT NULL," +
                    "kurs double(12, 2) DEFAULT NULL," +
                    "perdato date DEFAULT NULL," +
                    "periode_no int(11) DEFAULT NULL);";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE budget_job_exp_rel (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "budgetid int(11) DEFAULT NULL," +
                    "periodeid int(11) DEFAULT NULL," +
                    "aktid int(11) DEFAULT NULL," +
                    "belob double(12, 2) DEFAULT NULL," +
                    "dato date DEFAULT NULL);";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE budget_job_per (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "budgetid int(11) DEFAULT NULL," +
                    "beskrivelse varchar(250) DEFAULT NULL," +
                    "budget_app double(12, 2) DEFAULT NULL," +
                    "valuta int(11) DEFAULT NULL," +
                    "kurs double DEFAULT NULL," +
                    "stdato date DEFAULT NULL," +
                    "sldato date DEFAULT NULL," +
                    "currentperiod int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE budget_kunder (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "kundeid int(11) NOT NULL DEFAULT '0'," +
                    "salesgoal double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "date_year int(11) NOT NULL DEFAULT '2002');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE budget_medarb (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "navn varchar(255) NOT NULL DEFAULT ''," +
                    "budget_aar int(11) NOT NULL DEFAULT '0'," +
                    "dato timestamp(6) NULL DEFAULT NULL," +
                    "editor varchar(255) NOT NULL DEFAULT ''," +
                    "budget_brutto_id int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE budget_medarb_rel (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "dato timestamp(6) NULL DEFAULT NULL," +
                    "editor varchar(255) NOT NULL DEFAULT ''," +
                    "budget_id int(11) NOT NULL DEFAULT '0'," +
                    "medid int(11) NOT NULL DEFAULT '0'," +
                    "jan double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "feb double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "mar double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "apr double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "maj double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "jun double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "jul double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "aug double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "sep double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "okt double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "nov double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "des double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "timepris double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "ntimerpruge double(12, 2) NOT NULL DEFAULT '0.00');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE chat (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "jobid int(11) DEFAULT '0'," +
                    "message varchar(255) DEFAULT ''," +
                    "editor int(11) DEFAULT '0'," +
                    "editdate date NOT NULL DEFAULT '2002-01-01'," +
                    "edittime time DEFAULT '00:00:00');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE crmemne (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "editor varchar(50) DEFAULT NULL," +
                    "dato varchar(50) DEFAULT NULL," +
                    "navn varchar(50) DEFAULT NULL);";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE crmhistorik (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "editor varchar(50) DEFAULT '07.00.00'," +
                    "dato varchar(50) DEFAULT NULL," +
                    "crmdato date DEFAULT '0000-00-00'," +
                    "kontaktemne int(2) DEFAULT '0'," +
                    "kontaktpers varchar(50) DEFAULT '0'," +
                    "status int(11) DEFAULT '0'," +
                    "komm text," +
                    "navn varchar(50) DEFAULT NULL," +
                    "kundeid int(11) DEFAULT '0'," +
                    "kontaktform int(11) DEFAULT '0'," +
                    "editorid int(11) DEFAULT NULL," +
                    "crmklokkeslet time DEFAULT '08:00:00'," +
                    "crmklokkeslet_slut time DEFAULT '09:00:00'," +
                    "crmdato_slut date DEFAULT NULL," +
                    "serialnb int(12) DEFAULT '0'," +
                    "KEY inx_kundeid (kundeid));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE crmkontaktform (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "editor varchar(50) DEFAULT NULL," +
                    "dato varchar(50) DEFAULT NULL," +
                    "navn varchar(50) DEFAULT NULL," +
                    "ikon varchar(50) DEFAULT NULL);";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE crmstatus (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "editor varchar(50) DEFAULT NULL," +
                    "dato datetime NOT NULL DEFAULT '0000-00-00 00:00:00'," +
                    "navn varchar(50) DEFAULT NULL);";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE dbdownload (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "dato date DEFAULT NULL," +
                    "editor varchar(50) DEFAULT NULL," +
                    "email varchar(50) NOT NULL DEFAULT '');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE dbversion (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "dbversion double DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE delete_hist (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "deltype varchar(255) NOT NULL DEFAULT ''," +
                    "delid int(11) NOT NULL DEFAULT '0'," +
                    "delnr varchar(255) NOT NULL DEFAULT ''," +
                    "delnavn varchar(255) NOT NULL DEFAULT ''," +
                    "mid int(11) NOT NULL DEFAULT '0'," +
                    "mnavn varchar(255) NOT NULL DEFAULT ''," +
                    "dato timestamp(6) NULL DEFAULT NULL);";
                command.ExecuteNonQueryAsync();

                // to varchar kan ikke hedde det. (mail_to) hedder den pt.
                command.CommandText = "CREATE TABLE emails (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "kpersid int(11) DEFAULT NULL," +
                    "kundeid int(11) DEFAULT '0'," +
                    "domæne varchar(50) DEFAULT NULL," +
                    "dato timestamp(6) NULL DEFAULT NULL," +
                    "mail_to varchar(50) DEFAULT NULL," +
                    "cc varchar(50) DEFAULT NULL," +
                    "bcc varchar(50) DEFAULT NULL," +
                    "subject varchar(50) DEFAULT NULL," +
                    "msgtext text," +
                    "laest int(11) NOT NULL DEFAULT '0'," +
                    "laest_dato timestamp(6) NULL DEFAULT NULL," +
                    "att varchar(50) DEFAULT NULL," +
                    "KEY kids (kpersid,kundeid));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE emails_attment (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "attment varchar(50) DEFAULT NULL," +
                    "emailid int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE emails_laestaf (id int(11) NOT NULL AUTO_INCREMENT  PRIMARY KEY," +
                    "medarbejderid int(11) NOT NULL DEFAULT '0'," +
                    "emailid int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE enheds_typer (et_id int(11) NOT NULL DEFAULT '0' PRIMARY KEY," +
                    "et_navn varchar(100) NOT NULL DEFAULT '');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE eval (eval_id double NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "eval_jobid double NOT NULL DEFAULT '0'," +
                    "eval_evalvalue int(11) NOT NULL DEFAULT '0'," +
                    "eval_jobvaluesuggested double NOT NULL DEFAULT '0'," +
                    "eval_comment varchar(255) DEFAULT NULL," +
                    "eval_diff double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "eval_suggested_hours double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "eval_suggested_hourly_rate double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "eval_original_price double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "eval_fakbartimer double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "eval_fakbartimepris double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "ubemandet_maskine_timer double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "ubemandet_maskine_timePris double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "laer_timer double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "laer_timepris double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "easy_reg_timepris double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "ikke_fakbar_tid_timer double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "ikke_fakbar_tid_timepris double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "easy_reg_timer double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "KEY inx_eval_id2 (eval_id)," +
                    "KEY inx_eval_jobid2 (eval_jobid));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE fak_mat_spec (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "matfakid int(11) NOT NULL DEFAULT '0'," +
                    "matid int(11) NOT NULL DEFAULT '0'," +
                    "matantal double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "matnavn text," +
                    "matvarenr varchar(255) DEFAULT NULL," +
                    "matenhed varchar(255) DEFAULT NULL," +
                    "matenhedspris double(12, 3) NOT NULL DEFAULT '0.000'," +
                    "matrabat double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "matbeloeb double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "matshowonfak int(11) NOT NULL DEFAULT '0'," +
                    "ikkemoms int(11) NOT NULL DEFAULT '0'," +
                    "valuta int(11) NOT NULL DEFAULT '1'," +
                    "kurs double(12, 2) NOT NULL DEFAULT '100.00'," +
                    "matfrb_mid int(11) NOT NULL DEFAULT '0'," +
                    "matfrb_id double(12, 0) NOT NULL DEFAULT '0'," +
                    "matgrp int(11) NOT NULL DEFAULT '0'," +
                    "fms_aktid int(11) NOT NULL DEFAULT '0'," +
                    "KEY inx_fakid_mid (matfakid)," +
                    "KEY inx_akt_mid (matid)," +
                    "KEY inx_akt_mid_fak (matfakid,matid));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE fak_med_spec (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "fakid int(11) NOT NULL DEFAULT '0'," +
                    "aktid int(11) NOT NULL DEFAULT '0'," +
                    "mid int(11) NOT NULL DEFAULT '0'," +
                    "fak double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "venter double(12, 2) NOT NULL DEFAULT '2.00'," +
                    "ub double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "tekst text," +
                    "enhedspris double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "beloeb double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "showonfak int(11) NOT NULL DEFAULT '0'," +
                    "medrabat double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "venter_brugt double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "enhedsang int(11) NOT NULL DEFAULT '0'," +
                    "valuta int(11) NOT NULL DEFAULT '1'," +
                    "kurs double(12, 2) NOT NULL DEFAULT '100.00'," +
                    "KEY inx_fakid_mid (fakid,mid)," +
                    "KEY inx_akt_mid (aktid,mid)," +
                    "KEY inx_akt_mid_fak (fakid,aktid,mid));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE fak_moms (id int(11) NOT NULL AUTO_INCREMENT  PRIMARY KEY," +
                    "moms int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE fak_opr_faknr (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "faknr varchar(50) DEFAULT NULL," +
                    "sesid int(11) DEFAULT NULL);";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE fak_sprog (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "navn varchar(255) NOT NULL DEFAULT '');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE faktura_det (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "antal double(11, 2) DEFAULT NULL," +
                    "beskrivelse text," +
                    "aktpris double(11, 2) DEFAULT NULL," +
                    "fakid int(11) DEFAULT '0'," +
                    "enhedspris double(12, 2) DEFAULT '0.00'," +
                    "aktid int(11) NOT NULL DEFAULT '0'," +
                    "showonfak int(10) DEFAULT '0'," +
                    "rabat double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "enhedsang int(11) NOT NULL DEFAULT '0'," +
                    "valuta int(11) NOT NULL DEFAULT '1'," +
                    "kurs double(12, 2) NOT NULL DEFAULT '100.00'," +
                    "fak_sortorder double(12, 2) NOT NULL DEFAULT '1.00'," +
                    "momsfri int(11) NOT NULL DEFAULT '0'," +
                    "fase varchar(255) DEFAULT NULL," +
                    "KEY inx_fakid (fakid));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE faktura_rykker (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "rykkerantal double(11, 2) DEFAULT NULL," +
                    "rykkertxt text," +
                    "rykkerbelob double(11, 2) DEFAULT NULL," +
                    "fakid int(11) NOT NULL DEFAULT '0'," +
                    "rykkerdato date NOT NULL DEFAULT '0000-00-00'," +
                    "editor varchar(255) NOT NULL DEFAULT ''," +
                    "dato date NOT NULL DEFAULT '0000-00-00'," +
                    "valuta int(11) NOT NULL DEFAULT '1'," +
                    "kurs double(12, 2) NOT NULL DEFAULT '100.00');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE fakturaer (Fid int(11) NOT NULL AUTO_INCREMENT," +
                    "dato varchar(50) DEFAULT NULL," +
                    "editor varchar(50) DEFAULT NULL," +
                    "faknr varchar(11) NOT NULL DEFAULT '0'," +
                    "fakdato date DEFAULT NULL," +
                    "jobid int(11) DEFAULT '0'," +
                    "timer double(11, 2) DEFAULT '0.00'," +
                    "beloeb double DEFAULT '0'," +
                    "kommentar longtext," +
                    "tidspunkt time DEFAULT NULL," +
                    "betalt int(11) DEFAULT '0'," +
                    "b_dato date DEFAULT NULL," +
                    "fakadr int(11) DEFAULT '0'," +
                    "att varchar(50) DEFAULT NULL," +
                    "faktype tinyint(3) NOT NULL DEFAULT '0'," +
                    "konto double(12, 0) DEFAULT NULL," +
                    "modkonto double(12, 2) DEFAULT NULL," +
                    "oprid int(11) DEFAULT '0'," +
                    "parentfak int(11) NOT NULL DEFAULT '0'," +
                    "moms double(10, 2) NOT NULL DEFAULT '0.00'," +
                    "vismodland int(11) NOT NULL DEFAULT '0'," +
                    "vismodatt int(11) NOT NULL DEFAULT '0'," +
                    "vismodtlf int(11) NOT NULL DEFAULT '0'," +
                    "vismodcvr int(11) NOT NULL DEFAULT '0'," +
                    "visafstlf int(11) NOT NULL DEFAULT '0'," +
                    "visafsemail int(11) NOT NULL DEFAULT '0'," +
                    "visafsswift int(11) NOT NULL DEFAULT '0'," +
                    "visafsiban int(11) NOT NULL DEFAULT '0'," +
                    "visafscvr int(11) NOT NULL DEFAULT '0'," +
                    "enhedsang int(11) NOT NULL DEFAULT '0'," +
                    "varenr varchar(50) NOT NULL DEFAULT '0'," +
                    "aftaleid int(11) NOT NULL DEFAULT '0'," +
                    "jobbesk text," +
                    "timersubtotal double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "visjoblog int(11) NOT NULL DEFAULT '0'," +
                    "visrabatkol int(11) NOT NULL DEFAULT '0'," +
                    "vismatlog int(11) NOT NULL DEFAULT '0'," +
                    "shadowcopy int(11) NOT NULL DEFAULT '0'," +
                    "rabat double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "visjoblog_timepris int(11) NOT NULL DEFAULT '0'," +
                    "visjoblog_enheder int(11) NOT NULL DEFAULT '0'," +
                    "visafsfax int(11) NOT NULL DEFAULT '0'," +
                    "erfakbetalt int(11) NOT NULL DEFAULT '0'," +
                    "subtotaltilmoms double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "valuta int(11) NOT NULL DEFAULT '1'," +
                    "kurs double(12, 2) NOT NULL DEFAULT '100.00'," +
                    "sprog int(11) NOT NULL DEFAULT '1'," +
                    "istdato date NOT NULL DEFAULT '2002-01-01'," +
                    "momskonto int(11) NOT NULL DEFAULT '1'," +
                    "visperiode int(11) NOT NULL DEFAULT '0'," +
                    "visjoblog_mnavn int(11) NOT NULL DEFAULT '1'," +
                    "betbetint int(11) NOT NULL DEFAULT '0'," +
                    "jobfaktype int(11) NOT NULL DEFAULT '0'," +
                    "istdato2 date NOT NULL DEFAULT '2002-01-01'," +
                    "brugfakdatolabel int(11) NOT NULL DEFAULT '0'," +
                    "fakbetkom varchar(250) DEFAULT NULL," +
                    "momssats int(11) NOT NULL DEFAULT '0'," +
                    "modtageradr text," +
                    "usealtadr int(11) NOT NULL DEFAULT '0'," +
                    "vorref varchar(150) DEFAULT NULL," +
                    "showmatasgrp int(11) NOT NULL DEFAULT '0'," +
                    "fak_ski int(11) NOT NULL DEFAULT '0'," +
                    "hidesumaktlinier int(11) NOT NULL DEFAULT '0'," +
                    "sideskiftlinier int(11) NOT NULL DEFAULT '0'," +
                    "labeldato date DEFAULT NULL," +
                    "fak_abo int(11) NOT NULL DEFAULT '0'," +
                    "fak_ubv int(11) NOT NULL DEFAULT '0'," +
                    "visikkejobnavn int(11) NOT NULL DEFAULT '0'," +
                    "hidefasesum int(11) NOT NULL DEFAULT '0'," +
                    "hideantenh int(11) NOT NULL DEFAULT '0'," +
                    "medregnikkeioms int(11) NOT NULL DEFAULT '0'," +
                    "fak_laast int(11) NOT NULL DEFAULT '0'," +
                    "afsender int(11) NOT NULL DEFAULT '0'," +
                    "overfort_erp int(11) NOT NULL DEFAULT '0'," +
                    "vis_jobbesk int(11) NOT NULL DEFAULT '0'," +
                    "kontonr_sel int(11) NOT NULL DEFAULT '0'," +
                    "fakglobalfaktor double(12, 2) NOT NULL DEFAULT '1.00'," +
                    "totbel_afvige int(11) NOT NULL DEFAULT '0'," +
                    "fak_fomr int(11) NOT NULL DEFAULT '0'," +
                    "fak_rekvinr varchar(255) NOT NULL DEFAULT ''," +
                    "PRIMARY KEY (Fid,faknr)," +
                    "UNIQUE KEY Fid_inx (Fid)," +
                    "KEY Fdat_inx (fakdato)," +
                    "KEY Fakjid_inx (jobid)," +
                    "KEY fnr (faknr)," +
                    "KEY job_inx (jobid)," +
                    "KEY job_inx2 (jobid)," +
                    "KEY job_inx3 (jobid)," +
                    "KEY fakja_inx (jobid,aftaleid));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE filer (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "filnavn varchar(250) DEFAULT NULL," +
                    "editor varchar(50) DEFAULT NULL," +
                    "dato date NOT NULL DEFAULT '0000-00-00'," +
                    "type varchar(50) DEFAULT NULL," +
                    "jobid int(11) DEFAULT NULL," +
                    "kundeid int(11) DEFAULT NULL," +
                    "oprses varchar(50) DEFAULT NULL," +
                    "folderid int(11) NOT NULL DEFAULT '0'," +
                    "adg_kunde int(11) NOT NULL DEFAULT '0'," +
                    "adg_admin int(11) NOT NULL DEFAULT '1'," +
                    "adg_alle int(11) NOT NULL DEFAULT '0'," +
                    "incidentid int(11) NOT NULL DEFAULT '0'," +
                    "filertxt longtext," +
                    "KEY inx_fi_j (jobid)," +
                    "KEY inx_fi_j2 (jobid)," +
                    "KEY inx_fi_j3 (jobid)," +
                    "KEY inx_fi_k2 (kundeid)," +
                    "KEY inx_fi_n2 (filnavn));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE folder_grupper (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "navn varchar(50) DEFAULT NULL," +
                    "editor varchar(50) DEFAULT NULL," +
                    "dato varchar(50) DEFAULT NULL);";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE foldere (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "navn varchar(255) DEFAULT NULL," +
                    "kundese int(11) NOT NULL DEFAULT '0'," +
                    "kundeid int(11) NOT NULL DEFAULT '0'," +
                    "jobid int(11) NOT NULL DEFAULT '0'," +
                    "editor varchar(255) DEFAULT NULL," +
                    "dato date DEFAULT NULL," +
                    "stfoldergruppe int(11) NOT NULL DEFAULT '0'," +
                    "KEY inx_fo_k2 (kundeid)," +
                    "KEY inx_fo_j2 (jobid)," +
                    "KEY inx_fo_n2 (navn));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE fomr (id int(11) NOT NULL AUTO_INCREMENT  PRIMARY KEY," +
                    "editor varchar(50) DEFAULT NULL," +
                    "dato varchar(50) DEFAULT NULL," +
                    "navn varchar(50) DEFAULT NULL," +
                    "konto int(11) NOT NULL DEFAULT '0'," +
                    "jobok int(11) NOT NULL DEFAULT '1'," +
                    "aktok int(11) NOT NULL DEFAULT '1'," +
                    "business_unit varchar(250) DEFAULT NULL," +
                    "business_area_label varchar(250) DEFAULT NULL," +
                    "fomr_segment varchar(255) DEFAULT NULL);";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE fomr_rel (for_id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "for_fomr int(11) NOT NULL DEFAULT '0'," +
                    "for_jobid double(12, 0) NOT NULL DEFAULT '0'," +
                    "for_aktid double(12, 0) NOT NULL DEFAULT '0'," +
                    "for_faktor double(12, 2) NOT NULL DEFAULT '1.00'," +
                    "KEY for_aktid (for_aktid)," +
                    "KEY for_fomr (for_fomr)," +
                    "KEY inx_for_jobid (for_jobid));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE forgot_password (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "medid int(11) DEFAULT '0'," +
                    "stampdate date NOT NULL," +
                    "stamptime time NOT NULL," +
                    "password_updated int(11) DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE gantts (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "dato date DEFAULT NULL," +
                    "editor varchar(100) DEFAULT NULL," +
                    "medid int(11) NOT NULL DEFAULT '0'," +
                    "jobnrs text," +
                    "navn varchar(255) DEFAULT NULL," +
                    "datomidtpkt date NOT NULL DEFAULT '0000-00-00');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE incidentnoter (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "note mediumtext," +
                    "editor varchar(255) DEFAULT NULL," +
                    "dato date DEFAULT NULL," +
                    "timerid int(11) NOT NULL DEFAULT '0'," +
                    "todoid int(11) NOT NULL DEFAULT '0'," +
                    "status int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE info_screen (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "overskrift varchar(255) NOT NULL DEFAULT ''," +
                    "brodtext text," +
                    "datofra date NOT NULL DEFAULT '2002-01-01'," +
                    "datotil date NOT NULL DEFAULT '2002-01-01'," +
                    "editor varchar(255) DEFAULT ''," +
                    "type int(11) DEFAULT '2'," +
                    "vigtig int(11) DEFAULT '0'," +
                    "filnavn varchar(255) DEFAULT NULL);";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE infobase (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "navn varchar(50) DEFAULT NULL," +
                    "editor varchar(50) DEFAULT NULL," +
                    "dato datetime DEFAULT NULL," +
                    "parent int(11) NOT NULL DEFAULT '0'," +
                    "kuid int(11) DEFAULT NULL," +
                    "beskrivelse text," +
                    "globalinfobase int(11) DEFAULT '0'," +
                    "KEY navn_index (navn)," +
                    "FULLTEXT KEY besk_index (beskrivelse));";
                command.ExecuteNonQueryAsync();


                command.CommandText = "CREATE TABLE job (id int(11) NOT NULL AUTO_INCREMENT," +
                    "dato varchar(50) DEFAULT NULL," +
                    "jobnavn varchar(150) NOT NULL DEFAULT ''," +
                    "jobnr varchar(50) NOT NULL DEFAULT '0'," +
                    "jobstatus int(11) unsigned DEFAULT '0'," +
                    "jobknr int(11) DEFAULT '0'," +
                    "jobTpris double(10, 2) DEFAULT '0.00'," +
                    "editor varchar(50) DEFAULT NULL," +
                    "jobstartdato date DEFAULT NULL," +
                    "jobslutdato date DEFAULT NULL," +
                    "projektgruppe1 int(10) unsigned DEFAULT '0'," +
                    "projektgruppe2 int(10) unsigned DEFAULT '0'," +
                    "projektgruppe3 int(10) unsigned DEFAULT '0'," +
                    "projektgruppe4 int(10) unsigned DEFAULT '0'," +
                    "projektgruppe5 int(10) unsigned DEFAULT '0'," +
                    "fakturerbart int(10) unsigned zerofill DEFAULT '0000000000'," +
                    "budgettimer double(10, 2) DEFAULT '0.00'," +
                    "fastpris varchar(50) DEFAULT '0'," +
                    "kundeok int(11) DEFAULT '1'," +
                    "beskrivelse longtext," +
                    "ikkebudgettimer double(10, 2) DEFAULT '0.00'," +
                    "tilbudsnr int(12) DEFAULT '0'," +
                    "projektgruppe6 int(10) DEFAULT '1'," +
                    "projektgruppe7 int(10) DEFAULT '1'," +
                    "projektgruppe8 int(10) DEFAULT '1'," +
                    "projektgruppe9 int(10) DEFAULT '1'," +
                    "projektgruppe10 int(10) DEFAULT '1'," +
                    "jobans1 int(11) NOT NULL DEFAULT '0'," +
                    "jobans2 int(11) NOT NULL DEFAULT '0'," +
                    "serviceaft int(11) NOT NULL DEFAULT '0'," +
                    "kundekpers int(11) NOT NULL DEFAULT '0'," +
                    "forventetslut date DEFAULT NULL," +
                    "udgifter double NOT NULL DEFAULT '0'," +
                    "risiko int(11) NOT NULL DEFAULT '0'," +
                    "restestimat double NOT NULL DEFAULT '0'," +
                    "lukafmjob int(11) NOT NULL DEFAULT '0'," +
                    "valuta int(11) NOT NULL DEFAULT '1'," +
                    "jobfaktype int(11) NOT NULL DEFAULT '0'," +
                    "kommentar text," +
                    "rekvnr varchar(50) DEFAULT NULL," +
                    "jo_gnstpris double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "jo_gnsfaktor double(12, 2) NOT NULL DEFAULT '1.00'," +
                    "jo_gnsbelob double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "jo_bruttofortj double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "jo_dbproc double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "usejoborakt_tp int(11) NOT NULL DEFAULT '0'," +
                    "ski int(11) NOT NULL DEFAULT '0'," +
                    "job_internbesk text," +
                    "abo int(11) NOT NULL DEFAULT '0'," +
                    "ubv int(11) NOT NULL DEFAULT '0'," +
                    "stade_tim_proc int(11) NOT NULL DEFAULT '0'," +
                    "sandsynlighed double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "jobans3 int(11) NOT NULL DEFAULT '0'," +
                    "jobans4 int(11) NOT NULL DEFAULT '0'," +
                    "jobans5 int(11) NOT NULL DEFAULT '0'," +
                    "diff_timer double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "diff_sum double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "jo_udgifter_intern double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "jo_udgifter_ulev double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "jo_bruttooms double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "jobans_proc_1 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "jobans_proc_2 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "jobans_proc_3 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "jobans_proc_4 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "jobans_proc_5 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "virksomheds_proc double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "syncslutdato int(11) NOT NULL DEFAULT '0'," +
                    "lukkedato date NOT NULL DEFAULT '2002-01-01'," +
                    "altfakadr int(11) NOT NULL DEFAULT '0'," +
                    "preconditions_met int(11) NOT NULL DEFAULT '0'," +
                    "laasmedtpbudget int(11) DEFAULT '0'," +
                    "salgsans1 int(11) NOT NULL DEFAULT '0'," +
                    "salgsans2 int(11) NOT NULL DEFAULT '0'," +
                    "salgsans3 int(11) NOT NULL DEFAULT '0'," +
                    "salgsans4 int(11) NOT NULL DEFAULT '0'," +
                    "salgsans5 int(11) NOT NULL DEFAULT '0'," +
                    "salgsans1_proc double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "salgsans2_proc double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "salgsans3_proc double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "salgsans4_proc double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "salgsans5_proc double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "kunde_betbetint int(11) NOT NULL DEFAULT '0'," +
                    "kunde_levbetint int(11) NOT NULL DEFAULT '0'," +
                    "lev_betbetint int(11) NOT NULL DEFAULT '0'," +
                    "lev_levbetint int(11) NOT NULL DEFAULT '0'," +
                    "supplier int(11) NOT NULL DEFAULT '0'," +
                    "invoice_no varchar(50) NOT NULL DEFAULT '0'," +
                    "filepath1 varchar(255) DEFAULT NULL," +
                    "filepath2 varchar(255) DEFAULT NULL," +
                    "filepath3 varchar(255) DEFAULT NULL," +
                    "fomr_konto int(11) NOT NULL DEFAULT '0'," +
                    "jfak_sprog int(11) NOT NULL DEFAULT '1'," +
                    "jfak_moms int(11) NOT NULL DEFAULT '1'," +
                    "alert int(11) NOT NULL DEFAULT '0'," +
                    "lincensindehaver_faknr_prioritet_job varchar(255) DEFAULT NULL," +
                    "jo_valuta int(11) NOT NULL DEFAULT '0'," +
                    "jo_valuta_kurs double(12, 2) NOT NULL DEFAULT '100.00'," +
                    "jo_usefybudgetingt int(11) NOT NULL DEFAULT '0'," +
                    "extracost double(12,2) NOT NULL DEFAULT '0.00'," +
                    "extracost_txt varchar(255) DEFAULT NULL," +
                    "categories_process varchar(250) NOT NULL DEFAULT ''," +
                    "data_outside_eu int(11) NOT NULL DEFAULT '0'," +
                    "safeguard varchar(250) NOT NULL DEFAULT ''," +
                    "gdpr_projecttype int(11) NOT NULL DEFAULT '0'," +
                    "gdpr_personaldata int(11) NOT NULL DEFAULT '0'," +
                    "gdpr_safeguard_io int(11) NOT NULL DEFAULT '0'," +
                    "PRIMARY KEY (id,jobnavn(10),jobnr)," +
                    "project_tier varchar(250) DEFAULT ''," +
                    "UNIQUE KEY job_jobnr_inx (jobnr)," +
                    "UNIQUE KEY jobnr_inx (jobnr)," +
                    "UNIQUE KEY jobid_inx (id)," +
                    "KEY jobKnr_inx (jobknr)," +
                    "KEY jobnavn_inx (jobnavn)," +
                    "KEY job_st_inx (jobstatus)," +
                    "KEY job_st_fakb_pg1_inx (jobstatus,fakturerbart,projektgruppe1,projektgruppe2,projektgruppe3,projektgruppe4,projektgruppe5));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE job_igv_status (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "dato date DEFAULT NULL," +
                    "editor varchar(100) DEFAULT NULL," +
                    "medid int(11) NOT NULL DEFAULT '0'," +
                    "aar int(11) NOT NULL DEFAULT '0'," +
                    "maaned int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE job_status (js_id int(11) NOT NULL DEFAULT '0' PRIMARY KEY," +
                    "js_navn varchar(255) NOT NULL DEFAULT '');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE job_ulev_ju (ju_id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "ju_navn varchar(100) NOT NULL DEFAULT ''," +
                    "ju_ipris double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "ju_faktor double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "ju_belob double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "ju_jobid int(11) NOT NULL DEFAULT '0'," +
                    "ju_favorit int(11) NOT NULL DEFAULT '0'," +
                    "ju_fase varchar(200) DEFAULT NULL," +
                    "ju_stk double(12, 2) NOT NULL DEFAULT '1.00'," +
                    "ju_stkpris double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "ju_fravalgt int(11) NOT NULL DEFAULT '0'," +
                    "extsysid double(12, 0) NOT NULL DEFAULT '0'," +
                    "ju_konto int(11) NOT NULL DEFAULT '0'," +
                    "ju_konto_label varchar(50) DEFAULT NULL," +
                    "ju_matid int(11) NOT NULL DEFAULT '0'," +
                    "ju_date date NOT NULL DEFAULT '2002-01-01'," +
                    "ju_editor varchar(250) NOT NULL DEFAULT 'TO_sys');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE job_ulev_jugrp (jugrp_id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "jugrp_dato date DEFAULT NULL," +
                    "jugrp_editor varchar(100) DEFAULT NULL," +
                    "jugrp_navn varchar(200) DEFAULT NULL," +
                    "jugrp_forvalgt int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE kontaktpers (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "kundeid int(11) NOT NULL DEFAULT '0'," +
                    "editor varchar(50) DEFAULT NULL," +
                    "dato date NOT NULL DEFAULT '0000-00-00'," +
                    "navn varchar(150) DEFAULT NULL," +
                    "titel varchar(50) DEFAULT NULL," +
                    "adresse varchar(50) DEFAULT NULL," +
                    "postnr varchar(50) DEFAULT NULL," +
                    "town varchar(50) DEFAULT NULL," +
                    "land varchar(50) DEFAULT NULL," +
                    "afdeling varchar(50) NOT NULL DEFAULT ''," +
                    "email varchar(50) DEFAULT NULL," +
                    "password varchar(50) DEFAULT NULL," +
                    "dirtlf varchar(50) DEFAULT NULL," +
                    "mobiltlf varchar(50) DEFAULT NULL," +
                    "privattlf varchar(50) DEFAULT NULL," +
                    "lastlogin varchar(50) DEFAULT NULL," +
                    "photo varchar(50) DEFAULT NULL," +
                    "beskrivelse text," +
                    "kpean varchar(250) DEFAULT NULL," +
                    "kptype int(11) NOT NULL DEFAULT '0'," +
                    "kpcvr varchar(255) DEFAULT NULL," +
                    "kp_interest_christmas int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE kontoplan (id int(11) NOT NULL AUTO_INCREMENT," +
                    "navn varchar(50) DEFAULT NULL," +
                    "editor varchar(50) DEFAULT NULL," +
                    "dato date NOT NULL DEFAULT '2000-00-00'," +
                    "kid int(11) DEFAULT NULL," +
                    "kontonr double(12, 0) NOT NULL DEFAULT '0'," +
                    "keycode int(11) DEFAULT NULL," +
                    "momskode int(11) NOT NULL DEFAULT '0'," +
                    "debitkredit int(11) NOT NULL DEFAULT '1'," +
                    "status int(11) NOT NULL DEFAULT '1'," +
                    "type int(11) NOT NULL DEFAULT '0'," +
                    "PRIMARY KEY (id,kontonr));";
                command.ExecuteNonQueryAsync();


                command.CommandText = "CREATE TABLE kunder (Kid int(11) NOT NULL AUTO_INCREMENT," +
                    "Kdato varchar(50) DEFAULT NULL," +
                    "kkundenavn varchar(150) NOT NULL DEFAULT ''," +
                    "Kkundenr varchar(20) NOT NULL DEFAULT '0'," +
                    "editor varchar(50) DEFAULT NULL," +
                    "ketype varchar(50) DEFAULT '0'," +
                    "adresse varchar(200) DEFAULT NULL," +
                    "postnr varchar(25) DEFAULT NULL," +
                    "city varchar(50) DEFAULT NULL," +
                    "telefon varchar(20) DEFAULT NULL," +
                    "telefonmobil varchar(12) DEFAULT NULL," +
                    "telefonalt varchar(20) DEFAULT NULL," +
                    "fax varchar(12) DEFAULT NULL," +
                    "email varchar(50) DEFAULT NULL," +
                    "land varchar(50) DEFAULT NULL," +
                    "kontaktpers1 varchar(50) DEFAULT NULL," +
                    "kontaktpers2 varchar(50) DEFAULT NULL," +
                    "kontaktpers3 varchar(50) DEFAULT NULL," +
                    "kontaktpers4 varchar(50) DEFAULT NULL," +
                    "kontaktpers5 varchar(50) DEFAULT NULL," +
                    "kpersemail1 varchar(50) DEFAULT NULL," +
                    "kpersemail2 varchar(50) DEFAULT NULL," +
                    "kpersemail3 varchar(50) DEFAULT NULL," +
                    "kpersemail4 varchar(50) DEFAULT NULL," +
                    "kpersemail5 varchar(50) DEFAULT NULL," +
                    "kperstlf1 varchar(12) DEFAULT NULL," +
                    "kperstlf2 varchar(12) DEFAULT NULL," +
                    "kperstlf3 varchar(12) DEFAULT NULL," +
                    "kperstlf4 varchar(12) DEFAULT NULL," +
                    "kperstlf5 varchar(12) DEFAULT NULL," +
                    "kpersmobil1 varchar(12) DEFAULT NULL," +
                    "kpersmobil2 varchar(12) DEFAULT NULL," +
                    "kpersmobil3 varchar(12) DEFAULT NULL," +
                    "kpersmobil4 varchar(12) DEFAULT NULL," +
                    "kpersmobil5 varchar(12) DEFAULT NULL," +
                    "url varchar(100) DEFAULT NULL," +
                    "beskrivelse text," +
                    "regnr varchar(10) DEFAULT NULL," +
                    "kontonr varchar(50) DEFAULT NULL," +
                    "cvr varchar(25) DEFAULT NULL," +
                    "bank varchar(200) DEFAULT NULL," +
                    "pwkpers1 varchar(20) DEFAULT NULL," +
                    "pwkpers2 varchar(20) DEFAULT NULL," +
                    "pwkpers3 varchar(20) DEFAULT NULL," +
                    "pwkpers4 varchar(20) DEFAULT NULL," +
                    "pwkpers5 varchar(20) DEFAULT NULL," +
                    "lastloginkp1 varchar(50) DEFAULT NULL," +
                    "lastloginkp2 varchar(50) DEFAULT NULL," +
                    "lastloginkp3 varchar(50) DEFAULT NULL," +
                    "lastloginkp4 varchar(50) DEFAULT NULL," +
                    "lastloginkp5 varchar(50) DEFAULT NULL," +
                    "useasfak int(11) DEFAULT '0'," +
                    "logo int(11) DEFAULT NULL," +
                    "hot int(11) DEFAULT '0'," +
                    "swift varchar(8) DEFAULT NULL," +
                    "iban varchar(50) DEFAULT NULL," +
                    "ktype int(10) DEFAULT '0'," +
                    "kundeans1 int(11) NOT NULL DEFAULT '0'," +
                    "kundeans2 int(11) NOT NULL DEFAULT '0'," +
                    "nace varchar(50) NOT NULL DEFAULT '0'," +
                    "levbet text," +
                    "betbet text," +
                    "sdskpriogrp int(11) NOT NULL DEFAULT '0'," +
                    "ean varchar(50) DEFAULT NULL," +
                    "betbetint int(11) NOT NULL DEFAULT '0'," +
                    "regnr_b varchar(10) DEFAULT NULL," +
                    "kontonr_b varchar(50) DEFAULT NULL," +
                    "regnr_c varchar(10) DEFAULT NULL," +
                    "kontonr_c varchar(50) DEFAULT NULL," +
                    "bank_b varchar(250) DEFAULT NULL," +
                    "swift_b varchar(50) DEFAULT NULL," +
                    "iban_b varchar(50) DEFAULT NULL," +
                    "bank_c varchar(250) DEFAULT NULL," +
                    "swift_c varchar(50) DEFAULT NULL," +
                    "iban_c varchar(50) DEFAULT NULL," +
                    "regnr_d varchar(10) DEFAULT NULL," +
                    "kontonr_d varchar(50) DEFAULT NULL," +
                    "bank_d varchar(250) DEFAULT NULL," +
                    "swift_d varchar(50) DEFAULT NULL," +
                    "iban_d varchar(50) DEFAULT NULL," +
                    "regnr_e varchar(10) DEFAULT NULL," +
                    "kontonr_e varchar(50) DEFAULT NULL," +
                    "bank_e varchar(250) DEFAULT NULL," +
                    "swift_e varchar(50) DEFAULT NULL," +
                    "iban_e varchar(50) DEFAULT NULL," +
                    "kinit varchar(50) DEFAULT NULL," +
                    "regnr_f varchar(50) DEFAULT NULL," +
                    "kontonr_f varchar(50) DEFAULT NULL," +
                    "bank_f varchar(50) DEFAULT NULL," +
                    "swift_f varchar(50) DEFAULT NULL," +
                    "iban_f varchar(50) DEFAULT NULL," +
                    "kfak_sprog int(11) NOT NULL DEFAULT '1'," +
                    "kfak_moms int(11) NOT NULL DEFAULT '1'," +
                    "kfak_valuta int(11) NOT NULL DEFAULT '1'," +
                    "lincensindehaver_faknr_prioritet int(11) NOT NULL DEFAULT '0'," +
                    "kstatus int(11) NOT NULL DEFAULT '1'," +
                    "PRIMARY KEY (Kid,kkundenavn(10),Kkundenr)," +
                    "UNIQUE KEY kunderid_inx (Kid)," +
                    "KEY Kundenavn_inx (kkundenavn)," +
                    "KEY Kundenr_inx (Kkundenr));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE kundetyper (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "editor varchar(50) DEFAULT NULL," +
                    "dato varchar(50) DEFAULT NULL," +
                    "navn varchar(50) DEFAULT NULL," +
                    "rabat int(11) NOT NULL DEFAULT '0'," +
                    "ktlabel varchar(50) DEFAULT NULL);";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE leverand (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "editor varchar(50) DEFAULT NULL," +
                    "dato varchar(50) DEFAULT NULL," +
                    "navn varchar(50) DEFAULT NULL," +
                    "adresse varchar(255) DEFAULT NULL," +
                    "postnr varchar(255) DEFAULT NULL," +
                    "city varchar(255) DEFAULT NULL," +
                    "land varchar(255) DEFAULT NULL," +
                    "tlf varchar(255) DEFAULT NULL," +
                    "mobil varchar(255) DEFAULT NULL," +
                    "email varchar(255) DEFAULT NULL," +
                    "fax varchar(255) DEFAULT NULL," +
                    "besk text," +
                    "levnr varchar(255) DEFAULT NULL," +
                    "type int(11) DEFAULT NULL," +
                    "kpers1 varchar(255) DEFAULT NULL," +
                    "kperstlf1 varchar(255) DEFAULT NULL," +
                    "kpersemail1 varchar(255) DEFAULT NULL," +
                    "kpers2 varchar(255) DEFAULT NULL," +
                    "kperstlf2 varchar(255) DEFAULT NULL," +
                    "kpersemail2 varchar(255) DEFAULT NULL);";
                command.ExecuteNonQueryAsync();

                // Key navn kan ikke hedde key (licens_key) hedder den pt.
                command.CommandText = "CREATE TABLE licens (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "licens varchar(50) DEFAULT '0'," +
                    "licens_key varchar(50) DEFAULT '0'," +
                    "klienter int(11) DEFAULT '0'," +
                    "licenstype varchar(50) DEFAULT NULL," +
                    "owa varchar(50) DEFAULT ''," +
                    "dom varchar(50) DEFAULT ''," +
                    "kontonavn varchar(50) DEFAULT ''," +
                    "kontopw varchar(50) DEFAULT ''," +
                    "smiley int(11) NOT NULL DEFAULT '0'," +
                    "stempelur int(11) NOT NULL DEFAULT '0'," +
                    "lukafm int(11) NOT NULL DEFAULT '0'," +
                    "autogk int(11) NOT NULL DEFAULT '0'," +
                    "sdsk int(11) NOT NULL DEFAULT '0'," +
                    "normtid_st_man time NOT NULL DEFAULT '08:00:00'," +
                    "normtid_sl_man time NOT NULL DEFAULT '17:00:00'," +
                    "normtid_st_tir time NOT NULL DEFAULT '08:00:00'," +
                    "normtid_sl_tir time NOT NULL DEFAULT '17:00:00'," +
                    "normtid_st_ons time NOT NULL DEFAULT '08:00:00'," +
                    "normtid_sl_ons time NOT NULL DEFAULT '17:00:00'," +
                    "normtid_st_tor time NOT NULL DEFAULT '08:00:00'," +
                    "normtid_sl_tor time NOT NULL DEFAULT '17:00:00'," +
                    "normtid_st_fre time NOT NULL DEFAULT '08:00:00'," +
                    "normtid_sl_fre time NOT NULL DEFAULT '17:00:00'," +
                    "normtid_st_lor time NOT NULL DEFAULT '08:00:00'," +
                    "normtid_sl_lor time NOT NULL DEFAULT '17:00:00'," +
                    "normtid_st_son time NOT NULL DEFAULT '08:00:00'," +
                    "normtid_sl_son time NOT NULL DEFAULT '17:00:00'," +
                    "autolukvdato int(11) NOT NULL DEFAULT '0'," +
                    "autolukvdatodato int(11) NOT NULL DEFAULT '28'," +
                    "ignorertid_st time NOT NULL DEFAULT '08:00:00'," +
                    "ignorertid_sl time NOT NULL DEFAULT '08:00:00'," +
                    "stpause double NOT NULL DEFAULT '0'," +
                    "stpause2 double NOT NULL DEFAULT '0'," +
                    "p1_man int(11) NOT NULL DEFAULT '0'," +
                    "p1_tir int(11) NOT NULL DEFAULT '0'," +
                    "p1_ons int(11) NOT NULL DEFAULT '0'," +
                    "p1_tor int(11) NOT NULL DEFAULT '0'," +
                    "p1_fre int(11) NOT NULL DEFAULT '0'," +
                    "p1_lor int(11) NOT NULL DEFAULT '0'," +
                    "p1_son int(11) NOT NULL DEFAULT '0'," +
                    "p2_man int(11) NOT NULL DEFAULT '0'," +
                    "p2_tir int(11) NOT NULL DEFAULT '0'," +
                    "p2_ons int(11) NOT NULL DEFAULT '0'," +
                    "p2_tor int(11) NOT NULL DEFAULT '0'," +
                    "p2_fre int(11) NOT NULL DEFAULT '0'," +
                    "p2_lor int(11) NOT NULL DEFAULT '0'," +
                    "p2_son int(11) NOT NULL DEFAULT '0'," +
                    "brugabningstid int(11) NOT NULL DEFAULT '0'," +
                    "fakturanr double(12, 0) NOT NULL DEFAULT '0'," +
                    "rykkernr int(11) NOT NULL DEFAULT '0'," +
                    "kreditnr double(12, 0) NOT NULL DEFAULT '0'," +
                    "crm int(11) NOT NULL DEFAULT '0'," +
                    "erp int(11) NOT NULL DEFAULT '0'," +
                    "tilbudsnr double(12, 0) NOT NULL DEFAULT '0'," +
                    "jobnr double(12, 0) NOT NULL DEFAULT '0'," +
                    "fakprocent int(11) NOT NULL DEFAULT '0'," +
                    "bgt int(11) NOT NULL DEFAULT '0'," +
                    "autogktimer int(11) NOT NULL DEFAULT '0'," +
                    "listatus int(11) NOT NULL DEFAULT '1'," +
                    "kmdialog int(11) NOT NULL DEFAULT '0'," +
                    "licensstdato date NOT NULL DEFAULT '2002-01-01'," +
                    "timeout_version varchar(50) NOT NULL DEFAULT 'ver2_1'," +
                    "fakturanr_kladde double(12, 0) NOT NULL DEFAULT '3010001'," +
                    "kmpris double(12, 2) NOT NULL DEFAULT '3.56'," +
                    "p1_grp varchar(50) DEFAULT NULL," +
                    "p2_grp varchar(50) DEFAULT NULL," +
                    "jobasnvigv int(11) NOT NULL DEFAULT '1'," +
                    "positiv_aktivering_akt int(11) NOT NULL DEFAULT '0'," +
                    "showeasyreg int(11) NOT NULL DEFAULT '0'," +
                    "showupload int(11) NOT NULL DEFAULT '0'," +
                    "forcebudget_onakttreg int(11) NOT NULL DEFAULT '0'," +
                    "globalfaktor double(12, 2) NOT NULL DEFAULT '1.00'," +
                    "bdgmtypon int(11) NOT NULL DEFAULT '0'," +
                    "regnskabsaar_start date NOT NULL DEFAULT '2001-01-01'," +
                    "forcebudget_onakttreg_afgr int(11) NOT NULL DEFAULT '0'," +
                    "lukaktvdato int(11) NOT NULL DEFAULT '0'," +
                    "salgsans int(11) NOT NULL DEFAULT '0'," +
                    "smileyaggressiv int(11) NOT NULL DEFAULT '0'," +
                    "timerround int(11) NOT NULL DEFAULT '0'," +
                    "teamleder_flad int(11) NOT NULL DEFAULT '0'," +
                    "forcebudget_onakttreg_filt_viskunmbgt int(11) NOT NULL DEFAULT '0'," +
                    "stempelur_hidelogin int(11) NOT NULL DEFAULT '0'," +
                    "stempelur_igno_komkrav int(11) NOT NULL DEFAULT '0'," +
                    "SmiWeekOrMonth int(11) NOT NULL DEFAULT '0'," +
                    "SmiantaldageCount int(11) NOT NULL DEFAULT '1'," +
                    "SmiantaldageCountClock int(11) NOT NULL DEFAULT '12'," +
                    "SmiTeamlederCount int(11) NOT NULL DEFAULT '1'," +
                    "hidesmileyicon int(11) NOT NULL DEFAULT '0'," +
                    "visAktlinjerSimpel int(11) NOT NULL DEFAULT '0'," +
                    "fomr_mandatory int(11) NOT NULL DEFAULT '0'," +
                    "matlageruminemailnot date NOT NULL DEFAULT '2002-01-01'," +
                    "akt_maksbudget_treg int(11) NOT NULL DEFAULT '0'," +
                    "minimumslageremail int(11) NOT NULL DEFAULT '0'," +
                    "fomr_account int(11) NOT NULL DEFAULT '0'," +
                    "visAktlinjerSimpel_datoer int(11) NOT NULL DEFAULT '1'," +
                    "visAktlinjerSimpel_timebudget int(11) NOT NULL DEFAULT '1'," +
                    "visAktlinjerSimpel_realtimer int(11) NOT NULL DEFAULT '1'," +
                    "visAktlinjerSimpel_restimer int(11) NOT NULL DEFAULT '1'," +
                    "visAktlinjerSimpel_medarbtimepriser int(11) NOT NULL DEFAULT '1'," +
                    "visAktlinjerSimpel_medarbrealtimer int(11) NOT NULL DEFAULT '1'," +
                    "visAktlinjerSimpel_akttype int(11) NOT NULL DEFAULT '1'," +
                    "timesimon int(11) NOT NULL DEFAULT '0'," +
                    "timesimh1h2 int(11) NOT NULL DEFAULT '0'," +
                    "timesimtp int(11) NOT NULL DEFAULT '0'," +
                    "budgetakt int(11) NOT NULL DEFAULT '0'," +
                    "akt_maksforecast_treg int(11) NOT NULL DEFAULT '0'," +
                    "traveldietexp_on int(11) NOT NULL DEFAULT '0'," +
                    "traveldietexp_maxhours double(12,2) NOT NULL DEFAULT '0.00'," +
                    "medarbtypligmedarb int(11) NOT NULL DEFAULT '0'," +
                    "pa_aktlist int(11) NOT NULL DEFAULT '0'," +
                    "smiley_agg_lukhard int(11) NOT NULL DEFAULT '0'," +
                    "mobil_week_reg_job_dd int(11) NOT NULL DEFAULT '0'," +
                    "mobil_week_reg_akt_dd int(11) NOT NULL DEFAULT '0'," +
                    "week_showbase_norm_kommegaa int(11) NOT NULL DEFAULT '0'," +
                    "mobil_week_reg_akt_dd_forvalgt int(11) NOT NULL DEFAULT '0'," +
                    "budget_mandatory int(11) NOT NULL DEFAULT '0'," +
                    "tilbud_mandatory int(11) NOT NULL DEFAULT '0'," +
                    "show_salgsomk_mandatory int(11) NOT NULL DEFAULT '0'," +
                    "multible_licensindehavere int(11) NOT NULL DEFAULT '0'," +
                    "fakturanr_2 int(11) NOT NULL DEFAULT '0'," +
                    "kreditnr_2 int(11) NOT NULL DEFAULT '0'," +
                    "fakturanr_3 int(11) NOT NULL DEFAULT '0'," +
                    "kreditnr_3 int(11) NOT NULL DEFAULT '0'," +
                    "fakturanr_4 int(11) NOT NULL DEFAULT '0'," +
                    "kreditnr_4 int(11) NOT NULL DEFAULT '0'," +
                    "fakturanr_5 int(11) NOT NULL DEFAULT '0'," +
                    "kreditnr_5 int(11) NOT NULL DEFAULT '0'," +
                    "fakturanr_kladde_2 int(11) NOT NULL DEFAULT '0'," +
                    "fakturanr_kladde_3 int(11) NOT NULL DEFAULT '0'," +
                    "fakturanr_kladde_4 int(11) NOT NULL DEFAULT '0'," +
                    "fakturanr_kladde_5 int(11) NOT NULL DEFAULT '0'," +
                    "showeasyreg_per int(11) NOT NULL DEFAULT '0'," +
                    "smiweekormonth_hr int(11) NOT NULL DEFAULT '0'," +
                    "vis_resplanner int(11) DEFAULT '0'," +
                    "vis_favorit int(11) DEFAULT '1'," +
                    "vis_projektgodkend int(11) DEFAULT '0'," +
                    "pa_tilfojvmedopret int(11) DEFAULT '0'," +
                    "fakturanr_6 int(11) NOT NULL DEFAULT '0'," +
                    "kreditnr_6 int(11) NOT NULL DEFAULT '0'," +
                    "fakturanr_7 int(11) NOT NULL DEFAULT '0'," +
                    "kreditnr_7 int(11) NOT NULL DEFAULT '0'," +
                    "fakturanr_8 int(11) NOT NULL DEFAULT '0'," +
                    "kreditnr_8 int(11) NOT NULL DEFAULT '0'," +
                    "fakturanr_9 int(11) NOT NULL DEFAULT '0'," +
                    "kreditnr_9 int(11) NOT NULL DEFAULT '0'," +
                    "fakturanr_10 int(11) NOT NULL DEFAULT '0'," +
                    "kreditnr_10 int(11) NOT NULL DEFAULT '0'," +
                    "fakturanr_kladde_6 int(11) NOT NULL DEFAULT '0'," +
                    "fakturanr_kladde_7 int(11) NOT NULL DEFAULT '0'," +
                    "fakturanr_kladde_8 int(11) NOT NULL DEFAULT '0'," +
                    "fakturanr_kladde_9 int(11) NOT NULL DEFAULT '0'," +
                    "fakturanr_kladde_10 int(11) NOT NULL DEFAULT '0'," +
                    "stpause3 double NOT NULL DEFAULT '0'," +
                    "stpause4 double NOT NULL DEFAULT '0'," +
                    "p3_grp varchar(50) DEFAULT NULL," +
                    "p4_grp varchar(50) DEFAULT NULL," +
                    "p3_man int(11) NOT NULL DEFAULT '0'," +
                    "p3_tir int(11) NOT NULL DEFAULT '0'," +
                    "p3_ons int(11) NOT NULL DEFAULT '0'," +
                    "p3_tor int(11) NOT NULL DEFAULT '0'," +
                    "p3_fre int(11) NOT NULL DEFAULT '0'," +
                    "p3_lor int(11) NOT NULL DEFAULT '0'," +
                    "p3_son int(11) NOT NULL DEFAULT '0'," +
                    "p4_man int(11) NOT NULL DEFAULT '0'," +
                    "p4_tir int(11) NOT NULL DEFAULT '0'," +
                    "p4_ons int(11) NOT NULL DEFAULT '0'," +
                    "p4_tor int(11) NOT NULL DEFAULT '0'," +
                    "p4_fre int(11) NOT NULL DEFAULT '0'," +
                    "p4_lor int(11) NOT NULL DEFAULT '0'," +
                    "p4_son int(11) NOT NULL DEFAULT '0'," +
                    "stempelur_hidelogin_onlymonitor int(11) NOT NULL DEFAULT '0'," +
                    "ignorertid_st_logud time NOT NULL DEFAULT '18:00:00'," +
                    "ignorertid_sl_logud time NOT NULL DEFAULT '18:00:00'," +
                    "ignorertid_st_2 time NOT NULL DEFAULT '09:00:00'," +
                    "ignorertid_sl_2 time NOT NULL DEFAULT '09:00:00'," +
                    "ignorertid_st_logud_2 time NOT NULL DEFAULT '21:00:00'," +
                    "ignorertid_sl_logud_2 time NOT NULL DEFAULT '21:00:00'," +
                    "sperretid_grp varchar(50) DEFAULT ''," +
                    "sperretid_grp2 varchar(50) DEFAULT ''," +
                    "vis_lager int(11) DEFAULT '1'," +
                    "password_visning int(11) DEFAULT '0'," +
                    "job_felt_fomr int(11) DEFAULT '0'," +
                    "job_felt_salgsans int(11) DEFAULT '0'," +
                    "job_felt_rekno int(11) DEFAULT '0'," +
                    "job_felt_intnote int(11) DEFAULT '0'," +
                    "simple_joblist_jobnavn int(11) DEFAULT '0'," +
                    "simple_joblist_jobnr int(11) DEFAULT '0'," +
                    "simple_joblist_kunde int(11) DEFAULT '0'," +
                    "simple_joblist_ans int(11) DEFAULT '0'," +
                    "simple_joblist_salgsans int(11) DEFAULT '0'," +
                    "simple_joblist_fomr int(11) DEFAULT '0'," +
                    "simple_joblist_prg int(11) DEFAULT '0'," +
                    "simple_joblist_stdato int(11) DEFAULT '0'," +
                    "simple_joblist_sldato int(11) DEFAULT '0'," +
                    "simple_joblist_status int(11) DEFAULT '0'," +
                    "simple_joblist_tidsforbrug int(11) DEFAULT '0'," +
                    "simple_joblist_budgettid int(11) DEFAULT '0'," +
                    "smiley_agg_lukhard_igngrp varchar(250) NOT NULL DEFAULT '');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE login_historik (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "dato date DEFAULT NULL," +
                    "login datetime DEFAULT NULL," +
                    "logud datetime DEFAULT NULL," +
                    "mid int(11) NOT NULL DEFAULT '0'," +
                    "stempelurindstilling int(11) NOT NULL DEFAULT '0'," +
                    "minutter double NOT NULL DEFAULT '0'," +
                    "kommentar text," +
                    "manuelt_afsluttet int(11) NOT NULL DEFAULT '0'," +
                    "login_first datetime DEFAULT NULL," +
                    "logud_first datetime DEFAULT NULL," +
                    "manuelt_oprettet int(11) NOT NULL DEFAULT '0'," +
                    "ipn varchar(50) NOT NULL DEFAULT '0'," +
                    "lh_longitude double(25, 20) NOT NULL DEFAULT '0.00000000000000000000'," +
                    "lh_latitude double(25, 20) NOT NULL DEFAULT '0.00000000000000000000'," +
                    "lh_longitude_logud double(25, 20) NOT NULL DEFAULT '0.00000000000000000000'," +
                    "lh_latitude_logud double(25, 20) NOT NULL DEFAULT '0.00000000000000000000'," +
                    "KEY inx_lh_mid2 (mid)," +
                    "KEY inx_lh_mid_dt2 (mid,dato));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE login_historik_aktivejob_rel (lha_id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "lha_mid int(11) NOT NULL DEFAULT '0'," +
                    "lha_aktid int(11) NOT NULL DEFAULT '0'," +
                    "lha_jobid int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE login_historik_terminal (lht_id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "lht_dt timestamp(6) NULL DEFAULT NULL," +
                    "lht_logind datetime DEFAULT NULL," +
                    "lht_mid int(11) NOT NULL DEFAULT '0'," +
                    "lht_type int(11) NOT NULL DEFAULT '0'," +
                    "lht_status int(11) NOT NULL DEFAULT '0'," +
                    "fo_logud_kode int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE lon_korsel (lk_id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "lk_dato date DEFAULT NULL," +
                    "lk_editor varchar(100) DEFAULT NULL," +
                    "lk_stamp timestamp(6) NULL DEFAULT NULL," +
                    "lk_close_projects int(11) NOT NULL DEFAULT '0'," +
                    "lk_correction_on int(11) NOT NULL DEFAULT '0'," +
                    "lk_medarbtype varchar(250) DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE materiale_forbrug (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "matid int(11) NOT NULL DEFAULT '0'," +
                    "matantal double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "matnavn varchar(255) DEFAULT NULL," +
                    "matvarenr varchar(255) DEFAULT NULL," +
                    "matkobspris double(12, 3) NOT NULL DEFAULT '0.000'," +
                    "matsalgspris double(12, 3) NOT NULL DEFAULT '0.000'," +
                    "matenhed varchar(255) DEFAULT NULL," +
                    "jobid int(11) NOT NULL DEFAULT '0'," +
                    "dato date NOT NULL DEFAULT '0000-00-00'," +
                    "editor varchar(255) NOT NULL DEFAULT ''," +
                    "usrid int(11) NOT NULL DEFAULT '0'," +
                    "matgrp int(11) NOT NULL DEFAULT '0'," +
                    "forbrugsdato date DEFAULT NULL," +
                    "serviceaft int(11) NOT NULL DEFAULT '0'," +
                    "valuta int(11) NOT NULL DEFAULT '1'," +
                    "bilagsnr varchar(50) DEFAULT NULL," +
                    "intkode int(11) NOT NULL DEFAULT '0'," +
                    "kurs double(12, 2) NOT NULL DEFAULT '100.00'," +
                    "afregnet int(11) NOT NULL DEFAULT '0'," +
                    "personlig int(11) NOT NULL DEFAULT '0'," +
                    "godkendt int(11) NOT NULL DEFAULT '0'," +
                    "gkaf varchar(150) DEFAULT NULL," +
                    "gkdato date DEFAULT NULL," +
                    "erfak int(11) NOT NULL DEFAULT '0'," +
                    "extsysid double NOT NULL DEFAULT '0'," +
                    "aktid int(11) NOT NULL DEFAULT '0'," +
                    "ava double(12, 4) NOT NULL DEFAULT '0.0000'," +
                    "mf_konto int(11) NOT NULL DEFAULT '0'," +
                    "filnavn varchar(250) NOT NULL DEFAULT ''," +
                    "mattype int(12) DEFAULT '0'," +
                    "basic_valuta varchar(255) DEFAULT 'NA'," +
                    "basic_kurs double DEFAULT '0'," +
                    "basic_belob double DEFAULT '0'," +
                    "KEY mf_jobid_inx(jobid));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE materiale_grp (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "editor varchar(50) DEFAULT NULL," +
                    "dato varchar(50) DEFAULT NULL," +
                    "navn varchar(50) DEFAULT NULL," +
                    "nummer varchar(255) DEFAULT NULL," +
                    "av double NOT NULL DEFAULT '0'," +
                    "mgkundeid int(11) NOT NULL DEFAULT '0'," +
                    "medarbansv int(11) NOT NULL DEFAULT '0'," +
                    "matgrp_konto int(11) NOT NULL DEFAULT '0'," +
                    "fragtpris double(12, 2) DEFAULT '0.00'," +
                    "matgrp_type int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = " CREATE TABLE materiale_ordrer (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "ordrenr int(11) NOT NULL DEFAULT '0'," +
                    "editor varchar(50) NOT NULL DEFAULT ''," +
                    "dato date NOT NULL DEFAULT '0000-00-00'," +
                    "mid int(11) NOT NULL DEFAULT '0'," +
                    "ordredato date NOT NULL DEFAULT '0000-00-00'," +
                    "status enum('Igang','Afsluttet','Godkendt','Udført','Venter','Kladde') NOT NULL DEFAULT 'Kladde'," +
                    "ts timestamp(6) NULL DEFAULT NULL," +
                    "UNIQUE KEY inx_ordid(ordrenr));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE materiale_ordrer_linier (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "ordreid int(11) NOT NULL DEFAULT '0'," +
                    "matid int(11) NOT NULL DEFAULT '0'," +
                    "navn varchar(255) DEFAULT NULL," +
                    "betegnelse varchar(255) DEFAULT NULL," +
                    "varenr varchar(11) NOT NULL DEFAULT '0'," +
                    "gruppe varchar(11) DEFAULT '0'," +
                    "status int(11) NOT NULL DEFAULT '0'," +
                    "antal double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "KEY inx_ordreid(ordreid));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE materialer (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "editor varchar(50) DEFAULT NULL," +
                    "dato varchar(50) DEFAULT NULL," +
                    "navn varchar(50) DEFAULT NULL," +
                    "matgrp int(11) NOT NULL DEFAULT '0'," +
                    "varenr varchar(255) NOT NULL DEFAULT ''," +
                    "indkobspris double(12, 3) NOT NULL DEFAULT '0.000'," +
                    "salgspris double(12, 3) NOT NULL DEFAULT '0.000'," +
                    "antal int(11) NOT NULL DEFAULT '0'," +
                    "arrestordredato date DEFAULT NULL," +
                    "ptid int(11) NOT NULL DEFAULT '0'," +
                    "ptid_arr int(11) NOT NULL DEFAULT '0'," +
                    "enhed varchar(255) NOT NULL DEFAULT 'Stk.'," +
                    "pic int(255) NOT NULL DEFAULT '1'," +
                    "leva int(11) NOT NULL DEFAULT '0'," +
                    "levb int(11) NOT NULL DEFAULT '0'," +
                    "sera int(11) NOT NULL DEFAULT '0'," +
                    "serb int(11) NOT NULL DEFAULT '0'," +
                    "minlager int(11) NOT NULL DEFAULT '0'," +
                    "sortorder double NOT NULL DEFAULT '1000'," +
                    "betegnelse text," +
                    "lokation varchar(255) DEFAULT NULL," +
                    "fabrikat varchar(255) DEFAULT NULL," +
                    "UNIQUE KEY varenr_matgrp_inx(matgrp, varenr)," +
                    "KEY grp_inx(matgrp)," +
                    "KEY navn_inx(navn)," +
                    "KEY varenr_inx(varenr));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE medarbejder_flexsaldo (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "mf_dato date NOT NULL DEFAULT '2002-01-01'," +
                    "mf_mid double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "mf_flex_norm_real double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "mf_flex_norm_kommegaa double(12, 2) NOT NULL DEFAULT '0.00');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE medarbejdere (Mid int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "Mnavn varchar(50) DEFAULT NULL," +
                    "Mnr int(11) DEFAULT '0'," +
                    "Mansat varchar(50) DEFAULT 'No'," +
                    "Mkat_old varchar(50) DEFAULT NULL," +
                    "login varchar(50) DEFAULT NULL," +
                    "pw varchar(50) NOT NULL DEFAULT ''," +
                    "dato varchar(50) DEFAULT NULL," +
                    "editor varchar(50) DEFAULT NULL," +
                    "lastlogin varchar(50) DEFAULT NULL," +
                    "Brugergruppe int(11) DEFAULT '0'," +
                    "Medarbejdertype int(11) DEFAULT '0'," +
                    "Medarbejderinfo text," +
                    "email varchar(100) DEFAULT NULL," +
                    "tsacrm int(1) DEFAULT '0'," +
                    "visguide int(11) DEFAULT '0'," +
                    "smilord int(11) NOT NULL DEFAULT '1'," +
                    "exchkonto varchar(50) DEFAULT ''," +
                    "init varchar(50) DEFAULT NULL," +
                    "timereg int(11) NOT NULL DEFAULT '0'," +
                    "forecaststamp date DEFAULT NULL," +
                    "ansatdato date NOT NULL DEFAULT '2002-01-01'," +
                    "opsagtdato date NOT NULL DEFAULT '2044-01-01'," +
                    "sprog int(11) NOT NULL DEFAULT '1'," +
                    "nyhedsbrev int(11) NOT NULL DEFAULT '0'," +
                    "madr varchar(100) DEFAULT NULL," +
                    "mpostnr varchar(50) DEFAULT NULL," +
                    "mcity varchar(100) DEFAULT NULL," +
                    "mland varchar(50) DEFAULT NULL," +
                    "mtlf varchar(50) DEFAULT NULL," +
                    "mcpr varchar(50) DEFAULT NULL," +
                    "mkoregnr varchar(50) DEFAULT NULL," +
                    "bonus int(11) NOT NULL DEFAULT '0'," +
                    "bankaccount varchar(255) DEFAULT NULL," +
                    "visskiftversion int(11) NOT NULL DEFAULT '0'," +
                    "medarbejdertype_grp int(11) NOT NULL DEFAULT '0'," +
                    "showtimereg_start_stop int(11) NOT NULL DEFAULT '0'," +
                    "timer_ststop int(11) NOT NULL DEFAULT '0'," +
                    "create_newemployee int(11) NOT NULL DEFAULT '0'," +
                    "med_lincensindehaver int(11) NOT NULL DEFAULT '0'," +
                    "medarbejder_rfid varchar(255) DEFAULT NULL," +
                    "measyregtimer double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "med_cal varchar(50) NOT NULL DEFAULT 'DK'," +
                    "mfrokost int(11) DEFAULT '0'," +
                    "firma varchar(250) DEFAULT NULL," +
                    "UNIQUE KEY inx_mids(Mid)," +
                    "KEY inx_mansat(Mansat)," +
                    "KEY inx_minit(init));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE medarbejdertyper (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "type varchar(50) DEFAULT NULL," +
                    "dato varchar(50) DEFAULT NULL," +
                    "editor varchar(50) DEFAULT NULL," +
                    "timepris double DEFAULT '0'," +
                    "kostpris double DEFAULT NULL," +
                    "normtimer_man double(5, 2) DEFAULT '0.00'," +
                    "normtimer_tir double(5, 2) DEFAULT '0.00'," +
                    "normtimer_ons double(5, 2) DEFAULT '0.00'," +
                    "normtimer_tor double(5, 2) DEFAULT '0.00'," +
                    "normtimer_fre double(5, 2) DEFAULT '0.00'," +
                    "normtimer_lor double(5, 2) DEFAULT '0.00'," +
                    "normtimer_son double(5, 2) DEFAULT '0.00'," +
                    "timepris_a1 double NOT NULL DEFAULT '0'," +
                    "timepris_a2 double NOT NULL DEFAULT '0'," +
                    "timepris_a3 double NOT NULL DEFAULT '0'," +
                    "timepris_a4 double NOT NULL DEFAULT '0'," +
                    "timepris_a5 double NOT NULL DEFAULT '0'," +
                    "tp0_valuta int(11) NOT NULL DEFAULT '1'," +
                    "tp1_valuta int(11) NOT NULL DEFAULT '1'," +
                    "tp2_valuta int(11) NOT NULL DEFAULT '1'," +
                    "tp3_valuta int(11) NOT NULL DEFAULT '1'," +
                    "tp4_valuta int(11) NOT NULL DEFAULT '1'," +
                    "tp5_valuta int(11) NOT NULL DEFAULT '1'," +
                    "sostergp int(11) NOT NULL DEFAULT '0'," +
                    "mtsortorder int(11) NOT NULL DEFAULT '0'," +
                    "mgruppe int(11) NOT NULL DEFAULT '0'," +
                    "afslutugekri int(11) NOT NULL DEFAULT '0'," +
                    "afslutugekri_proc double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "noflex int(11) NOT NULL DEFAULT '0'," +
                    "kostpristarif_A double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "kostpristarif_B double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "kostpristarif_C double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "kostpristarif_D double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "kp1_valuta int(11) NOT NULL DEFAULT '1'," +
                    "mt_mobil_visstopur int(11) NOT NULL DEFAULT '0'," +
                    "feriesats double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "mtype_passiv int(11) DEFAULT '0'," +
                    "mtype_maxflex double(12, 2) DEFAULT '0.00');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE medarbejdertyper_historik (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "mid int(11) NOT NULL DEFAULT '0'," +
                    "mtype int(11) NOT NULL DEFAULT '0'," +
                    "mtypedato date NOT NULL DEFAULT '0000-00-00');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE medarbejdertyper_timebudget (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "jobid int(11) NOT NULL DEFAULT '0'," +
                    "typeid int(11) NOT NULL DEFAULT '0'," +
                    "timer double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "timepris double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "faktor double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "belob double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "belob_ff double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "kostpris double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "KEY inx2_mtb_jobid_typeid(jobid, typeid));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE medarbtyper_grp (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "dato date DEFAULT NULL," +
                    "editor varchar(100) DEFAULT NULL," +
                    "navn varchar(100) DEFAULT NULL," +
                    "opencalc int(11) NOT NULL DEFAULT '1');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE milepale (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "navn varchar(50) NOT NULL DEFAULT ''," +
                    "dato timestamp(6) NULL DEFAULT NULL," +
                    "type int(11) NOT NULL DEFAULT '0'," +
                    "beskrivelse text," +
                    "editor varchar(50) DEFAULT NULL," +
                    "milepal_dato date NOT NULL DEFAULT '0000-00-00'," +
                    "jid int(11) NOT NULL DEFAULT '0'," +
                    "aftaleid int(11) NOT NULL DEFAULT '0'," +
                    "belob double(12, 2) NOT NULL DEFAULT '0.00');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE milepale_typer (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "dato date NOT NULL DEFAULT '0000-00-00'," +
                    "editor varchar(50) DEFAULT NULL," +
                    "navn varchar(50) DEFAULT NULL," +
                    "ikon varchar(50) DEFAULT NULL," +
                    "mpt_fak int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE momsafsluttet (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "dato date NOT NULL DEFAULT '0000-00-00'," +
                    "editor varchar(255) DEFAULT NULL," +
                    "afslutdato date NOT NULL DEFAULT '0000-00-00'," +
                    "kommentar varchar(255) DEFAULT NULL);";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE momskoder (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "navn varchar(50) DEFAULT NULL," +
                    "editor varchar(50) DEFAULT NULL," +
                    "dato date DEFAULT '2000-00-00'," +
                    "kvotient double NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE mtype_mal_fordel_fomr (mmff_id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "mmff_mal double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "mmff_mtype int(11) NOT NULL DEFAULT '0'," +
                    "mmff_fomr int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE national_holidays (nh_id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "nh_name varchar(100) NOT NULL DEFAULT ''," +
                    "nh_duration int(11) NOT NULL DEFAULT '0'," +
                    "nh_editor varchar(50) DEFAULT NULL," +
                    "nh_editor_date date NOT NULL DEFAULT '2010-01-01'," +
                    "nh_date date NOT NULL DEFAULT '2010-01-01'," +
                    "nh_country varchar(50) NOT NULL DEFAULT ''," +
                    "nh_open int(11) NOT NULL DEFAULT '1'," +
                    "nh_sortorder int(11) NOT NULL DEFAULT '1'," +
                    "nh_projgrp varchar(50) DEFAULT NULL);";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE news_rel (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "medarbid int(11) NOT NULL," +
                    "newsid int(11) NOT NULL," +
                    "newsread int(11) DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE nogletalskoder (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "navn varchar(50) DEFAULT NULL," +
                    "editor varchar(50) DEFAULT NULL," +
                    "dato date DEFAULT '0000-00-00'," +
                    "ap_type int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE pdf_values (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "pdf_txt longtext," +
                    "pdf_footer text);";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE posteringer (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "editor varchar(50) DEFAULT NULL," +
                    "dato date DEFAULT '0000-00-00'," +
                    "bilagstype varchar(50) DEFAULT NULL," +
                    "modkontonr double(12, 2) DEFAULT NULL," +
                    "kontonr double(12, 0) DEFAULT NULL," +
                    "bilagsnr varchar(50) NOT NULL DEFAULT '0'," +
                    "beloeb double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "nettobeloeb double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "moms double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "tekst text," +
                    "reference varchar(50) DEFAULT NULL," +
                    "status int(11) NOT NULL DEFAULT '0'," +
                    "att int(11) NOT NULL DEFAULT '0'," +
                    "posteringsdato date NOT NULL DEFAULT '0000-00-00'," +
                    "oprid int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE produktion_mat_rel (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "produktionsId int(11) NOT NULL DEFAULT '0'," +
                    "materialeId int(11) NOT NULL DEFAULT '0'," +
                    "afhId int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE produktioner (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "editor varchar(50) DEFAULT NULL," +
                    "dato varchar(50) DEFAULT NULL," +
                    "navn varchar(50) DEFAULT NULL," +
                    "nummer varchar(255) DEFAULT NULL);";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE progrupperelationer (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "ProjektgruppeId int(11) DEFAULT '0'," +
                    "MedarbejderId int(11) DEFAULT '0'," +
                    "teamleder int(11) NOT NULL DEFAULT '0'," +
                    "notificer int(11) NOT NULL DEFAULT '0'," +
                    "KEY pro_medid_inx(id)," +
                    "KEY progrp_id_inx(ProjektgruppeId)," +
                    "KEY inx_progrprel_jm(ProjektgruppeId, MedarbejderId));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE projektgrupper (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "navn varchar(50) DEFAULT NULL," +
                    "dato date DEFAULT NULL," +
                    "editor varchar(50) DEFAULT NULL," +
                    "opengp int(11) NOT NULL DEFAULT '1'," +
                    "orgvir int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE ressourcer (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "jobid int(11) DEFAULT NULL," +
                    "aktid int(11) DEFAULT NULL," +
                    "dato date NOT NULL DEFAULT '0000-00-00'," +
                    "medarbtype int(11) DEFAULT NULL," +
                    "timer double DEFAULT '0'," +
                    "mid int(11) DEFAULT NULL," +
                    "starttp time NOT NULL DEFAULT '23:59:59'," +
                    "sluttp time NOT NULL DEFAULT '23:59:59'," +
                    "serieid int(11) NOT NULL DEFAULT '0'," +
                    "exch_appid varchar(50) DEFAULT ''," +
                    "KEY jobaktmid_dato(jobid, aktid, dato, mid)," +
                    "KEY mid_inx(mid));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE ressourcer_md (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "jobid int(11) NOT NULL DEFAULT '0'," +
                    "medid int(11) NOT NULL DEFAULT '0'," +
                    "md int(11) NOT NULL DEFAULT '0'," +
                    "aar int(11) NOT NULL DEFAULT '0'," +
                    "timer double(11, 2) DEFAULT '0.00'," +
                    "uge int(11) NOT NULL DEFAULT '0'," +
                    "aktid int(11) NOT NULL DEFAULT '0'," +
                    "proc double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "KEY medid_inx(medid)," +
                    "KEY jam_inx2(jobid, aktid, medid));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE ressourcer_md_spor (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "jobid int(11) NOT NULL DEFAULT '0'," +
                    "aktid int(11) NOT NULL DEFAULT '0'," +
                    "medid int(11) NOT NULL DEFAULT '0'," +
                    "aar int(11) NOT NULL DEFAULT '0'," +
                    "md int(11) NOT NULL DEFAULT '0'," +
                    "uge int(11) NOT NULL DEFAULT '0'," +
                    "timer double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "editor int(11) DEFAULT '0'," +
                    "editdate date NOT NULL DEFAULT '2002-01-01'," +
                    "edittime time NOT NULL DEFAULT '00:00:00');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE ressourcer_ramme (id int(11) NOT NULL AUTO_INCREMENT," +
                    "jobid int(11) NOT NULL DEFAULT '0'," +
                    "medid int(11) NOT NULL DEFAULT '0'," +
                    "aar int(11) NOT NULL DEFAULT '0'," +
                    "timer double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "aktid bigint(20) NOT NULL DEFAULT '0'," +
                    "fctimepris double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "fctimeprish2 double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "rr_budgetbelob double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "PRIMARY KEY(id));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE rettigheder (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "brugergruppeId int(11) DEFAULT '0'," +
                    "jobId int(11) DEFAULT '0'," +
                    "rettighedstype int(11) DEFAULT '0'," +
                    "dato date DEFAULT NULL," +
                    "editor varchar(50) DEFAULT NULL);";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE sdsk (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "emne text NOT NULL," +
                    "besk text NOT NULL," +
                    "fil varchar(255) NOT NULL DEFAULT ''," +
                    "type int(11) NOT NULL DEFAULT '0'," +
                    "prioitet int(11) NOT NULL DEFAULT '0'," +
                    "kundeid int(11) NOT NULL DEFAULT '0'," +
                    "jobid int(11) NOT NULL DEFAULT '0'," +
                    "ansvarlig int(11) NOT NULL DEFAULT '0'," +
                    "esttid double NOT NULL DEFAULT '0'," +
                    "dato date NOT NULL DEFAULT '0000-00-00'," +
                    "editor varchar(255) NOT NULL DEFAULT ''," +
                    "tidspunkt time NOT NULL DEFAULT '00:00:00'," +
                    "public int(11) NOT NULL DEFAULT '0'," +
                    "status int(11) NOT NULL DEFAULT '0'," +
                    "sogeord1 varchar(255) NOT NULL DEFAULT ''," +
                    "sogeord2 varchar(255) NOT NULL DEFAULT ''," +
                    "sogeord3 varchar(255) NOT NULL DEFAULT ''," +
                    "sogeord4 varchar(255) NOT NULL DEFAULT ''," +
                    "rsptid_b_id int(11) NOT NULL DEFAULT '0'," +
                    "rsptid_b_tid double NOT NULL DEFAULT '0'," +
                    "rsptid_a_id int(11) NOT NULL DEFAULT '0'," +
                    "rsptid_a_tid double NOT NULL DEFAULT '0'," +
                    "editor2 varchar(50) DEFAULT NULL," +
                    " dato2 date DEFAULT NULL," +
                    "creator int(11) NOT NULL DEFAULT '0'," +
                    "kpers int(11) NOT NULL DEFAULT '0'," +
                    "kpers2 varchar(50) DEFAULT NULL," +
                    "kpers2email varchar(50) DEFAULT NULL," +
                    "duedate datetime NOT NULL DEFAULT '2001-01-01 00:00:00'," +
                    "useduedate int(11) NOT NULL DEFAULT '0'," +
                    "priotype int(11) NOT NULL DEFAULT '0'," +
                    "sortorder int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE sdsk_prio_grp (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "editor varchar(50) DEFAULT NULL," +
                    "dato varchar(50) DEFAULT NULL," +
                    "navn varchar(50) DEFAULT NULL," +
                    "gemtider int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE sdsk_prio_typ (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "editor varchar(50) DEFAULT NULL," +
                    "dato varchar(50) DEFAULT NULL," +
                    "navn varchar(50) DEFAULT NULL);";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE sdsk_prioitet (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "editor varchar(50) DEFAULT NULL," +
                    "dato varchar(50) DEFAULT NULL," +
                    "navn varchar(50) DEFAULT NULL," +
                    "responstid double(6, 2) NOT NULL DEFAULT '0.00'," +
                    "faktor double(6, 2) NOT NULL DEFAULT '0.00'," +
                    "kunweekend int(11) NOT NULL DEFAULT '0'," +
                    "advisering double(6, 2) NOT NULL DEFAULT '0.00'," +
                    "priogrp int(11) NOT NULL DEFAULT '0'," +
                    "type int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE sdsk_rel (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "besk text NOT NULL," +
                    "dato date NOT NULL DEFAULT '0000-00-00'," +
                    "editor varchar(255) NOT NULL DEFAULT ''," +
                    "sdsk_rel int(11) NOT NULL DEFAULT '0'," +
                    "sdsktidspunkt time NOT NULL DEFAULT '00:00:00'," +
                    "public int(11) NOT NULL DEFAULT '0'," +
                    "losning int(11) NOT NULL DEFAULT '0'," +
                    "sogeord varchar(255) NOT NULL DEFAULT ''," +
                    "sdskdato date NOT NULL DEFAULT '0000-00-00'," +
                    "editor2 varchar(50) DEFAULT NULL," +
                    "dato2 date DEFAULT NULL);";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE sdsk_status (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "editor varchar(50) NOT NULL DEFAULT ''," +
                    "dato varchar(50) NOT NULL DEFAULT ''," +
                    "navn varchar(50) NOT NULL DEFAULT ''," +
                    "farve varchar(255) NOT NULL DEFAULT ''," +
                    "rsptid_a int(11) NOT NULL DEFAULT '0'," +
                    "rsptid_b int(11) NOT NULL DEFAULT '0'," +
                    "vispahovedliste int(11) NOT NULL DEFAULT '0'," +
                    "progrp int(11) NOT NULL DEFAULT '0'," +
                    "luk int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE sdsk_typer (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "editor varchar(50) DEFAULT NULL," +
                    "dato varchar(50) DEFAULT NULL," +
                    "navn varchar(50) DEFAULT NULL);";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE serviceaft (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "editor varchar(50) DEFAULT NULL," +
                    "dato varchar(50) DEFAULT NULL," +
                    "navn varchar(50) DEFAULT NULL," +
                    "varenr varchar(255) DEFAULT NULL," +
                    "pris double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "enheder double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "stdato date DEFAULT NULL," +
                    "sldato date DEFAULT NULL," +
                    "perafg int(11) NOT NULL DEFAULT '0'," +
                    "kundeid int(11) NOT NULL DEFAULT '0'," +
                    "status int(11) NOT NULL DEFAULT '0'," +
                    "advihvor int(11) NOT NULL DEFAULT '0'," +
                    "advitype int(11) NOT NULL DEFAULT '0'," +
                    "erfornyet int(11) NOT NULL DEFAULT '0'," +
                    "besk text," +
                    "aftalenr int(11) NOT NULL DEFAULT '0'," +
                    "erfaktureret int(11) NOT NULL DEFAULT '0'," +
                    "fordel int(11) NOT NULL DEFAULT '0'," +
                    "overfortsaldo double NOT NULL DEFAULT '0'," +
                    "valuta int(11) NOT NULL DEFAULT '1');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE signoff (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "editor varchar(50) DEFAULT NULL," +
                    "dato date NOT NULL DEFAULT '0000-00-00'," +
                    "navn varchar(50) DEFAULT NULL," +
                    "adresse varchar(50) DEFAULT NULL," +
                    "postnr varchar(50) DEFAULT NULL," +
                    "town varchar(50) DEFAULT NULL," +
                    "kpersnavn varchar(50) DEFAULT NULL," +
                    "kpersstilling varchar(50) DEFAULT NULL," +
                    "kpersfirma varchar(50) DEFAULT NULL," +
                    "kperstlf varchar(50) DEFAULT NULL," +
                    "kpersfax varchar(50) DEFAULT NULL," +
                    "kpersmobil varchar(50) DEFAULT NULL," +
                    "kpersemail varchar(50) DEFAULT NULL," +
                    "itansnavn varchar(50) DEFAULT NULL," +
                    "itansfirma varchar(50) DEFAULT NULL," +
                    "itanstlf varchar(50) DEFAULT NULL," +
                    "itansfax varchar(50) DEFAULT NULL," +
                    "itansmobil varchar(50) DEFAULT NULL," +
                    "itansemail varchar(50) DEFAULT NULL," +
                    "virktype varchar(50) DEFAULT NULL," +
                    "virktypeandet varchar(50) DEFAULT NULL," +
                    "pafgift varchar(50) DEFAULT NULL," +
                    "tohsalgskon varchar(50) DEFAULT NULL," +
                    "tohsyskon varchar(50) DEFAULT NULL," +
                    "placering varchar(50) DEFAULT NULL," +
                    "placeringandet varchar(50) DEFAULT NULL," +
                    "model varchar(50) DEFAULT NULL," +
                    "usednew varchar(50) DEFAULT NULL," +
                    "proveops varchar(50) DEFAULT NULL," +
                    "levdato date DEFAULT '0000-00-00'," +
                    "levtid time DEFAULT NULL," +
                    "insdato date DEFAULT '0000-00-00'," +
                    "instid time DEFAULT NULL," +
                    "kurdato date DEFAULT '0000-00-00'," +
                    "kurtid time DEFAULT NULL," +
                    "kurskonsul varchar(50) DEFAULT NULL," +
                    "rj45 varchar(50) DEFAULT NULL," +
                    "ethernet varchar(50) DEFAULT NULL," +
                    "token varchar(50) DEFAULT NULL," +
                    "IPX_SPX varchar(50) DEFAULT NULL," +
                    "TCP_IP varchar(50) DEFAULT NULL," +
                    "appleTalk varchar(50) DEFAULT NULL," +
                    "novell3 varchar(50) DEFAULT NULL," +
                    "novell4 varchar(50) DEFAULT NULL," +
                    "novell5 varchar(50) DEFAULT NULL," +
                    "novell6 varchar(50) DEFAULT NULL," +
                    "nt4 varchar(50) DEFAULT NULL," +
                    "k2 varchar(50) DEFAULT NULL," +
                    "k03 varchar(50) DEFAULT NULL," +
                    "ipadr varchar(50) DEFAULT NULL," +
                    "subnet varchar(50) DEFAULT NULL," +
                    "gateway varchar(50) DEFAULT NULL," +
                    "dns varchar(50) DEFAULT NULL," +
                    "smtp varchar(50) DEFAULT NULL," +
                    "ftp varchar(50) DEFAULT NULL," +
                    "bnavn varchar(50) DEFAULT NULL," +
                    "pw varchar(50) DEFAULT NULL," +
                    "arb_win98 varchar(50) DEFAULT NULL," +
                    "arb_winME varchar(50) DEFAULT NULL," +
                    "arb_winNT varchar(50) DEFAULT NULL," +
                    "arb_win2000 varchar(50) DEFAULT NULL," +
                    "arb_winXP varchar(50) DEFAULT NULL," +
                    "arb_mac8 varchar(50) DEFAULT NULL," +
                    "arb_mac9 varchar(50) DEFAULT NULL," +
                    "arb_mac10 varchar(50) DEFAULT NULL," +
                    "arb_dos varchar(50) DEFAULT NULL," +
                    "kopi_farve varchar(10) DEFAULT NULL," +
                    "kopi_sh varchar(10) DEFAULT NULL," +
                    "kopi_it varchar(10) DEFAULT NULL," +
                    "kopi_t1 varchar(10) DEFAULT NULL," +
                    "kopi_t2 varchar(10) DEFAULT NULL," +
                    "kopi_b1_a5 varchar(10) DEFAULT NULL," +
                    "kopi_b1_a4 varchar(10) DEFAULT NULL," +
                    "kopi_b1_a3 varchar(10) DEFAULT NULL," +
                    "kopi_b2_a5 varchar(10) DEFAULT NULL," +
                    "kopi_b2_a4 varchar(10) DEFAULT NULL," +
                    "kopi_b2_a3 varchar(10) DEFAULT NULL," +
                    "kopi_b3_a5 varchar(10) DEFAULT NULL," +
                    "kopi_b3_a4 varchar(10) DEFAULT NULL," +
                    "kopi_b3_a3 varchar(10) DEFAULT NULL," +
                    "faxdriver varchar(10) DEFAULT NULL," +
                    "faxnr varchar(10) DEFAULT NULL," +
                    "faxid varchar(10) DEFAULT NULL," +
                    "faxfarve varchar(10) DEFAULT NULL," +
                    "fax_papribk_1 varchar(10) DEFAULT NULL," +
                    "fax_papribk_2 varchar(10) DEFAULT NULL," +
                    "fax_papribk_3 varchar(10) DEFAULT NULL," +
                    "fax_papribk_4 varchar(10) DEFAULT NULL," +
                    "fax_it varchar(10) DEFAULT NULL," +
                    "fax_t1 varchar(10) DEFAULT NULL," +
                    "fax_t2 varchar(10) DEFAULT NULL," +
                    "fax_rap varchar(10) DEFAULT NULL," +
                    "print_plc5 varchar(10) DEFAULT NULL," +
                    "print_plc6 varchar(10) DEFAULT NULL," +
                    "print_psl2 varchar(10) DEFAULT NULL," +
                    "print_psl3 varchar(10) DEFAULT NULL," +
                    "print_farve varchar(10) DEFAULT NULL," +
                    "print_sh varchar(10) DEFAULT NULL," +
                    "print_it varchar(10) DEFAULT NULL," +
                    "print_t1 varchar(10) DEFAULT NULL," +
                    "print_t2 varchar(10) DEFAULT NULL," +
                    "print_admin varchar(10) DEFAULT NULL," +
                    "scan_email varchar(10) DEFAULT NULL," +
                    "scan_efile varchar(10) DEFAULT NULL," +
                    "scan_ftp varchar(10) DEFAULT NULL," +
                    "scan_esame varchar(10) DEFAULT NULL," +
                    "scan_twain varchar(10) DEFAULT NULL," +
                    "scan_filedownload varchar(10) DEFAULT NULL," +
                    "scan_topacc varchar(10) DEFAULT NULL," +
                    "scan_agfa varchar(10) DEFAULT NULL," +
                    "bemaerk text," +
                    "signoffdato date DEFAULT '0000-00-00'," +
                    "underskrift varchar(50) DEFAULT NULL," +
                    "pandet varchar(10) DEFAULT NULL);";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE sprog (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "sproglabel varchar(255) NOT NULL DEFAULT '');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE stempelur (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "editor varchar(50) DEFAULT NULL," +
                    "dato varchar(50) DEFAULT NULL," +
                    "navn varchar(50) DEFAULT NULL," +
                    "faktor double(12, 2) NOT NULL DEFAULT '1.00'," +
                    "forvalgt int(11) DEFAULT NULL," +
                    "minimum double(12, 2) NOT NULL DEFAULT '0.00');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE stopur (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "jobid int(11) NOT NULL DEFAULT '0'," +
                    "sttid datetime NOT NULL DEFAULT '2000-00-00 00:00:00'," +
                    "sltid datetime NOT NULL DEFAULT '2000-00-00 00:00:00'," +
                    "medid int(11) NOT NULL DEFAULT '0'," +
                    "incident int(11) NOT NULL DEFAULT '0'," +
                    "incidentlog int(11) NOT NULL DEFAULT '0'," +
                    "aktid int(11) NOT NULL DEFAULT '0'," +
                    "timereg_overfort int(11) NOT NULL DEFAULT '0'," +
                    "sid_godkendt int(11) NOT NULL DEFAULT '0'," +
                    "dato date NOT NULL DEFAULT '2002-01-01'," +
                    "editor varchar(255) DEFAULT NULL," +
                    "kommentar text);";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE tilbuds_skabeloner (ts_id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "ts_navn varchar(100) NOT NULL DEFAULT ''," +
                    "ts_txt longtext," +
                    "ts_dato date NOT NULL DEFAULT '0000-00-00'," +
                    "ts_editor varchar(100) NOT NULL DEFAULT ''," +
                    "ts_kundeid int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE timepriser (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "jobid int(11) NOT NULL DEFAULT '0'," +
                    "aktid int(11) NOT NULL DEFAULT '0'," +
                    "medarbid int(11) NOT NULL DEFAULT '0'," +
                    "timeprisalt int(11) NOT NULL DEFAULT '0'," +
                    "6timepris double(10, 2) DEFAULT '0.00'," +
                    "6valuta int(11) NOT NULL DEFAULT '1'," +
                    "KEY akt_inx(aktid)," +
                    "KEY medarb_aktid_jobid(medarbid, jobid, aktid)," +
                    "KEY jobid_tp_inx(jobid));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE timer (Tid int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "Tdato date DEFAULT NULL," +
                    "Tmnavn varchar(50) DEFAULT NULL," +
                    "Tmnr int(11) DEFAULT '0'," +
                    "tjobnavn varchar(150) NOT NULL DEFAULT ''," +
                    "tjobnr varchar(50) NOT NULL DEFAULT '0'," +
                    "Tknavn varchar(50) DEFAULT NULL," +
                    "Tknr int(11) DEFAULT '0'," +
                    "Timer double DEFAULT '0'," +
                    "Tfaktim int(11) DEFAULT '0'," +
                    "Timerkom text," +
                    "TAktivitetId int(11) DEFAULT '0'," +
                    "taktivitetnavn varchar(100) DEFAULT NULL," +
                    "Taar int(11) DEFAULT '0'," +
                    "TimePris double DEFAULT NULL," +
                    "TasteDato date DEFAULT NULL," +
                    "fastpris varchar(50) DEFAULT '0'," +
                    "tidspunkt time DEFAULT NULL," +
                    "editor varchar(50) DEFAULT NULL," +
                    "kostpris double DEFAULT '0'," +
                    "offentlig int(11) NOT NULL DEFAULT '0'," +
                    "seraft int(11) NOT NULL DEFAULT '0'," +
                    "godkendtstatus int(11) NOT NULL DEFAULT '0'," +
                    "godkendtstatusaf varchar(50) DEFAULT NULL," +
                    "sttid time DEFAULT NULL," +
                    "sltid time DEFAULT NULL," +
                    "valuta int(11) NOT NULL DEFAULT '1'," +
                    "kurs double(12, 4) NOT NULL DEFAULT '0.0000'," +
                    "bopal int(11) NOT NULL DEFAULT '0'," +
                    "destination varchar(50) DEFAULT NULL," +
                    "extSysId double(12, 0) NOT NULL DEFAULT '0'," +
                    "origin int(11) NOT NULL DEFAULT '0'," +
                    "timeralt double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "kpvaluta int(11) NOT NULL DEFAULT '1'," +
                    "kpvaluta_kurs double(12, 2) NOT NULL DEFAULT '100.00'," +
                    "godkendtdato date DEFAULT '2002-01-01'," +
                    "overfort int(11) NOT NULL DEFAULT '0'," +
                    "overfortdt date NOT NULL DEFAULT '2002-01-01'," +
                    "KEY timer_tknr(Tknr)," +
                    "KEY timer_tmnr(Tmnr)," +
                    "KEY timer_aktid_mnr_dato(TAktivitetId, Tmnr, Tdato, tjobnr)," +
                    "KEY timer_m_j_tfak(Tmnr, tjobnr, Tfaktim)," +
                    "KEY timer_tjobnr(tjobnr)," +
                    "KEY inx_timer_tregsum(Tdato, Tmnr, Tknr)," +
                    "KEY inx_timer_tregsum2(Tdato, Tmnr, tjobnr)," +
                    "KEY inx_tim_aktid(TAktivitetId));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE timer_imp_err (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "dato date DEFAULT NULL," +
                    "extsysid double(12, 0) NOT NULL DEFAULT '0'," +
                    "errId int(11) NOT NULL DEFAULT '0'," +
                    "jobnr varchar(50) DEFAULT NULL," +
                    "jobId double(12, 0) NOT NULL DEFAULT '0'," +
                    "med_init varchar(50) DEFAULT NULL," +
                    "timeregdato date DEFAULT NULL," +
                    "timer double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "origin int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE timer_import_temp (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "dato date DEFAULT NULL," +
                    "editor varchar(100) DEFAULT NULL," +
                    "origin int(11) NOT NULL DEFAULT '0'," +
                    "medarbejderid varchar(100) DEFAULT NULL," +
                    "jobid double(12, 0) NOT NULL DEFAULT '0'," +
                    "aktnavn varchar(100) DEFAULT NULL," +
                    "timer double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "tdato date DEFAULT NULL," +
                    "lto varchar(100) DEFAULT NULL," +
                    "overfort int(11) NOT NULL DEFAULT '0'," +
                    "timerkom varchar(250) DEFAULT NULL," +
                    "jobnavn varchar(250) DEFAULT NULL," +
                    "aktid double(12, 0) NOT NULL DEFAULT '0'," +
                    "errid int(11) NOT NULL DEFAULT '0'," +
                    "errmsg varchar(50) DEFAULT NULL);";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE timer_konsolideret (tk_id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "tk_dato date DEFAULT NULL," +
                    "tk_timerid double(12, 0) NOT NULL DEFAULT '0'," +
                    "tk_mnr double(12, 0) NOT NULL DEFAULT '0'," +
                    "tk_jobnr varchar(50) NOT NULL DEFAULT '0'," +
                    "tk_aid double(12, 0) NOT NULL DEFAULT '0'," +
                    "tk_timer double(12, 2) DEFAULT NULL," +
                    "dato_kons date DEFAULT NULL," +
                    "UNIQUE KEY tk_timerid_inx(tk_timerid));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE timer_konsolideret_tot (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "dato date DEFAULT NULL," +
                    "jobid int(11) NOT NULL DEFAULT '0'," +
                    "mtype int(11) NOT NULL DEFAULT '0'," +
                    "mtgid int(11) NOT NULL DEFAULT '0'," +
                    "timer double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "belob double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "kost double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "opd_dato timestamp(6) NULL DEFAULT NULL," +
                    "KEY inx_mj(mtgid, jobid)," +
                    "KEY inx_tojob(jobid));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE timereg_usejob (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "medarb int(11) NOT NULL DEFAULT '0'," +
                    "jobid int(11) NOT NULL DEFAULT '0'," +
                    "easyreg int(11) NOT NULL DEFAULT '0'," +
                    "forvalgt int(11) NOT NULL DEFAULT '0'," +
                    "forvalgt_sortorder int(11) NOT NULL DEFAULT '0'," +
                    "forvalgt_af int(11) NOT NULL DEFAULT '0'," +
                    "forvalgt_dt date DEFAULT NULL," +
                    "aktid double NOT NULL DEFAULT '0'," +
                    "favorit int(11) NOT NULL DEFAULT '0'," +
                    "KEY inx_medarb(medarb)," +
                    "KEY tu_inx(jobid, medarb)," +
                    "KEY inx_tu_mja(medarb, jobid, aktid));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE todo (id int(11) NOT NULL AUTO_INCREMENT  PRIMARY KEY," +
                    "editor varchar(50) DEFAULT NULL," +
                    "dato date NOT NULL DEFAULT '0000-00-00'," +
                    "startdato date NOT NULL DEFAULT '0000-00-00'," +
                    "slutdato date NOT NULL DEFAULT '0000-00-00'," +
                    "emne text NOT NULL," +
                    "beskrivelse text," +
                    "mid int(11) NOT NULL DEFAULT '0'," +
                    "udfort int(11) DEFAULT '0'," +
                    "sortorder int(4) NOT NULL DEFAULT '0'," +
                    "klassificering int(4) NOT NULL DEFAULT '0'," +
                    "kundeid int(4) NOT NULL DEFAULT '0'," +
                    "tidsstempel timestamp(6) NULL DEFAULT NULL);";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE todo_new (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "navn varchar(255) DEFAULT NULL," +
                    "parent int(11) NOT NULL DEFAULT '0'," +
                    "dato date NOT NULL DEFAULT '0000-00-00'," +
                    "level int(11) NOT NULL DEFAULT '0'," +
                    "delt int(11) DEFAULT '0'," +
                    "afsluttet int(11) NOT NULL DEFAULT '0'," +
                    "sortorder int(11) NOT NULL DEFAULT '1000'," +
                    "tododato date DEFAULT NULL," +
                    "forvafsl int(11) NOT NULL DEFAULT '0'," +
                    "public int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE todo_rel (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "medarbid int(11) NOT NULL DEFAULT '0'," +
                    "todoid int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE todo_rel_new (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "medarbid int(11) NOT NULL DEFAULT '0'," +
                    "todoid int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE travel_diet_tariff (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "tdf_diet_editor varchar(255) NOT NULL DEFAULT ''," +
                    "tdf_diet_dato date NOT NULL DEFAULT '2002-01-01'," +
                    "tdf_diet_name varchar(255) NOT NULL DEFAULT ''," +
                    "tdf_diet_current int(11) NOT NULL DEFAULT '0'," +
                    "tdf_diet_dayprice double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "tdf_diet_dayprice_half double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "tdf_diet_morgenamount double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "tdf_diet_middagamount double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "tdf_diet_aftenamount double(12, 2) NOT NULL DEFAULT '0.00');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE traveldietexp (diet_id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "diet_mid int(11) NOT NULL DEFAULT '0'," +
                    "diet_namedest varchar(250) NOT NULL DEFAULT ''," +
                    "diet_stdato datetime DEFAULT NULL," +
                    "diet_sldato datetime DEFAULT NULL," +
                    "diet_jobid int(11) NOT NULL DEFAULT '0'," +
                    "diet_konto int(11) NOT NULL DEFAULT '0'," +
                    "diet_dayprice double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "diet_maksamount double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "diet_morgen int(11) NOT NULL DEFAULT '0'," +
                    "diet_morgenamount double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "diet_middag int(11) NOT NULL DEFAULT '0'," +
                    "diet_middagamount double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "diet_aften int(11) NOT NULL DEFAULT '0'," +
                    "diet_aftenamount double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "diet_rest double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "diet_min25proc double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "diet_total double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "diet_bilag int(11) NOT NULL DEFAULT '0'," +
                    "diet_traveldays double(12, 2) DEFAULT NULL," +
                    "diet_dayprice_half double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "diet_delvis int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE traveldietexp_jobrel (tj_id int(11) NOT NULL AUTO_INCREMENT  PRIMARY KEY," +
                    "tj_trvlid int(11) NOT NULL DEFAULT '0'," +
                    "tj_jobid int(11) NOT NULL DEFAULT '0'," +
                    "tj_percent double(12, 2) NOT NULL DEFAULT '0.00');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE udeafhuset (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "medid int(11) NOT NULL," +
                    "fradato date NOT NULL," +
                    "fratidspunkt time DEFAULT NULL," +
                    "tiltidspunkt time DEFAULT NULL," +
                    "heledagen int(11) DEFAULT NULL," +
                    "tildato date NOT NULL," +
                    "fra datetime DEFAULT NULL," +
                    "til datetime DEFAULT NULL," +
                    "udeafhusettype int(11) DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE ugestatus (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "uge date NOT NULL DEFAULT '0000-00-00'," +
                    "afsluttet datetime NOT NULL DEFAULT '0000-00-00 00:00:00'," +
                    "mid int(11) NOT NULL DEFAULT '0'," +
                    "status int(11) NOT NULL DEFAULT '0'," +
                    "ugegodkendt int(11) NOT NULL DEFAULT '0'," +
                    "ugegodkendtaf int(11) NOT NULL DEFAULT '0'," +
                    "ugegodkendtTxt varchar(50) DEFAULT NULL," +
                    "ugegodkendtdt date DEFAULT NULL," +
                    "splithr int(11) NOT NULL DEFAULT '0'," +
                    "KEY inx_ugest(mid)," +
                    "KEY inx_ugest_mid_uge(mid, uge));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE valuta_historik (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "valutaid int(11) NOT NULL DEFAULT '0'," +
                    "kurs double NOT NULL DEFAULT '0'," +
                    "dato date NOT NULL DEFAULT '0000-00-00'," +
                    "editor varchar(255) NOT NULL DEFAULT '');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE valutaer (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "valuta varchar(255) NOT NULL DEFAULT ''," +
                    "valutakode varchar(255) NOT NULL DEFAULT ''," +
                    "kurs double NOT NULL DEFAULT '0'," +
                    "grundvaluta int(11) NOT NULL DEFAULT '0'," +
                    "editor varchar(255) NOT NULL DEFAULT ''," +
                    "dato date NOT NULL DEFAULT '0000-00-00');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE wip_historik (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "dato date DEFAULT NULL," +
                    "medid int(11) NOT NULL DEFAULT '0'," +
                    "jobid double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "restestimat double(12, 2) NOT NULL DEFAULT '0.00'," +
                    "stade_tim_proc int(11) NOT NULL DEFAULT '0');";
                command.ExecuteNonQueryAsync();


                command.CommandText = "CREATE TABLE xkunder (Kid int(11) NOT NULL AUTO_INCREMENT," +
                    "Kdato varchar(50) DEFAULT NULL," +
                    "Kkundenavn varchar(50) NOT NULL DEFAULT ''," +
                    "Kkundenr int(11) NOT NULL DEFAULT '0'," +
                    "editor varchar(50) DEFAULT NULL," +
                    "ketype varchar(50) DEFAULT '0'," +
                    "adresse varchar(50) DEFAULT NULL," +
                    "postnr varchar(11) DEFAULT NULL," +
                    "city varchar(50) DEFAULT NULL," +
                    "telefon varchar(20) DEFAULT NULL," +
                    "telefonmobil varchar(12) DEFAULT NULL," +
                    "telefonalt varchar(20) DEFAULT NULL," +
                    "fax varchar(12) DEFAULT NULL," +
                    "email varchar(50) DEFAULT NULL," +
                    "land varchar(50) DEFAULT NULL," +
                    "kontaktpers1 varchar(50) DEFAULT NULL," +
                    "kontaktpers2 varchar(50) DEFAULT NULL," +
                    "kontaktpers3 varchar(50) DEFAULT NULL," +
                    "kontaktpers4 varchar(50) DEFAULT NULL," +
                    "kontaktpers5 varchar(50) DEFAULT NULL," +
                    "kpersemail1 varchar(50) DEFAULT NULL," +
                    "kpersemail2 varchar(50) DEFAULT NULL," +
                    "kpersemail3 varchar(50) DEFAULT NULL," +
                    "kpersemail4 varchar(50) DEFAULT NULL," +
                    "kpersemail5 varchar(50) DEFAULT NULL," +
                    "kperstlf1 varchar(12) DEFAULT NULL," +
                    "kperstlf2 varchar(12) DEFAULT NULL," +
                    "kperstlf3 varchar(12) DEFAULT NULL," +
                    "kperstlf4 varchar(12) DEFAULT NULL," +
                    "kperstlf5 varchar(12) DEFAULT NULL," +
                    "kpersmobil1 varchar(12) DEFAULT NULL," +
                    "kpersmobil2 varchar(12) DEFAULT NULL," +
                    "kpersmobil3 varchar(12) DEFAULT NULL," +
                    "kpersmobil4 varchar(12) DEFAULT NULL," +
                    "kpersmobil5 varchar(12) DEFAULT NULL," +
                    "url varchar(100) DEFAULT NULL," +
                    "beskrivelse text," +
                    "regnr int(11) DEFAULT NULL," +
                    "kontonr int(11) DEFAULT NULL," +
                    "cvr int(11) DEFAULT NULL," +
                    "bank varchar(50) DEFAULT NULL," +
                    "pwkpers1 varchar(20) DEFAULT NULL," +
                    "pwkpers2 varchar(20) DEFAULT NULL," +
                    "pwkpers3 varchar(20) DEFAULT NULL," +
                    "pwkpers4 varchar(20) DEFAULT NULL," +
                    "pwkpers5 varchar(20) DEFAULT NULL," +
                    "lastloginkp1 varchar(50) DEFAULT NULL," +
                    "lastloginkp2 varchar(50) DEFAULT NULL," +
                    "lastloginkp3 varchar(50) DEFAULT NULL," +
                    "lastloginkp4 varchar(50) DEFAULT NULL," +
                    "lastloginkp5 varchar(50) DEFAULT NULL," +
                    "useasfak int(11) DEFAULT '0'," +
                    "logo int(11) DEFAULT NULL," +
                    "hot int(11) DEFAULT '0'," +
                    "swift varchar(8) DEFAULT NULL," +
                    "iban varchar(50) DEFAULT NULL," +
                    "oldid int(11) DEFAULT NULL," +
                    "PRIMARY KEY (Kid,Kkundenavn(10),Kkundenr)," +
                    "UNIQUE KEY kunderid_inx(Kid)," +
                    "KEY Kundenavn_inx(Kkundenavn)," +
                    "KEY Kundenr_inx(Kkundenr));";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE xtodo (id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "editor varchar(50) DEFAULT NULL," +
                    "dato date NOT NULL DEFAULT '0000-00-00'," +
                    "beskrivelse text," +
                    "mid int(11) NOT NULL DEFAULT '0'," +
                    "udfort int(11) DEFAULT '0');";
                command.ExecuteNonQueryAsync();

                command.CommandText = "CREATE TABLE your_rapports (rap_id int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY," +
                    "rap_mid int(11) NOT NULL," +
                    "rap_navn varchar(255) NOT NULL DEFAULT 'your_rapport'," +
                    "rap_url varchar(255) NOT NULL DEFAULT 'thisfile'," +
                    "rap_criteria text," +
                    "rap_dato date NOT NULL DEFAULT '2002-01-01'," +
                    "rap_editor varchar(255) DEFAULT NULL);";
                command.ExecuteNonQueryAsync();
                command.Connection.Close();

            }

        }

    }


}