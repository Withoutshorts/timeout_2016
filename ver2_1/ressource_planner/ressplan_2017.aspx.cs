using MySql.Data.MySqlClient;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Web.Services;

public partial class ressplan_2017 : System.Web.UI.Page
{
    #region Variable Declaration

    static string connString = Convert.ToString(System.Configuration.ConfigurationManager.ConnectionStrings["ConnectionString"]);

    #endregion

    #region Events

    /// <summary>
    /// Handles the Load event of the Page control.
    /// </summary>
    /// <param name="sender">The source of the event.</param>
    /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    protected void Page_Load(object sender, EventArgs e)
    {
        ////sqlquery = query;
        //lblsql.InnerHtml = "Hej 10";

        //string command = GetEmployeeQuery_SK("1,53,104", "2017-01-01", "2018-01-01", 3);
        //lblsql.InnerHtml = command;



    }

    #endregion

    #region Methods

    #region Dataset Methods

    /// <summary>
    /// Gets the data set by command.
    /// </summary>
    /// <param name="command">The command.</param>
    /// <returns>returns dataset by command</returns>
    private static DataSet GetDataSetByCommand(string command)
    {
        MySqlConnection sqlConnection = new MySqlConnection();
        MySqlCommand sqlCommand;

        try
        {
            sqlConnection = new MySqlConnection(connString);
            sqlConnection.Open();

            sqlCommand = sqlConnection.CreateCommand();
            sqlCommand.CommandText = command;

            MySqlDataAdapter sqlDataAdapter = new MySqlDataAdapter(sqlCommand);
            DataSet dsData = new DataSet();
            sqlDataAdapter.Fill(dsData);

            return dsData;
        }
        catch (Exception ex)
        {
            return null;
        }
        finally
        {
            if (sqlConnection != null && sqlConnection.State == System.Data.ConnectionState.Open)
            {

                sqlConnection.Close();
            }
        }

        
    }

    /// <summary>
    /// Gets the execute non query by command.
    /// </summary>
    /// <param name="command">The command.</param>
    /// <returns>return value by execute non query command</returns>
    private static int GetExecuteNonQueryByCommand(string command)
    {
        MySqlConnection sqlConnection = new MySqlConnection();
        MySqlCommand sqlCommand;

        try
        {
            sqlConnection = new MySqlConnection(connString);
            sqlConnection.Open();

            sqlCommand = sqlConnection.CreateCommand();
            sqlCommand.CommandText = command;

            return sqlCommand.ExecuteNonQuery();
        }
        catch (Exception ex)
        {
            return 0;
        }
        finally
        {
            if (sqlConnection != null && sqlConnection.State == System.Data.ConnectionState.Open)
            {
                sqlConnection.Close();
            }
        }
    }

    #endregion

    /// <summary>
    /// Gets the sort by type query.
    /// </summary>
    /// <param name="sortby">The sortby.</param>
    /// <returns>returns view type query</returns>
    private static string GetSortByTypeQuery(int sortby)
    {
        string query = string.Empty;

        // When 1st dropdown is selected
        if (sortby > 0)
        {
            if (sortby == 1) // Customer
            {
                //sqlCommand.CommandText = "SELECT DISTINCT(kunder.kid), kunder.kkundenavn FROM aktiviteter LEFT JOIN job ON job.`id` = aktiviteter.`job` LEFT JOIN kunder ON kunder.`Kid` = job.`jobknr` WHERE kunder.kid IS NOT NULL ORDER BY kunder.kkundenavn ";
                //query = "SELECT DISTINCT(kunder.kid), kunder.kkundenavn, projektgrupper.navn FROM aktiviteter LEFT JOIN medarbejdere ON aktiviteter.id = medarbejdere.MID LEFT JOIN progrupperelationer ON medarbejdere.MID = progrupperelationer.medarbejderid LEFT JOIN projektgrupper ON progrupperelationer.projektgruppeid = projektgrupper.id LEFT JOIN job ON job.`id` = aktiviteter.`job` LEFT JOIN kunder ON kunder.`Kid` = job.`jobknr` WHERE kunder.kid IS NOT NULL ORDER BY kunder.kkundenavn";
                query = "SELECT Kid, kkundenavn FROM kunder";
            }
            else if (sortby == 2) // job
            {
                //query = "SELECT DISTINCT(job.id), job.jobnavn  FROM aktiviteter LEFT JOIN job ON job.`id` = aktiviteter.`job` LEFT JOIN kunder ON kunder.`Kid` = job.`jobknr` WHERE job.id IS NOT NULL ORDER BY job.jobnavn";
                //query = "SELECT DISTINCT(job.id), job.jobnavn  FROM aktiviteter LEFT JOIN job ON job.`id`  = aktiviteter.`job` LEFT JOIN kunder ON kunder.`kid` = job.`jobknr` WHERE job.id IS NOT NULL";
                query = "SELECT DISTINCT(job.id), job.jobnavn, job.jobknr, kunder.kkundenavn FROM job LEFT JOIN kunder ON kunder.`Kid` = job.`jobknr` WHERE job.id IS NOT NULL AND jobstatus = 1 AND risiko > 0 ORDER BY kunder.kkundenavn";
            }
            else if (sortby == 3) // Group
            {
                //sqlCommand.CommandText = "SELECT projektgrupper.id,projektgrupper.navn FROM projektgrupper LEFT JOIN progrupperelationer ON (progrupperelationer.projektgruppeid = projektgrupper.id) LEFT JOIN medarbejdere ON(mid = progrupperelationer.medarbejderid) ORDER BY projektgrupper.navn";
                //query = "SELECT projektgrupper.id,projektgrupper.navn FROM projektgrupper LEFT JOIN progrupperelationer ON(progrupperelationer.projektgruppeid = projektgrupper.id) LEFT JOIN medarbejdere ON (MID = progrupperelationer.medarbejderid) LEFT JOIN aktiviteter ON medarbejdere.MID = aktiviteter.pl_employee ORDER BY projektgrupper.navn";
                query = "SELECT id,navn FROM projektgrupper";
            }
            else if (sortby == 4) // Employee
            {
                query = "SELECT mid, mnavn, init FROM medarbejdere WHERE mansat = 1 ORDER BY mnavn";
            }
        }
        return query;
    }

    /// <summary>
    /// Gets the customers query.
    /// </summary>
    /// <param name="customerIds">The group ids.</param>
    /// <param name="startDate">The start date.</param>
    /// <param name="endDate">The end date.</param>
    /// <returns>
    /// returns query for customers
    /// </returns>
    private static string GetCustomersQuery(List<string> customerIds, string startDate, string endDate, int viewType)
    {
        
        string query = string.Empty;
        string whereClause = string.Empty;
        string orderByClause = string.Empty;
        string strIds = GetStringFromList(customerIds);
        whereClause = "WHERE kunder.Kid IN (" + strIds + ")";
        whereClause += " AND (date(akt_bookings.ab_enddate) >= '" + startDate + "' AND date(akt_bookings.ab_startdate) <= '" + endDate + "') ";
        whereClause += GetViewTypeQuery(viewType);

        orderByClause = "ORDER BY CustomerId, JobId, ActivityName, EmployeeName";

        query = "SELECT akt_bookings.ab_id AS BookingId, " +
                        "aktiviteter.id AS ActivityId, " +
                        "aktiviteter.navn AS ActivityName, " +
                        "job.id AS JobId, " +
                        "job.jobnavn AS JobName, " +
                        "kunder.Kid AS CustomerId, " +
                        "kunder.kkundenavn AS CustomerName, " +
                        "medarbejdere.Mid AS EmployeeId, " +
                        "medarbejdere.Mnavn AS EmployeeName," +
                        "akt_bookings.ab_startdate AS StartDate, " +
                        "akt_bookings.ab_enddate AS EndDate, " +
                        "akt_bookings.ab_important AS IsImportant, " +
                        "akt_bookings.ab_serie AS Recurrence, " +
                        "akt_bookings.ab_end_after AS NoOfRecurrence, " +
                        "aktiviteter.fakturerbar AS ActivityType " +
                "FROM akt_bookings " +
                "LEFT JOIN aktiviteter ON akt_bookings.ab_aktid = aktiviteter.id " +
                "LEFT JOIN akt_bookings_rel ON akt_bookings.ab_id = akt_bookings_rel.abl_bookid " +
                "LEFT JOIN job ON job.id = akt_bookings.ab_jobid " +
                "LEFT JOIN kunder ON kunder.Kid = job.jobknr " +
                "LEFT JOIN medarbejdere ON medarbejdere.Mid = akt_bookings_rel.abl_medid ";
        query += whereClause + orderByClause;
        return query;
    }

    /// <summary>
    /// Gets the job query.
    /// </summary>
    /// <param name="jobIds">The job ids.</param>
    /// <param name="startDate">The start date.</param>
    /// <param name="endDate">The end date.</param>
    /// <returns>
    /// returns query for job
    /// </returns>
    /// private static string GetJobQuery(List<string> jobIds, string startDate, string endDate, int viewType)
    private static string GetJobQuery_SK(string jobIds, string startDate, string endDate, int viewType)
    {
        string query = string.Empty;
        string whereClause = string.Empty;
        string orderByClause = string.Empty;
        //string strIds = GetStringFromList(jobIds);
        string strIds = jobIds;
        whereClause = "WHERE job.id IN (" + strIds + ")";
        whereClause += " AND (date(akt_bookings.ab_enddate) >= '" + startDate + "' AND date(akt_bookings.ab_startdate) <= '" + endDate + "') ";
        whereClause += GetViewTypeQuery(viewType);

        orderByClause = "ORDER BY CustomerId, JobId, ActivityName, EmployeeName";

        query = "SELECT akt_bookings.ab_id AS BookingId, " +
                        "aktiviteter.id AS ActivityId, " +
                        "aktiviteter.navn AS ActivityName, " +
                        "job.id AS JobId, " +
                        "job.jobnavn AS JobName, " +
                        "kunder.Kid AS CustomerId, " +
                        "kunder.kkundenavn AS CustomerName, " +
                        "medarbejdere.Mid AS EmployeeId, " +
                        "medarbejdere.Mnavn AS EmployeeName," +
                        "akt_bookings.ab_startdate AS StartDate, " +
                        "akt_bookings.ab_enddate AS EndDate, " +
                        "akt_bookings.ab_important AS IsImportant, " +
                        "akt_bookings.ab_serie AS Recurrence, " +
                        "akt_bookings.ab_end_after AS NoOfRecurrence, " +
                        "aktiviteter.fakturerbar AS ActivityType " +
                "FROM akt_bookings " +
                "LEFT JOIN aktiviteter ON akt_bookings.ab_aktid = aktiviteter.id " +
                "LEFT JOIN akt_bookings_rel ON akt_bookings.ab_id = akt_bookings_rel.abl_bookid " +
                "LEFT JOIN job ON job.id = akt_bookings.ab_jobid " +
                "LEFT JOIN kunder ON kunder.Kid = job.jobknr " +
                "LEFT JOIN medarbejdere ON medarbejdere.Mid = akt_bookings_rel.abl_medid ";
        query += whereClause + orderByClause;
        return query;
    }

    /// <summary>
    /// Gets the job query.
    /// </summary>
    /// <param name="jobIds">The job ids.</param>
    /// <param name="startDate">The start date.</param>
    /// <param name="endDate">The end date.</param>
    /// <returns>
    /// returns query for job
    /// </returns>
    private static string GetJobQuery(List<string> jobIds, string startDate, string endDate, int viewType)
    {
        string query = string.Empty;
        string whereClause = string.Empty;
        string orderByClause = string.Empty;
        string strIds = GetStringFromList(jobIds);
        
        whereClause = "WHERE job.id IN (" + strIds + ")";
        whereClause += " AND (date(akt_bookings.ab_enddate) >= '" + startDate + "' AND date(akt_bookings.ab_startdate) <= '" + endDate + "') ";
        whereClause += GetViewTypeQuery(viewType);

        orderByClause = "ORDER BY CustomerId, JobId, ActivityName, EmployeeName";

        query = "SELECT akt_bookings.ab_id AS BookingId, " +
                        "aktiviteter.id AS ActivityId, " +
                        "aktiviteter.navn AS ActivityName, " +
                        "job.id AS JobId, " +
                        "job.jobnavn AS JobName, " +
                        "kunder.Kid AS CustomerId, " +
                        "kunder.kkundenavn AS CustomerName, " +
                        "medarbejdere.Mid AS EmployeeId, " +
                        "medarbejdere.Mnavn AS EmployeeName," +
                        "akt_bookings.ab_startdate AS StartDate, " +
                        "akt_bookings.ab_enddate AS EndDate, " +
                        "akt_bookings.ab_important AS IsImportant, " +
                        "akt_bookings.ab_serie AS Recurrence, " +
                        "akt_bookings.ab_end_after AS NoOfRecurrence, " +
                        "aktiviteter.fakturerbar AS ActivityType " +
                "FROM akt_bookings " +
                "LEFT JOIN aktiviteter ON akt_bookings.ab_aktid = aktiviteter.id " +
                "LEFT JOIN akt_bookings_rel ON akt_bookings.ab_id = akt_bookings_rel.abl_bookid " +
                "LEFT JOIN job ON job.id = akt_bookings.ab_jobid " +
                "LEFT JOIN kunder ON kunder.Kid = job.jobknr " +
                "LEFT JOIN medarbejdere ON medarbejdere.Mid = akt_bookings_rel.abl_medid ";
        query += whereClause + orderByClause;
        return query;
    }

    /// <summary>
    /// Gets the project name and employee query.
    /// </summary>
    /// <param name="groupIds">The group ids.</param>
    /// <param name="startDate">The start date.</param>
    /// <param name="endDate">The end date.</param>
    /// <returns>
    /// returns query for fetching project name and employee
    /// </returns>
    private static string GetGroupQuery(List<string> groupIds, string startDate, string endDate, int viewType)
    {
        string query = string.Empty;
        string whereClause = string.Empty;
        string orderByClause = string.Empty;
        string strIds = GetStringFromList(groupIds);
        whereClause = "WHERE progrupperelationer.ProjektgruppeId IN (" + strIds + ") ";
        whereClause += "AND (date(akt_bookings.ab_enddate) >= '" + startDate + "' AND date(akt_bookings.ab_startdate) <= '" + endDate + "') ";
        whereClause += GetViewTypeQuery(viewType);

        orderByClause = "ORDER BY ProjectGroupId, EmployeeName";

        query = "SELECT medarbejdere.Mid AS EmployeeId, " +
                        "medarbejdere.Mnavn AS EmployeeName, " +
                        "progrupperelationer.ProjektgruppeId AS ProjectGroupId, " +
                        "projektgrupper.navn AS ProjectName, " +
                        "akt_bookings.ab_id AS BookingId, " +
                        "aktiviteter.id AS ActivityId, " +
                        "aktiviteter.navn AS ActivityName," +
                        "akt_bookings.ab_startdate AS StartDate, " +
                        "akt_bookings.ab_enddate AS EndDate, " +
                        "akt_bookings.ab_important AS IsImportant, " +
                        "akt_bookings.ab_serie AS Recurrence, " +
                        "akt_bookings.ab_end_after AS NoOfRecurrence, " +
                        "aktiviteter.fakturerbar AS ActivityType, " +
                        "aktiviteter.pl_header AS ActivityHeader, " +
                        "job.jobnavn AS JobName, " +
                        "kunder.kkundenavn AS CustomerName " +
                "FROM medarbejdere " +
                "LEFT JOIN progrupperelationer ON progrupperelationer.MedarbejderId = medarbejdere.Mid " +
                "LEFT JOIN projektgrupper ON projektgrupper.id = progrupperelationer.ProjektgruppeId " +
                "LEFT JOIN akt_bookings_rel ON akt_bookings_rel.abl_medid = medarbejdere.Mid " +
                "LEFT JOIN akt_bookings ON akt_bookings.ab_id = akt_bookings_rel.abl_bookid " +
                "LEFT JOIN aktiviteter ON aktiviteter.id = akt_bookings.ab_aktid " +
                "LEFT JOIN job ON job.id = akt_bookings.ab_jobid " +
                "LEFT JOIN kunder ON kunder.Kid = job.jobknr " +
                whereClause + orderByClause + "; ";
        return query;
    }


    private static string GetEmployeeQuery_SK(string empids, string startDate, string endDate, int viewType)
    {
        string query = string.Empty;
        string whereClause = string.Empty;
        string orderByClause = string.Empty;
        //string strIds = GetStringFromList(employeeIds);
        string strIds = "1, 104, 53";
        whereClause = "WHERE medarbejdere.Mid IN (" + strIds + ")";
        whereClause += " AND (date(akt_bookings.ab_enddate) >= '" + startDate + "' AND date(akt_bookings.ab_startdate) <= '" + endDate + "') ";
        whereClause += GetViewTypeQuery(viewType);

        orderByClause = "ORDER BY EmployeeName";

        query = "SELECT medarbejdere.Mid AS EmployeeId, " +
                        "medarbejdere.Mnavn AS EmployeeName, " +
                        "akt_bookings.ab_id AS BookingId, " +
                        "aktiviteter.id AS ActivityId, " +
                        "aktiviteter.navn AS ActivityName, " +
                        "akt_bookings.ab_startdate AS StartDate, " +
                        "akt_bookings.ab_enddate AS EndDate, " +
                        "akt_bookings.ab_important AS IsImportant, " +
                        "akt_bookings.ab_serie AS Recurrence, " +
                        "akt_bookings.ab_end_after AS NoOfRecurrence, " +
                        "aktiviteter.fakturerbar AS ActivityType, " +
                        "aktiviteter.pl_header AS ActivityHeader, " +
                        "job.jobnavn AS JobName, " +
                        "kunder.kkundenavn AS CustomerName " +
                "FROM medarbejdere " +
                "LEFT JOIN akt_bookings_rel ON akt_bookings_rel.abl_medid = medarbejdere.Mid " +
                "LEFT JOIN akt_bookings ON akt_bookings.ab_id = akt_bookings_rel.abl_bookid  " +
                "LEFT JOIN aktiviteter ON aktiviteter.id = akt_bookings.ab_aktid " +
                "LEFT JOIN job ON job.id = akt_bookings.ab_jobid " +
                "LEFT JOIN kunder ON kunder.Kid = job.jobknr " +
                whereClause + orderByClause + ";";
        return query;
    }

    /// <summary>
    /// Gets the employee query.
    /// </summary>
    /// <param name="employeeIds">The employee ids.</param>
    /// <param name="startDate">The start date.</param>
    /// <param name="endDate">The end date.</param>
    /// <returns>returns employee data query</returns>
    private static string GetEmployeeQuery(List<string> employeeIds, string startDate, string endDate, int viewType)
    {
        string query = string.Empty;
        string whereClause = string.Empty;
        string orderByClause = string.Empty;
        string strIds = GetStringFromList(employeeIds);
        whereClause = "WHERE medarbejdere.Mid IN (" + strIds + ")";
        whereClause += " AND (date(akt_bookings.ab_enddate) >= '" + startDate + "' AND date(akt_bookings.ab_startdate) <= '" + endDate + "') ";
        whereClause += GetViewTypeQuery(viewType);

        orderByClause = "ORDER BY EmployeeName";

        query = "SELECT medarbejdere.Mid AS EmployeeId, " +
                        "medarbejdere.Mnavn AS EmployeeName, " +
                        "akt_bookings.ab_id AS BookingId, " +
                        "aktiviteter.id AS ActivityId, " +
                        "aktiviteter.navn AS ActivityName, " +
                        "akt_bookings.ab_startdate AS StartDate, " +
                        "akt_bookings.ab_enddate AS EndDate, " +
                        "akt_bookings.ab_important AS IsImportant, " +
                        "akt_bookings.ab_serie AS Recurrence, " +
                        "akt_bookings.ab_end_after AS NoOfRecurrence, " +
                        "aktiviteter.fakturerbar AS ActivityType, " +
                        "aktiviteter.pl_header AS ActivityHeader, " +
                        "job.jobnavn AS JobName, " +
                        "kunder.kkundenavn AS CustomerName " +
                "FROM medarbejdere " +
                "LEFT JOIN akt_bookings_rel ON akt_bookings_rel.abl_medid = medarbejdere.Mid " +
                "LEFT JOIN akt_bookings ON akt_bookings.ab_id = akt_bookings_rel.abl_bookid  " +
                "LEFT JOIN aktiviteter ON aktiviteter.id = akt_bookings.ab_aktid " +
                "LEFT JOIN job ON job.id = akt_bookings.ab_jobid " +
                "LEFT JOIN kunder ON kunder.Kid = job.jobknr " +
                whereClause + orderByClause + ";";
        return query;
    }

    /// <summary>
    /// Gets the activity model.
    /// </summary>
    /// <returns></returns>
    private static string GetEditActivityQuery(int bookingId)
    {
        string query = string.Empty;
        query = "SELECT aktiviteter.id AS ActivityId, " +
                        "aktiviteter.navn AS ActivityName, " +
                        "akt_bookings.ab_id AS BookingId, " +
                        "akt_bookings.ab_startdate AS StartDate, " +
                        "akt_bookings.ab_enddate AS EndDate, " +
                        "akt_bookings.ab_important AS IsImportant, " +
                        "akt_bookings.ab_serie AS Recurrence, " +
                        "akt_bookings.ab_end_after AS NoOfRecurrence, " +
                        "aktiviteter.fakturerbar AS ActivityType, " +
                        "aktiviteter.pl_header AS ActivityHeader, " +
                        "job.jobnavn AS JobName, " +
                        "kunder.kkundenavn AS CustomerName, " +
                        "akt_bookings_rel.abl_medid AS EmployeeId, " +
                        "medarbejdere.Mnavn AS EmployeeName " +
                "FROM akt_bookings " +
                "LEFT JOIN aktiviteter ON aktiviteter.id = akt_bookings.ab_aktid " +
                "LEFT JOIN akt_bookings_rel ON akt_bookings_rel.abl_bookid = akt_bookings.ab_id " +
                "LEFT JOIN job ON (job.id = aktiviteter.job) " +
                "LEFT JOIN kunder ON (kunder.kid = job.jobknr) " +
                "LEFT JOIN medarbejdere ON medarbejdere.Mid = akt_bookings_rel.abl_medid " +
                "WHERE akt_bookings.ab_id = " + bookingId;
        return query;
    }

    /// <summary>
    /// Saves the task query.
    /// </summary>
    /// <param name="taskObject">The task object.</param>
    /// <returns>returns save task query</returns>
    private static string SaveActivityQuery(ActivityModel model)
    {
        
        string query = string.Empty;
       

        query = "UPDATE akt_bookings SET ab_editor = '"+ model.Employees + "', ab_startdate = '" + model.StartDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "', ab_enddate = '" + model.EndDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "', ab_serie = 0 WHERE ab_id = " + model.BookingId + ";";
        query += "DELETE FROM akt_bookings_rel WHERE abl_bookid = " + model.BookingId + "";
        //query += ";INSERT INTO akt_bookings_rel (abl_bookid, abl_medid) VALUES (" + model.BookingId + ", 53)";


        foreach (string employeeId in model.Employees)
        {
        query += ";INSERT INTO akt_bookings_rel (abl_bookid, abl_medid) VALUES ("+ model.BookingId + ", "+ employeeId + ")";
        }

        // + model.ActivityId;
        // ";

        //query = "SET ab_startdate = " + "'" + model.StartDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "'" + ", " +
        //" ab_enddate = " + "'" + model.EndDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "'" +
        //" WHERE ab_id = " + model.BookingId + "; ";

        //"ab_serie = " + model.Recurrence + ", " +
        //"ab_end_after = " + model.NoOfRecurrence + ", " +
        //"ab_important = " + model.Important + " " +
        //"WHERE ab_id = " + model.BookingId + "; " +

        //"UPDATE aktiviteter " +
        //"SET pl_header = " + model.Heading + " " +
        //"WHERE ab_id = " + model.ActivityId + "; " +


        //foreach(string employeeId in model.Employees)
        //{
        //query += "INSERT INTO akt_bookings_rel (abl_bookid, abl_medid) VALUES ("+ model.BookingId + ", "+ employeeId + ");";
        //}

        return query;
    }

    /// <summary>
    /// Gets all employee query.
    /// </summary>
    /// <returns>returns query of all employees</returns>
    private static string GetAllEmployeeQuery()
    {
        string query = string.Empty;
        query = "SELECT mid, mnavn FROM medarbejdere WHERE mansat = 1 ORDER BY mnavn;";
        return query;
    }

    /// <summary>
    /// Gets the activity model.
    /// </summary>
    /// <param name="drUser">The dr user.</param>
    /// <returns>returns activity model</returns>
    private static ActivityModel GetActivityModel(DataRow drUser)
    {
        ActivityModel activity = new ActivityModel();
        activity.BookingId = Convert.ToInt32(drUser["BookingId"]);
        //activity.ActivityId = Convert.ToInt32(drUser["ActivityId"]);
        activity.ActivityId = Convert.ToInt32(drUser["ActivityId"]);
        activity.ActivityName = Convert.ToString(drUser["ActivityName"]);
        activity.StartDateTime = Convert.ToDateTime(drUser["StartDate"]);
        activity.EndDateTime = Convert.ToDateTime(drUser["EndDate"]);
        activity.Important = Convert.ToInt32(drUser["IsImportant"]);
        activity.Recurrence = Convert.ToInt32(drUser["Recurrence"]);
        activity.NoOfRecurrence = Convert.ToInt32(drUser["NoOfRecurrence"]);
        activity.ActivityType = Convert.ToInt32(drUser["ActivityType"]);

        return activity;
    }

    /// <summary>
    /// Gets the employee model.
    /// </summary>
    /// <param name="drUser">The dr user.</param>
    /// <returns>returns employee model</returns>
    private static Employee GetEmployeeModel(DataRow drUser)
    {
        Employee employee = new Employee();
        employee.Id = Convert.ToInt32(drUser["EmployeeId"]);
        employee.Name = Convert.ToString(drUser["EmployeeName"]);

        return employee;
    }

    /// <summary>
    /// Gets the string from list.
    /// </summary>
    /// <param name="strList">The string list.</param>
    /// <returns>returns comma separated string from list</returns>
    private static string GetStringFromList(List<string> strList)
    {
        string strReturnVal = string.Empty;
        foreach (string data in strList)
        {
            strReturnVal += data + ",";
        }
        strReturnVal = strReturnVal.TrimEnd(',');
        return strReturnVal;
    }

    /// <summary>
    /// Gets the view type query.
    /// </summary>
    /// <param name="viewType">Type of the view.</param>
    /// <returns>returns view type related query</returns>
    private static string GetViewTypeQuery(int viewType)
    {
        string query = string.Empty;

        switch (viewType)
        {
            case 1:
                {
                    //query = " AND aktiviteter.fakturerbar = 1 ";
                }
                break;
            case 2:
                {
                    //query = " AND (aktiviteter.fakturerbar = 7 OR aktiviteter.fakturerbar = 8 OR aktiviteter.fakturerbar = 13 OR aktiviteter.fakturerbar = 14 OR aktiviteter.fakturerbar = 20 OR aktiviteter.fakturerbar = 21 OR aktiviteter.fakturerbar = 22 OR aktiviteter.fakturerbar = 23 OR aktiviteter.fakturerbar = 25 OR aktiviteter.fakturerbar = 31 OR aktiviteter.fakturerbar = 81 OR aktiviteter.fakturerbar = 115 OR aktiviteter.fakturerbar = 125) ";
                }
                break;
            case 3:
                {
                    //query = " AND ab_important = 0 ";
                    //query = " AND aktiviteter.pl_important_act  = 1 ";
                }
                break;
        }

        return query;
    }

    /// <summary>
    /// Gets the customer job list.
    /// </summary>
    /// <param name="dsData">The ds data.</param>
    /// <returns>returns job list for customer and job view</returns>
    private static List<Job> GetCustomerJobList(DataSet dsData)
    {
        List<Job> jobList = new List<Job>();
        Job job;

        if (dsData != null && dsData.Tables.Count > 0 && dsData.Tables[0].Rows.Count > 0)
        {
            DataRow drUsers;
            ActivityModel activity;
            bool isJobExists = false;

            for (int i = 0; i < dsData.Tables[0].Rows.Count; i++)
            {
                drUsers = dsData.Tables[0].Rows[i];
                bool isActivityExists = false;

                Job existingJob = jobList.Find(delegate (Job j)
                {
                    return j.CustomerId == Convert.ToInt32(drUsers["CustomerId"]) && j.JobId == Convert.ToInt32(drUsers["JobId"]);
                });

                if (existingJob == null)
                {
                    job = new Job();
                    job.JobId = Convert.ToInt32(drUsers["JobId"]);
                    job.CustomerId = Convert.ToInt32(drUsers["CustomerId"]);
                    job.JobName = drUsers["JobName"].ToString();
                    job.CustomerName = drUsers["CustomerName"].ToString();
                    job.Activities = new List<ActivityModel>();

                    existingJob = job;
                    isJobExists = false;
                }
                else
                {
                    isJobExists = true;
                }

                if (isJobExists)
                {
                    ActivityModel existingActivity = existingJob.Activities.Find(delegate (ActivityModel j)
                    {
                        return j.ActivityId == Convert.ToInt32(drUsers["ActivityId"]);
                    });

                    if (existingActivity == null)
                    {
                        isActivityExists = false;
                    }
                    else
                    {
                        isActivityExists = true;
                        existingActivity.EmployeeList.Add(GetEmployeeModel(drUsers));
                    }
                }

                if (!isActivityExists)
                {
                    activity = new ActivityModel();
                    activity = GetActivityModel(drUsers);
                    activity.EmployeeList = new List<Employee>();
                    activity.EmployeeList.Add(GetEmployeeModel(drUsers));
                    existingJob.Activities.Add(activity);
                }

                if (!isJobExists)
                {
                    jobList.Add(existingJob);
                }
            }
        }

        return jobList;
    }

    /// <summary>
    /// Gets the holidays query.
    /// </summary>
    /// <returns>returns holidays query</returns>
    private static string GetHolidaysQuery()
    {
        string query = string.Empty;
        query = "SELECT nh_id AS HolidayId, nh_date AS HolidayDate FROM national_holidays;";
        return query;
    }

    #endregion

    #region Web Method

    #region Dropdown binding methods

    /// <summary>
    /// Views the type selected.
    /// </summary>
    /// <param name="sortby">The sortby.</param>
    /// <param name="viewtype">The viewtype.</param>
    /// <returns>returns 3rd dropdown options</returns>
    [WebMethod]
    public static string SortByTypeSelected(int sortby)
    {
        string oliverDD = string.Empty;
        string where = string.Empty;

        try
        {
            string command = GetSortByTypeQuery(sortby);
            DataSet dsData = GetDataSetByCommand(command);

            if (dsData != null && dsData.Tables.Count > 0 && dsData.Tables[0].Rows.Count > 0)
            {
                DataRow drUsers;
                string ddText = string.Empty;
                int value = 1;
                //oliverDD += "<option value='0'>Select</option>";

                for (int i = 0; i < dsData.Tables[0].Rows.Count; i++)
                {
                    drUsers = dsData.Tables[0].Rows[i];

                    if (sortby == 1)
                    {
                        //ddText = Convert.ToString(drUsers["kkundenavn"] + " " + drUsers["navn"]);
                        ddText = Convert.ToString(drUsers["kkundenavn"]);
                        value = Convert.ToInt32(drUsers["kid"]);
                    }
                    if (sortby == 2)
                    {
                        ddText = Convert.ToString(drUsers["jobnavn"]);
                        value = Convert.ToInt32(drUsers["id"]);
                    }
                    else if (sortby == 3)
                    {
                        ddText = Convert.ToString(drUsers["navn"]);
                        value = Convert.ToInt32(drUsers["id"]);
                    }
                    else if (sortby == 4)
                    {
                        ddText = Convert.ToString(drUsers["mnavn"]);
                        value = Convert.ToInt32(drUsers["mid"]);
                    }

                    if (sortby == 2)
                    {
                        ddText += " (" + Convert.ToString(drUsers["id"]) + ")";

                        string compare = "'" + Convert.ToString(drUsers["kkundenavn"]) + " " + Convert.ToString(drUsers["jobknr"]) + "'";
                        if (oliverDD.IndexOf(compare) > -1)
                        {
                            int lastIndex = oliverDD.LastIndexOf("</optgroup>");
                            oliverDD = oliverDD.Insert(lastIndex, compare + "<option value='" + value + "'>" + ddText + "</option>");
                        }
                        else
                        {
                            string optGroupLabel = Convert.ToString(drUsers["kkundenavn"]) + " " + Convert.ToString(drUsers["jobknr"]);
                            oliverDD += "<optgroup label='" + optGroupLabel + "'>" + "<option value='" + value + "'>" + ddText + "</option>" + "</optgroup>";
                        }
                    }
                    else
                    {
                        oliverDD += "<option value='" + value + "'>" + ddText + "</option>";
                    }

                }
            }

            return oliverDD;
        }
        catch (Exception ex)
        {
            return "0";
        }
    }

    /// <summary>
    /// Gets all employees.
    /// </summary>
    /// <returns>returns all employees</returns>
    [WebMethod]
    public static string GetAllEmployees()
    {
        string employees = string.Empty;
        try
        {
            string command = GetAllEmployeeQuery();
            DataSet dsData = GetDataSetByCommand(command);

            if (dsData != null && dsData.Tables.Count > 0)
            {
                DataRow drEmployee;

                for (int i = 0; i < dsData.Tables[0].Rows.Count; i++)
                {
                    drEmployee = dsData.Tables[0].Rows[i];
                    employees += "<option value='" + Convert.ToInt32(drEmployee["mid"]) + "'>" + Convert.ToString(drEmployee["mnavn"]) + "</option>";
                }
            }
        }
        catch (Exception ex)
        {
        }

        return employees;
    }

    /// <summary>
    /// Gets the holidays.
    /// </summary>
    /// <returns>returns list of holiday</returns>
    [WebMethod]
    public static List<Holiday> GetHolidays()
    {
        List<Holiday> holidayList = new List<Holiday>();
        try
        {
            string command = GetHolidaysQuery();
            DataSet dsData = GetDataSetByCommand(command);

            if (dsData != null && dsData.Tables.Count > 0)
            {
                DataRow drHoliday;
                Holiday holiday;

                for (int i = 0; i < dsData.Tables[0].Rows.Count; i++)
                {
                    drHoliday = dsData.Tables[0].Rows[i];

                    holiday = new Holiday();
                    holiday.HolidayId = Convert.ToInt32(drHoliday["HolidayId"]);
                    holiday.HolidayDate = Convert.ToDateTime(drHoliday["HolidayDate"]);

                    holidayList.Add(holiday);
                }
            }
        }
        catch (Exception ex)
        {
        }

        return holidayList;
    }

    #endregion

    #region Layout binding methods

    /// <summary>
    /// Gets the customer layout data.
    /// </summary>
    /// <param name="customerIds">The customer ids.</param>
    /// <param name="startDate">The start date.</param>
    /// <param name="endDate">The end date.</param>
    /// <returns>
    /// returns customer layout data
    /// </returns>
    [WebMethod]
    public static List<Job> GetCustomerLayoutData(List<string> customerIds, string startDate, string endDate, int viewType)
    {
        try
        {
            string command = GetCustomersQuery(customerIds, startDate, endDate, viewType);
            DataSet dsData = GetDataSetByCommand(command);
            return GetCustomerJobList(dsData);

        }
        catch (Exception ex)
        {
            return null;
        }
    }

    /// <summary>
    /// Gets the job layout data.
    /// </summary>
    /// <param name="groupIds">The group ids.</param>
    /// <returns>returns job layout data</returns>
    [WebMethod]
    public static List<Job> GetJobLayoutData(List<string> jobIds, string startDate, string endDate, int viewType)
    {
        try
        {
            string command = GetJobQuery(jobIds, startDate, endDate, viewType);
            DataSet dsData = GetDataSetByCommand(command);
            return GetCustomerJobList(dsData);
        }
        catch (Exception ex)
        {
            return null;
        }
    }

    /// <summary>
    /// Gets the group layout data.
    /// </summary>
    /// <param name="groupIds">The group ids.</param>
    /// <param name="startDate">The start date.</param>
    /// <param name="endDate">The end date.</param>
    /// <returns>
    /// returns group layout data
    /// </returns>
    [WebMethod]
    public static List<Employee> GetGroupLayoutData(List<string> groupIds, string startDate, string endDate, int viewType)
    {
        List<Employee> employeeList = new List<Employee>();
        Employee employee;

        try
        {
            string command = GetGroupQuery(groupIds, startDate, endDate, viewType);
            DataSet dsData = GetDataSetByCommand(command);

            if (dsData != null && dsData.Tables.Count > 0 && dsData.Tables[0].Rows.Count > 0)
            {
                DataRow drUsers;
                bool isEmployeeExists = false;
                ActivityModel activity;

                for (int i = 0; i < dsData.Tables[0].Rows.Count; i++)
                {
                    drUsers = dsData.Tables[0].Rows[i];

                    Employee existingEmployee = employeeList.Find(delegate (Employee e)
                    {
                        return e.Id == Convert.ToInt32(drUsers["EmployeeId"]) && e.ProjectId == Convert.ToInt32(drUsers["ProjectGroupId"]);
                    });

                    if (existingEmployee == null)
                    {
                        employee = new Employee();
                        employee.Id = Convert.ToInt32(drUsers["EmployeeId"]);
                        employee.Name = Convert.ToString(drUsers["EmployeeName"]);
                        employee.ProjectId = Convert.ToInt32(drUsers["ProjectGroupId"]);
                        employee.ProjectName = Convert.ToString(drUsers["ProjectName"]);
                        employee.Activities = new List<ActivityModel>();

                        existingEmployee = employee;
                        isEmployeeExists = false;
                    }
                    else
                    {
                        isEmployeeExists = true;
                    }

                    activity = new ActivityModel();
                    activity = GetActivityModel(drUsers);
                    activity.Heading = Convert.ToInt32(drUsers["ActivityHeader"]);
                    activity.JobName = Convert.ToString(drUsers["JobName"]);
                    activity.CustomerName = Convert.ToString(drUsers["CustomerName"]);

                    existingEmployee.Activities.Add(activity);

                    if (!isEmployeeExists)
                    {
                        employeeList.Add(existingEmployee);
                    }
                }
            }

            return employeeList;
        }
        catch (Exception ex)
        {
            return null;
        }
    }

    /// <summary>
    /// Gets the employee layout data.
    /// </summary>
    /// <param name="employeeIds">The employee ids.</param>
    /// <param name="startDate">The start date.</param>
    /// <param name="endDate">The end date.</param>
    /// <returns>returns list of employee</returns>
    [WebMethod]
    public static List<Employee> GetEmployeeLayoutData(List<string> employeeIds, string startDate, string endDate, int viewType)
    {
        List<Employee> employeeList = new List<Employee>();
        Employee employee;

        try
        {
            string command = GetEmployeeQuery(employeeIds, startDate, endDate, viewType);
            DataSet dsData = GetDataSetByCommand(command);

            if (dsData != null && dsData.Tables.Count > 0 && dsData.Tables[0].Rows.Count > 0)
            {
                DataRow drUsers;
                bool isEmployeeExists = false;
                ActivityModel activity;

                for (int i = 0; i < dsData.Tables[0].Rows.Count; i++)
                {
                    drUsers = dsData.Tables[0].Rows[i];

                    Employee existingEmployee = employeeList.Find(delegate (Employee e)
                    {
                        return e.Id == Convert.ToInt32(drUsers["EmployeeId"]);
                    });

                    if (existingEmployee == null)
                    {
                        employee = new Employee();
                        employee.Id = Convert.ToInt32(drUsers["EmployeeId"]);
                        employee.Name = Convert.ToString(drUsers["EmployeeName"]);
                        employee.Activities = new List<ActivityModel>();

                        existingEmployee = employee;
                        isEmployeeExists = false;
                    }
                    else
                    {
                        isEmployeeExists = true;
                    }

                    activity = new ActivityModel();
                    activity = GetActivityModel(drUsers);
                    activity.Heading = Convert.ToInt32(drUsers["ActivityHeader"]);
                    activity.JobName = Convert.ToString(drUsers["JobName"]);
                    activity.CustomerName = Convert.ToString(drUsers["CustomerName"]);

                    existingEmployee.Activities.Add(activity);

                    if (!isEmployeeExists)
                    {
                        employeeList.Add(existingEmployee);
                    }
                }
            }

            return employeeList;
        }
        catch (Exception ex)
        {
            return null;
        }
    }

    #endregion

    #region Activity get/save methods

    /// <summary>
    /// Gets the customer layout data.
    /// </summary>
    /// <param name="groupIds">The group ids.</param>
    /// <returns>returns customer layout data</returns>
    [WebMethod]
    public static ActivityModel EditActivity(int bookingId)
    {
        ActivityModel model = new ActivityModel();
        try
        {
            string command = GetEditActivityQuery(bookingId);
            DataSet dsData = GetDataSetByCommand(command);

            if (dsData != null && dsData.Tables.Count > 0 && dsData.Tables[0].Rows.Count > 0)
            {
                DataRow drUsers;

                for (int i = 0; i < dsData.Tables[0].Rows.Count; i++)
                {
                    drUsers = dsData.Tables[0].Rows[i];

                    if (model.BookingId == 0)
                    {
                        model = GetActivityModel(drUsers);

                        if (model.StartDateTime != null)
                        {
                            model.StartDate = model.StartDateTime.ToString("yyyy-MM-dd");
                            model.StartTime = model.StartDateTime.ToString("HH:mm:ss");
                        }

                        if (model.EndDateTime != null)
                        {
                            model.EndDate = model.EndDateTime.ToString("yyyy-MM-dd");
                            model.EndTime = model.EndDateTime.ToString("HH:mm:ss");
                        }

                        model.Customer = Convert.ToString(drUsers["CustomerName"]);
                        model.Heading = Convert.ToInt32(drUsers["ActivityHeader"]);
                        model.Project = Convert.ToString(drUsers["JobName"]);
                        model.EmployeeList = new List<Employee>();
                        model.EmployeeList.Add(GetEmployeeModel(drUsers));
                    }
                    else
                    {
                        model.EmployeeList.Add(GetEmployeeModel(drUsers));
                    }
                }
            }

            return model;
        }
        catch (Exception ex)
        {
            return null;
        }
    }

    /// <summary>
    /// Saves the task.
    /// </summary>
    /// <param name="taskData">The task data.</param>
    /// <returns>returns value after save task</returns>
    [WebMethod]
    public static int SaveActivity(ActivityModel model)
    {
        

       try
        {
             string command = SaveActivityQuery(model);
             return GetExecuteNonQueryByCommand(command);
         }
         catch (Exception ex)
         {
            return 0;
         }
    }

    #endregion

    #endregion
}

/// <summary>
/// Employee class
/// </summary>
public class Employee
{
    /// <summary>
    /// The identifier
    /// </summary>
    public int Id = 0;

    /// <summary>
    /// The name
    /// </summary>
    public string Name = string.Empty;

    /// <summary>
    /// The project identifier
    /// </summary>
    public int ProjectId = 0;

    /// <summary>
    /// The project name
    /// </summary>
    public string ProjectName = string.Empty;

    /// <summary>
    /// The activities
    /// </summary>
    public List<ActivityModel> Activities;
}

/// <summary>
/// Job class
/// </summary>
public class Job
{
    /// <summary>
    /// The job identifier
    /// </summary>
    public int JobId = 0;

    /// <summary>
    /// The customer identifier
    /// </summary>
    public int CustomerId = 0;

    /// <summary>
    /// The job name
    /// </summary>
    public string JobName = string.Empty;

    /// <summary>
    /// The customer name
    /// </summary>
    public string CustomerName = string.Empty;

    /// <summary>
    /// The activities
    /// </summary>
    public List<ActivityModel> Activities;

}

/// <summary>
/// Activity Model
/// </summary>
public class ActivityModel
{
    public int BookingId = 0;
    public int ActivityId = 0;
    public string StartDate = string.Empty;
    public string EndDate = string.Empty;
    public string StartTime = string.Empty;
    public string EndTime = string.Empty;
    public DateTime StartDateTime;
    public DateTime EndDateTime;
    public string From = string.Empty;
    public string To = string.Empty;
    public int Recurrence = 0;
    public int Important = 0;
    public string Employee = string.Empty;
    public string Customer = string.Empty;
    public int Heading = 0;
    public string Project = string.Empty;
    public int ActivityType = 0;
    public int NoOfRecurrence = 0;
    public string ActivityName = string.Empty;
    public string[] Employees;
    public string JobName = string.Empty;
    public string CustomerName = string.Empty;

    public List<Employee> EmployeeList { get; set; }
}

/// <summary>
/// Holiday model
/// </summary>
public class Holiday
{
    public int HolidayId { get; set; }
    public DateTime HolidayDate { get; set; }
}