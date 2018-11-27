using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BARevitRoutines
{
    public static class DatabaseOperations
    {
        private static readonly string integratedSecurity = "False";
        private static readonly string userId = BARevitRoutines.Properties.Settings.Default.SQLServerUser;
        private static readonly string password = BARevitRoutines.Properties.Settings.Default.SQLServerPwd;
        private static readonly string connectTimeout = "3";
        private static readonly string encrypt = "False";
        private static readonly string trustServerCertificate = "True";
        private static readonly string applicationIntent = "ReadWrite";
        private static readonly string multiSubnetFailover = "False";
        private static readonly string dbServer = BARevitRoutines.Properties.Settings.Default.SQLServerName;
        private static readonly string database = BARevitRoutines.Properties.Settings.Default.BABimDbName;
        public static string adminDataSqlConnectionString = "Server=" + dbServer +
                                ";Database=" + database +
                                ";Integrated Security=" + integratedSecurity +
                                ";User Id=" + userId +
                                ";Password=" + password +
                                ";Connect Timeout=" + connectTimeout +
                                ";Encrypt=" + encrypt +
                                ";TrustServerCertificate=" + trustServerCertificate +
                                ";ApplicationIntent=" + applicationIntent +
                                ";MultiSubnetFailover=" + multiSubnetFailover;

        public static SqlConnection SqlOpenConnection(string connectionString)
        {
            List<string> dataTableNames = new List<string>();
            SqlConnection sqlConnection;

            sqlConnection = new SqlConnection(connectionString);
            try
            {
                sqlConnection.Open();
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine(e.Message);
            }
            return sqlConnection;
        }
        public static void SqlCloseConnection(SqlConnection sqlConnection)
        {
            try
            {
                sqlConnection.Close();
            }
            catch
            {
                ;
            }
        }
        public static void SqlWriteDataTable(string tableName, SqlConnection sqlConnection, DataTable dataTable, CreateOutputLog log)
        {
            DataTable dt = sqlConnection.GetSchema("Tables");
            List<string> existingTables = new List<string>();
            foreach (DataRow row in dt.Rows)
            {
                string existingTableName = (string)row[2];
                existingTables.Add(existingTableName);
            }

            if (existingTables.Contains(tableName))
            {
                log.m_dbTableName = tableName;
                using (sqlConnection)
                {
                    //SOMETHING IS GOING ON HERE THAT IS CAUSING AN EXCEPTION WITH THE SQL COMMAND
                    //Possibly rewrite the date modified to pull from the links OR the project.
                    //This broke after the second loop.
                    System.Diagnostics.Debug.WriteLine("Copying Data To SQL Table");
                    foreach (DataRow row in dataTable.Rows)
                    {
                        try
                        {
                            bool skip = false;
                            try
                            {
                                using (SqlCommand command = new SqlCommand("SELECT COUNT (*) FROM " + tableName +
                                    " WHERE ProjectNumber LIKE '" + row["ProjectNumber"] + "'" +
                                    " AND HostModelName LIKE '" + row["HostModelName"] + "'" +
                                    " AND HostRevitVersion = '" + row["HostRevitVersion"] + "'" +
                                    " AND HostFilePath LIKE '" + row["HostFilePath"] + "'" +
                                    " AND HostSizeMB = '" + row["HostSizeMB"] + "'" +
                                    " AND LinkModelName LIKE '" + row["LinkModelName"] + "'" +
                                    " AND LinkRevitVersion = '" + row["LinkRevitVersion"] + "'" +
                                    " AND LinkFilePath LIKE '" + row["LinkFilePath"] + "'" +
                                    " AND LinkSizeMB = '" + row["LinkSizeMB"] + "'"))
                                {
                                    command.Connection = sqlConnection;
                                    try
                                    {
                                        Int32 count = Convert.ToInt32(command.ExecuteScalar());
                                        if (count > 0)
                                        {
                                            skip = true;
                                            StringBuilder sb = GeneralOperations.BuildCSVStringFromDataTableRow(dataTable, row);
                                            log.m_existingDbEntries.Add(sb.ToString());
                                        }
                                    }
                                    catch (Exception e) { MessageBox.Show(e.ToString()); }
                                }
                            }
                            catch (Exception e) { MessageBox.Show(e.ToString()); }

                            if (skip == false)
                            {
                                using (SqlCommand comm = new SqlCommand("INSERT INTO " + tableName + " (ProjectNumber, HostModelName, HostRevitVersion, HostFilePath, HostSizeMB, LinkModelName, LinkRevitVersion, LinkFilePath, LinkSizeMB, DateModified) VALUES (@v1, @v2, @v3, @v4, @v5, @v6, @v7, @v8, @v9, @v10)"))
                                {
                                    comm.Connection = sqlConnection;
                                    comm.Parameters.AddWithValue("@v1", row["ProjectNumber"]);
                                    comm.Parameters.AddWithValue("@v2", row["HostModelName"]);
                                    comm.Parameters.AddWithValue("@v3", row["HostRevitVersion"]);
                                    comm.Parameters.AddWithValue("@v4", row["HostFilePath"]);
                                    comm.Parameters.AddWithValue("@v5", row["HostSizeMB"]);
                                    comm.Parameters.AddWithValue("@v6", row["LinkModelName"]);
                                    comm.Parameters.AddWithValue("@v7", row["LinkRevitVersion"]);
                                    comm.Parameters.AddWithValue("@v8", row["LinkFilePath"]);
                                    comm.Parameters.AddWithValue("@v9", row["LinkSizeMB"]);
                                    comm.Parameters.AddWithValue("@v10", row["DateModified"]);
                                    try
                                    {
                                        comm.ExecuteNonQuery();
                                        StringBuilder sb = GeneralOperations.BuildCSVStringFromDataTableRow(dataTable, row);
                                        log.m_newDbEntries.Add(sb.ToString());
                                    }
                                    catch (Exception e)
                                    {
                                        System.Diagnostics.Debug.WriteLine(e.ToString());
                                    }
                                }
                            }
                        }
                        catch (Exception sqlException)
                        {
                            MessageBox.Show(sqlException.ToString());
                        }
                    }
                }
            }

            else
            {
                log.m_dbTableName = tableName;
                System.Diagnostics.Debug.WriteLine("Creating New SQL Table");
                try
                {
                    SqlCommand sqlCreateTable = new SqlCommand("CREATE TABLE " + tableName + " (ProjectNumber text, HostModelName text, HostRevitVersion int, HostFilePath text, HostSizeMB float, LinkModelName text, LinkRevitVersion int, LinkFilePath text, LinkSizeMB float, DateModified datetime)", sqlConnection);
                    sqlCreateTable.ExecuteNonQuery();
                }
                catch (SqlException f)
                {
                    System.Diagnostics.Debug.WriteLine(f.Message);
                }

                try
                {
                    SqlBulkCopyOptions options = SqlBulkCopyOptions.Default;

                    System.Diagnostics.Debug.WriteLine("Copying Data To SQL Table");
                    using (SqlBulkCopy s = new SqlBulkCopy(sqlConnection, options, null))
                    {
                        s.DestinationTableName = "[" + tableName + "]";
                        foreach (DataColumn appColumn in dataTable.Columns)
                        {
                            s.ColumnMappings.Add(appColumn.ToString(), appColumn.ToString());
                        }
                        s.WriteToServer(dataTable);
                    }
                    StringBuilder sb = GeneralOperations.BuildCSVStringFromDataTable(dt);
                    log.m_newDbEntries.Add(sb.ToString());
                }
                catch (SqlException g)
                {
                    System.Diagnostics.Debug.WriteLine(g.Errors);
                }
                finally
                {
                    DatabaseOperations.SqlCloseConnection(sqlConnection);
                }
            }
            SqlLogWriter(tableName);
        }
        public static void SqlLogWriter(string writtenTableName)
        {
            SqlConnection sqlConnection = DatabaseOperations.SqlOpenConnection(DatabaseOperations.adminDataSqlConnectionString);
            try
            {
                DataTable dt = sqlConnection.GetSchema("Tables");

                List<string> existingTables = new List<string>();
                foreach (DataRow row in dt.Rows)
                {
                    string tableName = (string)row[2];
                    existingTables.Add(tableName);
                }

                if (existingTables.Contains("BARevitTools_SQLWriterLog"))
                {
                    string commandString = "INSERT INTO [BARevitTools_SQLWriterLog] (UserName, TableName, WriteDate) VALUES (@userName, @tableName, @dateTime)";
                    using (SqlCommand sqlInsert = new SqlCommand(commandString, sqlConnection))
                    {
                        sqlInsert.Parameters.AddWithValue("@userName", Environment.UserName);
                        sqlInsert.Parameters.AddWithValue("@tableName", writtenTableName);
                        sqlInsert.Parameters.AddWithValue("@dateTime", DateTime.Now);
                        sqlInsert.ExecuteNonQuery();
                    }
                }
                else
                {
                    SqlCommand sqlCreateTable = new SqlCommand("CREATE TABLE BARevitTools_SQLWriterLog (UserName varchar(255), TableName varchar(255), WriteDate datetime)", sqlConnection);
                    sqlCreateTable.ExecuteNonQuery();
                    string commandString = "INSERT INTO [BARevitTools_SQLWriterLog] (UserName, TableName, WriteDate) VALUES (@userName, @tableName, @dateTime)";
                    using (SqlCommand sqlInsert = new SqlCommand(commandString, sqlConnection))
                    {
                        sqlInsert.Parameters.AddWithValue("@userName", Environment.UserName);
                        sqlInsert.Parameters.AddWithValue("@tableName", writtenTableName);
                        sqlInsert.Parameters.AddWithValue("@dateTime", DateTime.Now);
                        sqlInsert.ExecuteNonQuery();
                    }
                }
                DatabaseOperations.SqlCloseConnection(sqlConnection);
            }
            catch (Exception e) { System.Diagnostics.Debug.WriteLine(e.Message); }
            finally { SqlCloseConnection(sqlConnection); }
        }
    }

    public class CreateOutputLog
    {
        public List<string> m_fileReadErrors = new List<string>();
        public List<string> m_newDbEntries = new List<string>();
        public List<string> m_existingDbEntries = new List<string>();
        public string m_dbTableName;
        public string m_officeLocation;
        public DateTime m_startDateRange;
        public DateTime m_endDateRange;
        public string m_officeDrive;

        public CreateOutputLog(string officeLocation, DateTime startDateRange, DateTime endDateRange, string officeDrive)
        {
            m_officeLocation = officeLocation;
            m_startDateRange = startDateRange;
            m_endDateRange = endDateRange;
            m_officeDrive = officeDrive;
        }

        public void SetOutputLogData(string officeLocation, DateTime startDateRange, DateTime endDateRange, string officeDrive, List<string> slogReadErrors, List<string> newDbEntries, List<string> existingDbEntries, string dbTableName, string year, string month, string day)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(String.Format("Office: {0}", officeLocation));
            sb.Append(Environment.NewLine);
            sb.Append(String.Format("Date Range: {0} - {1}", startDateRange.ToString(), endDateRange.ToString()));
            sb.Append(Environment.NewLine);
            sb.Append(String.Format("Drive Searched: {0}", officeDrive));
            sb.Append(Environment.NewLine);
            sb.Append(Environment.NewLine);
            sb.Append("The Following Errors Occurred:");
            sb.Append(Environment.NewLine);
            foreach (string line in slogReadErrors)
            {
                sb.Append(line);
                sb.Append(Environment.NewLine);
            }
            sb.Append(Environment.NewLine);
            sb.Append(Environment.NewLine);
            sb.Append(String.Format("The Following Entries Were Added To {0}:", dbTableName));
            sb.Append(Environment.NewLine);
            foreach (string line in newDbEntries)
            {
                sb.Append(line);
            }
            sb.Append(Environment.NewLine);
            sb.Append(Environment.NewLine);
            sb.Append(String.Format("The Following Entries Already Existed And Were Skipped In {0}:", dbTableName));
            sb.Append(Environment.NewLine);
            foreach (string line in existingDbEntries)
            {
                sb.Append(line);
            }
            sb.AppendLine();

            System.IO.StreamWriter file = new StreamWriter(BARevitRoutines.Properties.Settings.Default.CsvOutputPath + "\\" + officeLocation + " Parser Log " + year + month + day + ".txt");
            file.WriteLine(sb.ToString());
        }
    }
}
