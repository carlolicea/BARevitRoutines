using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

using Autodesk.Revit.ApplicationServices;
using RVTApplication = Autodesk.Revit.ApplicationServices.Application;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitRoutinesAddin
{
    public partial class RequestHandler: IExternalEventHandler
    {
        UIControlledApplication uiApp = null;
        public RequestHandler(UIControlledApplication newUIApp)
        {
            uiApp = newUIApp;
        }
        public Request m_request = new Request();
        public Request Request
        {
            get { return m_request; }
        }
        public String GetName()
        {
            return "BA Revit Routines External Event";
        }
        public void Execute(UIApplication uiApp)
        {
            try
            {
                switch (Request.Take())
                {
                    case RequestId.None:
                        return;
                    case RequestId.CollectProjectLinksData:
                        CollectProjectLinksData(uiApp, "Get Data");
                        break;
                }
            }
            finally
            {
                ;
            }
            return;
        }

        public void CollectProjectLinksData(UIApplication uiApp, String text)
        {
            RVTApplication app = uiApp.Application;
            //Document doc = app.NewProjectDocument(@"C:\ProgramData\Autodesk\RVT 2018\Templates\Generic\Default_I_ENU.rte");
            int daysBack = -60;
            //DataTable dallasCollection = GeneralOperations.PerformOperations(daysBack, "Dallas", Properties.Settings.Default.DalDrivePath, "bimProjectData_DAL", Properties.Settings.Default.CsvOutputPath);
            DataTable sfCollection = GeneralOperations.PerformOperations(daysBack, "San Francisco", Properties.Settings.Default.SfDrivePath, "bimProjectData_SF", Properties.Settings.Default.CsvOutputPath);
            DataTable ocCollection = GeneralOperations.PerformOperations(daysBack, "Irvine", Properties.Settings.Default.OcDrivePath, "bimProjectData_OC", Properties.Settings.Default.CsvOutputPath);
            DataTable sacCollection = GeneralOperations.PerformOperations(daysBack, "Sacramento", Properties.Settings.Default.SacDrivePath, "bimProjectData_SAC", Properties.Settings.Default.CsvOutputPath);
            DataTable bldCollection = GeneralOperations.PerformOperations(daysBack, "Boulder", Properties.Settings.Default.BldDrivePath, "bimProjectData_BLD", Properties.Settings.Default.CsvOutputPath);
        }
    }

    public class GeneralOperations
    {
        public static DataTable PerformOperations(int daysBack, string officeLocation, string officeDrivePath, string officeDbTableName, string csvSaveDirectory)
        {
            DateTime date = DateTime.Now;
            DateTime startDate = DateTime.Now.AddDays(daysBack);
            string year = DateTime.Today.Year.ToString();
            string month = DateTime.Today.Month.ToString();
            string day = DateTime.Today.Day.ToString();

            CreateOutputLog LogFile = new CreateOutputLog(officeLocation, startDate, date, officeDrivePath);
            System.Diagnostics.Debug.WriteLine("Collecting " + officeLocation + " Project Files");
            List<string> filesToCheck = GetAllRvtProjectFiles(officeDrivePath, startDate, LogFile);
            System.Diagnostics.Debug.WriteLine("Collecting " + officeLocation + " Project Data");
            DataTable dataTable = FillDataTable(filesToCheck, LogFile.m_fileReadErrors);
            SqlConnection sqlConnection = DatabaseOperations.SqlOpenConnection(DatabaseOperations.adminDataSqlConnectionString);
            System.Diagnostics.Debug.WriteLine("Writing " + officeLocation + " Project Data to SQL Database");
            DatabaseOperations.SqlWriteDataTable(officeDbTableName, sqlConnection, dataTable, LogFile);
            DatabaseOperations.SqlCloseConnection(sqlConnection);
            CreateCSVFromDataTable(dataTable, officeLocation + " PROJECT FILES " + year + month + day, csvSaveDirectory);
            LogFile.SetOutputLogData(officeLocation, startDate, date, officeDrivePath, LogFile.m_fileReadErrors, LogFile.m_newDbEntries, LogFile.m_existingDbEntries, LogFile.m_dbTableName, year, month, day);
            return dataTable;
        }
        public static List<string> GetAllRvtProjectFiles(string directoryPath, DateTime date, CreateOutputLog log)
        {
            List<string> files = new List<string>();
            string[] directories = Directory.GetDirectories(directoryPath);
            foreach (string directory in directories)
            {
                try
                {
                    List<string> filePaths = Directory.EnumerateFiles(directory, "*.rvt", SearchOption.AllDirectories).ToList();
                    foreach (string file in filePaths)
                    {
                        try
                        {
                            if (file.Contains("E1 Revit"))
                            {
                                File.SetAttributes(file, FileAttributes.Normal);
                                File.SetAttributes(file, FileAttributes.Archive);
                                FileInfo fileInfo = new FileInfo(file);
                                if (fileInfo.LastWriteTime >= date)
                                {
                                    files.Add(file);
                                }
                            }
                        }
                        catch (Exception e)
                        {
                            System.Diagnostics.Debug.WriteLine(String.Format("{0} : Exception: {1}", file, e.Message));
                            log.m_fileReadErrors.Add(file);
                            log.m_fileReadErrors.Add("    Exception: " + e.Message);
                            continue;
                        }
                    }
                }
                catch { continue; }
            }
            return files;
        }
        public static DataTable FillDataTable(List<string> files, List<string> fileReadErrors)
        {
            DataTable dt = new DataTable();
            DataColumn projectNumberColumn = dt.Columns.Add("ProjectNumber", typeof(String));
            DataColumn hostModelNameColumn = dt.Columns.Add("HostModelName", typeof(String));
            DataColumn hostRevitVersionColumn = dt.Columns.Add("HostRevitVersion", typeof(Int32));
            DataColumn hostFilePathColumn = dt.Columns.Add("HostFilePath", typeof(String));
            DataColumn hostFileSizeColumn = dt.Columns.Add("HostSizeMB", typeof(Decimal));
            DataColumn linkModelNameColumn = dt.Columns.Add("LinkModelName", typeof(String));
            DataColumn linkRevitVersionColumn = dt.Columns.Add("LinkRevitVersion", typeof(Int32));
            DataColumn linkFilePathColumn = dt.Columns.Add("LinkFilePath", typeof(String));
            DataColumn linkFileSizeColumn = dt.Columns.Add("LinkSizeMB", typeof(Decimal));
            DataColumn dateColumn = dt.Columns.Add("DateModified", typeof(DateTime));


            foreach (string file in files)
            {
                bool skip = false;

                string projectNumber = "";
                Match matchProjectNumber1 = Regex.Match(file, @"[0-9][0-9][0-9][0-9][0-9][0-9].[0-9][0-9]", RegexOptions.IgnoreCase);
                Match matchProjectNumber2 = Regex.Match(file, @"[0-9][0-9][0-9][0-9][0-9][0-9].\w\w", RegexOptions.IgnoreCase);
                Match matchProjectNumber3 = Regex.Match(file, @"[0-9][0-9][0-9][0-9][0-9][0-9]", RegexOptions.IgnoreCase);

                if (matchProjectNumber1.Success)
                {
                    GroupCollection groups = Regex.Match(file, @"[0-9][0-9][0-9][0-9][0-9][0-9].[0-9][0-9]", RegexOptions.IgnoreCase).Groups;
                    projectNumber = matchProjectNumber1.Value;
                }
                else if (matchProjectNumber2.Success)
                {
                    GroupCollection groups = Regex.Match(file, @"[0-9][0-9][0-9][0-9][0-9][0-9].\w\w", RegexOptions.IgnoreCase).Groups;
                    projectNumber = matchProjectNumber2.Value;
                }
                else if (matchProjectNumber3.Success)
                {
                    GroupCollection groups = Regex.Match(file, @"[0-9][0-9][0-9][0-9][0-9][0-9]", RegexOptions.IgnoreCase).Groups;
                    projectNumber = matchProjectNumber3.Value;
                }
                else
                {
                    skip = true;
                }

                if (skip == false)
                {
                    try
                    {
                        FileInfo fileInfo = new FileInfo(file);
                        string hostModelName = Path.GetFileNameWithoutExtension(file);
                        int hostRevitVersion = GeneralOperations.GetRevitNumber(file);
                        string hostFilePath = file;
                        decimal hostFileSize = fileInfo.Length / 1000000m;
                        string linkModelName = "";
                        int linkRevitVersion = 0;
                        string linkFilePath = "";
                        decimal linkFileSize = 0m;
                        DateTime date = fileInfo.LastWriteTime;

                        ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(file);
                        TransmissionData transmissionData = TransmissionData.ReadTransmissionData(modelPath);
                        if (transmissionData != null)
                        {
                            ICollection<ElementId> elementIds = transmissionData.GetAllExternalFileReferenceIds();
                            foreach (ElementId elementId in elementIds)
                            {
                                ExternalFileReference externalFileReference = transmissionData.GetLastSavedReferenceData(elementId);
                                if (externalFileReference.ExternalFileReferenceType == ExternalFileReferenceType.RevitLink && externalFileReference.GetLinkedFileStatus() != LinkedFileStatus.NotFound)
                                {
                                    try
                                    {
                                        linkFilePath = ModelPathUtils.ConvertModelPathToUserVisiblePath(externalFileReference.GetAbsolutePath());
                                        linkModelName = Path.GetFileNameWithoutExtension(linkFilePath);
                                        linkRevitVersion = GeneralOperations.GetRevitNumber(linkFilePath);
                                        FileInfo linkFileInfo = new FileInfo(linkFilePath);
                                        linkFileSize = linkFileInfo.Length / 1000000m;

                                        DataRow row = dt.NewRow();
                                        row["ProjectNumber"] = projectNumber;
                                        row["HostModelName"] = hostModelName;
                                        row["HostRevitVersion"] = hostRevitVersion;
                                        row["HostFilePath"] = hostFilePath;
                                        row["HostSizeMB"] = hostFileSize;
                                        row["LinkModelName"] = linkModelName;
                                        row["LinkRevitVersion"] = linkRevitVersion;
                                        row["LinkFilePath"] = linkFilePath;
                                        row["LinkSizeMB"] = linkFileSize;
                                        row["DateModified"] = date;
                                        dt.Rows.Add(row);
                                    }
                                    catch (Exception e)
                                    {
                                        System.Diagnostics.Debug.WriteLine(String.Format("{0} : Exception: {1}", file, e.Message));
                                        fileReadErrors.Add(file);
                                        fileReadErrors.Add("    Exception: " + e.Message);
                                        continue;
                                    }
                                }
                            }
                        }
                        else
                        {
                            try
                            {
                                DataRow row = dt.NewRow();
                                row["ProjectNumber"] = projectNumber;
                                row["HostModelName"] = hostModelName;
                                row["HostRevitVersion"] = hostRevitVersion;
                                row["HostFilePath"] = hostFilePath;
                                row["HostSizeMB"] = hostFileSize;
                                row["LinkModelName"] = linkModelName;
                                row["LinkRevitVersion"] = linkRevitVersion;
                                row["LinkFilePath"] = linkFilePath;
                                row["LinkSizeMB"] = linkFileSize;
                                row["DateModified"] = date;
                                dt.Rows.Add(row);
                            }
                            catch (Exception e)
                            {
                                System.Diagnostics.Debug.WriteLine(String.Format("{0} : Exception: {1}", file, e.Message));
                                fileReadErrors.Add(file);
                                fileReadErrors.Add("    Exception: " + e.Message);
                                continue;
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        System.Diagnostics.Debug.WriteLine(String.Format("{0} : Exception: {1}", file, e.Message));
                        fileReadErrors.Add(file);
                        fileReadErrors.Add("    Exception: " + e.Message);
                        continue;
                    }

                }
            }
            return dt;
        }
        public static string GetRevitVersion(string filePath)
        {
            if (filePath != null && filePath != "")
            {
                try
                {
                    BasicFileInfo rvtInfo = BasicFileInfo.Extract(filePath);
                    string rvtVersion = rvtInfo.SavedInVersion.ToString();
                    return rvtVersion;
                }
                catch
                {
                    return string.Empty;
                }
            }
            else { return string.Empty; }
        }
        public static int GetRevitNumber(string filePath)
        {
            int rvtNumber = 0;
            try
            {
                string rvtVersion = GeneralOperations.GetRevitVersion(filePath);
                rvtNumber = Convert.ToInt32(rvtVersion.Substring(rvtVersion.Length - 4));
                return rvtNumber;

            }
            catch { return rvtNumber; }
        }
        public static StringBuilder BuildCSVStringFromDataTable(DataTable dt)
        {
            StringBuilder output = new StringBuilder();
            foreach (DataColumn column in dt.Columns)
            {
                var item = column.ColumnName;
                output.AppendFormat(string.Concat("\"", item.ToString(), "\"", ","));
            }
            output.AppendLine();

            foreach (DataRow rows in dt.Rows)
            {
                foreach (DataColumn col in dt.Columns)
                {
                    var item = rows[col];
                    output.AppendFormat(string.Concat("\"", item.ToString(), "\"", ","));
                }
                output.AppendLine();
            }
            return output;
        }
        public static StringBuilder BuildCSVStringFromDataTableRow(DataTable dt, DataRow row)
        {
            StringBuilder output = new StringBuilder();
            foreach (DataColumn column in dt.Columns)
            {
                var item = row[column].ToString();
                output.AppendFormat(string.Concat("\"", item.ToString(), "\"", ","));

            }
            output.AppendLine();
            return output;
        }
        public static StringBuilder BuildStringFromDataTable(DataTable dt)
        {
            StringBuilder output = new StringBuilder();
            foreach (DataRow rows in dt.Rows)
            {
                foreach (DataColumn col in dt.Columns)
                {
                    var item = rows[col];
                    output.AppendFormat(string.Concat(item.ToString(), " "));
                }
                output.AppendLine();
            }
            return output;
        }


        public static void CreateCSVFromDataTable(DataTable dt, string exportName, string exportDirectory)
        {
            string exportPath = exportDirectory + @"\" + exportName + ".csv";
            if (File.Exists(exportPath))
            {
                File.Delete(exportPath);
            }
            StringBuilder sb = GeneralOperations.BuildCSVStringFromDataTable(dt);
            File.WriteAllText(exportPath, sb.ToString());
        }
    }

    public static class DatabaseOperations
    {
        private static readonly string integratedSecurity = "False";
        private static readonly string userId = RevitRoutinesAddin.Properties.Settings.Default.SQLServerUser;
        private static readonly string password = RevitRoutinesAddin.Properties.Settings.Default.SQLServerPwd;
        private static readonly string connectTimeout = "3";
        private static readonly string encrypt = "False";
        private static readonly string trustServerCertificate = "True";
        private static readonly string applicationIntent = "ReadWrite";
        private static readonly string multiSubnetFailover = "False";
        private static readonly string dbServer = RevitRoutinesAddin.Properties.Settings.Default.SQLServerName;
        private static readonly string database = RevitRoutinesAddin.Properties.Settings.Default.BABimDbName;
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
                                    }catch (Exception e) { MessageBox.Show(e.ToString()); }
                                }                                
                            }catch (Exception e) { MessageBox.Show(e.ToString()); }

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
                        catch(Exception sqlException)
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
            finally{SqlCloseConnection(sqlConnection); }
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

            System.IO.StreamWriter file = new StreamWriter(RevitRoutinesAddin.Properties.Settings.Default.CsvOutputPath + "\\" + officeLocation + " Parser Log " + year + month + day + ".txt");
            file.WriteLine(sb.ToString());
        }
    }
}
