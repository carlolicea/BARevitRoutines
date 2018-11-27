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


namespace BARevitRoutines.RoutineRequests
{
    class CollectProjectLinksData
    {
        public CollectProjectLinksData(UIApplication uiApp, String text)
        {
            RVTApplication app = uiApp.Application;
            //Document doc = app.NewProjectDocument(@"C:\ProgramData\Autodesk\RVT 2018\Templates\Generic\Default_I_ENU.rte");
            int daysBack = -60;
            DataTable dallasCollection = PerformOperations(daysBack, "Dallas", Properties.Settings.Default.DalDrivePath, "bimProjectData_DAL", Properties.Settings.Default.CsvOutputPath);
            DataTable sfCollection = PerformOperations(daysBack, "San Francisco", Properties.Settings.Default.SfDrivePath, "bimProjectData_SF", Properties.Settings.Default.CsvOutputPath);
            DataTable ocCollection = PerformOperations(daysBack, "Irvine", Properties.Settings.Default.OcDrivePath, "bimProjectData_OC", Properties.Settings.Default.CsvOutputPath);
            DataTable sacCollection = PerformOperations(daysBack, "Sacramento", Properties.Settings.Default.SacDrivePath, "bimProjectData_SAC", Properties.Settings.Default.CsvOutputPath);
            DataTable bldCollection = PerformOperations(daysBack, "Boulder", Properties.Settings.Default.BldDrivePath, "bimProjectData_BLD", Properties.Settings.Default.CsvOutputPath);
        }

        public static DataTable PerformOperations(int daysBack, string officeLocation, string officeDrivePath, string officeDbTableName, string csvSaveDirectory)
        {
            DateTime date = DateTime.Now;
            DateTime startDate = DateTime.Now.AddDays(daysBack);
            string year = DateTime.Today.Year.ToString();
            string month = DateTime.Today.Month.ToString();
            string day = DateTime.Today.Day.ToString();

            CreateOutputLog LogFile = new CreateOutputLog(officeLocation, startDate, date, officeDrivePath);
            System.Diagnostics.Debug.WriteLine("Collecting " + officeLocation + " Project Files");
            List<string> filesToCheck = GeneralOperations.GetAllRvtProjectFiles(officeDrivePath, startDate, LogFile);
            System.Diagnostics.Debug.WriteLine("Collecting " + officeLocation + " Project Data");
            DataTable dataTable = FillDataTable(filesToCheck, LogFile.m_fileReadErrors);
            SqlConnection sqlConnection = DatabaseOperations.SqlOpenConnection(DatabaseOperations.adminDataSqlConnectionString);
            System.Diagnostics.Debug.WriteLine("Writing " + officeLocation + " Project Data to SQL Database");
            DatabaseOperations.SqlWriteDataTable(officeDbTableName, sqlConnection, dataTable, LogFile);
            DatabaseOperations.SqlCloseConnection(sqlConnection);
            GeneralOperations.CreateCSVFromDataTable(dataTable, officeLocation + " PROJECT FILES " + year + month + day, csvSaveDirectory);
            LogFile.SetOutputLogData(officeLocation, startDate, date, officeDrivePath, LogFile.m_fileReadErrors, LogFile.m_newDbEntries, LogFile.m_existingDbEntries, LogFile.m_dbTableName, year, month, day);
            return dataTable;
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

    }
}
