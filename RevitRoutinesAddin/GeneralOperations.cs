using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB;

namespace BARevitRoutines
{
    public class GeneralOperations
    {        
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
}
