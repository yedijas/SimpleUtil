using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace com.github.yedijas.util
{
    /// <summary>
    /// This class is used to ease the common process that are related with CSV file.
    /// </summary>
    class CSVUtil
    {
        #region static methods
        /// <summary>
        /// Export list to CSV file.
        /// List is shaped like a database, the rows are in Dictionary with string as key and value.
        /// </summary>
        /// <param name="ListToExport">List containing data to be exported to CSV file.</param>
        /// <param name="CompleteFilePath">Target path of export complete with
        /// the filename and extension.</param>
        public static void ExportListToCSV(List<Dictionary<string, string>> ListToExport,
            string CompleteFilePath)
        {
            List<string> columnNames = GetColumnNames(ListToExport[0]);
            StringBuilder csvContent = new StringBuilder();

            for (int i = 0; i < columnNames.Count; i++)
            {
                csvContent.Append(CheckCSVSafe(columnNames[i]));
                if (i != (columnNames.Count - 1))
                {
                    csvContent.Append(",");
                }
                else
                {
                    csvContent.AppendLine();
                }
            }
            for (int i = 0; i < ListToExport.Count; i++)
            {
                for (int j = 0; j < ListToExport[i].Count; j++)
                {
                    csvContent.Append(ListToExport[i][columnNames[j]].ToString());
                    if (j != (ListToExport[i].Count - 1))
                    {
                        csvContent.Append(",");
                    }
                    else
                    {
                        csvContent.AppendLine();
                    }
                }
            }
            System.IO.File.Create(CompleteFilePath);
            try
            {
                System.IO.File.WriteAllText(CompleteFilePath, csvContent.ToString(), Encoding.UTF8);
            }
            catch (Exception allEx)
            {
                throw allEx;
            }
        }

        /// <summary>
        /// Export list to CSV file.
        /// List is shaped like a database, the rows are in Dictionary with string as key and value.
        /// </summary>
        /// <param name="ListToExport">List containing data to be exported to CSV file.</param>
        /// <param name="Path">Target path of export.</param>
        /// <param name="FileName">Name of  the target CSV file.</param>
        public static void ExportListToCSV(List<Dictionary<string, string>> ListToExport,
            string Path, string FileName)
        {
            if (!String.IsNullOrEmpty(
                System.IO.Path.GetExtension(FileName)))
            {
                ExportListToCSV(ListToExport, (Path + @"\" + FileName));
            }
            else
            {
                ExportListToCSV(ListToExport, (Path + @"\" + FileName + ".csv"));
            }
        }

        /// <summary>
        /// Get the column name of the database shaped list.
        /// </summary>
        /// <param name="SingleRow">Single row of Dictionary which keys will be taken
        /// as the column name.</param>
        /// <returns>Column names shaped in list of string.</returns>
        public static List<string> GetColumnNames(Dictionary<string, string> SingleRow)
        {
            List<string> columnNames = new List<string>();
            foreach (KeyValuePair<string, string> column in SingleRow)
            {
                columnNames.Add(column.Key);
            }
            return columnNames;
        }

        /// <summary>
        /// Make the data passed in parameter as correct value to be inserted as CSV.
        /// </summary>
        /// <param name="StringToCheck">Data to check</param>
        /// <returns>Data safe to be put to CSV.</returns>
        public static string CheckCSVSafe(string StringToCheck)
        {
            bool mustQuote = (StringToCheck.Contains(",") || StringToCheck.Contains("\"") || StringToCheck.Contains("\r") || StringToCheck.Contains("\n"));
            if (mustQuote)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append("\"");
                foreach (char nextChar in StringToCheck)
                {
                    sb.Append(nextChar);
                    if (nextChar == '"')
                        sb.Append("\"");
                }
                sb.Append("\"");
                return sb.ToString();
            }

            return StringToCheck;
        }

        /// <summary>
        /// Export data in DataTable to CSV file.
        /// </summary>
        /// <param name="DataToExport">DataTable containing data to export.</param>
        /// <param name="Path">Path to target file.</param>
        /// <param name="FileName">Target file name and extension.</param>
        public static void DataTableToCSV(System.Data.DataTable DataToExport, string Path, string FileName)
        {
            if (!string.IsNullOrEmpty(
                System.IO.Path.GetExtension(FileName)))
            {
                DataTableToCSV(DataToExport, (Path + @"\" + FileName));
            }
            else
            {
                DataTableToCSV(DataToExport, (Path + @"\" + FileName + ".csv"));
            }
        }

        /// <summary>
        /// Export data in DataTable to CSV file.
        /// </summary>
        /// <param name="DataToExport">DataTable containing data to export.</param>
        /// <param name="CompletePath">Path and filename to target csv file.</param>
        public static void DataTableToCSV(System.Data.DataTable DataToExport, string CompleteFilePath)
        {
            List<string> columnNames = DataTableUtil.GetColumnNames(DataToExport);
            StringBuilder csvContent = new StringBuilder();

            for (int i = 0; i < columnNames.Count; i++)
            {
                csvContent.Append(CheckCSVSafe(columnNames[i]));
                if (i != (columnNames.Count - 1))
                {
                    csvContent.Append(",");
                }
                else
                {
                    csvContent.AppendLine();
                }
            }
            for (int i = 0; i < DataToExport.Rows.Count; i++)
            {
                for (int j = 0; j < DataToExport.Rows[i].ItemArray.Length; j++)
                {
                    csvContent.Append(DataToExport.Rows[i].ItemArray[j].ToString());
                    if (j != (DataToExport.Rows[i].ItemArray.Length - 1))
                    {
                        csvContent.Append(",");
                    }
                    else
                    {
                        csvContent.AppendLine();
                    }
                }
            }
            System.IO.File.Create(CompleteFilePath);
            try
            {
                System.IO.File.WriteAllText(CompleteFilePath, csvContent.ToString(), Encoding.UTF8);
            }
            catch (Exception allEx)
            {
                throw allEx;
            }
        }

        /// <summary>
        /// Export data from CSV file given the full filename.
        /// </summary>
        /// <param name="FileName">CSV file</param>
        /// <returns>DataTable containing data from CSV file. All columns are in string type.</returns>
        public static System.Data.DataTable CSVToDataTable(string FileName)
        {
            #region logic
            System.Data.DataTable result = new System.Data.DataTable();
            System.IO.StreamReader fileReader = null;
            if (!System.IO.File.Exists(FileName))
            {
                throw new System.IO.IOException("File not found!");
            }
            if (new System.IO.FileInfo(FileName).Length == 0)
            {
                throw new Exception("File is EMPTY!");
            }
            try
            {
                fileReader = new System.IO.StreamReader(FileName);
                List<string> headers = RowToList(fileReader.ReadLine());
                foreach (string header in headers)
                {
                    System.Data.DataColumn tempColumn = new System.Data.DataColumn();
                    tempColumn.ColumnName = header;
                    tempColumn.DataType = Type.GetType("System.String");
                    result.Columns.Add(tempColumn);
                    tempColumn = null;
                }
                string singleRow = "";
                while ((singleRow = fileReader.ReadLine()) != null)
                {
                    System.Data.DataRow tempRow = result.NewRow();
                    List<string> dataInList = RowToList(singleRow);
                    for (int i = 0; i < result.Columns.Count; i++)
                    {
                        tempRow.ItemArray[i] = dataInList[i];
                    }
                }
            }
            #endregion
            #region exception handling
            catch (Exception AllEx)
            {
                throw AllEx;
            }
            finally
            {
                if (fileReader.BaseStream.CanRead)
                {
                    fileReader.Close();
                }
                fileReader = null;
            }
            #endregion
            return result;
        }

        /// <summary>
        /// Convert from comma separated string to list of string.
        /// </summary>
        /// <param name="SingleRow">A comma separated string to process.</param>
        /// <returns>List of string containing data.</returns>
        public static List<string> RowToList(string SingleRow)
        {
            List<string> result = new List<string>();
            foreach (string data in SingleRow.Split(','))
            {
                result.Add(data);
            }
            return result;
        }
        #endregion
    }
}
