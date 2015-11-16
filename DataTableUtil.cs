using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.Sql;

namespace com.github.yedijas.util
{
    /// <summary>
    /// This class is used to ease the common process that use DataTable.
    /// </summary>
    public class DataTableUtil
    {
        #region static methods
        /// <summary>
        /// Export data in DataTable to Excel file.
        /// </summary>
        /// <param name="DataToExport">DataTable contains data to export.</param>
        /// <param name="Path">Target path of export without the file name.</param>
        /// <param name="FileName">Name of the target excel file.</param>
        public static void DataTableToExcel(DataTable DataToExport, string Path, string FileName)
        {
            ExcelUtil.DataTableToExcel(DataToExport, Path, FileName);
        }

        /// <summary>
        /// Export data in DataTable to Excel file.
        /// </summary>
        /// <param name="DataToExport">DataTable contains data to export.</param>
        /// <param name="CompleteFilePath">Target path of export complete with
        /// the file name and extension.</param>
        public static void DataTableToExcel(DataTable DataToExport, string CompleteFilePath)
        {
            ExcelUtil.DataTableToExcel(DataToExport, CompleteFilePath);
        }

        /// <summary>
        /// Export data in DataTable to CSV file.
        /// </summary>
        /// <param name="DataToExport">DataTable contains data to export.</param>
        /// <param name="Path">Target path of export without the file name.</param>
        /// <param name="FileName">Name of the target excel file.</param>
        public static void DataTableToCSV(DataTable DataToExport, string Path, string FileName)
        {
            CSVutil.DataTableToCSV(DataToExport, Path, FileName);
        }

        /// <summary>
        /// Export data in DataTable to CSV file.
        /// </summary>
        /// <param name="DataToExport">DataTable contains data to export.</param>
        /// <param name="CompleteFilePath">Target path of export complete with
        /// the file name and extension.</param>
        public static void DataTableToCSV(DataTable DataToExport, string CompleteFilePath)
        {
            CSVutil.DataTableToCSV(DataToExport, CompleteFilePath);
        }

        /// <summary>
        /// Get column names from a data table.
        /// </summary>
        /// <param name="DataToGet">DataTable containing data which column names will be taken.</param>
        /// <returns>A list of string containing column name.</returns>
        public static List<string> GetColumnNames(DataTable DataToGet)
        {
            List<string> columnNames = new List<string>();
            foreach (DataColumn column in DataToGet.Columns)
            {
                columnNames.Add(column.ColumnName);
            }
            return columnNames;
        }

        /// <summary>
        /// Method to help to realease excel object cleanly.
        /// </summary>
        /// <param name="obj">object to release.</param>
        public static void ReleaseObject(object obj)
        {
            if (obj != null &&
                System.Runtime.InteropServices.Marshal.IsComObject(obj))
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            }
            obj = null;
            GC.Collect();
        }
        #endregion
    }
}
