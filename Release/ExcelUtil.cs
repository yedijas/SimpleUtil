using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace com.github.yedijas.util
{
    class ExcelUtil
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
            if (!string.IsNullOrEmpty(
                System.IO.Path.GetExtension(FileName)))
            {
                DataTableToExcel(DataToExport, (Path + @"\" + FileName));
            }
            else
            {
                DataTableToExcel(DataToExport, (Path + @"\" + FileName + ".xls"));
            }
        }

        /// <summary>
        /// Export data in DataTable to Excel file.
        /// </summary>
        /// <param name="DataToExport">DataTable contains data to export.</param>
        /// <param name="CompleteFilePath">Target path of export complete with
        /// the file name and extension.</param>
        public static void DataTableToExcel(DataTable DataToExport, string CompleteFilePath)
        {
            Excel.Application excelApplication = null;
            Excel.Workbook excelWorkbook = null;
            Excel.Worksheet excelWorkSheet = null;

            #region logic
            if (DataToExport == null || DataToExport.Columns.Count == 0)
                throw new Exception("No data in DataTable");
            try
            {
                excelApplication = new Microsoft.Office.Interop.Excel.Application();
                excelWorkbook = excelApplication.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                excelWorkSheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item(1);

                for (int i = 0; i < DataToExport.Columns.Count; i++)
                {
                    excelWorkSheet.Cells[1, (i + 1)] = DataToExport.Columns[i].ColumnName;
                }
                for (int i = 0; i < DataToExport.Rows.Count; i++)
                {
                    for (int j = 0; j < DataToExport.Columns.Count; j++)
                    {
                        excelWorkSheet.Cells[(i + 2), (j + 1)] = DataToExport.Rows[i][j];
                    }
                }
                excelWorkSheet.UsedRange.NumberFormat = "@";

                try
                {
                    excelWorkSheet.SaveAs(CompleteFilePath,
                                          Excel.XlFileFormat.xlWorkbookNormal,
                                          Type.Missing,
                                          Type.Missing,
                                          Type.Missing,
                                          Type.Missing,
                                          Type.Missing,
                                          Type.Missing,
                                          Type.Missing,
                                          Type.Missing
                                          );

                }
                catch (Exception SaveExcelFileEx)
                {
                    throw SaveExcelFileEx;
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
                if (excelWorkSheet != null &&
                    System.Runtime.InteropServices.Marshal.IsComObject(excelWorkSheet))
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkSheet);
                }
                excelWorkSheet = null;
                excelWorkbook.Close(true, Type.Missing, Type.Missing);
                if (excelWorkbook != null &&
                    System.Runtime.InteropServices.Marshal.IsComObject(excelWorkbook))
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbook);
                }
                excelWorkSheet = null;
                if (excelApplication != null)
                {
                    excelApplication.Quit();
                }
                if (excelApplication != null &&
                    System.Runtime.InteropServices.Marshal.IsComObject(excelApplication))
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApplication);
                }
                excelApplication = null;
                GC.Collect();
            }
            #endregion
        }

        /// <summary>
        /// Export table in Excel to DataTable.
        /// Good for data which started from A1 to end.
        /// Row 1 will be automatically treated as header.
        /// Total row and column will automatically counted.
        /// </summary>
        /// <param name="FileName">Excel file name complete with path and extension.</param>
        /// <returns>DataTable containing data from exce file.</returns>
        public static DataTable ExcelToDataTable(string FileName)
        {
            DataTable result = ExcelToDataTable(FileName, 1, 1, 0, 0);
            return result;
        }

        /// <summary>
        /// Export table in Excel to DataTable.
        /// Good for data which are not started from A1 to end.
        /// Total row and column will be automatically counted.
        /// </summary>
        /// <param name="FileName">Excel file name complete with path and extension.</param>
        /// <param name="HeaderRowIndex">Row number that contains header.</param>
        /// <param name="ColumnStartIndex">Start Index of column.</param>
        /// <returns>DataTable containing data from exce file.</returns>
        public static DataTable ExcelToDataTable(string FileName,
                                                 int ColumnStartIndex,
                                                 int HeaderRowIndex)
        {
            DataTable result = ExcelToDataTable(FileName, HeaderRowIndex, ColumnStartIndex, 0, 0);
            return result;
        }

        /// <summary>
        /// Export table in Excel to DataTable.
        /// Good for data which are not started from A1 and ended to specific row number.
        /// Total column will be automatically counted.
        /// </summary>
        /// <param name="FileName">Excel file name complete with path and extension.</param>
        /// <param name="HeaderRowIndex">Row number that contains header.</param>
        /// <param name="ColumnStartIndex">Start Index of column.</param>
        /// <param name="TotalRow">Total row that contains data in excel table.</param>
        /// <returns>DataTable containing data from exce file.</returns>
        public static DataTable ExcelToDataTable(string FileName,
                                                 int HeaderRowIndex,
                                                 int ColumnStartIndex,
                                                 int TotalRow)
        {
            DataTable result = ExcelToDataTable(FileName, HeaderRowIndex, ColumnStartIndex, TotalRow, 0);
            return result;
        }

        /// <summary>
        /// Export table in Excel to DataTable.
        /// Total column will be automatically counted.
        /// </summary>
        /// <param name="FileName">Excel file name complete with path and extension.</param>
        /// <param name="HeaderRowIndex">Row number that contains header.</param>
        /// <param name="ColumnStartIndex">Start Index of column.</param>
        /// <param name="TotalRow">Total row that contains data in excel table.</param>
        /// <param name="WorkSheetIndex">Index of worksheet containing data to take.</param>
        /// <returns>DataTable containing data from exce file.</returns>
        public static DataTable ExcelToDataTable(string FileName,
                                                 int HeaderRowIndex,
                                                 int ColumnStartIndex,
                                                 int TotalRow,
                                                 int WorkSheetIndex)
        {
            DataTable result = null;
            Excel.Application excelApplication = null;
            Excel.Workbook excelWorkbook = null;
            Excel.Worksheet excelWorksheet = null;
            #region logic
            try
            {
                excelApplication = new Microsoft.Office.Interop.Excel.Application();
                excelWorkbook = excelApplication.Workbooks.Open(FileName,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing);
                if (WorkSheetIndex == 0)
                {// default active
                    excelWorksheet = (Excel.Worksheet)excelWorkbook.ActiveSheet;
                }
                else
                {
                    excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item(WorkSheetIndex);
                }
                if (TotalRow == 0)
                {// default is checked
                    TotalRow = GetTotalRowsFromWorksheet(excelWorksheet);
                }
                int TotalColumn = GetTotalTableColumnFromWorksheet(excelWorksheet,
                    HeaderRowIndex, ColumnStartIndex);
                List<string> headers = GetTableHeader(excelWorksheet, HeaderRowIndex);
                try
                {
                    foreach (string header in headers)
                    {
                        DataColumn tempColumn = new DataColumn();
                        tempColumn.ColumnName = header;
                        tempColumn.DataType = Type.GetType("System.String");
                        result.Columns.Add(tempColumn);
                    }
                    for (int i = HeaderRowIndex + 1; i <= (HeaderRowIndex + TotalRow); i++)
                    {
                        DataRow tempRow = result.NewRow();
                        for (int j = ColumnStartIndex; j < (ColumnStartIndex + TotalColumn); i++)
                        {
                            tempRow.ItemArray[j] = (excelWorksheet.Cells[i, j] as Excel.Range)
                                .Value2.ToString();
                        }
                        result.Rows.Add(tempRow);
                    }
                }
                catch
                {
                    throw;
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
                if (excelWorksheet != null &&
                    System.Runtime.InteropServices.Marshal.IsComObject(excelWorksheet))
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorksheet);
                }
                excelWorksheet = null;
                excelWorkbook.Close(false, Type.Missing, Type.Missing);
                if (excelWorkbook != null &&
                    System.Runtime.InteropServices.Marshal.IsComObject(excelWorkbook))
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbook);
                }
                excelWorkbook = null;
                if (excelApplication != null)
                {
                    excelApplication.Quit();
                }
                if (excelApplication != null &&
                    System.Runtime.InteropServices.Marshal.IsComObject(excelApplication))
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApplication);
                }
                excelApplication = null;
                GC.Collect();
            }
            #endregion
            return result;
        }

        /// <summary>
        /// Get total column filled with data in worksheet.
        /// </summary>
        /// <param name="ExcelWorksheet">Worksheet containing data to count.</param>
        /// <returns>Number of used (contain data) columns.</returns>
        public static int GetTotalColumnFromWorksheet(Excel.Worksheet ExcelWorksheet)
        {
            ExcelWorksheet.Columns.ClearFormats();
            ExcelWorksheet.Rows.ClearFormats();
            return ExcelWorksheet.UsedRange.Columns.Count;
        }

        /// <summary>
        /// Get total rows filled with data in worksheet. 
        /// </summary>
        /// <param name="ExcelWorksheet">Worksheet containing data to count.</param>
        /// <returns>Number of used (contain data) rows.</returns>
        public static int GetTotalRowsFromWorksheet(Excel.Worksheet ExcelWorksheet)
        {
            ExcelWorksheet.Columns.ClearFormats();
            ExcelWorksheet.Rows.ClearFormats();
            return ExcelWorksheet.UsedRange.Columns.Count;
        }

        /// <summary>
        /// Get total column in worksheet given the row with header.
        /// </summary>
        /// <param name="ExcelWorksheet">Worksheet containing data to count.</param>
        /// <param name="HeaderRowIndex">Row index which header is located.</param>
        /// <param name="ColumnStartIndex">Start index of column. Set 1 for column A.</param>
        /// <returns>Number of used (contain data) columns.</returns>
        public static int GetTotalTableColumnFromWorksheet(Excel.Worksheet ExcelWorksheet,
                                                           int HeaderRowIndex,
                                                           int ColumnStartIndex)
        {
            ExcelWorksheet.Columns.ClearFormats();
            ExcelWorksheet.Rows.ClearFormats();
            int index = ColumnStartIndex;
            bool flag = true;
            while (flag)
            {
                if (string.IsNullOrEmpty(
                    (ExcelWorksheet.Cells[HeaderRowIndex, index] as Excel.Range)
                        .Value2.ToString()))
                {
                    flag = false;
                }
                index++;
            }
            return index;
        }

        /// <summary>
        /// Get the table header in Worksheet, given the header index.
        /// </summary>
        /// <param name="ExcelWorksheet">Worksheet containing data to get.</param>
        /// <param name="HeaderRowIndex">Index of row that contain the header.</param>
        /// <returns>List of string containing header.</returns>
        public static List<string> GetTableHeader(Excel.Worksheet ExcelWorksheet,
            int HeaderRowIndex)
        {
            List<string> result = new List<string>();
            int TotalColumn = GetTotalColumnFromWorksheet(ExcelWorksheet);
            for (int i = 1; i <= TotalColumn; i++)
            {
                result.Add((ExcelWorksheet.Cells[HeaderRowIndex, i] as Excel.Range).Value2.ToString());
            }
            return result;
        }
        #endregion
    }
}
