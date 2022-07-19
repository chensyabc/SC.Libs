using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Cortland.Fund.Tools.Excel.Util
{
    public class ExcelUtil
    {
        public static string OleDbConnectionString_XLS = "Provider=Microsoft.Jet.OleDb.4.0;data source={0};Extended Properties='Excel 8.0; HDR=Yes;'";
        public static string OleDbConnectionString_XLSX = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0;HDR=Yes;'";

        /// <summary>
        /// Can fit for xls and xlsx.
        /// </summary>
        /// <param name="excelFilePathWithExtension"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static DataSet ReadExcelSheetToDataSet(string excelFilePathWithExtension, string sheetName)
        {
            try
            {
                string connectionString = "";

                if (excelFilePathWithExtension.Contains(".xlsx"))
                {
                    connectionString = string.Format(OleDbConnectionString_XLSX, excelFilePathWithExtension);
                }
                else
                {
                    connectionString = string.Format(OleDbConnectionString_XLS, excelFilePathWithExtension);
                }

                OleDbConnection conn = new OleDbConnection(connectionString);
                conn.Open();
                string excelQuerySheet = "";
                OleDbDataAdapter myCommand = null;
                DataSet ds = null;
                excelQuerySheet = "select * from [" + sheetName + "$]";
                myCommand = new OleDbDataAdapter(excelQuerySheet, connectionString);
                ds = new DataSet();
                myCommand.Fill(ds);

                conn.Close();
                conn.Dispose();

                return ds;
            }
            catch (OleDbException oledbe)
            {
                throw new Exception("OLEDB Error: " + oledbe.Message);
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static void DataSetToExcel(string excelFilePathWithExtension, DataSet oldDataSet, string sheetName)
        {
            //先得到汇总Excel的DataSet 主要目的是获得Excel在DataSet中的结构  
            string strCon = string.Format(OleDbConnectionString_XLSX, excelFilePathWithExtension);
            //string strCon = "Provider=Microsoft.Jet.OleDb.4.0;" + "data source=" +Path+ ";Extended Properties='Excel 8.0; HDR=Yes; IMEX=1'"; //此连接只能操作Excel2007之前(.xls)文件  
            //string strCon = "Provider=Microsoft.Ace.OleDb.12.0;" + "data source=" + Path + ";Extended Properties='Excel 12.0; HDR=No; IMEX=0'"; //此连接可以操作.xls与.xlsx文件 (支持Excel2003 和 Excel2007 的连接字符串)  
            //备注： "HDR=yes;"是说Excel文件的第一行是列名而不是数据，"HDR=No;"正好与前面的相反。//      "IMEX=1 "如果列中的数据类型不一致，使用"IMEX=1"可必免数据类型冲突。   

            OleDbConnection myConn = new OleDbConnection(strCon);
            string strCom = "select * from [" + sheetName + "$]";
            myConn.Open();
            OleDbDataAdapter myCommand = new OleDbDataAdapter(strCom, myConn);
            System.Data.OleDb.OleDbCommandBuilder builder = new OleDbCommandBuilder(myCommand);
            //QuotePrefix和QuoteSuffix主要是对builder生成InsertComment命令时使用。  
            builder.QuotePrefix = "[";     //获取insert语句中保留字符（起始位置）  
            builder.QuoteSuffix = "]"; //获取insert语句中保留字符（结束位置）  
            DataSet newDataSet = new DataSet();
            myCommand.Fill(newDataSet, "Table1");
            for (int i = 0; i < oldDataSet.Tables[0].Rows.Count; i++)
            {
                //在这里不能使用ImportRow方法将一行导入到news中，  
                //因为ImportRow将保留原来DataRow的所有设置(DataRowState状态不变)。  
                //在使用ImportRow后newds内有值，但不能更新到Excel中因为所有导入行的DataRowState!=Added  
                DataRow nrow = newDataSet.Tables["Table1"].NewRow();
                //nrow[0] = oldDataSet.Tables[0].Rows[i][0];

                for (int j = 0; j < oldDataSet.Tables[0].Columns.Count; j++)
                {
                    nrow[j] = oldDataSet.Tables[0].Rows[i][j];
                }

                newDataSet.Tables["Table1"].Rows.Add(nrow);
            }
            myCommand.Update(newDataSet, "Table1");
            myConn.Close();
            myConn.Dispose();
        }

        /// <summary>
        /// Excel without report template header
        /// </summary>
        /// <param name="exportFilePath"></param>
        /// <param name="sourceTable"></param>
        /// <returns></returns>
        public static string DataTableToExcelWithNPOI(string exportFilePath, DataTable sourceTable, string sheetName)
        {
            string message = "";
            try
            {
                XSSFWorkbook book = null;

                using (FileStream fs = File.OpenRead(exportFilePath))
                {
                    book = new XSSFWorkbook(fs);
                    fs.Close();
                    NPOI.SS.UserModel.ISheet sheet = book.GetSheet(sheetName);

                    #region Set Excel Table Header Style
                    XSSFCellStyle tableHeaderStyle = CreateTableHeaderXSSFCellStyle(book);
                    int i = 0;
                    NPOI.SS.UserModel.IRow titleRow = sheet.CreateRow(i++);
                    titleRow.Height = 40 * 20;
                    int columnOrdinal = 0;
                    foreach (DataColumn item in sourceTable.Columns)
                    {
                        int a = columnOrdinal++;
                        titleRow.CreateCell(a).SetCellValue(item.ColumnName);
                        titleRow.Cells[a].CellStyle = tableHeaderStyle;
                        sheet.SetColumnWidth(a, 25 * 256);
                    }
                    #endregion

                    #region Set Specified Data Type Style
                    IDataFormat commonFormat = book.CreateDataFormat();

                    XSSFCellStyle dateTimeCellStyle = CreateXSSFCellStyle(book);
                    dateTimeCellStyle.DataFormat = commonFormat.GetFormat("MM/dd/yyyy");

                    //XSSFCellStyle doubleCellStyle = CreateXSSFCellStyle(book);
                    //doubleCellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");

                    XSSFCellStyle moneyCellStyle = CreateXSSFCellStyle(book);
                    moneyCellStyle.DataFormat = commonFormat.GetFormat("?#,##0.00");

                    XSSFCellStyle percentCellStyle = CreateXSSFCellStyle(book);
                    percentCellStyle.DataFormat = commonFormat.GetFormat("0.0000000000%");

                    XSSFCellStyle intCellStyle = CreateXSSFCellStyle(book);
                    intCellStyle.DataFormat = commonFormat.GetFormat("0");

                    XSSFCellStyle stringCellStyle = CreateXSSFCellStyle(book);
                    #endregion

                    #region Set File Into Excel
                    foreach (DataRow dr in sourceTable.Rows)
                    {
                        NPOI.SS.UserModel.IRow tmpRow = sheet.CreateRow(i++);

                        for (int j = 0; j < sourceTable.Columns.Count; j++)
                        {
                            Type type = sourceTable.Columns[j].DataType;

                            switch (type.FullName)
                            {
                                case "System.Int32":
                                    tmpRow.CreateCell(j).SetCellValue(Convert.ToInt32(dr[j].ToString()));
                                    tmpRow.Cells[j].CellStyle = intCellStyle;
                                    break;

                                case "System.DateTime":
                                    if (string.IsNullOrWhiteSpace(dr[j].ToString()))
                                    {
                                        tmpRow.CreateCell(j).SetCellValue("");
                                    }
                                    else
                                    {
                                        tmpRow.CreateCell(j).SetCellValue(Convert.ToDateTime(dr[j].ToString()));
                                    }
                                    //tmpRow.Cells[j].CellStyle = dateTimeCellStyle;
                                    break;

                                case "System.Double":
                                    if (dr[j] == DBNull.Value)
                                    {
                                        tmpRow.CreateCell(j).SetCellValue("N/A");
                                    }
                                    else
                                    {
                                        tmpRow.CreateCell(j).SetCellValue(Convert.ToDouble(dr[j].ToString()));
                                    }
                                    tmpRow.Cells[j].CellStyle = moneyCellStyle;
                                    break;

                                case "System.Decimal":
                                    tmpRow.CreateCell(j).SetCellValue(Convert.ToDouble(dr[j].ToString()));
                                    tmpRow.Cells[j].CellStyle = percentCellStyle;
                                    break;

                                default:
                                    tmpRow.CreateCell(j).SetCellValue(dr[j].ToString());
                                    tmpRow.Cells[j].CellStyle = stringCellStyle;
                                    break;
                            }
                        }
                    }
                    #endregion

                    string testxlsx = "test.xlsx";
                    if (File.Exists(testxlsx))
                    {
                        File.Delete(testxlsx);
                    }
                    var exportFile = File.Create(testxlsx);
                    book.Write(exportFile);
                }

                //using (FileStream fs = File.OpenRead(exportFilePath))
                //{
                //    book.Write(fs);
                //    fs.Close();
                //}
            }
            catch (FileNotFoundException fnfe)
            {
                message = "File Not Found,Export File Failed.";
            }
            catch (IOException ioe)
            {
                message = "Read/Write Data Error,Export File Failed.";
            }
            catch (FormatException fe)
            {
                message = "Convert Data error,Export File Failed.";
            }
            catch (Exception e)
            {
                message = "Export File Failed.";
            }

            return message;
        }

        public static string DataTablesSheetNamesDicToExcelWithNPOI(string exportFilePath, List<KeyValuePair<string, DataTable>> dataSheetPairs)
        {
            string message = "";
            try
            {
                XSSFWorkbook book = null;

                using (FileStream fs = File.Open(exportFilePath, FileMode.Open, FileAccess.Read))
                {
                    book = new XSSFWorkbook(fs);
                    fs.Close();

                    foreach (KeyValuePair<string, DataTable> dataSheetPair in dataSheetPairs)
                    {
                        string sheetName = dataSheetPair.Key;
                        DataTable sourceTable = dataSheetPair.Value;

                        NPOI.SS.UserModel.ISheet sheet = book.GetSheet(dataSheetPair.Key);

                        #region Set Excel Table Header Style
                        XSSFCellStyle tableHeaderStyle = CreateTableHeaderXSSFCellStyle(book);
                        int i = 0;
                        NPOI.SS.UserModel.IRow titleRow = sheet.CreateRow(i++);

                        titleRow.Height = 40 * 20;
                        int columnOrdinal = 0;
                        foreach (DataColumn item in sourceTable.Columns)
                        {
                            int a = columnOrdinal++;
                            titleRow.CreateCell(a).SetCellValue(item.ColumnName);
                            titleRow.Cells[a].CellStyle = tableHeaderStyle;
                            sheet.SetColumnWidth(a, 25 * 256);
                        }
                        #endregion

                        #region Set Specified Data Type Style
                        //IDataFormat commonFormat = book.CreateDataFormat();

                        //XSSFCellStyle dateTimeCellStyle = CreateXSSFCellStyle(book);
                        //dateTimeCellStyle.DataFormat = commonFormat.GetFormat("MM/dd/yyyy");

                        //XSSFCellStyle doubleCellStyle = CreateXSSFCellStyle(book);
                        //doubleCellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");

                        //XSSFCellStyle moneyCellStyle = CreateXSSFCellStyle(book);
                        //moneyCellStyle.DataFormat = commonFormat.GetFormat("?#,##0.00");

                        //XSSFCellStyle percentCellStyle = CreateXSSFCellStyle(book);
                        //percentCellStyle.DataFormat = commonFormat.GetFormat("0.0000000000%");

                        //XSSFCellStyle intCellStyle = CreateXSSFCellStyle(book);
                        //intCellStyle.DataFormat = commonFormat.GetFormat("0");

                        //XSSFCellStyle stringCellStyle = CreateXSSFCellStyle(book);
                        #endregion

                        #region Set File Into Excel
                        foreach (DataRow dr in sourceTable.Rows)
                        {
                            NPOI.SS.UserModel.IRow tmpRow = sheet.CreateRow(i++);

                            for (int j = 0; j < sourceTable.Columns.Count; j++)
                            {
                                Type type = sourceTable.Columns[j].DataType;

                                switch (type.FullName)
                                {
                                    case "System.Int32":
                                        tmpRow.CreateCell(j).SetCellValue(Convert.ToInt32(dr[j].ToString()));
                                        //tmpRow.Cells[j].CellStyle = intCellStyle;
                                        break;

                                    case "System.DateTime":
                                        if (string.IsNullOrWhiteSpace(dr[j].ToString()))
                                        {
                                            tmpRow.CreateCell(j).SetCellValue("");
                                        }
                                        else
                                        {
                                            tmpRow.CreateCell(j).SetCellValue(Convert.ToDateTime(dr[j].ToString()));
                                        }
                                        //tmpRow.Cells[j].CellStyle = dateTimeCellStyle;
                                        break;

                                    case "System.Double":
                                        if (dr[j] == DBNull.Value)
                                        {
                                            tmpRow.CreateCell(j).SetCellValue("N/A");
                                        }
                                        else
                                        {
                                            tmpRow.CreateCell(j).SetCellValue(Convert.ToDouble(dr[j].ToString()));
                                        }
                                        //tmpRow.Cells[j].CellStyle = moneyCellStyle;
                                        break;

                                    case "System.Decimal":
                                        tmpRow.CreateCell(j).SetCellValue(Convert.ToDouble(dr[j].ToString()));
                                        //tmpRow.Cells[j].CellStyle = percentCellStyle;
                                        break;

                                    default:
                                        tmpRow.CreateCell(j).SetCellValue(dr[j].ToString());
                                        //tmpRow.Cells[j].CellStyle = stringCellStyle;
                                        break;
                                }
                            }
                        }

                        #endregion
                    }

                    string testxlsx = @"C:\Users\Admin\Desktop\FundVBATest\test.xlsx";
                    if (File.Exists(testxlsx))
                    {
                        File.Delete(testxlsx);
                    }
                    var exportFile = File.Create(testxlsx);
                    book.Write(exportFile);
                }

                //using (FileStream fs = File.Open(exportFilePath, FileMode.Open, FileAccess.Write))
                //{
                //    book.Write(fs);
                //    fs.Close();
                //}

                book.Close();
            }
            catch (FileNotFoundException fnfe)
            {
                message = "File Not Found,Export File Failed.";
            }
            catch (IOException ioe)
            {
                message = "Read/Write Data Error,Export File Failed.";
            }
            catch (FormatException fe)
            {
                message = "Convert Data error,Export File Failed.";
            }
            catch (Exception e)
            {
                message = "Export File Failed.";
            }

            return message;
        }

        public static string DataTablesSheetNamesDicToExcelXLSWithNPOI(string exportFilePath, List<KeyValuePair<string, DataTable>> dataSheetPairs)
        {
            string message = "";
            try
            {
                HSSFWorkbook book = null;

                using (FileStream fs = File.Open(exportFilePath, FileMode.Open, FileAccess.Read))
                {
                    book = new HSSFWorkbook(fs);
                    fs.Close();

                    foreach (KeyValuePair<string, DataTable> dataSheetPair in dataSheetPairs)
                    {
                        string sheetName = dataSheetPair.Key;
                        DataTable sourceTable = dataSheetPair.Value;

                        NPOI.SS.UserModel.ISheet sheet = book.GetSheet(dataSheetPair.Key);

                        #region Set Excel Table Header Style
                        //XSSFCellStyle tableHeaderStyle = CreateTableHeaderXSSFCellStyle(book);
                        int i = 0;
                        NPOI.SS.UserModel.IRow titleRow = sheet.CreateRow(i++);

                        titleRow.Height = 40 * 20;
                        int columnOrdinal = 0;
                        foreach (DataColumn item in sourceTable.Columns)
                        {
                            int a = columnOrdinal++;
                            titleRow.CreateCell(a).SetCellValue(item.ColumnName);
                            //titleRow.Cells[a].CellStyle = tableHeaderStyle;
                            sheet.SetColumnWidth(a, 25 * 256);
                        }
                        #endregion

                        #region Set Specified Data Type Style
                        IDataFormat commonFormat = book.CreateDataFormat();

                        //XSSFCellStyle dateTimeCellStyle = CreateXSSFCellStyle(book);
                        //dateTimeCellStyle.DataFormat = commonFormat.GetFormat("MM/dd/yyyy");

                        //XSSFCellStyle doubleCellStyle = CreateXSSFCellStyle(book);
                        //doubleCellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");

                        //HSSFCellStyle moneyCellStyle = CreateXSSFCellStyle(book);
                        //moneyCellStyle.DataFormat = commonFormat.GetFormat("?#,##0.00");

                        //HSSFCellStyle percentCellStyle = CreateXSSFCellStyle(book);
                        //percentCellStyle.DataFormat = commonFormat.GetFormat("0.0000000000%");

                        //HSSFCellStyle intCellStyle = CreateXSSFCellStyle(book);
                        //intCellStyle.DataFormat = commonFormat.GetFormat("0");

                        //HSSFCellStyle stringCellStyle = CreateXSSFCellStyle(book);
                        #endregion

                        #region Set File Into Excel
                        foreach (DataRow dr in sourceTable.Rows)
                        {
                            NPOI.SS.UserModel.IRow tmpRow = sheet.CreateRow(i++);

                            for (int j = 0; j < sourceTable.Columns.Count; j++)
                            {
                                Type type = sourceTable.Columns[j].DataType;

                                switch (type.FullName)
                                {
                                    case "System.Int32":
                                        tmpRow.CreateCell(j).SetCellValue(Convert.ToInt32(dr[j].ToString()));
                                        //tmpRow.Cells[j].CellStyle = intCellStyle;
                                        break;

                                    case "System.DateTime":
                                        if (string.IsNullOrWhiteSpace(dr[j].ToString()))
                                        {
                                            tmpRow.CreateCell(j).SetCellValue("");
                                        }
                                        else
                                        {
                                            tmpRow.CreateCell(j).SetCellValue(Convert.ToDateTime(dr[j].ToString()));
                                        }
                                        //tmpRow.Cells[j].CellStyle = dateTimeCellStyle;
                                        break;

                                    case "System.Double":
                                        if (dr[j] == DBNull.Value)
                                        {
                                            tmpRow.CreateCell(j).SetCellValue("N/A");
                                        }
                                        else
                                        {
                                            tmpRow.CreateCell(j).SetCellValue(Convert.ToDouble(dr[j].ToString()));
                                        }
                                        //tmpRow.Cells[j].CellStyle = moneyCellStyle;
                                        break;

                                    case "System.Decimal":
                                        tmpRow.CreateCell(j).SetCellValue(Convert.ToDouble(dr[j].ToString()));
                                        //tmpRow.Cells[j].CellStyle = percentCellStyle;
                                        break;

                                    default:
                                        tmpRow.CreateCell(j).SetCellValue(dr[j].ToString());
                                        //tmpRow.Cells[j].CellStyle = stringCellStyle;
                                        break;
                                }
                            }
                        }

                        #endregion
                    }

                    string testxlsx = @"C:\Users\Admin\Desktop\FundVBATest\test.xls";
                    if (File.Exists(testxlsx))
                    {
                        File.Delete(testxlsx);
                    }
                    var exportFile = File.Create(testxlsx);
                    book.Write(exportFile);
                }

                //using (FileStream fs = File.Open(exportFilePath, FileMode.Open, FileAccess.Write))
                //{
                //    book.Write(fs);
                //    fs.Close();
                //}

                book.Close();
            }
            catch (FileNotFoundException fnfe)
            {
                message = "File Not Found,Export File Failed.";
            }
            catch (IOException ioe)
            {
                message = "Read/Write Data Error,Export File Failed.";
            }
            catch (FormatException fe)
            {
                message = "Convert Data error,Export File Failed.";
            }
            catch (Exception e)
            {
                message = "Export File Failed.";
            }

            return message;
        }


        #region Help Method

        private static XSSFCellStyle CreateXSSFCellStyle(XSSFWorkbook book)
        {
            XSSFFont commonFont = (XSSFFont)book.CreateFont();
            commonFont.FontName = "Calibri";
            commonFont.FontHeightInPoints = 11;

            XSSFCellStyle cellStyle = (XSSFCellStyle)book.CreateCellStyle();
            cellStyle.VerticalAlignment = VerticalAlignment.Center;
            cellStyle.Alignment = HorizontalAlignment.Center;
            cellStyle.WrapText = true;
            cellStyle.SetFont(commonFont);
            return cellStyle;
        }

        private static XSSFCellStyle CreateReportHeaderXSSFCellStyle(XSSFWorkbook book)
        {
            XSSFFont headerFont = (XSSFFont)book.CreateFont();
            headerFont.FontName = "Calibri";
            headerFont.IsBold = true;
            headerFont.FontHeightInPoints = 11;

            XSSFCellStyle headerStyle = (XSSFCellStyle)book.CreateCellStyle();
            headerStyle.SetFont(headerFont);
            headerStyle.WrapText = true;
            return headerStyle;
        }

        private static XSSFCellStyle CreateTableHeaderXSSFCellStyle(XSSFWorkbook book)
        {
            XSSFFont font = (XSSFFont)book.CreateFont();
            font.FontName = "Calibri";
            font.IsBold = true;
            font.FontHeightInPoints = 11;
            font.SetColor(new XSSFColor(Color.FromArgb(255, 255, 255)));
            font.Underline = FontUnderlineType.Single;

            XSSFCellStyle tableHeaderStyle = (XSSFCellStyle)book.CreateCellStyle();
            tableHeaderStyle.FillPattern = FillPattern.SolidForeground;
            tableHeaderStyle.SetFillForegroundColor(new XSSFColor(Color.FromArgb(105, 105, 105)));
            tableHeaderStyle.Alignment = HorizontalAlignment.Center;
            tableHeaderStyle.VerticalAlignment = VerticalAlignment.Center;
            tableHeaderStyle.SetFont(font);
            return tableHeaderStyle;
        }

        #endregion
    }
}
