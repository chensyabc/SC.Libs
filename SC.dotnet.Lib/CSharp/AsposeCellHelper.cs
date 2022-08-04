using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

/*------------------------------------------------------------------------------------------------
 * Change Log:
DATE		Developer		Description
8/9/2018	Soyann  Creation this entity
-------------------------------------------------------------------------------------------------*/
namespace SC.dotnet.Lib.CSharp
{
    public partial class AsposeCellHelper
    {
        #region Members

        protected static readonly log4net.ILog Log = log4net.LogManager.GetLogger(typeof(AsposeCellHelper));

        private string licenseFileName = AppDomain.CurrentDomain.BaseDirectory + @"\License\Aspose.Total.lic";

        public static AsposeCellHelper _asposeCellHelperInstance;

        #endregion


        #region Constructor

        public static AsposeCellHelper GetInstance()
        {
            if (_asposeCellHelperInstance == null)
            {
                _asposeCellHelperInstance = new AsposeCellHelper();
            }
            return _asposeCellHelperInstance;
        }

        /// <summary>
        /// Aspose convert excel/ppt/word to pdf
        /// </summary>
        private AsposeCellHelper()
        {
            try
            {
                Aspose.Cells.License licenseExcel = new Aspose.Cells.License();
                licenseExcel.SetLicense(licenseFileName);
            }
            catch (Exception ex)
            {
                Log.Error(ex);
                throw;
            }
        }

        #endregion


        #region Model Manipulation

        public Workbook GetWorkbookByDataSet(DataSet ds, bool includeHeader = true, List<ExportField> fieldsList = null, IDictionary<string, object> headerFields = null, string dateFormat = null)
        {
            if (ds != null && ds.Tables != null && ds.Tables.Count > 0)
            {
                Workbook workbook = new Workbook();
                workbook.Worksheets.Clear();
                List<string> sheetNameList = new List<string>();

                var headerCellStyle = CreateHeaderCellStyle(workbook);
                string datetimeFormat = System.Globalization.CultureInfo.CurrentUICulture.DateTimeFormat.ShortDatePattern;
                if (!string.IsNullOrEmpty(dateFormat))
                {
                    datetimeFormat = dateFormat;
                }

                var dateTimeCellStyle = CreateCellStyle(workbook, -1, datetimeFormat);
                var doubleCellStyle = CreateCellStyle(workbook, 4);
                var moneyCellStyle = CreateCellStyle(workbook, -1, "?#,##0.00;[red](?#,##0.00)");
                moneyCellStyle.HorizontalAlignment = TextAlignmentType.Right;
                var percentCellStyle = CreateCellStyle(workbook, 6, "0.000000%");
                percentCellStyle.HorizontalAlignment = TextAlignmentType.Right;
                var intCellStyle = CreateCellStyle(workbook, 1);
                var stringCellStyle = CreateCellStyle(workbook, -1, "Text");
                var styleFlag = new StyleFlag() { All = true };

                for (int i = 0; i < ds.Tables.Count; i++)
                {
                    // Sheet Name
                    string sheetName = "sheet_" + i.ToString();
                    if (!string.IsNullOrEmpty(ds.Tables[i].TableName))
                    {
                        sheetName = ds.Tables[i].TableName;
                    }
                    if (sheetNameList.Contains(sheetName))
                    {
                        sheetName += i.ToString();
                    }

                    workbook.Worksheets.Add(sheetName);

                    // Header Data
                    int baseRowId = 0;
                    if (headerFields != null && headerFields.Count > 0)
                    {
                        baseRowId = 1;

                        int cellCount = ds.Tables[0].Columns.Count;
                        int halfCellCount = (int)Math.Ceiling(cellCount / 2.0);

                        string headerLeftValue = string.Empty;// = moduleName + "\r\n";
                        foreach (string key in headerFields.Keys)
                        {
                            //sheet.CreateRow(rowId++);
                            headerLeftValue += key + (headerFields[key] != null ? ": " + headerFields[key] : string.Empty) + "\r\n";
                        }

                        Aspose.Cells.Range leftRange = workbook.Worksheets[i].Cells.CreateRange(0, 0, 1, halfCellCount - 1);
                        leftRange.Merge();
                        leftRange[0, 0].PutValue(headerLeftValue);

                        Aspose.Cells.Range rightRange = workbook.Worksheets[i].Cells.CreateRange(0, halfCellCount - 1, 1, cellCount - halfCellCount + 1);
                        rightRange.Merge();
                        //rightRange[0, 0].PutValue(ReportConstants.AlterDomusLabel.LOGO_LABEL);
                    }

                    // Sheet Data
                    ImportTableOptions importTableOptions = new ImportTableOptions();
                    workbook.Worksheets[i].Cells.ImportData(ds.Tables[i], baseRowId, 0, importTableOptions);
                    //Below function has been marked as Obsolete
                    //workbook.Worksheets[i].Cells.ImportDataTable(ds.Tables[i], includeHeader, "A" + (1 + baseRowId));
                    workbook.Worksheets[i].AutoFitColumns();

                    // Column Style
                    for (int j = 0; j < ds.Tables[i].Columns.Count; j++)
                    {
                        var associateFieldList = fieldsList?.FirstOrDefault(f => f.DisplayName == ds.Tables[i].Columns[j].ColumnName);

                        Type type = ds.Tables[i].Columns[j].DataType;
                        switch (type.FullName)
                        {
                            case ExportConstants.SystemDataType.INT32:
                                {
                                    workbook.Worksheets[i].Cells.Columns[j].ApplyStyle(intCellStyle, styleFlag);
                                    break;
                                }
                            case ExportConstants.SystemDataType.DATETIME:
                                {
                                    if (associateFieldList != null && !string.IsNullOrEmpty(associateFieldList.DateTimeFormat))
                                    {
                                        var specialDateTimeCellStyle = CreateCellStyle(workbook, -1, associateFieldList.DateTimeFormat);
                                        workbook.Worksheets[i].Cells.Columns[j].ApplyStyle(specialDateTimeCellStyle, styleFlag);
                                    }
                                    else
                                    {
                                        workbook.Worksheets[i].Cells.Columns[j].ApplyStyle(dateTimeCellStyle, styleFlag);
                                    }
                                    break;
                                }
                            case ExportConstants.SystemDataType.DECIMAL:
                            case ExportConstants.SystemDataType.DOUBLE:
                                {
                                    if (ds.Tables[i].Columns[j].ColumnName.Trim().EndsWith("%"))
                                    {
                                        workbook.Worksheets[i].Cells.Columns[j].ApplyStyle(percentCellStyle, styleFlag);
                                        //workbook.Worksheets[i].Cells.Columns[j].Style.Number = 6;//Hard code to apply it into cell.
                                    }
                                    else
                                    {
                                        if (associateFieldList != null && associateFieldList.IsMoney)
                                        {
                                            workbook.Worksheets[i].Cells.Columns[j].ApplyStyle(moneyCellStyle, styleFlag);
                                        }
                                        else
                                        {
                                            workbook.Worksheets[i].Cells.Columns[j].ApplyStyle(doubleCellStyle, styleFlag);
                                        }
                                    }

                                    break;
                                }
                            default:
                                {
                                    workbook.Worksheets[i].Cells.Columns[j].ApplyStyle(stringCellStyle, styleFlag);
                                    break;
                                }
                        }
                    }

                    // Header Style
                    if (includeHeader)
                    {
                        // Header Row Height
                        workbook.Worksheets[i].Cells.SetRowHeight(baseRowId, 40);

                        for (int j = 0; j < ds.Tables[i].Columns.Count; j++)
                        {
                            workbook.Worksheets[i].Cells[baseRowId, j].SetStyle(headerCellStyle, true);

                            // Header Column Width
                            var associateFieldList = fieldsList?.FirstOrDefault(f => f.DisplayName == ds.Tables[i].Columns[j].ColumnName);
                            workbook.Worksheets[i].Cells.Columns[j].Width = associateFieldList != null ? associateFieldList.Length : 25;
                        }
                    }

                    if (headerFields != null && headerFields.Count > 0)
                    {
                        var headerStyle = CreateHeaderTitleCellStyle(workbook, TextAlignmentType.Left);
                        var headerStyleWithLogo = CreateHeaderTitleCellStyle(workbook, TextAlignmentType.Right);

                        workbook.Worksheets[i].Cells.SetRowHeight(0, headerFields.Count * 20);
                        workbook.Worksheets[i].Cells[0, 0].SetStyle(headerStyle, true);

                        int halfCellCount = (int)Math.Ceiling(ds.Tables[0].Columns.Count / 2.0);
                        workbook.Worksheets[i].Cells[0, halfCellCount - 1].SetStyle(headerStyleWithLogo, true);
                    }

                    // Cell Height exclude Header
                    //for (int j = (includeHeader ? 1 : 0); j < workbook.Worksheets[i].Cells.Rows.Count; j++)
                    //{
                    //    workbook.Worksheets[i].Cells.SetRowHeight(j, 20);
                    //}
                }

                return workbook;
            }
            else
            {
                return null;
            }
        }

        #endregion


        #region Help Method

        internal Style CreateHeaderTitleCellStyle(Workbook workbook, TextAlignmentType horizontalAlignment)
        {
            var cellStyle = workbook.CreateStyle();
            cellStyle.Font.Name = "Calibri";
            cellStyle.Font.IsBold = true;
            cellStyle.Font.Size = 11;
            cellStyle.VerticalAlignment = TextAlignmentType.Top;
            cellStyle.HorizontalAlignment = horizontalAlignment;

            return cellStyle;
        }

        internal Style CreateHeaderCellStyle(Workbook workbook)
        {
            var cellStyle = workbook.CreateStyle();
            cellStyle.Font.Name = "Calibri";
            cellStyle.Font.IsBold = true;
            cellStyle.Font.Size = 11;
            //cellStyle.Font.Color = Color.FromArgb(255, 255, 255);
            cellStyle.Font.Underline = FontUnderlineType.Single;
            cellStyle.Pattern = BackgroundType.Solid;
            //cellStyle.ForegroundColor = Color.FromArgb(105, 105, 105);
            cellStyle.VerticalAlignment = TextAlignmentType.Center;
            cellStyle.HorizontalAlignment = TextAlignmentType.Center;

            return cellStyle;
        }

        /// <summary>
        /// Create Format Cell Style
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="numberFormat">bigger than 0 will set Number,else not set;</param>
        /// <param name="customFormat">customFormat!=null && customFormat!="Text" will set custom format cell</param>
        /// <returns></returns>
        internal Style CreateCellStyle(Workbook workbook, int numberFormat, string customFormat = null)
        {
            var cellStyle = workbook.CreateStyle();
            cellStyle.HorizontalAlignment = TextAlignmentType.Center;
            if (numberFormat >= 0)
            {
                cellStyle.Number = numberFormat;
                cellStyle.HorizontalAlignment = TextAlignmentType.Right;
            }
            if (!string.IsNullOrEmpty(customFormat) && customFormat != "Text")
            {
                cellStyle.SetCustom(customFormat, true);
            }
            cellStyle.Font.Name = "Calibri";
            cellStyle.Font.Size = 11;
            cellStyle.VerticalAlignment = TextAlignmentType.Center;
            cellStyle.IsTextWrapped = true;
            return cellStyle;
        }

        #endregion
    }
}
