using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;
using NetOffice.OfficeApi.Enums;

namespace DocInOffice.XlsxDelimitedTextConversion
{
    public static partial class XlsxConverter
    {
        public static void Convert(String xlsxFilePath, IDelimitedTextWriter textWriter, WorksheetConverter worksheetConverter = null)
        {
            if (xlsxFilePath == null) throw new ArgumentNullException(nameof(xlsxFilePath));
            if (textWriter == null) throw new ArgumentNullException(nameof(textWriter));

            if (System.IO.File.Exists(xlsxFilePath) == false)
                throw new FileNotFoundException(String.Format("\"{0}\" file has not been found.", xlsxFilePath));

            if(worksheetConverter == null)
                worksheetConverter = new WorksheetConverter();

            using (Excel.Application excelApplication = new Excel.Application())
            {
                try
                {
                    excelApplication.Visible = false;
                    excelApplication.ScreenUpdating = false;
                    excelApplication.EnableEvents = false;

                    using (Excel.Workbooks workbooks = excelApplication.Workbooks)
                    {
                        using (Excel.Workbook emptyExcelWorkbook =
                            workbooks.Count == 0 ? workbooks.Add() : null)
                        {
                            try
                            {
                                excelApplication.Calculation = XlCalculation.xlCalculationManual;

                                using (var excelWorkbook = workbooks.Open(filename: Path.GetFullPath(xlsxFilePath)))
                                {
                                    Object activeSheet = excelWorkbook.ActiveSheet;
                                    try
                                    {
                                        if (activeSheet is Excel.Worksheet == false)
                                            throw new ActiveWorksheetException();

                                        Excel.Worksheet worksheet = (Excel.Worksheet)activeSheet;
                                        worksheetConverter.Convert(worksheet, textWriter);
                                    }
                                    finally
                                    {
                                        (activeSheet as IDisposable)?.Dispose();
                                    }
                                }
                            }
                            finally
                            {
                                emptyExcelWorkbook?.Close(saveChanges: false);
                            }
                        }
                    }
                }
                finally
                {
                    excelApplication.Quit();
                }
            }
        }
    }
}
