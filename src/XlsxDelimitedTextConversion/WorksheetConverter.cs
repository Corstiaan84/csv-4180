//
// Copyright © Leonid S. Maliutin <mals@live.ru> 2015
//

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Linq;
using System.Text;

using Excel = NetOffice.ExcelApi;

namespace DocInOffice.XlsxDelimitedTextConversion
{
    public class WorksheetConverter
    {
        public WorksheetConverter()
        {

        }

        public void Convert(Excel.Worksheet worksheet, IDelimitedTextWriter textWriter)
        {
            using (Excel.Range usedRange = worksheet.UsedRange)
            {
                Int32 rowCount;
                using (Excel.Range rows = usedRange.Rows)
                    rowCount = rows.Count;


                Int32 columnCount;
                using (Excel.Range columns = usedRange.Columns)
                    columnCount = columns.Count;

                for (Int32 rowIndex = 1; rowIndex <= rowCount; rowIndex += 1)
                {
                    Int32 rightColumn = 0;

                    for (Int32 columnIndex = 1; columnIndex <= columnCount; columnIndex += 1)
                    {
                        if (rowIndex > 1 && columnIndex == 1)
                            textWriter.EndRecord();
                        else if (columnIndex > 1)
                            textWriter.EndField();

                        using (Excel.Range cellRange = (Excel.Range)usedRange.Cells[rowIndex, columnIndex])
                        {
                            String text = (String)cellRange.Text;
                            textWriter.Write(text);

                            if (String.IsNullOrEmpty(text) == false)
                                rightColumn = columnIndex;
                        }
                    }

                    this.OnRowConverted(rowIndex, rowCount);
                }
            }
        }

        protected virtual void OnRowConverted(Int32 rowIndex, Int32 rowCount)
        {

        }
    }
}
