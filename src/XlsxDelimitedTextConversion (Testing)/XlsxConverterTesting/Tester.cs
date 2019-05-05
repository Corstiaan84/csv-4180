//
// Copyright © Leonid S. Maliutin <mals@live.ru> 2015
//

using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;

using NUnit.Framework;

using DocInOffice.Csv4180;

namespace DocInOffice.XlsxDelimitedTextConversion.XlsxConverterTesting
{
    public class Tester
    {
        [Test]
        public void Test_001()
        {
            Assert.Catch<ActiveWorksheetException>(
                () =>
                {
                    String xlsxFilePath = Path.GetFullPath(Path.Combine(TestContext.CurrentContext.TestDirectory, @"..\..\XlsxConverterTesting\Tester+Files\001\source.xlsx"));
                    String csvFilePath = Path.GetTempFileName();
                    try
                    {
                        using (StreamWriter streamWriter = new StreamWriter(csvFilePath))
                        {
                            CsvWriter csvWriter = new CsvWriter(streamWriter);
                            MoqDelimitedTextWriter delimitedTextWriter = new MoqDelimitedTextWriter(csvWriter);
                            XlsxConverter.Convert(xlsxFilePath, delimitedTextWriter);
                            csvWriter.Flush();
                            streamWriter.Flush();
                        }
                    }
                    finally
                    {
                        File.Delete(csvFilePath);
                    }
                });
        }
    }
}
