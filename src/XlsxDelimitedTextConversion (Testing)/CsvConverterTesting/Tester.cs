//
// Copyright © Leonid S. Maliutin <mals@live.ru> 2014
//

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.IO;
using System.Text;

using NUnit.Framework;

using DocInOffice.Csv4180;

namespace DocInOffice.XlsxDelimitedTextConversion
{
    [TestFixture]
    public class Tester
    {
        private String _tempDirPath;
        private String _derictoryPath;

        [SetUp]
        public void SetUp()
        {
            DirectoryInfo tempDir = new DirectoryInfo(Path.GetFullPath(Path.Combine(TestContext.CurrentContext.TestDirectory, "Temp")));

            if (!tempDir.Exists)
                tempDir.Create();

            _tempDirPath = tempDir.FullName;

            _derictoryPath = Path.GetFullPath(Path.Combine(TestContext.CurrentContext.TestDirectory, @"..\..\CsvConverterTesting\Tester+Files"));
        }

        [Test]
        [Explicit]
        [TestCase("01_total.csv")]
        public void T01_TestConversion(String fileName)
        {
            String csvFilePath = this.GetFilePath(fileName);
            String xlsxFilePath = this.PrepareXlsxFile(csvFilePath);

            using (StreamReader streamReader = new StreamReader(csvFilePath))
            using (FileStream fileStream = File.OpenWrite(xlsxFilePath))
            {
                CsvReader csvReader = new CsvReader(streamReader);
                MoqDelimitedTextReader delimitedTextReader = new MoqDelimitedTextReader(csvReader);
                CsvConverter.Convert(delimitedTextReader, fileStream);

                fileStream.Flush();
            }

            Process myProcess = new Process();
            myProcess.StartInfo.FileName = xlsxFilePath;
            myProcess.Start();

        }

        private String GetFilePath(String fileName)
        {
            string filePath = Path.Combine(_derictoryPath, fileName);

            Assert.That(File.Exists(filePath), "\"{0}\" файл не найден.", filePath);

            return filePath;
        }

        private String PrepareXlsxFile(String csvFilePath)
        {
            String name = Path.GetFileNameWithoutExtension(csvFilePath);
            String fileName = name + ".xlsx";

            string docFilePath = Path.Combine(_tempDirPath, fileName);

            if (File.Exists(docFilePath))
                File.Delete(docFilePath);

            return docFilePath;
        }
    }
}
