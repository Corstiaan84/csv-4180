using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;

using NUnit.Framework;

using DocInOffice.Csv4180;

namespace DocInOffice.XlsxDelimitedTextConversion.XlsxConverterTesting
{
    public static class TesterEngine
    {

        public static IEnumerable<TestCaseData> GetTestSource(String sourceFolderName, String sourceFolderPath)
        {

            String[] testDirectories = Directory.GetDirectories(Path.GetFullPath(Path.Combine(Path.Combine(TestContext.CurrentContext.WorkDirectory, sourceFolderPath), sourceFolderName)));

            foreach (String testDirectory in testDirectories)
            {
                TestCaseData testCaseData = new TestCaseData(sourceFolderName, testDirectory);

                yield return testCaseData;
            }
        }

        public static void Test(String tempFolderName, String sourceSubfolderPath)
        {
            String subfolderName = Path.GetFileName(sourceSubfolderPath);

            String tempFolder = Path.GetFullPath(Path.Combine(TestContext.CurrentContext.WorkDirectory, @"Temp\XlsxConverterTesting"));
            String tempFolderPath = Path.Combine(tempFolder, subfolderName);

            String xlsxFilePath = GetFilePath(Path.Combine(sourceSubfolderPath, "source.xlsx"));
            String templateFilePath = GetFilePath(Path.Combine(sourceSubfolderPath, "template.csv"));
            String csvFilePath = PrepareCsvFile(Path.Combine(tempFolder, tempFolderName, subfolderName, "result.csv"));

            using (StreamWriter streamWriter = new StreamWriter(csvFilePath))
            {
                CsvWriter csvWriter = new CsvWriter(streamWriter);
                MoqDelimitedTextWriter delimitedTextWriter = new MoqDelimitedTextWriter(csvWriter);
                XlsxConverter.Convert(xlsxFilePath, delimitedTextWriter);
                delimitedTextWriter.Flush();
                streamWriter.Flush();
            }

            FileAssert.AreEqual(csvFilePath, templateFilePath);
        }

        private static String GetFilePath(String fileNamePath)
        {
            Assert.That(File.Exists(fileNamePath), "\"{0}\" файл не найден.", fileNamePath);

            return fileNamePath;
        }

        private static String PrepareCsvFile(String csvFilePath)
        {
            String directoryPath = Path.GetDirectoryName(csvFilePath);

            if (Directory.Exists(directoryPath) == false)
                Directory.CreateDirectory(directoryPath);

            if (File.Exists(csvFilePath))
                File.Delete(csvFilePath);

            return csvFilePath;
        }
    }
}
